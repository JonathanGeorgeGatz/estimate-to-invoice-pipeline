/** THE ESTIMATE-TO-INVOICE PIPELINE - Google Sheets + Apps Script
 *  Build It With AI - Episode 9. Free to use. Sample data is fictional (@example.com).
 *
 *  SET-UP (one time, ~2 minutes):
 *   1. Open a new Google Sheet.
 *   2. Extensions -> Apps Script, delete whatever is in the editor, paste this
 *      whole file, hit the disk icon to save.
 *   3. Reload the spreadsheet tab. A "Pipeline" menu appears.
 *   4. Pipeline -> "Set up tabs + sample data" to see it working, or create the
 *      same tabs with your own rows (headers in each tab must match).
 *   5. Pipeline -> "Rebuild status board".
 */
/** The Estimate-to-Invoice Pipeline — core engine (final).
 *  Build It With AI, Episode 9. Pure compute: no SpreadsheetApp in here, so the
 *  same code runs in Apps Script, in a browser, and in the test harness.
 *
 *  SPEC (the one both implementations are held to):
 *  - Money: round half up to cents at every defined step (r2).
 *  - Estimate: subtotal = sum of r2(qty*unit_price); taxableBase = taxable lines only;
 *    estTax = r2(taxableBase * rate); estTotal = r2(subtotal + estTax).
 *  - Deposit: fixed at acceptance = r2(estTotal * deposit_pct/100). Change orders
 *    never change the deposit.
 *  - Invoice amount for a job: snapshot of the accepted estimate's lines PLUS every
 *    APPROVED change order (change orders are taxable). invSub = r2(subtotal + sumCO);
 *    invTaxBase = r2(taxableBase + sumCO); invTax = r2(invTaxBase * rate);
 *    invTotal = r2(invSub + invTax); amountDue = r2(invTotal - deposit).
 *  - A PENDING change order (approved blank) BLOCKS invoicing: the job is not
 *    "ready to invoice" and never silently drops the pending work.
 *  - Job stage from events: DONE > STARTED > SCHEDULED (highest wins; NOTE never
 *    changes the stage, only counts as activity). Jobs are matched to invoices by
 *    job_id — never by customer.
 *  - Invoice status: paid = sum of its payments; PAID if balance <= 0.005;
 *    PARTIAL if 0 < paid < due; OVERDUE if unpaid/partial and daysSince(date_sent)
 *    > overdue_days; else SENT.
 *  - Stalled: STARTED (not DONE) and no event of any kind within stale_days.
 *  - Board buckets, in order: WAITING (SENT estimates), ACCEPTED_NOT_STARTED,
 *    IN_PROGRESS, STALLED, BLOCKED (pending CO), DONE_NOT_INVOICED, INVOICED
 *    (incl. OVERDUE flag), PAID (payments within paid_window_days).
 *  - Reconciliation must tie to the cent: for every invoice,
 *    amountDue = paid + balance; and board value conservation: every ACCEPTED
 *    estimate lands in exactly one of {no-job exception, job bucket}.
 */
function r2(x) { return Math.round(x * 100 + (x >= 0 ? 0 : -0)) / 100; }
function days(a, b) { // whole days from a to b (ISO strings), UTC
  return Math.round((Date.parse(b + "T00:00:00Z") - Date.parse(a + "T00:00:00Z")) / 86400000);
}

function computePipeline(data, settings) {
  var rate = settings.tax_rate_pct / 100;
  var asOf = data.as_of;
  var linesBy = {};
  data.estimate_lines.forEach(function (l) {
    (linesBy[l.est_id] = linesBy[l.est_id] || []).push(l);
  });
  var estBy = {}, est = [];
  data.estimates.forEach(function (e) {
    var sub = 0, taxable = 0;
    (linesBy[e.est_id] || []).forEach(function (l) {
      var v = r2(l.qty * l.unit_price);
      sub = r2(sub + v);
      if (l.taxable === "Y") taxable = r2(taxable + v);
    });
    var estTax = r2(taxable * rate), estTotal = r2(sub + estTax);
    var deposit = e.status === "ACCEPTED" ? r2(estTotal * e.deposit_pct / 100) : 0;
    var row = { est_id: e.est_id, customer: e.customer, status: e.status,
      date_sent: e.date_sent, accepted_date: e.accepted_date,
      subtotal: sub, taxableBase: taxable, estTax: estTax, estTotal: estTotal,
      deposit: deposit, deposit_pct: e.deposit_pct,
      waitingDays: e.status === "SENT" ? days(e.date_sent, asOf) : 0 };
    estBy[e.est_id] = row; est.push(row);
  });

  var evBy = {}, coBy = {};
  data.job_events.forEach(function (ev) { (evBy[ev.job_id] = evBy[ev.job_id] || []).push(ev); });
  data.change_orders.forEach(function (c) { (coBy[c.job_id] = coBy[c.job_id] || []).push(c); });
  var invByJob = {};
  data.invoices.forEach(function (i) { invByJob[i.job_id] = i; });   // by JOB ID (v1 matched by customer; that hid every repeat customer's second job)
  var payBy = {};
  data.payments.forEach(function (p) { (payBy[p.inv_id] = payBy[p.inv_id] || []).push(p); });

  var ORDER = { SCHEDULED: 1, STARTED: 2, DONE: 3 };
  var jobs = [], jobByEst = {};
  data.jobs.forEach(function (j) {
    var e = estBy[j.est_id];
    var evs = (evBy[j.job_id] || []).slice().sort(function (a, b) { return a.date < b.date ? -1 : 1; });
    var stage = 0, doneDate = "", startDate = "", lastActivity = j.created_date;
    evs.forEach(function (ev) {
      if (ev.date > lastActivity) lastActivity = ev.date;
      var o = ORDER[ev.event] || 0;              // NOTE -> 0: activity, never a stage
      if (o > stage) stage = o;
      if (ev.event === "DONE" && !doneDate) doneDate = ev.date;
      if (ev.event === "STARTED" && !startDate) startDate = ev.date;
    });
    var cos = coBy[j.job_id] || [];
    var sumApproved = 0, pending = 0;
    cos.forEach(function (c) {
      if (c.approved === "Y") sumApproved = r2(sumApproved + c.amount);
      else if (c.approved !== "N") pending++;
    });
    var invSub = r2(e.subtotal + sumApproved);
    var invTaxBase = r2(e.taxableBase + sumApproved);
    var invTax = r2(invTaxBase * rate);
    var invTotal = r2(invSub + invTax);
    var amountDue = r2(invTotal - e.deposit);
    var row = { job_id: j.job_id, est_id: j.est_id, customer: e.customer,
      created_date: j.created_date, stage: stage, doneDate: doneDate,
      startDate: startDate, lastActivity: lastActivity,
      pendingCOs: pending, approvedCOTotal: sumApproved,
      invSub: invSub, invTax: invTax, invTotal: invTotal,
      deposit: e.deposit, amountDue: amountDue,
      stalled: stage === 2 && days(lastActivity, asOf) > settings.stale_days,
      idleDays: days(lastActivity, asOf),
      invoice: invByJob[j.job_id] ? invByJob[j.job_id].inv_id : "" };
    jobs.push(row); jobByEst[j.est_id] = row;
  });

  var invoices = data.invoices.map(function (i) {
    var j = null;
    for (var k = 0; k < jobs.length; k++) if (jobs[k].job_id === i.job_id) j = jobs[k];
    var paid = 0;
    (payBy[i.inv_id] || []).forEach(function (p) { paid = r2(paid + p.amount); });
    var due = j ? j.amountDue : 0;
    var balance = r2(due - paid);
    var age = days(i.date_sent, asOf);
    var status = balance <= 0.005 ? "PAID"
      : (age > settings.overdue_days ? "OVERDUE" : (paid > 0 ? "PARTIAL" : "SENT"));
    return { inv_id: i.inv_id, job_id: i.job_id, customer: j ? j.customer : "?",
      date_sent: i.date_sent, amountDue: due, paid: paid, balance: balance,
      ageDays: age, status: status,
      overpaid: balance < -0.005,
      lastPay: (payBy[i.inv_id] || []).reduce(function (m, p) { return p.date > m ? p.date : m; }, "") };
  });

  // ---- board ----
  var board = [], exceptions = [];
  est.forEach(function (e) {
    if (e.status === "SENT")
      board.push({ bucket: "WAITING", id: e.est_id, customer: e.customer,
        value: e.estTotal, daysLabel: e.waitingDays });
    if (e.status === "ACCEPTED" && !jobByEst[e.est_id])
      exceptions.push({ kind: "ACCEPTED_NO_JOB", id: e.est_id, customer: e.customer,
        value: e.estTotal, daysLabel: days(e.accepted_date, asOf) });
  });
  jobs.forEach(function (j) {
    if (j.invoice) return;                        // invoiced jobs live on the invoice list
    if (j.pendingCOs > 0)
      board.push({ bucket: "BLOCKED", id: j.job_id, customer: j.customer,
        value: j.amountDue, daysLabel: j.idleDays });
    else if (j.stage === 3)
      board.push({ bucket: "DONE_NOT_INVOICED", id: j.job_id, customer: j.customer,
        value: j.amountDue, daysLabel: days(j.doneDate, asOf) });
    else if (j.stage === 2)
      board.push({ bucket: j.stalled ? "STALLED" : "IN_PROGRESS", id: j.job_id,
        customer: j.customer, value: j.amountDue, daysLabel: j.idleDays });
    else
      board.push({ bucket: "ACCEPTED_NOT_STARTED", id: j.job_id, customer: j.customer,
        value: j.amountDue, daysLabel: days(j.created_date, asOf) });
  });
  invoices.forEach(function (v) {
    if (v.status === "PAID") {
      if (days(v.lastPay || v.date_sent, asOf) <= settings.paid_window_days)
        board.push({ bucket: "PAID", id: v.inv_id, customer: v.customer,
          value: v.amountDue, daysLabel: days(v.lastPay || v.date_sent, asOf) });
    } else {
      board.push({ bucket: "INVOICED", id: v.inv_id, customer: v.customer,
        value: v.balance, daysLabel: v.ageDays,
        flag: v.status === "OVERDUE" ? "OVERDUE" : (v.status === "PARTIAL" ? "PARTIAL" : "") });
    }
    if (v.overpaid)
      exceptions.push({ kind: "OVERPAID", id: v.inv_id, customer: v.customer,
        value: r2(-v.balance), daysLabel: v.ageDays });
  });

  function sumBucket(name) {
    var t = 0;
    board.forEach(function (b) { if (b.bucket === name) t = r2(t + b.value); });
    return t;
  }
  var totals = {
    waiting: sumBucket("WAITING"),
    unbilled: sumBucket("DONE_NOT_INVOICED"),
    blocked: sumBucket("BLOCKED"),
    inProgress: r2(sumBucket("IN_PROGRESS") + sumBucket("STALLED")),
    outstanding: sumBucket("INVOICED"),
    overdue: 0, overdueCount: 0, unbilledCount: 0, oldestUnbilledDays: 0
  };
  board.forEach(function (b) {
    if (b.bucket === "INVOICED" && b.flag === "OVERDUE") { totals.overdue = r2(totals.overdue + b.value); totals.overdueCount++; }
    if (b.bucket === "DONE_NOT_INVOICED") { totals.unbilledCount++; if (b.daysLabel > totals.oldestUnbilledDays) totals.oldestUnbilledDays = b.daysLabel; }
  });

  // ---- reconciliation: every invoice must tie to the cent ----
  var reconOK = true, reconRows = [];
  invoices.forEach(function (v) {
    var tie = r2(v.paid + v.balance);
    var ok = Math.abs(tie - v.amountDue) <= 0.005;
    if (!ok) reconOK = false;
    reconRows.push({ inv_id: v.inv_id, due: v.amountDue, paid: v.paid, balance: v.balance, ok: ok });
  });
  var accepted = est.filter(function (e) { return e.status === "ACCEPTED"; }).length;
  var placed = jobs.length + exceptions.filter(function (x) { return x.kind === "ACCEPTED_NO_JOB"; }).length;
  if (placed !== accepted) reconOK = false;

  return { estimates: est, jobs: jobs, invoices: invoices, board: board,
    exceptions: exceptions, totals: totals,
    reconciliation: { ok: reconOK, rows: reconRows, accepted: accepted, placed: placed } };
}



/* ================= Apps Script glue (sheets in, board out) ================= */

var TABS = {
  estimates: ["est_id","customer","email","date_sent","status","accepted_date","deposit_pct"],
  estimate_lines: ["est_id","item","qty","unit_price","taxable"],
  jobs: ["job_id","est_id","created_date"],
  job_events: ["job_id","date","event"],
  change_orders: ["co_id","job_id","date","description","amount","approved"],
  invoices: ["inv_id","job_id","date_sent"],
  payments: ["pay_id","inv_id","date","amount"]
};
var TAB_TITLES = { estimates: "Estimates", estimate_lines: "Estimate Lines", jobs: "Jobs",
  job_events: "Job Events", change_orders: "Change Orders", invoices: "Invoices", payments: "Payments" };

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Pipeline")
    .addItem("Set up tabs + sample data", "setupSampleData")
    .addItem("Rebuild status board", "rebuildBoard")
    .addItem("Convert accepted estimates to jobs", "convertAccepted")
    .addItem("Create invoices for ready jobs", "createInvoices")
    .addToUi();
}

function isoToday_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}
function cellStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v);
}

function readSettings_(ss) {
  var sh = ss.getSheetByName("Settings");
  var s = { tax_rate_pct: 7.4, overdue_days: 30, stale_days: 14, paid_window_days: 30, as_of: "" };
  if (!sh) return s;
  var vals = sh.getDataRange().getValues();
  for (var i = 0; i < vals.length; i++) {
    var k = String(vals[i][0]).trim(), v = vals[i][1];
    if (k === "Sales tax rate %") s.tax_rate_pct = Number(v);
    if (k === "Invoice overdue after (days)") s.overdue_days = Number(v);
    if (k === "Job stalled after (days)") s.stale_days = Number(v);
    if (k === "Show paid from last (days)") s.paid_window_days = Number(v);
    if (k === "As-of date (blank = today)" && v) s.as_of = cellStr_(v);
  }
  return s;
}

function readTable_(ss, key) {
  var sh = ss.getSheetByName(TAB_TITLES[key]);
  if (!sh) throw new Error("Missing tab: " + TAB_TITLES[key]);
  var vals = sh.getDataRange().getValues();
  var head = vals[0].map(function (h) { return String(h).trim().toLowerCase().replace(/ /g, "_"); });
  var rows = [];
  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][0]) === "") continue;
    var o = {};
    for (var c = 0; c < head.length; c++) {
      var v = vals[r][c];
      if (["qty","unit_price","amount","deposit_pct"].indexOf(head[c]) >= 0) v = Number(v);
      else v = cellStr_(v);
      o[head[c]] = v;
    }
    rows.push(o);
  }
  return rows;
}

function readData_(ss, settings) {
  return {
    as_of: settings.as_of || isoToday_(),
    estimates: readTable_(ss, "estimates"),
    estimate_lines: readTable_(ss, "estimate_lines"),
    jobs: readTable_(ss, "jobs"),
    job_events: readTable_(ss, "job_events"),
    change_orders: readTable_(ss, "change_orders"),
    invoices: readTable_(ss, "invoices"),
    payments: readTable_(ss, "payments")
  };
}

var BUCKET_TITLES = {
  WAITING: "WAITING ON CUSTOMER (estimate out)",
  ACCEPTED_NOT_STARTED: "ACCEPTED - NOT STARTED",
  IN_PROGRESS: "IN PROGRESS",
  STALLED: "STALLED - no activity past the limit",
  BLOCKED: "BLOCKED - change order awaiting approval",
  DONE_NOT_INVOICED: "DONE - NEVER INVOICED",
  INVOICED: "INVOICED - AWAITING PAYMENT",
  PAID: "PAID (recent)"
};
var BUCKET_ORDER = ["DONE_NOT_INVOICED","BLOCKED","STALLED","INVOICED","IN_PROGRESS",
                    "ACCEPTED_NOT_STARTED","WAITING","PAID"];

function rebuildBoard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = readSettings_(ss);
  var data = readData_(ss, settings);
  var rep = computePipeline(data, settings);
  var sh = ss.getSheetByName("Status Board") || ss.insertSheet("Status Board");
  sh.clearContents();
  var rows = [["SADDLE RIDGE STATUS BOARD", "as of " + data.as_of, "", ""]];
  rows.push(["Finished, never invoiced", rep.totals.unbilled,
             rep.totals.unbilledCount + " job(s), oldest " + rep.totals.oldestUnbilledDays + " days", ""]);
  rows.push(["Overdue invoices", rep.totals.overdue, rep.totals.overdueCount + " invoice(s)", ""]);
  rows.push(["Blocked by change orders", rep.totals.blocked, "", ""]);
  rows.push(["Out as estimates", rep.totals.waiting, "", ""]);
  rows.push(["", "", "", ""]);
  BUCKET_ORDER.forEach(function (bk) {
    var items = rep.board.filter(function (b) { return b.bucket === bk; })
      .sort(function (a, b) { return b.value - a.value; });
    if (!items.length) return;
    rows.push([BUCKET_TITLES[bk], "", "", ""]);
    rows.push(["who", "value", "days", "flag"]);
    items.forEach(function (b) {
      rows.push([b.customer + "  (" + b.id + ")", b.value, b.daysLabel, b.flag || ""]);
    });
    rows.push(["", "", "", ""]);
  });
  if (rep.exceptions.length) {
    rows.push(["NEEDS A DECISION", "", "", ""]);
    rep.exceptions.forEach(function (x) {
      rows.push([x.customer + "  (" + x.id + ")", x.value, x.daysLabel, x.kind]);
    });
    rows.push(["", "", "", ""]);
  }
  rows.push(["Reconciliation", rep.reconciliation.ok ? "TIES OUT" : "DOES NOT TIE - do not trust this board",
             rep.reconciliation.placed + "/" + rep.reconciliation.accepted + " accepted estimates placed", ""]);
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  return rep;
}

function convertAccepted() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = readSettings_(ss);
  var data = readData_(ss, settings);
  var have = {};
  data.jobs.forEach(function (j) { have[j.est_id] = true; });
  var maxN = 2000;
  data.jobs.forEach(function (j) {
    var n = Number(String(j.job_id).replace(/\D/g, "")); if (n > maxN) maxN = n;
  });
  var sh = ss.getSheetByName(TAB_TITLES.jobs);
  var added = 0;
  data.estimates.forEach(function (e) {
    if (e.status !== "ACCEPTED" || have[e.est_id]) return;
    maxN++; added++;
    sh.appendRow(["J-" + maxN, e.est_id, data.as_of]);
  });
  return added;
}

function createInvoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = readSettings_(ss);
  var data = readData_(ss, settings);
  var rep = computePipeline(data, settings);
  var sh = ss.getSheetByName(TAB_TITLES.invoices);
  var maxN = 5000;
  data.invoices.forEach(function (i) {
    var n = Number(String(i.inv_id).replace(/\D/g, "")); if (n > maxN) maxN = n;
  });
  var made = 0, blocked = 0;
  rep.jobs.forEach(function (j) {
    if (j.invoice || j.stage !== 3) return;
    if (j.pendingCOs > 0) { blocked++; return; }   // never invoice past a pending change order
    maxN++; made++;
    sh.appendRow(["INV-" + maxN, j.job_id, data.as_of]);
  });
  return { made: made, blocked: blocked };
}

function setupSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fx = JSON.parse(SAMPLE_JSON);
  Object.keys(TABS).forEach(function (key) {
    var title = TAB_TITLES[key];
    var sh = ss.getSheetByName(title) || ss.insertSheet(title);
    sh.clearContents();
    var head = TABS[key];
    var rows = [head];
    fx[key].forEach(function (o) {
      rows.push(head.map(function (h) { return o[h] !== undefined ? o[h] : ""; }));
    });
    sh.getRange(1, 1, rows.length, head.length).setValues(rows);
  });
  var st = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  st.clearContents();
  st.getRange(1, 1, 5, 2).setValues([
    ["Sales tax rate %", fx.settings.tax_rate_pct],
    ["Invoice overdue after (days)", fx.settings.overdue_days],
    ["Job stalled after (days)", fx.settings.stale_days],
    ["Show paid from last (days)", fx.settings.paid_window_days],
    ["As-of date (blank = today)", fx.as_of]
  ]);
}



var SAMPLE_JSON = "{\"as_of\":\"2026-10-03\",\"settings\":{\"tax_rate_pct\":7.4,\"overdue_days\":30,\"stale_days\":14,\"paid_window_days\":30},\"estimates\":[{\"est_id\":\"E-1001\",\"customer\":\"Stu Verhoeven\",\"email\":\"stu.verhoeven@example.com\",\"date_sent\":\"2026-07-15\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-07-25\",\"deposit_pct\":50},{\"est_id\":\"E-1002\",\"customer\":\"Nadia Farwell\",\"email\":\"nadia.farwell@example.com\",\"date_sent\":\"2026-03-30\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1003\",\"customer\":\"Dale Ames\",\"email\":\"dale.ames@example.com\",\"date_sent\":\"2026-09-15\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":0},{\"est_id\":\"E-1004\",\"customer\":\"Oscar Kowalczyk\",\"email\":\"oscar.kowalczyk@example.com\",\"date_sent\":\"2026-07-13\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-07-14\",\"deposit_pct\":25},{\"est_id\":\"E-1005\",\"customer\":\"Rhoda Mares\",\"email\":\"rhoda.mares@example.com\",\"date_sent\":\"2026-07-26\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":0},{\"est_id\":\"E-1006\",\"customer\":\"Stu Verhoeven\",\"email\":\"stu.verhoeven@example.com\",\"date_sent\":\"2026-06-12\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-06-22\",\"deposit_pct\":30},{\"est_id\":\"E-1007\",\"customer\":\"Anneke Yoon\",\"email\":\"anneke.yoon@example.com\",\"date_sent\":\"2026-07-21\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-07-26\",\"deposit_pct\":0},{\"est_id\":\"E-1008\",\"customer\":\"Nadia Whitfield\",\"email\":\"nadia.whitfield@example.com\",\"date_sent\":\"2026-05-31\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":50},{\"est_id\":\"E-1009\",\"customer\":\"Anneke Yoon\",\"email\":\"anneke.yoon@example.com\",\"date_sent\":\"2026-04-16\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-04-27\",\"deposit_pct\":30},{\"est_id\":\"E-1010\",\"customer\":\"Mack Whitfield\",\"email\":\"mack.whitfield@example.com\",\"date_sent\":\"2026-07-21\",\"status\":\"DECLINED\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1011\",\"customer\":\"Petra Whitfield\",\"email\":\"petra.whitfield@example.com\",\"date_sent\":\"2026-04-16\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":50},{\"est_id\":\"E-1012\",\"customer\":\"Rhoda Whitfield\",\"email\":\"rhoda.whitfield@example.com\",\"date_sent\":\"2026-09-22\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":50},{\"est_id\":\"E-1013\",\"customer\":\"Priya Farwell\",\"email\":\"priya.farwell@example.com\",\"date_sent\":\"2026-05-03\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-05-09\",\"deposit_pct\":25},{\"est_id\":\"E-1014\",\"customer\":\"Priya Ames\",\"email\":\"priya.ames@example.com\",\"date_sent\":\"2026-03-29\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1015\",\"customer\":\"Priya Abara\",\"email\":\"priya.abara@example.com\",\"date_sent\":\"2026-07-07\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-07-11\",\"deposit_pct\":0},{\"est_id\":\"E-1016\",\"customer\":\"Gideon Kowalczyk\",\"email\":\"gideon.kowalczyk@example.com\",\"date_sent\":\"2026-06-23\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-06-25\",\"deposit_pct\":25},{\"est_id\":\"E-1017\",\"customer\":\"Odile Calloway\",\"email\":\"odile.calloway@example.com\",\"date_sent\":\"2026-09-06\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-09-17\",\"deposit_pct\":25},{\"est_id\":\"E-1018\",\"customer\":\"Garrett Whitfield\",\"email\":\"garrett.whitfield@example.com\",\"date_sent\":\"2026-04-10\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-04-21\",\"deposit_pct\":25},{\"est_id\":\"E-1019\",\"customer\":\"Carlene Strand\",\"email\":\"carlene.strand@example.com\",\"date_sent\":\"2026-08-29\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1020\",\"customer\":\"Marcy Reyes-Soto\",\"email\":\"marcy.reyes-soto@example.com\",\"date_sent\":\"2026-08-12\",\"status\":\"DECLINED\",\"accepted_date\":\"\",\"deposit_pct\":50},{\"est_id\":\"E-1021\",\"customer\":\"Boyd Druitt\",\"email\":\"boyd.druitt@example.com\",\"date_sent\":\"2026-07-17\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-07-25\",\"deposit_pct\":30},{\"est_id\":\"E-1022\",\"customer\":\"Oscar Druitt\",\"email\":\"oscar.druitt@example.com\",\"date_sent\":\"2026-09-28\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-10-02\",\"deposit_pct\":0},{\"est_id\":\"E-1023\",\"customer\":\"Carlene Seversen\",\"email\":\"carlene.seversen@example.com\",\"date_sent\":\"2026-06-28\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":25},{\"est_id\":\"E-1024\",\"customer\":\"Carlene Whitfield\",\"email\":\"carlene.whitfield@example.com\",\"date_sent\":\"2026-06-15\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-06-27\",\"deposit_pct\":50},{\"est_id\":\"E-1025\",\"customer\":\"Lou Ames\",\"email\":\"lou.ames@example.com\",\"date_sent\":\"2026-08-24\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":50},{\"est_id\":\"E-1026\",\"customer\":\"Walt Belliveau\",\"email\":\"walt.belliveau@example.com\",\"date_sent\":\"2026-06-08\",\"status\":\"DECLINED\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1027\",\"customer\":\"Rhoda Strand\",\"email\":\"rhoda.strand@example.com\",\"date_sent\":\"2026-03-25\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-03-28\",\"deposit_pct\":0},{\"est_id\":\"E-1028\",\"customer\":\"Lou Druitt\",\"email\":\"lou.druitt@example.com\",\"date_sent\":\"2026-07-13\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1029\",\"customer\":\"Walt Mares\",\"email\":\"walt.mares@example.com\",\"date_sent\":\"2026-08-09\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":25},{\"est_id\":\"E-1030\",\"customer\":\"Esteban Lindqvist\",\"email\":\"esteban.lindqvist@example.com\",\"date_sent\":\"2026-08-11\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":0},{\"est_id\":\"E-1031\",\"customer\":\"Sabrina Whitfield\",\"email\":\"sabrina.whitfield@example.com\",\"date_sent\":\"2026-07-11\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1032\",\"customer\":\"Boyd Camm\",\"email\":\"boyd.camm@example.com\",\"date_sent\":\"2026-05-01\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-05-09\",\"deposit_pct\":50},{\"est_id\":\"E-1033\",\"customer\":\"Lou Pike\",\"email\":\"lou.pike@example.com\",\"date_sent\":\"2026-08-21\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1034\",\"customer\":\"Marcy Brandt\",\"email\":\"marcy.brandt@example.com\",\"date_sent\":\"2026-05-02\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":0},{\"est_id\":\"E-1035\",\"customer\":\"Mabel Calloway\",\"email\":\"mabel.calloway@example.com\",\"date_sent\":\"2026-09-25\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1036\",\"customer\":\"Fern Ames\",\"email\":\"fern.ames@example.com\",\"date_sent\":\"2026-05-09\",\"status\":\"DECLINED\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1037\",\"customer\":\"Hollis Lindqvist\",\"email\":\"hollis.lindqvist@example.com\",\"date_sent\":\"2026-09-10\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-09-14\",\"deposit_pct\":50},{\"est_id\":\"E-1038\",\"customer\":\"Mack Camm\",\"email\":\"mack.camm@example.com\",\"date_sent\":\"2026-06-10\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1039\",\"customer\":\"Junie Reyes-Soto\",\"email\":\"junie.reyes-soto@example.com\",\"date_sent\":\"2026-07-18\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-07-22\",\"deposit_pct\":25},{\"est_id\":\"E-1040\",\"customer\":\"Mack Reyes-Soto\",\"email\":\"mack.reyes-soto@example.com\",\"date_sent\":\"2026-08-14\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":0},{\"est_id\":\"E-1041\",\"customer\":\"Junie Brandt\",\"email\":\"junie.brandt@example.com\",\"date_sent\":\"2026-03-31\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-04-08\",\"deposit_pct\":30},{\"est_id\":\"E-1042\",\"customer\":\"Stu Salandra\",\"email\":\"stu.salandra@example.com\",\"date_sent\":\"2026-07-05\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-07-09\",\"deposit_pct\":0},{\"est_id\":\"E-1043\",\"customer\":\"Mabel Pettigrew\",\"email\":\"mabel.pettigrew@example.com\",\"date_sent\":\"2026-04-04\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-04-16\",\"deposit_pct\":50},{\"est_id\":\"E-1044\",\"customer\":\"Priya Dinh\",\"email\":\"priya.dinh@example.com\",\"date_sent\":\"2026-07-24\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1045\",\"customer\":\"Clive Dinh\",\"email\":\"clive.dinh@example.com\",\"date_sent\":\"2026-04-12\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-04-16\",\"deposit_pct\":0},{\"est_id\":\"E-1046\",\"customer\":\"Imogen Pettigrew\",\"email\":\"imogen.pettigrew@example.com\",\"date_sent\":\"2026-04-08\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-04-11\",\"deposit_pct\":0},{\"est_id\":\"E-1047\",\"customer\":\"Dale Reyes-Soto\",\"email\":\"dale.reyes-soto@example.com\",\"date_sent\":\"2026-05-18\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-05-30\",\"deposit_pct\":0},{\"est_id\":\"E-1048\",\"customer\":\"Burt Pike\",\"email\":\"burt.pike@example.com\",\"date_sent\":\"2026-08-22\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-09-01\",\"deposit_pct\":25},{\"est_id\":\"E-1049\",\"customer\":\"Petra Brandt\",\"email\":\"petra.brandt@example.com\",\"date_sent\":\"2026-05-04\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":0},{\"est_id\":\"E-1050\",\"customer\":\"Rhoda Okafor\",\"email\":\"rhoda.okafor@example.com\",\"date_sent\":\"2026-03-17\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":0},{\"est_id\":\"E-1051\",\"customer\":\"Junie Belliveau\",\"email\":\"junie.belliveau@example.com\",\"date_sent\":\"2026-09-05\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-09-12\",\"deposit_pct\":30},{\"est_id\":\"E-1052\",\"customer\":\"Mabel Belliveau\",\"email\":\"mabel.belliveau@example.com\",\"date_sent\":\"2026-08-29\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":25},{\"est_id\":\"E-1053\",\"customer\":\"Mack Dinh\",\"email\":\"mack.dinh@example.com\",\"date_sent\":\"2026-03-27\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1054\",\"customer\":\"Carlene Abara\",\"email\":\"carlene.abara@example.com\",\"date_sent\":\"2026-08-30\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-09-03\",\"deposit_pct\":0},{\"est_id\":\"E-1055\",\"customer\":\"Walt Pike\",\"email\":\"walt.pike@example.com\",\"date_sent\":\"2026-07-13\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1056\",\"customer\":\"Stu Abara\",\"email\":\"stu.abara@example.com\",\"date_sent\":\"2026-08-24\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-09-01\",\"deposit_pct\":0},{\"est_id\":\"E-1057\",\"customer\":\"Hollis Verhoeven\",\"email\":\"hollis.verhoeven@example.com\",\"date_sent\":\"2026-08-12\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-08-20\",\"deposit_pct\":0},{\"est_id\":\"E-1058\",\"customer\":\"Ray Mares\",\"email\":\"ray.mares@example.com\",\"date_sent\":\"2026-08-08\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-08-19\",\"deposit_pct\":0},{\"est_id\":\"E-1059\",\"customer\":\"Petra Belliveau\",\"email\":\"petra.belliveau@example.com\",\"date_sent\":\"2026-05-14\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-05-25\",\"deposit_pct\":50},{\"est_id\":\"E-1060\",\"customer\":\"Odile Strand\",\"email\":\"odile.strand@example.com\",\"date_sent\":\"2026-06-10\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-06-14\",\"deposit_pct\":30},{\"est_id\":\"E-1061\",\"customer\":\"Teo Okafor\",\"email\":\"teo.okafor@example.com\",\"date_sent\":\"2026-06-08\",\"status\":\"DRAFT\",\"accepted_date\":\"\",\"deposit_pct\":30},{\"est_id\":\"E-1062\",\"customer\":\"Fern Lindqvist\",\"email\":\"fern.lindqvist@example.com\",\"date_sent\":\"2026-07-29\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-08-10\",\"deposit_pct\":50},{\"est_id\":\"E-1063\",\"customer\":\"Lou Brandt\",\"email\":\"lou.brandt@example.com\",\"date_sent\":\"2026-04-06\",\"status\":\"ACCEPTED\",\"accepted_date\":\"2026-04-12\",\"deposit_pct\":30},{\"est_id\":\"E-1064\",\"customer\":\"Odile Ostrander\",\"email\":\"odile.ostrander@example.com\",\"date_sent\":\"2026-08-01\",\"status\":\"SENT\",\"accepted_date\":\"\",\"deposit_pct\":0}],\"estimate_lines\":[{\"est_id\":\"E-1001\",\"item\":\"Post set in concrete\",\"qty\":1,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1001\",\"item\":\"Gutter guard, per ft\",\"qty\":24,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1001\",\"item\":\"Gate, single walk-through\",\"qty\":1,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1001\",\"item\":\"Pressure wash prep\",\"qty\":3,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1001\",\"item\":\"Gutter guard, per ft\",\"qty\":300,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1002\",\"item\":\"Deck board replacement, per board\",\"qty\":1,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1002\",\"item\":\"Stain and seal, per sq ft\",\"qty\":60,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1002\",\"item\":\"Gutter guard, per ft\",\"qty\":60,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1002\",\"item\":\"Gutter run, per ft\",\"qty\":150,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1002\",\"item\":\"Railing section\",\"qty\":2,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1003\",\"item\":\"Gate, single walk-through\",\"qty\":1,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1003\",\"item\":\"Demolition of old fence, per ft\",\"qty\":220,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1003\",\"item\":\"Gate, single walk-through\",\"qty\":40,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1003\",\"item\":\"Post set in concrete\",\"qty\":20,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1003\",\"item\":\"Stain and seal, per sq ft\",\"qty\":40,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1004\",\"item\":\"Haul-away and disposal\",\"qty\":3,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1004\",\"item\":\"Composite decking, per sq ft\",\"qty\":110,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1004\",\"item\":\"Post set in concrete\",\"qty\":1,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1004\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":110,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1005\",\"item\":\"Composite decking, per sq ft\",\"qty\":150,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1005\",\"item\":\"Pressure wash prep\",\"qty\":80,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1005\",\"item\":\"Haul-away and disposal\",\"qty\":20,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1005\",\"item\":\"Permit filing\",\"qty\":20,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1006\",\"item\":\"Demolition of old fence, per ft\",\"qty\":110,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1006\",\"item\":\"Railing section\",\"qty\":20,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1006\",\"item\":\"Deck board replacement, per board\",\"qty\":1,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1007\",\"item\":\"Gate, single walk-through\",\"qty\":2,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1007\",\"item\":\"Gutter guard, per ft\",\"qty\":60,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1007\",\"item\":\"Demolition of old fence, per ft\",\"qty\":110,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1007\",\"item\":\"Stain and seal, per sq ft\",\"qty\":150,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1008\",\"item\":\"Stain and seal, per sq ft\",\"qty\":150,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1008\",\"item\":\"Downspout\",\"qty\":6,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1009\",\"item\":\"Permit filing\",\"qty\":12,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1009\",\"item\":\"Gutter run, per ft\",\"qty\":300,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1009\",\"item\":\"Gutter guard, per ft\",\"qty\":60,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1010\",\"item\":\"Railing section\",\"qty\":20,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1010\",\"item\":\"Deck board replacement, per board\",\"qty\":1,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1010\",\"item\":\"Composite decking, per sq ft\",\"qty\":85,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1010\",\"item\":\"Gate, single walk-through\",\"qty\":20,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1011\",\"item\":\"Downspout\",\"qty\":2,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1011\",\"item\":\"Gutter guard, per ft\",\"qty\":85,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1011\",\"item\":\"Post set in concrete\",\"qty\":8,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1011\",\"item\":\"Gutter guard, per ft\",\"qty\":220,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1011\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":110,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1011\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":220,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1012\",\"item\":\"Gate, single walk-through\",\"qty\":8,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1012\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":110,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1012\",\"item\":\"Gutter guard, per ft\",\"qty\":220,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1012\",\"item\":\"Deck board replacement, per board\",\"qty\":6,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1012\",\"item\":\"Pressure wash prep\",\"qty\":80,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1013\",\"item\":\"Post set in concrete\",\"qty\":120,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1013\",\"item\":\"Demolition of old fence, per ft\",\"qty\":60,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1014\",\"item\":\"Haul-away and disposal\",\"qty\":1,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1014\",\"item\":\"Railing section\",\"qty\":40,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1014\",\"item\":\"Gutter run, per ft\",\"qty\":110,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1015\",\"item\":\"Gutter run, per ft\",\"qty\":85,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1015\",\"item\":\"Demolition of old fence, per ft\",\"qty\":150,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1015\",\"item\":\"Composite decking, per sq ft\",\"qty\":40,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1015\",\"item\":\"Pressure wash prep\",\"qty\":1,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1016\",\"item\":\"Downspout\",\"qty\":80,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1016\",\"item\":\"Pressure wash prep\",\"qty\":2,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1016\",\"item\":\"Downspout\",\"qty\":40,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1017\",\"item\":\"Gate, single walk-through\",\"qty\":12,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1017\",\"item\":\"Stain and seal, per sq ft\",\"qty\":60,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1017\",\"item\":\"Post set in concrete\",\"qty\":8,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1018\",\"item\":\"Gate, single walk-through\",\"qty\":1,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1018\",\"item\":\"Deck board replacement, per board\",\"qty\":60,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1018\",\"item\":\"Demolition of old fence, per ft\",\"qty\":110,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1019\",\"item\":\"Deck board replacement, per board\",\"qty\":20,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1019\",\"item\":\"Stain and seal, per sq ft\",\"qty\":24,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1019\",\"item\":\"Gutter run, per ft\",\"qty\":24,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1020\",\"item\":\"Gutter run, per ft\",\"qty\":40,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1020\",\"item\":\"Downspout\",\"qty\":12,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1020\",\"item\":\"Haul-away and disposal\",\"qty\":6,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1020\",\"item\":\"Stain and seal, per sq ft\",\"qty\":85,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1020\",\"item\":\"Post set in concrete\",\"qty\":4,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1021\",\"item\":\"Downspout\",\"qty\":1,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1021\",\"item\":\"Gutter run, per ft\",\"qty\":24,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1021\",\"item\":\"Deck board replacement, per board\",\"qty\":120,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1021\",\"item\":\"Deck board replacement, per board\",\"qty\":60,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1022\",\"item\":\"Composite decking, per sq ft\",\"qty\":85,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1022\",\"item\":\"Stain and seal, per sq ft\",\"qty\":220,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1022\",\"item\":\"Composite decking, per sq ft\",\"qty\":150,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1022\",\"item\":\"Deck board replacement, per board\",\"qty\":20,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1022\",\"item\":\"Pressure wash prep\",\"qty\":1,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1022\",\"item\":\"Pressure wash prep\",\"qty\":4,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1023\",\"item\":\"Railing section\",\"qty\":60,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1023\",\"item\":\"Stain and seal, per sq ft\",\"qty\":60,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1023\",\"item\":\"Demolition of old fence, per ft\",\"qty\":150,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1023\",\"item\":\"Gutter guard, per ft\",\"qty\":150,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1023\",\"item\":\"Post set in concrete\",\"qty\":20,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1023\",\"item\":\"Deck board replacement, per board\",\"qty\":20,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1024\",\"item\":\"Stain and seal, per sq ft\",\"qty\":300,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1024\",\"item\":\"Demolition of old fence, per ft\",\"qty\":60,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1024\",\"item\":\"Permit filing\",\"qty\":60,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1025\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":220,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1025\",\"item\":\"Composite decking, per sq ft\",\"qty\":85,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1025\",\"item\":\"Gutter guard, per ft\",\"qty\":110,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1025\",\"item\":\"Composite decking, per sq ft\",\"qty\":150,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1026\",\"item\":\"Haul-away and disposal\",\"qty\":3,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1026\",\"item\":\"Pressure wash prep\",\"qty\":2,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1026\",\"item\":\"Post set in concrete\",\"qty\":120,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1026\",\"item\":\"Stain and seal, per sq ft\",\"qty\":85,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1026\",\"item\":\"Pressure wash prep\",\"qty\":4,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1026\",\"item\":\"Downspout\",\"qty\":20,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1027\",\"item\":\"Permit filing\",\"qty\":80,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1027\",\"item\":\"Pressure wash prep\",\"qty\":120,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1028\",\"item\":\"Stain and seal, per sq ft\",\"qty\":110,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1028\",\"item\":\"Pressure wash prep\",\"qty\":6,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1028\",\"item\":\"Composite decking, per sq ft\",\"qty\":110,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1028\",\"item\":\"Stain and seal, per sq ft\",\"qty\":60,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1028\",\"item\":\"Gate, single walk-through\",\"qty\":1,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1029\",\"item\":\"Pressure wash prep\",\"qty\":4,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1029\",\"item\":\"Stain and seal, per sq ft\",\"qty\":24,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1029\",\"item\":\"Gutter run, per ft\",\"qty\":220,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1030\",\"item\":\"Deck board replacement, per board\",\"qty\":1,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1030\",\"item\":\"Pressure wash prep\",\"qty\":6,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1030\",\"item\":\"Gutter run, per ft\",\"qty\":220,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1030\",\"item\":\"Haul-away and disposal\",\"qty\":8,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1030\",\"item\":\"Pressure wash prep\",\"qty\":40,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1031\",\"item\":\"Gutter run, per ft\",\"qty\":40,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1031\",\"item\":\"Demolition of old fence, per ft\",\"qty\":24,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1032\",\"item\":\"Permit filing\",\"qty\":3,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1032\",\"item\":\"Deck board replacement, per board\",\"qty\":1,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1032\",\"item\":\"Railing section\",\"qty\":120,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1032\",\"item\":\"Railing section\",\"qty\":120,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1033\",\"item\":\"Permit filing\",\"qty\":20,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1033\",\"item\":\"Post set in concrete\",\"qty\":20,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1033\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":150,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1033\",\"item\":\"Post set in concrete\",\"qty\":120,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1034\",\"item\":\"Post set in concrete\",\"qty\":80,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1034\",\"item\":\"Downspout\",\"qty\":2,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1035\",\"item\":\"Gutter guard, per ft\",\"qty\":85,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1035\",\"item\":\"Gutter guard, per ft\",\"qty\":60,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1036\",\"item\":\"Haul-away and disposal\",\"qty\":120,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1036\",\"item\":\"Haul-away and disposal\",\"qty\":20,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1036\",\"item\":\"Post set in concrete\",\"qty\":40,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1037\",\"item\":\"Pressure wash prep\",\"qty\":20,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1037\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":150,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1038\",\"item\":\"Pressure wash prep\",\"qty\":1,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1038\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":60,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1038\",\"item\":\"Gutter run, per ft\",\"qty\":60,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1038\",\"item\":\"Post set in concrete\",\"qty\":60,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1038\",\"item\":\"Gate, single walk-through\",\"qty\":1,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1039\",\"item\":\"Gutter guard, per ft\",\"qty\":85,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1039\",\"item\":\"Demolition of old fence, per ft\",\"qty\":24,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1039\",\"item\":\"Railing section\",\"qty\":1,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1040\",\"item\":\"Railing section\",\"qty\":6,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1040\",\"item\":\"Demolition of old fence, per ft\",\"qty\":85,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1041\",\"item\":\"Pressure wash prep\",\"qty\":20,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1041\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":300,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1041\",\"item\":\"Gutter guard, per ft\",\"qty\":150,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1041\",\"item\":\"Permit filing\",\"qty\":4,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1041\",\"item\":\"Permit filing\",\"qty\":12,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1041\",\"item\":\"Downspout\",\"qty\":6,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1042\",\"item\":\"Demolition of old fence, per ft\",\"qty\":24,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1042\",\"item\":\"Pressure wash prep\",\"qty\":20,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1042\",\"item\":\"Gutter run, per ft\",\"qty\":110,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1042\",\"item\":\"Gutter guard, per ft\",\"qty\":110,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1042\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":60,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1042\",\"item\":\"Haul-away and disposal\",\"qty\":8,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1043\",\"item\":\"Demolition of old fence, per ft\",\"qty\":85,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1043\",\"item\":\"Composite decking, per sq ft\",\"qty\":85,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1043\",\"item\":\"Pressure wash prep\",\"qty\":1,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1043\",\"item\":\"Pressure wash prep\",\"qty\":6,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1044\",\"item\":\"Post set in concrete\",\"qty\":120,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1044\",\"item\":\"Post set in concrete\",\"qty\":8,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1044\",\"item\":\"Demolition of old fence, per ft\",\"qty\":150,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1044\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":24,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1044\",\"item\":\"Permit filing\",\"qty\":6,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1044\",\"item\":\"Gutter guard, per ft\",\"qty\":24,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1045\",\"item\":\"Post set in concrete\",\"qty\":60,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1045\",\"item\":\"Permit filing\",\"qty\":1,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1045\",\"item\":\"Composite decking, per sq ft\",\"qty\":85,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1045\",\"item\":\"Stain and seal, per sq ft\",\"qty\":300,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1045\",\"item\":\"Deck board replacement, per board\",\"qty\":6,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1045\",\"item\":\"Post set in concrete\",\"qty\":12,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1046\",\"item\":\"Post set in concrete\",\"qty\":4,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1046\",\"item\":\"Deck board replacement, per board\",\"qty\":12,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1046\",\"item\":\"Gutter run, per ft\",\"qty\":60,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1047\",\"item\":\"Gutter guard, per ft\",\"qty\":85,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1047\",\"item\":\"Post set in concrete\",\"qty\":3,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1047\",\"item\":\"Haul-away and disposal\",\"qty\":4,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1047\",\"item\":\"Composite decking, per sq ft\",\"qty\":110,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1047\",\"item\":\"Pressure wash prep\",\"qty\":4,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1047\",\"item\":\"Downspout\",\"qty\":120,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1048\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":60,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1048\",\"item\":\"Post set in concrete\",\"qty\":4,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1048\",\"item\":\"Pressure wash prep\",\"qty\":1,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1049\",\"item\":\"Composite decking, per sq ft\",\"qty\":40,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1049\",\"item\":\"Gate, single walk-through\",\"qty\":20,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1050\",\"item\":\"Permit filing\",\"qty\":1,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1050\",\"item\":\"Post set in concrete\",\"qty\":40,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1050\",\"item\":\"Haul-away and disposal\",\"qty\":2,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1050\",\"item\":\"Permit filing\",\"qty\":4,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1051\",\"item\":\"Gutter guard, per ft\",\"qty\":85,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1051\",\"item\":\"Haul-away and disposal\",\"qty\":80,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1051\",\"item\":\"Stain and seal, per sq ft\",\"qty\":85,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1051\",\"item\":\"Haul-away and disposal\",\"qty\":1,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1051\",\"item\":\"Downspout\",\"qty\":80,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1052\",\"item\":\"Haul-away and disposal\",\"qty\":12,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1052\",\"item\":\"Railing section\",\"qty\":60,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1053\",\"item\":\"Gutter run, per ft\",\"qty\":60,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1053\",\"item\":\"Haul-away and disposal\",\"qty\":1,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1053\",\"item\":\"Gate, single walk-through\",\"qty\":1,\"unit_price\":420.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1054\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":110,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1054\",\"item\":\"Haul-away and disposal\",\"qty\":3,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1054\",\"item\":\"Downspout\",\"qty\":3,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1054\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":300,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1054\",\"item\":\"Pressure wash prep\",\"qty\":1,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1055\",\"item\":\"Demolition of old fence, per ft\",\"qty\":60,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1055\",\"item\":\"Deck board replacement, per board\",\"qty\":1,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1055\",\"item\":\"Post set in concrete\",\"qty\":120,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1055\",\"item\":\"Railing section\",\"qty\":1,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1056\",\"item\":\"Stain and seal, per sq ft\",\"qty\":40,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1056\",\"item\":\"Composite decking, per sq ft\",\"qty\":60,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1056\",\"item\":\"Composite decking, per sq ft\",\"qty\":24,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1056\",\"item\":\"Demolition of old fence, per ft\",\"qty\":40,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1057\",\"item\":\"Stain and seal, per sq ft\",\"qty\":150,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1057\",\"item\":\"Railing section\",\"qty\":1,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1057\",\"item\":\"Permit filing\",\"qty\":1,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1058\",\"item\":\"Gutter guard, per ft\",\"qty\":110,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1058\",\"item\":\"Gutter run, per ft\",\"qty\":24,\"unit_price\":11.25,\"taxable\":\"Y\"},{\"est_id\":\"E-1058\",\"item\":\"Composite decking, per sq ft\",\"qty\":300,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1058\",\"item\":\"Cedar privacy fence, per ft\",\"qty\":85,\"unit_price\":38.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1058\",\"item\":\"Demolition of old fence, per ft\",\"qty\":85,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1058\",\"item\":\"Stain and seal, per sq ft\",\"qty\":110,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1059\",\"item\":\"Composite decking, per sq ft\",\"qty\":40,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1059\",\"item\":\"Railing section\",\"qty\":1,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1059\",\"item\":\"Railing section\",\"qty\":60,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1059\",\"item\":\"Deck board replacement, per board\",\"qty\":1,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1060\",\"item\":\"Permit filing\",\"qty\":4,\"unit_price\":85.0,\"taxable\":\"N\"},{\"est_id\":\"E-1060\",\"item\":\"Post set in concrete\",\"qty\":12,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1060\",\"item\":\"Downspout\",\"qty\":4,\"unit_price\":96.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1060\",\"item\":\"Railing section\",\"qty\":6,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1060\",\"item\":\"Deck board replacement, per board\",\"qty\":3,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1061\",\"item\":\"Composite decking, per sq ft\",\"qty\":85,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1061\",\"item\":\"Stain and seal, per sq ft\",\"qty\":150,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1061\",\"item\":\"Deck board replacement, per board\",\"qty\":60,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1061\",\"item\":\"Haul-away and disposal\",\"qty\":120,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1062\",\"item\":\"Railing section\",\"qty\":6,\"unit_price\":185.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1062\",\"item\":\"Haul-away and disposal\",\"qty\":20,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1062\",\"item\":\"Gutter guard, per ft\",\"qty\":220,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1062\",\"item\":\"Gutter guard, per ft\",\"qty\":110,\"unit_price\":8.75,\"taxable\":\"Y\"},{\"est_id\":\"E-1062\",\"item\":\"Post set in concrete\",\"qty\":6,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1062\",\"item\":\"Post set in concrete\",\"qty\":8,\"unit_price\":65.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1063\",\"item\":\"Stain and seal, per sq ft\",\"qty\":110,\"unit_price\":3.6,\"taxable\":\"Y\"},{\"est_id\":\"E-1063\",\"item\":\"Composite decking, per sq ft\",\"qty\":110,\"unit_price\":21.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1063\",\"item\":\"Haul-away and disposal\",\"qty\":120,\"unit_price\":150.0,\"taxable\":\"N\"},{\"est_id\":\"E-1063\",\"item\":\"Demolition of old fence, per ft\",\"qty\":300,\"unit_price\":9.5,\"taxable\":\"Y\"},{\"est_id\":\"E-1064\",\"item\":\"Deck board replacement, per board\",\"qty\":120,\"unit_price\":24.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1064\",\"item\":\"Pressure wash prep\",\"qty\":12,\"unit_price\":220.0,\"taxable\":\"Y\"},{\"est_id\":\"E-1064\",\"item\":\"Permit filing\",\"qty\":1,\"unit_price\":85.0,\"taxable\":\"N\"}],\"jobs\":[{\"job_id\":\"J-2001\",\"est_id\":\"E-1001\",\"created_date\":\"2026-07-26\"},{\"job_id\":\"J-2002\",\"est_id\":\"E-1004\",\"created_date\":\"2026-07-16\"},{\"job_id\":\"J-2003\",\"est_id\":\"E-1006\",\"created_date\":\"2026-06-27\"},{\"job_id\":\"J-2004\",\"est_id\":\"E-1007\",\"created_date\":\"2026-07-29\"},{\"job_id\":\"J-2005\",\"est_id\":\"E-1009\",\"created_date\":\"2026-05-01\"},{\"job_id\":\"J-2006\",\"est_id\":\"E-1013\",\"created_date\":\"2026-05-12\"},{\"job_id\":\"J-2007\",\"est_id\":\"E-1015\",\"created_date\":\"2026-07-13\"},{\"job_id\":\"J-2008\",\"est_id\":\"E-1016\",\"created_date\":\"2026-06-30\"},{\"job_id\":\"J-2009\",\"est_id\":\"E-1017\",\"created_date\":\"2026-09-21\"},{\"job_id\":\"J-2010\",\"est_id\":\"E-1021\",\"created_date\":\"2026-07-26\"},{\"job_id\":\"J-2011\",\"est_id\":\"E-1022\",\"created_date\":\"2026-10-03\"},{\"job_id\":\"J-2012\",\"est_id\":\"E-1024\",\"created_date\":\"2026-06-27\"},{\"job_id\":\"J-2013\",\"est_id\":\"E-1027\",\"created_date\":\"2026-03-31\"},{\"job_id\":\"J-2014\",\"est_id\":\"E-1032\",\"created_date\":\"2026-05-09\"},{\"job_id\":\"J-2015\",\"est_id\":\"E-1037\",\"created_date\":\"2026-09-15\"},{\"job_id\":\"J-2016\",\"est_id\":\"E-1039\",\"created_date\":\"2026-07-23\"},{\"job_id\":\"J-2017\",\"est_id\":\"E-1041\",\"created_date\":\"2026-04-11\"},{\"job_id\":\"J-2018\",\"est_id\":\"E-1042\",\"created_date\":\"2026-07-14\"},{\"job_id\":\"J-2019\",\"est_id\":\"E-1043\",\"created_date\":\"2026-04-21\"},{\"job_id\":\"J-2020\",\"est_id\":\"E-1045\",\"created_date\":\"2026-04-20\"},{\"job_id\":\"J-2021\",\"est_id\":\"E-1046\",\"created_date\":\"2026-04-12\"},{\"job_id\":\"J-2022\",\"est_id\":\"E-1047\",\"created_date\":\"2026-06-01\"},{\"job_id\":\"J-2023\",\"est_id\":\"E-1048\",\"created_date\":\"2026-09-03\"},{\"job_id\":\"J-2024\",\"est_id\":\"E-1051\",\"created_date\":\"2026-09-12\"},{\"job_id\":\"J-2025\",\"est_id\":\"E-1054\",\"created_date\":\"2026-09-08\"},{\"job_id\":\"J-2026\",\"est_id\":\"E-1056\",\"created_date\":\"2026-09-02\"},{\"job_id\":\"J-2027\",\"est_id\":\"E-1057\",\"created_date\":\"2026-08-20\"},{\"job_id\":\"J-2028\",\"est_id\":\"E-1058\",\"created_date\":\"2026-08-20\"},{\"job_id\":\"J-2029\",\"est_id\":\"E-1059\",\"created_date\":\"2026-05-27\"},{\"job_id\":\"J-2030\",\"est_id\":\"E-1060\",\"created_date\":\"2026-06-16\"},{\"job_id\":\"J-2031\",\"est_id\":\"E-1062\",\"created_date\":\"2026-08-14\"},{\"job_id\":\"J-2032\",\"est_id\":\"E-1063\",\"created_date\":\"2026-04-13\"}],\"job_events\":[{\"job_id\":\"J-2001\",\"date\":\"2026-07-27\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2001\",\"date\":\"2026-08-02\",\"event\":\"STARTED\"},{\"job_id\":\"J-2001\",\"date\":\"2026-08-04\",\"event\":\"NOTE\"},{\"job_id\":\"J-2001\",\"date\":\"2026-08-11\",\"event\":\"DONE\"},{\"job_id\":\"J-2001\",\"date\":\"2026-08-13\",\"event\":\"NOTE\"},{\"job_id\":\"J-2002\",\"date\":\"2026-07-19\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2003\",\"date\":\"2026-06-27\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2003\",\"date\":\"2026-07-05\",\"event\":\"STARTED\"},{\"job_id\":\"J-2003\",\"date\":\"2026-07-10\",\"event\":\"DONE\"},{\"job_id\":\"J-2004\",\"date\":\"2026-07-31\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2004\",\"date\":\"2026-08-09\",\"event\":\"STARTED\"},{\"job_id\":\"J-2004\",\"date\":\"2026-08-12\",\"event\":\"NOTE\"},{\"job_id\":\"J-2004\",\"date\":\"2026-08-13\",\"event\":\"DONE\"},{\"job_id\":\"J-2004\",\"date\":\"2026-08-17\",\"event\":\"NOTE\"},{\"job_id\":\"J-2005\",\"date\":\"2026-05-01\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2005\",\"date\":\"2026-05-03\",\"event\":\"STARTED\"},{\"job_id\":\"J-2005\",\"date\":\"2026-05-06\",\"event\":\"NOTE\"},{\"job_id\":\"J-2005\",\"date\":\"2026-05-20\",\"event\":\"DONE\"},{\"job_id\":\"J-2005\",\"date\":\"2026-05-21\",\"event\":\"NOTE\"},{\"job_id\":\"J-2006\",\"date\":\"2026-05-16\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2007\",\"date\":\"2026-07-16\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2007\",\"date\":\"2026-07-25\",\"event\":\"STARTED\"},{\"job_id\":\"J-2007\",\"date\":\"2026-08-02\",\"event\":\"DONE\"},{\"job_id\":\"J-2008\",\"date\":\"2026-07-03\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2008\",\"date\":\"2026-07-10\",\"event\":\"STARTED\"},{\"job_id\":\"J-2008\",\"date\":\"2026-07-12\",\"event\":\"NOTE\"},{\"job_id\":\"J-2008\",\"date\":\"2026-07-15\",\"event\":\"DONE\"},{\"job_id\":\"J-2009\",\"date\":\"2026-09-24\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2010\",\"date\":\"2026-07-30\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2011\",\"date\":\"2026-10-03\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2012\",\"date\":\"2026-06-27\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2012\",\"date\":\"2026-07-02\",\"event\":\"STARTED\"},{\"job_id\":\"J-2012\",\"date\":\"2026-07-08\",\"event\":\"DONE\"},{\"job_id\":\"J-2013\",\"date\":\"2026-04-01\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2013\",\"date\":\"2026-04-04\",\"event\":\"STARTED\"},{\"job_id\":\"J-2013\",\"date\":\"2026-04-06\",\"event\":\"NOTE\"},{\"job_id\":\"J-2013\",\"date\":\"2026-04-07\",\"event\":\"DONE\"},{\"job_id\":\"J-2013\",\"date\":\"2026-04-08\",\"event\":\"NOTE\"},{\"job_id\":\"J-2014\",\"date\":\"2026-05-10\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2014\",\"date\":\"2026-05-16\",\"event\":\"STARTED\"},{\"job_id\":\"J-2014\",\"date\":\"2026-05-23\",\"event\":\"DONE\"},{\"job_id\":\"J-2014\",\"date\":\"2026-05-27\",\"event\":\"NOTE\"},{\"job_id\":\"J-2015\",\"date\":\"2026-09-15\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2015\",\"date\":\"2026-09-17\",\"event\":\"STARTED\"},{\"job_id\":\"J-2015\",\"date\":\"2026-09-24\",\"event\":\"DONE\"},{\"job_id\":\"J-2016\",\"date\":\"2026-07-23\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2016\",\"date\":\"2026-07-29\",\"event\":\"STARTED\"},{\"job_id\":\"J-2016\",\"date\":\"2026-07-30\",\"event\":\"NOTE\"},{\"job_id\":\"J-2016\",\"date\":\"2026-08-07\",\"event\":\"DONE\"},{\"job_id\":\"J-2017\",\"date\":\"2026-04-11\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2017\",\"date\":\"2026-04-17\",\"event\":\"STARTED\"},{\"job_id\":\"J-2017\",\"date\":\"2026-04-18\",\"event\":\"DONE\"},{\"job_id\":\"J-2018\",\"date\":\"2026-07-15\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2018\",\"date\":\"2026-07-17\",\"event\":\"STARTED\"},{\"job_id\":\"J-2018\",\"date\":\"2026-07-26\",\"event\":\"DONE\"},{\"job_id\":\"J-2018\",\"date\":\"2026-07-28\",\"event\":\"NOTE\"},{\"job_id\":\"J-2019\",\"date\":\"2026-04-21\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2019\",\"date\":\"2026-04-25\",\"event\":\"STARTED\"},{\"job_id\":\"J-2019\",\"date\":\"2026-05-07\",\"event\":\"DONE\"},{\"job_id\":\"J-2020\",\"date\":\"2026-04-22\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2020\",\"date\":\"2026-05-01\",\"event\":\"STARTED\"},{\"job_id\":\"J-2020\",\"date\":\"2026-05-05\",\"event\":\"DONE\"},{\"job_id\":\"J-2020\",\"date\":\"2026-05-06\",\"event\":\"NOTE\"},{\"job_id\":\"J-2021\",\"date\":\"2026-04-15\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2021\",\"date\":\"2026-04-22\",\"event\":\"STARTED\"},{\"job_id\":\"J-2021\",\"date\":\"2026-04-27\",\"event\":\"DONE\"},{\"job_id\":\"J-2022\",\"date\":\"2026-06-02\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2022\",\"date\":\"2026-06-03\",\"event\":\"STARTED\"},{\"job_id\":\"J-2023\",\"date\":\"2026-09-03\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2023\",\"date\":\"2026-09-12\",\"event\":\"STARTED\"},{\"job_id\":\"J-2023\",\"date\":\"2026-09-13\",\"event\":\"NOTE\"},{\"job_id\":\"J-2023\",\"date\":\"2026-09-23\",\"event\":\"DONE\"},{\"job_id\":\"J-2023\",\"date\":\"2026-09-28\",\"event\":\"NOTE\"},{\"job_id\":\"J-2024\",\"date\":\"2026-09-13\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2024\",\"date\":\"2026-09-21\",\"event\":\"STARTED\"},{\"job_id\":\"J-2024\",\"date\":\"2026-10-02\",\"event\":\"DONE\"},{\"job_id\":\"J-2024\",\"date\":\"2026-10-03\",\"event\":\"NOTE\"},{\"job_id\":\"J-2025\",\"date\":\"2026-09-12\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2025\",\"date\":\"2026-09-21\",\"event\":\"STARTED\"},{\"job_id\":\"J-2025\",\"date\":\"2026-09-22\",\"event\":\"NOTE\"},{\"job_id\":\"J-2025\",\"date\":\"2026-10-02\",\"event\":\"DONE\"},{\"job_id\":\"J-2026\",\"date\":\"2026-09-02\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2026\",\"date\":\"2026-09-11\",\"event\":\"STARTED\"},{\"job_id\":\"J-2026\",\"date\":\"2026-09-14\",\"event\":\"NOTE\"},{\"job_id\":\"J-2026\",\"date\":\"2026-09-16\",\"event\":\"DONE\"},{\"job_id\":\"J-2027\",\"date\":\"2026-08-24\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2028\",\"date\":\"2026-08-23\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2028\",\"date\":\"2026-08-24\",\"event\":\"STARTED\"},{\"job_id\":\"J-2028\",\"date\":\"2026-08-27\",\"event\":\"DONE\"},{\"job_id\":\"J-2029\",\"date\":\"2026-05-28\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2029\",\"date\":\"2026-05-30\",\"event\":\"STARTED\"},{\"job_id\":\"J-2029\",\"date\":\"2026-06-01\",\"event\":\"NOTE\"},{\"job_id\":\"J-2029\",\"date\":\"2026-06-10\",\"event\":\"DONE\"},{\"job_id\":\"J-2029\",\"date\":\"2026-06-15\",\"event\":\"NOTE\"},{\"job_id\":\"J-2030\",\"date\":\"2026-06-16\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2030\",\"date\":\"2026-06-17\",\"event\":\"STARTED\"},{\"job_id\":\"J-2030\",\"date\":\"2026-06-19\",\"event\":\"DONE\"},{\"job_id\":\"J-2030\",\"date\":\"2026-06-23\",\"event\":\"NOTE\"},{\"job_id\":\"J-2031\",\"date\":\"2026-08-16\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2031\",\"date\":\"2026-08-19\",\"event\":\"STARTED\"},{\"job_id\":\"J-2031\",\"date\":\"2026-08-22\",\"event\":\"DONE\"},{\"job_id\":\"J-2032\",\"date\":\"2026-04-17\",\"event\":\"SCHEDULED\"},{\"job_id\":\"J-2032\",\"date\":\"2026-04-24\",\"event\":\"STARTED\"},{\"job_id\":\"J-2002\",\"date\":\"2026-09-27\",\"event\":\"STARTED\"},{\"job_id\":\"J-2006\",\"date\":\"2026-09-27\",\"event\":\"STARTED\"}],\"change_orders\":[{\"co_id\":\"CO-301\",\"job_id\":\"J-2001\",\"date\":\"2026-08-13\",\"description\":\"Upgrade to black hardware\",\"amount\":240.0,\"approved\":\"Y\"},{\"co_id\":\"CO-302\",\"job_id\":\"J-2005\",\"date\":\"2026-05-21\",\"description\":\"Replace hidden fascia board\",\"amount\":420.0,\"approved\":\"N\"},{\"co_id\":\"CO-303\",\"job_id\":\"J-2011\",\"date\":\"2026-10-03\",\"description\":\"Two extra posts at the slope\",\"amount\":760.0,\"approved\":\"Y\"},{\"co_id\":\"CO-304\",\"job_id\":\"J-2014\",\"date\":\"2026-05-27\",\"description\":\"Rot discovered in rim joist\",\"amount\":270.0,\"approved\":\"Y\"},{\"co_id\":\"CO-305\",\"job_id\":\"J-2016\",\"date\":\"2026-08-07\",\"description\":\"Two extra posts at the slope\",\"amount\":560.0,\"approved\":\"Y\"},{\"co_id\":\"CO-306\",\"job_id\":\"J-2017\",\"date\":\"2026-04-18\",\"description\":\"Two extra posts at the slope\",\"amount\":760.0,\"approved\":\"Y\"},{\"co_id\":\"CO-307\",\"job_id\":\"J-2023\",\"date\":\"2026-09-28\",\"description\":\"Upgrade to black hardware\",\"amount\":240.0,\"approved\":\"N\"},{\"co_id\":\"CO-308\",\"job_id\":\"J-2026\",\"date\":\"2026-09-16\",\"description\":\"Buried irrigation line reroute\",\"amount\":760.0,\"approved\":\"Y\"},{\"co_id\":\"CO-309\",\"job_id\":\"J-2030\",\"date\":\"2026-06-23\",\"description\":\"Two extra posts at the slope\",\"amount\":570.0,\"approved\":\"Y\"},{\"co_id\":\"CO-399\",\"job_id\":\"J-2024\",\"date\":\"2026-10-02\",\"description\":\"Buried irrigation line reroute\",\"amount\":480.0,\"approved\":\"\"}],\"invoices\":[{\"inv_id\":\"INV-5001\",\"job_id\":\"J-2003\",\"date_sent\":\"2026-07-12\"},{\"inv_id\":\"INV-5002\",\"job_id\":\"J-2004\",\"date_sent\":\"2026-08-19\"},{\"inv_id\":\"INV-5003\",\"job_id\":\"J-2008\",\"date_sent\":\"2026-07-20\"},{\"inv_id\":\"INV-5004\",\"job_id\":\"J-2012\",\"date_sent\":\"2026-07-10\"},{\"inv_id\":\"INV-5005\",\"job_id\":\"J-2013\",\"date_sent\":\"2026-04-14\"},{\"inv_id\":\"INV-5006\",\"job_id\":\"J-2014\",\"date_sent\":\"2026-05-30\"},{\"inv_id\":\"INV-5007\",\"job_id\":\"J-2015\",\"date_sent\":\"2026-09-24\"},{\"inv_id\":\"INV-5008\",\"job_id\":\"J-2016\",\"date_sent\":\"2026-09-13\"},{\"inv_id\":\"INV-5009\",\"job_id\":\"J-2017\",\"date_sent\":\"2026-04-23\"},{\"inv_id\":\"INV-5010\",\"job_id\":\"J-2018\",\"date_sent\":\"2026-07-28\"},{\"inv_id\":\"INV-5011\",\"job_id\":\"J-2019\",\"date_sent\":\"2026-05-12\"},{\"inv_id\":\"INV-5012\",\"job_id\":\"J-2020\",\"date_sent\":\"2026-05-08\"},{\"inv_id\":\"INV-5013\",\"job_id\":\"J-2021\",\"date_sent\":\"2026-04-30\"},{\"inv_id\":\"INV-5016\",\"job_id\":\"J-2026\",\"date_sent\":\"2026-09-16\"},{\"inv_id\":\"INV-5017\",\"job_id\":\"J-2028\",\"date_sent\":\"2026-08-29\"},{\"inv_id\":\"INV-5018\",\"job_id\":\"J-2029\",\"date_sent\":\"2026-06-18\"},{\"inv_id\":\"INV-5019\",\"job_id\":\"J-2030\",\"date_sent\":\"2026-06-24\"}],\"payments\":[{\"pay_id\":\"P-7001\",\"inv_id\":\"INV-5001\",\"date\":\"2026-07-26\",\"amount\":3585.34},{\"pay_id\":\"P-7002\",\"inv_id\":\"INV-5002\",\"date\":\"2026-09-14\",\"amount\":3168.3},{\"pay_id\":\"P-7003\",\"inv_id\":\"INV-5003\",\"date\":\"2026-08-14\",\"amount\":9633.78},{\"pay_id\":\"P-7004\",\"inv_id\":\"INV-5004\",\"date\":\"2026-08-01\",\"amount\":3436.05},{\"pay_id\":\"P-7005\",\"inv_id\":\"INV-5005\",\"date\":\"2026-04-21\",\"amount\":35153.6},{\"pay_id\":\"P-7006\",\"inv_id\":\"INV-5006\",\"date\":\"2026-06-14\",\"amount\":24273.17},{\"pay_id\":\"P-7007\",\"inv_id\":\"INV-5007\",\"date\":\"2026-10-03\",\"amount\":5423.7},{\"pay_id\":\"P-7008\",\"inv_id\":\"INV-5009\",\"date\":\"2026-05-09\",\"amount\":15066.45},{\"pay_id\":\"P-7009\",\"inv_id\":\"INV-5010\",\"date\":\"2026-08-02\",\"amount\":10981.99},{\"pay_id\":\"P-7010\",\"inv_id\":\"INV-5013\",\"date\":\"2026-05-27\",\"amount\":525.4},{\"pay_id\":\"P-7011\",\"inv_id\":\"INV-5016\",\"date\":\"2026-10-03\",\"amount\":3318.66},{\"pay_id\":\"P-7012\",\"inv_id\":\"INV-5017\",\"date\":\"2026-09-02\",\"amount\":13012.58},{\"pay_id\":\"P-7013\",\"inv_id\":\"INV-5018\",\"date\":\"2026-06-23\",\"amount\":6534.75},{\"pay_id\":\"P-7014\",\"inv_id\":\"INV-5019\",\"date\":\"2026-07-04\",\"amount\":2613.9}]}";

if (typeof module !== "undefined") module.exports = {
  computePipeline: computePipeline, r2: r2, days: days,
  rebuildBoard: rebuildBoard, convertAccepted: convertAccepted,
  createInvoices: createInvoices, setupSampleData: setupSampleData, onOpen: onOpen
};
