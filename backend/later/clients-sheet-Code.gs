/**
 * Sahithiness clients sheet
 * 1. sahithiness.in sends each client's form details here when they tap "Pay" → saved as "Payment pending".
 * 2. Razorpay sends a webhook when a payment succeeds → the matching client (by phone, else email)
 *    is marked "Active", with payment ID, amount and "Paid till" (+1 month per payment, so renewals extend it).
 * 3. Every morning: clients past "Paid till" become "Expired", and Sahithi gets a list of renewals due soon.
 * Payments that cannot be matched to a client are listed in the "Payments" tab as "UNMATCHED".
 *
 * Setup: paste this file, run setup(), deploy as Web app (Execute as: Me, access: Anyone),
 * then run webhookUrl() and paste the URL it prints into Razorpay → Settings → Webhooks (event: payment.captured).
 */

var TZ = 'Asia/Kolkata';
var TABS = { postnatal: 'Postnatal', pcos: 'PCOS-PCOD' };
var BASE = ['Submitted', 'Status', 'Name', 'Phone', 'Email', 'Gmail for videos', 'Paid till', 'Months paid', 'Last paid on',
  'Last amount (₹)', 'Last payment ID', 'Videos status', 'Health conditions'];
var PAYMENTS = ['Received', 'Payment ID', 'Amount (₹)', 'Phone', 'Email', 'Matched to', 'Status'];

/* ============================== web endpoints ============================== */

function doGet() { return ContentService.createTextOutput('Sahithiness clients sheet is running.'); }

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false }); }
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (e.parameter && e.parameter.token) {
      if (e.parameter.token !== prop_('WEBHOOK_TOKEN')) return out_({ ok: false });
      return out_(onRazorpay_(body));
    }
    if (body.action === 'client') return out_(saveClient_(body));
    return out_({ ok: false });
  } catch (err) {
    console.error(err && err.stack || err);
    return out_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ============================== 1. client details from the website ============================== */

function saveClient_(d) {
  var tabName = TABS[d.kind];
  if (!tabName) return { ok: false };
  var phone = phone_(d.phone);
  if (!phone) return { ok: false };
  var answers = Array.isArray(d.answers) ? d.answers.slice(0, 30) : [];
  var sh = tab_(tabName, BASE);
  answers.forEach(function (a) { ensureCol_(sh, clean_(a[0], 80)); });

  var row = findClient_(sh, phone, '');
  var v = {};
  v['Name'] = clean_(d.name, 80); v['Phone'] = "'" + phone; v['Email'] = clean_(d.email, 120).toLowerCase();
  v['Gmail for videos'] = clean_(d.gmail, 120).toLowerCase();
  v['Health conditions'] = clean_(Array.isArray(d.flags) && d.flags.length ? d.flags.join(', ') : 'None', 400);
  answers.forEach(function (a) { v[clean_(a[0], 80)] = clean_(a[1], 500); });

  if (row) {                         // returning client or second tap: update details, keep payment history
    v['Submitted'] = new Date();
    setRow_(sh, row, v);
  } else {
    v['Submitted'] = new Date(); v['Status'] = 'Payment pending'; v['Months paid'] = 0; v['Videos status'] = '';
    appendRow_(sh, v);
  }
  return { ok: true };
}

/* ============================== 2. payment from Razorpay ============================== */

function onRazorpay_(body) {
  if (body.event !== 'payment.captured') return { ok: true, ignored: body.event };
  var p = body.payload && body.payload.payment && body.payload.payment.entity;
  if (!p || !p.id) return { ok: false };

  var pay = tab_('Payments', PAYMENTS);
  var ids = pay.getRange(1, 2, Math.max(pay.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < ids.length; i++) if (ids[i][0] === p.id) return { ok: true, duplicate: true };  // Razorpay may resend

  var phone = phone_(p.contact), email = String(p.email || '').toLowerCase(), amount = (p.amount || 0) / 100;
  var match = null;
  Object.keys(TABS).forEach(function (k) {
    var sh = tab_(TABS[k], BASE), r = findClient_(sh, phone, email);
    if (r) {
      var when = new Date(sh.getRange(r, col_(sh, 'Submitted')).getValue()).getTime();
      if (!match || when > match.when) match = { sh: sh, row: r, when: when, tab: TABS[k] };
    }
  });

  var status = 'UNMATCHED';
  if (match) {
    var sh = match.sh, r = match.row;
    var till = sh.getRange(r, col_(sh, 'Paid till')).getValue();
    var start = till && new Date(till) > new Date() ? new Date(till) : new Date();
    var newTill = addMonth_(start);
    var v = {};
    v['Status'] = 'Active'; v['Paid till'] = newTill; v['Last paid on'] = new Date(); v['Last amount (₹)'] = amount;
    v['Last payment ID'] = p.id; v['Months paid'] = Number(sh.getRange(r, col_(sh, 'Months paid')).getValue() || 0) + 1;
    if (!sh.getRange(r, col_(sh, 'Videos status')).getValue()) v['Videos status'] = 'To send';
    setRow_(sh, r, v);
    status = 'Matched';
    alert_('Payment received: ' + sh.getRange(r, col_(sh, 'Name')).getValue() + ' · ' + match.tab + ' · ₹' + amount,
      'Paid till ' + fmt_(newTill) + '\nPhone: ' + phone + '\nPayment ID: ' + p.id + '\n\nOpen the sheet: ' + sheetUrl_());
  } else {
    alert_('[CHECK] Payment of ₹' + amount + ' could not be matched to a client',
      'Phone: ' + (phone || p.contact) + '\nEmail: ' + email + '\nPayment ID: ' + p.id +
      '\n\nThe client may have used a different phone number. Find them in the sheet and update their row: ' + sheetUrl_());
  }
  pay.appendRow([new Date(), p.id, amount, "'" + (phone || p.contact || ''), email, match ? match.tab : '', status]);
  return { ok: true, status: status };
}

/* ============================== 3. daily check (set up by setup()) ============================== */

function dailyCheck() {
  var today = new Date(), soon = new Date(today.getTime() + 3 * 86400000), due = [];
  Object.keys(TABS).forEach(function (k) {
    var sh = tab_(TABS[k], BASE), rows = sh.getDataRange().getValues(), h = rows[0];
    var cS = h.indexOf('Status'), cT = h.indexOf('Paid till'), cN = h.indexOf('Name'), cP = h.indexOf('Phone');
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][cS] !== 'Active' || !rows[i][cT]) continue;
      var till = new Date(rows[i][cT]);
      if (till < today) sh.getRange(i + 1, cS + 1).setValue('Expired');
      else if (till <= soon) due.push(rows[i][cN] + ' (' + TABS[k] + ', ' + rows[i][cP] + ') · renewal due ' + fmt_(till));
    }
  });
  if (due.length) alert_('Renewals due in the next 3 days: ' + due.length, due.join('\n') + '\n\nPayment link: https://rzp.io/rzp/Mb2OjID');
}

/* ============================== setup ============================== */

function setup() {
  var props = PropertiesService.getScriptProperties();
  var ss, id = props.getProperty('SHEET_ID');
  try { ss = id && SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  if (!ss) { ss = SpreadsheetApp.create('Sahithiness Clients'); props.setProperty('SHEET_ID', ss.getId()); }
  tab_(TABS.postnatal, BASE); tab_(TABS.pcos, BASE); tab_('Payments', PAYMENTS);
  var s1 = ss.getSheetByName('Sheet1'); if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  if (!props.getProperty('WEBHOOK_TOKEN')) props.setProperty('WEBHOOK_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyCheck') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(9).create();
  console.log('Clients sheet: ' + ss.getUrl() + '\nAlerts go to: ' + notify_() + '\nNext: Deploy → New deployment → Web app, then run webhookUrl().');
}

/** Run after deploying. Prints the webhook URL to paste into Razorpay. */
function webhookUrl() {
  var url = ScriptApp.getService().getUrl();
  if (!url) { console.log('Deploy the Web app first (Deploy → New deployment → Web app).'); return; }
  console.log('Website SHEET_URL (send this to Claude):\n' + url + '\n\nRazorpay webhook URL (Settings → Webhooks → Add, event payment.captured):\n' + url + '?token=' + prop_('WEBHOOK_TOKEN'));
}

/* ============================== helpers ============================== */

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function notify_() { return prop_('NOTIFY_EMAIL') || Session.getEffectiveUser().getEmail(); }
function sheetUrl_() { return 'https://docs.google.com/spreadsheets/d/' + prop_('SHEET_ID'); }
function alert_(subject, body) { try { MailApp.sendEmail(notify_(), subject, body, { name: 'Sahithiness website' }); } catch (e) { console.error(e); } }

function tab_(name, headers) {
  var ss = SpreadsheetApp.openById(prop_('SHEET_ID')), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#DCE5D3');
    sh.setFrozenRows(1);
  }
  return sh;
}
function headers_(sh) { return sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]; }
function col_(sh, name) { return headers_(sh).indexOf(name) + 1; }
function ensureCol_(sh, name) {
  if (!name || col_(sh, name)) return;
  var c = sh.getLastColumn() + 1;
  sh.getRange(1, c).setValue(name).setFontWeight('bold').setBackground('#F3EFE8');
}
function appendRow_(sh, v) { var h = headers_(sh); sh.appendRow(h.map(function (k) { return v[k] === undefined ? '' : v[k]; })); }
function setRow_(sh, r, v) { var h = headers_(sh); Object.keys(v).forEach(function (k) { var c = h.indexOf(k); if (c >= 0) sh.getRange(r, c + 1).setValue(v[k]); }); }

function findClient_(sh, phone, email) {
  var rows = sh.getDataRange().getValues(), h = rows[0], cP = h.indexOf('Phone'), cE = h.indexOf('Email');
  for (var i = rows.length - 1; i > 0; i--) if (phone && phone_(rows[i][cP]) === phone) return i + 1;
  for (var j = rows.length - 1; j > 0; j--) if (email && String(rows[j][cE]).toLowerCase() === email) return j + 1;
  return 0;
}

function phone_(v) { var d = String(v || '').replace(/\D/g, '').slice(-10); return /^[6-9]\d{9}$/.test(d) ? d : ''; }
function addMonth_(d) { var n = new Date(d); n.setMonth(n.getMonth() + 1); return n; }
function fmt_(d) { return Utilities.formatDate(new Date(d), TZ, 'd MMM yyyy'); }
function clean_(v, max) { var s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B-\u001F]/g, '').trim().slice(0, max); return /^[=+\-@]/.test(s) ? "'" + s : s; }
