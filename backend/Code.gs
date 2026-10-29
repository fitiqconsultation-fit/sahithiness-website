/**
 * Sahithiness backend for sahithiness.in
 * Takes Razorpay payments and delivers every program automatically:
 *   - 1:1 Consultation: Google Calendar event + Meet link, slot blocked, ₹999 credit code (48 hours)
 *   - General Yoga: Google Drive video folder shared for 90 days, removed automatically after
 *   - Prenatal / Postnatal: welcome email to the client, full health details to Sahithi
 * Every order is saved in the "Sahithiness Orders" Google Sheet.
 *
 * Script properties (Project Settings → Script properties):
 *   RAZORPAY_KEY_ID      rzp_test_... (later rzp_live_...)
 *   RAZORPAY_KEY_SECRET  the matching key secret
 *   VIDEO_FOLDER_ID      ID of the Drive folder with the 45 General Yoga videos
 *   NOTIFY_EMAIL         where new-order alerts go (optional, defaults to this Google account)
 * Then run setup() once and deploy as a Web app (Execute as: Me, Who has access: Anyone).
 */

var CFG = {
  brand: 'Sahithiness',
  site: 'https://sahithiness.in',
  whatsapp: '+91 99801 84295',
  whatsappLink: 'https://wa.me/919980184295',
  tz: 'Asia/Kolkata',
  consultMinutes: 30,
  creditPaise: 99900,     // ₹999 consultation credit
  creditHours: 48,        // valid for 48 hours after the call ends
  generalDays: 90,        // General Yoga access period
  firstSlot: 14 * 60,     // 2:00 PM
  lastSlot: 22 * 60 + 30, // last call starts 10:30 PM
  daysAhead: 60,
  leadMinutes: 60,        // no bookings within the next hour
  holdMinutes: 15         // a slot is held this long while someone is paying
};

// Prices are set here, not in the website, so nobody can change what they pay.
var PROGRAMS = {
  prenatal:  { name: 'Prenatal Movement + Nutrition', paise: 3600000, credit: true },
  postnatal: { name: 'Postnatal Movement + Nutrition', paise: 3600000, credit: true },
  general:   { name: 'General Yoga · 45 recorded classes · 90 days', paise: 350000, credit: true },
  consult:   { name: '1:1 Consultation Call · 30 minutes', paise: 99900, credit: false }
};

var ORDERS = ['Created', 'Order ID', 'Status', 'Program', 'Kind', 'Amount (₹)', 'Credit code used', 'Name', 'Phone', 'Email',
  'Gmail for videos', 'Consult slot', 'Payment ID', 'Paid at', 'Delivered', 'Meet link', 'Code issued', 'Health flags', 'Details'];
var C = {}; ORDERS.forEach(function (h, i) { C[h] = i + 1; });
var CREDITS = ['Code', 'Issued', 'Expires', 'Consult order', 'Name', 'Phone', 'Email', 'Used by order', 'Used at'];
var ACCESS = ['Order ID', 'Name', 'Gmail', 'Granted', 'Expires', 'Status'];

/* ============================== web endpoints ============================== */

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  try {
    if (action === 'slots') return json_({ ok: true, taken: takenSlots_() });
    if (action === 'ping') return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Something went wrong. Please try again.' });
  }
  return ContentService.createTextOutput('Sahithiness API is running.');
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  try {
    if (d.action === 'quote') return json_(quote_(d));
    if (d.action === 'createOrder') return json_(createOrder_(d));
    if (d.action === 'verify') return json_(verify_(d));
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    console.error(err && err.stack || err);
    return json_({ ok: false, error: err && err.userMessage ? err.userMessage : 'Something went wrong. Please try again or WhatsApp us on ' + CFG.whatsapp + '.' });
  }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function userError_(msg) { var e = new Error(msg); e.userMessage = msg; return e; }

/* ============================== pricing and credit codes ============================== */

function quote_(d) {
  var p = PROGRAMS[d.kind];
  if (!p) return { ok: false, error: 'Unknown program.' };
  var amount = p.paise, applied = false;
  if (d.code) {
    if (!p.credit) return { ok: false, error: 'Codes can be used on programs only.' };
    var st = creditStatus_(d.code);
    if (!st.ok) return st;
    amount -= CFG.creditPaise; applied = true;
  }
  return { ok: true, amount: amount, full: p.paise, applied: applied };
}

function normCode_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9-]/g, '').slice(0, 20); }

function creditStatus_(code) {
  code = normCode_(code);
  var sh = sheet_('Credits', CREDITS), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== code) continue;
    if (rows[i][7]) return { ok: false, error: 'This code has already been used.' };
    if (new Date() > new Date(rows[i][2])) return { ok: false, error: 'This code has expired. Codes are valid for 48 hours after your call.' };
    return { ok: true, row: i + 1, code: code };
  }
  return { ok: false, error: 'This code is not valid. Please check and try again.' };
}

function newCode_() {
  var abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', existing = {}, code;
  sheet_('Credits', CREDITS).getDataRange().getValues().forEach(function (r) { existing[r[0]] = 1; });
  do {
    code = 'SAHI-';
    for (var i = 0; i < 6; i++) code += abc.charAt(Math.floor(Math.random() * abc.length));
  } while (existing[code]);
  return code;
}

/* ============================== consultation slots ============================== */

function slotKey_(date) { return Utilities.formatDate(date, CFG.tz, 'yyyy-MM-dd HH:mm'); }

function parseSlot_(key) {
  if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(String(key))) throw userError_('Please choose a date and time for your call.');
  var dt = Utilities.parseDate(key, CFG.tz, 'yyyy-MM-dd HH:mm');
  var mins = Number(key.slice(11, 13)) * 60 + Number(key.slice(14, 16));
  var now = Date.now();
  if (mins < CFG.firstSlot || mins > CFG.lastSlot || mins % 30) throw userError_('Please choose a time between 2 PM and 11 PM.');
  if (dt.getTime() < now + CFG.leadMinutes * 60000) throw userError_('That time has passed. Please choose a later slot.');
  if (dt.getTime() > now + (CFG.daysAhead + 1) * 86400000) throw userError_('Please choose a date within the next 60 days.');
  return dt;
}

function takenSlots_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('taken');
  if (hit) return JSON.parse(hit);
  var taken = {}, now = new Date(), until = new Date(now.getTime() + (CFG.daysAhead + 1) * 86400000);
  // paid bookings and slots currently being paid for
  var rows = sheet_('Orders', ORDERS).getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[C['Kind'] - 1] !== 'consult' || !r[C['Consult slot'] - 1]) continue;
    var st = r[C['Status'] - 1];
    var holding = st === 'Created' && now - new Date(r[0]) < CFG.holdMinutes * 60000;
    if (st === 'Paid' || holding) taken[String(r[C['Consult slot'] - 1])] = 1;
  }
  // anything else in Sahithi's calendar
  CalendarApp.getDefaultCalendar().getEvents(now, until).forEach(function (ev) {
    if (ev.isAllDayEvent()) return;
    var s = ev.getStartTime().getTime(), e = ev.getEndTime().getTime();
    var t = new Date(Math.floor(s / 1800000) * 1800000);
    for (; t.getTime() < e; t = new Date(t.getTime() + 1800000)) taken[slotKey_(t)] = 1;
  });
  var list = Object.keys(taken);
  cache.put('taken', JSON.stringify(list), 60);
  return list;
}

/* ============================== orders ============================== */

function createOrder_(d) {
  var p = PROGRAMS[d.kind];
  if (!p) throw userError_('Unknown program.');
  var name = clean_(d.name, 80), email = String(d.email || '').trim().toLowerCase(), phone = String(d.phone || '').replace(/\D/g, '').slice(-10);
  if (name.length < 2) throw userError_('Please enter your full name.');
  if (!/^[6-9]\d{9}$/.test(phone)) throw userError_('Please enter a valid 10-digit mobile number.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw userError_('Please enter a valid email address.');
  var gmail = '';
  if (d.kind === 'general') {
    gmail = String(d.gmail || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(gmail)) throw userError_('Please enter the Gmail ID you will use to watch the videos.');
  }
  var slot = '';
  if (d.kind === 'consult') { parseSlot_(d.slot); slot = d.slot; }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (slot) {
      CacheService.getScriptCache().remove('taken');
      if (takenSlots_().indexOf(slot) >= 0) throw userError_('Sorry, this time was just booked. Please choose another slot.');
    }
    var amount = p.paise, code = '';
    if (d.code && p.credit) {
      var st = creditStatus_(d.code);
      if (!st.ok) throw userError_(st.error);
      amount -= CFG.creditPaise; code = st.code;
    }
    var order = rzp_('post', '/orders', {
      amount: amount, currency: 'INR', receipt: 'SAH-' + Date.now(), payment_capture: 1,
      notes: { program: d.kind, name: name, phone: phone, slot: slot }
    });
    var details = [];
    if (d.details && typeof d.details === 'object') {
      Object.keys(d.details).slice(0, 40).forEach(function (k) { details.push(clean_(k, 80) + ': ' + clean_(d.details[k], 500)); });
    }
    var flags = Array.isArray(d.flags) ? d.flags.slice(0, 20).map(function (f) { return clean_(f, 80); }).join(', ') : '';
    var row = {};
    row['Created'] = new Date(); row['Order ID'] = order.id; row['Status'] = 'Created'; row['Program'] = p.name; row['Kind'] = d.kind;
    row['Amount (₹)'] = amount / 100; row['Credit code used'] = code; row['Name'] = name; row['Phone'] = "'" + phone; row['Email'] = email;
    row['Gmail for videos'] = gmail; row['Consult slot'] = slot; row['Health flags'] = flags || 'None'; row['Details'] = details.join('\n');
    sheet_('Orders', ORDERS).appendRow(ORDERS.map(function (h) { return row[h] === undefined ? '' : row[h]; }));
    CacheService.getScriptCache().remove('taken');
  } finally {
    lock.releaseLock();
  }
  return {
    ok: true, keyId: prop_('RAZORPAY_KEY_ID'), orderId: order.id, amount: amount, currency: 'INR',
    name: CFG.brand, description: p.name, prefill: { name: name, email: email, contact: '+91' + phone }
  };
}

function verify_(d) {
  var orderId = String(d.orderId || ''), paymentId = String(d.paymentId || ''), sig = String(d.signature || '');
  if (!/^order_\w+$/.test(orderId) || !/^pay_\w+$/.test(paymentId)) throw userError_('Payment details missing.');
  var expected = hex_(Utilities.computeHmacSha256Signature(orderId + '|' + paymentId, prop_('RAZORPAY_KEY_SECRET')));
  if (expected !== sig) throw userError_('We could not verify this payment. If money was debited, WhatsApp us on ' + CFG.whatsapp + '.');
  confirmPayment_(orderId, paymentId);
  return deliver_(orderId, paymentId);
}

function confirmPayment_(orderId, paymentId) {
  var pay = rzp_('get', '/payments/' + paymentId);
  if (pay.order_id !== orderId) throw userError_('Payment does not match this order.');
  if (pay.status === 'authorized') pay = rzp_('post', '/payments/' + paymentId + '/capture', { amount: pay.amount, currency: pay.currency });
  if (pay.status !== 'captured') throw userError_('Payment is not complete yet. If money was debited, it will be confirmed by email within 15 minutes.');
  return pay;
}

/** Marks an order paid and delivers it. Safe to call more than once. */
function deliver_(orderId, paymentId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_('Orders', ORDERS), r = findRow_(sh, orderId);
    if (!r) throw userError_('Order not found. Please WhatsApp us on ' + CFG.whatsapp + '.');
    var o = readRow_(sh, r);
    if (o['Status'] === 'Paid') return result_(o);

    o['Status'] = 'Paid'; o['Payment ID'] = paymentId; o['Paid at'] = new Date();
    setCells_(sh, r, o, ['Status', 'Payment ID', 'Paid at']);
    if (o['Credit code used']) markCreditUsed_(o['Credit code used'], orderId);

    try {
      if (o['Kind'] === 'consult') deliverConsult_(o);
      else if (o['Kind'] === 'general') deliverGeneral_(o);
      o['Delivered'] = 'Yes';
    } catch (err) {
      console.error(err && err.stack || err);
      o['Delivered'] = 'ERROR: ' + (err && err.message || err);
    }
    setCells_(sh, r, o, ['Delivered', 'Meet link', 'Code issued']);
    CacheService.getScriptCache().remove('taken');

    try { customerEmail_(o); } catch (err) { console.error(err); }
    try { adminEmail_(o); } catch (err) { console.error(err); }
    return result_(o);
  } finally {
    lock.releaseLock();
  }
}

function result_(o) {
  var res = { ok: true, kind: o['Kind'], name: o['Name'], email: o['Email'], amount: o['Amount (₹)'], orderId: o['Order ID'] };
  if (o['Kind'] === 'consult') {
    res.slot = slotLabel_(o['Consult slot']); res.meetLink = o['Meet link'];
    res.code = o['Code issued']; res.codeExpires = o['Code issued'] ? fmt_(creditExpiry_(o['Consult slot'])) : '';
  }
  if (o['Kind'] === 'general') res.gmail = o['Gmail for videos'];
  if (String(o['Delivered']).indexOf('ERROR') === 0) res.pending = true;
  return res;
}

/* ============================== delivery ============================== */

function deliverConsult_(o) {
  var start = Utilities.parseDate(String(o['Consult slot']), CFG.tz, 'yyyy-MM-dd HH:mm');
  var end = new Date(start.getTime() + CFG.consultMinutes * 60000);
  if (!o['Meet link']) {
    var ev = Calendar.Events.insert({
      summary: 'Sahithiness consultation · ' + o['Name'],
      description: 'Consultation call booked on sahithiness.in\n\nPhone: ' + String(o['Phone']).replace("'", '') +
        '\nHealth: ' + o['Health flags'] + '\n\n' + o['Details'],
      start: { dateTime: start.toISOString(), timeZone: CFG.tz },
      end: { dateTime: end.toISOString(), timeZone: CFG.tz },
      attendees: [{ email: o['Email'] }],
      conferenceData: { createRequest: { requestId: o['Order ID'], conferenceSolutionKey: { type: 'hangoutsMeet' } } },
      reminders: { useDefault: false, overrides: [{ method: 'email', minutes: 120 }, { method: 'popup', minutes: 15 }] }
    }, 'primary', { conferenceDataVersion: 1, sendUpdates: 'all' });
    o['Meet link'] = ev.hangoutLink || ((ev.conferenceData && ev.conferenceData.entryPoints || [])[0] || {}).uri || ev.htmlLink;
  }
  if (!o['Code issued']) {
    var code = newCode_();
    sheet_('Credits', CREDITS).appendRow([code, new Date(), creditExpiry_(o['Consult slot']), o['Order ID'], o['Name'], o['Phone'], o['Email'], '', '']);
    o['Code issued'] = code;
  }
}

function creditExpiry_(slot) {
  var start = Utilities.parseDate(String(slot), CFG.tz, 'yyyy-MM-dd HH:mm');
  return new Date(start.getTime() + (CFG.consultMinutes + CFG.creditHours * 60) * 60000);
}

function deliverGeneral_(o) {
  var folder = DriveApp.getFolderById(prop_('VIDEO_FOLDER_ID'));
  folder.addViewer(o['Gmail for videos']);
  var until = new Date(Date.now() + CFG.generalDays * 86400000);
  sheet_('Access', ACCESS).appendRow([o['Order ID'], o['Name'], o['Gmail for videos'], new Date(), until, 'Active']);
}

function markCreditUsed_(code, orderId) {
  var sh = sheet_('Credits', CREDITS), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === code && !rows[i][7]) { sh.getRange(i + 1, 8, 1, 2).setValues([[orderId, new Date()]]); return; }
  }
}

/* ============================== automatic jobs (set up by setup()) ============================== */

/** Every 10 minutes: catches payments where the browser closed before confirming. */
function reconcile() {
  var sh = sheet_('Orders', ORDERS), rows = sh.getDataRange().getValues(), now = Date.now();
  for (var i = 1; i < rows.length; i++) {
    var st = rows[i][C['Status'] - 1], orderId = rows[i][C['Order ID'] - 1], age = now - new Date(rows[i][0]).getTime();
    if (st !== 'Created' || !orderId) continue;
    try {
      var pays = rzp_('get', '/orders/' + orderId + '/payments').items || [];
      var good = pays.filter(function (p) { return p.status === 'captured' || p.status === 'authorized'; })[0];
      if (good) {
        confirmPayment_(orderId, good.id);
        deliver_(orderId, good.id);
      } else if (age > 24 * 3600000) {
        sh.getRange(i + 1, C['Status']).setValue('Not paid');
      }
    } catch (err) { console.error('reconcile ' + orderId + ': ' + err); }
  }
}

/** Daily: removes General Yoga video access after 90 days. */
function expireAccess() {
  var sh = sheet_('Access', ACCESS), rows = sh.getDataRange().getValues(), folder = null;
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][5] !== 'Active' || new Date(rows[i][4]) > new Date()) continue;
    try {
      folder = folder || DriveApp.getFolderById(prop_('VIDEO_FOLDER_ID'));
      folder.removeViewer(rows[i][2]);
      sh.getRange(i + 1, 6).setValue('Expired');
    } catch (err) { console.error('expire ' + rows[i][2] + ': ' + err); sh.getRange(i + 1, 6).setValue('Expired'); }
  }
}

/* ============================== emails ============================== */

function customerEmail_(o) {
  var first = String(o['Name']).split(' ')[0], k = o['Kind'], subject, body;
  var receipt = box_('<b>Payment received:</b> ₹' + inr_(o['Amount (₹)']) + '<br>Program: ' + esc_(o['Program']) +
    '<br>Order ID: ' + esc_(o['Order ID']) + '<br>Payment ID: ' + esc_(o['Payment ID']));
  if (k === 'consult') {
    subject = 'Your call with Sahithi is booked · ' + slotLabel_(o['Consult slot']);
    body = p_('Hi ' + esc_(first) + ', your 1:1 consultation call is confirmed.') +
      box_('<b>' + esc_(slotLabel_(o['Consult slot'])) + '</b> (30 minutes, India time)<br>' +
        (o['Meet link'] ? '<a href="' + esc_(o['Meet link']) + '" style="color:#4F6446;font-weight:600">Join on Google Meet</a>' : 'Your Google Meet link will be emailed shortly.')) +
      p_('A calendar invite has also been sent to this email. Please join from a quiet place with your questions ready.') +
      (o['Code issued'] ? box_('<b>Your ₹999 is not extra.</b><br>Use code <b style="font-size:18px;letter-spacing:1px">' + esc_(o['Code issued']) +
        '</b> when you enroll in any program on <a href="' + CFG.site + '" style="color:#4F6446">sahithiness.in</a> and ₹999 is taken off the fee.<br>Valid until ' + esc_(fmt_(creditExpiry_(o['Consult slot']))) + '.') : '');
  } else if (k === 'general') {
    subject = 'Welcome to General Yoga · your classes are ready';
    body = p_('Hi ' + esc_(first) + ', welcome to General Yoga with Sahithi!') +
      box_('Your 45 video classes are shared with <b>' + esc_(o['Gmail for videos']) + '</b>.<br><br>' +
        '<a href="https://drive.google.com/drive/folders/' + esc_(prop_('VIDEO_FOLDER_ID')) + '" style="display:inline-block;background:#4F6446;color:#fff;padding:12px 20px;border-radius:999px;text-decoration:none;font-weight:600">Open my classes</a><br><br>' +
        'Open the link while signed in to that Gmail account. You have access for 90 days, until ' + esc_(fmt_(new Date(Date.now() + CFG.generalDays * 86400000), 'd MMM yyyy')) + '.') +
      p_('Start with class 1 and follow the order. Practise within your comfort, and stop if anything hurts.');
  } else {
    subject = 'Welcome to ' + o['Program'];
    body = p_('Hi ' + esc_(first) + ', welcome to Sahithiness! Your place in <b>' + esc_(o['Program']) + '</b> is confirmed.') +
      box_('<b>What happens next</b><br>1. Sahithi reviews your health details.<br>2. She will message you on WhatsApp within 24 hours to fix your live session days and times.<br>3. Your personalised nutrition plan follows after your first session.') +
      (o['Health flags'] && o['Health flags'] !== 'None' ? p_('Because you shared some health conditions, Sahithi may ask for your doctor’s clearance before your first session.') : '');
  }
  body += receipt + p_('Questions? WhatsApp us on <a href="' + CFG.whatsappLink + '" style="color:#4F6446">' + CFG.whatsapp + '</a>.') + p_('With care,<br>Sahithi');
  MailApp.sendEmail({ to: o['Email'], subject: subject, htmlBody: wrap_(body), name: CFG.brand, replyTo: notifyEmail_() });
}

function adminEmail_(o) {
  var flagged = o['Health flags'] && o['Health flags'] !== 'None';
  var subject = (String(o['Delivered']).indexOf('ERROR') === 0 ? '[ACTION NEEDED] ' : flagged ? '[CHECK HEALTH] ' : '') +
    'New ' + o['Kind'] + ': ' + o['Name'] + ' · ₹' + inr_(o['Amount (₹)']);
  var lines = ['Program', 'Amount (₹)', 'Name', 'Phone', 'Email', 'Gmail for videos', 'Consult slot', 'Meet link', 'Credit code used', 'Code issued', 'Health flags', 'Delivered', 'Order ID', 'Payment ID'];
  var html = '<table style="border-collapse:collapse;font:14px Arial">' + lines.filter(function (h) { return o[h]; }).map(function (h) {
    return '<tr><td style="padding:6px 12px 6px 0;color:#777">' + h + '</td><td style="padding:6px 0">' + esc_(String(o[h]).replace(/^'/, '')) + '</td></tr>';
  }).join('') + '</table><pre style="font:13px Arial;white-space:pre-wrap;background:#f6f3ee;padding:12px;border-radius:8px">' + esc_(o['Details']) + '</pre>';
  MailApp.sendEmail({ to: notifyEmail_(), subject: subject, htmlBody: html, name: CFG.brand + ' website' });
}

function wrap_(inner) {
  return '<div style="background:#FBF8F3;padding:24px 12px;font:15px/1.6 Arial,sans-serif;color:#2E2A27"><div style="max-width:560px;margin:0 auto;background:#fff;border-radius:16px;padding:28px">' +
    '<div style="font:700 24px Georgia,serif;color:#4F6446;margin-bottom:16px">Sahithiness</div>' + inner + '</div></div>';
}
function p_(t) { return '<p style="margin:0 0 14px">' + t + '</p>'; }
function box_(t) { return '<div style="background:#F3EFE8;border-radius:12px;padding:16px;margin:0 0 14px">' + t + '</div>'; }

/* ============================== setup and helpers ============================== */

/** Run once after adding the script properties. Creates the Sheet and the automatic jobs, and checks everything. */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID'), ss;
  try { ss = id && SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  if (!ss) { ss = SpreadsheetApp.create('Sahithiness Orders'); props.setProperty('SHEET_ID', ss.getId()); }
  sheet_('Orders', ORDERS); sheet_('Credits', CREDITS); sheet_('Access', ACCESS);
  var first = ss.getSheetByName('Sheet1'); if (first && ss.getSheets().length > 1) ss.deleteSheet(first);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['reconcile', 'expireAccess'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('reconcile').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('expireAccess').timeBased().everyDays(1).atHour(3).create();

  var report = ['Orders sheet: ' + ss.getUrl()];
  ['RAZORPAY_KEY_ID', 'RAZORPAY_KEY_SECRET', 'VIDEO_FOLDER_ID'].forEach(function (k) { if (!props.getProperty(k)) report.push('MISSING script property: ' + k); });
  try { rzp_('get', '/orders?count=1'); report.push('Razorpay keys: OK (' + (prop_('RAZORPAY_KEY_ID').indexOf('rzp_live') === 0 ? 'LIVE' : 'TEST') + ' mode)'); } catch (e) { report.push('Razorpay keys: FAILED – ' + e.message); }
  try { report.push('Video folder: OK – ' + DriveApp.getFolderById(prop_('VIDEO_FOLDER_ID')).getName()); } catch (e) { report.push('Video folder: FAILED – check VIDEO_FOLDER_ID'); }
  try { Calendar.Events.list('primary', { maxResults: 1 }); report.push('Calendar: OK'); } catch (e) { report.push('Calendar: FAILED – add the Google Calendar API service (Services +)'); }
  report.push('Alerts go to: ' + notifyEmail_() + ' · Emails left today: ' + MailApp.getRemainingDailyQuota());
  console.log(report.join('\n'));
  return report.join('\n');
}

function prop_(k) {
  var v = PropertiesService.getScriptProperties().getProperty(k);
  if (!v) throw new Error('Missing script property ' + k);
  return v.trim();
}
function notifyEmail_() { return PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || Session.getEffectiveUser().getEmail(); }

function sheet_(name, headers) {
  var ss = SpreadsheetApp.openById(prop_('SHEET_ID')), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#DCE5D3');
    sh.setFrozenRows(1);
  }
  return sh;
}
function findRow_(sh, orderId) {
  var ids = sh.getRange(1, C['Order ID'], sh.getLastRow(), 1).getValues();
  for (var i = ids.length - 1; i > 0; i--) if (ids[i][0] === orderId) return i + 1;
  return 0;
}
function readRow_(sh, r) {
  var v = sh.getRange(r, 1, 1, ORDERS.length).getValues()[0], o = {};
  ORDERS.forEach(function (h, i) { o[h] = v[i]; });
  return o;
}
function setCells_(sh, r, o, cols) { cols.forEach(function (h) { sh.getRange(r, C[h]).setValue(o[h] === undefined ? '' : o[h]); }); }

function rzp_(method, path, payload) {
  var opts = {
    method: method, muteHttpExceptions: true,
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(prop_('RAZORPAY_KEY_ID') + ':' + prop_('RAZORPAY_KEY_SECRET')) }
  };
  if (payload) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(payload); }
  var res = UrlFetchApp.fetch('https://api.razorpay.com/v1' + path, opts), body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() >= 300) throw new Error('Razorpay ' + res.getResponseCode() + ': ' + (body.error && body.error.description || res.getContentText()));
  return body;
}

function hex_(bytes) { return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join(''); }
function clean_(v, max) { var s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B-\u001F]/g, '').trim().slice(0, max); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function inr_(n) { return Number(n).toLocaleString('en-IN'); }
function fmt_(d, pattern) { return Utilities.formatDate(new Date(d), CFG.tz, pattern || 'EEE d MMM yyyy, h:mm a'); }
function slotLabel_(key) { return key ? fmt_(Utilities.parseDate(String(key), CFG.tz, 'yyyy-MM-dd HH:mm')) : ''; }
