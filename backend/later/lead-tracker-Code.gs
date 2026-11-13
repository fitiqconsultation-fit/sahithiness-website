/**
 * Sahithiness – Lead & Payment Tracker
 * Receives enquiries from sahithiness.in and adds them to this sheet.
 * Open the sheet → Extensions → Apps Script → paste this file → Save →
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) → copy the Web app URL.
 */

var TAB_ID = 1864207904;   // the tab from the sheet link (gid=1864207904)
var TZ = 'Asia/Kolkata';

/* ===== Welcome messages: edit the Zoom link and texts here ===== */
// The Zoom link is kept private: add it under Project Settings → Script properties as ZOOM_LINK
var ZOOM_LINK = PropertiesService.getScriptProperties().getProperty('ZOOM_LINK') || 'Zoom link will be shared soon';
var WELCOME = {
  'Strength Training Classes':
    'Hi {name}!\n\nWelcome to Strength Training Classes with Sahithi. Your payment is confirmed.\n\n' +
    'Classes: Monday to Friday, 7-8 PM IST\nYour start date: {start}\nZoom link: {zoom}\n\n' +
    'Please join 5 minutes early with a mat, a water bottle and comfortable clothes. Keep light dumbbells or filled water bottles handy.\n\n' +
    'See you in class!\nTeam Sahithiness',
  'default':
    'Hi {name}!\n\nWelcome to {program} with Sahithi. Your payment is confirmed.\n\n' +
    'Your start date: {start}\nYou will get 5 new recorded videos every week. We will share your video access with you here.\n\n' +
    'Practise within your comfort, and message us here anytime.\nTeam Sahithiness'
};

function doGet() { return ContentService.createTextOutput('Sahithiness lead tracker is running.'); }

var WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxhVIIac981nqxuiM_RMf4D6P1bfkPCoXVUmwMc-c-IBta5kHpN3r385RyYlozRi2-WKQ/exec';

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false }); }
  // Razorpay payment (relayed by sahithiness.in/razorpay-webhook.php after checking Razorpay's signature)
  if (e.parameter && e.parameter.token) {
    if (e.parameter.token !== PropertiesService.getScriptProperties().getProperty('WEBHOOK_TOKEN')) return out_({ ok: false });
    var res = withLock_(function () { return markPaidFromRazorpay_(body); });
    if (res.row && PropertiesService.getScriptProperties().getProperty('MANYCHAT_API_KEY')) {
      try { sendWelcome_(tab_(), res.row); } catch (err) { console.error(err); }
    }
    return out_({ ok: true, status: res.status });
  }
  return out_(withLock_(function () { return saveLead_(body); }));
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } catch (err) { console.error(err && err.stack || err); return { ok: false }; } finally { lock.releaseLock(); }
}

/** Enquiry from the website (client tapped Pay / Register). */
function saveLead_(d) {
  {
    var phone = phoneText_(d.phone), wa = phoneText_(d.whatsapp);
    if (!phone || !String(d.name || '').trim()) return { ok: false };

    var sh = tab_();
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    ['Start Date', 'Enquiry Date', 'Postpartum Stage / Fitness Category', 'WhatsApp (if different)', 'Program Answers'].forEach(function (h) {           // extra columns, added once at the end
      if (head.indexOf(h) < 0) { addHeader_(sh, head.length + 1, h); head.push(h); }
    });

    var v = {
      'Full Name': clean_(d.name, 80),
      'Email': clean_(d.email, 120).toLowerCase(),
      'Phone / WhatsApp': "'" + phone,
      'WhatsApp (if different)': wa && wa !== phone ? "'" + wa : '',
      'Program': clean_(d.program, 80),
      'Payment Status': 'Pending',
      'Start Date': clean_(d.startDate, 40),
      'Postpartum Stage / Fitness Category': clean_(d.stage, 80),
      'Program Answers': clean_(d.answers, 500),
      'Enquiry Date': Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy, h:mm a')
    };

    // Same person tapping Pay again for the same program: update their row instead of adding a duplicate
    var rows = sh.getDataRange().getValues(), cP = head.indexOf('Phone / WhatsApp'), cG = head.indexOf('Program'), cS = head.indexOf('Payment Status');
    for (var i = rows.length - 1; i > 0; i--) {
      if (key_(rows[i][cP]) === key_(phone) && rows[i][cG] === v['Program'] && rows[i][cS] !== 'Paid') {
        head.forEach(function (h, c) { if (v[h] !== undefined) sh.getRange(i + 1, c + 1).setValue(v[h]); });
        return { ok: true, updated: true };
      }
    }
    sh.appendRow(head.map(function (h) { return v[h] === undefined ? '' : v[h]; }));
    return { ok: true };
  }
}

/** Razorpay says a payment succeeded: mark the client's row Paid with the real payment date and time. */
function markPaidFromRazorpay_(body) {
  if (body.event !== 'payment.captured') return { status: 'ignored' };
  var p = body.payload && body.payload.payment && body.payload.payment.entity;
  if (!p || !p.id) return { status: 'no payment' };
  var sh = tab_(), head = headers_(sh);
  ['Payment ID', 'Amount (₹)'].forEach(function (h) { if (head.indexOf(h) < 0) { addHeader_(sh, head.length + 1, h); head.push(h); } });
  var rows = sh.getDataRange().getValues(), c = function (h) { return head.indexOf(h); };
  for (var k = 1; k < rows.length; k++) if (rows[k][c('Payment ID')] === p.id) return { status: 'duplicate' };   // Razorpay may resend

  var phone = String(p.contact || '').replace(/\D/g, '').slice(-10), email = String(p.email || '').trim().toLowerCase();
  var amount = (p.amount || 0) / 100, guess = amount === 3000 ? 'Strength Training Classes' : '';
  var samePerson = function (r) {
    return (phone && String(r[c('Phone / WhatsApp')]).replace(/\D/g, '').slice(-10) === phone) ||
           (email && String(r[c('Email')]).trim().toLowerCase() === email);
  };
  var match = 0, earlier = null;
  for (var i = rows.length - 1; i > 0; i--) {
    if (!samePerson(rows[i])) continue;
    earlier = earlier || rows[i];
    var prog = String(rows[i][c('Program')]).replace(/\u00a0/g, ' ');
    if (rows[i][c('Payment Status')] !== 'Paid' && (!guess || prog === guess)) { match = i + 1; break; }
  }
  var when = p.created_at ? new Date(p.created_at * 1000) : new Date();
  var v = {
    'Payment Status': 'Paid', 'Payment Date': Utilities.formatDate(when, TZ, 'dd MMM yyyy'),
    'Payment Time': Utilities.formatDate(when, TZ, 'h:mm a'), 'Payment ID': p.id, 'Amount (₹)': amount
  };
  if (match) {
    head.forEach(function (h, col) { if (v[h] !== undefined) sh.getRange(match, col + 1).setValue(v[h]); });
    return { status: 'matched', row: match };
  }
  // No enquiry row: a renewal, or someone who paid without using the website form
  v['Full Name'] = earlier ? earlier[c('Full Name')] : clean_((p.notes && (p.notes.name || p.notes.full_name)) || email || 'Unknown', 80);
  v['Email'] = email || (earlier ? earlier[c('Email')] : '');
  v['Phone / WhatsApp'] = phoneText_(p.contact) ? "'" + phoneText_(p.contact) : '';
  v['Program'] = guess || (earlier ? earlier[c('Program')] : 'Check program (₹' + amount + ')');
  v['Enquiry Date'] = earlier ? 'Renewal' : 'Paid without form';
  // No automatic welcome for renewals or unknown programs (clear this cell to allow sending)
  if (head.indexOf('Welcome Sent') >= 0 && (earlier || !guess)) v['Welcome Sent'] = earlier ? '✓ renewal (no welcome)' : '✓ check program first, then clear this cell';
  sh.appendRow(head.map(function (h) { return v[h] === undefined ? '' : v[h]; }));
  return { status: 'new row', row: sh.getLastRow() };
}

/** Menu: prepares the private link that lets Razorpay update this sheet. */
function setupRazorpay() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('WEBHOOK_TOKEN')) props.setProperty('WEBHOOK_TOKEN', Utilities.getUuid().replace(/-/g, ''));
  var url = (ScriptApp.getService().getUrl() || WEB_APP_URL).replace(/\/dev$/, '/exec') + '?token=' + props.getProperty('WEBHOOK_TOKEN');
  var html = HtmlService.createHtmlOutput('<div style="font:14px Arial"><p>Copy this private link into <b>sahithiness-config.php</b> on Hostinger as <b>sheet_webhook_url</b>:</p>' +
    '<textarea style="width:100%;height:90px" onclick="this.select()">' + url + '</textarea><p>Keep it private.</p></div>').setWidth(520).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Razorpay → Sheet link');
}

/** Phone as typed: Indian 10-digit numbers stay as 10 digits, others keep their country code (+...). */
function phoneText_(v) {
  var s = String(v || '').trim(), d = s.replace(/\D/g, '');
  if (d.length < 5) return '';
  if (/^[6-9]\d{9}$/.test(d)) return d;                       // Indian 10-digit
  if (/^(91|0)[6-9]\d{9}$/.test(d)) return d.slice(-10);       // Indian with 91 / 0
  return s.charAt(0) === '+' ? '+' + d : d;                    // anything else, kept as typed
}
/** Digits for WhatsApp/ManyChat (91 added for Indian 10-digit numbers). */
function waDigits_(v) {
  var p = phoneText_(v);
  if (!p) return '';
  return p.charAt(0) === '+' ? p.slice(1) : (/^[6-9]\d{9}$/.test(p) ? '91' + p : p);
}
/** Matching key: last 10 digits. */
function key_(v) { return String(v || '').replace(/\D/g, '').slice(-10); }

function tab_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().filter(function (s) { return s.getSheetId() === TAB_ID; })[0] || ss.getSheets()[0];
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(v, max) { var s = String(v == null ? '' : v).replace(/[\u0000-\u001F\u00a0]/g, ' ').trim().slice(0, max); return /^[=+\-@]/.test(s) ? "'" + s : s; }

/* ============ Welcome on WhatsApp ============
 * When "Payment Status" is changed to Paid, this fills Payment Date / Payment Time (if empty)
 * and adds a "WhatsApp Welcome" link: tap it, WhatsApp opens with the message typed, press Send.
 * For rows already marked Paid, use the menu: Sahithiness → Create welcome links for Paid clients.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sahithiness')
    .addItem('Send welcome (ManyChat) to Paid clients not yet sent', 'sendPendingWelcomes')
    .addItem('Create WhatsApp links for Paid clients (manual backup)', 'createWelcomeLinks')
    .addItem('Fix header colours', 'styleHeaders')
    .addItem('Set up Razorpay auto-Paid (get private link)', 'setupRazorpay')
    .addSeparator()
    .addItem('Set up automatic welcome (run once)', 'setupAutoWelcome')
    .addToUi();
}

function onEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getSheetId() !== TAB_ID || e.range.getRow() < 2) return;
  var head = headers_(sh), cS = head.indexOf('Payment Status') + 1;
  if (cS < 1 || cS < e.range.getColumn() || cS > e.range.getLastColumn()) return;
  for (var r = e.range.getRow(); r <= e.range.getLastRow(); r++) welcomeRow_(sh, r, true);
}

function createWelcomeLinks() {
  var sh = tab_(), n = 0;
  for (var r = 2; r <= sh.getLastRow(); r++) if (welcomeRow_(sh, r)) n++;
  SpreadsheetApp.getActive().toast(n + ' welcome link(s) ready in the "WhatsApp Welcome" column.', 'Sahithiness', 6);
}

function welcomeRow_(sh, r, datesOnly) {
  var head = headers_(sh);
  if (!datesOnly) ['WhatsApp Welcome'].forEach(function (h) {
    if (head.indexOf(h) < 0) { addHeader_(sh, head.length + 1, h); head.push(h); }
  });
  var row = sh.getRange(r, 1, 1, head.length).getValues()[0], get = function (h) { return row[head.indexOf(h)]; };
  if (String(get('Payment Status')).trim().toLowerCase() !== 'paid') return false;
  var now = new Date();
  if (head.indexOf('Payment Date') >= 0 && !get('Payment Date')) sh.getRange(r, head.indexOf('Payment Date') + 1).setValue(Utilities.formatDate(now, TZ, 'dd MMM yyyy'));
  if (head.indexOf('Payment Time') >= 0 && !get('Payment Time')) sh.getRange(r, head.indexOf('Payment Time') + 1).setValue(Utilities.formatDate(now, TZ, 'h:mm a'));
  if (datesOnly) return true;                       // ManyChat sends the message; no link needed
  var phone = waDigits_(get('WhatsApp (if different)') || get('Phone / WhatsApp'));
  if (!phone) return false;
  var program = String(get('Program')).replace(/\u00a0/g, ' ').trim();
  var msg = (WELCOME[program] || WELCOME['default'])
    .replace(/{name}/g, String(get('Full Name')).trim().split(' ')[0] || 'there')
    .replace(/{program}/g, program)
    .replace(/{start}/g, dateText_(get('Start Date')) || 'as discussed')
    .replace(/{zoom}/g, ZOOM_LINK);
  var url = 'https://api.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(msg);
  sh.getRange(r, head.indexOf('WhatsApp Welcome') + 1).setFormula('=HYPERLINK("' + url.replace(/"/g, '""') + '","📲 Send welcome")');
  return true;
}

// The sheet turns "Wed, 4 Nov, 2026" into a date; show it as a readable date again
function dateText_(v) {
  if (v && typeof v.getTime === 'function') return Utilities.formatDate(v, TZ, 'EEE, d MMM yyyy');
  return String(v || '').trim();
}
// ManyChat "Date" fields need YYYY-MM-DD. Handles a sheet date or text like "Wed, 4 Nov, 2026".
function isoDate_(v) {
  var d = v && typeof v.getTime === 'function' ? v : new Date(String(v || '').replace(/^[A-Za-z]{3},\s*/, ''));
  if (isNaN(d)) d = new Date();
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function headers_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }

/* ============ Automatic welcome through ManyChat (WhatsApp 98885 88825) ============
 * When Payment Status becomes Paid, the sheet sends the client to ManyChat, fills her details
 * (wa_number, program, start_date, zoom_link) and starts the welcome flow, which sends the approved template.
 * Script properties needed (Project Settings → Script properties):
 *   MANYCHAT_API_KEY      ManyChat → Settings → API → Generate API key
 *   MANYCHAT_FLOW_NS      the welcome flow's ID (from the flow's URL, looks like content20251104123456_123456)
 *   MANYCHAT_WA_FIELD_ID  the ID of the custom field "wa_number" (Settings → Fields → wa_number)
 * Then run setupAutoWelcome() once.
 */
var MC = 'https://api.manychat.com';
// ManyChat custom field IDs (from the template preview: cuf_15128968 = program, etc.)
var MC_FIELDS = { program: 15128968, start_date: 15128969, zoom_link: 15128970 };

function setupAutoWelcome() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'onPaidEdit') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onPaidEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  var ok = mc_('get', '/fb/page/getInfo');
  waFieldId_();
  SpreadsheetApp.getActive().toast('ManyChat connected: ' + (ok.data && ok.data.name || 'OK') + '. Marking a client Paid will now send the welcome automatically.', 'Sahithiness', 10);
}

/** Runs on every edit (installed by setupAutoWelcome). */
function onPaidEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getSheetId() !== TAB_ID || e.range.getRow() < 2) return;
  var cS = headers_(sh).indexOf('Payment Status') + 1;
  if (cS < 1 || cS < e.range.getColumn() || cS > e.range.getLastColumn()) return;
  for (var r = e.range.getRow(); r <= e.range.getLastRow(); r++) sendWelcome_(sh, r);
}

/** Menu: send to every Paid client who has not been sent a welcome yet. */
function sendPendingWelcomes() {
  var sh = tab_(), n = 0;
  for (var r = 2; r <= sh.getLastRow(); r++) if (sendWelcome_(sh, r) === 'sent') n++;
  SpreadsheetApp.getActive().toast(n + ' welcome message(s) sent through ManyChat.', 'Sahithiness', 8);
}

function sendWelcome_(sh, r) {
  var head = headers_(sh);
  ['Welcome Sent'].forEach(function (h) {
    if (head.indexOf(h) < 0) { addHeader_(sh, head.length + 1, h); head.push(h); }
  });
  var row = sh.getRange(r, 1, 1, head.length).getValues()[0], get = function (h) { return row[head.indexOf(h)]; };
  var cell = sh.getRange(r, head.indexOf('Welcome Sent') + 1);
  if (String(get('Payment Status')).trim().toLowerCase() !== 'paid') return 'skip';
  if (/^✓/.test(String(get('Welcome Sent')))) return 'skip';                 // never send twice
  var waNum = waDigits_(get('WhatsApp (if different)') || get('Phone / WhatsApp'));
  if (!waNum) { cell.setValue('✗ Invalid phone'); return 'error'; }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (/^✓/.test(String(cell.getValue()))) return 'skip';
    var wa = waNum, name = String(get('Full Name')).trim(), program = String(get('Program')).replace(/ /g, ' ').trim();
    var id = findContact_(wa);
    if (!id) {
      var made = mc_('post', '/fb/subscriber/createSubscriber', {
        first_name: name.split(' ')[0] || name, last_name: name.split(' ').slice(1).join(' '),
        whatsapp_phone: '+' + wa, consent_phrase: 'Paid on sahithiness.in and agreed to receive class updates on WhatsApp'
      });
      id = made.data && made.data.id;
    }
    if (!id) throw new Error('Could not create contact');
    [[Number(waFieldId_()), wa], [MC_FIELDS.program, program], [MC_FIELDS.start_date, isoDate_(get('Start Date'))],
     [MC_FIELDS.zoom_link, /strength/i.test(program) ? ZOOM_LINK : 'Your recorded videos will be shared here on WhatsApp']].forEach(function (f) {
      mc_('post', '/fb/subscriber/setCustomField', { subscriber_id: id, field_id: f[0], field_value: f[1] });
    });
    mc_('post', '/fb/sending/sendFlow', { subscriber_id: id, flow_ns: prop_('MANYCHAT_FLOW_NS') });
    cell.setValue('✓ ' + Utilities.formatDate(new Date(), TZ, 'dd MMM, h:mm a'));
    return 'sent';
  } catch (err) {
    var msg = String(err && err.message || err);
    if (/already exists/i.test(msg)) msg = 'Contact already in ManyChat without wa_number. Run the one-time "copy WhatsApp ID" broadcast, then try again';
    cell.setValue('✗ ' + msg.slice(0, 180));
    return 'error';
  } finally {
    lock.releaseLock();
  }
}

/** ID of the ManyChat field "wa_number": from Script properties if set, otherwise looked up once and remembered. */
function waFieldId_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('MANYCHAT_WA_FIELD_ID');
  if (id) return id.trim();
  var fields = mc_('get', '/fb/page/getCustomFields').data || [];
  var f = fields.filter(function (x) { return String(x.name).trim().toLowerCase() === 'wa_number'; })[0];
  if (!f) throw new Error('ManyChat field "wa_number" not found');
  props.setProperty('MANYCHAT_WA_FIELD_ID', String(f.id));
  return String(f.id);
}

function findContact_(wa) {
  var res = mc_('get', '/fb/subscriber/findByCustomField?field_id=' + encodeURIComponent(waFieldId_()) + '&field_value=' + wa);
  var list = res.data || [];
  return list.length ? list[0].id : null;
}

function mc_(method, path, body) {
  var opts = { method: method, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + prop_('MANYCHAT_API_KEY') } };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  var res = UrlFetchApp.fetch(MC + path, opts), out = {};
  try { out = JSON.parse(res.getContentText()); } catch (e) {}
  if (res.getResponseCode() >= 300 || out.status === 'error') {
    var detail = [out.message, out.details && JSON.stringify(out.details.messages || out.details)].filter(Boolean).join(' ') || res.getContentText();
    throw new Error('ManyChat: ' + detail);
  }
  return out;
}

function prop_(k) {
  var v = PropertiesService.getScriptProperties().getProperty(k);
  if (!v) throw new Error('Missing script property ' + k);
  return v.trim();
}

/** New header cells copy the look (green fill, bold, borders) of the "Full Name" header. */
function addHeader_(sh, col, title) {
  sh.getRange(1, 1).copyFormatToRange(sh, col, col, 1, 1);
  sh.getRange(1, col).setValue(title);
}

/** Menu: make every header cell match the "Full Name" header. */
function styleHeaders() {
  var sh = tab_();
  sh.getRange(1, 1).copyFormatToRange(sh, 2, sh.getLastColumn(), 1, 1);
  SpreadsheetApp.getActive().toast('Header colours updated.', 'Sahithiness', 4);
}
