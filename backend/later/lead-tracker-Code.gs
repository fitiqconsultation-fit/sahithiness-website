/**
 * Sahithiness – Lead & Payment Tracker
 * Receives enquiries from sahithiness.in and adds them to this sheet.
 * Open the sheet → Extensions → Apps Script → paste this file → Save →
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) → copy the Web app URL.
 */

var TAB_ID = 1864207904;   // the tab from the sheet link (gid=1864207904)
var TZ = 'Asia/Kolkata';

/* ===== Welcome messages: edit the Zoom link and texts here ===== */
var ZOOM_LINK = 'PASTE_ZOOM_LINK_HERE';
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

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var phone = String(d.phone || '').replace(/\D/g, '').slice(-10);
    if (!/^[6-9]\d{9}$/.test(phone) || !String(d.name || '').trim()) return out_({ ok: false });

    var sh = tab_();
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    ['Start Date', 'Enquiry Date', 'Postpartum Stage / Fitness Category'].forEach(function (h) {           // extra columns, added once at the end
      if (head.indexOf(h) < 0) { sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold'); head.push(h); }
    });

    var v = {
      'Full Name': clean_(d.name, 80),
      'Email': clean_(d.email, 120).toLowerCase(),
      'Phone / WhatsApp': "'" + phone,
      'Program': clean_(d.program, 80),
      'Payment Status': 'Pending',
      'Start Date': clean_(d.startDate, 40),
      'Postpartum Stage / Fitness Category': clean_(d.stage, 80),
      'Enquiry Date': Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy, h:mm a')
    };

    // Same person tapping Pay again for the same program: update their row instead of adding a duplicate
    var rows = sh.getDataRange().getValues(), cP = head.indexOf('Phone / WhatsApp'), cG = head.indexOf('Program'), cS = head.indexOf('Payment Status');
    for (var i = rows.length - 1; i > 0; i--) {
      if (String(rows[i][cP]).replace(/\D/g, '').slice(-10) === phone && rows[i][cG] === v['Program'] && rows[i][cS] !== 'Paid') {
        head.forEach(function (h, c) { if (v[h] !== undefined) sh.getRange(i + 1, c + 1).setValue(v[h]); });
        return out_({ ok: true, updated: true });
      }
    }
    sh.appendRow(head.map(function (h) { return v[h] === undefined ? '' : v[h]; }));
    return out_({ ok: true });
  } catch (err) {
    console.error(err && err.stack || err);
    return out_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

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
    .addItem('Create welcome links for Paid clients', 'createWelcomeLinks')
    .addToUi();
}

function onEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getSheetId() !== TAB_ID || e.range.getRow() < 2) return;
  var head = headers_(sh), cS = head.indexOf('Payment Status') + 1;
  if (cS < 1 || cS < e.range.getColumn() || cS > e.range.getLastColumn()) return;
  for (var r = e.range.getRow(); r <= e.range.getLastRow(); r++) welcomeRow_(sh, r);
}

function createWelcomeLinks() {
  var sh = tab_(), n = 0;
  for (var r = 2; r <= sh.getLastRow(); r++) if (welcomeRow_(sh, r)) n++;
  SpreadsheetApp.getActive().toast(n + ' welcome link(s) ready in the "WhatsApp Welcome" column.', 'Sahithiness', 6);
}

function welcomeRow_(sh, r) {
  var head = headers_(sh);
  ['WhatsApp Welcome'].forEach(function (h) {
    if (head.indexOf(h) < 0) { sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold'); head.push(h); }
  });
  var row = sh.getRange(r, 1, 1, head.length).getValues()[0], get = function (h) { return row[head.indexOf(h)]; };
  if (String(get('Payment Status')).trim().toLowerCase() !== 'paid') return false;
  var now = new Date();
  if (head.indexOf('Payment Date') >= 0 && !get('Payment Date')) sh.getRange(r, head.indexOf('Payment Date') + 1).setValue(Utilities.formatDate(now, TZ, 'dd MMM yyyy'));
  if (head.indexOf('Payment Time') >= 0 && !get('Payment Time')) sh.getRange(r, head.indexOf('Payment Time') + 1).setValue(Utilities.formatDate(now, TZ, 'h:mm a'));
  var phone = String(get('Phone / WhatsApp')).replace(/\D/g, '').slice(-10);
  if (!/^[6-9]\d{9}$/.test(phone)) return false;
  var program = String(get('Program')).replace(/\u00a0/g, ' ').trim();
  var msg = (WELCOME[program] || WELCOME['default'])
    .replace(/{name}/g, String(get('Full Name')).trim().split(' ')[0] || 'there')
    .replace(/{program}/g, program)
    .replace(/{start}/g, dateText_(get('Start Date')) || 'as discussed')
    .replace(/{zoom}/g, ZOOM_LINK);
  var url = 'https://api.whatsapp.com/send?phone=91' + phone + '&text=' + encodeURIComponent(msg);
  sh.getRange(r, head.indexOf('WhatsApp Welcome') + 1).setFormula('=HYPERLINK("' + url.replace(/"/g, '""') + '","📲 Send welcome")');
  return true;
}

// The sheet turns "Wed, 4 Nov, 2026" into a date; show it as a readable date again
function dateText_(v) {
  if (v && typeof v.getTime === 'function') return Utilities.formatDate(v, TZ, 'EEE, d MMM yyyy');
  return String(v || '').trim();
}
function headers_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]; }
