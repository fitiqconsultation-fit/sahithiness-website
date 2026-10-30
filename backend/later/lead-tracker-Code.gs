/**
 * Sahithiness – Lead & Payment Tracker
 * Receives enquiries from sahithiness.in and adds them to this sheet.
 * Open the sheet → Extensions → Apps Script → paste this file → Save →
 * Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) → copy the Web app URL.
 */

var TAB_ID = 1864207904;   // the tab from the sheet link (gid=1864207904)
var TZ = 'Asia/Kolkata';

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
    ['Start Date', 'Enquiry Date'].forEach(function (h) {           // extra columns, added once at the end
      if (head.indexOf(h) < 0) { sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold'); head.push(h); }
    });

    var v = {
      'Full Name': clean_(d.name, 80),
      'Email': clean_(d.email, 120).toLowerCase(),
      'Phone / WhatsApp': "'" + phone,
      'Program': clean_(d.program, 80),
      'Payment Status': 'Pending',
      'Start Date': clean_(d.startDate, 40),
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
