/**
 * Meeting invitations -> Google Sheet (optional log).
 *
 * Paste this file into the Apps Script editor of a Google Sheet
 * (Extensions > Apps Script), deploy it as a web app and put the /exec URL in
 * CONFIG.scriptUrl in index.html. Full steps are in README.md.
 *
 * The web app only appends rows. It never returns meeting data, so the public
 * URL can't be used to read who is coming.
 */

var SHEET_NAME = 'זימונים';

var FIELDS = [
  ['timestamp', 'נשלח בתאריך'],
  ['date', 'תאריך הפגישה'],
  ['start', 'שעת התחלה'],
  ['end', 'שעת סיום'],
  ['guest_name', 'שם המוזמן'],
  ['guest_phone', 'טלפון'],
  ['host', 'הפגישה עם'],
  ['notes', 'הערה'],
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(FIELDS.map(function (f) { return f[1]; }));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, FIELDS.length).setFontWeight('bold');
    sheet.setRightToLeft(true);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Strips a leading =, +, - or @ so a submitted value can't run as a formula.
function clean_(value) {
  var s = String(value || '').trim().slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  var p = (e && e.parameter) || {};

  var required = ['guest_name', 'guest_phone', 'host', 'date', 'start', 'end'];
  for (var i = 0; i < required.length; i++) {
    if (!String(p[required[i]] || '').trim()) {
      return json_({ ok: false, error: 'missing_' + required[i] });
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var row = FIELDS.map(function (f) {
      return f[0] === 'timestamp' ? new Date() : clean_(p[f[0]]);
    });
    getSheet_().appendRow(row);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
