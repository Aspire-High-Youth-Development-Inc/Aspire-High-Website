/**
 * Aspire High - sponsor form responses
 * ------------------------------------------------------------------------
 * Receives submissions from sponsor-form.html and adds one row per
 * submission to the "Sponsor Responses" tab of THIS spreadsheet.
 *
 * Setup (once):
 *   1. Open your Google Sheet -> Extensions -> Apps Script.
 *   2. Delete everything in the editor, paste this whole file, and Save.
 *   3. Project Settings (gear) -> Time zone: (GMT-04:00) New York.
 *   4. Choose "setup" in the function dropdown and click Run (approve the
 *      permissions). This creates the tab and the column headings.
 *   5. Deploy -> New deployment -> type "Web app".
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Deploy, approve the permissions, copy the Web app URL (ends in /exec).
 *   6. In sponsor-form.html, paste that URL between the quotes of
 *        var ENDPOINT = "";
 *      and upload the file to GitHub again.
 *
 * "Anyone" only lets people SEND the form. Nothing is ever sent back, and the
 * sheet stays private. Keep the sheet's sharing on Restricted.
 * After editing this script later: Deploy -> Manage deployments -> edit ->
 * Version: New version.
 */
var SHEET_NAME = 'Sponsor Responses';
var NOTIFY_EMAIL = '';   // optional: an address to tell when a form arrives
var COLUMNS = [          // [field name on the form, column heading]
  ['submitted_at',         'Submitted at'],
  ['email',                'Email'],
  ['company_name',         'Company name'],
  ['company_website',      'Company website'],
  ['company_size',         'Estimated company size'],
  ['contact_name',         'Contact name'],
  ['contact_email',        'Contact email'],
  ['location',             'Location'],
  ['location_other',       'Other location'],
  ['partnership_interests','Ways they would like to partner'],
  ['message',              'Additional comments']
];
var TEXT_FIELDS = { company_size: 1 };   // keep "10-50" as text, not a date
var HONEYPOT = 'website';
var MAX_FIELD = 3000;
var MAX_FIELDS = 40;

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = (e && e.parameters) || {};
    if (Object.keys(p).length > MAX_FIELDS) return json_({ ok: false, error: 'too many fields' });
    if (p[HONEYPOT] && String(p[HONEYPOT][0] || '').trim() !== '') return json_({ ok: true });
    var sheet = sheet_();
    var headers = ensureHeaders_(sheet, p);
    var byHeader = {};
    COLUMNS.forEach(function (c) { byHeader[c[1]] = c[0]; });
    var row = headers.map(function (h) {
      var field = byHeader[h] || h;
      if (field === 'submitted_at') return new Date();
      var v = p[field];
      if (!v) return '';
      var out = clean_(v.join(', '));
      return (TEXT_FIELDS[field] && out !== '' && out.charAt(0) !== "'") ? "'" + out : out;
    });
    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm');
    notify_();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* Opening the /exec URL in a browser shows this, to confirm the deployment is live. */
function doGet() {
  return json_({ ok: true, service: 'Aspire High sponsor form', accepts: 'POST' });
}

/* Run once from the editor (Run -> setup) to create the tab and headings. */
function setup() {
  var sheet = sheet_();
  ensureHeaders_(sheet, {});
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
       .setFontWeight('bold').setBackground('#0B2D5B').setFontColor('#FFFFFF');
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function ensureHeaders_(sheet, p) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  if (headers.every(function (h) { return h === ''; })) headers = [];
  var wanted = COLUMNS.map(function (c) { return c[1]; });
  var known = {};
  COLUMNS.forEach(function (c) { known[c[0]] = true; });
  Object.keys(p).forEach(function (f) {
    if (!known[f] && f !== HONEYPOT && /^[A-Za-z0-9_]{1,60}$/.test(f)) wanted.push(f);
  });
  var added = false;
  wanted.forEach(function (h) {
    if (headers.indexOf(h) === -1) { headers.push(h); added = true; }
  });
  if (added) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}

/* Answers come from the public, so they are treated as text: a value starting
   with = + - or @ would otherwise run as a formula in Sheets and Excel. */
function clean_(s) {
  s = String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, MAX_FIELD);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function notify_() {
  if (!NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New Aspire High sponsor form',
      body: 'A new sponsor form was submitted.\n\nOpen the sheet:\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });
  } catch (err) { /* a mail failure must never lose a submission */ }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
