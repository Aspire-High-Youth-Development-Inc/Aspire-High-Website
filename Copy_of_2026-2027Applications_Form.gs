/**
 * Aspire High — application responses
 * ------------------------------------------------------------------------
 * Receives submissions from the website's application forms
 * (English and Spanish, both inside index.html) and adds one row per
 * application to the "Responses" tab of THIS spreadsheet.
 *
 * Setup (once):
 *   1. Open the responses spreadsheet → Extensions → Apps Script.
 *   2. Replace everything in Code.gs with this file. Save.
 *   3. Project Settings (gear) → Time zone: (GMT-04:00) New York.
 *   4. Deploy → New deployment → type "Web app".
 *        Execute as:      Me
 *        Who has access:  Anyone      (NOT "Anyone with Google account" —
 *                                      families are not signed in to Google)
 *      Deploy, approve the permissions, copy the Web app URL (ends in /exec).
 *   5. Put that URL in index.html, between the quotes of
 *        applicationEndpoint: ""
 *      near the top of the script, and upload index.html again.
 *
 * "Anyone" only lets people SEND an application. Nothing here ever returns
 * a response back — the spreadsheet itself stays private to whoever it is
 * shared with. Keep the spreadsheet's sharing on Restricted.
 *
 * After editing this script later: Deploy → Manage deployments → edit →
 * Version: New version. Editing the code alone does not update the live URL.
 */

var SHEET_NAME = 'Responses';

/* Optional: an address to notify when an application arrives, e.g. the
   Franklin High School coordinator. The email says only that one arrived and
   links to the sheet — no student details travel by email. Leave '' for none. */
var NOTIFY_EMAIL = '';

/* Field name on the form → column heading in the sheet, in form order.
   Columns are matched by heading, so staff may reorder or add columns in the
   sheet without breaking anything. A field not listed here still gets saved,
   in a new column named after the field. */
var COLUMNS = [
  ['submitted_at',            'Submitted at'],
  ['form_language',           'Form language'],
  ['email',                   'Email'],
  ['student_first_name',      'Student first name'],
  ['student_last_name',       'Student last name'],
  ['preferred_name',          'Preferred name'],
  ['student_dob',             'Student date of birth'],
  ['grade_level',             'Grade level'],
  ['school_name',             'School'],
  ['school_district',         'School district'],
  ['zip_code',                'ZIP code'],
  ['student_email',           'Student email'],
  ['student_phone',           'Student phone'],
  ['parent_name',             'Parent/guardian name'],
  ['parent_relationship',     'Relationship to student'],
  ['parent_email',            'Parent/guardian email'],
  ['parent_phone',            'Parent/guardian phone'],
  ['contact_method',          'Preferred contact method'],
  ['household_size',          'Household size'],
  ['parent_education',        'Parent/guardian education'],
  ['first_gen',               'First-generation college'],
  ['school_lunch',            'School lunch'],
  ['english_learner',         'English learner'],
  ['race_ethnicity',          'Race/ethnicity'],
  ['barriers',                'Barriers to participation'],
  ['household_income',        'Household income range'],
  ['programs',                'Programs of interest'],
  ['goals',                   'Student goals'],
  ['after_high_school',       'Plan after high school'],
  ['gpa',                     'GPA'],
  ['academic_performance',    'Academic performance'],
  ['academic_support_areas',  'Academic support needed'],
  ['year_goal',               'Goal for this year'],
  ['virtual_tutoring',        'Wants virtual tutoring'],
  ['tutoring_subjects',       'Tutoring subjects'],
  ['tutoring_time',           'Tutoring time'],
  ['parent_engagement',       'Parent engagement interest'],
  ['parent_session_time',     'Parent session times'],
  ['parent_topics',           'Parent topics'],
  ['emergency_contact_name',  'Emergency contact name'],
  ['emergency_relationship',  'Emergency contact relationship'],
  ['emergency_contact_phone', 'Emergency contact phone'],
  ['accessibility_needs',     'Accessibility needs'],
  ['heard_about',             'How they heard about us'],
  ['consent_participation',   'Participation consent'],
  ['consent_evaluation',      'Evaluation consent'],
  ['photo_release',           'Photo/video release'],
  ['signature',               'Signature'],
  ['signature_date',          'Signature date']
];

/* Kept as text exactly as typed. Sheets would otherwise read "08724" as the
   number 8724 — and every ZIP code in New Jersey starts with a zero. */
var TEXT_FIELDS = { zip_code: 1, student_phone: 1, parent_phone: 1, emergency_contact_phone: 1, gpa: 1 };

var HONEYPOT = 'website';      // hidden field only a bot fills in
var MAX_FIELD = 3000;          // characters kept per answer
var MAX_FIELDS = 120;          // a real application has about fifty

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);      // two parents submitting at once get two rows, not one

    var p = (e && e.parameters) || {};
    if (Object.keys(p).length > MAX_FIELDS) return json_({ ok: false, error: 'too many fields' });

    /* A bot filled the hidden field. Say yes, keep nothing. */
    if (p[HONEYPOT] && String(p[HONEYPOT][0] || '').trim() !== '') return json_({ ok: true });

    var sheet = sheet_();
    var headers = ensureHeaders_(sheet, p);
    var byHeader = {};
    COLUMNS.forEach(function (c) { byHeader[c[1]] = c[0]; });

    var row = headers.map(function (h) {
      var field = byHeader[h] || h;            // extra columns are named after their field
      if (field === 'submitted_at') return new Date();
      var v = p[field];
      if (!v) return '';
      var out = clean_(v.join(', '));
      return (TEXT_FIELDS[field] && out !== '' && out.charAt(0) !== "'") ? "'" + out : out;
    });

    sheet.appendRow(row);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm');

    notify_(p);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* Opening the /exec URL in a browser shows this, so you can confirm the
   deployment is live. It never returns any application data. */
function doGet() {
  return json_({ ok: true, service: 'Aspire High applications', accepts: 'POST' });
}

/* Run once from the editor (Run → setup) to create the tab and headings
   before the first application arrives. Optional — doPost does it too. */
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

/* Make sure every known heading, and a heading for any unknown field in this
   submission, exists in row 1. Returns the full heading row. */
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

/* Answers come from the public, so they are treated as text: a value that
   starts with = + - or @ would otherwise run as a formula in Sheets and in
   Excel when staff open the download. Phone numbers like "+1 732…" keep
   their plus sign — the apostrophe is not shown in the cell. */
function clean_(s) {
  s = String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, MAX_FIELD);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function notify_(p) {
  if (!NOTIFY_EMAIL) return;
  try {
    var lang = p.form_language ? p.form_language[0] : '';
    var grade = p.grade_level ? String(p.grade_level[0]).slice(0, 20) : '';
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New Aspire High application',
      body: 'A new application has been submitted' + (lang ? ' (' + lang + ' form' + (grade ? ', grade ' + grade : '') + ')' : '') +
            '.\n\nOpen the responses sheet:\n' + SpreadsheetApp.getActiveSpreadsheet().getUrl() +
            '\n\nStudent details are in the sheet only, not in this email.'
    });
  } catch (err) { /* a mail failure must never lose the application */ }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
