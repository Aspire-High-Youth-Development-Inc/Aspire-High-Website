
/**
 * Aspire High — application responses
 * ------------------------------------------------------------------------
 * Receives English and Spanish student applications and saves each
 * submission as one row in the "Responses" tab of this spreadsheet.
 *
 * SETUP
 * 1. Open the response spreadsheet.
 * 2. Select Extensions -> Apps Script.
 * 3. Back up Code.gs, then replace its contents with this code.
 * 4. Save the project.
 * 5. Confirm the project time zone is America/New_York.
 * 6. Keep the existing web app deployment URL if it uses this project.
 *    After changing this code, deploy a NEW VERSION:
 *    Deploy -> Manage deployments -> Edit -> New version -> Deploy.
 *
 * Keep the spreadsheet sharing set to Restricted.
 */

var SHEET_NAME = 'Responses';

/* Optional notification email. Leave blank to disable notifications. */
var NOTIFY_EMAIL = '';

/*
 * Form field name -> spreadsheet column heading.
 *
 * IMPORTANT:
 * Each field name on the left must match the HTML name="..." attribute.
 * Unrecognized fields are still saved automatically in additional columns.
 */
var COLUMNS = [
  ['submitted_at',              'Submitted at'],
  ['form_language',             'Form language'],

  // Student information
  ['student_first_name',        'Student first name'],
  ['student_last_name',         'Student last name'],
  ['preferred_name',            'Preferred name'],
  ['student_dob',               'Student date of birth'],
  ['grade_level',               'Grade level'],
  ['student_personal_email',    'Student personal email'],
  ['student_email',             'Student school email'],
  ['student_phone',             'Student phone'],

  // Student address
  ['student_street',            'Student street address'],
  ['student_apartment',         'Student apartment or unit'],
  ['student_city',              'Student city'],
  ['student_state',             'Student state'],
  ['student_zip',               'Student ZIP code'],

  // Parent or guardian
  ['parent_name',               'Parent/guardian name'],
  ['parent_relationship',       'Relationship to student'],
  ['parent_email',              'Parent/guardian email'],
  ['parent_phone',              'Parent/guardian phone'],
  ['parent_phone_type',         'Parent phone type'],
  ['contact_method',            'Preferred contact methods'],

  // Household information
  ['household_size',            'Household size'],
  ['parent_education',          'Highest parent/guardian education'],

  // Community demographics
  ['first_gen',                 'First-generation college'],
  ['school_lunch',              'School lunch'],
  ['english_learner',           'English learner'],
  ['race_ethnicity',            'Race/ethnicity'],
  ['nationality',               'Nationality'],

  // Student goals and academic needs
  ['goals',                     'Student goals'],
  ['after_high_school',         'Plan after high school'],
  ['gpa',                       'GPA'],
  ['academic_support_areas',    'Academic support needed'],
  ['year_goal',                 'Goal for this year'],

  // Parent/guardian engagement
  ['parent_session_time',       'Preferred parent session times'],

  // Emergency and accessibility information
  ['emergency_contact_name',    'Emergency contact name'],
  ['emergency_relationship',    'Emergency contact relationship'],
  ['emergency_contact_phone',   'Emergency contact phone'],
  ['accessibility_needs',       'Accessibility needs'],
  ['student_iep',               'Student has an IEP'],

  // Referral and consent
  ['heard_about',               'How they heard about us'],
  ['consent_participation',     'Participation consent'],
  ['consent_evaluation',        'Evaluation consent'],
  ['photo_release',             'Photo/video release'],

  // Signatures
  ['student_signature',         'Student signature'],
  ['student_signature_date',    'Student signature date'],
  ['signature',                 'Parent/guardian signature'],
  ['signature_date',            'Parent/guardian signature date']
];

/*
 * Keep these values as text to preserve leading zeroes and punctuation.
 * Add additional field names here if necessary.
 */
var TEXT_FIELDS = {
  student_zip: 1,
  zip_code: 1,
  student_phone: 1,
  parent_phone: 1,
  emergency_contact_phone: 1
};

var HONEYPOT = 'website';
var MAX_FIELD = 3000;
var MAX_FIELDS = 120;

/**
 * Receives form submissions.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    // Use parameters so multiple checkbox selections are preserved.
    var p = (e && e.parameters) || {};

    if (Object.keys(p).length > MAX_FIELDS) {
      return json_({ ok: false, error: 'Too many fields.' });
    }

    // Silently discard bot submissions that fill in the hidden honeypot.
    if (
      p[HONEYPOT] &&
      String(p[HONEYPOT][0] || '').trim() !== ''
    ) {
      return json_({ ok: true });
    }

    var sheet = sheet_();
    var headers = ensureHeaders_(sheet, p);
    var fieldByHeader = {};

    COLUMNS.forEach(function (column) {
      fieldByHeader[column[1]] = column[0];
    });

    var row = headers.map(function (header) {
      var field = fieldByHeader[header] || header;

      if (field === 'submitted_at') {
        return new Date();
      }

      var values = p[field];

      if (!values || !values.length) {
        return '';
      }

      // Multiple checkbox values are stored together in one cell.
      var answer = clean_(values.join(', '));

      // Preserve ZIP codes and phone numbers as text.
      if (
        TEXT_FIELDS[field] &&
        answer !== '' &&
        answer.charAt(0) !== "'"
      ) {
        answer = "'" + answer;
      }

      return answer;
    });

    sheet.appendRow(row);

    // Format submission timestamp.
    var lastRow = sheet.getLastRow();
    var submittedAtColumn = headers.indexOf('Submitted at') + 1;

    if (submittedAtColumn > 0) {
      sheet.getRange(lastRow, submittedAtColumn)
        .setNumberFormat('yyyy-mm-dd hh:mm');
    }

    notify_(p);

    return json_({ ok: true });

  } catch (err) {
    return json_({
      ok: false,
      error: String(err && err.message || err)
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

/**
 * Opening the deployed /exec URL in a browser checks whether the service
 * is responding. It does not display application data.
 */
function doGet() {
  return json_({
    ok: true,
    service: 'Aspire High applications',
    accepts: 'POST'
  });
}

/**
 * Optional setup function.
 * Run from the Apps Script editor to create the Responses tab and headings.
 */
function setup() {
  var sheet = sheet_();
  ensureHeaders_(sheet, {});

  sheet.setFrozenRows(1);

  if (sheet.getLastColumn() > 0) {
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setFontWeight('bold')
      .setBackground('#0B2D5B')
      .setFontColor('#FFFFFF');
  }
}

/**
 * Gets or creates the Responses tab.
 */
function sheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'No active spreadsheet. Open Apps Script from the response spreadsheet.'
    );
  }

  return spreadsheet.getSheetByName(SHEET_NAME) ||
    spreadsheet.insertSheet(SHEET_NAME);
}

/**
 * Ensures the sheet has every expected column.
 *
 * Existing headings are retained. New expected headings and new form fields
 * are added to the right. Existing response data is not deleted.
 */
function ensureHeaders_(sheet, p) {
  var lastColumn = sheet.getLastColumn();
  var headers = lastColumn
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String)
    : [];

  if (headers.every(function (header) {
    return header === '';
  })) {
    headers = [];
  }

  var wanted = COLUMNS.map(function (column) {
    return column[1];
  });

  var knownFields = {};

  COLUMNS.forEach(function (column) {
    knownFields[column[0]] = true;
  });

  // Include any submitted field that is not in the mapping.
  Object.keys(p || {}).forEach(function (field) {
    if (
      !knownFields[field] &&
      field !== HONEYPOT &&
      /^[A-Za-z0-9_]{1,60}$/.test(field)
    ) {
      wanted.push(field);
    }
  });

  var added = false;

  wanted.forEach(function (heading) {
    if (headers.indexOf(heading) === -1) {
      headers.push(heading);
      added = true;
    }
  });

  if (added) {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]);
  }

  return headers;
}

/**
 * Cleans user input and prevents spreadsheet formula injection.
 */
function clean_(value) {
  var result = String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .slice(0, MAX_FIELD);

  if (/^[=+\-@\t\r]/.test(result)) {
    return "'" + result;
  }

  return result;
}

/**
 * Optional notification email.
 * Sends no student details, only a notice and spreadsheet link.
 */
function notify_(p) {
  if (!NOTIFY_EMAIL) {
    return;
  }

  try {
    var language = p.form_language
      ? String(p.form_language[0] || '')
      : '';

    var grade = p.grade_level
      ? String(p.grade_level[0] || '').slice(0, 20)
      : '';

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New Aspire High application',
      body:
        'A new application has been submitted' +
        (language
          ? ' (' + language + ' form' +
            (grade ? ', grade ' + grade : '') + ')'
          : '') +
        '.\n\nOpen the responses spreadsheet:\n' +
        SpreadsheetApp.getActiveSpreadsheet().getUrl() +
        '\n\nStudent details are in the spreadsheet only, not in this email.'
    });

  } catch (err) {
    // Email failure must not prevent the application from being saved.
  }
}

/**
 * Returns a JSON response.
 */
function json_(object) {
  return ContentService
    .createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}
