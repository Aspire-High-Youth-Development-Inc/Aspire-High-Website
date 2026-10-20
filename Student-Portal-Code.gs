/**
 * Aspire High - student hub content and sign-ups
 * ------------------------------------------------------------------------
 * One Google Sheet runs the student hub. Staff change what students see by
 * editing its tabs (or by posting announcements from the site):
 *   Announcements       Posted at | Audience (All / Seniors) | Title | Message | Link | Posted by
 *   Workshops           Track | Workshop | Date | Time | Where | Description | Link
 *   Links               tutoring_booking and virtual_meeting: type each URL in column C
 *   Senior Checklist    Item | Details
 *   Mentorship Signups  filled in by students from the site (private)
 *
 * Setup (once), using a NEW blank Google Sheet:
 *   1. Extensions -> Apps Script. Delete what is there, paste this whole file, Save.
 *   2. Project Settings (gear) -> Time zone: (GMT-04:00) New York.
 *      Then, still in Project Settings, under "Script properties" click
 *      "Add script property": name  STAFF_KEY , value = a passphrase that only
 *      staff will know. Staff type it on the site to post announcements.
 *   3. Back in the editor, choose "setup" in the function dropdown and click
 *      Run (approve the permissions). It creates the tabs and the headings.
 *   4. Deploy -> New deployment -> type "Web app".
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Copy the Web app URL (ends in /exec).
 *   5. In index.html, paste it between the quotes of  studentPortalEndpoint: ""
 *      and upload index.html to GitHub again.
 *
 * Anyone with the /exec link can READ announcements, workshops and links (it is
 * what students see). Posting an announcement needs the STAFF_KEY. Keep this
 * Google Sheet's own sharing on Restricted: it holds the mentorship sign-ups.
 * After editing this script later: Deploy -> Manage deployments -> edit ->
 * Version: New version.
 */
var TABS = {
  announcements: { name: 'Announcements',       headers: ['Posted at', 'Audience', 'Title', 'Message', 'Link', 'Posted by'] },
  workshops:     { name: 'Workshops',           headers: ['Track', 'Workshop', 'Date', 'Time', 'Where', 'Description', 'Link'] },
  links:         { name: 'Links',               headers: ['Key', 'Label', 'URL'] },
  checklist:     { name: 'Senior Checklist',    headers: ['Item', 'Details'] },
  mentorship:    { name: 'Mentorship Signups',  headers: ['Submitted at', 'Name', 'Email', 'Grade', 'What they want from a mentor'] }
};
var TRACKS = ['Life Skills', 'Freshman', 'Sophomore', 'Junior', 'Senior', 'After School'];
var AUDIENCES = ['All', 'Seniors'];
var MAX_FIELD = 1500;
var HONEYPOT = 'website_hp';

/* ---------- what the site reads ---------- */
function doGet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    var out = { ok: true, announcements: [], workshops: [], links: {}, checklist: [] };
    rows_(ss, TABS.announcements.name).reverse().slice(0, 40).forEach(function (r) {
      if (!r[2] && !r[3]) return;
      out.announcements.push({ when: str_(r[0], tz, 'yyyy-MM-dd'), audience: str_(r[1]) || 'All', title: str_(r[2]), message: str_(r[3]), link: str_(r[4]), by: str_(r[5]) });
    });
    rows_(ss, TABS.workshops.name).forEach(function (r) {
      if (!r[1]) return;
      out.workshops.push({ track: str_(r[0]), title: str_(r[1]), date: str_(r[2], tz, 'yyyy-MM-dd'), time: str_(r[3], tz, 'h:mm a'), where: str_(r[4]), description: str_(r[5]), link: str_(r[6]) });
    });
    rows_(ss, TABS.links.name).forEach(function (r) {
      if (r[0]) out.links[str_(r[0])] = { label: str_(r[1]), url: str_(r[2]) };
    });
    rows_(ss, TABS.checklist.name).forEach(function (r) {
      if (r[0]) out.checklist.push({ item: str_(r[0]), details: str_(r[1]) });
    });
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- what the site sends ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = (e && e.parameter) || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p[HONEYPOT]) return json_({ ok: true });

    if (p.action === 'announce') {
      var key = PropertiesService.getScriptProperties().getProperty('STAFF_KEY');
      if (!key || String(p.key || '') !== key) return json_({ ok: false, error: 'bad key' });
      var title = clean_(p.title), message = clean_(p.message);
      if (!title || !message) return json_({ ok: false, error: 'title and message are required' });
      var audience = AUDIENCES.indexOf(p.audience) === -1 ? 'All' : p.audience;
      var link = /^https?:\/\//i.test(String(p.link || '').trim()) ? clean_(p.link) : '';
      ss.getSheetByName(TABS.announcements.name).appendRow([new Date(), audience, title, message, link, clean_(p.posted_by)]);
      return json_({ ok: true });
    }
    if (p.action === 'mentorship') {
      var name = clean_(p.name), email = clean_(p.email);
      if (!name || !email) return json_({ ok: false, error: 'name and email are required' });
      ss.getSheetByName(TABS.mentorship.name).appendRow([new Date(), name, email, clean_(p.grade), clean_(p.goals)]);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* Run once from the editor (Run -> setup): creates the tabs, headings, drop-downs and starter rows. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function (k) {
    var t = TABS[k];
    var sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, t.headers.length).setFontWeight('bold').setBackground('#0B2D5B').setFontColor('#FFFFFF');
    sh.setColumnWidths(1, t.headers.length, 170);
  });
  var rule = function (list) { return SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build(); };
  ss.getSheetByName(TABS.workshops.name).getRange(2, 1, 500, 1).setDataValidation(rule(TRACKS));
  ss.getSheetByName(TABS.workshops.name).getRange(2, 3, 500, 1).setNumberFormat('yyyy-mm-dd');
  ss.getSheetByName(TABS.workshops.name).getRange(2, 4, 500, 1).setNumberFormat('h:mm AM/PM');
  ss.getSheetByName(TABS.announcements.name).getRange(2, 2, 500, 1).setDataValidation(rule(AUDIENCES));
  ss.getSheetByName(TABS.announcements.name).getRange(2, 1, 500, 1).setNumberFormat('yyyy-mm-dd');

  var links = ss.getSheetByName(TABS.links.name);
  if (links.getLastRow() < 2) links.getRange(2, 1, 2, 3).setValues([
    ['tutoring_booking', 'Book tutoring', ''],
    ['virtual_meeting', 'Join a virtual meeting', '']
  ]);
  var cl = ss.getSheetByName(TABS.checklist.name);
  if (cl.getLastRow() < 2) cl.getRange(2, 1, 8, 2).setValues([
    ['Finish and submit your college applications', ''],
    ['Complete the FAFSA or state financial aid form', ''],
    ['Apply for scholarships', ''],
    ['Ask teachers for recommendation letters', ''],
    ['Send your final transcript to your college', ''],
    ['Choose your college and pay the deposit', ''],
    ['Sign up for housing and orientation', ''],
    ['Update your resume', '']
  ]);
  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
}

/* ---------- helpers ---------- */
function rows_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().filter(function (r) {
    return r.some(function (c) { return String(c).trim() !== ''; });
  });
}
function str_(v, tz, fmt) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, fmt || 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}
/* Answers come from the public, so they are treated as text: a value starting
   with = + - or @ would otherwise run as a formula in Sheets. */
function clean_(s) {
  s = String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, MAX_FIELD);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
