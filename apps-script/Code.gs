/**
 * Status Report Tracker Agent — Apps Script backend
 * ---------------------------------------------------
 * Bind this script to the Google Sheet that holds your status report.
 * Extensions > Apps Script > paste this file (and appsscript.json's
 * settings, via Project Settings > "Show appsscript.json").
 *
 * WHAT THIS DOES
 * 1. Exposes a tiny web API (doGet / doPost) the desktop widget talks to:
 *      GET  ?action=tabs                 -> list of sheet tab names
 *      GET  ?action=columns&tab=NAME     -> header row for that tab
 *      POST { token, tab, values{...} }  -> appends one row to that tab
 * 2. Two time-driven jobs you wire up once via setupTriggers():
 *      sendFridayReminder()      -> 6pm Friday, nudges the team
 *      compileAndSendReport()    -> ~11pm Friday, emails the whole
 *                                    spreadsheet (as .xlsx) to RECIPIENTS
 *
 * Each sheet tab's row 1 must be column headers. Any tab whose name
 * starts with "_" is treated as config/internal and hidden from the
 * widget (handy for a "_Team" tab listing reminder recipients, etc).
 */

// ============================= CONFIG =====================================

// Shared secret the desktop widget must send. Keeps random visitors who
// somehow get the web app URL from writing junk rows. Change this to your
// own random string, and put the same value in the widget's config.json.
var SHARED_SECRET = 'change-me-to-a-random-string';

// Who gets the Friday reminder + the final compiled report.
// Add more addresses any time — no redeploy needed for RECIPIENTS/REMINDER_RECIPIENTS,
// just re-run setupTriggers() is NOT required for this, edits take effect immediately.
var REPORT_RECIPIENTS = [
  'Suryaraj.Rathanasamy@fairview.org',
  'Amulya.Kumar@fairview.org',
  'nav13418@fairview.org'
];

var REMINDER_RECIPIENTS = REPORT_RECIPIENTS; // usually the same list

// Optional: a Microsoft Teams "Incoming Webhook" URL for a channel, if you
// want the Friday reminder posted to Teams as well as emailed. Leave blank
// to skip Teams entirely. (Team > channel > Connectors > Incoming Webhook —
// no admin/IT app registration needed for this, most tenants allow it.)
var TEAMS_WEBHOOK_URL = '';

// Column name (must match a header in every tab) auto-filled with the
// submission time. Leave as-is unless your sheet uses a different name.
var TIMESTAMP_COLUMN = 'Timestamp';

// Names of tabs to exclude from the widget's tab list (besides any
// starting with "_", which are always excluded).
var HIDDEN_TABS = [];

// =============================== API ======================================

function doGet(e) {
  var action = e.parameter.action;
  try {
    if (action === 'tabs') {
      return jsonOut({ ok: true, tabs: listVisibleTabs() });
    }
    if (action === 'columns') {
      var tab = e.parameter.tab;
      return jsonOut({ ok: true, tab: tab, columns: getColumns(tab) });
    }
    return jsonOut({ ok: true, message: 'Status Report Tracker Agent API is running.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (SHARED_SECRET && body.token !== SHARED_SECRET) {
      return jsonOut({ ok: false, error: 'Invalid token.' });
    }
    if (!body.tab) {
      return jsonOut({ ok: false, error: 'Missing "tab".' });
    }

    var sheet = getSheetByName(body.tab);
    if (!sheet) {
      return jsonOut({ ok: false, error: 'Unknown tab: ' + body.tab });
    }

    var headers = getHeaderRow(sheet);
    var values = body.values || {};
    var row = headers.map(function (col) {
      if (col === TIMESTAMP_COLUMN) return new Date();
      return Object.prototype.hasOwnProperty.call(values, col) ? values[col] : '';
    });

    sheet.appendRow(row);
    return jsonOut({ ok: true, message: 'Saved to "' + body.tab + '".' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ============================ SCHEDULED JOBS ==============================

/**
 * Run this once manually (Run > setupTriggers) to install the two weekly
 * triggers. Safe to re-run: it clears old triggers from this project first
 * so you never end up with duplicates.
 */
function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  ScriptApp.newTrigger('sendFridayReminder')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(18) // 6pm, script timezone (see appsscript.json -> Asia/Kolkata)
    .create();

  ScriptApp.newTrigger('compileAndSendReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(23) // 11pm
    .nearMinute(0)
    .create();

  Logger.log('Triggers installed: Friday 6pm reminder, Friday 11pm report.');
}

/** Friday 6pm nudge to fill in the status sheet before the 10pm cutoff. */
function sendFridayReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subject = 'Reminder: submit your status update by 10pm tonight';
  var body =
    'Hi team,\n\n' +
    'Friendly reminder to log your status update for the week before 10pm tonight.\n' +
    'Use the desktop widget, or update the sheet directly here:\n' +
    ss.getUrl() + '\n\n' +
    'The completed sheet goes out automatically at 11pm.\n';

  REMINDER_RECIPIENTS.forEach(function (addr) {
    MailApp.sendEmail(addr, subject, body);
  });

  if (TEAMS_WEBHOOK_URL) {
    UrlFetchApp.fetch(TEAMS_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: '⏰ ' + subject + ' — ' + ss.getUrl() }),
      muteHttpExceptions: true,
    });
  }
}

/** Friday ~11pm: export the whole sheet as .xlsx and email it out. */
function compileAndSendReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = ss.getId();
  var url =
    'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  var token = ScriptApp.getOAuthToken();

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    // Fall back to alerting the recipients that the export failed, rather
    // than failing silently.
    REPORT_RECIPIENTS.forEach(function (addr) {
      MailApp.sendEmail(
        addr,
        'Status report export FAILED',
        'The weekly status report could not be exported automatically. ' +
          'Please check the sheet directly: ' + ss.getUrl()
      );
    });
    return;
  }

  var blob = response.getBlob().setName(ss.getName() + '.xlsx');
  var subject = 'Weekly Status Report — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy');
  var body = 'Attached is this week\'s status report.\n\nSheet: ' + ss.getUrl();

  REPORT_RECIPIENTS.forEach(function (addr) {
    MailApp.sendEmail({
      to: addr,
      subject: subject,
      body: body,
      attachments: [blob],
    });
  });
}

// ============================== HELPERS ===================================

function listVisibleTabs() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) {
      return name.indexOf('_') !== 0 && HIDDEN_TABS.indexOf(name) === -1;
    });
}

function getSheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getHeaderRow(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(function (v) {
    return v !== '';
  });
}

function getColumns(tabName) {
  var sheet = getSheetByName(tabName);
  if (!sheet) return [];
  return getHeaderRow(sheet).filter(function (col) {
    return col !== TIMESTAMP_COLUMN; // widget doesn't ask the user for this
  });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
