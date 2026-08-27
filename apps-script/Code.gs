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
 *    tabs/columns responses are cached for CACHE_SECONDS since they rarely
 *    change — this is most of the speed-up if things feel slow. Run
 *    clearCache() manually after adding/renaming a tab or column if you
 *    want the change to show up immediately instead of waiting it out.
 * 2. Two time-driven jobs you wire up once via setupTriggers():
 *      sendFridayReminder()      -> 6pm Friday, nudges the team
 *      compileAndSendReport()    -> ~11pm Friday, emails THIS WEEK'S rows
 *                                    only (Mon-Fri), as a .xlsx, to RECIPIENTS
 *
 * Each sheet tab's row 1 must be column headers. Any tab whose name
 * starts with "_" is treated as config/internal and hidden from the
 * widget (handy for a "_Team" tab listing reminder recipients, etc).
 */

// ============================= CONFIG =====================================

// Shared secret the desktop widget must send. Keeps random visitors who
// somehow get the web app URL from writing junk rows. Change this to your
// own random string, and put the same value in the widget's config.json.
var SHARED_SECRET = 'FnKzihGF3xrFOthBuApKgGAQVd1s1aKQ';

// Who gets the Friday reminder + the final compiled report.
// Add more addresses any time — no redeploy needed for RECIPIENTS/REMINDER_RECIPIENTS.
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
// starting with "_", which are always excluded). This one has cumulative
// formulas, not per-entry data, so it doesn't belong in the input flow —
// it's still emailed as-is if you reference it manually, just never shown
// as a "log an entry" choice.
var HIDDEN_TABS = ['Weekly_Monthly_Summary'];

// How long tabs/columns responses are cached (seconds). Higher = faster
// widget, but slower to notice a newly added tab/column. Run clearCache()
// after such a change if you don't want to wait this out.
var CACHE_SECONDS = 300;

// =============================== API ======================================

function doGet(e) {
  var action = e.parameter.action;
  try {
    if (action === 'tabs') {
      return jsonOut({ ok: true, tabs: cached('tabs', listVisibleTabs) });
    }
    if (action === 'columns') {
      var tab = e.parameter.tab;
      var columns = cached('columns:' + tab, function () { return getColumns(tab); });
      return jsonOut({ ok: true, tab: tab, columns: columns });
    }
    if (action === 'nextNumber') {
      var seqTab = e.parameter.tab;
      var seqColumn = e.parameter.column;
      return jsonOut({ ok: true, next: getNextSequenceNumber(seqTab, seqColumn) });
    }
    if (action === 'version') {
      return jsonOut({
        ok: true,
        latest: getConfigValue('LatestVersion'),
        downloadUrl: getConfigValue('DownloadUrl'),
      });
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

    if (body.action === 'sendReportNow') {
      compileAndSendReport();
      return jsonOut({ ok: true, message: 'Report sent.' });
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

/** Run manually any time you add/rename a tab or column and don't want
 * to wait out CACHE_SECONDS for the widget to notice. */
function clearCache() {
  CacheService.getScriptCache().removeAll(
    listVisibleTabsUncached().map(function (t) { return 'columns:' + t; }).concat(['tabs'])
  );
  Logger.log('Cache cleared.');
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

/**
 * Friday ~11pm: build a fresh workbook containing only THIS WEEK's rows
 * (Monday 00:00 through now, per TIMESTAMP_COLUMN) for each visible tab,
 * export it as .xlsx, email it, then discard the temp file.
 */
function compileAndSendReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = getCurrentWeekRange();
  var tempName = ss.getName() + ' — Week of ' +
    Utilities.formatDate(range.start, Session.getScriptTimeZone(), 'MMM d, yyyy');

  var temp = SpreadsheetApp.create(tempName);
  var tabNames = listVisibleTabsUncached();

  tabNames.forEach(function (tabName, i) {
    var source = getSheetByName(tabName);
    var headers = getHeaderRow(source);
    var tsIndex = headers.indexOf(TIMESTAMP_COLUMN);
    var lastRow = source.getLastRow();

    var rows = [];
    if (lastRow > 1) {
      var all = source.getRange(2, 1, lastRow - 1, headers.length).getValues();
      rows = tsIndex === -1
        ? all // no Timestamp column to filter by — include everything
        : all.filter(function (r) {
            var ts = r[tsIndex];
            return ts instanceof Date && ts >= range.start && ts <= range.end;
          });
    }

    var dest = i === 0 ? temp.getSheets()[0].setName(tabName) : temp.insertSheet(tabName);
    dest.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      dest.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
  });

  SpreadsheetApp.flush();
  var blob = DriveApp.getFileById(temp.getId())
    .getAs(MimeType.MICROSOFT_EXCEL)
    .setName(tempName + '.xlsx');

  var subject = 'Weekly Status Report — ' +
    Utilities.formatDate(range.start, Session.getScriptTimeZone(), 'MMM d') + ' to ' +
    Utilities.formatDate(range.end, Session.getScriptTimeZone(), 'MMM d, yyyy');
  var body = 'Attached is this week\'s status report (' +
    Utilities.formatDate(range.start, Session.getScriptTimeZone(), 'MMM d') + '–' +
    Utilities.formatDate(range.end, Session.getScriptTimeZone(), 'MMM d') + ').\n\n' +
    'Full sheet (all history): ' + ss.getUrl();

  REPORT_RECIPIENTS.forEach(function (addr) {
    MailApp.sendEmail({ to: addr, subject: subject, body: body, attachments: [blob] });
  });

  DriveApp.getFileById(temp.getId()).setTrashed(true); // clean up the temp file
}

/** Monday 00:00:00 of the current week through right now. */
function getCurrentWeekRange() {
  var now = new Date();
  var dayIndex = (now.getDay() + 6) % 7; // Monday=0 ... Sunday=6
  var monday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayIndex, 0, 0, 0);
  return { start: monday, end: now };
}

// ============================== HELPERS ===================================

function cached(key, computeFn) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(key);
  if (hit !== null) return JSON.parse(hit);
  var value = computeFn();
  cache.put(key, JSON.stringify(value), CACHE_SECONDS);
  return value;
}

function listVisibleTabsUncached() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) {
      return name.indexOf('_') !== 0 && HIDDEN_TABS.indexOf(name) === -1;
    });
}

function listVisibleTabs() {
  return listVisibleTabsUncached();
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

/** Next value for a numbered column (e.g. "Cleanup Number"): the highest
 * existing number in that column, plus one. Not cached — it has to reflect
 * every submission so far, including ones from a few seconds ago. */
function getNextSequenceNumber(tabName, columnName) {
  var sheet = getSheetByName(tabName);
  if (!sheet) return 1;
  var headers = getHeaderRow(sheet);
  var target = String(columnName).trim().toLowerCase();
  var colIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === target) { colIndex = i; break; }
  }
  if (colIndex === -1) return 1;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;

  var values = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  var max = 0;
  values.forEach(function (row) {
    var n = Number(row[0]);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function getColumns(tabName) {
  var sheet = getSheetByName(tabName);
  if (!sheet) return [];
  return getHeaderRow(sheet).filter(function (col) {
    return col !== TIMESTAMP_COLUMN; // widget doesn't ask the user for this
  });
}

// Name of the hidden config tab holding app-version/download-link info the
// widget checks on launch. Created automatically (with sensible defaults)
// the first time it's needed — no manual setup required. Starts with "_"
// so it's already excluded from the widget's tab list.
var CONFIG_TAB_NAME = '_Config';

function ensureConfigTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_TAB_NAME);
    sheet.getRange(1, 1, 3, 2).setValues([
      ['Key', 'Value'],
      ['LatestVersion', '1.1.0'],
      ['DownloadUrl', ''],
    ]);
    sheet.hideSheet();
  }
  return sheet;
}

/** Bump LatestVersion (and DownloadUrl) in the _Config tab whenever you cut
 * a new installer — every running widget picks it up on next launch and
 * nudges whoever's on an older version to grab the update. No redeploy
 * needed, since this reads live sheet data, not script code. */
function getConfigValue(key) {
  var sheet = ensureConfigTab();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return data[i][1];
  }
  return '';
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
