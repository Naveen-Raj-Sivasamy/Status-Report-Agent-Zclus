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
 *      sendFridayReminder()              -> 6pm Friday, nudges the team
 *      scheduledCompileAndSendReport()   -> ~11pm Friday, emails THIS WEEK'S
 *                                    rows only (Mon-Fri), as a .xlsx, to
 *                                    RECIPIENTS. Thin wrapper around
 *                                    compileAndSendReport() — see below.
 *    Both are skipped automatically on any date listed in the optional
 *    _Holidays tab (see HOLIDAYS_TAB_NAME) — a manual send always goes
 *    through compileAndSendReport() directly and ignores holidays.
 *
 * Each sheet tab's row 1 must be column headers. Any tab whose name
 * starts with "_" is treated as config/internal and hidden from the
 * widget (handy for a "_Team" tab listing reminder recipients, etc).
 *
 * A "Leave" tab is created automatically (see ensureLeaveTab()) the first
 * time the widget asks for the tab list, so people can log their own
 * leave/holiday/WFH days out of the box. New tabs — Leave included — show
 * up in the widget automatically too, but the compiled weekly report only
 * ever pulls from REPORT_TABS below; add a tab's name there yourself if
 * you want it folded into the report as well.
 *
 * Which fields get a dropdown/date-picker/multiselect/etc. — and each
 * dropdown's actual choices — are NOT hardcoded in the widget either:
 * _FieldSchema and _Options (both auto-created, see ensureFieldSchemaTab()
 * / ensureOptionsTab()) drive both. A (Tab, Column) with no _FieldSchema
 * row just falls back to a plain text box, same as always — nothing
 * breaks if you never touch either tab.
 *
 * Which category each tab is grouped under on the landing screen is the
 * same story: _Categories (auto-created, see ensureCategoriesTab()) drives
 * it, and an org that never adds a row there just gets one flat tab list —
 * no category picker screen at all.
 *
 * REQUIRED ONE-TIME SETUP: SHARED_SECRET and the recipient lists are NOT
 * hardcoded in this file (this repo is public — a committed secret/PII
 * would be readable by anyone). Run setupScriptProperties() once from the
 * editor first, or the API will refuse every write with "Invalid token."
 * See that function's comment for details.
 */

// ============================= CONFIG =====================================

// SHARED_SECRET and the recipient lists below used to be hardcoded right
// here — fine for a private script, but this repo is public, which meant
// a real, live auth token and real people's email addresses sat in plain
// text for anyone on the internet to read. Both now come from this Apps
// Script project's own Script Properties instead (Project Settings >
// Script Properties in the editor — or run setupScriptProperties() once,
// see below): never committed, never visible in this file or its history
// going forward. Nothing here needs editing again to rotate the token or
// change who gets emailed; that all happens in Script Properties now.
function getScriptProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}
function getScriptPropList_(key) {
  return getScriptProp_(key).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

// Shared secret the desktop widget must send. Keeps random visitors who
// somehow get the web app URL from writing junk rows. Rotate it any time
// via Script Properties (or setupScriptProperties()) — put the matching
// value in the widget's config.json / config.template.json.
var SHARED_SECRET = getScriptProp_('SHARED_SECRET');

// Who gets the Friday reminder + the final compiled report. Comma-
// separated in the SCRIPT_PROPERTIES value; add/remove addresses any
// time there — no redeploy needed.
var REPORT_RECIPIENTS = getScriptPropList_('REPORT_RECIPIENTS');

// Optional separate list for just the Friday reminder — set a
// REMINDER_RECIPIENTS Script Property only if you want it to differ from
// REPORT_RECIPIENTS; leave it unset to reuse the same list (the common case).
var REMINDER_RECIPIENTS_OVERRIDE_ = getScriptPropList_('REMINDER_RECIPIENTS');
var REMINDER_RECIPIENTS = REMINDER_RECIPIENTS_OVERRIDE_.length ? REMINDER_RECIPIENTS_OVERRIDE_ : REPORT_RECIPIENTS;

// Optional: a Microsoft Teams "Incoming Webhook" URL for a channel, if you
// want the Friday reminder posted to Teams as well as emailed. Leave blank
// to skip Teams entirely. (Team > channel > Connectors > Incoming Webhook —
// no admin/IT app registration needed for this, most tenants allow it.)
var TEAMS_WEBHOOK_URL = '';

// Column name (must match a header in every tab) auto-filled with the
// submission time, IF such a column exists. None of the current tabs have
// one, so this is currently a no-op — safe to ignore unless you add one.
var TIMESTAMP_COLUMN = 'Timestamp';

// The column the weekly report actually filters by: each tab's "Date"
// field (the one the widget's date picker fills in), matched
// case-insensitively. This is what decides "this week's rows".
var REPORT_DATE_COLUMN = 'Date';

// Names of tabs to exclude from the widget's tab list (besides any
// starting with "_", which are always excluded). This one has cumulative
// formulas, not per-entry data, so it doesn't belong in the input flow —
// it's still emailed as-is if you reference it manually, just never shown
// as a "log an entry" choice.
var HIDDEN_TABS = ['Weekly_Monthly_Summary'];

// Tabs actually pulled into the compiled weekly report (email + download)
// — an explicit allowlist, deliberately NOT "whatever tabs happen to be
// visible in the widget right now". New tabs (Leave, or anything else you
// or the team add later) show up in the widget just fine without being
// added here — they just don't show up IN the report unless you
// explicitly opt them in by adding their name to this list.
var REPORT_TABS = ['Daily Status', 'Adhoc_Mails', 'Cleanup_Activities', 'Drupal_Bugs_&_Improvements'];

// How long tabs/columns responses are cached (seconds). Higher = faster
// widget, but slower to notice a newly added tab/column. Run clearCache()
// after such a change if you don't want to wait this out.
var CACHE_SECONDS = 300;

// How long (ms) a save waits for the write lock before giving up and
// telling the widget to retry, instead of queuing silently. Kept well
// under the widget's own per-attempt timeout (see REQUEST_TIMEOUT_MS in
// main.js) so a busy backend fails fast and visibly rather than looking
// like a permanent hang.
var LOCK_WAIT_MS = 10000;

// Optional tab listing full-team holidays, one date per row under a
// "Date" column (same header convention as every other tab). When today
// is listed here, the SCHEDULED Friday reminder and report-send are
// skipped — no point nagging people, or mailing out a near-empty report,
// on a day nobody's expected to be working. Doesn't affect a manual
// on-demand send (the widget's "Submit Report", or the API's
// sendReportNow) — those still always send when a person actually asks
// for them. Starts with "_" so it's already hidden from the widget's tab
// list like _Config, with zero setup needed beyond creating the tab.
var HOLIDAYS_TAB_NAME = '_Holidays';

// =============================== API ======================================

function doGet(e) {
  // `e` (and e.parameter) is only ever missing when someone runs doGet
  // directly from the Apps Script editor's "Run" button instead of hitting
  // the real web app URL — the editor calls it with no arguments. That's a
  // harmless way to end up in the Executions log with a TypeError here, but
  // guard it anyway so it can't crash outside the try/catch below for any
  // other reason either.
  var params = (e && e.parameter) || {};
  var action = params.action;
  try {
    if (action === 'tabs') {
      return jsonOut({ ok: true, tabs: cached('tabs', listVisibleTabs) });
    }
    if (action === 'columns') {
      var tab = params.tab;
      var columns = cached('columns:' + tab, function () { return getColumns(tab); });
      return jsonOut({ ok: true, tab: tab, columns: columns });
    }
    if (action === 'nextNumber') {
      var seqTab = params.tab;
      var seqColumn = params.column;
      return jsonOut({ ok: true, next: getNextSequenceNumber(seqTab, seqColumn) });
    }
    if (action === 'version') {
      return jsonOut({
        ok: true,
        latest: getConfigValue('LatestVersion'),
        downloadUrl: getConfigValue('DownloadUrl'),
      });
    }
    /** Every dropdown/multiselect's option list, read from the _Options
     * tab (see ensureOptionsTab()) — lets you add/remove/reorder choices
     * by editing that tab directly instead of needing a code change and a
     * redeploy for every tweak. */
    if (action === 'options') {
      return jsonOut({ ok: true, options: cached('options', getOptionsMap) });
    }
    /** Which fields on which tabs get a dropdown/date/multiselect/etc. —
     * read from _FieldSchema (see ensureFieldSchemaTab()) instead of
     * being hardcoded in the widget. */
    if (action === 'fieldSchema') {
      return jsonOut({ ok: true, fieldSchema: cached('fieldSchema', getFieldSchemaMap) });
    }
    /** Which category each tab is grouped under on the landing screen —
     * read from _Categories (see ensureCategoriesTab()). Empty ({}) for an
     * org that hasn't configured any, which the widget treats as "show a
     * flat tab list, skip the category picker entirely". */
    if (action === 'categories') {
      return jsonOut({ ok: true, categories: cached('categories', getCategoriesMap) });
    }
    return jsonOut({ ok: true, message: 'Status Report Tracker Agent API is running.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // Was `if (SHARED_SECRET && body.token !== SHARED_SECRET)` — meaning
    // an EMPTY SHARED_SECRET disabled the check entirely (auth optional).
    // That's exactly the state this project is in right until
    // setupScriptProperties() is run for the first time, which made "not
    // configured yet" silently equivalent to "wide open to anyone with the
    // URL" instead of refusing writes until it's actually set up. Fail
    // closed instead: no secret configured means no writes accepted.
    if (!SHARED_SECRET || body.token !== SHARED_SECRET) {
      return jsonOut({ ok: false, error: 'Invalid token.' });
    }

    if (body.action === 'sendReportNow') {
      compileAndSendReport(parseRangeFromRequest(body));
      return jsonOut({ ok: true, message: 'Report sent.' });
    }

    /** Same report as sendReportNow, but to exactly one address instead of
     * the whole REPORT_RECIPIENTS list — for previewing a design/content
     * change (like the HTML email) without emailing the entire team every
     * time. Not exposed anywhere in the widget UI on purpose; call it
     * directly (e.g. from the Apps Script editor's Run, or a one-off
     * request) when you actually want a preview. */
    if (body.action === 'sendTestReport') {
      if (!body.to) return jsonOut({ ok: false, error: 'Missing "to".' });
      // Restricted to REPORT_RECIPIENTS, not an arbitrary address — the
      // shared TOKEN is baked into every installed widget (and has leaked
      // once before, which is why it exists as a Script Property instead
      // of hardcoded now), so letting `to` be attacker-controlled would
      // turn this into a way to mail the full report — attachment
      // included — to any outside address using nothing but that token.
      if (REPORT_RECIPIENTS.indexOf(body.to) === -1) {
        return jsonOut({ ok: false, error: '"to" must be one of REPORT_RECIPIENTS.' });
      }
      var testRange = parseRangeFromRequest(body) || getCurrentWeekRange();
      var testBuilt = buildReportBlob(testRange);
      var testSheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      MailApp.sendEmail({
        to: body.to,
        subject: '[TEST] Status Report — ' + formatRangeLabel(testRange),
        body: reportEmailPlainText(testRange, testBuilt, testSheetUrl),
        htmlBody: reportEmailHtml(testRange, testBuilt, testSheetUrl),
        attachments: [testBuilt.blob],
      });
      return jsonOut({ ok: true, message: 'Test report sent to ' + body.to + '.' });
    }

    if (body.action === 'downloadReport') {
      var file = getReportFileBase64(parseRangeFromRequest(body));
      return jsonOut({ ok: true, fileName: file.fileName, base64: file.base64 });
    }

    /** Lets the widget's Settings screen trigger the same clearCache()
     * that used to require opening the Apps Script editor by hand — e.g.
     * right after reordering tabs in the Sheet, so the new order shows up
     * immediately instead of waiting out CACHE_SECONDS. */
    if (body.action === 'clearCache') {
      clearCache();
      return jsonOut({ ok: true, message: 'Cache cleared.' });
    }

    /** Called by the build-release CI workflow right after it publishes a
     * new installer, so the widget's "update available" banner and its
     * DownloadUrl-driven "Get it" button stay in sync automatically. */
    if (body.action === 'setLatestVersion') {
      if (!body.version) {
        return jsonOut({ ok: false, error: 'Missing "version".' });
      }
      var cfgSheet = ensureConfigTab();
      var cfgData = cfgSheet.getDataRange().getValues();
      var updated = false;
      for (var vi = 1; vi < cfgData.length; vi++) {
        if (String(cfgData[vi][0]).trim() === 'LatestVersion') {
          cfgSheet.getRange(vi + 1, 2).setValue(body.version);
          updated = true;
        }
      }
      return jsonOut({ ok: true, message: 'LatestVersion set to ' + body.version, updated: updated });
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

    // Without a lock, two people submitting within the same second or two
    // (common right before the Friday cutoff) race on appendRow, and on top
    // of that every append forces Sheets to recalculate every formula in
    // the workbook (see the Weekly_Monthly_Summary comment above) — so a
    // pile-up of concurrent doPost calls can each end up waiting on the
    // others for minutes with no feedback, which is what made the widget
    // look permanently stuck on "Saving..." for everyone at once. Bound the
    // wait instead: fail fast with a clear, retryable error rather than
    // letting requests queue silently. (LOCK_WAIT_MS applies only around
    // the actual write — reads elsewhere are unaffected.)
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(LOCK_WAIT_MS)) {
      return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
    }
    try {
      // Checked INSIDE the lock, not before it — two people submitting
      // leave for the same date within the same second or two would
      // otherwise both pass a "is this date free" check done before
      // either had actually written their row. Leave-specific: no other
      // tab has this one-entry-per-date rule.
      if (body.tab === LEAVE_TAB_NAME) {
        var conflictName = findLeaveDateConflict(headers, sheet, values);
        if (conflictName) {
          return jsonOut({ ok: false, error: conflictName + ' has already applied for leave on this date.' });
        }
      }
      sheet.appendRow(row);
    } finally {
      lock.releaseLock();
    }
    return jsonOut({ ok: true, message: 'Saved to "' + body.tab + '".' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/** Client sends start/end as ISO date strings (or omits them entirely for
 * "This Week", which each caller defaults on its own). */
function parseRangeFromRequest(body) {
  if (!body.start || !body.end) return null;
  return { start: new Date(body.start), end: new Date(body.end) };
}

/**
 * Run this ONCE from the editor (select it in the function dropdown, click
 * Run) to grant the Drive + Gmail-sending permissions compileAndSendReport
 * needs. clearCache()/setupTriggers() don't touch those APIs, so running
 * them doesn't trigger the consent screen — this does, safely:
 * it creates a throwaway Drive file and deletes it immediately, and checks
 * your mail quota (a real MailApp call that sends nothing). No report is
 * emailed and nothing is left behind. You'll see a Google permission
 * screen the first time — approve it, then this feature is unblocked.
 */
function authorizeApis() {
  var temp = DriveApp.createFile('status-report-agent-auth-test.txt', 'temporary — safe to ignore/delete');
  DriveApp.getFileById(temp.getId()).setTrashed(true);
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log('Authorized. Remaining mail quota today: ' + quota);
}

/**
 * Run this ONCE from the editor (select it in the function dropdown,
 * click Run) after pasting your real values into the three lines below —
 * saves them into this project's Script Properties, which is where
 * SHARED_SECRET / REPORT_RECIPIENTS / REMINDER_RECIPIENTS actually live
 * from then on (see the CONFIG section up top). Safe to re-run any time
 * you want to rotate the token or update who gets emailed — this
 * overwrites, it doesn't append. You can also skip this function
 * entirely and edit the same three properties directly under Project
 * Settings > Script Properties, if you'd rather not have the values
 * pasted into a function body at all, even briefly and un-committed.
 *
 * IMPORTANT: after running this, delete/blank the pasted values below
 * again before you save/commit this file anywhere — the whole point is
 * that these never end up in source control (this project's included).
 */
function setupScriptProperties() {
  PropertiesService.getScriptProperties().setProperties({
    SHARED_SECRET: 'PASTE-A-NEW-RANDOM-SECRET-HERE',
    REPORT_RECIPIENTS: 'name1@example.com, name2@example.com',
    // REMINDER_RECIPIENTS: 'only-set-this-if-it-should-differ-from-REPORT_RECIPIENTS@example.com',
  });
  Logger.log('Script Properties saved. Blank the values above and re-run (or just leave this function alone) from now on.');
}

/** Run manually any time you add/rename a tab or column and don't want
 * to wait out CACHE_SECONDS for the widget to notice. */
function clearCache() {
  CacheService.getScriptCache().removeAll(
    listVisibleTabsUncached().map(function (t) { return 'columns:' + t; }).concat(['tabs', 'options', 'fieldSchema', 'categories'])
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

  ScriptApp.newTrigger('scheduledCompileAndSendReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(23) // 11pm
    .nearMinute(0)
    .create();

  Logger.log('Triggers installed: Friday 6pm reminder, Friday 11pm report.');
}

/** Rows under HOLIDAYS_TAB_NAME, if that tab exists, are compared against
 * `date` (calendar day only, script timezone — same rule REPORT_DATE_COLUMN
 * filtering already uses elsewhere). Returns false, quietly, if the tab
 * doesn't exist yet — holidays are opt-in, not a hard requirement. */
function isHoliday(date) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOLIDAYS_TAB_NAME);
  if (!sheet) return false;
  var headers = getHeaderRow(sheet);
  var dateIndex = findColumnIndex(headers, REPORT_DATE_COLUMN);
  if (dateIndex === -1) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  var tz = Session.getScriptTimeZone();
  var target = normalizeDateForCompare(date, tz);
  var values = sheet.getRange(2, dateIndex + 1, lastRow - 1, 1).getValues();
  return values.some(function (row) {
    return normalizeDateForCompare(row[0], tz) === target;
  });
}

/** What the Friday 11pm trigger actually calls — skips entirely on a
 * listed holiday. A manual send (widget's "Submit Report", or the API's
 * sendReportNow action) goes straight to compileAndSendReport() below and
 * is never skipped this way, since a person explicitly asked for it. */
function scheduledCompileAndSendReport() {
  if (isHoliday(new Date())) {
    Logger.log('Skipping scheduled report — today is listed in ' + HOLIDAYS_TAB_NAME + '.');
    return;
  }
  compileAndSendReport();
}

/** Friday 6pm nudge to fill in the status sheet before the 10pm cutoff. */
function sendFridayReminder() {
  if (isHoliday(new Date())) {
    Logger.log('Skipping Friday reminder — today is listed in ' + HOLIDAYS_TAB_NAME + '.');
    return;
  }
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
 * Builds a fresh workbook containing only the rows whose REPORT_DATE_COLUMN
 * falls within `range` (for each visible tab), exports it as .xlsx, and
 * discards the temp Drive file. Shared by both the email path
 * (compileAndSendReport) and the download path (getReportFileBase64) so
 * there's exactly one place that knows how to filter/build a report.
 */
function buildReportBlob(range) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempName = ss.getName() + ' — ' + formatRangeLabel(range);

  var temp = SpreadsheetApp.create(tempName);
  // REPORT_TABS, not listVisibleTabsUncached() — the report is scoped to
  // an explicit list of real tabs, not "whatever's visible in the widget
  // right now" (filtered defensively in case one's been renamed/removed).
  var tabNames = REPORT_TABS.filter(function (name) { return !!getSheetByName(name); });
  var counts = []; // [{tab, count}, ...] in the same order as tabNames — used by the email body

  tabNames.forEach(function (tabName, i) {
    var source = getSheetByName(tabName);
    var headers = getHeaderRow(source);
    var dateIndex = findColumnIndex(headers, REPORT_DATE_COLUMN);
    var lastRow = source.getLastRow();

    var rows = [];
    if (lastRow > 1) {
      var all = source.getRange(2, 1, lastRow - 1, headers.length).getValues();
      rows = dateIndex === -1
        ? all // no Date column on this tab — nothing to filter by, include everything
        : all.filter(function (r) {
            var raw = r[dateIndex];
            var d = raw instanceof Date ? raw : new Date(raw);
            return !isNaN(d.getTime()) && d >= range.start && d <= range.end;
          });
    }
    counts.push({ tab: tabName, count: rows.length });

    var dest = i === 0 ? temp.getSheets()[0].setName(tabName) : temp.insertSheet(tabName);
    dest.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      dest.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
  });

  SpreadsheetApp.flush();
  var fileName = tempName + '.xlsx';
  var blob = exportSpreadsheetAsXlsx(temp.getId(), fileName);
  DriveApp.getFileById(temp.getId()).setTrashed(true); // clean up the temp file
  return { blob: blob, fileName: fileName, counts: counts };
}

function formatRangeLabel(range) {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(range.start, tz, 'MMM d, yyyy') + ' to ' +
    Utilities.formatDate(range.end, tz, 'MMM d, yyyy');
}

/**
 * Emails the report for `range` (defaults to the current Mon-Fri work week —
 * this is what the Friday 11pm trigger calls with no argument) to
 * REPORT_RECIPIENTS.
 */
function compileAndSendReport(range) {
  range = range || getCurrentWeekRange();
  var built = buildReportBlob(range);

  var subject = 'Status Report — ' + formatRangeLabel(range);
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  REPORT_RECIPIENTS.forEach(function (addr) {
    MailApp.sendEmail({
      to: addr,
      subject: subject,
      body: reportEmailPlainText(range, built, sheetUrl), // plain-text fallback for clients that don't render HTML
      htmlBody: reportEmailHtml(range, built, sheetUrl),
      attachments: [built.blob],
    });
  });
}

/** Row-count summary table + a styled banner matching the desktop
 * widget's own maroon branding — see EMAIL_BRAND_COLOR below to change
 * it. Kept to inline styles only (no <style> block, no external
 * anything): most corporate mail clients (Outlook especially) strip or
 * ignore a <style> tag, so this is the one approach that actually
 * renders consistently rather than degrading in exactly the inboxes
 * REPORT_RECIPIENTS are most likely reading from. */
function reportEmailHtml(range, built, sheetUrl) {
  var rowsHtml = built.counts.map(function (c) {
    return '<tr>' +
      '<td style="padding:8px 14px;border-bottom:1px solid #e7dade;color:#2a2020;">' + escapeHtml_(c.tab) + '</td>' +
      '<td style="padding:8px 14px;border-bottom:1px solid #e7dade;color:#2a2020;text-align:right;font-weight:700;">' + c.count + '</td>' +
      '</tr>';
  }).join('');
  var total = built.counts.reduce(function (sum, c) { return sum + c.count; }, 0);

  return '' +
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:0 auto;">' +
      '<div style="background:' + EMAIL_BRAND_COLOR + ';border-radius:12px 12px 0 0;padding:20px 24px;">' +
        '<div style="color:#ffffff;font-size:18px;font-weight:700;">Status Report</div>' +
        '<div style="color:#e8c9cf;font-size:13px;font-weight:600;margin-top:4px;">' + escapeHtml_(formatRangeLabel(range)) + '</div>' +
      '</div>' +
      '<div style="border:1px solid #e7dade;border-top:none;border-radius:0 0 12px 12px;padding:20px 24px;">' +
        '<div style="color:#2a2020;font-size:13px;margin-bottom:14px;">' +
          'The attached <b>.xlsx</b> covers ' + total + ' ' + (total === 1 ? 'entry' : 'entries') + ' across ' + built.counts.length + ' ' + (built.counts.length === 1 ? 'tab' : 'tabs') + ':' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;font-size:12.5px;">' + rowsHtml + '</table>' +
        '<div style="margin-top:20px;">' +
          '<a href="' + sheetUrl + '" style="display:inline-block;background:#ffb700;color:#2a2020;font-size:13px;font-weight:700;text-decoration:none;padding:10px 18px;border-radius:9px;">View Full Sheet</a>' +
        '</div>' +
      '</div>' +
    '</div>';
}

/** Same content as reportEmailHtml(), no markup — MailApp.sendEmail's
 * `body` param, shown by mail clients that ignore htmlBody entirely. */
function reportEmailPlainText(range, built, sheetUrl) {
  var lines = built.counts.map(function (c) { return '  ' + c.tab + ': ' + c.count; });
  var total = built.counts.reduce(function (sum, c) { return sum + c.count; }, 0);
  return 'Attached is the status report for ' + formatRangeLabel(range) + '.\n\n' +
    total + ' entries across ' + built.counts.length + ' tabs:\n' + lines.join('\n') + '\n\n' +
    'Full sheet (all history): ' + sheetUrl;
}

// Change this to re-theme the report email — currently the same maroon as
// the desktop widget's own header (var(--brand) in index.html), so the
// two feel like one product instead of the email looking bolted-on.
var EMAIL_BRAND_COLOR = '#6e1b2c';

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Same report as compileAndSendReport, but handed back as base64 instead
 * of emailed — for the app's "Download" option. */
function getReportFileBase64(range) {
  // Was missing the same `range || getCurrentWeekRange()` default
  // compileAndSendReport has right above — meant every "Download Report"
  // with the default "This Week" selection (range comes through as null)
  // crashed with "Cannot read properties of null (reading 'start')"
  // instead of actually downloading anything. Found via a live test
  // during a full test sweep.
  range = range || getCurrentWeekRange();
  var built = buildReportBlob(range);
  return { fileName: built.fileName, base64: Utilities.base64Encode(built.blob.getBytes()) };
}

/**
 * Converts a Google Sheet to a real .xlsx blob. DriveApp's own
 * File.getAs(MimeType.MICROSOFT_EXCEL) can NOT do this — Google Workspace
 * files (Sheets/Docs/Slides) aren't convertible that way, and it fails with
 * "Converting from application/vnd.google-apps.spreadsheet ... is not
 * supported." The documented way is Sheets' own export endpoint, using the
 * script's own OAuth token (no extra scope needed beyond what's already
 * declared in appsscript.json).
 */
function exportSpreadsheetAsXlsx(spreadsheetId, fileName) {
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to export report as .xlsx (HTTP ' + response.getResponseCode() + ')');
  }
  return response.getBlob().setName(fileName);
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
  ensureLeaveTab();
  ensureOptionsTab();
  ensureFieldSchemaTab();
  ensureCategoriesTab();
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) {
      return name.indexOf('_') !== 0 && HIDDEN_TABS.indexOf(name) === -1;
    });
}

// Name of the auto-created "Leave" tab — a normal, visible tab (unlike
// _Config/_Holidays) so it shows up in the widget's tab list like any
// other "log an entry" category, letting people log their own leave/
// holiday/WFH days without you having to set the tab up by hand first.
// Deliberately not in REPORT_TABS: it's the same shape of data as
// everything else, it just isn't part of the weekly work report.
var LEAVE_TAB_NAME = 'Leave';

/** Created once, the first time anything asks for the tab list, if it
 * doesn't already exist — same self-creating pattern as _Config. A no-op
 * every time after that (cheap existence check), and never touches the
 * tab again once it's there, so renaming columns or adding your own is
 * completely safe. */
function ensureLeaveTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(LEAVE_TAB_NAME)) return;
  var sheet = ss.insertSheet(LEAVE_TAB_NAME);
  sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Date', 'Type', 'Reason', 'Week']]);
}

/** One leave entry per date, team-wide — if `values`' Date already has a
 * row on the Leave tab (from anyone, including the same person), returns
 * that row's Name so doPost can reject the new one with a clear message;
 * returns '' when the date's free. Only meaningful called while already
 * holding the write lock (see doPost) — that's what makes the check
 * atomic against two near-simultaneous submissions for the same date. */
function findLeaveDateConflict(headers, sheet, values) {
  var dateIndex = findColumnIndex(headers, 'Date');
  var nameIndex = findColumnIndex(headers, 'Name');
  if (dateIndex === -1) return ''; // no Date column on this tab — nothing to check

  var tz = Session.getScriptTimeZone();
  var target = normalizeDateForCompare(values[headers[dateIndex]], tz);
  if (!target) return ''; // no valid incoming date — let the normal required-field flow handle it

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return '';

  var numCols = Math.max(dateIndex, nameIndex) + 1;
  var rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (normalizeDateForCompare(rows[i][dateIndex], tz) === target) {
      return nameIndex === -1 ? 'Someone' : (String(rows[i][nameIndex]).trim() || 'Someone');
    }
  }
  return '';
}

/** Calendar-day-only comparison (script timezone), same rule isHoliday()
 * and the report's date filtering already use — returns '' for anything
 * that isn't a valid date instead of throwing. */
function normalizeDateForCompare(raw, tz) {
  var d = raw instanceof Date ? raw : new Date(raw);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function listVisibleTabs() {
  return listVisibleTabsUncached();
}

// Name of the auto-created config tab holding every dropdown/multiselect's
// option list — a "List | Option" table, one row per option, grouped by
// the "List" column into whatever the widget asks for by name (SITES,
// ASSIGNEES, "Daily Status.Status", etc. — see field-config.js's
// DEFAULT_OPTIONS keys, which this tab's starting rows exactly mirror).
// Add, remove, reorder, or retype rows here any time — no code change, no
// redeploy, just wait out CACHE_SECONDS or use the widget's Settings ->
// "Refresh tabs & fields" to see it immediately. Starts with "_" so it's
// hidden from the widget's tab list like _Config/_Holidays/_Leave-tab-isn't.
//
// Want this locked down so only specific people can edit it? This tab
// (like any tab) can be protected natively in Sheets — no password to
// manage or leak: select the _Options tab -> right-click -> "Protect
// sheet" -> choose exactly who's allowed to edit. Everyone else sees it
// read-only (or not at all, depending on their sharing access) while the
// widget itself keeps reading it fine either way, since it reads via this
// script's own authorization, not the requesting person's.
var OPTIONS_TAB_NAME = '_Options';

/** Created once, the first time anything asks for the tab list, pre-filled
 * with every option list this app currently ships with — so editing means
 * changing real starting values, not building a list from an empty tab.
 * Never touched again after that first creation, same as ensureLeaveTab(). */
function ensureOptionsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(OPTIONS_TAB_NAME);
  if (sheet) {
    // Was hidden here originally, same as _Config — wrong call for this
    // one specifically: unlike _Config (set-once, never touched again),
    // _Options exists FOR you to keep editing, so hiding it just got in
    // the way of its own purpose. Un-hides it every time this runs in
    // case it's still hidden from an earlier version of this function.
    sheet.showSheet();
    return;
  }
  sheet = ss.insertSheet(OPTIONS_TAB_NAME);

  var defaults = {
    SITES: ['MHF', 'FV', 'Peds', 'GI', 'Specialty Pharmacy', 'All'],
    REQUESTERS: ['Cassandra', 'Tseten', 'Grant', 'Tammy', 'Naveen'],
    ASSIGNEES: ['Naveen', 'Surya', 'Amulya', 'Cassandra', 'Tseten', 'Grant', 'Tammy', 'Lucy', 'Erika'],
    LEAVE_NAMES: ['Amulya Kumar', 'Suryaraj', 'Naveen Raj'],
    'Daily Status.Status': ['Done', 'Pending', 'In Progress', 'Open'],
    'Daily Status.Priority': ['High', 'Medium', 'Low'],
    'Cleanup_Activities.Volume': ['Large', 'Medium', 'Small'],
    'Drupal_Bugs_&_Improvements.Type': ['Bug', 'Fix', 'Suggestion'],
    'Leave.Type': ['Vacation', 'Sick', 'WFH', 'Holiday', 'Other'],
  };

  var rows = [['List', 'Option']];
  Object.keys(defaults).forEach(function (list) {
    defaults[list].forEach(function (option) {
      rows.push([list, option]);
    });
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

/** Groups _Options' rows by the "List" column, in the order they appear
 * on the sheet (so reordering rows there reorders the dropdown too). A
 * row with a blank List or Option is skipped rather than producing a
 * broken/empty entry. */
function getOptionsMap() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPTIONS_TAB_NAME);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  values.forEach(function (row) {
    var list = String(row[0]).trim();
    var option = String(row[1]).trim();
    if (!list || !option) return;
    if (!map[list]) map[list] = [];
    map[list].push(option);
  });
  return map;
}

// Which form fields get a dropdown/date-picker/multiselect/etc., instead
// of the plain-text default, per (Tab, Column) — used to live entirely in
// the widget's own code (FIELD_CONFIG in field-config.js). Now it's a
// sheet tab like everything else configurable, so a new organization
// setting this app up from scratch starts with NOTHING hardcoded about
// what fields exist — every tab just gets plain text boxes until rows are
// added here, same graceful fallback fieldSpecFor() already had for any
// (Tab, Column) it didn't recognize.
var FIELD_SCHEMA_TAB_NAME = '_FieldSchema';

/** Created once, seeded with exactly what FIELD_CONFIG used to hardcode —
 * this is a MIGRATION for an existing deployment (yours): it moves that
 * data from code into your sheet without touching any of your other
 * tabs' actual submitted rows, so the app keeps behaving identically
 * while the "no hardcoded data" rule now genuinely holds. A brand-new
 * organization starting from zero just gets an empty tab with headers —
 * nothing pre-filled for a team whose tabs don't look like yours. */
function ensureFieldSchemaTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(FIELD_SCHEMA_TAB_NAME)) return;
  var sheet = ss.insertSheet(FIELD_SCHEMA_TAB_NAME);
  sheet.getRange(1, 1, 1, 5).setValues([['Tab', 'Column', 'Type', 'OptionsKey', 'BasedOn']]);

  // [Tab, Column, Type, OptionsKey, BasedOn] — only seeded if these exact
  // tabs already exist in THIS spreadsheet (checked below), so a fresh
  // org's blank _FieldSchema tab doesn't get rows referencing tabs they
  // don't have.
  var migration = [
    ['Daily Status', 'Site', 'select', 'SITES', ''],
    ['Daily Status', 'Date', 'date', '', ''],
    ['Daily Status', 'Week', 'week-auto', '', 'Date'],
    ['Daily Status', 'Status', 'select', 'Daily Status.Status', ''],
    ['Daily Status', 'Priority', 'select', 'Daily Status.Priority', ''],
    ['Daily Status', 'Assigned to', 'multiselect', 'ASSIGNEES', ''],
    ['Adhoc_Mails', 'Requester', 'select', 'REQUESTERS', ''],
    ['Adhoc_Mails', 'Site', 'select', 'SITES', ''],
    ['Adhoc_Mails', 'Date', 'date', '', ''],
    ['Adhoc_Mails', 'Week', 'week-auto', '', 'Date'],
    ['Adhoc_Mails', 'Assigned to', 'multiselect', 'ASSIGNEES', ''],
    ['Cleanup_Activities', 'Cleanup Number', 'sequence', '', ''],
    ['Cleanup_Activities', 'Volume', 'select', 'Cleanup_Activities.Volume', ''],
    ['Cleanup_Activities', 'Requester', 'select', 'REQUESTERS', ''],
    ['Cleanup_Activities', 'Site Impacted', 'select', 'SITES', ''],
    ['Cleanup_Activities', 'Date', 'date', '', ''],
    ['Cleanup_Activities', 'Week', 'week-auto', '', 'Date'],
    ['Cleanup_Activities', 'Assigned to', 'multiselect', 'ASSIGNEES', ''],
    ['Drupal_Bugs_&_Improvements', 'Website', 'select', 'SITES', ''],
    ['Drupal_Bugs_&_Improvements', 'Date', 'date', '', ''],
    ['Drupal_Bugs_&_Improvements', 'Week', 'week-auto', '', 'Date'],
    ['Drupal_Bugs_&_Improvements', 'Type', 'select', 'Drupal_Bugs_&_Improvements.Type', ''],
    ['Drupal_Bugs_&_Improvements', 'Assigned to', 'multiselect', 'ASSIGNEES', ''],
    ['Leave', 'Name', 'select', 'LEAVE_NAMES', ''],
    ['Leave', 'Date', 'date', '', ''],
    ['Leave', 'Week', 'week-auto', '', 'Date'],
    ['Leave', 'Type', 'select', 'Leave.Type', ''],
  ];
  var existingTabNames = ss.getSheets().map(function (s) { return s.getName(); });
  var rows = migration.filter(function (r) { return existingTabNames.indexOf(r[0]) !== -1; });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

/** Groups _FieldSchema's rows into { Tab: { Column: {type, optionsKey,
 * basedOn} } }, matching the shape the widget's fieldSpecFor() used to
 * get from the hardcoded FIELD_CONFIG. A row with a blank Tab, Column, or
 * Type is skipped. */
function getFieldSchemaMap() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIELD_SCHEMA_TAB_NAME);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var map = {};
  values.forEach(function (row) {
    var tab = String(row[0]).trim();
    var column = String(row[1]).trim();
    var type = String(row[2]).trim();
    if (!tab || !column || !type) return;
    var optionsKey = String(row[3]).trim();
    var basedOn = String(row[4]).trim();
    if (!map[tab]) map[tab] = {};
    map[tab][column] = { type: type, optionsKey: optionsKey || undefined, basedOn: basedOn || undefined };
  });
  return map;
}

// Which landing-screen category each tab is grouped under (the "Report
// Generator" / "Team Management" picker) — used to live entirely in the
// widget's own code (CATEGORIES in field-config.js). Now it's a sheet tab
// like everything else configurable, so a new organization starts with NO
// grouping at all: the widget just shows every tab in one flat list until
// rows are added here. Add, remove, reorder, or retype rows here any time —
// no code change, no redeploy needed, just wait out CACHE_SECONDS or use
// Settings -> "Refresh tabs & fields".
var CATEGORIES_TAB_NAME = '_Categories';

/** Created once, seeded with exactly what CATEGORIES used to hardcode —
 * this is a MIGRATION for an existing deployment (yours): it moves that
 * data from code into your sheet without touching any of your other tabs'
 * actual submitted rows, so the app's landing screen keeps looking
 * identical while the "no hardcoded data" rule now genuinely holds. A
 * brand-new organization starting from zero just gets an empty tab with
 * headers — nothing pre-filled for a team whose tabs don't look like
 * yours, and the widget falls back to one flat tab list with no category
 * screen at all when this tab has no rows (see groupTabsIntoCategories()
 * in field-config.js). */
function ensureCategoriesTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(CATEGORIES_TAB_NAME)) return;
  var sheet = ss.insertSheet(CATEGORIES_TAB_NAME);
  sheet.getRange(1, 1, 1, 2).setValues([['Category', 'Tab']]);

  // [Category, Tab] — only seeded if the tab already exists in THIS
  // spreadsheet (checked below), so a fresh org's blank _Categories tab
  // doesn't get rows referencing tabs they don't have.
  var migration = [
    ['Report Generator', 'Daily Status'],
    ['Report Generator', 'Adhoc_Mails'],
    ['Report Generator', 'Cleanup_Activities'],
    ['Report Generator', 'Drupal_Bugs_&_Improvements'],
    ['Team Management', 'Leave'],
  ];
  var existingTabNames = ss.getSheets().map(function (s) { return s.getName(); });
  var rows = migration.filter(function (r) { return existingTabNames.indexOf(r[1]) !== -1; });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

/** Groups _Categories' rows by the "Category" column, in the order they
 * appear on the sheet (so reordering rows there reorders the landing
 * screen too). A row with a blank Category or Tab is skipped. Returns {}
 * (no categories at all) for a brand-new org that hasn't added any rows —
 * that's the signal the widget uses to skip the category picker entirely. */
function getCategoriesMap() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORIES_TAB_NAME);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  values.forEach(function (row) {
    var category = String(row[0]).trim();
    var tab = String(row[1]).trim();
    if (!category || !tab) return;
    if (!map[category]) map[category] = [];
    map[category].push(tab);
  });
  return map;
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
function findColumnIndex(headers, columnName) {
  var target = String(columnName).trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === target) return i;
  }
  return -1;
}

function getNextSequenceNumber(tabName, columnName) {
  var sheet = getSheetByName(tabName);
  if (!sheet) return 1;
  var headers = getHeaderRow(sheet);
  var colIndex = findColumnIndex(headers, columnName);
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
