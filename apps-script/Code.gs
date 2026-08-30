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
 * 2. Three time-driven jobs you wire up once via setupTriggers():
 *      sendFridayReminder()              -> 6pm Friday, nudges the team
 *      scheduledCompileAndSendReport()   -> ~11pm Friday, emails THIS WEEK'S
 *                                    rows only (Mon-Fri), as a .xlsx, to
 *                                    RECIPIENTS. Thin wrapper around
 *                                    compileAndSendReport() — see below.
 *      checkUrgentMail()                 -> every 15 min, flags unread mail
 *                                    from URGENT_SENDERS or matching
 *                                    URGENT_KEYWORDS (both in _Options) as
 *                                    a red badge in URGENT_ALERT_FOR's
 *                                    widget. Can only ever read THIS
 *                                    script's own authorized Gmail inbox —
 *                                    see URGENT_ALERT_FOR's comment.
 *    The two Friday jobs are skipped automatically on any date listed in
 *    the optional _Holidays tab (see HOLIDAYS_TAB_NAME) — a manual send
 *    always goes through compileAndSendReport() directly and ignores
 *    holidays. checkUrgentMail() runs regardless of holidays.
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

// checkUrgentMail() (see SCHEDULED JOBS below) can only ever read the
// Gmail inbox of whoever this Apps Script project is authorized under —
// there's no way for a shared backend like this to read each teammate's
// own separate inbox. So the flag it produces is only ever "this one
// person's inbox has something urgent", and the desktop widget only
// shows the red badge when its own YOUR_NAME matches this exactly
// (case-insensitive). Blank means nobody's widget ever shows it.
var URGENT_ALERT_FOR = 'Naveen Raj';

// Base name for the two Gmail labels checkUrgentMail() manages:
//   <this>/Scanned  — applied to EVERY thread it looks at, match or not.
//                      This is the dedup marker: a thread with this label
//                      is never re-evaluated on a later run.
//   <this>/Flagged  — applied only to threads that actually matched. A
//                      thread only counts toward the badge while it's
//                      BOTH labeled Flagged AND still unread — reading it
//                      in Gmail is what clears it, no separate "dismiss"
//                      step needed in the app.
var URGENT_LABEL_NAME = 'StatusReportAgent';

// How far back checkUrgentMail() looks for new (not-yet-labeled) mail on
// each run. Wide enough that a 15-minute trigger never misses a message
// that arrived since the last run, without re-scanning your whole inbox.
var URGENT_LOOKBACK = 'newer_than:1d';

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
    /** Polled by the widget every few minutes for the urgent-mail red
     * badge (see checkUrgentMail(), which is what actually updates this —
     * this endpoint just reads the last result, it never scans Gmail
     * itself, so it's cheap and fast regardless of poll frequency).
     * alertFor tells the widget WHOSE flag this is, since it can only
     * ever be one person's inbox — the widget only shows the badge when
     * its own YOUR_NAME matches. */
    if (action === 'urgentStatus') {
      var raw = PropertiesService.getScriptProperties().getProperty('URGENT_STATUS');
      var status = raw ? JSON.parse(raw) : { count: 0, updatedAt: null };
      return jsonOut({ ok: true, alertFor: URGENT_ALERT_FOR, count: status.count, updatedAt: status.updatedAt });
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
    listVisibleTabsUncached().map(function (t) { return 'columns:' + t; }).concat(['tabs', 'options'])
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

  ScriptApp.newTrigger('checkUrgentMail')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('Triggers installed: Friday 6pm reminder, Friday 11pm report, urgent-mail check every 15 min.');
}

/**
 * Scans for new unread mail matching URGENT_SENDERS/URGENT_KEYWORDS (both
 * from _Options — see ensureOptionsTab()), labels every message it looks
 * at (matched or not) with URGENT_LABEL_NAME so it's never re-evaluated,
 * then recounts how many *still-unread, already-labeled* messages exist —
 * that recount, not the scan itself, is what decides the badge. Reading a
 * flagged email in Gmail is what clears it; there's no separate "dismiss"
 * button in the app. Writes { count, updatedAt } to a Script Property the
 * widget polls via GET ?action=urgentStatus.
 *
 * Requires the gmail.modify scope in appsscript.json (already added) and
 * your own one-time re-authorization — run this function once manually
 * from the editor and approve the Gmail permission prompt.
 */
function checkUrgentMail() {
  if (!URGENT_ALERT_FOR) return; // feature off if nobody's configured to receive it

  var scannedLabelName = URGENT_LABEL_NAME + '/Scanned';
  var flaggedLabelName = URGENT_LABEL_NAME + '/Flagged';

  var options = getOptionsMap();
  var senders = (options.URGENT_SENDERS || []).map(function (s) { return s.toLowerCase(); });
  var keywords = (options.URGENT_KEYWORDS || []).map(function (k) { return k.toLowerCase(); });

  if (senders.length || keywords.length) {
    var scannedLabel = GmailApp.getUserLabelByName(scannedLabelName) || GmailApp.createLabel(scannedLabelName);
    var flaggedLabel = GmailApp.getUserLabelByName(flaggedLabelName) || GmailApp.createLabel(flaggedLabelName);
    // Excludes anything already Scanned (matched or not) — that's the
    // dedup — NOT anything already Flagged, since a Flagged-but-unread
    // thread is also always Scanned already and so is excluded anyway.
    var threads = GmailApp.search('is:unread ' + URGENT_LOOKBACK + ' -label:"' + scannedLabelName + '"');
    threads.forEach(function (thread) {
      var messages = thread.getMessages();
      var matched = messages.some(function (message) {
        var from = String(message.getFrom() || '').toLowerCase();
        var subject = String(message.getSubject() || '').toLowerCase();
        var snippet = String(message.getPlainBody() || '').slice(0, 500).toLowerCase();
        var fromMatch = senders.some(function (s) { return from.indexOf(s) !== -1; });
        var keywordMatch = keywords.some(function (k) {
          return subject.indexOf(k) !== -1 || snippet.indexOf(k) !== -1;
        });
        return fromMatch || keywordMatch;
      });
      if (matched) thread.addLabel(flaggedLabel);
      thread.addLabel(scannedLabel); // evaluated either way — never scan this thread again
    });
  }

  var stillFlagged = senders.length || keywords.length
    ? GmailApp.search('is:unread label:"' + flaggedLabelName + '"').length
    : 0;

  PropertiesService.getScriptProperties().setProperty('URGENT_STATUS', JSON.stringify({
    count: stillFlagged,
    updatedAt: new Date().toISOString(),
  }));
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
  var target = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  var values = sheet.getRange(2, dateIndex + 1, lastRow - 1, 1).getValues();
  return values.some(function (row) {
    var raw = row[0];
    var d = raw instanceof Date ? raw : new Date(raw);
    return !isNaN(d.getTime()) && Utilities.formatDate(d, tz, 'yyyy-MM-dd') === target;
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
  return { blob: blob, fileName: fileName };
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
  var body = 'Attached is the status report for ' + formatRangeLabel(range) + '.\n\n' +
    'Full sheet (all history): ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();

  REPORT_RECIPIENTS.forEach(function (addr) {
    MailApp.sendEmail({ to: addr, subject: subject, body: body, attachments: [built.blob] });
  });
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
  var isNew = !sheet;
  if (isNew) {
    sheet = ss.insertSheet(OPTIONS_TAB_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['List', 'Option']]);
    sheet.hideSheet();
  }

  // Not just "create if missing" — also backfills any DEFAULT_LISTS key
  // that doesn't have real rows yet, so adding a new configurable list to
  // this app later (like URGENT_KEYWORDS below) appends its starting
  // values into your *already-existing* tab instead of silently doing
  // nothing because the tab was already there. Never touches a list that
  // already has at least one row — your edits to SITES/ASSIGNEES/etc.
  // are never overwritten, only genuinely-new lists get added.
  var existing = isNew ? {} : getOptionsMap();
  var appendRows = [];
  Object.keys(DEFAULT_LISTS).forEach(function (list) {
    if (existing[list] && existing[list].length) return;
    DEFAULT_LISTS[list].forEach(function (option) {
      appendRows.push([list, option]);
    });
  });
  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, 2).setValues(appendRows);
  }
}

// The starting values ensureOptionsTab() seeds _Options with — both for a
// brand-new tab and for backfilling any key added to this app after your
// tab already existed (see the comment above). URGENT_SENDERS isn't here
// deliberately: there's no sensible default VIP list, so it's just not
// created until you add rows for it yourself.
var DEFAULT_LISTS = {
  SITES: ['MHF', 'FV', 'Peds', 'GI', 'Specialty Pharmacy', 'All'],
  REQUESTERS: ['Cassandra', 'Tseten', 'Grant', 'Tammy', 'Naveen'],
  ASSIGNEES: ['Naveen', 'Surya', 'Amulya', 'Cassandra', 'Tseten', 'Grant', 'Tammy', 'Lucy', 'Erika'],
  LEAVE_NAMES: ['Amulya Kumar', 'Suryaraj', 'Naveen Raj'],
  'Daily Status.Status': ['Done', 'Pending', 'In Progress', 'Open'],
  'Daily Status.Priority': ['High', 'Medium', 'Low'],
  'Cleanup_Activities.Volume': ['Large', 'Medium', 'Small'],
  'Drupal_Bugs_&_Improvements.Type': ['Bug', 'Fix', 'Suggestion'],
  'Leave.Type': ['Vacation', 'Sick', 'WFH', 'Holiday', 'Other'],
  URGENT_KEYWORDS: ['urgent', 'asap', 'emergency', 'immediately', 'critical'],
};

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
