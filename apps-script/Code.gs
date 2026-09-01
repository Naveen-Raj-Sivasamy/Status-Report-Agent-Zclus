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
 * ever pulls from _Config's ReportTabs (see getReportTabs()); add a tab's
 * name there — via the Sheet or the widget's "Manage Fields & Options"
 * screen — if you want it folded into the report as well.
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
 * _Features (auto-created, see ensureFeaturesTab()) is a plain step-by-step
 * guide to all of the above, written directly into the sheet for whoever's
 * editing it day to day — nothing reads it programmatically.
 *
 * A "Weekly_Connect" tab is created automatically (see
 * ensureWeeklyConnectTab()) the same way Leave is — tracks questions/
 * issues raised through the week, posts a running list to Teams (if a
 * webhook URL is set — see getTeamsWebhookUrl()) scoped to the current
 * Wed-to-Wed window, and supports editing a ticket's Status/Comments in
 * place from the widget's "View & Update Tickets" screen — the one place
 * in this app that edits an existing row instead of appending.
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

// Who gets the Friday reminder + the final compiled report. Editable from
// _Config (ReportRecipients/ReminderRecipients — see getReportRecipients()
// / getReminderRecipients() near ensureConfigTab()) or the widget's
// in-app "Manage Fields & Options" screen; the REPORT_RECIPIENTS /
// REMINDER_RECIPIENTS Script Properties below are the ORIGINAL mechanism
// (this repo is public, so these could never be hardcoded here directly —
// see the comment above) and still work as a fallback for any deployment
// that hasn't moved to _Config for this yet.

// Optional: a Microsoft Teams "Incoming Webhook" URL for a channel — used
// for the Friday reminder, and for Weekly Connect's running ticket post
// (see postWeeklyConnectToTeams()). Leave unset to skip Teams entirely.
// (Team > channel > Connectors > Incoming Webhook — no admin/IT app
// registration needed for this, most tenants allow it.)
//
// Editable from _Config's TeamsWebhookUrl row, or the in-app "Manage
// Fields & Options" -> Report Settings screen — see getTeamsWebhookUrl()
// near ensureConfigTab(), same tiered pattern as ReportRecipients: _Config
// wins if set, else this Script Property is the fallback (kept working
// for a deployment that set it up before _Config supported this). Unlike
// SHARED_SECRET — which genuinely can't live in a sheet cell without
// defeating its own purpose — a Teams webhook only grants posting into
// one specific channel, so making it Sheet/App-configurable is a
// reasonable trade for "any client using this app can set/change it
// themselves, no code touched."
var TEAMS_WEBHOOK_URL = getScriptProp_('TEAMS_WEBHOOK_URL');

// Column name (must match a header in every tab) auto-filled with the
// submission time, IF such a column exists. None of the current tabs have
// one, so this is currently a no-op — safe to ignore unless you add one.
var TIMESTAMP_COLUMN = 'Timestamp';

// The column the weekly report actually filters by: each tab's "Date"
// field (the one the widget's date picker fills in), matched
// case-insensitively. This is what decides "this week's rows".
var REPORT_DATE_COLUMN = 'Date';

// Names of tabs to exclude from the widget's tab list (besides any
// starting with "_", which are always excluded) — e.g. a tab with
// cumulative formulas, not per-entry data, that doesn't belong in the
// input flow (still emailed as-is if you reference it manually, just
// never shown as a "log an entry" choice). Editable from _Config's
// HiddenTabs row, or the in-app "Manage Fields & Options" screen — see
// getHiddenTabs() near ensureConfigTab().

// Tabs actually pulled into the compiled weekly report (email + download)
// — an explicit allowlist, deliberately NOT "whatever tabs happen to be
// visible in the widget right now". New tabs (Leave, or anything else you
// or the team add later) show up in the widget just fine without being
// added here — they just don't show up IN the report unless you
// explicitly opt them in. Editable from _Config's ReportTabs row, or the
// in-app "Manage Fields & Options" screen — see getReportTabs() near
// ensureConfigTab().

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
    /** Every Weekly Connect ticket — powers the widget's "View & Update
     * Tickets" screen. Deliberately NOT run through cached() like tabs/
     * options/etc above: tickets change constantly (new ones logged
     * through the week, Status/Comments edited every Wednesday), and
     * someone opening this screen needs the real current state, not up to
     * CACHE_SECONDS stale — the whole point of "pick a ticket to update"
     * breaks if it's working from an outdated list. */
    if (action === 'weeklyConnectTickets') {
      return jsonOut({ ok: true, tickets: getWeeklyConnectTickets() });
    }
    // Deliberately NOT a doGet action like tabs/options/fieldSchema/
    // categories above, even though it's a read — those return tab
    // structure or (at most) first names; this one returns real email
    // addresses, and doGet has no token check at all (see its comment).
    // getReportSettings lives in doPost below instead, gated by the same
    // SHARED_SECRET check every write already goes through.
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
      compileAndSendReport(parseRangeFromRequest(body), body.configName);
      return jsonOut({ ok: true, message: 'Report sent.' });
    }

    /** Same report as sendReportNow, but to exactly one address instead of
     * the whole config's recipient list — for previewing a design/content
     * change (like the HTML email) without emailing the entire team every
     * time. Not exposed anywhere in the widget UI on purpose; call it
     * directly (e.g. from the Apps Script editor's Run, or a one-off
     * request) when you actually want a preview. */
    if (body.action === 'sendTestReport') {
      if (!body.to) return jsonOut({ ok: false, error: 'Missing "to".' });
      var testConfig = getReportConfigByName_(body.configName);
      if (!testConfig) return jsonOut({ ok: false, error: 'No report configured yet.' });
      // Restricted to that config's own recipients, not an arbitrary
      // address — the shared TOKEN is baked into every installed widget
      // (and has leaked once before, which is why it exists as a Script
      // Property instead of hardcoded now), so letting `to` be
      // attacker-controlled would turn this into a way to mail the full
      // report — attachment included — to any outside address using
      // nothing but that token.
      if (testConfig.recipients.indexOf(body.to) === -1) {
        return jsonOut({ ok: false, error: '"to" must be one of "' + testConfig.name + '"\'s current recipients.' });
      }
      var testRange = parseRangeFromRequest(body) || getCurrentWeekRange();
      var testBuilt = buildReportBlob(testRange, testConfig);
      var testSheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
      MailApp.sendEmail({
        to: body.to,
        subject: '[TEST] ' + testConfig.name + ' — ' + formatRangeLabel(testRange),
        body: reportEmailPlainText(testRange, testBuilt, testSheetUrl),
        htmlBody: reportEmailHtml(testRange, testBuilt, testSheetUrl),
        attachments: [testBuilt.blob],
      });
      return jsonOut({ ok: true, message: 'Test report sent to ' + body.to + '.' });
    }

    if (body.action === 'downloadReport') {
      var file = getReportFileBase64(parseRangeFromRequest(body), body.configName);
      return jsonOut({ ok: true, fileName: file.fileName, base64: file.base64 });
    }

    /** Read/write for _ReportConfigs — the app/Sheet comparison table's
     * "In App" column for this needs a real backend action, same
     * reasoning as saveOptions/saveFieldSchema/saveCategories above. */
    if (body.action === 'getReportConfigs') {
      return jsonOut({ ok: true, configs: getReportConfigs() });
    }
    if (body.action === 'saveReportConfigs') {
      var configsLock = LockService.getScriptLock();
      if (!configsLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        writeReportConfigs(body.configs || []);
      } finally {
        configsLock.releaseLock();
      }
      return jsonOut({ ok: true, message: 'Report configs saved.' });
    }

    /** Lets the widget's Settings screen trigger the same clearCache()
     * that used to require opening the Apps Script editor by hand — e.g.
     * right after reordering tabs in the Sheet, so the new order shows up
     * immediately instead of waiting out CACHE_SECONDS. */
    if (body.action === 'clearCache') {
      clearCache();
      return jsonOut({ ok: true, message: 'Cache cleared.' });
    }

    /** Powers the widget's in-app "Manage Fields & Options" screen (local-
     * password-gated on the client side, same admin password as connecting
     * a new backend — see main.js). Each of these three actions replaces
     * that ENTIRE tab's data rows with whatever the client sends, rather
     * than patching individual rows — the client always sends the full,
     * current map back (it fetched the same shape via GET first), so a
     * full replace is simpler and can't drift out of sync from a partial
     * edit. Row 1 (headers) is never touched. */
    if (body.action === 'saveOptions' || body.action === 'saveFieldSchema' || body.action === 'saveCategories') {
      var writeLock = LockService.getScriptLock();
      if (!writeLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        if (body.action === 'saveOptions') {
          writeOptionsMap(body.options || {});
          CacheService.getScriptCache().remove('options');
          return jsonOut({ ok: true, message: 'Options saved.' });
        }
        if (body.action === 'saveFieldSchema') {
          writeFieldSchemaMap(body.fieldSchema || {});
          CacheService.getScriptCache().remove('fieldSchema');
          return jsonOut({ ok: true, message: 'Field types saved.' });
        }
        writeCategoriesMap(body.categories || {});
        CacheService.getScriptCache().remove('categories');
        return jsonOut({ ok: true, message: 'Categories saved.' });
      } finally {
        writeLock.releaseLock();
      }
    }

    /** Read-side of Report Settings (ReportTabs/HiddenTabs/
     * ReportRecipients/ReminderRecipients — see ensureConfigDefaults_()).
     * A POST, not a doGet action like tabs/options/fieldSchema/categories
     * — see the comment on this action's absence from doGet above; email
     * addresses are more sensitive than doGet's unauthenticated reads
     * elsewhere are comfortable exposing. Returns the RAW _Config values,
     * not the resolved (with-Script-Properties-fallback) ones getReportRecipients()/
     * getReminderRecipients() use when actually sending mail — so the
     * editor UI shows a field genuinely blank when nothing's set there
     * yet, rather than looking pre-filled with a value that isn't really
     * stored in the sheet. */
    if (body.action === 'getReportSettings') {
      // Every real tab except internal "_"-prefixed ones — deliberately
      // NOT listVisibleTabs()/getTabs(), which also excludes anything
      // already in HiddenTabs. This screen needs to see (and let you
      // un-hide) a currently-hidden tab, which the widget's normal tab
      // list can never show you in the first place.
      var allTabs = SpreadsheetApp.getActiveSpreadsheet()
        .getSheets()
        .map(function (s) { return s.getName(); })
        .filter(function (name) { return name.indexOf('_') !== 0; });
      return jsonOut({
        ok: true,
        allTabs: allTabs,
        reportSettings: {
          reportTabs: getReportTabs(),
          hiddenTabs: getHiddenTabs(),
          reportRecipients: getConfigList_('ReportRecipients'),
          reminderRecipients: getConfigList_('ReminderRecipients'),
          teamsWebhookUrl: getConfigValue('TeamsWebhookUrl'),
        },
      });
    }

    /** Write-side — same "replace the whole thing" shape as saveOptions/
     * saveFieldSchema/saveCategories above. HiddenTabs affects the tab
     * list, so its cache key is cleared too (the others don't feed any
     * cached('...') read here, but clearing 'tabs' is cheap and correct
     * either way). */
    if (body.action === 'saveReportSettings') {
      var reportLock = LockService.getScriptLock();
      if (!reportLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        var settings = body.settings || {};
        setConfigValue('ReportTabs', (settings.reportTabs || []).join(', '));
        setConfigValue('HiddenTabs', (settings.hiddenTabs || []).join(', '));
        setConfigValue('ReportRecipients', (settings.reportRecipients || []).join(', '));
        setConfigValue('ReminderRecipients', (settings.reminderRecipients || []).join(', '));
        setConfigValue('TeamsWebhookUrl', (settings.teamsWebhookUrl || '').trim());
      } finally {
        reportLock.releaseLock();
      }
      CacheService.getScriptCache().remove('tabs');
      return jsonOut({ ok: true, message: 'Report settings saved.' });
    }

    /** Updates one Weekly Connect ticket's Status/Comments in place — the
     * one write in this whole app that edits an existing row instead of
     * appending. Lock-guarded like every other write, so two people
     * closing out the same ticket at once can't race each other. */
    if (body.action === 'updateWeeklyConnectTicket') {
      if (!body.ticketId) {
        return jsonOut({ ok: false, error: 'Missing "ticketId".' });
      }
      var ticketLock = LockService.getScriptLock();
      if (!ticketLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      var found;
      try {
        found = updateWeeklyConnectTicket(body.ticketId, body.status || '', body.comments || '');
      } finally {
        ticketLock.releaseLock();
      }
      if (!found) {
        return jsonOut({ ok: false, error: 'Ticket "' + body.ticketId + '" not found.' });
      }
      logAudit_(body.sessionToken, 'update', WEEKLY_CONNECT_TAB_NAME,
        'Ticket #' + body.ticketId + ' -> Status: ' + (body.status || '(unchanged)'));
      return jsonOut({ ok: true, message: 'Ticket updated.' });
    }

    /** The manual "Post to Teams" action — replaces the old auto-post on
     * every save (see the comment where that used to be called, right
     * after appendRow above). Runs synchronously and returns the real
     * outcome: unlike a ticket save, nothing else is waiting on this, and
     * whoever clicked the button should see whether it actually worked,
     * not just trust it happened. start/end are required (not defaulted
     * to "this week") — the whole point is picking the range that matches
     * whenever the real discussion actually happened. */
    if (body.action === 'postWeeklyConnectToTeamsNow') {
      var postRange = parseRangeFromRequest(body);
      if (!postRange) {
        return jsonOut({ ok: false, error: 'Pick a date range first.' });
      }
      var postResult = postWeeklyConnectToTeams(body.group || '', postRange);
      return jsonOut({ ok: postResult.ok, message: postResult.message, error: postResult.ok ? undefined : postResult.message });
    }

    /** Read/write for _ConnectGroups — a webhook URL is a write capability,
     * same reasoning as getReportSettings above for going through doPost
     * (token-gated) instead of doGet. */
    if (body.action === 'getConnectGroups') {
      return jsonOut({ ok: true, groups: getConnectGroups() });
    }
    if (body.action === 'saveConnectGroups') {
      var groupsLock = LockService.getScriptLock();
      if (!groupsLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        writeConnectGroups(body.groups || []);
      } finally {
        groupsLock.releaseLock();
      }
      CacheService.getScriptCache().remove('options'); // ConnectGroups synced into _Options
      return jsonOut({ ok: true, message: 'Connect groups saved.' });
    }

    /** Per-person accounts — see the _Users/_Sessions section above for
     * why these exist and why they're not enforced on any other action
     * yet. bootstrapFirstAdmin only ever succeeds once, the very first
     * time it's called against a backend with no accounts at all; every
     * account after that comes from an existing admin's createUser. */
    if (body.action === 'bootstrapFirstAdmin') {
      return jsonOut(bootstrapFirstAdmin_(body.username, body.password));
    }
    if (body.action === 'login') {
      return jsonOut(loginUser_(body.username, body.password));
    }
    /** Lets the widget ask "does an account exist for me yet, and is
     * logging in mandatory right now?" using the name already saved
     * locally (YOUR_NAME), without needing a password to check.
     * `exists` is what makes rollout to existing installs work: nothing
     * changes for anyone until an admin runs createUser for their name,
     * and the very next time they open the app after that, it finds
     * `exists: true` here and shows the login screen instead of silently
     * requiring a reinstall or some other manual step. `requireLogin`
     * (from _Config's RequireLogin, default FALSE) is the separate,
     * later switch: once an admin flips that on — after everyone already
     * has an account, not before — EVERY install gets gated regardless of
     * whether its saved name happens to match an account, not just ones
     * that do. No account info beyond these two booleans leaves the
     * server; a blank/no username still gets a real requireLogin value
     * back, just with exists always false. */
    if (body.action === 'hasAccount') {
      var hasData = usersSheetRows_();
      var hasIdx = body.username ? findUserRowIndex_(hasData.headers, hasData.rows, body.username) : -1;
      return jsonOut({ ok: true, exists: hasIdx !== -1, requireLogin: getConfigValue('RequireLogin') === 'TRUE' });
    }
    if (body.action === 'listUsers') {
      var listAuth = requireAdminSession_(body);
      if (listAuth.error) return jsonOut({ ok: false, error: listAuth.error });
      return jsonOut({ ok: true, users: listUsersPublic_() });
    }
    if (body.action === 'createUser') {
      var createAuth = requireAdminSession_(body);
      if (createAuth.error) return jsonOut({ ok: false, error: createAuth.error });
      if (!body.username || !body.password) {
        return jsonOut({ ok: false, error: 'Username and password are required.' });
      }
      var createLock = LockService.getScriptLock();
      if (!createLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        var createData = usersSheetRows_();
        if (findUserRowIndex_(createData.headers, createData.rows, body.username) !== -1) {
          return jsonOut({ ok: false, error: 'That username is already taken.' });
        }
        var newSalt = Utilities.getUuid();
        createData.sheet.appendRow([
          String(body.username).trim(),
          hashPassword_(body.password, newSalt),
          newSalt,
          body.role === 'Admin' ? 'Admin' : 'User',
          'TRUE',
          new Date(),
          '',
        ]);
      } finally {
        createLock.releaseLock();
      }
      return jsonOut({ ok: true, message: 'Account created.' });
    }
    if (body.action === 'resetUserPassword') {
      var resetAuth = requireAdminSession_(body);
      if (resetAuth.error) return jsonOut({ ok: false, error: resetAuth.error });
      if (!body.username || !body.newPassword) {
        return jsonOut({ ok: false, error: 'Username and new password are required.' });
      }
      var resetLock = LockService.getScriptLock();
      if (!resetLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        var resetData = usersSheetRows_();
        var resetIdx = findUserRowIndex_(resetData.headers, resetData.rows, body.username);
        if (resetIdx === -1) return jsonOut({ ok: false, error: 'No such account: ' + body.username });
        var resetSalt = Utilities.getUuid();
        var hashCol = findColumnIndex(resetData.headers, 'PasswordHash');
        var saltCol = findColumnIndex(resetData.headers, 'Salt');
        resetData.sheet.getRange(resetIdx + 2, hashCol + 1).setValue(hashPassword_(body.newPassword, resetSalt));
        resetData.sheet.getRange(resetIdx + 2, saltCol + 1).setValue(resetSalt);
      } finally {
        resetLock.releaseLock();
      }
      return jsonOut({ ok: true, message: 'Password reset.' });
    }
    if (body.action === 'setUserActive') {
      var activeAuth = requireAdminSession_(body);
      if (activeAuth.error) return jsonOut({ ok: false, error: activeAuth.error });
      if (!body.username) return jsonOut({ ok: false, error: 'Missing "username".' });
      var activeLock = LockService.getScriptLock();
      if (!activeLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        var activeData = usersSheetRows_();
        var activeIdx = findUserRowIndex_(activeData.headers, activeData.rows, body.username);
        if (activeIdx === -1) return jsonOut({ ok: false, error: 'No such account: ' + body.username });
        var targetRec = rowToUserRecord_(activeData.headers, activeData.rows[activeIdx]);
        // Refuse to deactivate the last remaining active admin — the one
        // way this app could otherwise permanently lock everyone out of
        // user management, with no other admin left to undo it.
        if (!body.active && targetRec['Role'] === 'Admin' && countOtherActiveAdmins_(activeData, activeIdx) === 0) {
          return jsonOut({ ok: false, error: 'Can\'t deactivate the only remaining admin account.' });
        }
        var activeCol = findColumnIndex(activeData.headers, 'Active');
        activeData.sheet.getRange(activeIdx + 2, activeCol + 1).setValue(body.active ? 'TRUE' : 'FALSE');
      } finally {
        activeLock.releaseLock();
      }
      return jsonOut({ ok: true, message: body.active ? 'Account reactivated.' : 'Account deactivated.' });
    }
    if (body.action === 'setUserRole') {
      var roleAuth = requireAdminSession_(body);
      if (roleAuth.error) return jsonOut({ ok: false, error: roleAuth.error });
      if (!body.username || (body.role !== 'Admin' && body.role !== 'User')) {
        return jsonOut({ ok: false, error: 'Username and a valid role (Admin/User) are required.' });
      }
      var roleLock = LockService.getScriptLock();
      if (!roleLock.tryLock(LOCK_WAIT_MS)) {
        return jsonOut({ ok: false, error: 'Server is busy — please try again in a few seconds.' });
      }
      try {
        var roleData = usersSheetRows_();
        var roleIdx = findUserRowIndex_(roleData.headers, roleData.rows, body.username);
        if (roleIdx === -1) return jsonOut({ ok: false, error: 'No such account: ' + body.username });
        var roleTargetRec = rowToUserRecord_(roleData.headers, roleData.rows[roleIdx]);
        // Same guard as setUserActive, same reason — demoting the last
        // admin to User is just as much a lockout as deactivating them.
        if (body.role === 'User' && roleTargetRec['Role'] === 'Admin' && countOtherActiveAdmins_(roleData, roleIdx) === 0) {
          return jsonOut({ ok: false, error: 'Can\'t demote the only remaining admin account.' });
        }
        var roleCol = findColumnIndex(roleData.headers, 'Role');
        roleData.sheet.getRange(roleIdx + 2, roleCol + 1).setValue(body.role);
      } finally {
        roleLock.releaseLock();
      }
      return jsonOut({ ok: true, message: 'Role updated.' });
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
    // Used to auto-post to Teams right here on every single ticket saved,
    // compiling whatever fell in a fixed Wednesday-to-Wednesday window.
    // Dropped that: the real discussion this was compiling for doesn't
    // happen on a fixed day — it moves around week to week — so a rigid
    // auto-window kept posting the wrong set of tickets. Posting to Teams
    // is now its own explicit action (postWeeklyConnectToTeamsNow below),
    // triggered from the widget with whatever date range actually matches
    // this week's real discussion, same custom-range picker Report
    // Generator already uses — see the ticket list screen's "Post to
    // Teams" button.

    // Best-effort, after the lock is already released — see logAudit_'s
    // own comment for why this never blocks or fails the actual save.
    // A short, generic summary (first few populated fields) rather than
    // anything per-tab special-cased, since this runs for every tab.
    var auditSummary = headers.slice(0, 3).map(function (h) {
      return values[h] ? h + ': ' + String(values[h]).slice(0, 40) : null;
    }).filter(Boolean).join(', ');
    logAudit_(body.sessionToken, 'submit', body.tab, auditSummary || 'New entry added.');

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
 * click Run) after pasting your real values into the lines below — saves
 * them into this project's Script Properties, which is where
 * SHARED_SECRET lives from then on (see the CONFIG section up top; it has
 * no sheet-editable equivalent — a token that lived in a sheet cell would
 * defeat its own purpose). REPORT_RECIPIENTS/REMINDER_RECIPIENTS here are
 * now OPTIONAL — recipients can be set directly on _Config's
 * ReportRecipients/ReminderRecipients instead (Sheet or the widget's
 * "Manage Fields & Options" screen — see getReportRecipients()), which is
 * the easier path going forward. These two Script Properties still work
 * as a fallback for a deployment that set them up before that existed.
 * Safe to re-run any time you want to rotate the token — this overwrites,
 * it doesn't append. You can also skip this function entirely and edit
 * the same properties directly under Project Settings > Script
 * Properties, if you'd rather not have the values pasted into a function
 * body at all, even briefly and un-committed.
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
    // TEAMS_WEBHOOK_URL: 'https://....webhook.office.com/...' — optional, only if you want Weekly Connect / the Friday reminder posted to Teams.
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
 * Run this once manually (Run > setupTriggers) to install the Friday
 * reminder trigger (and, as of the multi-config Report Configs system,
 * the shared hourly report-schedule dispatcher too). Safe to re-run: it
 * clears old triggers from this project first so you never end up with
 * duplicates.
 *
 * No longer creates a fixed Friday-11pm report trigger here —
 * scheduledCompileAndSendReport() (still defined below, just no longer
 * wired to a new trigger by this function) has been superseded by
 * runScheduledReports_(), which checks every _ReportConfigs entry's own
 * ScheduleDay/ScheduleHour instead of one hardcoded time. That dispatcher
 * gets created below via ensureScheduledReportsTrigger_() — the same
 * self-healing call that also fires automatically the moment
 * _ReportConfigs is first created (see ensureReportConfigsTab), so this
 * function existing/being re-run isn't actually required for scheduling
 * to work at all; it's here mainly for the Friday reminder, and as a
 * manual "make sure both exist" fallback.
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

  ensureScheduledReportsTrigger_();

  Logger.log('Triggers installed: Friday 6pm reminder, hourly report-schedule dispatcher.');
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
/** "Week N (Month)" — N is 1-indexed from that date's day-of-month / 7,
 * so it resets to Week 1 on its own every month (day 1-7 = Week 1, day
 * 8-14 = Week 2, ...) with no stored counter to maintain or roll over.
 * Deliberately NOT the same thing as Weekly Connect's Ticket ID, which
 * stays a single global-forever counter per the earlier decision on
 * that — this is purely a display label for the Report Generator's
 * reminder/report, on a date within the work week it's labeling (the
 * Friday it's sent on, or a custom range's end date). */
function weekLabelForDate_(date) {
  var tz = Session.getScriptTimeZone();
  var weekNum = Math.ceil(Number(Utilities.formatDate(date, tz, 'd')) / 7);
  return 'Week ' + weekNum + ' (' + Utilities.formatDate(date, tz, 'MMMM') + ')';
}

function sendFridayReminder() {
  if (isHoliday(new Date())) {
    Logger.log('Skipping Friday reminder — today is listed in ' + HOLIDAYS_TAB_NAME + '.');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subject = 'Reminder: submit your status update for ' + weekLabelForDate_(new Date()) + ' (Mon-Fri) by 10pm tonight';
  var body =
    'Hi team,\n\n' +
    'Friendly reminder to log your status update for the week before 10pm tonight.\n' +
    'Use the desktop widget, or update the sheet directly here:\n' +
    ss.getUrl() + '\n\n' +
    'The completed sheet goes out automatically at 11pm.\n';

  getReminderRecipients().forEach(function (addr) {
    MailApp.sendEmail(addr, subject, body);
  });

  var teamsWebhookUrl = getTeamsWebhookUrl();
  if (teamsWebhookUrl) {
    // buildAdaptiveCardPayload_(), not a plain {text: ...} — see its own
    // comment (near PENDING_TEAMS_GROUPS_PROP) for why: the same shape
    // Weekly Connect's Teams post uses, confirmed against a real failure
    // from a webhook backed by Teams' current Workflows/Power Automate
    // route rather than the older Connectors/Incoming Webhook one.
    UrlFetchApp.fetch(teamsWebhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(buildAdaptiveCardPayload_('⏰ ' + subject + ' — ' + ss.getUrl())),
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
/** `config` is a getReportConfigs()-shaped object — {name, tabs: [{tab,
 * fields}], ...}. Date-range filtering always uses the tab's FULL header
 * row (so filtering still works correctly even if Date itself isn't one
 * of the selected output columns); column selection is applied
 * separately, after filtering, to decide what actually gets written to
 * the output sheet. An empty `fields` list means every column, same as
 * today's report always included everything — this is what makes the
 * migrated default config behave identically to before. */
function buildReportBlob(range, config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempName = ss.getName() + ' — ' + config.name + ' — ' + formatRangeLabel(range);

  var temp = SpreadsheetApp.create(tempName);
  // config.tabs, not listVisibleTabsUncached() — the report is scoped to
  // this config's explicit list of tabs, not "whatever's visible in the
  // widget right now" (filtered defensively in case one's been
  // renamed/removed since this config was set up).
  var tabEntries = config.tabs.filter(function (t) { return !!getSheetByName(t.tab); });
  var counts = []; // [{tab, count}, ...] in the same order as tabEntries — used by the email body

  tabEntries.forEach(function (entry, i) {
    var tabName = entry.tab;
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

    // Column selection, applied AFTER date filtering above (which needs
    // the tab's real header indexes) — a non-empty Fields list picks out
    // just those columns' indexes, in the order the config listed them,
    // not the sheet's own column order, so an admin can reorder what
    // shows up in the report just by reordering Fields.
    var outHeaders = headers;
    var outRows = rows;
    if (entry.fields && entry.fields.length) {
      var colIndexes = entry.fields
        .map(function (f) { return findColumnIndex(headers, f); })
        .filter(function (idx) { return idx !== -1; });
      if (colIndexes.length) {
        outHeaders = colIndexes.map(function (idx) { return headers[idx]; });
        outRows = rows.map(function (r) { return colIndexes.map(function (idx) { return r[idx]; }); });
      }
    }

    var dest = i === 0 ? temp.getSheets()[0].setName(tabName) : temp.insertSheet(tabName);
    dest.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);
    if (outRows.length) {
      dest.getRange(2, 1, outRows.length, outHeaders.length).setValues(outRows);
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

// Only meaningful for the default Mon-Fri work week the Friday trigger
// sends — a custom range picked from the app (a whole month, a quarter,
// a year) isn't "a week" at all, so this only prefixes the subject when
// the range actually looks like one (5 days or fewer, matching Mon-Fri).
function isWorkWeekRange_(range) {
  var days = Math.round((range.end - range.start) / (24 * 60 * 60 * 1000));
  return days >= 0 && days <= 6;
}

/**
 * Emails the report for `range` (defaults to the current Mon-Fri work week)
 * under `configName`'s setup — its tabs, its selected fields, its
 * recipients. `configName` blank/not found falls back to the first
 * config (see getReportConfigByName_) — what makes this work exactly as
 * before for anyone not sending an explicit one yet, since almost every
 * deployment only has the one migrated "Report Generator" config.
 * Silently does nothing if no config exists at all (a brand-new org that
 * hasn't set one up yet) rather than erroring on a scheduled trigger
 * with no one watching.
 */
function compileAndSendReport(range, configName) {
  var config = getReportConfigByName_(configName);
  if (!config) return;
  range = range || getCurrentWeekRange();
  var built = buildReportBlob(range, config);

  var subject = config.name + ' — ' +
    (isWorkWeekRange_(range) ? weekLabelForDate_(range.end) + ' — ' : '') +
    formatRangeLabel(range);
  var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  config.recipients.forEach(function (addr) {
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
function getReportFileBase64(range, configName) {
  // Was missing the same `range || getCurrentWeekRange()` default
  // compileAndSendReport has right above — meant every "Download Report"
  // with the default "This Week" selection (range comes through as null)
  // crashed with "Cannot read properties of null (reading 'start')"
  // instead of actually downloading anything. Found via a live test
  // during a full test sweep.
  range = range || getCurrentWeekRange();
  var config = getReportConfigByName_(configName);
  if (!config) throw new Error('No report configured yet — set one up in Manage Fields & Options first.');
  var built = buildReportBlob(range, config);
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
  // Must come AFTER the three ensure*Tab() calls above — it merges into
  // whatever they just seeded (or already had) via get/write*Map(), and
  // running first would let its narrower write clobber their migration.
  // ensureConnectGroupsTab() must come before ensureWeeklyConnectTab(), so
  // the Group field's dropdown already has a real "ConnectGroups" _Options
  // entry to point at by the time Weekly_Connect's own schema is seeded.
  ensureConnectGroupsTab();
  ensureWeeklyConnectTab();
  ensureFeaturesTab();
  var hiddenTabs = getHiddenTabs();
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) {
      return name.indexOf('_') !== 0 && hiddenTabs.indexOf(name) === -1;
    });
}

// Name of the auto-created "Leave" tab — a normal, visible tab (unlike
// _Config/_Holidays) so it shows up in the widget's tab list like any
// other "log an entry" category, letting people log their own leave/
// holiday/WFH days without you having to set the tab up by hand first.
// Deliberately not in _Config's ReportTabs by default: it's the same
// shape of data as everything else, it just isn't part of the weekly work
// report unless you explicitly add it there.
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

// ------------------------- Weekly Connect ----------------------------

// Name of the auto-created "Weekly Connect" tab — a normal, visible tab,
// same self-creating pattern as ensureLeaveTab(). Tracks questions/issues
// raised through the week and settled every Wednesday's CMS Weekly
// Connect meeting. Unlike Daily Status/Leave/etc, this tab never existed
// as hardcoded config in an older version of this app, so there's nothing
// to "migrate" — ensureWeeklyConnectTab() seeds its own _FieldSchema/
// _Options/_Categories rows itself, the first time it creates the tab.
var WEEKLY_CONNECT_TAB_NAME = 'Weekly_Connect';
var WEEKLY_CONNECT_COLUMNS = [
  'Ticket ID', 'Group', 'Date', 'Requester', 'Owner', 'Site', 'Type',
  'Issue', 'URL', 'Attachment', 'Priority', 'Status', 'Comments',
];

/** Created once, the first time anything asks for the tab list, if it
 * doesn't already exist. Never touches the tab again once it's there —
 * same as ensureLeaveTab() — so renaming columns or adding your own is
 * completely safe (just keep 'Ticket ID', 'Group', and 'Status' if you
 * want updateWeeklyConnectTicket()/the Teams post to keep working). */
function ensureWeeklyConnectTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(WEEKLY_CONNECT_TAB_NAME)) return;
  var sheet = ss.insertSheet(WEEKLY_CONNECT_TAB_NAME);
  sheet.getRange(1, 1, 1, WEEKLY_CONNECT_COLUMNS.length).setValues([WEEKLY_CONNECT_COLUMNS]);
  seedWeeklyConnectConfig_();
}

/** One-time seed of this tab's own field types/options/category, merged
 * into whatever's already on _FieldSchema/_Options/_Categories (reuses
 * the exact same write*Map() functions the in-app "Manage Fields &
 * Options" screen uses, so this is just those same edits made
 * automatically instead of by hand). Requester/Owner/Site reuse option
 * lists that already exist elsewhere in the app (REQUESTERS/ASSIGNEES/
 * SITES) rather than inventing new ones. Group's own choices come from
 * _ConnectGroups (see ensureConnectGroupsTab()), synced into _Options'
 * ConnectGroups key — this just points the Group field at that key, same
 * as every other select field. */
function seedWeeklyConnectConfig_() {
  var fsMap = getFieldSchemaMap();
  fsMap[WEEKLY_CONNECT_TAB_NAME] = {
    'Ticket ID': { type: 'sequence' },
    'Group': { type: 'select', optionsKey: 'ConnectGroups' },
    'Date': { type: 'date' },
    'Requester': { type: 'select', optionsKey: 'REQUESTERS' },
    'Owner': { type: 'select', optionsKey: 'ASSIGNEES' },
    'Site': { type: 'select', optionsKey: 'SITES' },
    'Type': { type: 'select', optionsKey: 'Weekly_Connect.Type' },
    'Priority': { type: 'select', optionsKey: 'Weekly_Connect.Priority' },
    'Status': { type: 'select', optionsKey: 'Weekly_Connect.Status' },
  };
  writeFieldSchemaMap(fsMap);

  var optMap = getOptionsMap();
  optMap['Weekly_Connect.Type'] = ['Bug', 'Question', 'Enhancement', 'Blocker'];
  optMap['Weekly_Connect.Priority'] = ['Low', 'Medium', 'High', 'Urgent'];
  optMap['Weekly_Connect.Status'] = ['Open', 'In Progress', 'Resolved', 'Deferred'];
  writeOptionsMap(optMap);

  var catMap = getCategoriesMap();
  var existing = catMap['Weekly Connect'] || [];
  if (existing.indexOf(WEEKLY_CONNECT_TAB_NAME) === -1) {
    catMap['Weekly Connect'] = existing.concat([WEEKLY_CONNECT_TAB_NAME]);
    writeCategoriesMap(catMap);
  }
}

// One Weekly Connect can serve more than one recurring meeting/channel —
// "CMS Weekly Connect" today, something else you add later, each posting
// to its own Teams channel. _ConnectGroups (Group Name | Teams Webhook
// URL) is the list of those groups: every ticket picks one via the Group
// field (seeded above), and postWeeklyConnectToTeams() posts each
// ticket's current week to THAT group's webhook, not a single shared one.
// Fully Sheet/App-editable — add a row here (or use the in-app "Manage
// Fields & Options" -> Connect Groups screen) any time you need another
// group; no code change, no redeploy.
var CONNECT_GROUPS_TAB_NAME = '_ConnectGroups';

/** Created once. Seeds a single starting group, "CMS Weekly Connect" —
 * this app's own origin story, not a hardcoded requirement — carrying
 * forward whatever _Config's TeamsWebhookUrl already held (from the
 * single-webhook mode this replaces) as that group's starting webhook, so
 * nothing already configured gets silently dropped. An org that renames
 * or deletes that row, or adds more, is fully in control from here on. */
function ensureConnectGroupsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(CONNECT_GROUPS_TAB_NAME)) return;
  var sheet = ss.insertSheet(CONNECT_GROUPS_TAB_NAME);
  sheet.getRange(1, 1, 1, 2).setValues([['Group Name', 'Teams Webhook URL']]);
  var carriedOverWebhook = getConfigValue('TeamsWebhookUrl') || getScriptProp_('TEAMS_WEBHOOK_URL');
  sheet.getRange(2, 1, 1, 2).setValues([['CMS Weekly Connect', carriedOverWebhook]]);
  syncConnectGroupOptions_([{ name: 'CMS Weekly Connect', webhookUrl: carriedOverWebhook }]);
}

/** [{name, webhookUrl}, ...] in sheet row order — order matters here,
 * since it's also the order the Group dropdown offers choices in. */
function getConnectGroups() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONNECT_GROUPS_TAB_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var groups = [];
  values.forEach(function (row) {
    var name = String(row[0]).trim();
    if (!name) return;
    groups.push({ name: name, webhookUrl: String(row[1]).trim() });
  });
  return groups;
}

/** Replaces _ConnectGroups' data rows wholesale (same "replace the whole
 * thing" shape as writeOptionsMap() etc) and keeps _Options' ConnectGroups
 * key in sync with the new group names, so the Group field's dropdown
 * (an ordinary select, see seedWeeklyConnectConfig_()) never needs its
 * own bespoke read path — it just reads _Options like every other select
 * field already does. */
function writeConnectGroups(groups) {
  ensureConnectGroupsTab();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONNECT_GROUPS_TAB_NAME);
  var rows = (groups || [])
    .map(function (g) { return [String(g.name || '').trim(), String(g.webhookUrl || '').trim()]; })
    .filter(function (r) { return r[0]; });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  syncConnectGroupOptions_(rows.map(function (r) { return { name: r[0] }; }));
}

function syncConnectGroupOptions_(groups) {
  var optMap = getOptionsMap();
  optMap['ConnectGroups'] = groups.map(function (g) { return g.name; });
  writeOptionsMap(optMap);
}

/** The webhook URL for one named group, or '' if that group isn't
 * configured (a blank Group on the ticket, a group whose webhook was
 * never set, or a group that's since been removed from _ConnectGroups —
 * postWeeklyConnectToTeams() treats all of these as "nothing to post",
 * same silent no-op as Teams being off entirely). */
function getConnectGroupWebhookUrl(groupName) {
  var match = getConnectGroups().filter(function (g) { return g.name === groupName; })[0];
  return match ? match.webhookUrl : '';
}

// -------------------- per-person accounts (_Users / _Sessions) --------------------
//
// Sprint 1 of "restrict who can use this": right now every write only
// checks the one shared SHARED_SECRET baked into every installer — there's
// no concept server-side of WHO is submitting, and the client-side "admin
// password" gating Manage Fields & Options (see saveOptions/saveFieldSchema/
// saveCategories above) is checked ONLY on the widget, never re-checked
// here — a request with a valid SHARED_SECRET can call those actions
// directly with no admin password at all.
//
// This section adds real accounts, stored the same way as everything else
// in this project: a Sheet tab, admin-editable, self-creating. It's
// deliberately NOT wired into saveOptions/saveFieldSchema/saveCategories/
// saveConnectGroups or the regular submit-a-report actions yet — those
// still work exactly as before. Flipping them to require a valid session
// is a follow-up step, once the widget itself has a login screen to get
// that session from; adding the requirement here first would just lock
// everyone out with no way back in.
var USERS_TAB_NAME = '_Users';
var USERS_COLUMNS = ['Username', 'PasswordHash', 'Salt', 'Role', 'Active', 'CreatedAt', 'LastLoginAt'];
var SESSIONS_TAB_NAME = '_Sessions';
var SESSIONS_COLUMNS = ['SessionToken', 'Username', 'CreatedAt'];

function ensureUsersTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(USERS_TAB_NAME)) return;
  var sheet = ss.insertSheet(USERS_TAB_NAME);
  sheet.getRange(1, 1, 1, USERS_COLUMNS.length).setValues([USERS_COLUMNS]);
  // No seed row — there's no credential to invent. The very first account
  // is created by bootstrapFirstAdmin below, which only works while this
  // tab has zero data rows, then refuses forever after.
}

function ensureSessionsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SESSIONS_TAB_NAME)) return;
  var sheet = ss.insertSheet(SESSIONS_TAB_NAME);
  sheet.getRange(1, 1, 1, SESSIONS_COLUMNS.length).setValues([SESSIONS_COLUMNS]);
}

// SHA-256(salt + password), base64 — Apps Script has no bcrypt/scrypt, this
// is what Utilities.computeDigest offers. Salted and per-user, which is
// what actually matters against the realistic threat here (an internal
// tool's Sheet backend, not a bank); a from-scratch bcrypt port would add
// real complexity for a threat model this doesn't face.
function hashPassword_(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function usersSheetRows_() {
  ensureUsersTab();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_TAB_NAME);
  var headers = getHeaderRow(sheet);
  var lastRow = sheet.getLastRow();
  var rows = lastRow <= 1 ? [] : sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return { sheet: sheet, headers: headers, rows: rows };
}

// Username lookup is case-insensitive (so "Naveen" and "naveen" are the
// same account — the honest expectation for a login field) but the
// original casing someone registered with is what's stored and shown.
function findUserRowIndex_(headers, rows, username) {
  var col = findColumnIndex(headers, 'Username');
  var needle = String(username || '').trim().toLowerCase();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][col]).trim().toLowerCase() === needle) return i;
  }
  return -1;
}

function rowToUserRecord_(headers, row) {
  var rec = {};
  headers.forEach(function (h, i) { rec[h] = row[i]; });
  return rec;
}

function isActiveValue_(v) {
  return String(v).trim().toUpperCase() === 'TRUE';
}

/** Only succeeds while _Users is completely empty — creates the very
 * first account (always Admin) and logs it straight in, so whoever sets
 * up a fresh backend doesn't need a separate "now go log in" step. Once
 * one account exists, this refuses forever — the normal path from there
 * is an existing admin using createUser. */
function bootstrapFirstAdmin_(username, password) {
  if (!username || !password) return { ok: false, error: 'Username and password are required.' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) {
    return { ok: false, error: 'Server is busy — please try again in a few seconds.' };
  }
  try {
    var data = usersSheetRows_();
    if (data.rows.length > 0) {
      return { ok: false, error: 'An admin account already exists — use the normal login instead.' };
    }
    var salt = Utilities.getUuid();
    data.sheet.appendRow([
      String(username).trim(),
      hashPassword_(password, salt),
      salt,
      'Admin',
      'TRUE',
      new Date(),
      '',
    ]);
  } finally {
    lock.releaseLock();
  }
  return loginUser_(username, password);
}

function loginUser_(username, password) {
  var data = usersSheetRows_();
  var idx = findUserRowIndex_(data.headers, data.rows, username);
  if (idx === -1) return { ok: false, error: 'Invalid username or password.' };
  var rec = rowToUserRecord_(data.headers, data.rows[idx]);
  if (!isActiveValue_(rec['Active'])) {
    return { ok: false, error: 'This account has been deactivated — contact your admin.' };
  }
  if (hashPassword_(password, rec['Salt']) !== rec['PasswordHash']) {
    return { ok: false, error: 'Invalid username or password.' };
  }
  var token = Utilities.getUuid();
  ensureSessionsTab();
  var sessSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_TAB_NAME);
  sessSheet.appendRow([token, rec['Username'], new Date()]);
  var lastLoginCol = findColumnIndex(data.headers, 'LastLoginAt');
  if (lastLoginCol !== -1) data.sheet.getRange(idx + 2, lastLoginCol + 1).setValue(new Date());
  return { ok: true, sessionToken: token, username: rec['Username'], role: rec['Role'] };
}

/** {username, role} for a valid, still-active session, or null — null
 * covers "no such token", "token belongs to a deactivated account", and
 * "no token sent" all the same way (the caller doesn't need to tell them
 * apart, just refuse). Not yet called from the general write path (see
 * the section comment above) — only from the admin-only user-management
 * actions below, until the widget itself sends a session on every call. */
function validateSession_(token) {
  if (!token) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sessSheet = ss.getSheetByName(SESSIONS_TAB_NAME);
  if (!sessSheet) return null;
  var lastRow = sessSheet.getLastRow();
  if (lastRow <= 1) return null;
  var sessRows = sessSheet.getRange(2, 1, lastRow - 1, SESSIONS_COLUMNS.length).getValues();
  var username = null;
  for (var i = 0; i < sessRows.length; i++) {
    if (sessRows[i][0] === token) { username = sessRows[i][1]; break; }
  }
  if (!username) return null;
  var data = usersSheetRows_();
  var idx = findUserRowIndex_(data.headers, data.rows, username);
  if (idx === -1) return null;
  var rec = rowToUserRecord_(data.headers, data.rows[idx]);
  if (!isActiveValue_(rec['Active'])) return null;
  return { username: rec['Username'], role: rec['Role'] };
}

/** Every user-management action below needs this — returns the validated
 * admin session on success, or a ready-to-return jsonOut() error object
 * (never both) so callers can just `var s = requireAdminSession_(body); if
 * (s.error) return jsonOut(s);`. */
function requireAdminSession_(body) {
  var session = validateSession_(body.sessionToken);
  if (!session) return { error: 'Please log in again.' };
  if (session.role !== 'Admin') return { error: 'Admin access required.' };
  return { session: session };
}

/** How many OTHER active admins exist besides the row at `excludeIdx` —
 * shared by setUserActive's deactivate guard and setUserRole's demote
 * guard, same reasoning both times: either one is a way to end up with
 * zero admins and no one left who can undo it. */
function countOtherActiveAdmins_(data, excludeIdx) {
  return data.rows.filter(function (row, i) {
    if (i === excludeIdx) return false;
    var r = rowToUserRecord_(data.headers, row);
    return r['Role'] === 'Admin' && isActiveValue_(r['Active']);
  }).length;
}

/** Every account, minus PasswordHash/Salt — those never leave the server,
 * same rule getReportSettings/getConnectionConfig etc already follow for
 * anything sensitive. */
function listUsersPublic_() {
  var data = usersSheetRows_();
  return data.rows.map(function (row) {
    var rec = rowToUserRecord_(data.headers, row);
    return {
      username: rec['Username'],
      role: rec['Role'],
      active: isActiveValue_(rec['Active']),
      createdAt: rec['CreatedAt'],
      lastLoginAt: rec['LastLoginAt'],
    };
  });
}

// -------------------- audit log (who did what) --------------------
//
// _Sessions already answers "who logged in, and when" — this answers the
// separate question "who actually submitted/edited THIS". Deliberately
// its own tab, not a new column bolted onto every existing tab (Daily
// Status, Weekly_Connect, Leave, ...): those already have a fixed column
// set other things (report generation, _FieldSchema rows) expect, and
// this needs to work the same way across all of them without touching
// any of that. And deliberately NO formulas here, ever — the whole
// reason this app had a real "stuck on Saving" problem for a long time
// was formulas with wide references recalculating on every single edit
// anywhere in the workbook. A plain, formula-free log tab that nothing
// else references can't trigger that, no matter how many rows pile up.
var AUDIT_LOG_TAB_NAME = '_AuditLog';
var AUDIT_LOG_COLUMNS = ['Timestamp', 'Username', 'Action', 'Tab', 'Summary'];

function ensureAuditLogTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(AUDIT_LOG_TAB_NAME)) return;
  var sheet = ss.insertSheet(AUDIT_LOG_TAB_NAME);
  sheet.getRange(1, 1, 1, AUDIT_LOG_COLUMNS.length).setValues([AUDIT_LOG_COLUMNS]);
}

/** Resolves a sessionToken to a username for logging, without ever
 * throwing or blocking the real action over it — logging in isn't
 * required for regular submissions yet (see RequireLogin above), so a
 * missing/invalid token here just means "not logged in", not an error.
 * Never call this inside the write lock other actions already hold for
 * their own appendRow — it's a plain read (validateSession_), safe
 * either way, but keeping it outside keeps the lock's own critical
 * section exactly what it was before this existed. */
function resolveAuditUsername_(sessionToken) {
  try {
    var session = validateSession_(sessionToken);
    return session ? session.username : '(not logged in)';
  } catch (err) {
    return '(not logged in)';
  }
}

/** Best-effort, always — a failure to log an action should never turn a
 * successful save into an error, or block it in any way. Called AFTER
 * the real write's lock is already released, same reasoning as
 * everything else in this app that does something non-essential after
 * the actual data is safely saved (e.g. the old Teams-posting design). */
function logAudit_(sessionToken, action, tab, summary) {
  try {
    ensureAuditLogTab();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AUDIT_LOG_TAB_NAME);
    sheet.appendRow([new Date(), resolveAuditUsername_(sessionToken), action, tab || '', summary || '']);
  } catch (err) {
    Logger.log('logAudit_ failed: ' + err);
  }
}

/** Every Weekly Connect ticket as a plain object per row, newest first —
 * powers the widget's "View & Update Tickets" list (the same list also
 * serves as the "open items" agenda view and the picker for
 * updateWeeklyConnectTicket()). `stale` is true for anything still Open
 * after STALE_TICKET_DAYS — a nudge, not a hard rule; nothing else changes
 * about a stale ticket. */
var STALE_TICKET_DAYS = 21;

function getWeeklyConnectTickets() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEEKLY_CONNECT_TAB_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var headers = getHeaderRow(sheet);
  var idIndex = findColumnIndex(headers, 'Ticket ID');
  var dateIndex = findColumnIndex(headers, 'Date');
  var statusIndex = findColumnIndex(headers, 'Status');
  var tz = Session.getScriptTimeZone();
  var staleBefore = new Date(Date.now() - STALE_TICKET_DAYS * 24 * 60 * 60 * 1000);

  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var tickets = values.map(function (row) {
    var ticket = {};
    headers.forEach(function (h, i) {
      var v = row[i];
      ticket[h] = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v;
    });
    var status = statusIndex === -1 ? '' : String(row[statusIndex]).trim();
    var rawDate = dateIndex === -1 ? null : row[dateIndex];
    var d = rawDate instanceof Date ? rawDate : new Date(rawDate);
    ticket.stale = status === 'Open' && !isNaN(d.getTime()) && d < staleBefore;
    return ticket;
  });

  if (idIndex !== -1) {
    tickets.sort(function (a, b) { return Number(b['Ticket ID']) - Number(a['Ticket ID']); });
  } else {
    tickets.reverse();
  }
  return tickets;
}

/** Finds the row for `ticketId` and overwrites its Status/Comments cells
 * in place — the one place in this whole app that edits an existing row
 * rather than appending. Only ever called while already holding the
 * write lock (see doPost), same as every other write. Returns false (no
 * throw) if the ticket ID isn't found, so the caller can return a clean
 * error instead of a stack trace. */
function updateWeeklyConnectTicket(ticketId, status, comments) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEEKLY_CONNECT_TAB_NAME);
  if (!sheet) return false;
  var headers = getHeaderRow(sheet);
  var idIndex = findColumnIndex(headers, 'Ticket ID');
  var statusIndex = findColumnIndex(headers, 'Status');
  var commentsIndex = findColumnIndex(headers, 'Comments');
  if (idIndex === -1) return false;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  var ids = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(ticketId).trim()) {
      var row = i + 2;
      if (statusIndex !== -1) sheet.getRange(row, statusIndex + 1).setValue(status);
      if (commentsIndex !== -1) sheet.getRange(row, commentsIndex + 1).setValue(comments);
      return true;
    }
  }
  return false;
}

/** The current "connect week" window: the 7 days ending on the most
 * recent Wednesday at/before today (so on a Wednesday itself, today is
 * included — the meeting day's own late-breaking tickets still show).
 * Used to scope the Teams post to just this week's tickets, not the
 * tab's entire history — see postWeeklyConnectToTeams(). */
function getCurrentConnectWeekRange() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var today = new Date(todayStr + 'T00:00:00');
  var WEDNESDAY = 3;
  var daysSinceWednesday = (today.getDay() - WEDNESDAY + 7) % 7;
  var weekEnd = new Date(today);
  weekEnd.setDate(today.getDate() - daysSinceWednesday);
  var weekStart = new Date(weekEnd);
  weekStart.setDate(weekEnd.getDate() - 6);
  return { start: weekStart, end: new Date(weekEnd.getFullYear(), weekEnd.getMonth(), weekEnd.getDate(), 23, 59, 59) };
}

/** Wraps `bodyOrText` into a minimal Adaptive Card — the shape Teams'
 * current "Post to a channel when a webhook request is received"
 * Workflow template (the modern replacement for the old Connectors/
 * Incoming Webhook route — see the README) actually expects, confirmed
 * against a real failure: sending a plain {text: ...} body got back
 * "Property 'type' must be 'AdaptiveCard'" from the flow's Post-card
 * action, which deserializes the request body directly as a card. If
 * your webhook is a genuinely classic Incoming Webhook instead (rare
 * now — Microsoft has been steering everyone toward Workflows), it
 * likely expects the older {text: ...} shape instead; this is the one
 * place to change that back if so.
 *
 * Accepts either a plain string (the Friday reminder's use — wrapped as
 * one TextBlock, same lightweight-Markdown shape richtext fields produce,
 * which TextBlock understands a limited version of natively) or a real
 * Adaptive Card body array (Weekly Connect's structured per-ticket cards
 * below — FactSet/Container/etc., not just a wall of markdown text). */
function buildAdaptiveCardPayload_(bodyOrText) {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.4',
    body: typeof bodyOrText === 'string'
      ? [{ type: 'TextBlock', text: bodyOrText, wrap: true }]
      : bodyOrText,
  };
}

/** Adaptive Cards' small set of semantic text colors — used to make
 * Priority readable at a glance without needing real CSS (which Adaptive
 * Cards don't support at all — no arbitrary styling, only whatever
 * building blocks the format itself defines). Anything not in this list
 * (a custom Priority value an org added) just falls back to 'default'
 * rather than guessing. */
var PRIORITY_COLOR_MAP_ = { Urgent: 'attention', High: 'attention', Medium: 'warning', Low: 'good' };

/** One ticket's card content — a heading, a color-coded priority line
 * if set, a FactSet for the other populated fields (aligned label:value
 * columns, not plain text lines), and the URL as a real markdown link if
 * present. Wrapped by the caller in a Container per ticket so each
 * question reads as its own visually separated block instead of a wall
 * of text run together. */
function buildTicketCardItems_(t) {
  var items = [
    { type: 'TextBlock', text: 'Q' + t['Ticket ID'] + ' — ' + (t['Issue'] || '(no description)'), wrap: true, weight: 'Bolder' },
  ];
  if (t['Priority']) {
    items.push({
      type: 'TextBlock',
      text: 'Priority: ' + t['Priority'],
      wrap: true,
      size: 'Small',
      spacing: 'None',
      color: PRIORITY_COLOR_MAP_[t['Priority']] || 'default',
    });
  }
  var facts = [];
  ['Requester', 'Owner', 'Site', 'Type'].forEach(function (f) {
    if (t[f]) facts.push({ title: f, value: String(t[f]) });
  });
  if (facts.length) {
    items.push({ type: 'FactSet', facts: facts, spacing: 'Small' });
  }
  if (t['URL']) {
    items.push({ type: 'TextBlock', text: '[Reference link](' + t['URL'] + ')', wrap: true, size: 'Small', spacing: 'Small' });
  }
  return items;
}


/** Posts (as a fresh message — see the comment on CONNECT_GROUPS_TAB_NAME/
 * the Weekly Connect plan for why this isn't a truly edited single
 * message) every `groupName` ticket raised within `range`, to THAT
 * group's own webhook — not a single shared one, since one Weekly Connect
 * can serve more than one recurring meeting/channel. `range` defaults to
 * getCurrentConnectWeekRange() (Wednesday-anchored) only as a fallback for
 * anywhere this gets called with no explicit range at all — the normal
 * path is the widget's own "Post to Teams" button, which always sends a
 * real range the person picked themselves (their actual discussion date
 * moves around, not fixed to any one weekday — see the doPost action
 * below). Silently no-ops if `groupName` isn't blank but has no matching
 * _ConnectGroups webhook, or falls back to the legacy single
 * getTeamsWebhookUrl() if `groupName` itself is blank (a ticket saved
 * before Group existed, or with no groups configured at all yet) —
 * either way, Weekly Connect works fine without Teams, same as the
 * Friday reminder does. Never throws: returns {ok, message} either way,
 * for the caller to relay straight back to whoever clicked the button. */
function postWeeklyConnectToTeams(groupName, range) {
  var teamsWebhookUrl = groupName ? getConnectGroupWebhookUrl(groupName) : getTeamsWebhookUrl();
  var heading = groupName || 'Weekly Connect';
  if (!teamsWebhookUrl) {
    return { ok: false, message: 'No Teams webhook configured for "' + heading + '".' };
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEEKLY_CONNECT_TAB_NAME);
  if (!sheet) return { ok: false, message: 'Weekly Connect tab not found.' };

  range = range || getCurrentConnectWeekRange();
  var tickets = getWeeklyConnectTickets().filter(function (t) {
    if (groupName && (t['Group'] || '') !== groupName) return false;
    // 'T00:00:00' (no zone suffix) forces local-time parsing — a bare
    // 'yyyy-MM-dd' string parses as UTC midnight instead, which silently
    // shifts a day backward once compared against range.start/end (built
    // in the script's own timezone) for any negative-UTC-offset org.
    var d = new Date(t['Date'] + 'T00:00:00');
    return !isNaN(d.getTime()) && d >= range.start && d <= range.end;
  });
  // Oldest first within the range, so the list reads top-to-bottom in the
  // order tickets actually came in.
  tickets.reverse();

  var tz = Session.getScriptTimeZone();
  var rangeLabel = Utilities.formatDate(range.start, tz, 'MMM d') + ' – ' + Utilities.formatDate(range.end, tz, 'MMM d');
  // Real Adaptive Card structure now, not a wall of markdown text lines —
  // a heading, then each ticket as its own visually separated Container
  // (a subtle background box) holding a bold Q#/Issue line, a
  // color-coded Priority line, a FactSet for the other populated fields
  // (aligned label:value columns, not plain text lines), and the URL as
  // a real link. See buildTicketCardItems_'s own comment for why colors
  // are limited to Adaptive Cards' handful of semantic ones — there's no
  // real CSS available here at all, this is the actual ceiling.
  var cardBody = [
    { type: 'TextBlock', text: heading + ' — ' + rangeLabel, wrap: true, size: 'Medium', weight: 'Bolder' },
  ];
  if (!tickets.length) {
    cardBody.push({ type: 'TextBlock', text: 'Nothing logged in this range.', wrap: true, isSubtle: true, spacing: 'Medium' });
  } else {
    tickets.forEach(function (t) {
      cardBody.push({ type: 'Container', style: 'emphasis', spacing: 'Medium', items: buildTicketCardItems_(t) });
    });
  }

  try {
    var response = UrlFetchApp.fetch(teamsWebhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(buildAdaptiveCardPayload_(cardBody)),
      muteHttpExceptions: true, // so a non-2xx response lands here to inspect, instead of throwing
    });
    var statusCode = response.getResponseCode();
    var succeeded = statusCode >= 200 && statusCode < 300;
    // muteHttpExceptions means a wrong payload shape (e.g. a Power Automate
    // flow expecting different field names than the plain {text: ...} this
    // sends) fails SILENTLY otherwise — no exception, nothing in the
    // widget. Still written to _Config too (not just returned), so it's
    // checkable directly in the Sheet without needing this exact click's
    // response, or Apps Script's Executions log.
    var statusMsg = (succeeded ? 'OK' : 'FAILED') + ' (' + statusCode + ') — ' + heading + ' — ' + new Date().toISOString() +
      (succeeded ? '' : ' — response: ' + response.getContentText().slice(0, 300));
    setConfigValue('LastTeamsPostStatus', statusMsg);
    return succeeded
      ? { ok: true, message: 'Posted ' + tickets.length + ' ticket(s) to "' + heading + '".' }
      : { ok: false, message: 'Teams rejected the post (HTTP ' + statusCode + ') — check LastTeamsPostStatus in _Config for details.' };
  } catch (err) {
    setConfigValue('LastTeamsPostStatus', 'ERROR — ' + heading + ' — ' + new Date().toISOString() + ' — ' + err);
    Logger.log('postWeeklyConnectToTeams failed: ' + err);
    return { ok: false, message: 'Could not reach Teams: ' + err };
  }
}

function listVisibleTabs() {
  return listVisibleTabsUncached();
}

// Name of the auto-created config tab holding every dropdown/multiselect's
// option list — an "OptionsKey | Option" table, one row per option,
// grouped by the OptionsKey column into whatever list a _FieldSchema row
// asks for by name (SITES, ASSIGNEES, "Daily Status.Status", etc. — see
// ensureFieldSchemaTab()'s OptionsKey values, which this tab's starting
// rows exactly match).
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
 * Never touched again after that first creation, same as ensureLeaveTab()
 * — except for the header-rename self-heal below, which is safe to keep
 * re-running forever since it's a no-op once the header already matches. */
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
    // This column used to be labeled "List", which read as an unrelated
    // term next to _FieldSchema's "OptionsKey" column — same value, two
    // different names. Purely cosmetic (the code below reads column A by
    // position, never by header text), so renaming it here can't break
    // anything; only touches the header cell, never a data row, and only
    // when it's still the old text.
    var headerCell = sheet.getRange(1, 1);
    if (headerCell.getValue() === 'List') {
      headerCell.setValue('OptionsKey');
    }
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

  var rows = [['OptionsKey', 'Option']];
  Object.keys(defaults).forEach(function (list) {
    defaults[list].forEach(function (option) {
      rows.push([list, option]);
    });
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}

/** Groups _Options' rows by the "OptionsKey" column (matched against
 * _FieldSchema's OptionsKey column of the same name), in the order they
 * appear on the sheet (so reordering rows there reorders the dropdown
 * too). A row with a blank OptionsKey or Option is skipped rather than
 * producing a broken/empty entry. */
function getOptionsMap() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPTIONS_TAB_NAME);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  values.forEach(function (row) {
    var optionsKey = String(row[0]).trim();
    var option = String(row[1]).trim();
    if (!optionsKey || !option) return;
    if (!map[optionsKey]) map[optionsKey] = [];
    map[optionsKey].push(option);
  });
  return map;
}

/** Replaces every data row (everything below the header) on _Options with
 * `map` — used by the widget's in-app "Manage Fields & Options" screen.
 * Only ever called while already holding the write lock (see doPost). Blank
 * keys/options are dropped rather than written as broken rows. */
function writeOptionsMap(map) {
  ensureOptionsTab();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPTIONS_TAB_NAME);
  var rows = [];
  Object.keys(map).forEach(function (key) {
    var optionsKey = String(key).trim();
    if (!optionsKey) return;
    (map[key] || []).forEach(function (opt) {
      var option = String(opt).trim();
      if (option) rows.push([optionsKey, option]);
    });
  });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
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

/** Replaces every data row on _FieldSchema with `map` — used by the
 * widget's in-app "Manage Fields & Options" screen. Only ever called while
 * already holding the write lock (see doPost). A (Tab, Column) with a
 * blank Tab, Column, or Type is dropped rather than written as a broken
 * row. Also creates the actual column on its tab's sheet if it doesn't
 * exist yet (ensureColumnExists_) — this is what lets "define a field's
 * type from the app" genuinely CREATE that field, not just assign a type
 * to a column that has to already exist on the Sheet first. */
function writeFieldSchemaMap(map) {
  ensureFieldSchemaTab();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIELD_SCHEMA_TAB_NAME);
  var rows = [];
  Object.keys(map).forEach(function (tabName) {
    var tab = String(tabName).trim();
    if (!tab) return;
    var columns = map[tabName] || {};
    Object.keys(columns).forEach(function (columnName) {
      var column = String(columnName).trim();
      var spec = columns[columnName] || {};
      var type = String(spec.type || '').trim();
      if (!column || !type) return;
      ensureColumnExists_(tab, column);
      rows.push([tab, column, type, String(spec.optionsKey || '').trim(), String(spec.basedOn || '').trim()]);
    });
  });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 5).setValues(rows);
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

/** Replaces every data row on _Categories with `map` — used by the
 * widget's in-app "Manage Fields & Options" screen. Only ever called while
 * already holding the write lock (see doPost). A category with a blank
 * name, or no tabs left under it, is dropped rather than written as an
 * empty/broken row. */
function writeCategoriesMap(map) {
  ensureCategoriesTab();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CATEGORIES_TAB_NAME);
  var rows = [];
  Object.keys(map).forEach(function (categoryName) {
    var category = String(categoryName).trim();
    if (!category) return;
    (map[categoryName] || []).forEach(function (tabName) {
      var tab = String(tabName).trim();
      if (tab) rows.push([category, tab]);
    });
  });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

// A plain-English, step-by-step guide to everything above, written INTO
// the sheet itself so whoever's editing _Options/_FieldSchema/_Categories
// doesn't need to go find this source file to remember the rules. Starts
// with "_" so it's hidden from the widget's tab list like every other
// config tab — it's for a human reading the Sheet directly, not a widget
// screen.
//
// Unlike _Options/_FieldSchema/_Categories, this tab is pure reference
// content, not something you're expected to hand-edit — so instead of
// "created once, never touched again", it re-writes itself whenever
// FEATURES_GUIDE_VERSION below is bumped (tracked via a note on cell A1,
// invisible in the normal grid view), so an existing deployment's guide
// stays in sync with whatever this function actually knows how to
// document. Bump the version any time you change the rows below.
var FEATURES_TAB_NAME = '_Features';
var FEATURES_GUIDE_VERSION = '11';

function ensureFeaturesTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(FEATURES_TAB_NAME);
  if (sheet) {
    if (sheet.getRange(1, 1).getNote() === FEATURES_GUIDE_VERSION) return; // already current
  } else {
    sheet = ss.insertSheet(FEATURES_TAB_NAME);
  }

  var rows = [
    ['Task', 'Steps', 'Notes'],
    [
      'Use the in-app editor instead of this sheet',
      'Open the widget -> Settings -> "Manage fields & options" (asks for the admin password).',
      'Covers everything below except adding a brand-new tab. Edits the same _Options/_FieldSchema/_Categories rows either way — this sheet stays the source of truth.',
    ],
    [
      'Add a new dropdown/multiselect option',
      '1. Open _Options.\n2. Add a row: OptionsKey = the OptionsKey from _FieldSchema for that field (e.g. "Leave.Type").\n3. Option = the new value.',
      'No new OptionsKey needed if it already exists elsewhere.',
    ],
    [
      'Edit an existing option',
      '1. Find its row in _Options.\n2. Edit the Option cell directly.',
      "Doesn't change past rows already saved with the old text — only future choices.",
    ],
    [
      'Delete an option',
      '1. Delete its row in _Options.',
      'Past entries keep the old value; it just stops being offered.',
    ],
    [
      'Turn a plain-text field into a dropdown/date/checkbox/etc.',
      '1. Open _FieldSchema.\n2. Add a row: Tab = sheet tab name, Column = exact header text.\n3. Type = one of the Field Types below.\n4. OptionsKey = an OptionsKey in _Options (add it there too if new) — only needed for select/multiselect.',
      'See "Field Types" rows below for the full list and what each needs.',
    ],
    [
      'Revert a field back to plain text',
      '1. Delete its row in _FieldSchema.',
      '',
    ],
    [
      'Add a new landing-screen category',
      '1. Open _Categories.\n2. Add rows: Category = new or existing name, Tab = the sheet tab to put under it.',
      'A tab left out of _Categories entirely just won\'t show up on the landing screen until you add it to one — no catch-all bucket for it.',
    ],
    [
      'Add a brand-new tab (a whole new kind of entry)',
      '1. Right-click the tabs bar in Sheets -> Insert sheet.\n2. Name it, then put your column headers in row 1.',
      "Shows up in the widget automatically. To fold it into the Friday emailed report too, add its exact name to _Config's ReportTabs (see the row below) — or use the app's \"Manage Fields & Options\" screen.",
    ],
    [
      'Choose which tabs feed the Friday report',
      '1. Open _Config.\n2. Edit the ReportTabs row: comma-separated exact tab names.',
      'Or use the app\'s "Manage Fields & Options" -> Report Settings, which shows checkboxes instead of typed names.',
    ],
    [
      'Hide a tab from the app (still usable directly in the Sheet)',
      '1. Open _Config.\n2. Edit the HiddenTabs row: comma-separated exact tab names.',
      'Same "Report Settings" screen in the app covers this too. A hidden tab can still be in ReportTabs — hiding it from the app and including it in the report are independent.',
    ],
    [
      'Change who gets the report / Friday reminder emails',
      '1. Open _Config.\n2. Edit ReportRecipients and/or ReminderRecipients: comma-separated emails.',
      'Leave ReminderRecipients blank to reuse ReportRecipients. Also settable from the app\'s "Manage Fields & Options" -> Report Settings.',
    ],
    [
      'See your changes in the app right away',
      '1. Open the app -> Settings -> "Refresh tabs & fields (clear cache)".',
      'Otherwise changes show up on their own within ~5 minutes (CACHE_SECONDS) — the cache is why a change can look like it "didn\'t work" for a few minutes.',
    ],
    [
      'Weekly Connect: log a ticket / update one',
      '1. Widget -> Weekly Connect category -> Weekly_Connect -> pick a Group, fill the rest, submit.\n2. From that same screen, "View & update tickets" -> pick one -> set Status/Comments -> Save.',
      "New tickets auto-post to their Group's own Teams channel (see the row below) as a running list scoped to the current Wed-to-Wed week. Status/Comments are the one thing in this whole app that edits an existing row instead of appending.",
    ],
    [
      'Add a Connect Group (another recurring meeting/channel)',
      '1. Open _ConnectGroups.\n2. Add a row: Group Name = whatever you\'ll call it, Teams Webhook URL = that channel\'s webhook.',
      'Or use the app\'s "Manage Fields & Options" -> Connect Groups. Shows up as a new Group choice on the Weekly_Connect form immediately — no code change. Started with one group, "CMS Weekly Connect" — add as many more as you actually run.',
    ],
    [
      'Teams webhook for the Friday reminder',
      '1. Open _Config.\n2. Edit the TeamsWebhookUrl row.',
      'Separate from Connect Groups above — this one\'s just for the Friday reminder. Or use the app\'s "Manage Fields & Options" -> Report Settings. Leave blank to turn it off.',
    ],
    [
      "What's still a code change (not sheet- or app-editable)",
      '- Reminder schedule (day/time) — see setupTriggers() in Code.gs\n- Adding a genuinely new field Type beyond the ones listed below\n- Rotating SHARED_SECRET itself (the token value the app connects with is set from the app — see the comparison table)',
      'Ask whoever set up the backend for these.',
    ],
    ['— Field Types (_FieldSchema "Type" column) —', '', ''],
    [
      'select',
      'Dropdown, one choice.',
      'Needs OptionsKey pointing at an _Options list.',
    ],
    [
      'multiselect',
      'Tap-to-toggle chips, any number of choices; saved as a comma-separated list.',
      'Needs OptionsKey pointing at an _Options list.',
    ],
    [
      'checkbox',
      'A single Yes/No checkbox. Saves "Yes" when checked, blank when not.',
      'No OptionsKey needed.',
    ],
    [
      'date',
      'A date picker.',
      'No OptionsKey needed.',
    ],
    [
      'week-auto',
      'Read-only text, auto-filled as "Week N" from another field on the same form.',
      'Needs BasedOn = that other column\'s exact header (usually a "date" field). No OptionsKey.',
    ],
    [
      'sequence',
      'Read-only text, auto-numbered: highest existing value in that column on the sheet, plus one.',
      'No OptionsKey needed.',
    ],
    [
      'richtext',
      'A text box with a small Bold / Bullets / Numbered / Link toolbar above it, plus a live rendered preview underneath as you type — and shows fully formatted (not raw **/- syntax) wherever it\'s displayed read-only, e.g. Weekly Connect\'s ticket detail screen.',
      'Stores lightweight Markdown (**bold**, - item, 1. item, [text](url)), not HTML — reads fine as plain text directly in the sheet cell too. No OptionsKey needed.',
    ],
    [
      '(blank / no row at all)',
      'Plain text box (or a taller text area if the column name contains "notes", "description", "blocker", "comment", or "detail").',
      'The default for any field you never add to _FieldSchema.',
    ],
  ];

  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 1, 420);
  sheet.setColumnWidths(3, 1, 260);
  sheet.getRange(2, 1, rows.length - 1, 3).setWrap(true);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setNote(FEATURES_GUIDE_VERSION);

  writeSheetVsAppTable_(sheet);
}

// A color-coded "can I do this in the Sheet, or in the App?" table, placed
// two columns to the right of the main guide (column E on) so both are
// visible side by side without scrolling. Green = yes, that surface can do
// it; amber = no, it's still a real code change either way (not a gap in
// this feature — those things were never made editable from anywhere);
// everything else is a plain, uncolored "No".
function writeSheetVsAppTable_(sheet) {
  var YES = '#d9ead3'; // soft green
  var NO = '#f4cccc'; // soft red
  var CODE = '#fce5cd'; // soft amber — "still a real code change", not a gap

  var startCol = 5; // column E
  var header = ['Where can I make this change?', 'In Sheet', 'In App'];
  var data = [
    ['Add / edit / delete a dropdown option', 'Yes', 'Yes'],
    ["Add / edit / delete a field's Type", 'Yes', 'Yes'],
    ['Add / edit / delete a category', 'Yes', 'Yes'],
    ['Choose which tabs feed the Friday report', 'Yes', 'Yes'],
    ['Hide a tab from the app', 'Yes', 'Yes'],
    ['Change who gets report/reminder emails', 'Yes', 'Yes'],
    ['Set/change the Friday-reminder Teams webhook', 'Yes', 'Yes'],
    ['Add/edit/delete a Weekly Connect group', 'Yes', 'Yes'],
    ['Create a brand-new tab', 'Yes', 'No'],
    ['Rename or delete an existing tab', 'Yes', 'No'],
    ['Change the connection token / admin password', 'No', 'Yes'],
    ['Reminder schedule (day/time) and holidays', 'Code change', 'Code change'],
  ];

  var values = [header].concat(data);
  var range = sheet.getRange(1, startCol, values.length, 3);
  range.setValues(values);
  sheet.getRange(1, startCol, 1, 3).setFontWeight('bold').setBackground('#6e1b2c').setFontColor('#ffffff');

  var backgrounds = data.map(function (row) {
    return ['#ffffff', colorFor_(row[1]), colorFor_(row[2])];
  });
  sheet.getRange(2, startCol, data.length, 3).setBackgrounds(backgrounds);
  sheet.getRange(2, startCol + 1, data.length, 2).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(1, startCol, values.length, 3).setBorder(true, true, true, true, true, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidth(startCol, 300);
  sheet.setColumnWidths(startCol + 1, 2, 110);

  function colorFor_(v) {
    if (v === 'Yes') return YES;
    if (v === 'Code change') return CODE;
    return NO;
  }
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

/** Adds `column` as a new header on `tab`'s actual sheet if it isn't
 * already there (case-insensitive match, so "Site"/"site" don't end up
 * as two separate columns) — a no-op if it already exists. This is what
 * makes "define a field's Type from the app" also able to genuinely
 * CREATE that field, not just assign a type to a column that has to
 * already exist. Matches the promise made in the app/Sheet comparison
 * table (_Features / the Manage screen's own "In App" column) that this
 * is doable from either side, not Sheet-only. No-ops quietly if the tab
 * itself doesn't exist — the caller (writeFieldSchemaMap) already has
 * nothing to attach a schema row to in that case either. */
function ensureColumnExists_(tab, column) {
  var sheet = getSheetByName(tab);
  if (!sheet) return;
  var trimmed = String(column).trim();
  if (!trimmed) return;
  var headers = getHeaderRow(sheet);
  var exists = headers.some(function (h) { return String(h).trim().toLowerCase() === trimmed.toLowerCase(); });
  if (exists) return;
  sheet.getRange(1, sheet.getLastColumn() + 1).setValue(trimmed);
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
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    sheet.hideSheet();
  }
  ensureConfigDefaults_(sheet);
  return sheet;
}

/** Adds any of these keys that are missing — never touches one that's
 * already there, even if its value is blank (blank can be deliberate,
 * e.g. "use the Script Properties fallback"). Runs on every
 * ensureConfigTab() call, so an existing deployment (this one included)
 * migrates in place the first time this runs after a key was added here,
 * without touching anything already on the tab. ReportTabs/HiddenTabs
 * only seed with this app's original hardcoded values for the specific
 * tab names that already exist in THIS spreadsheet — same non-destructive-
 * migration pattern as ensureFieldSchemaTab()/ensureCategoriesTab(), so a
 * brand-new organization gets an honest blank instead of another org's
 * tab names. */
function ensureConfigDefaults_(sheet) {
  var existingTabNames = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) { return s.getName(); });
  function migratedTabList_(names) {
    return names.filter(function (n) { return existingTabNames.indexOf(n) !== -1; }).join(', ');
  }
  var defaults = [
    ['LatestVersion', '1.1.0'],
    ['DownloadUrl', ''],
    ['ReportTabs', migratedTabList_(['Daily Status', 'Adhoc_Mails', 'Cleanup_Activities', 'Drupal_Bugs_&_Improvements'])],
    ['HiddenTabs', migratedTabList_(['Weekly_Monthly_Summary'])],
    ['ReportRecipients', ''],
    ['ReminderRecipients', ''],
    ['TeamsWebhookUrl', ''],
    // 'FALSE' by default on purpose — flipping this to 'TRUE' makes login
    // mandatory for everyone (see the hasAccount action's requireLogin
    // field, and index.html's checkLoginGate). Defaulting to FALSE means
    // shipping the code that reads this never locks anyone out on its
    // own — it only takes effect once you deliberately set it here,
    // after every current teammate already has an account (createUser).
    // Editable right here in the Sheet, or from a future Manage screen.
    ['RequireLogin', 'FALSE'],
  ];
  var data = sheet.getDataRange().getValues();
  var existingKeys = data.slice(1).map(function (r) { return String(r[0]).trim(); });
  var toAdd = defaults.filter(function (d) { return existingKeys.indexOf(d[0]) === -1; });
  if (toAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
  }
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

/** Sets a _Config key's value, adding the row if it somehow doesn't exist
 * yet (shouldn't normally happen — ensureConfigDefaults_ seeds every known
 * key). Used by the widget's in-app "Manage Fields & Options" screen. */
function setConfigValue(key, value) {
  var sheet = ensureConfigTab();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2).setValues([[key, value]]);
}

function getConfigList_(key) {
  return String(getConfigValue(key)).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

/** _Config's HiddenTabs — see the comment above HIDDEN_TABS' old
 * definition near the top of this file. */
function getHiddenTabs() {
  return getConfigList_('HiddenTabs');
}

/** _Config's ReportTabs — see the comment above REPORT_TABS' old
 * definition near the top of this file. */
function getReportTabs() {
  return getConfigList_('ReportTabs');
}

/** _Config's ReportRecipients if set, else the original REPORT_RECIPIENTS
 * Script Property (kept working for any deployment that hasn't moved to
 * _Config for this yet). */
function getReportRecipients() {
  var fromSheet = getConfigList_('ReportRecipients');
  if (fromSheet.length) return fromSheet;
  return getScriptPropList_('REPORT_RECIPIENTS');
}

/** _Config's ReminderRecipients if set, else the REMINDER_RECIPIENTS
 * Script Property override if set, else whatever getReportRecipients()
 * resolves to — same three-tier fallback the Script-Properties-only
 * mechanism always had, just with _Config checked first now. */
function getReminderRecipients() {
  var fromSheet = getConfigList_('ReminderRecipients');
  if (fromSheet.length) return fromSheet;
  var override = getScriptPropList_('REMINDER_RECIPIENTS');
  if (override.length) return override;
  return getReportRecipients();
}

/** _Config's TeamsWebhookUrl if set, else the TEAMS_WEBHOOK_URL Script
 * Property (the original mechanism, kept working for any deployment that
 * already relies on it). Empty string if neither is set — every caller
 * treats that as "Teams is off", never an error. */
function getTeamsWebhookUrl() {
  var fromSheet = getConfigValue('TeamsWebhookUrl');
  return fromSheet || TEAMS_WEBHOOK_URL;
}

// -------------------- Report Configs (multiple, independent reports) --------------------
//
// Generalizes what used to be exactly one hardcoded report (_Config's
// ReportTabs/ReportRecipients, sent every Friday 11pm to a fixed
// address list) into as many independent, named report setups as an
// org actually needs — each with its own tabs, its own selected
// columns per tab, its own recipients, and (Sprint 2) its own
// schedule. A client that only ever wants the one report they have
// today still gets exactly that; one that wants a second report for a
// different team/category with different fields and a different
// audience can just add another row here — no code change, same as
// everything else in this file.
//
// One row per (ConfigName, Tab) pair — a config spanning several tabs
// is just several rows sharing the same ConfigName, same shape as
// _FieldSchema being one row per (Tab, Column). Fields is a
// comma-separated column allowlist for THAT tab within THAT config;
// blank means "every column on that tab", which is what preserves
// today's exact behavior for the migrated default config below.
var REPORT_CONFIGS_TAB_NAME = '_ReportConfigs';
var REPORT_CONFIGS_COLUMNS = ['ConfigName', 'Tab', 'Fields', 'Recipients', 'ScheduleDay', 'ScheduleHour', 'Enabled'];

/** Created once. If _Config already had a ReportTabs/ReportRecipients
 * setup (every existing deployment, this one included), migrates it
 * into a single "Report Generator" config — one row per tab already in
 * ReportTabs, Fields left blank (= all columns, so this migration
 * changes nothing about what the existing report contains), same
 * recipients, and the same Friday 11pm schedule the fixed trigger
 * already sends at. A brand-new organization with nothing in ReportTabs
 * yet just gets an empty tab with headers, same "honest blank instead
 * of someone else's setup" rule every other ensure*Tab() here follows. */
function ensureReportConfigsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(REPORT_CONFIGS_TAB_NAME)) return;
  var sheet = ss.insertSheet(REPORT_CONFIGS_TAB_NAME);
  sheet.getRange(1, 1, 1, REPORT_CONFIGS_COLUMNS.length).setValues([REPORT_CONFIGS_COLUMNS]);

  var migratedTabs = getReportTabs();
  if (!migratedTabs.length) return; // nothing to migrate — leave it empty
  var migratedRecipients = getReportRecipients().join(', ');
  var rows = migratedTabs.map(function (tab) {
    // Named 'Status Report' (not 'Report Generator', the app's own
    // landing-screen category name) specifically because the email
    // subject line uses config.name as its prefix now — this is what
    // keeps the migrated default's subject text byte-for-byte identical
    // to what it's always been, so nobody's inbox filter/rule keyed on
    // "Status Report" silently breaks the first time this runs.
    return ['Status Report', tab, '', migratedRecipients, 'Friday', 23, 'TRUE'];
  });
  sheet.getRange(2, 1, rows.length, REPORT_CONFIGS_COLUMNS.length).setValues(rows);
  // The migrated config already carries the old fixed Friday/23
  // schedule — set up the new hourly dispatcher (and retire the old
  // fixed trigger it replaces) right away, in the same breath as the
  // migration itself, rather than waiting on someone to open Manage and
  // hit Save first. See ensureScheduledReportsTrigger_'s own comment for
  // why this is one shared trigger, not one per config.
  ensureScheduledReportsTrigger_();
}

/** Every report config, grouped from the flat sheet rows into
 * {name, tabs: [{tab, fields}], recipients, scheduleDay, scheduleHour,
 * enabled} objects, in the sheet's own row order (first-seen ConfigName
 * order, tabs in the order their rows appear). `fields` is [] for "every
 * column on that tab" (a blank Fields cell), or the explicit allowlist
 * otherwise. Config-level fields (Recipients/Schedule/Enabled) are read
 * off that config's FIRST row — every row for the same config is
 * expected to repeat the same values, same as how a duplicate-detection
 * problem would show up in Field Types if they ever drifted, but nothing
 * here currently guards against that beyond the Manage screen always
 * writing them consistently in the first place. */
function getReportConfigs() {
  ensureReportConfigsTab();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_CONFIGS_TAB_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, REPORT_CONFIGS_COLUMNS.length).getValues();
  var byName = {};
  var order = [];
  values.forEach(function (row) {
    var name = String(row[0]).trim();
    var tab = String(row[1]).trim();
    if (!name || !tab) return;
    if (!byName[name]) {
      byName[name] = {
        name: name,
        tabs: [],
        recipients: String(row[3]).split(',').map(function (s) { return s.trim(); }).filter(Boolean),
        scheduleDay: String(row[4]).trim(),
        scheduleHour: row[5] === '' || row[5] === null ? null : Number(row[5]),
        enabled: String(row[6]).trim().toUpperCase() === 'TRUE',
      };
      order.push(name);
    }
    var fields = String(row[2]).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    byName[name].tabs.push({ tab: tab, fields: fields });
  });
  return order.map(function (name) { return byName[name]; });
}

/** One config by name, or the first one if `name` is blank/not found —
 * same "no name given, use whatever's first" default every report
 * action (sendReportNow/downloadReport/sendTestReport) falls back to for
 * anyone not yet sending an explicit configName, so an org with exactly
 * one report (the common case) never has to think about this at all. */
function getReportConfigByName_(name) {
  var configs = getReportConfigs();
  if (!configs.length) return null;
  if (!name) return configs[0];
  var match = configs.filter(function (c) { return c.name === name; })[0];
  return match || configs[0];
}

/** Replaces every row on _ReportConfigs — same wholesale-replace shape
 * as writeFieldSchemaMap/writeOptionsMap. Takes the same structured
 * shape getReportConfigs() returns, flattens back into one row per
 * (ConfigName, Tab). A config with no tabs, or a blank name, is dropped
 * rather than written as a broken row. */
function writeReportConfigs(configs) {
  ensureReportConfigsTab();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_CONFIGS_TAB_NAME);
  var rows = [];
  (configs || []).forEach(function (config) {
    var name = String(config.name || '').trim();
    if (!name) return;
    var recipients = (config.recipients || []).join(', ');
    var scheduleDay = String(config.scheduleDay || '').trim();
    var scheduleHour = config.scheduleHour === '' || config.scheduleHour == null ? '' : Number(config.scheduleHour);
    var enabled = config.enabled ? 'TRUE' : 'FALSE';
    (config.tabs || []).forEach(function (t) {
      var tab = String(t.tab || '').trim();
      if (!tab) return;
      var fields = (t.fields || []).join(', ');
      rows.push([name, tab, fields, recipients, scheduleDay, scheduleHour, enabled]);
    });
  });
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, REPORT_CONFIGS_COLUMNS.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, REPORT_CONFIGS_COLUMNS.length).setValues(rows);
  ensureScheduledReportsTrigger_();
}

// -------------------- configurable per-config scheduling --------------------
//
// Deliberately ONE trigger, shared by every config that has a schedule,
// rather than a separate trigger per config. A per-config trigger would
// mean creating/updating/deleting an actual ScriptApp trigger every
// single time a schedule changes in the Manage screen — real risk of
// orphaned triggers left behind if a config gets renamed or removed
// (nothing left to tell the old trigger to clean itself up), on top of
// Apps Script's own per-project trigger count limit. This way there's
// nothing to keep in sync at all: the trigger itself never changes,
// only what runScheduledReports_() finds in _ReportConfigs when it
// fires does — so editing/adding/removing a config's schedule just
// takes effect on the next hourly check, with zero trigger management.
var SCHEDULED_REPORTS_TRIGGER_FN = 'runScheduledReports_';

/** Idempotent — a no-op if the trigger already exists, never a
 * duplicate. Called from saveReportConfigs so turning on a schedule for
 * the first time (or after every config was previously manual-only)
 * doesn't need a separate "now go set this up" step anywhere else -
 * same self-healing spirit as every ensure*Tab() in this file, just for
 * a trigger instead of a sheet. Uses the SAME script.scriptapp OAuth
 * scope Weekly Connect's original trigger design already required (see
 * appsscript.json) - already granted from that one-time
 * re-authorization, so this doesn't need it done again. */
function ensureScheduledReportsTrigger_() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  var exists = existingTriggers.some(function (t) {
    return t.getHandlerFunction() === SCHEDULED_REPORTS_TRIGGER_FN;
  });
  if (!exists) {
    ScriptApp.newTrigger(SCHEDULED_REPORTS_TRIGGER_FN).timeBased().everyHours(1).create();
  }
  // The OLD fixed Friday-11pm trigger (scheduledCompileAndSendReport,
  // from setupTriggers()) is now fully superseded by the hourly
  // dispatcher above — the migrated "Status Report" config already has
  // that exact same Friday/23 schedule, so leaving the old trigger in
  // place would send the SAME report TWICE every Friday, not just be
  // redundant. Removed here in code, not left for a manual
  // setupTriggers() re-run, so this can't sit in a double-sending state
  // waiting on someone to notice. setupTriggers() itself no longer
  // creates this one either, going forward — see its own comment.
  existingTriggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'scheduledCompileAndSendReport') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** Fired roughly hourly — Apps Script's own time-based triggers have
 * some natural jitter and aren't guaranteed to fire at exactly :00, so
 * this checks "does the current hour match" rather than needing
 * exact-minute precision. Sends every enabled config whose
 * ScheduleDay/ScheduleHour matches right now; a config with no schedule
 * set (blank ScheduleDay or ScheduleHour) is manual-only and never
 * matches here, same as it working today via the widget's own "Send
 * Report Now" only. Each config's send is wrapped in its own try/catch —
 * one failing (a bad recipient address, anything) should never skip the
 * rest, same reasoning as Weekly Connect's multi-group Teams posting. */
function runScheduledReports_() {
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var currentDay = Utilities.formatDate(now, tz, 'EEEE'); // 'Monday', 'Tuesday', ...
  var currentHour = Number(Utilities.formatDate(now, tz, 'H'));

  getReportConfigs().forEach(function (config) {
    if (!config.enabled) return;
    if (!config.scheduleDay || config.scheduleHour === null || isNaN(config.scheduleHour)) return;
    if (config.scheduleDay !== currentDay) return;
    if (config.scheduleHour !== currentHour) return;
    try {
      compileAndSendReport(null, config.name);
    } catch (err) {
      Logger.log('runScheduledReports_ failed for config "' + config.name + '": ' + err);
    }
  });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
