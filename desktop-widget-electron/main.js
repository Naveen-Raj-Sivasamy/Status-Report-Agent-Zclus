const { app, Tray, Menu, BrowserWindow, ipcMain, screen, nativeImage, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const APP_VERSION = require('./package.json').version;

// Dev convenience: a full config.json next to main.js (gitignored) overrides
// everything, including YOUR_NAME — this is what `npm start` uses locally.
const CONFIG_PATH = path.join(__dirname, 'config.json');
// Shared connection details (WEBHOOK_URL/TOKEN/SHEET_URL) baked into the
// installer at build time — an org building their OWN internal-only
// installer (not distributing through a store) can still do this; it's
// just no longer the only way to configure a connection.
const TEMPLATE_PATH = path.join(__dirname, 'config.template.json');
// Per-person name, written on first run into each teammate's own profile
// folder — never baked into the installer, never shared between installs.
const USER_CONFIG_PATH = path.join(app.getPath('userData'), 'user-config.json');
const POSITION_PATH = path.join(app.getPath('userData'), 'float-position.json');
// Connection details entered through the app itself (the "Connect to your
// organization" screen) — this is what makes a store-distributed build
// (no config.template.json baked in at all) work: whoever installs it
// enters their own org's WEBHOOK_URL/TOKEN/SHEET_URL at first launch,
// nothing about which organization this is is compiled into the binary.
// Password-gated to change once set — see hashPassword()/verifyPassword().
const CONNECTION_CONFIG_PATH = path.join(app.getPath('userData'), 'connection-config.json');

function loadSharedConfig() {
  // Order matters: an existing dev config.json or a baked-in
  // config.template.json (an org's own internal build) both still work
  // exactly as before, unchanged. Only when NEITHER exists does this fall
  // through to whatever's been entered at runtime through the app's own
  // "Connect" screen — and if that's empty too, returns null instead of
  // throwing, so app.whenReady() can show that screen instead of just
  // erroring out and quitting.
  if (fs.existsSync(CONFIG_PATH)) {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  }
  if (fs.existsSync(TEMPLATE_PATH)) {
    return JSON.parse(fs.readFileSync(TEMPLATE_PATH, 'utf8'));
  }
  var saved = loadConnectionConfig();
  if (saved) {
    return { WEBHOOK_URL: saved.WEBHOOK_URL, TOKEN: saved.TOKEN, SHEET_URL: saved.SHEET_URL };
  }
  return null;
}

// -------------------- runtime connection config + admin password --------------------

function loadConnectionConfig() {
  try {
    return JSON.parse(fs.readFileSync(CONNECTION_CONFIG_PATH, 'utf8'));
  } catch {
    return null;
  }
}

function hashPassword(password) {
  var salt = crypto.randomBytes(16).toString('hex');
  var hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return { salt: salt, hash: hash };
}

function verifyPassword(password, salt, hash) {
  var candidate = crypto.scryptSync(password, salt, 64).toString('hex');
  var a = Buffer.from(candidate, 'hex');
  var b = Buffer.from(hash, 'hex');
  // Length check first — timingSafeEqual throws on mismatched lengths
  // rather than just returning false, and a wrong-length hash should
  // never happen anyway (scrypt output length is fixed), so this isn't
  // itself a timing leak worth worrying about.
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

// Saves WEBHOOK_URL/TOKEN/SHEET_URL (+ the password hash, new or carried
// over from before) to disk and updates the in-memory `config` so the
// running app picks up the change immediately without a restart — same
// pattern saveYourNameToDisk() already uses for YOUR_NAME.
function saveConnectionConfig(conn) {
  var existing = loadConnectionConfig();
  var passwordFields;
  if (conn.newPassword) {
    passwordFields = hashPassword(conn.newPassword);
  } else if (existing) {
    passwordFields = { salt: existing.salt, hash: existing.hash };
  } else {
    throw new Error('A password is required the first time you connect.');
  }
  var toSave = {
    WEBHOOK_URL: conn.WEBHOOK_URL,
    TOKEN: conn.TOKEN,
    SHEET_URL: conn.SHEET_URL,
    salt: passwordFields.salt,
    hash: passwordFields.hash,
  };
  fs.writeFileSync(CONNECTION_CONFIG_PATH, JSON.stringify(toSave));
  config = Object.assign({}, config, {
    WEBHOOK_URL: toSave.WEBHOOK_URL,
    TOKEN: toSave.TOKEN,
    SHEET_URL: toSave.SHEET_URL,
  });
}

// Both YOUR_NAME and the openAtLogin preference below live in this same
// per-person file — read-modify-write through these two helpers so saving
// one never clobbers the other.
function loadUserConfig() {
  try {
    return JSON.parse(fs.readFileSync(USER_CONFIG_PATH, 'utf8'));
  } catch {
    return {};
  }
}

function saveUserConfig(patch) {
  try {
    fs.writeFileSync(USER_CONFIG_PATH, JSON.stringify(Object.assign(loadUserConfig(), patch)));
  } catch {
    /* best-effort */
  }
}

function loadSavedName() {
  return loadUserConfig().YOUR_NAME || '';
}

function saveYourNameToDisk(name) {
  saveUserConfig({ YOUR_NAME: name });
}

// null here means "nothing configured yet, and that's fine" — no
// config.json/config.template.json AND nothing saved from the Connect
// screen — app.whenReady() below shows that screen instead of erroring
// out. A THROWN error (e.g. malformed JSON in an existing file) is a
// different, real problem that can't be fixed by re-running Connect, so
// that path still surfaces a dialog and quits like before.
let config;
let configLoadFailed = false; // distinguishes "genuinely broken file" from "nothing configured yet" below — both leave config null
try {
  config = loadSharedConfig();
  if (config && !config.YOUR_NAME) config.YOUR_NAME = loadSavedName();
} catch (err) {
  app.whenReady().then(() => {
    dialog.showErrorBox('Status Update — setup needed', err.message);
    app.quit();
  });
  config = null;
  configLoadFailed = true;
}

let tray = null;
let popup = null;
let floatBtn = null;
let setupWin = null;
let connectWin = null;
let manageWin = null;

// -------------------- backend calls (with retry) --------------------

// Was 6 attempts x a 45s timeout each (worst case ~4.5 minutes of the
// widget silently showing "Saving...") — that's what made a single slow
// backend call (e.g. Sheets doing a full-workbook formula recalc, or many
// people submitting near the Friday cutoff) look like it was stuck forever,
// identically for everyone hitting the same backend. Fail faster instead,
// and let the caller report retry progress so the UI can say what's
// actually happening rather than a blank "Saving...".
const RETRY_ATTEMPTS = 3;
const REQUEST_TIMEOUT_MS = 15000;

async function callWithRetry(fn, { attempts = RETRY_ATTEMPTS, onRetry } = {}) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      if (i < attempts) {
        if (onRetry) onRetry(i, attempts, err);
        // Exponential backoff (500ms, 1s, 2s, ...): fast to recover from
        // a single blip, patient enough for a slow cold start.
        const delay = Math.min(500 * 2 ** (i - 1), 4000);
        await new Promise((r) => setTimeout(r, delay));
      }
    }
  }
  throw lastErr;
}

async function apiGet(params, opts) {
  const url = new URL(config.WEBHOOK_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  return callWithRetry(async () => {
    const resp = await fetch(url.toString(), { signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    if (!data.ok) throw new Error(data.error || 'Unknown API error');
    return data;
  }, opts);
}

async function apiPostBody(body, opts) {
  return callWithRetry(async () => {
    const resp = await fetch(config.WEBHOOK_URL, {
      method: 'POST',
      body: JSON.stringify(Object.assign({ token: config.TOKEN }, body)),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    if (!data.ok) throw new Error(data.error || 'Unknown API error');
    return data;
  }, opts);
}

async function apiPost(tab, values, opts) {
  return apiPostBody({ tab, values }, opts);
}

// In-memory cache the widget serves from instantly, refreshed in the
// background — the popup opening should never wait on a network round trip
// just to show the tab list it already showed a minute ago. The Apps
// Script side still caches for 5 minutes too, so a background refresh here
// is cheap even when it does need to hit the network.
let tabsCache = null;
let columnsCache = {}; // tab -> columns[]
let optionsCache = null; // dropdown/multiselect option lists from the _Options tab
let fieldSchemaCache = null; // which fields get which widget type, from the _FieldSchema tab
let categoriesCache = null; // landing-screen tab grouping, from the _Categories tab

async function refreshTabsCache() {
  try {
    const data = await apiGet({ action: 'tabs' });
    tabsCache = data;
    return data;
  } catch (err) {
    if (tabsCache) return tabsCache; // stale is better than nothing
    throw err;
  }
}

async function refreshColumnsCache(tab) {
  try {
    const data = await apiGet({ action: 'columns', tab });
    columnsCache[tab] = data;
    return data;
  } catch (err) {
    if (columnsCache[tab]) return columnsCache[tab];
    throw err;
  }
}

async function refreshOptionsCache() {
  try {
    const data = await apiGet({ action: 'options' });
    optionsCache = data;
    return data;
  } catch (err) {
    if (optionsCache) return optionsCache; // stale is better than nothing
    throw err;
  }
}

async function refreshFieldSchemaCache() {
  try {
    const data = await apiGet({ action: 'fieldSchema' });
    fieldSchemaCache = data;
    return data;
  } catch (err) {
    if (fieldSchemaCache) return fieldSchemaCache; // stale is better than nothing
    throw err;
  }
}

async function refreshCategoriesCache() {
  try {
    const data = await apiGet({ action: 'categories' });
    categoriesCache = data;
    return data;
  } catch (err) {
    if (categoriesCache) return categoriesCache; // stale is better than nothing
    throw err;
  }
}

function prefetchAll() {
  refreshTabsCache()
    .then((data) => Promise.all((data.tabs || []).map((t) => refreshColumnsCache(t))))
    .catch(() => {
      /* best-effort — a real request will retry when the user actually opens the form */
    });
  refreshOptionsCache().catch(() => {
    /* same — a real request will retry when a form actually needs it */
  });
  refreshFieldSchemaCache().catch(() => {
    /* same */
  });
  refreshCategoriesCache().catch(() => {
    /* same */
  });
}

ipcMain.handle('get-tabs', async () => {
  if (tabsCache) {
    refreshTabsCache(); // stale-while-revalidate: return now, update quietly
    return tabsCache;
  }
  return refreshTabsCache();
});
ipcMain.handle('get-columns', async (_e, tab) => {
  if (columnsCache[tab]) {
    refreshColumnsCache(tab);
    return columnsCache[tab];
  }
  return refreshColumnsCache(tab);
});
ipcMain.handle('get-options', async () => {
  if (optionsCache) {
    refreshOptionsCache();
    return optionsCache;
  }
  return refreshOptionsCache();
});
ipcMain.handle('get-field-schema', async () => {
  if (fieldSchemaCache) {
    refreshFieldSchemaCache();
    return fieldSchemaCache;
  }
  return refreshFieldSchemaCache();
});
ipcMain.handle('get-categories', async () => {
  if (categoriesCache) {
    refreshCategoriesCache();
    return categoriesCache;
  }
  return refreshCategoriesCache();
});
ipcMain.handle('submit-entry', async (event, { tab, values }) => {
  const result = await apiPost(tab, values, {
    onRetry: (attempt, attempts) => {
      // Let the popup show real progress ("retrying 2/3") instead of a
      // static "Saving..." that looks frozen while we're still working on it.
      if (!event.sender.isDestroyed()) {
        event.sender.send('submit-retry', { attempt: attempt + 1, attempts });
      }
    },
  });
  refreshColumnsCache(tab); // e.g. so the next "Cleanup Number" reflects this new row right away
  return result;
});
ipcMain.handle('send-report-now', async (_e, range) =>
  apiPostBody(Object.assign({ action: 'sendReportNow' }, range || {}))
);
ipcMain.handle('clear-cache', async () => {
  const result = await apiPostBody({ action: 'clearCache' });
  // Clearing just the server-side cache isn't enough on its own — this
  // process also keeps its own in-memory copy (see refreshTabsCache /
  // refreshColumnsCache above) that outlives any single popup open/close.
  // Drop both here so the very next tab-list/columns request goes fully
  // fresh instead of quietly serving stale data for the rest of this run.
  tabsCache = null;
  columnsCache = {};
  optionsCache = null;
  fieldSchemaCache = null;
  categoriesCache = null;
  return result;
});
// Powers the in-app "Manage Fields & Options" screen — each just posts the
// client's full, current map back to the backend (which replaces that
// tab's data rows wholesale, see writeOptionsMap() etc. in Code.gs), then
// drops this process's own in-memory copy so the very next read is fresh
// instead of serving what's now stale for up to CACHE_SECONDS.
ipcMain.handle('save-options', async (_e, options) => {
  const result = await apiPostBody({ action: 'saveOptions', options });
  optionsCache = null;
  return result;
});
ipcMain.handle('save-field-schema', async (_e, fieldSchema) => {
  const result = await apiPostBody({ action: 'saveFieldSchema', fieldSchema });
  fieldSchemaCache = null;
  return result;
});
ipcMain.handle('save-categories', async (_e, categories) => {
  const result = await apiPostBody({ action: 'saveCategories', categories });
  categoriesCache = null;
  return result;
});
// Report settings (ReportTabs/HiddenTabs/ReportRecipients/
// ReminderRecipients) go through doPost on the backend, not doGet like
// tabs/options/fieldSchema/categories — see the comment on doGet's
// 'getReportSettings' omission in Code.gs. No in-memory cache here either:
// this screen is opened rarely enough that a live read each time is fine,
// and there's nothing else in this process that reads these values to
// keep in sync.
ipcMain.handle('get-report-settings', async () => apiPostBody({ action: 'getReportSettings' }));
ipcMain.handle('save-report-settings', async (_e, settings) => {
  const result = await apiPostBody({ action: 'saveReportSettings', settings });
  tabsCache = null; // HiddenTabs can change which tabs the widget shows
  return result;
});
// Connect Groups — same doPost/token-gated reasoning as report settings
// (a webhook URL is a write capability). Group names are synced into
// _Options server-side, so a save here also invalidates this process's
// options cache — otherwise the ticket form's Group dropdown would keep
// showing whatever list was cached before the edit.
ipcMain.handle('get-connect-groups', async () => apiPostBody({ action: 'getConnectGroups' }));
ipcMain.handle('save-connect-groups', async (_e, groups) => {
  const result = await apiPostBody({ action: 'saveConnectGroups', groups });
  optionsCache = null;
  return result;
});
// Weekly Connect's ticket list — same reasoning as report settings above
// for going through GET on the backend, not a doPost round-trip: it's a
// plain GET action there (see Code.gs), fetched fresh every time this
// screen opens rather than cached, since "pick a ticket to update" needs
// the real current state.
ipcMain.handle('get-weekly-connect-tickets', async () => apiGet({ action: 'weeklyConnectTickets' }));
ipcMain.handle('update-weekly-connect-ticket', async (_e, { ticketId, status, comments }) =>
  apiPostBody({ action: 'updateWeeklyConnectTicket', ticketId, status, comments })
);
ipcMain.handle('download-report', async (_e, range) => {
  const data = await apiPostBody(Object.assign({ action: 'downloadReport' }, range || {}));
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: 'Save Report',
    defaultPath: data.fileName,
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
  });
  if (canceled || !filePath) return { canceled: true };
  fs.writeFileSync(filePath, Buffer.from(data.base64, 'base64'));
  return { canceled: false, filePath };
});
ipcMain.handle('get-next-number', async (_e, { tab, column }) => apiGet({ action: 'nextNumber', tab, column }));
ipcMain.handle('get-your-name', () => config.YOUR_NAME || '');
ipcMain.handle('get-app-version', () => APP_VERSION);
ipcMain.handle('check-latest-version', async () => {
  try {
    // Unlike a save, nobody is staring at a "Saving..." button waiting on
    // this — it runs once in the background at launch and only ever
    // updates a small badge. So it can afford to be more patient than
    // RETRY_ATTEMPTS (3) and just quietly retry longer instead of giving
    // up and silently hiding the update banner the moment the backend has
    // one slow beat — which is what made the banner flicker in and out
    // across relaunches even though a newer version really was available.
    return await apiGet({ action: 'version' }, { attempts: 6 });
  } catch {
    return null; // still never block the app on a version check
  }
});
ipcMain.handle('open-external', (_e, url) => {
  if (url) shell.openExternal(url);
});
ipcMain.handle('hide-window', () => hidePopup());

// The popup's content height varies a lot (4 tabs vs. 12, a short error
// screen vs. the report-options form) — rather than a fixed window height
// that leaves a dead gap for short screens or clips a tall one, the
// renderer measures its own content and asks us to resize to fit, on
// every render. Clamped so it can never grow absurdly tall or shrink to
// nothing; content beyond the max just scrolls (main already does that).
const MIN_POPUP_HEIGHT = 317; // 264 +20%
const MAX_POPUP_HEIGHT = 922; // 768 +20%
ipcMain.on('resize-window', (_e, height) => {
  if (!popup) return;
  const clamped = Math.max(MIN_POPUP_HEIGHT, Math.min(MAX_POPUP_HEIGHT, Math.round(height)));
  const [width, currentHeight] = popup.getSize();
  if (clamped === currentHeight) return;
  popup.setSize(width, clamped);
  if (popup.isVisible()) {
    const refBounds = floatBtn ? floatBtn.getBounds() : tray.getBounds();
    positionPopupNearWindow(refBounds);
  }
});
// Manual drag for the floating icon (see floatbtn.html) — the renderer
// tracks the mouse itself and sends screen-pixel deltas here rather than
// us relying on -webkit-app-region: drag on the icon, which on Windows
// ended up swallowing plain clicks along with actual drags. setPosition
// still fires the window's own 'moved' event same as a native OS drag
// would, so the existing debounced position-save in createFloatButton
// keeps working unchanged.
ipcMain.on('float-btn-move-by', (_e, { dx, dy }) => {
  if (!floatBtn) return;
  const [x, y] = floatBtn.getPosition();
  floatBtn.setPosition(Math.round(x + dx), Math.round(y + dy));
});
// The floating icon's window is `transparent: true`, but a transparent
// Electron window is NOT click-through by default — its entire rectangle
// (67x67, bigger than the 55x55 visible circle — see floatbtn.html) blocks
// every click to whatever's underneath it on the desktop/other windows,
// even on the fully-transparent pixels. setIgnoreMouseEvents(true) makes
// the whole window pass clicks straight through to the OS instead; the
// renderer flips it back to false while the cursor is actually over the
// visible icon (see floatbtn.html's mouseenter/mouseleave), so the icon
// itself stays clickable/draggable but nothing around it is a dead zone
// any more. `forward: true` keeps mousemove events reaching the renderer
// even while ignoring, which is what lets it notice the cursor entering
// the icon in the first place.
ipcMain.on('float-btn-set-ignore-mouse', (_e, ignore) => {
  if (!floatBtn) return;
  floatBtn.setIgnoreMouseEvents(!!ignore, { forward: true });
});
ipcMain.handle('open-sheet', () => {
  if (config.SHEET_URL) shell.openExternal(config.SHEET_URL);
});
ipcMain.handle('fab-clicked', () => toggleWindow());
ipcMain.handle('show-fab-menu', () => {
  Menu.buildFromTemplate(buildAppMenuTemplate()).popup({ window: floatBtn });
});
ipcMain.handle('save-your-name', (_e, name) => {
  config.YOUR_NAME = name;
  saveYourNameToDisk(name);
  if (setupWin) {
    setupWin.close();
    setupWin = null;
  }
  startMainApp();
});

// -------------------- connect-to-your-organization --------------------

ipcMain.handle('has-connection-config', () => !!loadConnectionConfig());

ipcMain.handle('verify-admin-password', (_e, password) => {
  var saved = loadConnectionConfig();
  if (!saved) return false; // nothing set yet — see the comment on saveConnectionConfig()
  return verifyPassword(password, saved.salt, saved.hash);
});

// Deliberately never returns the salt/hash to the renderer — those never
// need to leave this process.
ipcMain.handle('get-connection-config', () => {
  var saved = loadConnectionConfig();
  return saved ? { WEBHOOK_URL: saved.WEBHOOK_URL, TOKEN: saved.TOKEN, SHEET_URL: saved.SHEET_URL } : null;
});

ipcMain.handle('save-connection-config', (event, conn) => {
  try {
    saveConnectionConfig(conn);
  } catch (err) {
    return { ok: false, error: err.message };
  }
  // Only the FIRST-run Connect window (no config existed before this
  // call) needs to advance the app past onboarding — a reconfigure from
  // the running app's own Settings screen just reports success back and
  // lets that screen decide what to show next.
  if (connectWin && !connectWin.isDestroyed() && event.sender === connectWin.webContents) {
    connectWin.close();
    connectWin = null;
    if (!config.YOUR_NAME) {
      setupWin = createSetupWindow();
    } else {
      startMainApp();
    }
  }
  return { ok: true };
});

ipcMain.handle('open-connection-settings', () => {
  if (!connectWin || connectWin.isDestroyed()) {
    connectWin = createConnectWindow();
  } else {
    connectWin.show();
    connectWin.focus();
  }
});

function createConnectWindow() {
  var win = new BrowserWindow({
    width: 420,
    height: 480,
    resizable: false,
    fullscreenable: false,
    minimizable: false,
    maximizable: false,
    frame: false, // native Windows chrome rendered as a plain white/blue
    // title bar that didn't match the app's own maroon branding anywhere
    // else — connect.html now draws its own title bar/close button
    // instead, same as the main popup already does.
    title: 'Connect — Report Generator',
    backgroundColor: '#6e1b2c',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.setMenuBarVisibility(false);
  win.loadFile(path.join(__dirname, 'renderer', 'connect.html'));
  win.on('closed', () => {
    if (connectWin === win) connectWin = null;
  });
  return win;
}

// -------------------- in-app "Manage Fields & Options" --------------------

// Reuses the exact same admin password as connecting to a new backend —
// verifyAdminPassword above — rather than a second password to set up and
// remember. A separate resizable window (like createConnectWindow), not a
// screen inside the popup: this can get tall (many options/fields/
// categories), and keeping it independent of the popup's own show*()
// screen-switching state avoids tangling two very different UIs together.
ipcMain.handle('open-manage-screen', () => {
  if (!manageWin || manageWin.isDestroyed()) {
    manageWin = createManageWindow();
  } else {
    manageWin.show();
    manageWin.focus();
  }
});

function createManageWindow() {
  var win = new BrowserWindow({
    width: 560,
    height: 700,
    minWidth: 480,
    minHeight: 420,
    fullscreenable: false,
    frame: false, // see the same comment in createConnectWindow() above
    title: 'Manage Fields & Options',
    backgroundColor: '#6e1b2c',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.setMenuBarVisibility(false);
  win.loadFile(path.join(__dirname, 'renderer', 'manage.html'));
  win.on('closed', () => {
    if (manageWin === win) manageWin = null;
  });
  return win;
}

// -------------------- first-run "what's your name" prompt --------------------

function createSetupWindow() {
  const win = new BrowserWindow({
    width: 360,
    height: 260,
    resizable: false,
    fullscreenable: false,
    minimizable: false,
    maximizable: false,
    title: 'Welcome — Report Generator',
    backgroundColor: '#6e1b2c',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.setMenuBarVisibility(false);
  win.loadFile(path.join(__dirname, 'renderer', 'setup.html'));
  win.on('closed', () => {
    if (setupWin === win) setupWin = null;
  });
  return win;
}

// -------------------- popup (the form) --------------------

function createPopup() {
  const win = new BrowserWindow({
    width: 456, // 380 +20%
    height: 605, // 504 +20% — corrected to fit real content right after first paint — see resize-window
    show: false,
    frame: false,
    resizable: false,
    fullscreenable: false,
    skipTaskbar: true,
    backgroundColor: '#00000000',
    transparent: true,
    alwaysOnTop: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.setAlwaysOnTop(true, 'screen-saver'); // stay above everything, including other apps
  win.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  // Deliberately no hide-on-blur: closing is only via the ✕ button, so
  // switching to another app to look something up doesn't lose your form.
  return win;
}

function positionPopupNearWindow(refBounds) {
  const winBounds = popup.getBounds();
  const display = screen.getDisplayNearestPoint({ x: refBounds.x, y: refBounds.y });
  const workArea = display.workArea;

  let x = Math.round(refBounds.x + refBounds.width / 2 - winBounds.width / 2);
  let y;
  if (refBounds.y < workArea.y + workArea.height / 2) {
    y = Math.round(refBounds.y + refBounds.height + 8);
  } else {
    y = Math.round(refBounds.y - winBounds.height - 8);
  }
  x = Math.min(Math.max(x, workArea.x + 8), workArea.x + workArea.width - winBounds.width - 8);
  y = Math.min(Math.max(y, workArea.y + 8), workArea.y + workArea.height - winBounds.height - 8);
  popup.setPosition(x, y, false);
}

function hidePopup() {
  popup.hide();
  if (floatBtn) floatBtn.show(); // bring the floating icon back once the form is closed
}

function toggleWindow() {
  if (popup.isVisible() && popup.isFocused()) {
    hidePopup();
    return;
  }
  if (popup.isVisible() && !popup.isFocused()) {
    // It's open but buried behind another app — bring it back instead of
    // hiding it, so clicking the button always gets you back to your form.
    popup.focus();
    return;
  }
  const refBounds = floatBtn ? floatBtn.getBounds() : tray.getBounds();
  positionPopupNearWindow(refBounds);
  if (floatBtn) floatBtn.hide(); // avoid showing both the icon and the open form at once
  popup.show();
  popup.focus();
  popup.webContents.send('opened');
}

// -------------------- floating draggable button --------------------

function loadSavedPosition() {
  try {
    return JSON.parse(fs.readFileSync(POSITION_PATH, 'utf8'));
  } catch {
    return null;
  }
}

function savePosition(x, y) {
  try {
    fs.writeFileSync(POSITION_PATH, JSON.stringify({ x, y }));
  } catch {
    /* best-effort */
  }
}

function createFloatButton() {
  const saved = loadSavedPosition();
  const primary = screen.getPrimaryDisplay().workArea;
  const size = 67; // 56 +20% — was reading smaller than other desktop icons
  const x = saved ? saved.x : primary.x + primary.width - size - 24;
  const y = saved ? saved.y : primary.y + Math.round(primary.height / 2);

  const win = new BrowserWindow({
    width: size,
    height: size,
    x,
    y,
    show: false,
    frame: false,
    transparent: true,
    backgroundColor: '#00000000',
    alwaysOnTop: true,
    resizable: false,
    movable: true,
    skipTaskbar: true,
    hasShadow: false, // native window shadow renders as a square around our round content — CSS box-shadow handles it instead
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.setAlwaysOnTop(true, 'screen-saver');
  win.setVisibleOnAllWorkspaces(true, { visibleOnFullScreen: true });
  // Starts click-through (see the float-btn-set-ignore-mouse handler
  // above) — the renderer takes it back out of this mode the instant the
  // cursor actually enters the visible icon.
  win.setIgnoreMouseEvents(true, { forward: true });
  win.loadFile(path.join(__dirname, 'renderer', 'floatbtn.html'));
  win.once('ready-to-show', () => win.show());

  let saveTimer = null;
  win.on('moved', () => {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      const [px, py] = win.getPosition();
      savePosition(px, py);
    }, 250);
  });

  return win;
}

// Defaults to on (most people opening this widget at all want it running
// every login, same as any tray app) but is a one-time default, not an
// enforced setting — read from the per-person profile, which the tray
// menu's checkbox toggle below writes to, so someone who explicitly turns
// it off stays off across restarts instead of being reset back to on.
function applyLoginItemSetting() {
  const saved = loadUserConfig();
  const openAtLogin = typeof saved.openAtLogin === 'boolean' ? saved.openAtLogin : true;
  if (typeof saved.openAtLogin !== 'boolean') saveUserConfig({ openAtLogin });
  app.setLoginItemSettings({ openAtLogin });
}

function toggleLoginItemSetting() {
  const openAtLogin = !app.getLoginItemSettings().openAtLogin;
  app.setLoginItemSettings({ openAtLogin });
  saveUserConfig({ openAtLogin });
  rebuildTrayMenu();
}

// Shared by the tray's own right-click menu and the floating icon's
// right-click menu (see the show-fab-menu handler below) — one definition
// so the two never drift apart.
function buildAppMenuTemplate() {
  return [
    { label: 'Submit status update', click: toggleWindow },
    { label: 'Open sheet', click: () => config.SHEET_URL && shell.openExternal(config.SHEET_URL) },
    { type: 'separator' },
    {
      label: 'Start with Windows',
      type: 'checkbox',
      checked: app.getLoginItemSettings().openAtLogin,
      click: toggleLoginItemSetting,
    },
    { type: 'separator' },
    { label: 'Quit', click: () => app.quit() },
  ];
}

function rebuildTrayMenu() {
  tray.setContextMenu(Menu.buildFromTemplate(buildAppMenuTemplate()));
}

function startMainApp() {
  if (app.dock) app.dock.hide(); // tray/floating-button app, no dock icon needed
  applyLoginItemSetting();

  // Windows always fits this to its own fixed notification-area slot
  // regardless of the source size we pass in — so a bigger source doesn't
  // risk an oversized tray icon, it just gives Windows more detail to
  // downscale from instead of upscaling a too-small 20x20 bitmap on
  // anything above 100% display scaling, which is what was making it look
  // softer/smaller than other apps' tray icons.
  const trayImg = nativeImage.createFromPath(path.join(__dirname, 'build', 'tray.png'));
  tray = new Tray(trayImg.resize({ width: 32, height: 32 }));
  tray.setToolTip('Report Generator');
  tray.on('click', toggleWindow);
  rebuildTrayMenu();

  popup = createPopup();
  floatBtn = createFloatButton();
}

app.whenReady().then(() => {
  if (configLoadFailed) return; // already showing the error dialog + quitting above

  if (!config) {
    // No config.json, no config.template.json, and nothing saved from a
    // previous Connect — this is what makes a store-distributed build
    // (no organization baked in at all) work: ask for it here instead.
    connectWin = createConnectWindow();
    return;
  }

  prefetchAll(); // warm the tabs/columns cache immediately, don't make the first click wait

  if (!config.YOUR_NAME) {
    // First run on this machine: ask for a name once, then continue.
    setupWin = createSetupWindow();
    return;
  }

  startMainApp();
});

app.on('window-all-closed', () => {
  // Keep running via the tray/floating button even if a window is hidden/closed.
});
