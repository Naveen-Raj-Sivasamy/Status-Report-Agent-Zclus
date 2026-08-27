const { app, Tray, Menu, BrowserWindow, ipcMain, screen, nativeImage, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const APP_VERSION = require('./package.json').version;

// Dev convenience: a full config.json next to main.js (gitignored) overrides
// everything, including YOUR_NAME — this is what `npm start` uses locally.
const CONFIG_PATH = path.join(__dirname, 'config.json');
// Shared connection details (WEBHOOK_URL/TOKEN/SHEET_URL) baked into the
// installer at build time — same for every teammate.
const TEMPLATE_PATH = path.join(__dirname, 'config.template.json');
// Per-person name, written on first run into each teammate's own profile
// folder — never baked into the installer, never shared between installs.
const USER_CONFIG_PATH = path.join(app.getPath('userData'), 'user-config.json');
const POSITION_PATH = path.join(app.getPath('userData'), 'float-position.json');

function loadSharedConfig() {
  if (fs.existsSync(CONFIG_PATH)) {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  }
  if (fs.existsSync(TEMPLATE_PATH)) {
    return JSON.parse(fs.readFileSync(TEMPLATE_PATH, 'utf8'));
  }
  throw new Error(
    `Missing config.json or config.template.json next to main.js.\nExpected at: ${CONFIG_PATH}`
  );
}

function loadSavedName() {
  try {
    const saved = JSON.parse(fs.readFileSync(USER_CONFIG_PATH, 'utf8'));
    return saved.YOUR_NAME || '';
  } catch {
    return '';
  }
}

function saveYourNameToDisk(name) {
  try {
    fs.writeFileSync(USER_CONFIG_PATH, JSON.stringify({ YOUR_NAME: name }));
  } catch {
    /* best-effort */
  }
}

let config;
try {
  config = loadSharedConfig();
  if (!config.YOUR_NAME) config.YOUR_NAME = loadSavedName();
} catch (err) {
  app.whenReady().then(() => {
    dialog.showErrorBox('Status Update — setup needed', err.message);
    app.quit();
  });
  config = null;
}

let tray = null;
let popup = null;
let floatBtn = null;
let setupWin = null;

// -------------------- backend calls (with retry) --------------------

async function callWithRetry(fn, attempts = 6) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      if (i < attempts) {
        // Exponential backoff (500ms, 1s, 2s, 4s, ...): fast to recover from
        // a single blip, patient enough for a slow cold start.
        const delay = Math.min(500 * 2 ** (i - 1), 4000);
        await new Promise((r) => setTimeout(r, delay));
      }
    }
  }
  throw lastErr;
}

async function apiGet(params) {
  const url = new URL(config.WEBHOOK_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  return callWithRetry(async () => {
    const resp = await fetch(url.toString(), { signal: AbortSignal.timeout(45000) });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    if (!data.ok) throw new Error(data.error || 'Unknown API error');
    return data;
  });
}

async function apiPostBody(body) {
  return callWithRetry(async () => {
    const resp = await fetch(config.WEBHOOK_URL, {
      method: 'POST',
      body: JSON.stringify(Object.assign({ token: config.TOKEN }, body)),
      signal: AbortSignal.timeout(45000),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    if (!data.ok) throw new Error(data.error || 'Unknown API error');
    return data;
  });
}

async function apiPost(tab, values) {
  return apiPostBody({ tab, values });
}

// In-memory cache the widget serves from instantly, refreshed in the
// background — the popup opening should never wait on a network round trip
// just to show the tab list it already showed a minute ago. The Apps
// Script side still caches for 5 minutes too, so a background refresh here
// is cheap even when it does need to hit the network.
let tabsCache = null;
let columnsCache = {}; // tab -> columns[]

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

function prefetchAll() {
  refreshTabsCache()
    .then((data) => Promise.all((data.tabs || []).map((t) => refreshColumnsCache(t))))
    .catch(() => {
      /* best-effort — a real request will retry when the user actually opens the form */
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
ipcMain.handle('submit-entry', async (_e, { tab, values }) => {
  const result = await apiPost(tab, values);
  refreshColumnsCache(tab); // e.g. so the next "Cleanup Number" reflects this new row right away
  return result;
});
ipcMain.handle('send-report-now', async (_e, range) =>
  apiPostBody(Object.assign({ action: 'sendReportNow' }, range || {}))
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
    return await apiGet({ action: 'version' });
  } catch {
    return null; // never block the app on a version check
  }
});
ipcMain.handle('open-external', (_e, url) => {
  if (url) shell.openExternal(url);
});
ipcMain.handle('hide-window', () => hidePopup());
ipcMain.handle('open-sheet', () => {
  if (config.SHEET_URL) shell.openExternal(config.SHEET_URL);
});
ipcMain.handle('fab-clicked', () => toggleWindow());
ipcMain.handle('show-fab-menu', () => {
  Menu.buildFromTemplate([
    { label: 'Submit status update', click: toggleWindow },
    { label: 'Open sheet', click: () => config.SHEET_URL && shell.openExternal(config.SHEET_URL) },
    { type: 'separator' },
    { label: 'Quit', click: () => app.quit() },
  ]).popup({ window: floatBtn });
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

// -------------------- first-run "what's your name" prompt --------------------

function createSetupWindow() {
  const win = new BrowserWindow({
    width: 360,
    height: 260,
    resizable: false,
    fullscreenable: false,
    minimizable: false,
    maximizable: false,
    title: 'Welcome — Status Report Generator',
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
    width: 380,
    height: 500,
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
  const size = 56;
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

function startMainApp() {
  if (app.dock) app.dock.hide(); // tray/floating-button app, no dock icon needed

  const trayImg = nativeImage.createFromPath(path.join(__dirname, 'build', 'tray.png'));
  tray = new Tray(trayImg.resize({ width: 20, height: 20 }));
  tray.setToolTip('Status Report Generator');
  tray.on('click', toggleWindow);
  tray.setContextMenu(
    Menu.buildFromTemplate([
      { label: 'Submit status update', click: toggleWindow },
      { label: 'Open sheet', click: () => config.SHEET_URL && shell.openExternal(config.SHEET_URL) },
      { type: 'separator' },
      { label: 'Quit', click: () => app.quit() },
    ])
  );

  popup = createPopup();
  floatBtn = createFloatButton();
}

app.whenReady().then(() => {
  if (!config) return;

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
