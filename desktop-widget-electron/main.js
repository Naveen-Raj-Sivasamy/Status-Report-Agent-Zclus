const { app, Tray, Menu, BrowserWindow, ipcMain, screen, nativeImage } = require('electron');
const path = require('path');
const fs = require('fs');

const CONFIG_PATH = path.join(__dirname, 'config.json');

function loadConfig() {
  if (!fs.existsSync(CONFIG_PATH)) {
    throw new Error(
      `Missing config.json next to main.js.\nCopy config.example.json to config.json and fill in WEBHOOK_URL / TOKEN.\nExpected at: ${CONFIG_PATH}`
    );
  }
  return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
}

let config;
try {
  config = loadConfig();
} catch (err) {
  // Show the error once a window/dialog is possible.
  app.whenReady().then(() => {
    const { dialog } = require('electron');
    dialog.showErrorBox('Status Update — setup needed', err.message);
    app.quit();
  });
  config = null;
}

let tray = null;
let popup = null;

// -------------------- backend calls (with retry) --------------------

async function callWithRetry(fn, attempts = 3, delayMs = 2000) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      if (i < attempts) await new Promise((r) => setTimeout(r, delayMs));
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

async function apiPost(tab, values) {
  return callWithRetry(async () => {
    const resp = await fetch(config.WEBHOOK_URL, {
      method: 'POST',
      body: JSON.stringify({ token: config.TOKEN, tab, values }),
      signal: AbortSignal.timeout(45000),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    if (!data.ok) throw new Error(data.error || 'Unknown API error');
    return data;
  });
}

ipcMain.handle('get-tabs', async () => apiGet({ action: 'tabs' }));
ipcMain.handle('get-columns', async (_e, tab) => apiGet({ action: 'columns', tab }));
ipcMain.handle('submit-entry', async (_e, { tab, values }) => apiPost(tab, values));
ipcMain.handle('get-your-name', () => config.YOUR_NAME || '');
ipcMain.handle('hide-window', () => popup && popup.hide());

// -------------------- tray + popup window --------------------

function createPopup() {
  const win = new BrowserWindow({
    width: 380,
    height: 480,
    show: false,
    frame: false,
    resizable: false,
    fullscreenable: false,
    skipTaskbar: true,
    backgroundColor: '#00000000',
    transparent: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.on('blur', () => {
    if (!win.webContents.isDevToolsOpened()) win.hide();
  });
  return win;
}

function positionPopupNearTray() {
  const trayBounds = tray.getBounds();
  const winBounds = popup.getBounds();
  const display = screen.getDisplayNearestPoint({ x: trayBounds.x, y: trayBounds.y });
  const workArea = display.workArea;

  let x = Math.round(trayBounds.x + trayBounds.width / 2 - winBounds.width / 2);
  let y;
  // Tray at top of screen (Windows top? usually bottom; macOS top) — decide by tray y position.
  if (trayBounds.y < workArea.y + workArea.height / 2) {
    y = Math.round(trayBounds.y + trayBounds.height + 8); // menu bar at top (macOS)
  } else {
    y = Math.round(trayBounds.y - winBounds.height - 8); // taskbar at bottom (Windows)
  }
  x = Math.min(Math.max(x, workArea.x + 8), workArea.x + workArea.width - winBounds.width - 8);
  popup.setPosition(x, y, false);
}

function toggleWindow() {
  if (popup.isVisible()) {
    popup.hide();
    return;
  }
  positionPopupNearTray();
  popup.show();
  popup.focus();
  popup.webContents.send('opened');
}

app.whenReady().then(() => {
  if (!config) return; // error dialog already shown

  app.dock && app.dock.hide(); // macOS: this is a tray-only app, no dock icon

  const trayIcon = nativeImage.createFromPath(path.join(__dirname, 'build', 'tray.png'));
  tray = new Tray(trayIcon.resize({ width: 20, height: 20 }));
  tray.setToolTip('Status Update');

  popup = createPopup();

  tray.on('click', toggleWindow);
  tray.setContextMenu(
    Menu.buildFromTemplate([
      { label: 'Submit status update', click: toggleWindow },
      { type: 'separator' },
      { label: 'Quit', click: () => app.quit() },
    ])
  );
});

app.on('window-all-closed', (e) => {
  // Keep running in the tray even if the popup is closed.
  if (process.platform !== 'darwin') {
    // no-op: we don't want the app to quit just because the (hidden) window closed
  }
});
