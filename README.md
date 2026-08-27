# Status Report Tracker Agent

A floating desktop icon your team clicks to log a status update, which is
saved straight into a shared Google Sheet — plus a Friday-evening reminder
and an automatic end-of-week email of the completed sheet.

Everything runs on your personal Google account (Sheets + Apps Script +
Gmail). No IT/admin approval needed — this deliberately avoids the
Microsoft 365 / Azure app-registration route.

## How it fits together

```
 Desktop widget (each teammate's PC)
        │  click icon → pick tab → fill form → submit
        ▼
 Apps Script Web App  (doGet/doPost, bound to the Sheet)
        │  appends a row
        ▼
 Google Sheet  (one tab per project/category; row 1 = column headers)
        │
        ├── Fri 6pm  → sendFridayReminder()   (email, + optional Teams webhook)
        └── Fri 11pm → compileAndSendReport() (emails the whole sheet as .xlsx)
```

## 1. Set up the Google Sheet

Create (or open) the Google Sheet you want to use. Each **tab** is a
category the widget will offer as a choice (e.g. "Engineering", "Design",
"Client Calls" — whatever fits your team). **Row 1 of every tab must be
column headers** — those become the form fields in the widget.

Add a `Timestamp` column to each tab if you want to know when each entry
came in (the script fills it in automatically; leave it out of the form).

## 2. Deploy the Apps Script backend

1. In the Sheet: **Extensions → Apps Script**.
2. Delete the placeholder `Code.gs` content and paste in
   [`apps-script/Code.gs`](apps-script/Code.gs).
3. Project Settings (gear icon) → check **"Show appsscript.json manifest
   file in editor"** → open `appsscript.json` and replace it with
   [`apps-script/appsscript.json`](apps-script/appsscript.json).
4. Back in `Code.gs`, edit the CONFIG block at the top:
   - `SHARED_SECRET` — make up a random string (this is your `TOKEN`).
   - `REPORT_RECIPIENTS` — the email addresses that get the Friday report.
   - `TEAMS_WEBHOOK_URL` — optional, see below.
5. **Deploy → New deployment → Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy, and authorize the requested permissions (Sheets, Gmail,
     Drive-export) — it's your own script acting on your own Sheet.
   - Copy the `.../exec` URL it gives you — that's your `WEBHOOK_URL`.
6. In the Apps Script editor, run `setupTriggers` once (select it from the
   function dropdown → Run). This installs the two weekly triggers. You'll
   be asked to authorize again the first time — that's expected.

Whenever you edit `Code.gs` later, **redeploy** (Deploy → Manage
deployments → edit → new version) — editing the script alone doesn't
update a live Web App deployment.

### Optional: Teams reminder

In your Teams channel: **⋯ → Connectors → Incoming Webhook** → name it,
copy the URL, paste into `TEAMS_WEBHOOK_URL`. No admin approval needed on
most tenants for this (unlike a full Azure app registration).

## 3. Configure and run the desktop widget

There are two widgets in this repo — pick one:

- **`desktop-widget-electron/`** (recommended) — modern custom UI, true
  floating tray icon. Needs Node.js.
- **`desktop-widget/`** — plain Python/Tkinter version, works but looks
  basic. Simpler prerequisites if Node.js isn't an option.

### Electron version (recommended)

```
cd desktop-widget-electron
cp config.example.json config.json
```

Edit `config.json`:
- `WEBHOOK_URL` — the `.../exec` URL from step 2.5.
- `TOKEN` — must match `SHARED_SECRET` in `Code.gs`.
- `YOUR_NAME` — optional, prefills any field whose column header contains
  "name".

Run it (needs [Node.js](https://nodejs.org) 18+):

```
npm install
npm start
```

A small icon appears in the system tray (Windows) / menu bar (macOS).
Click it → a dark, rounded popup opens near the icon → pick a tab → fill
in the form → Submit. Click the icon again (or click away) to dismiss it.

### Python/Tkinter version (fallback)

```
cd desktop-widget
cp config.example.json config.json
```

Same `config.json` fields as above. Run it directly (needs Python 3.9+):

```
pip install -r requirements.txt
python status_widget.py
```

## 4. Package it for teammates

Neither PyInstaller nor electron-builder cross-compiles, so build on each OS:

- **Electron, Windows**: `cd desktop-widget-electron && npm install && npm run build:win` → `dist\StatusUpdate Setup *.exe`.
- **Electron, Mac**: `cd desktop-widget-electron && npm install && npm run build:mac` → `dist/StatusUpdate-*.dmg`.
- **Python, Windows**: run `build_windows.bat` on a Windows machine →
  `dist\StatusUpdate.exe`.
- **Python, Mac**: run `build_mac.sh` on a Mac → `dist/StatusUpdate.app`.

Give each teammate the built app **plus their own `config.json`** in the
same folder (everyone shares the same `WEBHOOK_URL`/`TOKEN` — it's the
Sheet that's shared, not the machine).

These builds are unsigned (no paid code-signing certificate), so:
- **Windows**: SmartScreen may warn — "More info" → "Run anyway".
- **Mac**: Gatekeeper will block it — right-click the app → "Open" the
  first time.

If that friction is a problem, teammates can instead just run the app
from source (`npm start` or `python status_widget.py`) if they have
Node.js or Python installed — no build step needed.

## 5. Adjust the schedule

Defaults: reminder at 6pm Friday, report at 11pm Friday, both in the
`Asia/Kolkata` timezone (set in `appsscript.json`). To change:

- Edit `.atHour(...)` in `setupTriggers()` in `Code.gs`.
- Re-run `setupTriggers()` (it clears old triggers first, so it's always
  safe to re-run).

## Notes / limitations

- The Web App is reachable by anyone who has the exact URL (Google
  doesn't index or guess these), and the `TOKEN` check in `doPost` blocks
  casual misuse — but treat the URL and token as semi-secret.
- `compileAndSendReport` emails the **entire spreadsheet** as an `.xlsx`
  attachment every Friday — it does not filter to "this week only". If
  you want a week-scoped report instead (only rows since last Monday),
  say so and the report function can be changed to build a filtered
  HTML/CSV summary instead of exporting the whole file.
- Recipient list: currently just `nav13418@fairview.org` — add more
  addresses to `REPORT_RECIPIENTS` in `Code.gs` any time.
