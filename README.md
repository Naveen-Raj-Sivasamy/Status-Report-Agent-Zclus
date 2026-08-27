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
- `YOUR_NAME` — your name, used to label entries and reports.
- `SHEET_URL` — the sheet's normal (non-API) URL, used by the "Open sheet"/"View Report" buttons.

Run it (needs [Node.js](https://nodejs.org) 18+):

```
npm install
npm start
```

A small floating icon appears on screen (and in the system tray). Click it
→ a maroon popup opens near the icon → pick a tab → fill in the form →
Submit. The popup stays open and on top of everything until you close it
with the ✕ — switching to another app to look something up won't lose
your form.

`config.json` here is for **your own dev/testing runs only** — it's
gitignored and never shipped. The installer built in the next section
does not use it at all.

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

## 4. Package it for teammates — one file, zero setup

Teammates get a **single `.exe`** — no Node.js, no `config.json` to edit,
no manual install steps. Double-clicking it installs the app silently and
launches it; the first time it runs, it just asks for the person's name,
then remembers it.

This works because the app now reads connection details
(`WEBHOOK_URL`/`TOKEN`/`SHEET_URL`) from `config.template.json`, which gets
baked into the installer at build time and is the same for every
teammate. Each person's name is asked once on first launch and saved to
their own Windows profile — never shared between installs, never baked
into the installer.

**One-time setup before building** (you only do this once, not per teammate):

```
cd desktop-widget-electron
cp config.json config.template.json   # if you haven't already made config.json
```

Then open `config.template.json` and **delete the `YOUR_NAME` line** —
it should only contain `WEBHOOK_URL`, `TOKEN`, and `SHEET_URL`. This file
is gitignored (like `config.json`), since it holds the shared token, but
it does get bundled into the `.exe` you build next.

**Build it** — must run natively on Windows (not through any bridge/VM),
since electron-builder needs the real OS to produce a proper signed-icon
NSIS installer:

```
cd desktop-widget-electron
npm install
npm run build:win
```

First run downloads Electron + builder tooling (a few hundred MB), so it
can take a few minutes depending on your connection. When it finishes,
you'll have:

```
dist\StatusUpdate-Setup.exe
```

That's the one file to hand to teammates (Slack/Teams/email/shared
drive — however you'd normally share a file). They just double-click it:
it installs to their own user profile (no admin rights needed), adds a
Start Menu + Desktop shortcut, launches automatically, asks for their
name once, and they're in.

**Mac**: same idea — `npm run build:mac` → `dist/StatusUpdate-*.dmg`.
Mac doesn't support a fully silent one-click install the way Windows NSIS
does; teammates drag the app into Applications as usual, but still get
the same first-run name prompt and shared `config.template.json`.

This build is unsigned (no paid code-signing certificate), so:
- **Windows**: SmartScreen may warn on first run — "More info" → "Run anyway".
- **Mac**: Gatekeeper will block it — right-click the app → "Open" the first time.

If you ever change `WEBHOOK_URL`, `TOKEN`, or `SHEET_URL`, update
`config.template.json` and rebuild — everyone will need the new `.exe`
(this is the one thing that isn't picked up automatically, since it's
baked in at build time).

### Letting teammates know a new version exists

The app shows its version number (e.g. `v1.1.0`) next to the title, and
checks once per launch whether a newer one exists. When you cut a new
`.exe` for a code/UI change (not just a data change — those apply on their
own within a few minutes, see the note at the top of `Code.gs`):

1. Bump `"version"` in `desktop-widget-electron/package.json`.
2. Rebuild (`npm run build:win`) and put the new `.exe` somewhere
   teammates can grab it from (shared drive, Slack, wherever).
3. Open the Google Sheet, unhide and open the `_Config` tab (created
   automatically the first time anyone used the app — right-click any tab
   → "Unhide sheet" if you don't see it), and set:
   - `LatestVersion` → the version you just built (e.g. `1.1.0`)
   - `DownloadUrl` → a link to the new `.exe`

That's it — no redeploy needed, since this reads live sheet data. Within
a few minutes, everyone still on the old version sees a banner in the app
telling them to update, with a button that opens `DownloadUrl` directly.

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
