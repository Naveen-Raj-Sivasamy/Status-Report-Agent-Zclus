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
4. `TEAMS_WEBHOOK_URL` in the CONFIG block is the one thing still edited
   directly in `Code.gs` (optional, see below) — `SHARED_SECRET` and
   `REPORT_RECIPIENTS` are NOT: this repo is public, so hardcoding a real
   token/email list here would mean anyone on the internet could read
   them. Instead:
   - In the function dropdown, select **`setupScriptProperties`**, paste
     your real secret and recipient list into the three lines inside it,
     click **Run**, then delete those pasted values again (that function's
     own comment explains why). `SHARED_SECRET` is your `TOKEN`.
   - Or skip that function entirely and set `SHARED_SECRET` /
     `REPORT_RECIPIENTS` (comma-separated) / `REMINDER_RECIPIENTS`
     directly under **Project Settings → Script Properties** — same
     effect, without ever pasting real values into a function body.
   - The API refuses every write with "Invalid token" until `SHARED_SECRET`
     is actually set this way — that's deliberate (fails closed, not open).
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

The widget lives in `desktop-widget-electron/` — a floating tray icon
with a custom popup UI. Needs [Node.js](https://nodejs.org) 18+.

(There used to be a second, plain Python/Tkinter widget in
`desktop-widget/` as a no-Node.js fallback. It was removed — it hadn't
been touched since this repo's very first commit, before essentially
every feature below existed, so it was just a trap for anyone who
happened to install the wrong, badly outdated one.)

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

## 4. Package it for teammates — one file, zero setup

Teammates get a **single `.exe`** — no Node.js, no manual install steps.
Double-clicking it installs the app silently and launches it. What happens
next depends on whether you baked connection details into the build:

- **`config.template.json` present** (below) — same as before: connects
  automatically, just asks for the person's name once.
- **No `config.template.json` at all** (the App Store / generic-distribution
  case — see the next section) — first launch instead asks the person to
  enter their own organization's `WEBHOOK_URL`/`TOKEN`/`SHEET_URL`
  themselves, and to set an admin password that protects those details
  from being casually changed again later (Settings → "Change organization
  connection" asks for it). Nothing about which organization this is gets
  compiled into the binary either way — that's what makes a single build
  distributable to more than one team.

Each person's name is asked once on first launch and saved to their own
Windows profile regardless of which path above applies — never shared
between installs, never baked into the installer.

### Baking in one organization's details (internal-only distribution)

If every teammate is on the same team as you (this repo's original use
case, not the App Store one), it's still simplest to bake the connection
in once so nobody has to type anything:

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

**Mac**: same idea — `npm run build:mac` → `dist/StatusUpdate-Setup.dmg`.
Mac doesn't support a fully silent one-click install the way Windows NSIS
does; teammates drag the app into Applications as usual, but still get
the same first-run name prompt and shared `config.template.json`. (You
don't actually need to run this by hand — see the CI/CD section below,
which builds and releases both platforms automatically on every version
bump.)

This build is unsigned (no paid code-signing certificate), so:
- **Windows**: SmartScreen may warn on first run — "More info" → "Run anyway".
- **Mac**: Gatekeeper will block it — right-click the app → "Open" the first time.

If you ever change `WEBHOOK_URL`, `TOKEN`, or `SHEET_URL`, update
`config.template.json` and rebuild — everyone will need the new `.exe`
(this is the one thing that isn't picked up automatically, since it's
baked in at build time).

### Continuous deployment (CI/CD) — everything below is automatic

Three GitHub Actions workflows do all of this for you now. You never need
to open the Apps Script editor by hand, you never need to build the
installer on your own machine, and you never need to manually tag a
release either.

**`.github/workflows/deploy-apps-script.yml`** — runs on every push to
`main` that touches `apps-script/**`. It pushes `Code.gs` /
`appsscript.json` to the live Apps Script project via
[`clasp`](https://github.com/google/clasp) and creates a new deployment
version, re-using the SAME deployment ID the widget already talks to — so
the web app URL never changes and nobody needs to update `config.json`.

**`.github/workflows/auto-tag-release.yml`** — runs on every push to `main`
that touches `desktop-widget-electron/package.json`. If the `version` field
changed, it creates and pushes the matching `v<version>` git tag itself —
that tag push is what actually triggers a release, so you no longer touch
`git tag` by hand at all.

**`.github/workflows/build-release.yml`** — runs whenever a `v*.*.*` tag is
pushed (now done automatically by the workflow above). It builds the
Windows installer AND the Mac disk image, on their respective native
runners (`windows-latest` / `macos-latest`, one after the other, not in
parallel — both upload to the same release), and publishes them as
GitHub Release assets named `StatusUpdate-Setup.exe` and
`StatusUpdate-Setup.dmg`. It also calls the Apps Script web app's
`setLatestVersion` action once (not once per platform) so the `_Config`
tab's `LatestVersion` updates itself. Neither release asset's filename
changes between releases, so these links always point at the newest
build of each and the `.exe` one is what's stored as `DownloadUrl` (set
once, permanently):

```
https://github.com/<owner>/<repo>/releases/latest/download/StatusUpdate-Setup.exe
https://github.com/<owner>/<repo>/releases/latest/download/StatusUpdate-Setup.dmg
```

The Mac build is unsigned (no Apple Developer account here), so it needs
the same right-click-to-open workaround described above.

**One-time setup** (only needed once, by whoever owns this repo):

1. **Turn on the Apps Script API** for the Google account that owns the
   script: https://script.google.com/home/usersettings → toggle it on.
2. **Authenticate clasp locally**: on your own machine, run
   `npx @google/clasp login` (this opens a browser for you to log into the
   same Google account). It writes a credentials file
   (`~/.clasprc.json` on macOS/Linux, `%USERPROFILE%\.clasprc.json` on
   Windows — clasp prints the exact path when it finishes).
3. Copy that file's entire contents and paste them into a new repo secret
   named **`CLASP_CREDENTIALS`** (GitHub repo → Settings → Secrets and
   variables → Actions → New repository secret).
4. Copy the entire contents of `desktop-widget-electron/config.template.json`
   into a second repo secret named **`CONFIG_TEMPLATE_JSON`**.
5. Set `_Config`'s `DownloadUrl` (see step 2 above) to the `releases/latest/download/...`
   link once — it never needs to change again.

**Shipping a new version from then on:**

- **Backend-only change** (Code.gs): just `git push` to `main`. Deployed
  automatically within ~1 minute. (Changing who's on `REPORT_RECIPIENTS`,
  or rotating `SHARED_SECRET`, isn't a code change at all any more — both
  live in Script Properties now, edited directly in the Apps Script editor,
  no push needed.)
- **App/UI change** (needs a new installer): bump `"version"` in
  `desktop-widget-electron/package.json` and push to `main`. That's it —
  the tag, the build, the GitHub Release, and the `LatestVersion` bump all
  happen automatically from there. Everyone still on an older version sees
  the in-app update banner within a few minutes.

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
- `compileAndSendReport` emails only **this week's rows (Mon-Fri)** from
  `REPORT_TABS`, as a filtered `.xlsx` — not the entire spreadsheet.
- Recipient list and the shared token both live in this Apps Script
  project's Script Properties, not in `Code.gs` — see step 4 above.
