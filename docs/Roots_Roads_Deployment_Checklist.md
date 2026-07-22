# Roots & Roads — Deployment Checklist

Follow in order. Steps marked (once) are one-time.
Verified against the code on 2026-07-21 (clasp migration).

Code now deploys from source with [clasp](https://github.com/google/clasp).
No more copy-paste into the Apps Script editor. The two projects live in:

- `apps-script/booking/` — **BookingSheet** project ("RootsRoadsBookings"):
  `bookingList_v2.gs`, `websiteAvailabilityUpdate.gs`
- `apps-script/control/` — **Control** project ("mobBoss", bound to
  Roots_Roads_Control_v1): `guidePortal.gs`, `assignShifts.gs`,
  `mobileControls.gs`

Each folder has its own `.clasp.json` (which project to push to) and
`appsscript.json` (manifest — time zone `Europe/Madrid`, web-app access).

## 0. One-time machine setup (once, per computer)

1. Install [Node.js](https://nodejs.org) 18+ (includes `npm`) and make sure
   `node` and `npm` work in a terminal. `bash` is also required to run the
   test suite (Git Bash on Windows, or WSL/macOS/Linux).
2. From the repo root: `npm install` (installs clasp locally into
   `node_modules`; all `npm run …` commands below use it automatically).
3. `npm run login` — runs `clasp login`. **This opens a browser and asks for
   your Google account.** Do this yourself; it stores a token in
   `~/.clasprc.json` (git-ignored, never commit it).
4. Fill in the four IDs (they are placeholders in the repo until you do):
   - **Script IDs** → `apps-script/booking/.clasp.json` and
     `apps-script/control/.clasp.json`. Get each from the Apps Script editor:
     **Project Settings → IDs → Script ID**. Replace `PASTE_…_SCRIPT_ID_HERE`.
   - **Deployment IDs** → `package.json`, in the `deploy:booking` and
     `deploy:control` scripts. Get them by `cd`-ing into the folder and running
     `npx clasp deployments` — copy the ID of the **active web-app deployment**
     (the one whose `/exec` URL is referenced by `guide/index.html` and the
     `BOOKING_WEBAPP_URL` script property). Replace `PASTE_…_DEPLOYMENT_ID_HERE`.

   > Why deployment IDs matter: `deploy:*` runs `clasp deploy -i <id>`, which
   > publishes a **new version of the existing deployment** so the `/exec` URL
   > never changes. Deploying **without** `-i` would mint a brand-new URL and
   > silently break the portal and the booking hook.
5. (once, recommended) Run `npx clasp pull` inside each folder and diff the
   result against the committed files. This confirms the `scriptId` is right
   and lets you reconcile the `appsscript.json` manifest (time zone, web-app
   access, and OAuth scopes) against what each live project currently has.
   The manifests here omit `oauthScopes` on purpose so Apps Script
   auto-detects them — the first deploy may prompt you to re-authorize.

## 1. Push / deploy the code

All commands run from the **repo root**. Every `push`/`deploy` first runs the
gate (`npm test` → reference check + Sheets-logic suite) and **aborts if it
fails** — a broken build never reaches Apps Script.

- Push code only (updates the editor, does **not** publish a new version):
  - `npm run push:booking`
  - `npm run push:control`
- Push **and** publish a new version of the existing web-app deployment:
  - `npm run deploy:booking`
  - `npm run deploy:control`

Use `push:*` while iterating and running editor-side setup functions; use
`deploy:*` when you want the live `/exec` URL to serve the new code.

## 2. BookingSheet project — first-time editor setup (once)

After `npm run push:booking`, in the Apps Script editor / Project Settings:

1. (once) Project Settings → Script Properties:
   `ADMIN_KEY` = a long random string. Keep `BREVO_API_KEY`,
   `BREVO_TEMPLATE_ID`.
2. (once) Run `setupBookingSystem` → authorize when prompted. This creates the
   per-language tabs — **English / German / Spanish / Italian / French Tours** —
   plus Done Tours / Errors / Status (idempotent: re-running never duplicates).
3. Run `testBookingParsers` → expect **0 failed** (includes the Italian/French
   routing + parser checks).
4. Triggers (delete any trigger pointing at a function that no longer exists):
   - `runBookingSystem` — time-driven, every 5 minutes
   - `runBookingAudit` — time-driven, every 8 hours
   There is **no daily self-test trigger** — the system does not email you.
5. `npm run deploy:booking` to publish. The `/exec` URL is unchanged; if this
   is the very first clasp deploy, copy the `/exec` URL for step 3 below.

## 3. Control project — first-time editor setup (once)

After `npm run push:control`, in the editor / Project Settings:

1. (once) Script Properties: `BOOKING_WEBAPP_URL` = the BookingSheet `/exec`
   URL; `ADMIN_KEY` = the same string as 2.1. Keep `LEDGER_ID`. Replace
   `TOKEN_SECRET` in guidePortal.gs if it is still the placeholder.
2. (once) Run, in this order:
   - `setupLedger` — creates/migrates ledger tabs, adds the queue tabs with
     their CLEAR buttons, installs the ledger edit trigger.
   - `setupMobileControls` — Control!A1 functions block (top-left, run buttons)
     with the SYSTEM HEALTH block written below it, + edit trigger.
   - `setupScheduleEditValidation` — auto-lock + validate manual grid edits.
   - `validateMobileControls` — must report no problems.
   - Guides tab: language columns (English, German, Spanish, **Italian**,
     **French**, …) are read **by header name**, so adding a language column
     needs no code change — a guide with TRUE in the Italian/French column is
     recognised and offered for those tours automatically.
3. Offer + schedule refresh, in this order:
   `updateWeeklyScheduleToCurrentOffer` → `syncAvailabilityFile` →
   `makeSchedule`. Then check Schedule_English has 11:00 columns and the
   Control sheet Errors tab for flagged conflicts.
4. Triggers:
   - `runWeeklyScheduling` — time-driven, Friday 18:00–19:00
   - `updateManagementQueues` — time-driven, hourly
   - `archiveLedgerMonthly` — time-driven, monthly, day 1, 02:00–03:00
   (The three onEdit triggers were created by the setup functions in 3.2 —
   do not add them by hand.)
5. `npm run deploy:control` to publish. The manifest sets **Execute as: Me**
   and **Who has access: Anyone** (anonymous). Confirm this on the deployment
   if it is the first clasp deploy.

## 4. Website repository

1. Commit and push: `guide/index.html`, `index.html`, `css/main_css.css`,
   `apps-script/`, `docs/`, `test/`, `package.json`, `.github/`.
2. Only if a portal deployment URL changed (it should **not**, since deploys
   reuse the existing deployment ID): update `PORTAL_URL` near the top of
   `guide/index.html`'s script block, and `BOOKING_WEBAPP_URL` in the Control
   script properties.

## 5. Verification (do these, in order)

- Phone browser: `<portal /exec URL>?action=health` → `"ok":true`,
  `"timezoneOk":true`, `deployment":"portal-v4"`.
- Portal on a phone: log in; tours load; a booking with children shows `2+3`;
  a private tour shows the Private pill; tapping **Check in** greys the button;
  minimise and reopen → tours still there, check-in still recorded.
- Manager view: All tours shows only this week; the assign dropdown offers
  only compatible guides; assigning a conflicted guide asks "Assign anyway?".
- Control sheet → Control tab: health block A1:B14 populated; tick
  "Run booking update" (O3) → Running… → Done.
- Guide_Ledger_v1: each queue tab has row 1 CLEAR button, row 2 headers.
  Tick CLEAR on an empty tab → status says "Cleared 0 entries".
- BookingSheet → Status tab: "Last run finished" within the last 5 minutes.
- Editor: `testQueueIdempotency` → PASS; `validateScheduleGrids` → no
  violations.

## 6. Local test gate (runs automatically before every push/deploy)

`npm run push:*` and `npm run deploy:*` run this first and block on failure.
To run it by hand:

```
npm test
```

This runs `node test/check-references.js` (every internal call resolves, no
duplicate definitions — the class of bug that shipped a ReferenceError to
production) followed by `bash test/run-tests.sh` (Sheets-logic assertions).
Expect: reference check PASS for both projects, then the Sheets-logic suites
all passing. The parser suite runs inside Apps Script via `testBookingParsers`
(41 checks). The same two checks run in CI on every push — see
`.github/workflows/checks.yml`.

## 7. Rollback

- Apps Script code: `git revert` the offending commit, then
  `npm run deploy:booking` / `npm run deploy:control` to republish the previous
  code to the same deployment. Or, in the editor: Deploy → Manage deployments →
  Edit → select a previous version; File → See version history for code.
- Website: `git revert` the deploy commit and push.
- Spreadsheets: no destructive migrations. The only schema changes are
  additive (ledger `Children` column; queue tabs gained a button row above
  the headers) and both are re-runnable via `repairLedgers` / `repairQueueTabs`.
