# Roots & Roads — Deployment Checklist

Follow in order. Steps marked (once) are one-time.
Verified against the code on 2026-07-20 (post-audit).

## 0. Before you start

- Both Apps Script projects must have **Project Settings → Time zone =
  Europe/Madrid**. `systemStatus` and the portal health endpoint check this
  and will warn you if it is wrong.

## 1. BookingSheet project ("RootsRoadsBookings")

1. Paste `apps-script/bookingList_v2.gs` over the existing booking file. Save.
2. Paste `apps-script/websiteAvailabilityUpdate.gs` (adds the admin remote-run
   hook used by the phone controls). Save.
3. (once) Project Settings → Script Properties:
   `ADMIN_KEY` = a long random string. Keep `BREVO_API_KEY`,
   `BREVO_TEMPLATE_ID`.
4. (once) Run `setupBookingSystem` → authorize when prompted.
5. Run `testBookingParsers` → expect **41 passed, 0 failed**.
6. Triggers (delete any trigger pointing at a function that no longer exists):
   - `runBookingSystem` — time-driven, every 5 minutes
   - `runBookingAudit` — time-driven, every 8 hours
   There is **no daily self-test trigger** — the system does not email you.
7. Deploy → Manage deployments → pencil on the ACTIVE deployment →
   Version: **New version** → Deploy. Copy the `/exec` URL.

## 2. Control project ("mobBoss", bound to Roots_Roads_Control_v1)

1. Paste `apps-script/assignShifts.gs`, `apps-script/guidePortal.gs`,
   `apps-script/mobileControls.gs`. Save.
2. (once) Script Properties: `BOOKING_WEBAPP_URL` = the `/exec` URL from 1.7;
   `ADMIN_KEY` = the same string as 1.3. Keep `LEDGER_ID`. Replace
   `TOKEN_SECRET` in guidePortal.gs if it is still the placeholder.
3. (once) Run, in this order:
   - `setupLedger` — creates/migrates ledger tabs, adds the queue tabs with
     their CLEAR buttons, installs the ledger edit trigger.
   - `setupMobileControls` — Control!N2:P12 block + edit trigger.
   - `setupScheduleEditValidation` — auto-lock + validate manual grid edits.
   - `validateMobileControls` — must report no problems.
4. Offer + schedule refresh, in this order:
   `updateWeeklyScheduleToCurrentOffer` → `syncAvailabilityFile` →
   `makeSchedule`. Then check Schedule_English has 11:00 columns and the
   Control sheet Errors tab for flagged conflicts.
5. Triggers:
   - `runWeeklyScheduling` — time-driven, Friday 18:00–19:00
   - `updateManagementQueues` — time-driven, hourly
   - `archiveLedgerMonthly` — time-driven, monthly, day 1, 02:00–03:00
   (The three onEdit triggers were created by the setup functions in 2.3 —
   do not add them by hand.)
6. Deploy → Manage deployments → pencil on the ACTIVE deployment →
   **New version** → Deploy. Confirm **Execute as: Me** and
   **Who has access: Anyone** (not "Anyone with Google account").

## 3. Website repository

1. Commit and push: `guide/index.html`, `index.html`, `css/main_css.css`,
   `apps-script/`, `docs/`, `test/`.
2. Only if the portal deployment URL changed: update `PORTAL_URL` near the top
   of `guide/index.html`'s script block first.

## 4. Verification (do these, in order)

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

## 5. Local test suite (optional, on a computer with node)

```
bash test/run-tests.sh
```
Expect: 17 + 13 + 17 + 17 Sheets-logic assertions, all passing.
Parser suite runs inside Apps Script via `testBookingParsers` (41 checks).

## 6. Rollback

- Apps Script: Deploy → Manage deployments → Edit → select the previous
  version. Code: File → See version history.
- Website: `git revert` the deploy commit and push.
- Spreadsheets: no destructive migrations. The only schema changes are
  additive (ledger `Children` column; queue tabs gained a button row above
  the headers) and both are re-runnable via `repairLedgers` / `repairQueueTabs`.
