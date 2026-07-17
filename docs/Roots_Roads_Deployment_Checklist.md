# Roots & Roads — Deployment Checklist

Follow in order. Steps you run once are marked (once).

## 1. BookingSheet project (Extensions → Apps Script inside BookingSheet)

1. Replace the booking script contents with `apps-script/bookingList_v2.gs`.
2. Replace/keep `websiteAvailabilityUpdate.gs` with the repo version (adds the
   admin remote-run hook).
3. Script Properties: add `ADMIN_KEY` = a long random string. Keep existing
   `BREVO_API_KEY`, `BREVO_TEMPLATE_ID`.
4. Run `setupBookingSystem` (once) → authorize.
5. Run `testBookingParsers` → expect "29 passed, 0 failed".
6. Triggers: `runBookingSystem` every 5 min · `runBookingAudit` 2-3x/day.
   (Delete any old trigger pointing at removed functions.)
7. Web app: Deploy → Manage deployments → Edit → **New version** on the
   EXISTING deployment (the website's booking form + calendar use this URL).
   Copy the /exec URL.

## 2. Control project (Extensions → Apps Script inside Roots_Roads_Control_v1)

1. Replace `assignShifts.gs` and `guidePortal.gs` with the repo versions; add
   `mobileControls.gs`.
2. Script Properties: `BOOKING_WEBAPP_URL` = the /exec URL from step 1.7;
   `ADMIN_KEY` = same string as step 1.3. Keep `LEDGER_ID`. Change
   `TOKEN_SECRET` in guidePortal.gs if still the placeholder.
3. Run `setupLedger` (once) → creates/migrates ledger tabs (adds Children
   column) + queue tabs.
4. Run `setupMobileControls` (once) → Control tab block N2:P12 + installable
   onEdit trigger.
5. Run `validateMobileControls` and `validateScheduleGrids` → fix anything
   reported.
6. Triggers: `runWeeklyScheduling` Friday 18:00-19:00 ·
   `updateManagementQueues` hourly · `sendGuruwalkCheckinReminder` daily
   16:00-17:00 · `archiveLedgerMonthly` monthly day 1, 02:00-03:00.
7. Portal web app: Deploy → Manage deployments → Edit → **New version** on the
   existing deployment (URL must stay `AKfycbxzl4...` or update the front
   end). Execute as: Me. Access: Anyone.

## 3. Website repository (GitHub Pages)

1. Commit and push the changed files: `guide/index.html`, `index.html`,
   `css/main_css.css`, `docs/*`, `apps-script/*`.
2. If the portal deployment URL changed, update `PORTAL_URL` in
   `guide/index.html` first.

## 4. Tests after deployment

- Phone browser → `<portal>/exec?action=health` → `ok: true`.
- Portal login on a phone; confirm tours load, `2+3` shows on a booking with
  children, Private pill on a private tour, check-in greys out.
- BookingSheet: `testBookingParsers` (29/29), send yourself a test website
  booking, confirm it appears.
- Control: tick "Run booking update" on the phone → status Running… → Done.
- Ledger: `testQueueIdempotency` → PASS.
- Grids: `validateScheduleGrids` → "no violations found".

## 5. Rollback

- Apps Script: Deploy → Manage deployments → Edit → select the PREVIOUS
  version. Code editor: File → See version history.
- Website: `git revert` the deploy commit and push.
- Sheets are never destructively migrated; the only schema change (ledger
  Children column) is additive and can be left in place.
