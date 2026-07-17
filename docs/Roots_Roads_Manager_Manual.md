# Roots & Roads — Manager Manual (Albert & Carlos)

## Daily workflow

1. Nothing to do for normal bookings — emails are parsed every 5 minutes into
   the BookingSheet, and the portal follows automatically.
2. Glance at the portal or BookingSheet for today's tours.
3. Afternoon: read the daily queue email ("GuruWalk check-ins / no-shows to
   process") and act on it:
   - **GuruWalk Check-ins** tab (Guide_Ledger_v1): open the GuruWalk platform,
     mark the listed guests as attended, then tick "Reported in GuruWalk".
     You have **48 h from tour start** — the deadline column shows it.
   - **Viator No-shows / GYG No-shows** tabs: mark those bookings as no-shows
     inside Viator / GetYourGuide (this is the required management action in
     the OTA platform), then tick "OTA action done".
4. Check **Unassigned** (Guide_Ledger_v1): upcoming bookings whose shift has
   no guide. Assign someone in the schedule grid (or wait for Friday's run);
   the entry disappears automatically once assigned.

## Weekly scheduling workflow

- Guides tick availability in Roots_Roads_Guide_Availability_v1 ("Week NN"
  tabs) during the week.
- **Friday evening the trigger runs `runWeeklyScheduling`**: refreshes week
  tabs, syncs availability, generates schedules, and emails you the
  Schedule_<Language> tables ready to forward.
- Review the grids (Schedule_English / Schedule_German / Schedule_Spanish).
  Hand-edit any cell — the portal reads the grids directly.

### Locking an assignment (bold)

Write (or keep) a guide's name **in bold** in a Schedule_<Language> cell =
management lock. `makeSchedule` will never move or overwrite it and assigns
everyone else around it. Generated names are normal weight.

If your lock is impossible (guide doesn't speak the language, isn't
available, or overlaps another tour < 5 h apart), the run KEEPS your lock but
tints the cell red and writes the reason to the Control sheet's `Errors` tab.
Fix it or ignore it deliberately — nothing is silently changed.
`Albert` can never be auto-assigned to German (Guides tab languages rule);
only a manual bold lock could put him there, and it would be flagged.

### Private tours & children

- Private bookings (Viator private products, GYG "Tour privado"/"Private
  tour") appear with 🔒 in the grids, "Private" pills in the portal, and run
  at their REAL booked time (a 10:00 private stays at 10:00).
- Children show in BookingSheet Notes ("3 children") and in the portal as
  `2+3` (adults + smaller child count). They never count toward paying
  headcount, income, or guide pay.

## Phone controls (Google Sheets app)

Open Roots_Roads_Control_v1 → **Control** tab → block at **N2:P12**:

| Action | What it runs |
|---|---|
| Run booking update | runBookingSystem (remote) |
| Run booking audit | runBookingAudit (remote) |
| Sync availability | syncAvailabilityFile |
| Generate schedules | makeSchedule |
| Refresh ledger & queues | updateManagementQueues |
| Weekly full run | runWeeklyScheduling |
| Full operational refresh | booking update + sync + schedule + queues |

Tick the checkbox in column O. Column P shows `Running…`, then
`Done: 2026-07-17 18:04` or `Error: <reason>`. The checkbox resets itself.
If it says `Busy`, someone else's run is in progress — wait a minute.

One-time setup: run `setupMobileControls()` from the Apps Script editor
(creates block + installable trigger), and set Script Properties
`BOOKING_WEBAPP_URL` (BookingSheet web-app /exec URL) and `ADMIN_KEY`
(same random string in BOTH the Control and BookingSheet projects).
`validateMobileControls()` reports anything missing.

## Gmail: filters and labels

Filters (already configured; recreate exactly if lost — all "Never send to
Spam", never "Skip the Inbox"):

- Viator: `from:(booking@t1.viator.com)` + subject `"New Booking for"` →
  Viator/Confirmations · `"Cancelled Booking:"` → Viator/Cancellations ·
  `"Amended Booking:"` → Viator/Modifications (there is NO Amended label).
- GuruWalk: `from:(no-reply@guruwalk.com)` + `"Confirmed booking"` →
  Confirmations · `{"has canceled booking" "has canceled a booking"
  "has cancelled booking" "has cancelled a booking"}` → Cancellations ·
  `{"you have a modification on booking" "booking modification"}` →
  Modifications.
- GetYourGuide: `from:(do-not-reply@notification.getyourguide.com)` +
  `{"Booking -" "Urgente: nueva reserva recibida" "New booking received"}` →
  Confirmations · `{cancelado cancelada cancelled canceled}` → Cancellations ·
  `{"Booking detail change" "Cambio en los datos de la reserva"}` →
  Modifications.
- Website: `from:(rootsandroadstours@gmail.com)` +
  `"NEW WEBSITE RESERVATION"` → Webpage/Confirmations.
- Airbnb: `from:(inforootsandroads@gmail.com)` + `"booked your experience"` →
  Airbnb/Confirmations. FreeTour filters: to be added when the first real
  freetour.com email arrives (verify sender + subject first; do not guess).

Label meanings:
- `Publishing Pages/Processed` = the algorithm has fully handled that thread.
  A confirmation WITHOUT Processed is not in the BookingSheet yet (or failed
  to parse — check Errors).
- `<Source>/Done` = tour completed; all working labels removed. The live
  Confirmations/Modifications labels therefore contain only upcoming tours —
  anything old sitting there deserves a look.

## Functions reference

Automatic (time triggers): runBookingSystem, runBookingAudit,
runWeeklyScheduling, updateManagementQueues, sendGuruwalkCheckinReminder,
archiveLedgerMonthly.

Safe to run manually any time (all idempotent): everything above plus
makeSchedule, syncAvailabilityFile, setupBookingSystem, setupLedger,
validateScheduleGrids, validateMobileControls, testBookingParsers,
testQueueIdempotency, debugWhereIsBooking (edit the id inside),
debugGygParsingOnly, debugFreetourParsingOnly.

Do NOT run: archiveLedgerMonthly mid-month (it's guarded, but don't),
setupWeeklySchedule (overwrites the tour offer with defaults).

## Portal deployment & health

After changing guidePortal.gs: Apps Script → Deploy → **Manage deployments →
Edit → New version** (keep the SAME deployment so the URL doesn't change).
If you create a NEW deployment, update `PORTAL_URL` in `guide/index.html` and
push to GitHub.

Health check from any phone browser:
`<portal /exec URL>?action=health` → shows server time, deployment id, and
whether the Control sheet / BookingSheet / ledger are reachable. `ok: true`
means the backend is fine — a portal problem is then front-end or network.

## Gmail quota — why it used to break, and how it's prevented now

The old script listed every thread under every label and inspected each one on
every 5-minute run — hundreds of Gmail calls per run, which hit Google's
"Service invoked too many times for one day: gmail" limit by the afternoon.

The current fast run uses **the Processed label as a Gmail search filter**: it
asks Gmail directly for "mail under this label that is NOT yet Processed", so
an idle run costs about 20 Gmail calls instead of 400. It stays comfortably
under quota at 5-minute frequency — you do NOT need to slow it down.

The twice-daily **audit** deliberately re-reads everything (Processed
included) to repair drift; that's the heavier run, and twice a day is fine.

If you ever see the quota error again in the BookingSheet **Errors** tab:
it self-heals within ~24h (Google's window is rolling), and nothing is lost
because mail without the Processed label is retried automatically. To watch
recovery, open the **Status** tab (BookingSheet) — it refreshes after every
run.

## Repair & audit functions (run from the editor when needed)

- `systemStatus` (bookingList_v2.gs) — full diagnosis in the log: Gmail quota
  state, email quota, unprocessed mail per label, triggers, latest errors.
- `forceProcessEverythingNow` (bookingList_v2.gs) — force a full re-read and
  repair now; safe, idempotent.
- `testInternalAlertEmail` (bookingList_v2.gs) — confirms the script can email.
- `repairLedgers` (guidePortal.gs) — fixes guide-tab columns if headers look
  misaligned (e.g. "R&R makes" missing / Children column).
- `repairQueueTabs` (guidePortal.gs) — clears stray checkboxes that could
  inflate a queue tab's row count.
- `repairLedgers` + `repairQueueTabs` also run automatically inside
  `setupLedger`.

## Troubleshooting

- **Booking missing**: is the email under the right label? No label → fix the
  Gmail filter. Labelled but no Processed → check BookingSheet Errors tab;
  run "Run booking update"; the audit retries failed parses automatically.
  Use debugWhereIsBooking with the booking id to locate the thread.
- **Duplicate booking**: run "Run booking audit" (dedupe pass). If it
  persists, the two rows differ in booking id — delete the wrong one.
- **Stale schedule in portal**: the portal reads the grids live; refresh the
  page. If a grid cell is wrong, edit it (bold if you want it locked).
- **Portal "network error"**: open the health URL. `ok:true` → front-end;
  hard-refresh, check PORTAL_URL. Unreachable → redeploy the web app
  (Anyone, execute as you).
- **Check-in failed**: guide should retry (saves are idempotent — a repeated
  save overwrites, never duplicates). Managers can check in on any guide's
  behalf from All tours.
- **Recovery**: never mark mail unread. Run "Run booking audit" — it re-reads
  every labelled thread (Processed included) and repairs sheet state.
