# Roots & Roads — Manager Manual (Albert & Carlos)

Updated 2026-07-20 (post-audit). Everything here matches the deployed code.

## The three places you look

1. **Guide portal → All tours** (phone): this week's tours, bookings, check-in
   state, assign/reassign dropdowns, language moves. Your day-to-day tool.
2. **Control sheet → Control tab**: health dashboard (A1:B14) + phone controls
   (N2:P12).
3. **Gmail inbox**: the live list of upcoming tours. An email stays in the
   inbox until its tour ends or is cancelled. `Processed` label = it's in the
   BookingSheet. In inbox WITHOUT `Processed` = not parsed yet (recent, or
   check Errors).

## Daily

- Nothing for normal bookings — automatic every 5 minutes.
- **The system does not email you.** Check the three queue tabs in
  Guide_Ledger_v1 whenever it suits you:
  - **GuruWalk Check-ins** — report those guests on GuruWalk (All / Some (N of
    M) / None) within **48h of tour start**, then tick the **CLEAR** checkbox
    in row 1 to empty the tab.
  - **Viator No-shows** / **GYG No-shows** — mark them as no-shows in the OTA,
    then tick **CLEAR**.
  Each tab is: row 1 = CLEAR button, row 2 = headers, row 3+ = entries.
- Status at a glance: **Control sheet → Control tab, cells A1:B14** (last
  booking run, last schedule, pending counts, error count).

## Weekly

Friday 18:00 the system refreshes week tabs, syncs availability, generates
schedules, emails you the grids. Review Schedule_<Language>; adjust by typing
(auto-locks + validates) or from the portal.

## Assigning & changing guides

- **Portal (preferred)**: All tours → dropdown per tour. Only guides who speak
  the language AND have no tour within 5h are offered. Picking one writes a
  bold LOCK. If you pick someone conflicted anyway, it asks "Assign anyway?" —
  confirming enforces it.
- **Grid**: type a name in Schedule_<Language>. It auto-bolds (= lock, survives
  regeneration) and validates instantly — a ⚠ note appears on the cell if the
  guide doesn't speak the language, is unknown/inactive, or overlaps another
  tour (<5h). Your edit is kept either way; red cells/⚠ notes are your flag.
- Sick guide: arrange the replacement on WhatsApp, then set it in the portal.

## Bookings

- **Your edits to booking rows stick.** Confirmations only insert; audits
  never overwrite an existing row. Only modification/cancellation emails (the
  guest changed it on the OTA) update fields.
- **Language move** (e.g. German guest joins English tour): portal → booking
  row → language selector → confirm. Traceable "moved from German" note;
  emails can never move it back.
- Guest-detail fixes after a phone call (count, date, name): edit the row in
  the BookingSheet directly.
- A booking at a time outside the offer still gets a tour: it appears in the
  grids/portal as its own column, flagged "Extra tour (not in Weekly_Schedule)"
  in the Schedule tab, assigned from availability or "Not assigned".

## Offer changes

English/Spanish: edit `updateWeeklyScheduleToCurrentOffer` in assignShifts.gs
and run it. German: edit Weekly_Schedule rows directly (Time column is text —
type 10:00). Private slots: `ASSIGN_CFG.PRIVATE_AVAILABILITY` ("Private" is
not a language and never goes in Weekly_Schedule). After any offer change:
Sync availability → Generate schedules (or "Weekly full run" from the phone).

## Phone controls — Control tab N2:P12

Run booking update · Run booking audit · Sync availability · Generate
schedules · Refresh ledger & queues · Weekly full run · Full operational
refresh. Tick column O; column P shows Running… → Done/Error; the box resets
itself. "Busy" = someone else's run; wait a minute. Stuck "Running…" (e.g.
after a timeout): run `clearStaleControls` (mobileControls.gs).

## Function reference (Apps Script editor)

**Automatic** (see trigger table in the Architecture doc): runBookingSystem,
runBookingAudit · runWeeklyScheduling, updateManagementQueues,
archiveLedgerMonthly, and the three installable onEdit handlers
(handleMobileControlsEdit, handleScheduleEdit, handleLedgerEdit).

**Diagnosis** — safe anytime:
- `systemStatus` (bookingList_v2.gs): quota, timezone, unprocessed mail,
  triggers, latest errors.
- `debugBooking` (bookingList_v2.gs): set `DEBUG_BOOKING_ID` first; full
  parsing story for one booking.
- `validateScheduleGrids`, `validateMobileControls` (Control project).
- `testBookingParsers` (41 checks) and `testQueueIdempotency`.
- Portal health from any browser: `<portal /exec URL>?action=health`.

**Recovery** — all idempotent, never touch unread state:

| Problem | Run |
|---|---|
| One booking wrong/missing | `debugBooking`, then `reprocessBookingById` (both use DEBUG_BOOKING_ID) |
| Many rows wrong after a parser fix | `reparseActiveRowsFromEmail` |
| Anything possibly stale | `forceProcessEverythingNow` (full audit + report) |
| Upcoming mail missing from inbox | `restoreUpcomingThreadsToInbox` |
| Ledger columns misaligned / duplicates | `repairLedgers`, `repairLedgerDuplicates` |
| Queue tabs weird row counts | `repairQueueTabs` |
| Stuck phone control | `clearStaleControls` |
| Queue tab layout wrong / stray boxes | `repairQueueTabs` (also upgrades old layouts) |
| CLEAR buttons not responding | `setupLedgerControls` (installs the ledger trigger) |
| Corrupted Weekly_Schedule times (12/30/1899) | `updateWeeklyScheduleToCurrentOffer` |

**One-time setup**: setupBookingSystem, setupLedger (also installs the queue
CLEAR buttons), setupMobileControls, setupScheduleEditValidation,
setupLedgerControls. **Do not run**: setupWeeklySchedule (obsolete
defaults), archiveLedgerMonthly mid-month.

## Gmail filters (recreate exactly if lost — never "Skip the Inbox")

Unchanged from before: Viator (`booking@t1.viator.com`; "New Booking for" /
"Cancelled Booking:" / "Amended Booking:"→Modifications), GuruWalk
(`no-reply@guruwalk.com`; "Confirmed booking" / canceled variants /
modification variants), GetYourGuide
(`do-not-reply@notification.getyourguide.com`; "Booking -"+"Urgente: nueva
reserva recibida"+"New booking received" / cancelado-cancelled variants /
"Booking detail change"+"Cambio en los datos de la reserva"), Website
(`rootsandroadstours@gmail.com`; "NEW WEBSITE RESERVATION"), Airbnb
(`inforootsandroads@gmail.com`; "booked your experience"). FreeTour: add only
after verifying a real email's sender + subject.

## Troubleshooting quick table

| Symptom | First look | Fix |
|---|---|---|
| Booking missing | `debugBooking` | reprocessBookingById |
| Wrong guest count | `debugBooking` (shows email vs sheet) | reparseActiveRowsFromEmail |
| Tour missing from schedule | Weekly_Schedule row exists? Errors tab | updateWeeklySchedule… → sync → makeSchedule |
| Impossible assignment shown | red cell / ⚠ note / Errors | reassign via portal |
| Portal error on phone | tap "Test server connection" on the error screen | health JSON tells you server vs network |
| Quota error in Errors | systemStatus | wait (heals <24h); mail is retried automatically |
| Duplicate check-ins rows | — | repairLedgerDuplicates |

Never mark mail unread; never mass-remove Processed. Every recovery above is
targeted and safe to re-run.
