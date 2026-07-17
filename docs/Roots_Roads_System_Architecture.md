# Roots & Roads — System Architecture

Last updated: 2026-07-17

## Components

```
                 OTA emails (Viator, GYG, GuruWalk, FreeTour, Airbnb)
                                   |
                       Gmail filters (label per source/type)
                                   |
   Website booking form ----> [BookingSheet project]
   (doPost, instant)          bookingList_v2.gs  (runBookingSystem every 5 min,
                              runBookingAudit 2-3x/day)
                              websiteAvailabilityUpdate.gs (doGet: website
                              calendar + admin remote-run hook)
                                   |
                     BookingSheet spreadsheet
                     (1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw)
                     English/German/Spanish Tours · Done Tours · Errors
                     · Completed Log (hidden handoff)
                                   |
        +--------------------------+---------------------------+
        |                                                      |
[Control project]                                     [Guide portal front end]
Roots_Roads_Control_v1                                rootsandroadsbcn.com/guide/
(1A8RrqIoWw-HpxCLDRGVJLcflja8tSGEOgaXd-2sSkLs)        (guide/index.html, JSONP)
assignShifts.gs   — availability + scheduling                  |
guidePortal.gs    — portal API (doGet), ledger, queues  <------+
mobileControls.gs — phone checkbox controls
        |
        +--> Availability spreadsheet (1ZkF0yDVE5Q7V2XUO5051ojJdTXtsvb8FA2sh15rqf-0)
        |    "Week NN" checkbox tabs guides tick
        +--> Guide_Ledger_v1 (auto-created, folder 1AkSO3hS5aoUP8vZXXIBKCjQrhmavUz5j)
             Rates · per-guide tabs · Unassigned
             · Viator No-shows · GYG No-shows · GuruWalk Check-ins
```

## Authoritative data

| Data | Source of truth |
|---|---|
| Active bookings | BookingSheet `English/German/Spanish Tours` tabs |
| Completed tours (aggregate) | BookingSheet `Done Tours` |
| Completed bookings (detail, w/ ids) | BookingSheet `Completed Log` (hidden, 21-day retention) |
| Guides, languages, seniority, portal logins | Control `Guides` tab |
| Weekly tour offer | Control `Weekly_Schedule` (+ `Website_Schedule` extras) |
| Assignments shown everywhere | Control `Schedule_<Language>` grids (bold = manager lock) |
| Guide availability | Availability spreadsheet `Week NN` tabs |
| Check-ins & money per guide | Guide_Ledger_v1 per-guide tabs |
| Management queues | Guide_Ledger_v1 `Unassigned`, `Viator No-shows`, `GYG No-shows`, `GuruWalk Check-ins` |

## Booking lifecycle

1. **Confirmation** email arrives → Gmail filter labels it (e.g. `Publishing
   Pages/GetYourGuide/Confirmations`). `runBookingSystem` (every 5 min) parses
   it: adults → `Number of Guests`, children/infants → Notes ("3 children"),
   private keyword → Notes "Private", income per source model. Row upserted
   into the language tab. Thread gets `Publishing Pages/Processed` **only after
   a successful parse+write** and is archived. A thread that parses to nothing
   is logged to Errors and left un-Processed so the audit retries it.
2. **Modification** → same label logic per source. Viator "Amended Booking:"
   emails (delta style: traveler added/removed) are summed per booking id and
   applied against the ORIGINAL confirmation guest count (idempotent). GYG
   reschedules keep only the newest state per booking code and reinstate the
   code against stale cancellations. Guruwalk/FreeTour carry the new state
   directly. Modification threads also get Processed on success.
3. **Cancellation always wins.** The cancellation pass removes the sheet row;
   `reconcileCancelledThreadLabels_` strips Confirm/Modify labels from that
   booking's threads, applies the Cancel label, marks Processed, archives.
   `upsertActiveBooking_` refuses to re-add any booking present in the
   cancellation cache, so an old confirmation can never resurrect it.
4. **Completion** (start + 2h): row is written in full to `Completed Log`,
   aggregated into `Done Tours`, deleted from the active tab. The Gmail
   threads lose all working labels and keep only `<Source>/Done` — the live
   Confirmations/Modifications labels show only upcoming operational mail.
5. Read/unread status is NEVER a signal anywhere.

## Scheduling lifecycle

`runWeeklyScheduling` (Friday evening trigger): `ensureWeekTabs_` →
`syncAvailabilityFile` → `makeSchedule` → `emailWeeklySchedule` (HTML tables of
each `Schedule_<Language>` grid to rootsandroadstours@gmail.com).

`makeSchedule`:
1. Reads guides (languages/seniority), Weekly_Schedule offer, availability.
2. Reads **manager locks**: bold names in the current `Schedule_<Language>`
   grids. Locks are seated first and never overwritten; conflicts (wrong
   language, unavailable, overlapping, unknown name) keep the lock, tint the
   cell red, and log to the Control `Errors` tab.
3. Auto-assigns remaining seats deterministically (see `ASSIGN_CFG` in
   assignShifts.gs): seniority → carried value → tour count → availability
   volume → sheet order. Two tours per guide must start ≥ 5 h apart.
4. Private bookings (Notes contains "Private") become their own shifts at
   their REAL booked time, never folded into a regular slot; several private
   groups at one time stay distinct.
5. Writes the `Schedule` detail tab + `Schedule_<Language>` grids (locked
   names re-written in bold so locks survive regeneration).

## Portal & check-in lifecycle

Portal front end calls the Control web app (JSONP):
`login` → token · `tours` → my tours + all-tours/schedule + rates ·
`save` → check-ins · `health` → deployment sanity.

Check-in writes replace that shift's rows in the guide's ledger tab
(idempotent; LockService guarded). Booking rows show adults and children as
`2+3` (children never counted or paid). After the tour completes, any paid
booking with no check-in row lands once in `Viator No-shows` / `GYG No-shows`;
every GuruWalk check-in lands once in `GuruWalk Check-ins` with a 48 h
deadline. `updateManagementQueues` (hourly) + `sendGuruwalkCheckinReminder`
(daily afternoon) keep managers on top of both. `archiveLedgerMonthly`
(1st of month) copies the ledger to `YYYY_MM_Guide_Ledger_v1` and clears the
live one.

## Trigger schedule (all Europe/Madrid)

| Function | Project | Type | Frequency |
|---|---|---|---|
| runBookingSystem | BookingSheet | time | every 5 min |
| runBookingAudit | BookingSheet | time | 2-3x/day |
| runWeeklyScheduling | Control | time | Friday 18:00-19:00 |
| updateManagementQueues | Control | time | hourly |
| sendGuruwalkCheckinReminder | Control | time | daily 16:00-17:00 |
| archiveLedgerMonthly | Control | time | monthly, 1st, 02:00-03:00 |
| handleMobileControlsEdit | Control | installable onEdit | on checkbox tick |

## Failure handling

Booking-side errors log to BookingSheet `Errors`; schedule conflicts to
Control `Errors`. Per-thread Gmail failures are contained (safe wrappers);
a failed parse is retried by the audit because Processed is only applied on
success. All caches are per-execution; every write path is idempotent, so
re-running anything is safe.
