# Roots & Roads — System Architecture

Last updated: 2026-07-20 (post-audit)

## Components

```
             OTA emails (Viator, GYG, GuruWalk, FreeTour, Airbnb)
                               |
                   Gmail filters (label per source/type)
                               |
 Website form ---> [BookingSheet project "RootsRoadsBookings"]
 (doPost, instant) bookingList_v2.gs   (runBookingSystem 5min / audit 8h)
                   websiteAvailabilityUpdate.gs (doGet: calendar + admin hook)
                               |
                 BookingSheet (1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw)
                 English/German/Spanish Tours · Done Tours · Errors · Status
                 · Completed Log (hidden)
                               |
      +------------------------+---------------------------+
      |                                                    |
[Control project "mobBoss"]                       [Portal front end]
Roots_Roads_Control_v1                            rootsandroadsbcn.com/guide/
(1A8RrqIoWw-HpxCLDRGVJLcflja8tSGEOgaXd-2sSkLs)    guide/index.html (fetch+JSONP)
assignShifts.gs    scheduling + grid locks + edits         |
guidePortal.gs     portal API + ledger + queues + health <-+
mobileControls.gs  phone checkbox controls
   Tabs: Guides · Weekly_Schedule · Website_Schedule · Schedule ·
         Schedule_<Language> grids · Control (health A1:B14 + controls N2:P12)
         · Errors
      |
      +--> Availability (1ZkF0yDVE5Q7V2XUO5051ojJdTXtsvb8FA2sh15rqf-0) Week NN tabs
      +--> Guide_Ledger_v1 (folder 1AkSO3hS5aoUP8vZXXIBKCjQrhmavUz5j)
           Rates · per-guide tabs · Unassigned · Viator No-shows ·
           GYG No-shows · GuruWalk Check-ins
```

Timezone: **Europe/Madrid everywhere.** Both Apps Script projects must have it
set in Project Settings; `systemStatus` and `?action=health` verify it.

## Authoritative data

| Data | Source of truth |
|---|---|
| Active bookings | BookingSheet language tabs. Once a row exists, **manual edits win**: confirmation re-reads never touch it; only modification/cancellation emails, managers, or `reparseActiveRowsFromEmail` change it. A booking's TAB (language) can only be changed by management (portal "move"). |
| Completed tours | `Done Tours` (aggregate) + `Completed Log` (detail w/ ids, 21-day retention, feeds no-show queues) |
| Guides, languages, seniority, logins | Control `Guides` tab |
| Tour offer | Control `Weekly_Schedule` (real languages only — "Private" is NOT a language). Private availability slots: `ASSIGN_CFG.PRIVATE_AVAILABILITY` in assignShifts.gs |
| Operational schedule | `Schedule_<Language>` grids. **Bold = management lock** (from typing — auto-bolded by the edit trigger — or portal assignment). Private tours get their own columns (`10:00 · Private`). |
| Availability | Availability spreadsheet Week tabs (tick columns from offer + private slots) |
| Check-ins & money | Guide_Ledger_v1 per-guide tabs (Time column text-formatted; dedupe by shift AND booking id) |
| Management queues | Ledger `Unassigned`, `Viator No-shows`, `GYG No-shows`, `GuruWalk Check-ins`. Queue tabs are laid out **row 1 = CLEAR button, row 2 = headers, row 3+ = entries**; ticking CLEAR wipes the entries once you have entered them on the platform. |

## Enforced invariants (checked by `checkInvariants_` on every audit)

I1 one booking id ⇒ max one active row (all tabs) · I2 cancellation always wins
and cancelled bookings cannot return (cancellation cache guards every write
path) · I3 completed tours (start+2h) leave the active tabs · I4 rows are
complete · idempotent reprocessing everywhere (thread level, sheet level,
ledger level, queue level) · Processed only after successful parse+write ·
read/unread is never a signal · children/infants never affect paid counts or
income · no auto-assignment across unsupported languages or within 5h overlap
· locks preserved, conflicts flagged (soft: availability; hard: language/
overlap → red cell + Errors) · private tours keep their real booked time.

## Lifecycles (condensed)

**Booking**: filter → label → fast run (Gmail SEARCH `label:x -label:processed`,
~20 calls/run) → parse (subject-first classification; both body flavours
normalised — asterisk bold stripped) → adults to Guests, children/infants/
private to Notes → insert (never update) into language tab → Processed +
mark-read, **stays in inbox until tour ends** (inbox = upcoming tours).
Modifications update fields (never the tab); Viator deltas anchor to the
original confirmation; GYG keeps newest state + reinstates against stale
cancellations. Cancellation: row deleted, threads relabelled Cancel +
archived immediately, id blocked from re-insertion. Completion (start+2h):
row → Completed Log + Done Tours; threads stripped to `<Source>/Done` +
archived. Errors tab is structured + deduplicated (Count / Last seen).

**Scheduling**: rules × dates (availability columns can NEVER hide a tour) +
private shifts per booking at real times + **orphan shifts** for any booking
outside the offer → locks seated first → deterministic scoring (seniority →
carried value → tour count → availability volume → order; 5h separation) →
all Schedule_ grids regenerated (stale tabs cleared) → `validateScheduleGrids`
auto-runs → Unassigned queue → portal.

**Portal**: login (localStorage, 30-day token) → tours (cached render, then
refresh) → per-shift `eligible` guides (language + no 5h clash) → assign
(bold lock; conflicts require confirmed `force`) → move booking between
languages (traceable note; emails can't move it back) → idempotent check-in
(shift + booking-id dedupe) → guide ledger → GuruWalk queue (All/Some/None) /
no-show queues from Completed Log → 2h cutoff hides finished tours; All tours
shows current week only.

## Triggers

| Project | Function | Type | Frequency |
|---|---|---|---|
| BookingSheet | runBookingSystem | time | every 5 min |
| BookingSheet | runBookingAudit | time | every 8 h |
| Control | runWeeklyScheduling | time | Friday 18–19 |
| Control | updateManagementQueues | time | hourly |
| Control | archiveLedgerMonthly | time | monthly, 1st, 02–03 |
| Control | handleMobileControlsEdit | installable onEdit (Control sheet) | — |
| Control | handleScheduleEdit | installable onEdit (Control sheet) | — |
| Control | handleLedgerEdit | installable onEdit (Guide_Ledger_v1) | — |

## Observability

BookingSheet `Status` tab: heartbeat after every run. Control tab `A1:B14`:
manager dashboard (last runs, unassigned/pending counts, error count) —
refreshed hourly. **No recurring notification emails**: the system is silent by design. The
only mail it sends is the Friday schedule (requested), the per-booking website
alert, and an emergency "manual entry needed" note if a website reservation
arrives while the system is locked. Status is *pulled* from the dashboard, not
pushed. Deduped Errors tabs in both spreadsheets. `systemStatus` / `?action=health` for full
diagnosis. Test harness in `test/` (`bash test/run-tests.sh`).
