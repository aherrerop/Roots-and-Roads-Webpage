# Guide Portal — Staging copy (test the second-account setup without touching live)

Goal: run a **fully independent copy** of the guide portal — its own sheets, its own web-app
URL, its own login page — so we can load-test and tweak it while the **live portal keeps
serving guides untouched**. This is also the dry run for moving the portal to the second
Google account (`inforootsandroads@gmail.com`).

## Why it's safe (data isolation)

- The **Control** script is *container-bound* to the Control spreadsheet (`control_()` =
  `getActiveSpreadsheet()`). So **copying the Control spreadsheet copies the whole script with
  it** — that copy IS the staging portal code, owned by whoever made the copy.
- The **external** sheets are referenced by id: the BookingSheet (holds the `Portal Feed`,
  bookings, the `near-future Portal feed`), the Ledger, and the two availability files. The
  portal **reads and writes** the BookingSheet (check-ins, assignments) and the Ledger, so
  staging MUST point these at **copies** or it would corrupt live data.
- The code now reads those ids from Script Properties, falling back to the live constants
  (`extSheetId_` / `bookingSheetId_` / `guideFileId_` / `sfGuideFileId_`). **Set a property →
  staging uses the copy. Set nothing → prod uses the live sheet.** No code fork.

## The staging config = four Script Properties

On the **staging** Control project only (Project Settings → Script Properties):

| Property | Set to | If unset |
|---|---|---|
| `BOOKING_SHEET_ID`   | id of the **copy** BookingSheet | live BookingSheet (unsafe for staging) |
| `GUIDE_FILE_ID`      | id of the **copy** Guide_3h_Availability | live |
| `SF_GUIDE_FILE_ID`   | id of the **copy** Guide_SF_Availability | live |
| `LEDGER_ID`          | id of a **copy** Ledger (or leave unset to let staging create a fresh one) | recreated on first use |
| `BOOKING_WEBAPP_URL` | **leave unset** on staging (disables the write-bridge to the live Booking project) | bridge off |

The live constants stay as the defaults in `assignShifts.gs`, so **production is unchanged**.

## Setup checklist

Done once, mostly in the Google Drive / Apps Script UI. Steps marked **(B)** must be done
signed in as `inforootsandroads@gmail.com`; the rest I can drive once you share ids.

1. **(B) Copy the sheets** (Drive → right-click → *Make a copy*), into one "R&R Staging" folder:
   - Control spreadsheet → `Control_v1 (STAGING)` — this also copies the bound portal script.
   - BookingSheet → `BookingSheet (STAGING)`.
   - Guide_3h_Availability and Guide_SF_Availability → `(STAGING)` copies.
   - (Optional) Ledger → `(STAGING)` copy, or skip and let staging make a fresh one.
2. **(B) Share** all the staging copies **as Editor with your main account**, so I can push code
   and you can both inspect them.
3. **(B) Open the staging Control copy → Extensions → Apps Script → Project Settings** and send me
   the **Script ID** and the ids of the copied sheets.
4. **(me) Push the latest code** to the staging script via clasp (as your main account, now an
   editor) so staging == repo, with no manual edits.
5. **(me / B) Set the Script Properties** from the table above (I can script this, or you paste
   them into Project Settings).
6. **(B) Publish the staging web app**: Deploy → New deployment → *Web app*, Execute as
   **inforootsandroads**, Access *Anyone*. Copy the `/exec` URL — that's `STAGING_PORTAL_URL`.
   (Publishing as B is what makes the request run in B's own lane — the whole point.)
7. **(me) Build the staging login page**: copy `guide/index.html`, set `PORTAL_URL =
   STAGING_PORTAL_URL`, serve it (locally, or at a `/guide-staging/` path). Guides are NOT given
   this URL — it's ours for testing.
8. **(me/B) Seed test data**: the copy BookingSheet already holds a snapshot of real bookings
   from the copy; add a few test rows as needed. Staging won't receive live emails (the live
   Booking project writes only to the live BookingSheet), which is exactly what we want.

## Validate (prove it works, in isolation)

- `…STAGING_PORTAL_URL?action=health` → `controlOk/bookingSheetOk/ledgerOk` true, and the ids it
  reports are the **copies**, not live.
- Log in on the staging page, assign a guide, check someone in → confirm it writes to the
  **copy** sheets and the live portal is unaffected.
- Load-test: hit `…?action=tours` repeatedly / from several sessions and watch
  `…?action=timings` on staging — this is where we confirm the second-account lane is faster.

## Cutover (later, when proven)

Flip the live login page's `PORTAL_URL` to the B-owned **production** deployment (a separate
step from staging), keep the triggers on the main account, point the properties at the LIVE
sheets (or just unset them). Reversible: flip `PORTAL_URL` back.

## Teardown

Delete the staging folder (sheets + bound script + its deployment). Nothing in production
references it.
