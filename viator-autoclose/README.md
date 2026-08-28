# Viator auto-close

Closes empty Viator departures inside a per-product time window, so you get an
effective "close N hours before if nobody booked" **without** setting Viator's
native cutoff (which would cost the Excellence badge). Only departures that are
empty across **every** platform (Website, Viator, GYG, Airbnb, Guruwalk, Free
Tour) are closed — the count comes from the BookingSheet.

## How it works

```
BookingSheet Apps Script (the brain)         GitHub Action (the hands)
  ?closecheck  → empty departures + windows     every 30 min:
  ?viatorcode  → latest 2FA code from Gmail        1. GET ?closecheck
  ?viatoralert → email mgmt when blocked           2. log into supplier.viator.com
Control sheet: per-product close-hours              (reuse session; 2FA via ?viatorcode)
                                                    3. mark empty departures "Sold out"
                                                    4. CAPTCHA? → GET ?viatoralert (email mgmt)
```

Only ever clicks **"Sold out"** (reversible, not a cancellation). Never "Not
operating" / "Available".

## One-time setup

**1. Deploy the Apps Script changes**

```bash
npm run deploy:booking     # adds ?closecheck / ?viatorcode / ?viatoralert
npm run deploy:control     # adds setupViatorControls()
```

**2. Build the Control table** — in the Control Apps Script editor, run
`setupViatorControls()` once. It writes an editable table at `Control!E1:H`:

| Viator product | Product code | Close hrs before (empty) | Enabled |
|---|---|---|---|
| Barcelona Walking Tour… | 5631527P3 | 10 | ✓ |
| Private Complete… | 5631527P5 | 24 | ✓ |

Edit the hours / uncheck Enabled anytime — the bot re-reads it every run.

**3. Script properties** (booking project — most already set):
`ADMIN_KEY`, and optionally `VIATOR_ALERT_EMAIL` (where the "close these by
hand" email goes; defaults to the booking account).

**4. Capture a verified session** (avoids the ~monthly 2FA on most CI runs):

```bash
cd viator-autoclose
npm install
npx playwright install chromium
node capture-session.js       # log in by hand, press Enter, copy the base64
```

**5. GitHub secrets** (repo → Settings → Secrets and variables → Actions):
`BOOKING_WEBAPP_URL`, `ADMIN_KEY`, `VIATOR_EMAIL`, `VIATOR_PASSWORD`,
`VIATOR_STORAGE_STATE_B64` (the base64 from step 4).

**6. Dry-run first.** The bot ships with `DRY_RUN=true` — it logs what it *would*
close without clicking. Trigger it manually (Actions → "Viator auto-close" → Run
workflow), read the log. When happy, add repo **variable** `VIATOR_DRY_RUN=false`
to arm it.

## Two selectors to confirm on the first real run

The Viator DOM couldn't be inspected live while building, so two spots use
best-guess selectors and are marked `TUNING POINT` in `close-empty-tours.js`:

1. **Login form** — email / password / 2FA-code inputs + submit buttons.
2. **Date navigation** — whether `?date=YYYY-MM-DD` works or the on-page date
   field must be driven.

Departure parsing and the "Sold out" click are anchored on stable text tokens
(`(5631527P3)`, `TG8~16:00`, the "Sold out" action) and should be robust. Run
`node capture-session.js` headed once and, if a dry run mismatches, send the log
— tuning these two is quick.

## Coverage & one edge case

The bot enumerates departures **from Viator's own availability page**, so it
covers every product Viator shows — group **P3** and private **P5** — with no
schedule edits and no risk of private tours leaking onto the public website.
`?closecheck` supplies the all-platform guest counts (`guestMap`) and the
per-product windows; the bot closes a card only when its count is 0 and it's
inside that product's window.

Edge case: the sheet counts guests by `date + time + language`, not by product.
The one overlapping slot is **English 17:00** (exists on both P3 and P5). If a
private booking lands there, both the P3 and P5 17:00 departures are treated as
non-empty, so an empty group departure at 17:00 could be left open. This is
**safe** (it never *wrongly* closes a booked tour) — just an occasional missed
close on that single overlapping time. Fixing it would require the sheet to tag
Viator bookings by product (P3 vs P5), which it doesn't today.
