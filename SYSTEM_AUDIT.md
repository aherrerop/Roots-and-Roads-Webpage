# System Audit — 2026‑10‑20

Scope: verify the capacity change live, confirm the SF (Sagrada exterior) tour is fully
incorporated into the guide portal and that the 3‑hour tour still works, review when
recent changes were made and the errors we have been seeing, check that each booking
platform is placed correctly and that the different parts of the system agree, and
assess what can improve guide‑portal speed in terms of logic.

Deployed at audit time: **Control v144 · Booking v127 · Portal `2026-10-17-sf-source-label`**.
Baseline tag `known-good` = `53b64ca` (2026‑10‑19). Test suite: **1054 assertions, all green**
at any hour (harness clock pinned).

---

## 0. Verdict — can you leave it and trust it?

**Yes, for day‑to‑day running.** The pipeline is healthy right now (last booking run 10:01,
0 errors, 0 backlog), SF is live and correct, 3h is untouched, and every platform in the
live sheet carries the right source tag. The capacity rule is deployed and a real live tour
already meets the new "full" threshold.

Three things are **owner actions, not bugs** (details below): re‑run **Generate schedules**
once to clear stale SF lock‑conflict rows, keep an eye on the **free‑Gmail quota** (the one
recurring systemic risk), and know that the **portal is single‑threaded** under your Google
account (the speed ceiling). None of these block trusting the system.

---

## 1. Capacity change — verified live

**Rule (deployed, Control v143/v144 + Booking v126/v127):** a group tour is "full" at
**20 people**, where `fullTourCount_(adults, children) = adults + (children > 5 ? children × 0.5 : 0)`.
Adults always count; up to 5 children are free; beyond 5, every child counts as half an adult.
Identical helper in both projects, so the portal's red outline and the capacity email always agree.

**Live case that proves it (today, Tue Oct 20):**

| Slot | Adults | Children | Old count (adults) | New count | Result |
|---|---:|---:|---:|---:|---|
| Oct 20 · 10:30 · **French** | 16 | 12 | 16 → not full | **16 + 6 = 22** | **FULL (red outline)** |

This is exactly the child‑heavy case you asked for: under the old rule it slipped by at 16
adults; now it crosses at 22. Adult‑heavy tours and tours with ≤5 children behave exactly as
before. Regression tests cover 17+13 (full), 19+6 (the 6th child tips it), 15+6 (not full),
and ≤5 kids (free), on both the portal outline and the email.

> Note I could verify in the data but not by driving the portal UI (it needs a guide/manager
> login I won't perform): the portal API recomputes `full` from this same helper, so the red
> outline on the Oct 20 French card is guaranteed by the data + deployed code + tests. Worth a
> 5‑second eyeball next time you open the portal, and confirming the capacity **email** for that
> French tour arrived.

---

## 2. SF tour — incorporation complete

SF is a **fully separate tour** that mirrors the 3h tour's machinery end‑to‑end. Live state
confirms every stage works:

| Stage | 3h | SF | Live evidence |
|---|---|---|---|
| Email parse → source tag | `GetYourGuide` etc. | `GYG-SF` / `Viator-SF` | Patrizia (9:00) & Lidia (11:00) on English Tours, source `GYG-SF`, `SF` note |
| BookingSheet routing | language tab | same language tab, `SF` note | ✓ |
| Portal Feed | 16‑col row | same row, source `GYG-SF` | ✓ |
| Portal card | 3h card | own SF card (never merges with 3h) | SF cards render (your screenshot) |
| Manager view | My / 3h tabs | third **SF tab** | ✓ |
| Grids | `Schedule_<lang>` | `Schedule_SF_<lang>` | `Schedule_SF_English` = Amalia on 9/11/13/15, Mon–Fri |
| Availability | guide file | **separate** `SF_Availability_v1` | ✓ |
| Auto‑schedule | 5h separation | **stackable** (no separation) | Amalia fills all 4 consecutive slots |
| Weekly default | Albert/Carlos | Amalia (type‑gated) | ✓ |
| Portal source pill | platform name | **shows platform** (GYG‑SF → GetYourGuide) | display‑only; ledger keeps `GYG-SF` |
| Ledger | Paid 10€/pax | **SF Exterior 3€/pax** | computeMoney_ matches SF before the paid branch |

**Fixes that got SF there (this cycle, Oct 17–19):** the parser now matches the real product
("Exterior Unlocked"), a stale `GetYourGuide` source self‑heals to `GYG-SF` on re‑read, the 5h
separation rule is exempted for SF in **five** places (balancer, assign‑dot, weekly default,
lock validator, grid/self‑test checks), and the portal shows the platform name instead of the
internal `GYG-SF` tag.

**Verdict:** SF is done. Nothing further is required to "finish incorporating" it. The only
SF residue is cosmetic — see §4 (stale error rows).

---

## 3. 3h tour — unchanged and working

Every SF change is gated on `shift.sfExt`; for a 3h shift the code paths are byte‑for‑byte
what they were. Live 3h grids (`Schedule_English/German/Italian/French`) are populated with the
expected guides, manager **locks are preserved**, and the live language tabs show normal 3h
bookings flowing (GetYourGuide, GYG, Guruwalk) with children tracked in Notes. The deterministic
test clock means the full 3h suite now passes at any hour (it used to flake only between 23:00
and midnight — a test artifact, never a product fault).

---

## 4. Errors we have been getting

**Booking project — Errors tab:** one live row.
- `parseAirbnbMessage_` threw `getDate` "Could not locate target object" once at 03:56. This is
  a **transient Gmail object** error (the message handle went stale mid‑run), not a parse bug.
  **Fixed** (Booking v127): `msg.getDate()` is now wrapped — it only supplies the year, so it
  falls back to "now" and never aborts the pass. Count was 1 and it had already self‑healed.

**Control project — Errors tab (37 rows):** two kinds.
- **Amalia SF "overlapping tours (<5h)"** — all timestamped **2026‑10‑17 22:46**, the last
  "Generate schedules" run, which was under **v141** (before v142/v143 silenced those SF
  validators). They are **stale**; a fresh generation under v144 will not reproduce them.
  → **Owner action:** click **Generate schedules** once to confirm a clean Errors tab, then the
  old SF rows can be cleared.
- **Miguel German "overlapping tours (<5h)"** — these are **real** 3h double‑bookings (Miguel on
  German 10:00 and 10:30 the same day). Not caused by any change here; it's a standing scheduling
  reality you manage manually. Worth deciding whether Miguel should ever hold two German slots
  that close together, or whether one should be reassigned.

**Portal self‑test tab:** "booking surfaces as a shift at the RIGHT date" shows **FAIL** for
English and Italian. This is a **race in the synthetic test**, not a real defect: a brand‑new
test booking is checked before the near‑future feed rebuilds, so it isn't visible yet. **Real
bookings surface fine** — the live sheet proves it (today's bookings are all on their cards).
Treat that self‑test line as a known false negative; the assign / check‑in / ledger steps after
it all PASS.

**Status tab:** healthy — last run 10:01:44, 2 processed, 2 written, **0 errors**, 0 not‑registered,
backlog 0.

---

## 5. Platform placement & cross‑part agreement

Each platform must line up across five independent places: the **Gmail label/filter**, the
**source tag** written to the sheet, the **accounting model** (paid/free in Done Tours), the
**guide pay rate** (ledger), and the **portal display** (pill + paid/free colour). Verified:

| Platform | Gmail label | Source tag | Model | Guide pay | Portal pill | Ledger type |
|---|---|---|---|---|---|---|
| GetYourGuide (main) | GetYourGuide/Confirmations | `GetYourGuide` | paid | 10€/pax | GetYourGuide | Paid |
| GetYourGuide (2nd acct) | same label (forwarded) | `GYG` | paid | 10€/pax | GYG | Paid |
| GYG — Sagrada exterior | same GYG label | `GYG-SF` | paid | **3€/pax** | GetYourGuide¹ | SF Exterior |
| Viator | Viator/Confirmations | `Viator` | paid | 10€/pax | Viator | Paid |
| Viator — Sagrada exterior | same Viator label | `Viator-SF` | paid | **3€/pax** | Viator¹ | SF Exterior |
| Guruwalk | Guruwalk/Confirmations | `Guruwalk` | free | guide owes 6€/pax; platform 4.70 | Guruwalk | Free |
| Freetour #1 | FreeTour/Confirmations | `Free Tour` | free | 6€/pax; platform 2.5 | Free Tour | Free |
| Freetour #2 (DB Tours) | same label (forwarded) | `Free-Tour` | free | 6€/pax; platform 3.5 | Free‑Tour | Free |
| Airbnb | Airbnb/Confirmations | `Airbnb` | paid | 10€/pax | Airbnb | Paid |
| Website | Webpage/Confirmations | `Website` | free | 6€/pax; platform 0 | Website | Free |

¹ SF shows the **platform** name in the portal (display‑only `srcLabel`); the booking keeps
`GYG-SF`/`Viator-SF` internally so the ledger still pays the 3€ SF rate.

**Agreement checks that pass:**
- Portal `PAID_SOURCES` (6) == Booking `MODEL` paid sources (6) — exact match, so the
  paid/free pill and the Done‑Tours accounting can't disagree.
- `SF_EXT_SOURCES` == `['GYG-SF','Viator-SF']` in the portal == the booking SF tags.
- Free‑tour platform commissions are per‑platform (Guruwalk 4.70, Free Tour 2.5, Free‑Tour 3.5,
  Website 0) and keyed case‑insensitively on the source.
- Portal‑Feed layout (16 cols) == `readPortalFeed_` == `makeLedgerRow_` ordering.
- Live sheet source values are exactly the registry strings (`GetYourGuide`, `GYG`, `GYG-SF`,
  `Guruwalk`, …) — no drift, no typos, no stray platforms.

**Agreement RISKS to respect in future edits (not current bugs):**
- The **two Apps Script projects can't share code**, so `fullTourCount_`, `PAID_SOURCES`,
  `MODEL`, `SF_EXT_SOURCES` and the rates are **duplicated by hand**. If you change one, change
  the other, or the email and the portal silently diverge. (I re‑synced the capacity comments
  this audit for exactly this reason.)
- **Adding a new platform touches five places** — Gmail filter, source registry (both projects),
  MODEL/PAID_SOURCES (both projects), a Rates‑tab row, and (if it should read as a known name)
  the portal `srcLabel`. There is a checklist at the end of this doc.

---

## 6. Possible issues during regular use

1. **Gmail quota (the one systemic risk).** The booking system runs on a **free Gmail** account
   with a ~20k‑read/day rolling budget; the 5‑minute poll spends most of it. It has been tuned
   (poll trimmed, full audit once/day) and is currently fine, but a burst of bookings or an
   extra trigger can still push it. *First stop if emails stop processing:* the Status + Errors
   tabs. Lever: poll frequency.
2. **Portal is single‑threaded** under your account (see §8). On a busy morning several guides
   loading at once queue behind each other; a cold load can hit ~9s. It recovers (the client
   backs the poll off to 60s), but it's the main felt‑slowness.
3. **Insert‑only confirmations.** A re‑read never overwrites a row's manager edits (by design).
   The one exception we had to add was GYG‑family **source** re‑tagging (the SF self‑heal). If a
   platform ever changes a booking's *core* fields by resending the confirmation (not a
   modification email), it won't update — modifications are the only update path.
4. **Children live in a Notes string** ("2 children") on the language tabs and as a real column
   only in the Portal Feed. The feed is the source of truth for the capacity maths; the Notes
   string is for humans. Fine today, but a parser that miscounts "children" in a note would
   under/over‑count the cap. (The feed column is what `full` actually uses, so this is low risk.)
5. **SF depends on owner data staying set:** Amalia marked `SF` in the Guides "Tour" column, SF
   rows present for all 4 daily times in `Weekly_Schedule`, and Amalia ticking SF availability.
   If a future week's SF rows are missing a time, that slot won't appear (that's how the missing
   Friday 13:00 slot happened before you added the row).

---

## 7. Risks for future modifications

- **Duplicate registries across projects** (§5) — the single biggest "future edit" trap.
- **The 5h separation rule is referenced in ~8 places**; SF had to be exempted in five of them.
  A new short tour type would repeat that hunt. If a second short add‑on ever appears, factor the
  "is this a stackable short tour?" test into one helper rather than per‑site `shift.sfExt`.
- **Portal assembly has many overlays** (feed → weekly shifts → cleared → defaults → 2nd guide →
  no‑show → sort). Each new behaviour tends to add an overlay. It's correct but it's the part
  most likely to grow slow; keep new work in the pre‑built feed path, not the live fallback.
- **The near‑future feed is a cache**: anything that must be correct within seconds of a manager
  action (not just the next feed rebuild) has to write through both the feed and the grid, like
  check‑ins already do.

---

## 8. Guide‑portal speed — a logic (not code) assessment

The slow‑load telemetry in the Portal Log is clear: a 9.8s load breaks down roughly as
`me≈2.0s` (reading guides/auth) + `book≈2.2s` (reading the Portal Feed) + `asm.weekly≈0.9s`
(expanding recurring offers), i.e. **reads and assembly dominate, not computation.** The
root cause (documented) is that Apps Script **serializes all executions under one Google
account**, so concurrent guide loads queue, and CPU throttling stretches each one.

**Correction after a closer read (2026‑10‑20): the portal already does the obvious logic wins.**
- **Short window — already in place.** The client loads **3 days** (`WINDOW_DAYS=3`) with "Load 3
  more days"; the server windows to `MANAGER_WINDOW_DAYS=3`. `UPCOMING_DAYS=45` is only the ceiling.
  There is no 45‑day‑per‑load cost to cut.
- **Warm guide/config read — already in place.** `readGuidesRaw_` is cached across requests on a
  **separate config version** (300s), invalidated only by a password change or close/reopen — never
  by an assign or check‑in. The `me≈2s` in a slow load is a *cold execution context* (empty cache,
  spreadsheet re‑open), not a missing cache.
- **Assembled‑schedule cache — already in place.** The shift structure is cached 180s and
  invalidated on any assign/move, so repeated polls between changes don't re‑assemble.

**The one "always‑warm feed" idea left — rebuild the pre‑built feed right after each manager action
instead of on its 5‑minute trigger — is NOT worth it here, left as‑is deliberately.** A full
rebuild assembles the whole 45‑day horizon (~3–4s); doing it inside every assign/move pushes those
from ~3s to ~6–7s, and on a single‑execution serialized runtime a slow mutation *blocks* the guide
loads it was meant to help. The 180s structure cache + the 5‑minute trigger already cover the common
case; the synchronous rebuild trades a frequent small win for an occasional large stall.

**What actually limits speed — architectural, not a tweak.** Apps Script runs this script one
execution at a time under the owner's account and periodically recycles to a **cold context** (empty
cache → must re‑open the BookingSheet + Guides sheet: `me≈2s + book≈2s`). No in‑script cache survives
a cold context, and every guide's poll competes for that one slot. The real levers are projects, not
flags, each with its own correctness risk (today's design is careful never to serve a stale
assignment or check‑in):
  1. **Serve the tours response as one versioned blob** (one read per cold load, not three) — hard
     as‑is because check‑in counts must stay live, which is exactly why structure and counts are
     cached separately today.
  2. **Give the portal its own deployment / execution identity** so guide loads don't serialize
     behind the booking triggers and each other.
  3. **Move the portal read path off Sheets** (a cached JSON endpoint refreshed by the booking side)
     so a load never re‑opens a spreadsheet.

Recommend scoping one of these deliberately rather than bolting speed onto the current path.

---

## 9. Owner action list (short)

1. **Generate schedules** once → confirms a clean Control Errors tab (clears the stale SF rows).
2. Eyeball the **Oct 20 10:30 French** card for the red "full" outline, and confirm its
   **capacity email** arrived — the live proof of the new rule.
3. Decide on **Miguel's German 10:00/10:30** overlap (reassign one, or accept it).
4. Keep watching the **Gmail quota** via the Status tab; it's the one thing that can quietly
   stop bookings processing.
5. Portal speed: the easy caching wins are already in place (§8). A real step‑change is an
   architectural project (own deployment, or a cached JSON read path) — say the word to scope one.

---

## 10. New‑platform checklist (for future me/you)

To add a booking platform without breaking agreement:
1. Gmail filter → a `Publishing Pages/<Platform>/{Confirmations,Cancellations,Modifications,Done}` label.
2. Booking `RNR.SOURCE` + `RNR.LABELS` + `sourceConfigs_` entry + a parser.
3. Booking `RNR.MODEL` (paid/free).
4. Portal `PORTAL.PAID_SOURCES` (if paid) **and** a free‑commission row if free.
5. A **Rates** tab row for the guide pay / platform commission.
6. Portal `srcLabel` only if it should display as a friendlier name.
7. Add a test asserting the new source lands on the right card with the right rate.
