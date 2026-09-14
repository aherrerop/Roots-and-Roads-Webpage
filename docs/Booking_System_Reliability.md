# Booking System — Reliability & Damage Control

How the booking pipeline guarantees **no booking is ever missed**, how to **recover and verify** in one click when something looks wrong, and how to **change the system safely** without breaking those guarantees.

> The two maxims everything below serves: **(1) the manager can change anything**, and **(2) we can never miss an assignment, a check-in, or a booking.** A missed booking can ruin the company, so the system is built so that even a *future* bug self-heals.

---

## 1. How a booking flows

```
OTA email  ──(your Gmail filters label it)──►  Confirmations / Modifications / Cancellations labels
     │
     ▼
Booking script (runs every 5 min + a full audit)
     │  parse per source (GetYourGuide, Viator, GuruWalk, Airbnb, Website, Free Tour)
     ▼
Language tabs (English/German/Spanish/Italian/French Tours)  ──►  Done Tours (after the tour runs)
     │
     ▼
Portal Feed tab (rebuilt every run)  ──►  Guide portal reads this one tab
```

- A confirmed booking **stays in the inbox until its tour is over.** So **the inbox is the live list of upcoming bookings** — that single fact is what the safety nets rely on.
- Guest **name + phone** live only on the language tabs and the ledger; the **Done** tab keeps only anonymized counts (see the data-retention purge).

---

## 2. The safety nets (defence in depth)

Each net catches a different failure. Together they mean a booking has to slip past *all* of them to be lost — and it can't, because the last one reads the whole inbox.

| Net | What it catches | Runs |
|---|---|---|
| **Fast pass** (`runBookingSystem`) | New confirmations / modifications / cancellations | every 5 min |
| **Reliability net** (`getThreadsSafe_`, confirmations) | A just-arrived confirmation the Gmail *search index* hasn't caught yet — reads the newest few by label membership | every 5 min |
| **Loud failure flag** (`flagUnprocessedConfirmation_`) | A confirmation that parsed to nothing / an invalid booking → counted, left **UNREAD**, escalates to the audit | every 5 min |
| **Reconcile** (`reconcileConfirmationsToBookingList_`) | Any confirmed **upcoming** booking missing from the list — reads **every confirmation in the inbox**, re-inserts the missing ones | audit |
| **Invariants** (`checkInvariants_`) | Rows that should NOT be there: duplicates, cancelled-still-active, completed-still-active, contaminated names | audit |
| **Integrity check** (`verifyBookingIntegrity_`) | The reconciliation itself — counts Gmail vs the list and reports any gap on the **Verification** tab | audit |

**Why a missed booking now self-heals:** reconcile reads `label:<confirmations> in:inbox` (not a fixed newest-N slice — that was the bug that lost Olga's group of 30). Because the inbox = the upcoming set, *every* upcoming booking is re-checked on every audit regardless of how old its email is. So even if some future change wrongly deletes a row, the next audit puts it back.

---

## 3. Damage control — run these from the Apps Script editor

> After any deploy, **reload the editor tab first** so it runs the new code, not a stale copy.

| I want to… | Run | What it does |
|---|---|---|
| **Fix the sheet now** (bookings missing) | `recoverMissingBookings()` | Re-inserts every missing upcoming booking, dedupes, sorts, rebuilds the feed. Idempotent — safe to run any time. |
| **Prove nothing is missing** | `verifyBookingsNow()` | Writes the **Verification** tab: inbox counts vs the list, and a ✓ / ⚠ result. |
| **Full re-read of everything** | `runBookingAudit()` | Reconcile + invariants + integrity + cleanup, ignoring the Processed label. |
| **Apply stuck modifications / drop superseded rows** | `fixModificationsNow()` | Heals Guruwalk move duplicates + Done-stuck modifications. |
| **Debug one booking** | `diagnoseBooking('GYG…')` | Prints how each of that booking's emails parses, field by field. |

**The three tabs to glance at:**
- **Status** — "Last run finished" (if >10 min old, the trigger isn't running), and "Confirmations NOT registered this run".
- **Verification** — the reconciliation: every source's inbox confirmations vs how many are on the list, and **Missing** (must be 0).
- **Errors** — deduplicated problems, each with a count + last-seen.

---

## 4. The invariants (what must always be true)

1. **Every confirmation still in the inbox is on a language tab** — unless it is cancelled, superseded, or already run. *Guaranteed by reconcile; verified by the integrity check.*
2. **Nothing cancelled or already-run stays on a language tab.** *Guaranteed by the cancellation sweeps; verified by `checkInvariants_`.*
3. **No duplicate booking id across the language tabs.** *`checkInvariants_` (I1); dedupe is surgical — it only removes real duplicates, never a singleton, so a manual row is safe.*
4. **A name never contains an email / URL / "customer-" token.** *`checkInvariants_` (I5) catches parser contamination for any source.*
5. **A check-in shows only after the ledger write is confirmed**, and is read from the feed ∪ ledger so it can't be missed.

---

## 5. Changing the system safely (evolution checklist)

The tests are the guardrail. **`node test/run-tests.js` must stay green** (it also runs `check-references.js`, which fails the build on any undefined call or duplicate function). Before deploying a change to the booking script:

- [ ] **Ran `node test/run-tests.js`** — all suites green (≈700 assertions).
- [ ] If you touched a **parser**, added a test with the *real* email body (see `RNR_FIXTURES_` and `gygBody` in the tests) proving id + name + date + time + language all parse.
- [ ] If you touched a **removal path** (cancellation, superseded, dedupe), confirmed reconcile still re-inserts a valid booking (the `gmail-tests.js` reconcile tests).
- [ ] Field extraction stays **tolerant** — read a field by several label variants and don't hard-require exact punctuation (a single comma once broke every GetYourGuide date; a fixed newest-60 cap once lost a group booking). When in doubt, add a fallback rather than a stricter regex.
- [ ] Kept the **loud-not-silent** rule: a booking that can't be registered must be counted + left unread, never dropped quietly.

**Adding a new OTA source:** add its labels to `RNR.LABELS`, an entry to `sourceConfigs_` (`source/confirm/cancel/modify/done`), a `parse<Source>Message_`, and — if its subject carries the id — a branch in `confirmationIdFromSubject_` so the cheap cross-check covers it. The safety nets then apply to it automatically.

**Golden rule:** any change that could *remove* a row is acceptable only because reconcile will restore it if wrong — so never weaken reconcile or the "confirmation stays in the inbox until the tour is done" rule. Those two are the floor the whole guarantee stands on.

---

*Deploy: `npm run deploy:booking` (fixed deployment id, prod URL never changes). Live diagnosis IDs and the deploy details are in the team's internal notes.*
