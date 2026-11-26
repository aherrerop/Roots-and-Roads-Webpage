# Removed code — where to find it if something breaks

This file records code deleted from the active source because a newer path replaced it.
**Nothing is lost.** Every removed version still lives in git history and in a dated,
pushed tag. To read or restore a removed function, check out the tag below and copy it
back.

## How to get an old version back

```powershell
# See the exact code as it was live, right before the cleanup:
git show known-good-2026-11-26-planA-live:apps-script/control/guidePortal.gs > old-guidePortal.gs
# ...then copy the function you need out of old-guidePortal.gs.
```

The tag `known-good-2026-11-26-planA-live` is the **full live-verified state** (Plan A on
info @184 / main @183) with **all** of the code below still present. It is pushed to
GitHub, so it survives even a local disk loss.

---

## 2026-11-26 — dead-function cleanup (Control project, `guidePortal.gs`)

Removed ONE private helper that no longer had any caller because a newer implementation
replaced it. Behaviour-neutral, confirmed by the full suite and `check-references`.

| Removed | What it did | Replaced by | Why it was dead |
|---|---|---|---|
| `readLedgerHistory_(guides, loKey)` | Old manager-history reader: scanned each guide's ledger tab for check-ins since a cutoff. | The month-history path (`monthBoundsKeys_` + the ledger-feed helpers) that `apiHistory_` uses for the month-picker History view. | Superseded when History gained the any-month picker; the old fixed-window reader was left unreferenced (no refs in source OR tests). |

Restore source: `known-good-2026-11-26-planA-live` (or any commit at/before `b7ad8de`).

### ⚠️ Correction (2026-11-26): `writeFeedCheckin_` was RESTORED — it was NOT dead
I first removed `writeFeedCheckin_` too, thinking the batched `writeFeedCheckins_` superseded it.
It had no caller in `guidePortal.gs`, but **the feed-check-in tests call it directly**
(`portal-e2e-tests.js`, `tests6.js`) because it has **distinct semantics**: the singular NEVER
downgrades a count (`Math.max`) and keeps the first time, whereas the batched plural is
AUTHORITATIVE (a count can be corrected DOWN). They are not interchangeable. Removing it crashed
those suites. It is restored, with a comment saying to keep it.

**Two lessons (now habits):**
1. A dead-code scan must grep the **`test/`** directory too, not just the source — a test-only
   caller still means the function is referenced.
2. Verify the suite by the runner's **exit code** (`npm test; echo $?` → must be `0`), never by
   summing the per-file "RESULT:" lines — a suite that *crashes* prints no RESULT line, so the
   hand-sum looks green while a suite is actually failing.

### What was deliberately KEPT (looks unused, but isn't)

A reference scan flags ~19 functions as "never called." **Almost all are intentional** —
they are run by hand from the Apps Script editor, by time triggers, or via the admin hook
(string-dispatched by name, which a scan can't see). Do **not** delete these:

- Backup / watchdog / repair: `backupEverythingNow`, `backupLedgerNow`, `runPreflightNow`,
  `installHealthCheckTrigger`, `installUnstaffedAlertTrigger`, `repairLedgerDuplicates`,
  `setupLedgerHistoryTab` (calls the live `historyTab_` + backfill).
- Diagnostics: `portalTimingReport`, `debugPortal`, `testQueueIdempotency`.
- Manual setup / toggles: `setFreeTourCommissions`, `setupScheduleEditValidation`,
  `enableAutoScheduling`, `disableAutoScheduling`, `validateMobileControls`,
  `clearStaleControls`.
- Platform entry point: `doPost`.

These are the runbook's tools. They stay.
