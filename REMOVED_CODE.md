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

Removed two private helpers that no longer had any caller because a newer implementation
replaced them. Both removals are **behaviour-neutral** (dead code — nothing called them),
confirmed by the full test suite (153 passed) and `check-references`.

| Removed | What it did | Replaced by | Why it was dead |
|---|---|---|---|
| `writeFeedCheckin_(bookingId, checkedIn, checkedAt)` | Wrote ONE check-in onto a Portal Feed row (cols M/N), one read + one write per booking. | `writeFeedCheckins_(updates, at)` — mirrors MANY check-ins in ONE read + ONE write per tour. | The batched version (commit `1fccc5c`) took over the check-in→feed write; the single-booking writer was left with no callers. |
| `readLedgerHistory_(guides, loKey)` | Old manager-history reader: scanned each guide's ledger tab for check-ins since a cutoff. | The month-history path (`monthBoundsKeys_` + the ledger-feed helpers) that `apiHistory_` uses for the month-picker History view. | Superseded when History gained the any-month picker; the old fixed-window reader was left unreferenced. |

Restore source: `known-good-2026-11-26-planA-live` (or any commit at/before `b7ad8de`).

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
