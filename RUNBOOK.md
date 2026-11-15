# Roots & Roads — Operations Runbook

What runs automatically to keep the system safe, what to check when something seems wrong,
how to back up and restore, and the safe way to make a change. Keep this short and current.

---

## 1. Automatic safety nets (already running)

| What | When | Where |
|---|---|---|
| Booking email processing + reconcile | fast poll (time-of-day aware) | Booking project |
| Portal feed rebuild (keeps the portal fast/fresh) | every 5 min | `rebuildNearFuturePortalFeed_` |
| **All-sheet backups** (ledger, Control, BookingSheet, availability files) | weekly (Sun ~03:00) + monthly (1st) | `backupLedgerWeekly` / `backupLedgerMonthly` |
| Monthly ledger archive (freezes the month, clears the live ledger) | 1st ~02:00 | `archiveLedgerMonthly` |
| **Daily health watchdog** (emails only on a new critical failure) | ~07:00 | `dailyHealthCheck` |
| Unstaffed-tour alarm (guests, no guide) | ~19:00 day before | `alertUnstaffedToursTomorrow` |
| Data-retention purge (anonymise old PII) | monthly 1st ~04:00 | `runDataRetentionPurge` |

One-time installers (run once from the **Control** Apps Script editor): `installLedgerBackupTriggers`,
`installNearFutureFeedTrigger`, `installHealthCheckTrigger`, `installUnstaffedAlertTrigger`,
`installDataRetentionTrigger_`.

---

## 2. When something seems wrong — check in this order

1. **Control sheet → "Control" tab → SYSTEM HEALTH block.** All timestamps should be recent;
   pending counts and recent-error count should be low.
2. **`runPreflightNow()`** (Control editor) — prints PASS/FAIL for every critical check
   (timezone, booking heartbeat fresh, ledger + Portal Feed reachable, a recent backup exists,
   no duplicate booking ids). This is the fastest "is the system healthy?" answer.
3. **`systemStatus()`** (BookingSheet editor) — Gmail read access/quota, email quota,
   unprocessed threads per label, installed triggers, newest errors.
4. **Errors tab** (Control sheet, and BookingSheet) — deduped, with a Count + Last-seen.

### Common failures → fix
- **Emails stopped processing** → almost always the Gmail daily read quota. Check the
  BookingSheet Status/Errors tab; it recovers within 24h and unprocessed mail is retried. The
  lever is poll frequency, not logic.
- **Portal shows stale/no tours** → the feed-rebuild trigger is missing
  (`runPreflightNow()` flags it) → run `installNearFutureFeedTrigger`.
- **A booking is missing from the portal** → `recoverMissingBookings()` (BookingSheet editor).
- **A booking won't go away after you remove it** → its confirmation email is re-adding it;
  the portal's Remove already trashes it, else run the booking-side cleanup.
- **Dates/times look shifted** → project timezone drifted from Europe/Madrid (health check
  flags it) → Apps Script → Project Settings → Time zone.

---

## 3. Backups & restore

- **What + where:** every critical spreadsheet (Guide_Ledger_v1, Control_v1, the BookingSheet,
  the two availability files) is copied into the **"Ledger Backups"** folder, dated + tagged
  `weekly` / `monthly` / `manual`, and pruned (12 weekly ≈ 3 months, 24 monthly ≈ 2 years).
- **Code backup:** the git repo + the `known-good` tag. `npm run compare` diffs known-good→HEAD.
- **Before any risky change:** `backupEverythingNow()` (Control editor) for a fresh copy of
  everything.
- **Restore the ledger (non-destructive):** `listLedgerBackups()` → copy an id →
  `restoreLedgerFromBackup('<id>')`. It makes a fresh working copy and repoints `LEDGER_ID`;
  the previous (possibly corrupted) ledger is left untouched in Drive.
- **Restore any other sheet:** open the dated copy in the Backups folder and copy the tab(s)
  back, or re-point the relevant `*_ID` property/constant.
- **Restore the code:** `git checkout known-good -- <file>` (or `git reset`), then redeploy.

---

## 4. Safe way to make a change (the one that has saved us)

1. One change at a time. Keep a clean diff.
2. `node test/run-tests.js` must be green (deploy scripts run it first). The reference check
   also fails on any undefined call / duplicate function.
3. Deploy: `npm run deploy:control` or `npm run deploy:booking` (code) and `git push`
   (publishes the portal page via GitHub Pages).
4. **Verify live** when the change is observable (open the portal, confirm the feature). Node
   tests do not catch client-JS runtime errors.
5. When the owner confirms it works: `npm run bless` (moves the `known-good` tag). If a change
   breaks, fall back to `known-good`.

See the memory playbooks for the deeper methodology:
`closed-loop-system-design`, `fail-safe-patterns`, `load-cache-and-data-access`,
`failproofing-methodology`.
