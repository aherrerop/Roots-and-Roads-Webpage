# Regression audit — what changed vs. the last known-good working portal

Baseline: **known-good = 53b64ca** (2026-10-19, "SF parity + capacity + green suite").
Scope: behavioral changes to the portal (`guidePortal.gs`, `assignShifts.gs`, `guide/index.html`)
since then. The bulk of the raw diff (~2800 lines) is **comment conversion** (JSDoc → teaching
`//`), which changes no behavior and is covered by the green suite — not listed individually.

Legend: **KEEP** = justified, low risk · **WATCH** = justified but a real-failure area (risk
comment added in code) · **FIX** = fixed this pass · **OWNER** = justified by an explicit
request; a cheaper alternative is flagged.

---

## The one that caused the incident

| # | Change | Why | Risk | Status |
|---|---|---|---|---|
| 1 | **`extendDailyAvailabilityUntilMarch`** — all languages × 10:30 & 15:30 × every day as standing offer rows | Owner wanted availability open through March | **HIGH — this is the main load driver.** It turns every day into a wall of empty "offer" cards the portal assembles, serializes and (used to) compute guide-dots for on every load → the CPU pile-up behind the Nov-19 wrong-guide incident | **OWNER.** Kept (you asked for it). Mitigated by #2. **Recommended real fix:** stop the portal from *showing* empty, unbooked, unassigned offer cards (they still drive website availability) — biggest remaining load win, but it changes manager UX so it needs your yes. |

## Performance changes (this session — they fight the overload)

| # | Change | Why | Risk | Status |
|---|---|---|---|---|
| 2 | Assign-dropdown dots computed **only for tours with guests or a guide** | mgrOpts was 4–6 s/manager-load | LOW — empty cards keep their guide list (client falls back), just no colour dots | KEEP |
| 3 | **Whole-response cache** — a repeat poll with nothing changed is served from cache | Cut repeat-poll cost to ~0 | MEDIUM — a stale response if a mutation forgets to bump a version. Keyed on BOTH versions + 25 s TTL + auth runs first | **WATCH** (risk comment in code) |
| 4 | **Snapshot self-correction** (`reapplySnapshotMutations_`) — fast path re-derives assignments from the fresh feed instead of a cold rebuild | A new assignment shows instantly; no 5-min stale window | **HIGH — this is the assignment-display path (the incident path).** A divergence from a cold build = wrong guide shown | **WATCH** (mirrors `buildScheduleFromFeed_` + overlays + "new slot ⇒ cold build" guard; 5 unit + e2e tests; risk comment in code) |
| 5 | Ledger backstop keyed on guide-set + TTL only (`cachedReadStable_`), not the global version | An assign burst was re-reading every guide ledger (3–9 s each) | MEDIUM — the backstop is up-to-TTL stale; a feed-less check-in shows late. Feed is primary + hourly reconcile covers it | KEEP |
| 6 | Manager dropdown options cached off the poll path (`mgrOpts` / `gopts:`) | Same O(shifts×guides) recompute every poll | LOW | KEEP |

## Two-account / staging scaffolding (this session)

| # | Change | Why | Risk | Status |
|---|---|---|---|---|
| 7 | External sheet ids read from Script Properties, default to the live constants (`extSheetId_`) | A staging copy can point at copy sheets with no code fork | LOW on prod (no property ⇒ identical) — adds a cheap memoised property read | KEEP |
| 8 | **Fail-closed guard** — a copy (`ENV=staging`) throws rather than fall back to a live sheet | The Nov-19 staging write-to-live incident | LOW on prod (no `ENV` ⇒ inert) | KEEP |
| 9 | `health` returns `sheets` (the ids it resolves) + `env` | Verify a deployment's isolation instead of assuming | **LOW-but-real — exposes full sheet ids on a public endpoint** | **FIX** (truncate to last 6 chars — enough to tell copy from live, not the full id) |
| 10 | `control_()` opens the Control **by id** when `CONTROL_ID` is set, else the bound sheet | Lets info run a standalone portal (two-account) | LOW on prod (no property ⇒ bound handle, as before) | KEEP |
| 11 | **Auto-failover** — `api()` falls over to `BACKUP_PORTAL_URL` on a dead/wrong/auth primary | Self-heal a bad two-account cutover | LOW — no-op until a backup url is set | KEEP |

## Earlier changes (pre-session, already live a while)

| # | Change | Why | Risk | Status |
|---|---|---|---|---|
| 12 | Persist deleted cards (closed-shifts overlay) + refresh Rates before monthly archive | Deleted cards were reappearing | MEDIUM — a closed slot must RESURFACE when a new booking lands (maxim). Verified: the closed filter resurfaces on any real booking | KEEP |
| 13 | Check-in: correct the count **down**, instant + crash-proof | Record who actually came | MEDIUM — "omit a check-in in a save must NOT wipe it" (Aldo bug). Covered by test | KEEP |
| 14 | History: month picker + harden the month-ledger read | View past months | LOW — manager-only, read-only | KEEP |
| 15 | Unstaffed-tour alert (19:00 day-before email) | Never leave guests with no guide | LOW — fire-once, subject-only | KEEP |
| 16 | Backups + daily health watchdog | Fail-proofing | LOW-MED — adds triggers ⇒ more owner-serialized load (another reason for the account split) | KEEP |
| 17 | Consolidate 5 feed per-id writers → `eachFeedRowById_` | DRY | LOW — behaviourally identical | KEEP |
| 18 | Rates Summary: exclude managers; tolerate a missing trailing Note column | Reporting correctness | LOW — reporting tab, not the serving path | KEEP |

---

## Verdict
No **unjustified** behavioral differences were found — every change traces to a request or a
fix. The genuine failure-risk areas are #1 (load), #3 and #4 (the cache/assignment-display
paths). #9 is fixed this pass. #1's deeper fix (hide empty offer cards) is left for an owner
decision because it changes what managers see.
