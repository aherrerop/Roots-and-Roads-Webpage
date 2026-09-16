# Debugging & recovery — start here

One page that points at everything for diagnosing a problem fast. (Deep detail lives
in each file's **CRISIS PLAYBOOK** comment block at the very top.)

## 1. First move when something breaks: see what changed
```bash
npm run compare        # files changed since the last known-good ("what did we touch?")
npm run compare:full   # the full line-by-line diff
```
Then read the specific file's crisis playbook, or `git diff known-good..HEAD -- <file>`.
Fall back with `git checkout known-good -- .` and redeploy. See `KNOWN_GOOD.md`.

## 2. Diagnose without guessing
| Surface | How to look |
|---|---|
| **Portal server up? sheets reachable? contention?** | open `…/exec?action=health` |
| **Portal slow vs phone slow? per-action p50/p95 + slow log** | open `…/exec?action=timings&n=40` |
| **Which portal build is on a phone** | the small `Portal 2026-…` version in the page footer |
| **Booking problems** | the BookingSheet **Errors** and **Status** tabs |
| **Portal problems** | the Control sheet **Portal Log** tab |
| **A guest-facing JS error** | open the browser Console on `rootsandroadsbcn.com/guide/` |

## 3. Recovery functions (run from the Apps Script editor; reload the tab first)
| Symptom | Run |
|---|---|
| Bookings missing from the sheet | `recoverMissingBookings()` (idempotent; recovers 0 if already complete) |
| A modification stuck / duplicate after a move | `fixModificationsNow()` |
| Full re-read of all bookings | `runBookingAudit()` |
| How does one booking parse? | `diagnoseBooking('GYG…')` |
| Feed & ledger disagree / a check-in looks missing | `updateManagementQueues()` |
| Duplicate/malformed ledger rows | `repairLedgers()` |
| Availability/schedule wrong | edit **Weekly_Schedule** (the only source), then `runWeeklyScheduling()` |

## 4. Rules that make debugging possible (for whoever edits — usually Claude)
- **Verify LIVE after a portal change**: deploy, then drive `showDash()` in the browser and
  confirm "Updating…" + no boot error before calling it fixed. Node tests don't catch client
  runtime errors.
- **One change at a time**, tested, so a break can be isolated.
- **Take every console error seriously** (an "Unterminated string in JSON" was a real
  apostrophe bug, not noise).
- `node test/run-tests.js` (green + reference check) before every deploy.
- After a change is confirmed working, `npm run bless` to update the fall-back copy.

## 5. Deploy targets (so a fix reaches the right place)
- Website + guide portal (`/`, `/guide/`, `/en/…`, `llms.txt`, sitemap): `git push` → GitHub Pages.
- Booking Apps Script (`apps-script/booking`): `npm run deploy:booking`.
- Control/portal Apps Script (`apps-script/control`): `npm run deploy:control`.
