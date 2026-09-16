# Known-good baseline — fall-back & comparison copy

A **`known-good` git tag** always points at the last code we agreed was working well.
It is a permanent snapshot on GitHub, so it is a real fall-back copy *and* a
comparison baseline. When something breaks, we compare current vs `known-good` to
see exactly what changed; if needed, we roll back to it.

> Golden habit when a change causes a problem: **diff before vs after.**
> `npm run compare` shows what changed since the last known-good. If that's not
> enough, `git diff <last-working-commit>..HEAD` on the specific file isolates it.

## The three commands (run from the project folder)

```bash
npm run compare
```
Shows, file by file, **what changed since the last known-good** (a quick "what did
we touch?"). Use `npm run compare:full` for the full line-by-line diff.

```bash
npm run knowngood
```
Prints which commit the current known-good is (hash, date, message).

```bash
npm run bless
```
**Promotes the current code to the new known-good** — but only after the full test
suite passes (so we never bless broken code). Run this when you're happy with how
everything is working. It moves the `known-good` tag to now and pushes it to GitHub.

## Falling back to known-good

The safe way (keeps history, re-deployable):

```bash
git checkout known-good -- .        # restore every file to the known-good copy
```
Then redeploy what changed: `npm run deploy:control`, `npm run deploy:booking`,
and/or `git commit` + `git push` for the website/portal. Ask Claude to do this — it
knows which surface each change belongs to.

To just *look at* the old copy without changing anything:
`git stash` (park current work) then `git checkout known-good` … `git checkout main`.

## Log of blessed baselines (newest first)

| Blessed (date) | Commit | Deployed versions at that point | Notes |
|---|---|---|---|
| 2026-09-16 | `known-good` tag | Control **v97** · Booking **v101** · Portal **2026-09-16-assign-dots** | Full system verified green (709 assertions); crisis playbooks added; check-in/assign/apostrophe fixes, website-recovery, Hide-from-website, availability dots all in and working. (`npm run knowngood` prints the exact commit.) |

_When you run `npm run bless`, add a row here (date, `git rev-parse --short HEAD`,
the deployed versions from the deploy output, one-line note) so we keep a readable
history of what "working" meant each time._
