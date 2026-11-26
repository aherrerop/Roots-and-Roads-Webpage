# Plan: private-file read-model (Albert's idea)

**Goal:** cut cold portal loads by reading the whole read-model from ONE private Drive JSON file
instead of reading + assembling from Sheets — without ever hiding a booking, check-in, or
assignment, and without slowing writes.

**Status (2026-11-26):** Stages 0–2 built + verified on **staging** (branch `file-read-model`).
Nothing on live. Stage 3 (freshness) is designed below, not yet built.

## Why (measured on staging, present-day data)
- A cold load serves in **~0.8s WITH the file** (file read ~0.44s) vs **~3.7–4.7s** reading +
  assembling from Sheets. Parity confirmed (identical tours).
- A **warm** cache hit is ~0.15s either way — the file isn't needed then.
- The data is 185 KB (over the 95 KB memory-cache cap), so a file is the only store that fits it.

## The safety rule (the heart of it — "Option C")
The file is a **read-only mirror**; all writes still go to the Sheets/ledger (source of truth), so
nothing is ever lost. The only risk is a *stale read* hiding a recent change. Resolved by:

**Use the file ONLY when nothing has changed since it was built.** Then:
- **Quiet period → cache cold → use the file** (fast, and safe because it's current).
- **Recent change → cache warm → use the normal fresh path** (~0.15s, always correct).

A recent change routes itself down the fresh path, so the file can never show stale data. This
needs a precise, CHEAP "has anything changed since the build?" check covering all three sources:

| Change source | Cheap signal the portal checks |
|---|---|
| Assign / move / close (portal) | `cacheVersion_()` — already bumped by these; file records it at build |
| Check-in (portal) | `feedCacheVersion_()` — bumped by check-ins; file records it at build |
| **New booking (separate booking system)** | **a tiny Drive "feed stamp" file** (below) |

The first two are in-project and free. The third is the real work.

## The new-booking signal (cross-project)
The booking system rebuilds the whole Portal Feed via `rebuildPortalFeed_()` /
`safeRebuildPortalFeed_()` (apps-script/booking/bookingList_v2.gs) whenever bookings change. So:

1. **Booking project:** at the end of `safeRebuildPortalFeed_`, write/update a tiny Drive file
   (`portal-feed-stamp`, content = `Date.now()`). Cheap, background, no user impact.
2. **Control project (file build):** `buildReadModelFile_` records that stamp value in the model.
3. **Control project (file read):** `readModelFromFile_` reads the big file + the stamp (both cheap
   Drive reads, NO big-sheet open). If `stamp > file.stamp` → a booking landed since the build →
   **fall back to the fresh path** (usually warm → fast). Else → file is current → use it.

This never opens the big booking sheet on the read path, so it stays cheap, and a new booking is
never missed — it just routes that load to the fresh path until the next 5-min rebuild catches up.

(An even faster variant — serve the file instantly, then merge a tiny "recent bookings" delta a
beat behind — is possible later; start with the fall-back-when-changed version: simpler, already
fast, always correct.)

## Stages
- **Stage 0 — Benchmark** ✅ (file read ~0.44s vs sheet read+assemble ~2.3s).
- **Stage 1 — Build the file** ✅ staging. `buildReadModelFile_` writes schedule + feed (bookings/
  guests) + check-ins + roster to a private Drive file. Flag `READMODEL_FILE`. Diagnostics:
  `action=buildreadmodel`, `action=readmodelcheck`.
- **Stage 2 — Portal reads the file on the cold path** ✅ staging. `readModelFromFile_` + injection
  into `apiTours_` (feed + schedule from the file), version-guarded, failover to Sheets. Flag
  `READMODEL_READ`. Diagnostics: `action=rmset`, `action=toursbench`.
- **Stage 3 — Freshness (Option C), NOT built yet:** add the `feedCacheVersion` + the cross-project
  **feed-stamp** to the guard so the file is used only when provably current; booking project writes
  the stamp; keep the 5-min rebuild. This is what makes it safe to actually turn on.
- **Stage 4 — Guests from the file** (already in the file; once Stage 3 is safe, guest lists are
  instant, private, no per-card fetch).
- **Stage 5 — Promote to live** after staging sign-off (deploy both projects, flip flags; reversible).

## Guardrails
- Writes never change; the file is read-only; Sheets/ledger stay the source of truth.
- Every freshness check fails **safe**: any doubt → the fresh Sheet path, never a stale file.
- Everything flag-gated (`READMODEL_FILE`, `READMODEL_READ`) and staging-first.
- Staging diagnostics (`buildreadmodel`/`readmodelcheck`/`rmset`/`toursbench`/`refreshstaging`) are
  ENV-guarded to staging and should be removed or left dormant before any live promotion.
