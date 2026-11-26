# Plan: static read path (reads off Apps Script, writes stay on Apps Script)

**Status:** PLANNED — not started. Written 2026-11-26.
**Prereq done:** two-account split (isolation) + Plan A fast path (no cold-rebuild spikes), both live.

## The problem this solves

Every portal **read** still wakes a cold Apps Script context that re-opens ~5 Sheets,
serialized under one account → 8–13 s loads. Plan A removed the *spikes*; it did not remove
this baseline. No amount of in-Apps-Script caching removes it, because the cold context itself
is the cost.

## The idea (CQRS / "static read, dynamic write")

The schedule changes slowly (a ~5-min rebuild already computes the whole near-future feed).
So **stop reading it back through Apps Script**:

- The rebuild (and each mutation) **publishes** the read-model to a fast front door.
- The client fetches that directly for **reads** — no Apps Script run, no Sheet open, not
  serialized. Target: **sub-second, unthrottleable** (it never touches the slow runtime).
- Apps Script is hit **only for mutations** (assign / check-in), which are infrequent and
  already ~3–5 s, and for the small amount of guest PII (see below).

The client already merges its own just-made changes over slightly-stale data
(`PENDING_ASSIGN/CHECKIN/CLOSE`), so staleness is a solved problem.

## The two decisions to make first (Stage 0)

### A. What goes in the static read-model — split PII out
The slow part is assembling the **schedule** (times, languages, assigned guides, head-counts,
status, flags) — which is **not guest PII**. Guest names/phones are a small slice.
So publish the **schedule only**; fetch a card's **guests lazily** from Apps Script when a
manager opens that card (rare). This kills the latency AND keeps PII off the front door.

### B. Where to host it (the real open question)
Even PII-free, the schedule (which tours run, how full, which guide) is **business-sensitive**,
so the front door should be access-gated, not a public URL. Options:

| Host | Reads | Gating | New infra | Notes |
|---|---|---|---|---|
| **Firebase (RTDB/Firestore) free tier** ⭐ | ~50–100 ms, CDN | Real auth + security rules | Firebase project; Apps Script writes via REST | Recommended for the real thing |
| GitHub Pages (already serving the client) | ~50 ms, CDN | **None (public)** | none | Only OK if we accept the schedule being public; fastest PoC |
| Token-protected Apps Script `doGet` w/ cache | still runs Apps Script | token | none | Does NOT remove the cold start — rejected |

**Recommendation:** Firebase for the real build (free, fast, proper auth). A GitHub-Pages
**redacted** feed is fine only as a throwaway proof-of-speed if we want numbers before
committing to Firebase.

## Stages (incremental, test-first, staging-before-live — like Plan A)

- **Stage 0 — Decide A (schema) + B (host).** Define the PII-free read-model schema =
  exactly what a card needs to render without guests. Add a test asserting the published
  blob contains no PII fields.
- **Stage 1 — Publisher (write side only), flag-gated, no client change.**
  The rebuild (and mutations) also write the read-model to the front door, behind
  `STATIC_FEED_PUBLISH`. Verify the blob matches the assembled schedule. Zero user-facing change.
- **Stage 2 — Client reads the static schedule, lazily fetches guests, flag-gated.**
  Client renders the grid from the static blob; opening a card fetches that card's guests from
  Apps Script. Behind a client flag so it A/B's against the current `PORTAL_URL` path.
- **Stage 3 — Staging verification** with a real login; measure schedule render sub-second.
- **Stage 4 — Mutations republish immediately** so other managers see a change in seconds,
  not ≤5 min (optional; the 5-min rebuild already covers correctness).
- **Stage 5 — Live cutover + kill-switch.** Client flag flips to static-read. A client
  **failover** (static fetch fails → fall back to Apps Script reads) mirrors the deployment
  failover already built, and a flag mirrors `PLANA_BUILDIN` for instant rollback.

## Risks / guardrails
- The front door is a new moving part → the client **must** fall back to Apps Script reads on
  any fetch failure (reuse the existing failover pattern). Never a hard dependency.
- PII discipline: a test must prove the published blob is PII-free before Stage 1 ships.
- Firebase creds live in Script Properties (like `ADMIN_KEY`); rotate-able, never in the repo.
- Rollback at every stage is a flag flip, not a deploy.

## Where the effort is
Stage 1 (Apps Script → front-door writer + Firebase setup) is the bulk. Stage 2 reuses the
existing card render; only the data source changes. Everything else is flags + tests.
