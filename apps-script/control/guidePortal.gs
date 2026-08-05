/******************************************************
 * ROOTS & ROADS — GUIDE PORTAL BACKEND  (Apps Script Web App)
 *
 * WHERE THIS LIVES
 *   Bind this script to the "Roots_Roads_Control_v1" spreadsheet
 *   (Extensions -> Apps Script inside that sheet), then Deploy as a Web App.
 *
 * WHAT IT DOES
 *   Serves the guide portal on your website as a JSON(P) API. The website
 *   (rootsandroadsbcn.com/guide) calls these actions:
 *     ?action=login   &email=&password=            -> validates, returns a token
 *     ?action=tours   &token=                       -> that guide's upcoming tours,
 *                                                       guests per source, contacts,
 *                                                       co-guides, and everyone's schedule
 *     ?action=save    &token=&data=<json>           -> writes check-ins to the ledger
 *
 *   Guides never open Control_v1. The script runs AS management (deploy setting),
 *   reads Control_v1 + BookingSheet server-side, and writes check-ins to a separate
 *   "Guide_Ledger_v1" spreadsheet (auto-created in the Guide Management folder,
 *   one tab per guide + a Rates tab you can edit).
 *
 * WHY JSONP (not fetch/POST)
 *   A static GitHub Pages site can't read a normal Apps Script response
 *   cross-origin (no CORS headers). JSONP (a <script> callback) works everywhere.
 *   Check-in payloads are tiny, so GET is fine.
 *
 * MONEY MODEL (editable in the ledger's "Rates" tab)
 *   Paid tours (Viator / GetYourGuide / Airbnb): WE OWE the guide 10 € per
 *     checked-in person.
 *   Free tours (Guruwalk / Free Tour / Website): the guide OWES US 6 € per
 *     checked-in person.
 *
 * MULTI-LANGUAGE
 *   Languages are read from the Guides header row (columns between "Seniority"
 *   and "Email"), so adding a "French"/"Italian" column just works. Tour language
 *   matching is by name, so a new BookingSheet language tab is picked up too.
 ******************************************************/


/******************************************************
 * 1. CONFIGURATION  — set these once.
 ******************************************************/

const PORTAL = {
  // The BookingSheet (upcoming tours + guests). Same id used by the booking script.
  BOOKING_SHEET_ID: '1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw',

  // Guide Management Drive folder (parent of Control_v1). The ledger is created here.
  LEDGER_FOLDER_ID: '1AkSO3hS5aoUP8vZXXIBKCjQrhmavUz5j',
  LEDGER_NAME: 'Guide_Ledger_v1',

  // Secret used to sign login tokens. CHANGE THIS to any long random string once.
  TOKEN_SECRET: 'CHANGE_ME_to_a_long_random_string',
  TOKEN_TTL_HOURS: 720,   // 30 days — guides stay logged in on their phones

  // Which sources are "paid" (we owe the guide). Everything else is "free".
  PAID_SOURCES: ['Viator', 'GetYourGuide', 'Airbnb'],

  // Default rates (€ per checked-in person). The Rates tab overrides these.
  DEFAULT_PAID_RATE: 10,       // paid tours: we owe the guide, € per checked-in person
  DEFAULT_FREE_RATE: 6,        // free tours: the guide owes us, € per checked-in person
  DEFAULT_PRIVATE_PAY: 75,     // private tours: flat € we owe the guide who runs it

  // Free-tour platform commission the platform charges US, € per CHECKED-IN
  // person (a cost to R&R). We charge the guide DEFAULT_FREE_RATE/person and
  // keep the difference. Editable per platform on the Rates tab. Keys are
  // matched case-insensitively against the booking Source; '' is the default.
  DEFAULT_FREE_COMMISSIONS: {
    guruwalk: 4.70,
    'free tour': 0,
    website: 0,
    '': 0                       // any other free source
  },

  // Show tours from today up to this many days ahead.
  UPCOMING_DAYS: 45,

  // Manager "All tours" loads only the near-term window first (today .. this many
  // days). Far-out tours (e.g. next month) load on demand via "Load more", which
  // adds MANAGER_WINDOW_MORE days each tap — they are costly to build and rarely
  // needed when assigning this week's tours.
  MANAGER_WINDOW_DAYS: 5,
  MANAGER_WINDOW_MORE: 14,

  // A tour stays visible on the portal until this hour (24h) of its own day,
  // so management can check prepaid/free guests after it ran.
  TOUR_VISIBLE_UNTIL_HOUR: 23,

  // BookingSheet tab name pattern per language, e.g. "English Tours".
  BOOKING_TAB_SUFFIX: ' Tours',

  // Control_v1 tab that holds the generated assignments.
  SCHEDULE_TAB: 'Schedule',
  GUIDES_TAB: 'Guides',

  // Shifts a manager has CLOSED. Keyed by tour id, this durably hides a shift
  // from the portal no matter which source produced it (grid, live booking, or
  // a Weekly_Schedule offer rule). Reversible: delete the row to reopen.
  CLOSED_TAB: 'Closed_Shifts',

  // Timing budget (ms). An operation slower than this is flagged for review in
  // the timing report, and a slow tours READ is logged (normal fast reads are
  // not, to keep the auto-refresh cheap).
  SLOW_MS: 6000,

  // Seconds the slow, shared cross-file reads (schedule grids, BookingSheet,
  // Completed Log) are cached so the frequent auto-refresh does not re-open
  // those spreadsheets every time. Invalidated immediately on any assignment /
  // move / note change. Check-ins are NEVER cached — they read live so a
  // guest ticked in shows on the very next poll.
  CACHE_TTL: 20
};


/******************************************************
 * 2. WEB APP ENTRY POINT (JSONP)
 ******************************************************/

// Actions that CHANGE data — these get a timing/result line in the Portal Log
// so a manager can see, after the fact, how long each save took, whether it
// worked, and what went wrong. Reads (tours/login/ping) are not logged, to keep
// the frequent auto-refresh cheap.
const PORTAL_MUTATIONS = { assign: 1, save: 1, move: 1, setNote: 1, closeShift: 1 };

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const callback = p.callback || '';
  const t0 = Date.now();

  let out;
  try {
    switch (p.action) {
      case 'login':  out = apiLogin_(p); break;
      case 'tours':  out = apiTours_(p); break;
      case 'save':   out = apiSave_(p);  break;
      case 'assign': out = apiAssign_(p); break;
      case 'move':   out = apiMoveBooking_(p); break;
      case 'setNote': out = apiSetNote_(p); break;
      case 'closeShift': out = apiCloseShift_(p); break;
      case 'ping':   out = { ok: true, pong: true }; break;
      case 'health': out = apiHealth_(); break;
      default:       out = { ok: false, error: 'Unknown action: ' + String(p.action || '(none)') };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }

  const ms = Date.now() - t0;
  if (PORTAL_MUTATIONS[p.action]) {
    const detail = (out && out.error) ? out.error :
      ['date=' + (p.dateKey || ''), 'time=' + (p.time || ''), 'lang=' + (p.language || ''),
       'guide=' + (p.guide || ''), 'id=' + (p.bookingId || p.id || '')].filter(x => !/=$/.test(x)).join(' ');
    portalLog_(p.action, ms, !!(out && out.ok), detail);
  } else if (p.action === 'tours' && ms > PORTAL.SLOW_MS) {
    // Reads are not normally logged (the frequent poll would flood the log), but
    // a SLOW load is exactly what we want to catch — with its per-phase breakdown
    // (schedule vs booking list vs ledger) so we can see where the time went.
    const tim = (out && out.timings) ?
      ' | ' + Object.keys(out.timings).map(function (k) { return k + '=' + out.timings[k] + 'ms'; }).join(' ') : '';
    portalLog_('tours (slow)', ms, !!(out && out.ok), 'load exceeded ' + PORTAL.SLOW_MS + 'ms' + tim);
  }

  return jsonp_(callback, out);
}

/**
 * Append a timing/result line to the "Portal Log" tab (auto-created, capped).
 * Best-effort: a logging failure must never break the actual request.
 */
function portalLog_(action, ms, ok, detail) {
  try {
    const ss = control_();
    let sh = ss.getSheetByName('Portal Log');
    if (!sh) {
      sh = ss.insertSheet('Portal Log');
      sh.getRange(1, 1, 1, 5).setValues([['When', 'Action', 'ms', 'Result', 'Detail']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      action, ms, ok ? 'OK' : 'ERROR', String(detail || '').slice(0, 300)
    ]);
    const last = sh.getLastRow();
    if (last > 700) sh.deleteRows(2, last - 500);   // keep the newest ~500 lines
  } catch (e) { /* never let logging break a request */ }
}

/**
 * SHORT-LIVED READ CACHE (the fix for 8-127s portal loads).
 *
 * The heavy cost of a tours load is re-opening the BookingSheet + reading every
 * schedule grid on EVERY 8-second poll, for every guide watching. Those reads
 * are identical for everyone for many seconds, so we cache their JSON in the
 * script cache under a version key and reuse it. Any assignment / move / note
 * change bumps the version, so a real change is never hidden behind the cache.
 *
 * Fail-open by design: if the cache is unavailable or the value is too big, the
 * read simply runs live — never an error, never stale beyond the TTL.
 * Check-ins are deliberately NOT cached (they read live every request).
 */
function cacheVersion_() {
  try { return PropertiesService.getScriptProperties().getProperty('PORTAL_CACHE_VER') || '0'; }
  catch (e) { return '0'; }
}
function bumpCacheVersion_() {
  try {
    const p = PropertiesService.getScriptProperties();
    p.setProperty('PORTAL_CACHE_VER', String((Number(p.getProperty('PORTAL_CACHE_VER')) || 0) + 1));
  } catch (e) { /* best-effort; a missed bump only means <=TTL staleness */ }
}
function cachedRead_(name, ttlSeconds, fn) {
  let cache = null, key = '';
  try {
    cache = CacheService.getScriptCache();
    key = 'rd:' + name + ':' + cacheVersion_();
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) { cache = null; }
  const val = fn();
  try {
    if (cache) {
      const s = JSON.stringify(val);
      if (s.length < 95000) cache.put(key, s, ttlSeconds || PORTAL.CACHE_TTL);   // 100KB cap
    }
  } catch (e) { /* value not cacheable this time; live read already returned */ }
  return val;
}

/** The p-th percentile (0-100) of a numeric array. Nearest-rank, no interpolation. */
function percentile_(arr, p) {
  const a = (arr || []).filter(x => typeof x === 'number' && !isNaN(x)).slice().sort((x, y) => x - y);
  if (!a.length) return 0;
  const rank = Math.ceil((p / 100) * a.length);
  return a[Math.min(a.length - 1, Math.max(0, rank - 1))];
}

/** Aggregate the raw Portal Log rows into per-action timing stats. Pure, testable. */
function summarisePortalTimings_(rows, slowMs) {
  const slow = Number(slowMs || 6000);
  const by = {};
  (rows || []).forEach(r => {
    const action = String((r[1] != null ? r[1] : '')).trim();
    if (!action) return;
    const ms = Number(r[2] || 0);
    const ok = String(r[3] || '') === 'OK';
    const g = by[action] || (by[action] = { ms: [], errors: 0 });
    g.ms.push(ms);
    if (!ok) g.errors++;
  });
  return Object.keys(by).sort().map(action => {
    const g = by[action];
    const p50 = percentile_(g.ms, 50), p95 = percentile_(g.ms, 95);
    const max = g.ms.reduce((m, x) => Math.max(m, x), 0);
    return { action, count: g.ms.length, errors: g.errors, p50, p95, max,
             status: (p95 > slow || g.errors > 0) ? 'REVIEW' : 'OK' };
  });
}

/**
 * TIMING REPORT: read the Portal Log and print (and write to a "Portal Timing"
 * tab) how long each portal action really takes — count, p50, p95, max, and a
 * REVIEW flag for anything slower than PORTAL.SLOW_MS or with errors. Run it
 * from the editor to see, from real usage, whether the portal is fast enough.
 */
function portalTimingReport() {
  const ss = control_();
  const log = ss.getSheetByName('Portal Log');
  if (!log || log.getLastRow() < 2) { console.log('No Portal Log data yet — use the portal, then run this.'); return []; }
  const rows = log.getRange(2, 1, log.getLastRow() - 1, 5).getValues();
  const stats = summarisePortalTimings_(rows, PORTAL.SLOW_MS);

  const table = [['Action', 'Count', 'Errors', 'p50 ms', 'p95 ms', 'max ms', 'Status']]
    .concat(stats.map(s => [s.action, s.count, s.errors, s.p50, s.p95, s.max, s.status]));
  let out = ss.getSheetByName('Portal Timing') || ss.insertSheet('Portal Timing');
  out.clear();
  out.getRange(1, 1, table.length, table[0].length).setValues(table);
  out.getRange(1, 1, 1, table[0].length).setFontWeight('bold');
  out.setFrozenRows(1);

  console.log('================ PORTAL TIMING (from ' + rows.length + ' logged ops) ================');
  table.forEach(r => console.log(r.join('  |  ')));
  const review = stats.filter(s => s.status === 'REVIEW');
  console.log(review.length ? ('REVIEW: ' + review.map(s => s.action).join(', ') + ' (slow or erroring)')
                            : 'All actions within the ' + PORTAL.SLOW_MS + 'ms budget.');
  return stats;
}

function jsonp_(callback, obj) {
  const json = JSON.stringify(obj);
  // If a callback name is supplied, wrap for JSONP; else return raw JSON.
  const body = callback ? `${callback}(${json});` : json;
  return ContentService
    .createTextOutput(body)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}


/******************************************************
 * 3. ACTIONS
 ******************************************************/

/** action=login  -> { ok, token, guide, languages } */
function apiLogin_(p) {
  const email = String(p.email || '').trim().toLowerCase();
  const password = String(p.password || '');

  if (!email || !password) return { ok: false, error: 'Missing email or password' };

  const guide = findGuideByEmail_(email);
  if (!guide) return { ok: false, error: 'No guide with that email' };

  // Plaintext compare (management keeps readable passwords, by design).
  if (String(guide.password) !== password) return { ok: false, error: 'Wrong password' };
  if (!guide.active) return { ok: false, error: 'This guide account is inactive' };

  return {
    ok: true,
    token: makeToken_(guide.name),
    guide: { name: guide.name, languages: guide.languages }
  };
}


/** action=tours -> { ok, guide, rates, tours:[...], schedule:[...] } */
function apiTours_(p) {
  const name = requireToken_(p.token);
  if (!name) return { ok: false, error: 'Session expired, please log in again' };

  const today = todayKey_();
  // Per-phase stopwatch, returned as `timings` so we can see exactly where a
  // load spends its milliseconds (schedule vs booking list vs ledger) — real
  // numbers, not guesses. A cached phase reads ~0ms; a cold one shows its cost.
  const T = {};
  const _t = (k, fn) => { const s = Date.now(); const r = fn(); T[k] = Date.now() - s; return r; };

  // These three cross-file reads are identical for every guide for many seconds,
  // so they are cached (invalidated on any assign/move/note change). This is the
  // core fix for the 8-127s loads: the frequent poll stops re-opening the
  // BookingSheet and re-reading every schedule grid each time.
  const rates = _t('rates', function () { return cachedRead_('rates', 60, readRates_); });
  let schedule = _t('sched', function () { return cachedRead_('sched', PORTAL.CACHE_TTL, function () { return readSchedule_(); }); });
  // Reservations come from the single "Portal Feed" tab (one read). If it is not
  // there yet, fall back to scanning the six language tabs — so the portal keeps
  // working through the feed's first rollout.
  const bookingsByKey = _t('book', function () {
    return cachedRead_('bk', PORTAL.CACHE_TTL, function () {
      return readPortalFeed_() || readBookingsIndex_();
    });
  });
  appendOrphanBookingShifts_(schedule, bookingsByKey); // bookings with no grid slot yet -> live extra shifts
  appendWeeklyScheduleShifts_(schedule);            // #4: new Weekly_Schedule offer rows show at once
  applyWeeklyDefaults_(schedule);                   // recurring default guide fills unassigned weekly slots
  // Order every shift by date then start time (so appended orphans slot into
  // their real time position, not at the end of the list).
  schedule.sort((a, b) =>
    (a.dateKey < b.dateKey ? -1 : a.dateKey > b.dateKey ? 1 : 0) ||
    (a.minutes - b.minutes) ||
    String(a.language).localeCompare(String(b.language)));

  // Who is asking — decides how much we load. A regular guide only needs their
  // OWN check-ins (their My-tours); a manager needs everyone's (they watch all).
  const me = findGuideByName_(name);
  const isManager = !!(me && me.manager);

  // CHECK-INS live in the ledger tab of the guide the tour is ASSIGNED to, and
  // only a tour TODAY or earlier can have any (a future tour has none yet). So
  // read ONLY those few tabs — never all twelve. A regular guide reads just their
  // own tab; a manager reads today's assigned guides. Read live (never cached)
  // so a guest ticked in shows on the very next poll.
  const ledgerGuides = {};
  ledgerGuides[name] = true;                        // my own tab (I may have checked in)
  if (isManager) {
    schedule.forEach(s => { if (s.dateKey <= today) (s.assigned || []).forEach(g => { if (g) ledgerGuides[g] = true; }); });
  }
  const ledger = _t('ledger', function () { return readLedgerForGuides_(Object.keys(ledgerGuides)); });   // { reservations, checkins }
  const priorCheckins = ledger.checkins;            // key|bookingId -> {n, at}

  // A completed tour's live rows move to Done, so fall back to durable sources
  // for its guests + check-in status. Only a shift TODAY-or-earlier that has
  // lost its live bookings needs this — future tours never do — so on the common
  // pre-tour poll we skip the Completed Log read entirely.
  const needBackfill = schedule.some(s => s.dateKey <= today &&
    !((bookingsByKey[shiftKey_(s.dateKey, s.minutes, s.language)] || []).length));
  if (needBackfill) {
    const doneByKey = cachedRead_('cl', PORTAL.CACHE_TTL, readCompletedLogReservations_);
    const ledgerByKey = ledger.reservations;
    new Set(Object.keys(doneByKey).concat(Object.keys(ledgerByKey))).forEach(k => {
      if (!bookingsByKey[k] || !bookingsByKey[k].length) {
        bookingsByKey[k] = doneByKey[k] || ledgerByKey[k];
      }
    });
  }

  // A deleted (closed) shift stays gone — it does NOT reappear even if it still
  // has a booking (manager's explicit call; the grid cell was also cleared, which
  // frees its guide). Reopen from the Closed_Shifts tab to bring it back.
  const closed = readClosedShifts_();
  schedule = schedule.filter(s => !closed[shiftDomId_(s)]);

  const mine = schedule.filter(s => s.assigned.some(a => sameName_(a, name)));

  const tours = mine.map(shift => {
    const key = shiftKey_(shift.dateKey, shift.minutes, shift.language);
    // A private shift shows only its private booking(s); a regular shift only
    // the non-private ones. That's what splits Fake (regular) from Fake 2 (private).
    const bookings = (bookingsByKey[key] || [])
      .filter(b => shift.private ? /privat/i.test(b.note || '') : !/privat/i.test(b.note || ''))
      .map(b => {
        const kk = key + '|' + b.bookingId;
        const ck = priorCheckins[kk];                 // ledger check-in, if any
        const feedCk = (b.feedCheckedIn != null);     // feed check-in, if any
        const isCk = feedCk || !!ck;                  // UNION: checked in on either
        return {
          bookingId: b.bookingId,
          name: b.name,
          phone: b.phone,
          source: b.source,
          guests: b.guests,               // adults
          children: Number(b.children || 0),
          infants: Number(b.infants || 0),
          paid: isPaidSource_(b.source),
          income: Number(b.income || 0),
          isPrivate: /privat/i.test(b.note || ''),
          note: String(b.note || ''),
          manualNote: String(b.manualNote || ''),   // editable per-booking note
          checked: isCk,
          checkedIn: feedCk ? Number(b.feedCheckedIn)
                            : (ck ? Number(ck.n) : Number(b.guests || 0)), // feed, else ledger, else booked default
          checkedAt: feedCk ? String(b.feedCheckedAt || '') : (ck ? (ck.at || '') : '')
        };
      });

    const bookedGuests = bookings.reduce((s, b) => s + Number(b.guests || 0), 0);
    const bookedChildren = bookings.reduce((s, b) => s + Number(b.children || 0), 0);
    const checkedGuests = bookings.reduce((s, b) => s + (b.checked ? Number(b.checkedIn || 0) : 0), 0);

    const id = shift.private ? key + '|P' + (shift.privIndex || 1) : key;
    return {
      id,
      dateKey: shift.dateKey,
      dateText: shift.dateText,
      day: shift.day,
      time: shift.time,          // display "11:00"
      timeLabel: shift.timeLabel, // display "11:00 AM"
      language: shift.language,
      coGuides: shift.assigned.filter(a => !sameName_(a, name)),
      status: shift.status,
      isPrivate: !!shift.private,
      bookedGuests,
      bookedChildren,
      checkedGuests,
      bookings
    };
  });

  // Shared tour list (shown to every guide): only THIS WEEK (today .. Sunday),
  // so the compact list stays scannable. The manager "All tours" tab below uses
  // the full upcoming window instead.
  const weekEnd = weekEndKey_();
  let thisWeek = schedule.filter(s => s.dateKey <= weekEnd);

  // A guide sees every ASSIGNED tour (so they know who is working), plus the
  // UNASSIGNED ones they could actually take: their own language, and no clash
  // with their own shifts (same 5h separation rule the assigner uses). That way
  // they can offer to cover an open tour without being shown ones they can't do.
  // Managers keep seeing everything.
  thisWeek = visibleShiftsForGuide_(thisWeek, mine, me, isManager);

  const scheduleView = thisWeek.map(s => ({
    dateKey: s.dateKey, dateText: s.dateText, day: s.day,
    time: s.time, language: s.language, assigned: s.assigned, status: s.status,
    private: !!s.private
  }));

  // Managers get the full My-tours-style view of EVERY tour (with bookings +
  // check-ins), and can save on a guide's behalf.
  let allTours = [];
  let guidesByLanguage = null;
  let busyMap = null;
  if (isManager) {
    guidesByLanguage = {};
    const seniorityOf = {};
    const raw = readGuidesRaw_();
    const cols = guideColumns_(raw.header);
    raw.rows.forEach(row => {
      const g = parseGuideRow_(row, cols);
      if (!g.name || !g.active) return;
      seniorityOf[g.name] = g.seniority;
      Object.keys(g.languages).forEach(l => {
        if (g.languages[l] === true) {
          (guidesByLanguage[l] = guidesByLanguage[l] || []).push(g.name);
        }
      });
    });
    // Most senior first (then alphabetical), so the assign list's top choice is
    // the guide management would usually pick — often a one-tap decision.
    Object.keys(guidesByLanguage).forEach(l => {
      guidesByLanguage[l].sort((a, b) =>
        (seniorityOf[a] - seniorityOf[b]) || a.localeCompare(b));
    });
    busyMap = buildBusyMap_(schedule);
  }
  // The manager "All tours" list loads the near-term window first (today .. +N
  // days); "Load more" re-requests with a bigger `days`. This keeps the common
  // refresh light — a manager assigning this week does not pay to build next
  // month's tours every 20 seconds.
  const windowDays = Math.min(PORTAL.UPCOMING_DAYS, Math.max(1, Number(p.days) || PORTAL.MANAGER_WINDOW_DAYS));
  const managerHorizon = addDaysKey_(today, windowDays);
  let hasMore = false;
  if (isManager) {
    hasMore = schedule.some(s => s.dateKey > managerHorizon);
    // Check-ins come from the targeted ledger read above, so a check-in shows no
    // matter which assigned guide on the tour tapped it.
    allTours = schedule.filter(s => s.dateKey <= managerHorizon).map(shift => {
      const key = shiftKey_(shift.dateKey, shift.minutes, shift.language);
      const primary = shift.assigned[0] || '';
      const bookings = (bookingsByKey[key] || [])
        .filter(b => shift.private ? /privat/i.test(b.note || '') : !/privat/i.test(b.note || ''))
        .map(b => {
          const kk = key + '|' + b.bookingId;
          const cke = priorCheckins[kk];              // ledger check-in, if any
          const feedCk = (b.feedCheckedIn != null);   // feed check-in, if any
          const isCk = feedCk || !!cke;               // UNION: checked in on either
          return {
            bookingId: b.bookingId, name: b.name, phone: b.phone, source: b.source, guests: b.guests,
            children: Number(b.children || 0), infants: Number(b.infants || 0),
            income: Number(b.income || 0),
            paid: isPaidSource_(b.source), isPrivate: /privat/i.test(b.note || ''),
            note: String(b.note || ''),
            manualNote: String(b.manualNote || ''),
            checked: isCk,
            checkedIn: feedCk ? Number(b.feedCheckedIn) : (cke ? Number(cke.n) : Number(b.guests || 0)),
            checkedAt: feedCk ? String(b.feedCheckedAt || '') : (cke ? (cke.at || '') : '')
          };
        });
      const aid = shift.private ? key + '|P' + (shift.privIndex || 1) : key;
      return {
        id: aid,
        dateKey: shift.dateKey, dateText: shift.dateText, day: shift.day,
        time: shift.time, timeLabel: shift.timeLabel, language: shift.language,
        privIndex: shift.privIndex || 1,
        eligible: eligibleGuidesForShift_(shift, busyMap, guidesByLanguage),
        assigned: shift.assigned, guide: primary, coGuides: shift.assigned, status: shift.status,
        isPrivate: !!shift.private,
        bookedGuests: bookings.reduce((s, b) => s + Number(b.guests || 0), 0),
        bookedChildren: bookings.reduce((s, b) => s + Number(b.children || 0), 0),
        checkedGuests: bookings.reduce((s, b) => s + (b.checked ? Number(b.checkedIn || 0) : 0), 0),
        bookings
      };
    });
  }

  return { ok: true, guide: name, manager: isManager, rates, tours,
           schedule: scheduleView, allTours, guidesByLanguage,
           hasMore: hasMore, windowDays: windowDays,
           timings: T,   // { rates, sched, book, ledger } ms — where the load went
           // Freshness: the server's own clock at the moment this data was built,
           // so the phone can show "Updated HH:mm:ss" truthfully (not its own clock).
           now: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss') };
}


/**
 * action=assign — MANAGER ONLY. Writes a guide into a Schedule_<Language>
 * grid cell from the portal. The name is written in BOLD, i.e. it becomes a
 * management LOCK that makeSchedule preserves. Empty guide = clear to
 * "Not assigned".
 *   params: token, dateKey (yyyy-MM-dd), time (24h "17:00"), language,
 *           isPrivate ("1"/""), privIndex, guide
 */
/**
 * action=setNote — MANAGER ONLY. Save (or clear) the note on a single
 * BOOKING. The note lives in column J of the booking's row in the BookingSheet
 * language tab (the sync only writes A:I, so column J is the portal's alone),
 * and rides into the ledger's Note column when the guest is checked in.
 *   params: token, bookingId, language (optional hint), note
 */
function apiSetNote_(p) {
  const name = requireToken_(p.token);
  if (!name) return { ok: false, error: 'Session expired, please log in again' };
  const me = findGuideByName_(name);
  if (!me || !me.manager) return { ok: false, error: 'Managers only' };

  const bookingId = String(p.bookingId || '').trim();
  if (!bookingId) return { ok: false, error: 'Missing booking id' };
  const note = String(p.note || '').trim();
  const language = String(p.language || '').trim();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Server busy, try again' };
  try {
    const out = writeBookingNote_(bookingId, language, note);
    if (out && out.ok) bumpCacheVersion_();   // the note is a cached field
    return out;
  } finally {
    lock.releaseLock();
  }
}

/** Column J of the booking's active row -> the per-booking note. */
function writeBookingNote_(bookingId, language, note) {
  const ss = bookingSS_();
  const idNorm = bookingId.toUpperCase().replace(/\s+/g, '');

  // Prefer the hinted language tab, then scan every active language tab (a
  // manager may have moved the guest to another language).
  const tabs = [];
  if (language) { const s = ss.getSheetByName(language + PORTAL.BOOKING_TAB_SUFFIX); if (s) tabs.push(s); }
  ss.getSheets().forEach(sh => {
    const tab = sh.getName();
    if (tab.indexOf(PORTAL.BOOKING_TAB_SUFFIX) === -1 || /^done\b/i.test(tab)) return;
    if (tabs.indexOf(sh) === -1) tabs.push(sh);
  });

  for (const sh of tabs) {
    const last = sh.getLastRow();
    if (last < 2) continue;
    const ids = sh.getRange(2, 8, last - 1, 1).getValues();   // col H = Booking ID
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').toUpperCase().replace(/\s+/g, '') === idNorm) {
        if (String(sh.getRange(1, 10).getValue() || '').trim() !== 'Note') {
          sh.getRange(1, 10).setValue('Note').setFontWeight('bold');   // ensure the header once
        }
        sh.getRange(i + 2, 10).setValue(note);   // col J
        return { ok: true, bookingId: bookingId, note: note };
      }
    }
  }
  return { ok: false, error: 'Booking ' + bookingId + ' not found' };
}

/** The tour id a shift shows under — regular = key, private = key|P<index>. */
function shiftDomId_(s) {
  const k = shiftKey_(s.dateKey, s.minutes, s.language);
  return s.private ? k + '|P' + (s.privIndex || 1) : k;
}

/** Set of tour ids a manager has closed (hidden from the portal). */
function readClosedShifts_() {
  const set = {};
  const sh = control_().getSheetByName(PORTAL.CLOSED_TAB);
  if (!sh || sh.getLastRow() < 2) return set;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
    const id = String(r[0] || '').trim();
    if (id) set[id] = true;
  });
  return set;
}

/**
 * action=closeShift — MANAGER ONLY. Hide a shift from the portal (it
 * disappears), or reopen it. The suppression is durable and source-agnostic:
 * it outlives makeSchedule and also stops a Weekly_Schedule offer rule from
 * re-surfacing the shift.
 *   params: token, id (the tour id), reopen ("1" to un-hide)
 */
function apiCloseShift_(p) {
  const name = requireToken_(p.token);
  if (!name) return { ok: false, error: 'Session expired, please log in again' };
  const me = findGuideByName_(name);
  if (!me || !me.manager) return { ok: false, error: 'Managers only' };

  const id = String(p.id || '').trim();
  if (!id) return { ok: false, error: 'Missing tour id' };
  const reopen = String(p.reopen || '') === '1';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Server busy, try again' };
  try {
    const ss = control_();
    let sh = ss.getSheetByName(PORTAL.CLOSED_TAB);
    if (!sh) {
      sh = ss.insertSheet(PORTAL.CLOSED_TAB);
      sh.getRange(1, 1, 1, 3).setValues([['Tour id', 'Closed by', 'Closed at']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    const last = sh.getLastRow();
    let row = -1;
    if (last >= 2) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) { if (String(ids[i][0] || '').trim() === id) { row = i + 2; break; } }
    }
    if (reopen) {
      if (row !== -1) sh.deleteRow(row);
      bumpCacheVersion_();
      return { ok: true, id, closed: false };
    }
    if (row === -1) {
      const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      sh.getRange(sh.getLastRow() + 1, 1, 1, 3).setValues([[id, name, stamp]]);
    }
    // TRUE DELETE: clear the guide out of the grid cell so they are freed to be
    // assigned to another tour, and prune any column/row left empty (kills the
    // leftover empty "9:00" column).
    deleteShiftFromGrid_(id);
    SpreadsheetApp.flush();
    bumpCacheVersion_();
    return { ok: true, id, closed: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Clear a shift's cell from its Schedule_<language> grid (which FREES the guide
 * that was on it) and prune any time-column now empty on every date, plus empty
 * date rows. Id = "dateKey|minutes|language[|P<index>]".
 */
function deleteShiftFromGrid_(id) {
  const parts = String(id).split('|');
  const dateKey = parts[0];
  const minutes = Number(parts[1]);
  const langLower = parts[2] || '';
  const isPriv = /^P/i.test(parts[3] || '');
  const privIndex = isPriv ? (Number((parts[3] || '').replace(/[^0-9]/g, '')) || 1) : 1;
  if (!dateKey || !Number.isFinite(minutes) || !langLower) return false;
  const language = LANGUAGES.find(l => l.toLowerCase() === langLower) ||
                   (langLower.charAt(0).toUpperCase() + langLower.slice(1));
  const time = minutesToTime_(minutes);

  const sh = control_().getSheetByName('Schedule_' + language);
  if (!sh || sh.getLastRow() < 3) return false;
  const anchor = gridAnchor_(String(sh.getRange(1, 1).getDisplayValue() || ''));

  const timeRow = sh.getRange(2, 1, 1, Math.max(2, sh.getLastColumn())).getDisplayValues()[0];
  let colNum = -1;
  for (let c = 1; c < timeRow.length; c++) {
    const h = parseGridTimeHeader_(timeRow[c]);
    if (h && h.time === time && h.isPrivate === isPriv && (!isPriv || (h.index || 1) === privIndex)) { colNum = c + 1; break; }
  }
  const colA = sh.getRange(3, 1, sh.getLastRow() - 2, 1).getDisplayValues();
  let rowNum = -1;
  for (let i = 0; i < colA.length; i++) {
    if (gridLabelToKey_(String(colA[i][0] || '').trim(), anchor) === dateKey) { rowNum = i + 3; break; }
  }
  if (colNum > -1 && rowNum > -1) sh.getRange(rowNum, colNum).clearContent();

  pruneEmptyGridColumnsAndRows_(sh);
  styleScheduleGrid_(sh);
  return true;
}

/** Delete any TIME column empty on every date, and any date row with no
 *  assignments. The Date column, title, and header row are never touched. */
function pruneEmptyGridColumnsAndRows_(sh) {
  let lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 3 || lastCol < 2) return;
  for (let c = lastCol; c >= 2; c--) {                 // right-to-left so indices stay valid
    if (!parseGridTimeHeader_(sh.getRange(2, c).getDisplayValue())) continue;   // not a time column
    const filled = sh.getRange(3, c, lastRow - 2, 1).getDisplayValues().some(v => String(v[0] || '').trim() !== '');
    if (!filled) sh.deleteColumn(c);
  }
  lastRow = sh.getLastRow(); lastCol = sh.getLastColumn();
  for (let r = lastRow; r >= 3; r--) {                 // then empty date rows
    const filled = lastCol >= 2 &&
      sh.getRange(r, 2, 1, lastCol - 1).getDisplayValues()[0].some(v => String(v || '').trim() !== '');
    if (!filled) sh.deleteRow(r);
  }
}

/**
 * #4 — a Weekly_Schedule offer rule shows in the portal IMMEDIATELY, before
 * the weekly makeSchedule run materialises it into a Schedule_<Language> grid.
 * Expands the rules (RULES × DATES) across the visible window into "Not
 * assigned" shifts, skipping any (date, time, language) the grids or live
 * bookings already produced. "Private" rows are availability windows, not
 * offered tours, so they are not surfaced here (a private tour appears only
 * when a private booking exists).
 */
function appendWeeklyScheduleShifts_(schedule) {
  let rules;
  try { rules = readWeeklySchedule_(control_()); }   // shared with assignShifts.gs (same project)
  catch (e) { return; }                              // no Weekly_Schedule tab -> nothing to add
  if (!rules || !rules.length) return;

  const today = todayKey_();
  const maxKey = addDaysKey_(today, PORTAL.UPCOMING_DAYS);
  const have = new Set(schedule.filter(s => !s.private)
    .map(s => shiftKey_(s.dateKey, s.minutes, s.language)));

  for (let dateKey = today; dateKey <= maxKey; dateKey = addDaysKey_(dateKey, 1)) {
    const day = dayNameFromKey_(dateKey);
    rules.forEach(rule => {
      if (rule.day !== day) return;
      if (/^private$/i.test(rule.language)) return;
      if (rule.activeFrom && dateKey < toDateKey_(rule.activeFrom)) return;
      if (rule.activeUntil && dateKey > toDateKey_(rule.activeUntil)) return;
      const time = normTime24_(rule.time);
      const minutes = timeToMinutes_(time);
      if (shiftIsOver_(dateKey, minutes)) return;
      const k = shiftKey_(dateKey, minutes, rule.language);
      if (have.has(k)) return;
      have.add(k);
      schedule.push({
        dateKey, dateText: prettyDate_(dateKey), day, time, timeLabel: to12h_(time),
        minutes, language: rule.language, private: false, assigned: [], status: 'Not assigned'
      });
    });
  }
}

function apiAssign_(p) {
  const name = requireToken_(p.token);
  if (!name) return { ok: false, error: 'Session expired, please log in again' };
  const me = findGuideByName_(name);
  if (!me || !me.manager) return { ok: false, error: 'Managers only' };

  const dateKey = String(p.dateKey || '').trim();
  const language = String(p.language || '').trim();
  const time = normTime24_(String(p.time || ''));
  const isPriv = String(p.isPrivate || '') === '1';
  const privIndex = Number(p.privIndex) || 1;
  const guide = String(p.guide || '').trim();   // '' -> unassign

  if (!dateKey || !language || !time) return { ok: false, error: 'Missing shift info' };

  if (guide) {
    const g = findGuideByName_(guide);
    if (!g) return { ok: false, error: 'Unknown guide: ' + guide };
    if (!g.active) return { ok: false, error: guide + ' is inactive' };
    if (g.languages[language] !== true) {
      return { ok: false, error: guide + ' does not speak ' + language };
    }

    // Incompatibility check: another tour within MIN_SEPARATION_HOURS.
    // Management can ENFORCE the change anyway with force=1 (the portal asks
    // for confirmation first) — the decision is theirs, but never accidental.
    if (String(p.force || '') !== '1') {
      const target = { dateKey, minutes: timeToMinutes_(time), language, private: isPriv, privIndex };
      const myKey = shiftKeyFull_(target);
      const st = shiftStartMs_(dateKey, target.minutes);
      const sepMs = ASSIGN_CFG.MIN_SEPARATION_HOURS * 3600000;
      const b = buildBusyMap_(readSchedule_())[guide.trim().toLowerCase()] || [];
      const clash = b.find(x => x.k !== myKey && Math.abs(x.ms - st) < sepMs);
      if (clash) {
        const parts = clash.k.split('|');
        return {
          ok: false, conflict: true,
          error: guide + ' already has a tour on ' + parts[0] + ' at ' +
                 to12h_(minutesToTime_(Number(parts[1]))) + ' (' + parts[2] + ') — less than ' +
                 ASSIGN_CFG.MIN_SEPARATION_HOURS + 'h apart.'
        };
      }
    }
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Server busy, try again' };
  try {
    const out = writeAssignmentToGrid_(language, dateKey, time, isPriv, privIndex, guide);
    SpreadsheetApp.flush();   // commit before returning, so the phone's next read is guaranteed fresh
    bumpCacheVersion_();      // the schedule changed -> next read must not serve a cached grid
    return out;
  } finally {
    lock.releaseLock();
  }
}


/**
 * action=move — MANAGER ONLY. Moves a booking to another language's tour
 * (e.g. a German or French guest who agrees to join the English tour).
 * The row physically moves between BookingSheet language tabs; a
 * "moved from X" note is appended so the change is traceable. The booking
 * system treats the tab as authoritative once a row exists, so emails will
 * NOT move it back.
 */
function apiMoveBooking_(p) {
  const name = requireToken_(p.token);
  if (!name) return { ok: false, error: 'Session expired, please log in again' };
  const me = findGuideByName_(name);
  if (!me || !me.manager) return { ok: false, error: 'Managers only' };

  const bookingId = String(p.bookingId || '').trim();
  const fromLanguage = String(p.fromLanguage || '').trim();
  const toLanguage = String(p.toLanguage || '').trim();
  if (!bookingId || !fromLanguage || !toLanguage) return { ok: false, error: 'Missing booking info' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: 'Server busy, try again' };
  try {
    const out = moveBookingRowBetweenTabs_(bookingId, fromLanguage, toLanguage);
    if (out && out.ok && out.moved) bumpCacheVersion_();   // a booking changed language tab
    return out;
  } finally {
    lock.releaseLock();
  }
}

function moveBookingRowBetweenTabs_(bookingId, fromLanguage, toLanguage) {
  if (fromLanguage === toLanguage) return { ok: true, moved: false };
  const ss = bookingSS_();
  const fromSh = ss.getSheetByName(fromLanguage + PORTAL.BOOKING_TAB_SUFFIX);
  const toSh = ss.getSheetByName(toLanguage + PORTAL.BOOKING_TAB_SUFFIX);
  if (!fromSh) return { ok: false, error: fromLanguage + ' Tours tab not found' };
  if (!toSh) return { ok: false, error: toLanguage + ' Tours tab not found' };

  const idNorm = bookingId.toUpperCase().replace(/\s+/g, '');
  let vals = null;

  const last = fromSh.getLastRow();
  if (last >= 2) {
    const rows = fromSh.getRange(2, 1, last - 1, 9).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][7] || '').toUpperCase().replace(/\s+/g, '') === idNorm) {
        vals = rows[i];
        fromSh.deleteRow(i + 2);
      }
    }
  }
  if (!vals) return { ok: false, error: bookingId + ' not found in ' + fromLanguage + ' Tours' };

  // Never duplicate if it somehow already exists in the target tab.
  const tLast = toSh.getLastRow();
  let exists = false;
  if (tLast >= 2) {
    exists = toSh.getRange(2, 8, tLast - 1, 1).getValues()
      .some(r => String(r[0] || '').toUpperCase().replace(/\s+/g, '') === idNorm);
  }
  if (!exists) {
    const note = String(vals[8] || '');
    if (!/moved from/i.test(note)) {
      vals[8] = (note ? note + ' · ' : '') + 'moved from ' + fromLanguage;
    }
    const row = toSh.getLastRow() + 1;
    toSh.getRange(row, 2, 1, 1).setNumberFormat('@');
    toSh.getRange(row, 5, 1, 1).setNumberFormat('@');
    toSh.getRange(row, 8, 1, 1).setNumberFormat('@');
    toSh.getRange(row, 1, 1, 9).setValues([vals]);
  }
  return { ok: true, moved: true, to: toLanguage };
}

/**
 * Write a manager assignment into Schedule_<language>, CREATING the tab, the
 * time column and/or the date row if they are missing. Management can always
 * assign a booked tour even when it was never in the pre-written offer ("there
 * are people signed up, it must be assignable"). A blank guide clears the cell.
 */
function writeAssignmentToGrid_(language, dateKey, time, isPriv, privIndex, guide) {
  const ss = control_();
  const tabName = 'Schedule_' + language;
  let sh = ss.getSheetByName(tabName);
  const header = gridHeaderForColumn_(time, isPriv, privIndex);
  const dayLabel = Utilities.formatDate(new Date(dateKey + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE MMM d');
  const newMinutes = timeToMinutes_(time);

  // Repair/create the tab if it is missing or not a real grid (e.g. the
  // "<lang>: no tours in the scheduling window" placeholder).
  let isGrid = false;
  if (sh) {
    const head = sh.getRange(2, 1).getDisplayValue();
    isGrid = String(head || '').trim().toLowerCase() === 'date';
  }
  if (!sh || !isGrid) {
    if (!sh) sh = ss.insertSheet(tabName);
    sh.clear();
    sh.getRange(1, 1).setValue(language + ' schedule (' + dateKey + ' to ' + dateKey + ')')
      .setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
    sh.getRange(2, 1, 1, 2).setValues([['Date', header]])
      .setFontWeight('bold').setBackground('#bfdbfe').setHorizontalAlignment('center');
  }

  const anchor = gridAnchor_(String(sh.getRange(1, 1).getDisplayValue() || ''));

  // ---- COLUMN: find by (time, private, index); else insert IN CHRONOLOGICAL
  //      ORDER (so a new 11:00 slots between 10:00 and 17:00, never at the end). ----
  const timeRow = sh.getRange(2, 1, 1, Math.max(2, sh.getLastColumn())).getDisplayValues()[0];
  let colNum = -1, insertColAt = -1;
  for (let c = 1; c < timeRow.length; c++) {
    const h = parseGridTimeHeader_(timeRow[c]);
    if (!h) continue;
    if (h.time === time && h.isPrivate === isPriv && (!isPriv || h.index === privIndex)) { colNum = c + 1; break; }
    if (insertColAt === -1) {
      const hm = timeToMinutes_(h.time);
      const after = hm > newMinutes ||
        (hm === newMinutes && (h.isPrivate ? 1 : 0) > (isPriv ? 1 : 0)) ||
        (hm === newMinutes && h.isPrivate === isPriv && (h.index || 1) > (privIndex || 1));
      if (after) insertColAt = c + 1;
    }
  }
  if (colNum === -1) {
    if (insertColAt === -1) {
      colNum = Math.max(2, sh.getLastColumn() + 1);
    } else {
      sh.getRange(1, 1, 1, sh.getMaxColumns()).breakApart();   // title merge would block the insert
      sh.insertColumnBefore(insertColAt);
      colNum = insertColAt;
    }
    sh.getRange(2, colNum).setValue(header)
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground(isPriv ? '#fde68a' : '#bfdbfe');
  }

  // ---- ROW: find by resolved date (window-aware, so it matches even a stale
  //      title); else insert IN CHRONOLOGICAL ORDER. ----
  const lastRow = sh.getLastRow();
  const colA = lastRow >= 3 ? sh.getRange(3, 1, lastRow - 2, 1).getDisplayValues() : [];
  let rowNum = -1, insertRowAt = -1;
  for (let i = 0; i < colA.length; i++) {
    const rk = gridLabelToKey_(String(colA[i][0] || '').trim(), anchor);
    if (rk === dateKey) { rowNum = i + 3; break; }
    if (insertRowAt === -1 && rk && rk > dateKey) insertRowAt = i + 3;
  }
  if (rowNum === -1) {
    if (insertRowAt === -1) {
      rowNum = Math.max(3, sh.getLastRow() + 1);
    } else {
      sh.insertRowBefore(insertRowAt);
      rowNum = insertRowAt;
    }
    sh.getRange(rowNum, 1).setValue(dayLabel).setFontWeight('bold').setBackground('#dbeafe');
  }

  const cell = sh.getRange(rowNum, colNum);
  if (!guide) {
    cell.setValue('Not assigned');
    cell.setFontWeight('normal').setFontStyle('italic').setFontColor('#94a3b8');
  } else {
    // Manager assignment = LOCK -> bold, so makeSchedule never moves it.
    cell.setFontStyle('normal').setFontColor('#1a2b49');
    cell.setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(guide)
        .setTextStyle(0, guide.length, SpreadsheetApp.newTextStyle().setBold(true).build())
        .build());
  }

  // Self-heal the title to the true min/max of the rows, so the anchor is always
  // accurate for the next read/write (kills the degenerate "(X to X)" title).
  refreshGridTitle_(sh, language, anchor);
  // Re-apply the scheduler's professional look so a portal-inserted row/column
  // never leaves the grid half-styled (blank rows like the old screenshot).
  styleScheduleGrid_(sh);
  return { ok: true, assigned: guide || '' };
}

/**
 * Give a Schedule_<language> grid the same clean look makeSchedule produces, no
 * matter how its rows/columns were added — so a row the PORTAL inserts is styled
 * exactly like one the weekly scheduler wrote. Idempotent and cheap (a handful
 * of range ops over a small grid). Only touches background/alignment/borders,
 * never the guide-name font, so a bold "lock" stays bold.
 */
function styleScheduleGrid_(sh) {
  try {
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    // Title (row 1): re-merge across the used width and paint it.
    sh.getRange(1, 1, 1, sh.getMaxColumns()).breakApart();
    sh.getRange(1, 1, 1, lastCol).merge()
      .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center')
      .setBackground('#2563eb').setFontColor('#ffffff');
    // Header (row 2): time columns; private ones tinted amber.
    sh.getRange(2, 1, 1, lastCol).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#bfdbfe');
    const times = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
    for (let c = 2; c <= lastCol; c++) {
      const h = parseGridTimeHeader_(times[c - 1]);
      if (h && h.isPrivate) sh.getRange(2, c).setBackground('#fde68a');
    }
    // Body (rows 3+): bordered, date column bold on light blue, cells centred.
    if (lastRow >= 3) {
      sh.getRange(3, 1, lastRow - 2, lastCol)
        .setBorder(true, true, true, true, true, true).setVerticalAlignment('middle').setWrap(true);
      sh.getRange(3, 1, lastRow - 2, 1).setFontWeight('bold').setBackground('#dbeafe');
      sh.getRange(3, 2, lastRow - 2, lastCol - 1).setHorizontalAlignment('center').setBackground('#f8fbff');
    }
    sh.setFrozenRows(2);
  } catch (e) { /* styling is cosmetic — never fail an assignment over it */ }
}

/** Rewrite the grid title to span the actual first..last dated row. */
function refreshGridTitle_(sh, language, anchor) {
  const last = sh.getLastRow();
  if (last < 3) return;
  const labels = sh.getRange(3, 1, last - 2, 1).getDisplayValues();
  let min = '', max = '';
  labels.forEach(r => {
    const k = gridLabelToKey_(String(r[0] || '').trim(), anchor);
    if (!k) return;
    if (!min || k < min) min = k;
    if (!max || k > max) max = k;
  });
  if (min && max) {
    sh.getRange(1, 1).setValue(language + ' schedule (' + min + ' to ' + max + ')');
  }
}


/** action=save -> { ok } . data = JSON: { tourId, dateKey, time, language, bookings:[{bookingId,source,name,phone,guests,checkedIn}], walkins:[{source,count}] } */
function apiSave_(p) {
  const name = requireToken_(p.token);
  if (!name) return { ok: false, error: 'Session expired, please log in again' };

  let d;
  try { d = JSON.parse(p.data || '{}'); }
  catch (err) { return { ok: false, error: 'Bad data' }; }

  if (!d.dateKey || !d.language) return { ok: false, error: 'Missing tour info' };

  // A manager may save on another guide's behalf (writes to that guide's tab).
  const me = findGuideByName_(name);
  const isManager = !!(me && me.manager);
  const targetGuide = (isManager && d.guide) ? d.guide : name;

  const rates = readRates_();
  const rows = [];
  const day = d.day || dayNameFromKey_(d.dateKey);
  const timeLabel = d.timeLabel || d.time || '';

  (d.bookings || []).forEach(b => {
    if (!b.checked) return; // only checked-in reservations get a ledger row; absence = not checked in
    const checkedIn = Math.max(0, Number(b.checkedIn || 0));
    const m = computeMoney_(b.source, checkedIn, b.isPrivate, b.income, rates);
    rows.push(makeLedgerRow_({
      dateKey: d.dateKey, day, timeLabel, language: d.language,
      bookingName: b.name || '', phone: b.phone || '', source: b.source || '',
      guests: Number(b.guests || 0), children: Number(b.children || 0), checkedIn,
      weOwe: m.weOwe, theyOwe: m.theyOwe, rrMakes: m.rrMakes, type: m.type,
      bookingId: b.bookingId || '', note: b.manualNote || ''
    }));
  });

  // Walk-ins were removed: OTA tours are prepaid, and a free-tour walk-in would
  // only mean the guide owes us commission we never generated — no guide reports
  // that. The feature was noise, so the portal no longer sends walk-ins.

  // LockService: two guides saving at once (or a double-tap) must not
  // interleave ledger writes. writeGuideLedger_ itself replaces the shift's
  // rows, so a repeated identical save is a clean overwrite, not a duplicate.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Server busy, try again in a moment' };
  try {
    writeGuideLedger_(targetGuide, d.dateKey, d.time || timeLabel, d.language, rows);
    // Mirror the check-ins onto the Portal Feed (its M/N columns) so the portal
    // reads them from the one tab. Written together with the ledger, both under
    // this lock: if a guest shows as checked-in in EITHER, they are checked in.
    const at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
    (d.bookings || []).forEach(b => {
      if (b.checked && b.bookingId) writeFeedCheckin_(b.bookingId, Math.max(0, Number(b.checkedIn || 0)), at);
    });
    SpreadsheetApp.flush();   // commit the check-in before returning, so a reload can't read stale
    // NOTE: the GuruWalk queue used to be rebuilt HERE, on every save. That read
    // the whole ledger + Completed Log and made check-ins take 7-22s to return.
    // It is idempotent and rebuilt by the hourly updateManagementQueues trigger,
    // so the guide's tap now returns immediately and the queue catches up within
    // the hour (well inside the 48h GuruWalk reporting window).
  } finally {
    lock.releaseLock();
  }
  return { ok: true, saved: rows.length, guide: targetGuide };
}


/**
 * action=health -> deployment sanity from a phone browser. No secrets, no
 * personal data, no spreadsheet contents.
 */
function apiHealth_() {
  const out = {
    ok: true,
    time: Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm:ss'),
    tz: 'Europe/Madrid',
    deployment: 'portal-v4'
  };
  try {
    const control = control_();
    out.controlOk = true;
    out.tabs = {
      guides: !!control.getSheetByName(PORTAL.GUIDES_TAB),
      scheduleGrids: control.getSheets().filter(s => s.getName().indexOf('Schedule_') === 0).length
    };
  } catch (e) { out.ok = false; out.controlOk = false; out.error = 'Control sheet unreachable'; }
  try {
    const b = bookingSS_();
    out.bookingSheetOk = true;
    out.bookingTabs = b.getSheets().filter(s =>
      s.getName().indexOf(PORTAL.BOOKING_TAB_SUFFIX) !== -1 && !/^done\b/i.test(s.getName())).length;
  } catch (e) { out.ok = false; out.bookingSheetOk = false; out.error = 'BookingSheet unreachable'; }
  try {
    out.ledgerOk = !!ledgerSS_();
  } catch (e) { out.ledgerOk = false; }
  try {
    const scriptTz = Session.getScriptTimeZone();
    out.timezoneOk = scriptTz === 'Europe/Madrid';
    if (!out.timezoneOk) { out.ok = false; out.error = 'Script timezone is ' + scriptTz + ', must be Europe/Madrid'; }
  } catch (e) { /* ignore */ }
  return out;
}


/******************************************************
 * 4. CONTROL_V1 READERS  (guides, rates, schedule)
 ******************************************************/

function control_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function readGuidesRaw_() {
  const sh = control_().getSheetByName(PORTAL.GUIDES_TAB);
  if (!sh) throw new Error('Guides tab not found');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { header: [], rows: [] };
  return { header: values[0].map(h => String(h).trim()), rows: values.slice(1) };
}

/** Language-agnostic: language columns are those between "Seniority" and "Email". */
function guideColumns_(header) {
  const idx = name => header.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const seniority = idx('Seniority');
  const email = idx('Email');
  const password = idx('Password');
  const manager = idx('Manager');
  const langStart = seniority + 1;
  // Languages are the columns after Seniority, up to the first of Manager/Email.
  const stops = [email, manager].filter(c => c > seniority);
  const langEnd = stops.length ? Math.min.apply(null, stops) : header.length;
  const languages = [];
  for (let c = langStart; c < langEnd; c++) {
    if (header[c]) languages.push({ col: c, name: header[c] });
  }
  return { nameCol: idx('Guide'), activeCol: idx('Active?'), emailCol: email, passwordCol: password,
           managerCol: idx('Manager'), seniorityCol: seniority, languages };
}

function parseGuideRow_(row, cols) {
  const languages = {};
  cols.languages.forEach(l => { languages[l.name] = row[l.col] === true; });
  return {
    name: String(row[cols.nameCol] || '').trim(),
    active: row[cols.activeCol] === true,
    email: String(row[cols.emailCol] || '').trim().toLowerCase(),
    password: row[cols.passwordCol],
    // Add a "Manager" column (TRUE/FALSE) in the Guides tab to grant the manager view.
    manager: cols.managerCol > -1 ? row[cols.managerCol] === true : false,
    // Seniority: lower number = more senior (1 first). Missing -> sorts last.
    seniority: (cols.seniorityCol > -1 && row[cols.seniorityCol] !== '' && !isNaN(row[cols.seniorityCol]))
      ? Number(row[cols.seniorityCol]) : 999,
    languages
  };
}

function findGuideByEmail_(email) {
  const { header, rows } = readGuidesRaw_();
  const cols = guideColumns_(header);
  for (const row of rows) {
    const g = parseGuideRow_(row, cols);
    if (g.email && g.email === email) return g;
  }
  return null;
}

function findGuideByName_(guideName) {
  const { header, rows } = readGuidesRaw_();
  const cols = guideColumns_(header);
  for (const row of rows) {
    const g = parseGuideRow_(row, cols);
    if (g.name && sameName_(g.name, guideName)) return g;
  }
  return null;
}

/**
 * Reads the per-language grids (Schedule_English, Schedule_German, ...) into
 * upcoming shift objects. These grids are the SOURCE OF TRUTH: makeSchedule
 * generates them, but a manager can hand-edit a cell and the portal follows.
 *
 * Grid layout (makeOneLanguageScheduleTab_):
 *   row 1  merged title "<Lang> schedule (2026-07-16 to 2026-07-24)"  (has the year)
 *   row 2  ["Date", "10:00", "11:00", "17:00", ...]
 *   row 3+ ["Thu Jul 16", <cell>, <cell>, ...]
 * A cell may stack a regular block and private blocks, separated by newlines:
 *   "Carlos, Albert"                 regular, assigned
 *   "Carlos\nNot assigned"           regular, partially/means unassigned
 *   "🔒 Bob (private)"               private group
 *   "Carlos\n🔒 Bob (private)"       regular + private in one cell
 */
function readSchedule_(opts) {
  const includePast = !!(opts && opts.includePast);
  const ss = control_();
  const today = todayKey_();
  const maxKey = addDaysKey_(today, PORTAL.UPCOMING_DAYS);
  const out = [];

  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name.indexOf('Schedule_') !== 0) return;      // only the per-language grids
    const language = name.substring('Schedule_'.length).trim();
    if (!language || sh.getLastRow() < 3) return;

    const vals = sh.getDataRange().getDisplayValues();
    const anchor = gridAnchor_(String((vals[0] && vals[0][0]) || ''));
    const timeRow = vals[1] || [];
    const times = [];
    for (let c = 1; c < timeRow.length; c++) {
      // New grids: "11:00" or "10:00 · Private [2]" headers. parseGridTimeHeader_
      // is shared with assignShifts.gs (same project).
      const h = parseGridTimeHeader_(timeRow[c]);
      if (h) times.push({ col: c, time: h.time, headerPrivate: h.isPrivate, privIndex: h.index });
    }

    for (let r = 2; r < vals.length; r++) {
      const label = String(vals[r][0] || '').trim();
      if (!label) continue;
      const dateKey = gridLabelToKey_(label, anchor);
      if (!dateKey) continue;
      if (!includePast && (dateKey < today || dateKey > maxKey)) continue;

      times.forEach(t => {
        const raw = String(vals[r][t.col] || '').trim();
        if (!raw) return;
        const minutes = timeToMinutes_(t.time);
        // A tour disappears from the portal 2h after its start. The
        // includePast reader (used to audit tours that already ran) keeps them.
        if (!includePast && shiftIsOver_(dateKey, minutes)) return;

        const lines = raw.split('\n').map(s => s.trim()).filter(Boolean);

        const base = {
          dateKey, dateText: prettyDate_(dateKey), day: dayNameFromKey_(dateKey),
          time: t.time, timeLabel: to12h_(t.time), minutes,
          language
        };

        const namesFrom = ls => ls.filter(l => !/not assigned|need \d|lock conflict/i.test(l))
          .join(',').split(',').map(s => s.trim()).filter(Boolean);

        if (t.headerPrivate) {
          // Whole column is one private group.
          out.push(Object.assign({}, base, {
            private: true, privIndex: t.privIndex,
            assigned: namesFrom(lines.map(l => l.replace(/🔒/g, '').replace(/\(private\)/ig, ''))),
            status: /not assigned/i.test(raw) ? 'Not assigned' : 'OK'
          }));
          return;
        }

        // Legacy grids could stack 🔒 private lines inside a regular cell.
        const privLines = lines.filter(l => /🔒|\(private\)/i.test(l));
        const regLines = lines.filter(l => !/🔒|\(private\)/i.test(l));

        if (regLines.length) {
          const names = namesFrom(regLines);
          out.push(Object.assign({}, base, {
            private: false, assigned: names,
            status: names.length ? 'OK' : 'Not assigned'
          }));
        }

        if (privLines.length) {
          const names = namesFrom(privLines.map(l =>
            l.replace(/🔒/g, '').replace(/\(private\)/ig, '')));
          out.push(Object.assign({}, base, {
            private: true, privIndex: 1, assigned: names,
            status: names.length ? 'OK' : 'Not assigned'
          }));
        }
      });
    }
  });

  // Dedupe: one shift per (date, time, language, private-flag). Merges
  // accidental duplicates so the portal never shows the same card twice.
  const seen = {};
  const deduped = [];
  out.forEach(s => {
    const k = shiftKey_(s.dateKey, s.minutes, s.language) + (s.private ? '|P' + (s.privIndex || 1) : '|R');
    if (seen[k]) {
      s.assigned.forEach(n => { if (seen[k].assigned.indexOf(n) === -1) seen[k].assigned.push(n); });
      if (seen[k].status !== 'OK' && s.status === 'OK') seen[k].status = 'OK';
      return;
    }
    seen[k] = s;
    deduped.push(s);
  });

  deduped.sort((a, b) => (a.dateKey + a.time).localeCompare(b.dateKey + b.time));
  return deduped;
}

/**
 * The grid window {startKey, endKey} from a title's yyyy-MM-dd dates, plus the
 * legacy {year, month} of the first one. gridLabelToKey_ uses the window to pick
 * the right YEAR for a year-less row label — the single source of a nasty class
 * of bugs where a "Jul 31" row in an "August" grid silently became next year, so
 * a portal assignment never matched its own shift and never "stuck".
 */
function gridAnchor_(title) {
  const all = String(title).match(/20\d{2}-\d{2}-\d{2}/g) || [];
  if (all.length) {
    const start = all[0], end = all[all.length - 1];
    const m = start.match(/(20\d{2})-(\d{2})-(\d{2})/);
    return { year: Number(m[1]), month: Number(m[2]) - 1, startKey: start, endKey: end };
  }
  const now = new Date();
  const todayK = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return { year: now.getFullYear(), month: now.getMonth(), startKey: todayK, endKey: todayK };
}

/**
 * "Thu Jul 16" + anchor -> "2026-07-16". Year-boundary safe AND resilient to a
 * stale or degenerate grid title: of the candidate years around the anchor, it
 * picks the one whose date falls INSIDE the grid window, else the one closest to
 * it. That way a "Jul 31" row resolves to the same 2026-07-31 the portal shift
 * uses, no matter what month the title happens to name.
 */
function gridLabelToKey_(label, anchor) {
  const m = String(label).match(/([A-Za-z]{3,})\s+(\d{1,2})\s*$/);
  if (!m) return '';
  const months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  const mon = months[m[1].slice(0, 3).toLowerCase()];
  if (mon == null) return '';
  const day = Number(m[2]);
  const baseYear = (anchor && anchor.year) || new Date().getFullYear();
  const tz = Session.getScriptTimeZone();
  const startT = (anchor && anchor.startKey) ? new Date(anchor.startKey + 'T12:00:00').getTime() : null;
  const endT   = (anchor && anchor.endKey)   ? new Date(anchor.endKey   + 'T12:00:00').getTime() : null;

  let best = '', bestScore = Infinity;
  for (let dy = -1; dy <= 1; dy++) {
    const d = new Date(baseYear + dy, mon, day, 12, 0, 0);
    if (isNaN(d.getTime())) continue;
    const t = d.getTime();
    let score;
    if (startT != null && endT != null && t >= startT && t <= endT) score = 0;         // inside the window
    else if (startT != null && endT != null) score = Math.min(Math.abs(t - startT), Math.abs(t - endT));
    else score = Math.abs(t - Date.now());
    if (score < bestScore) { bestScore = score; best = Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  }
  return best;
}


/******************************************************
 * 5. BOOKINGSHEET READER  (guests + contacts per shift)
 ******************************************************/

function bookingSS_() { return SpreadsheetApp.openById(PORTAL.BOOKING_SHEET_ID); }

/**
 * Read the BookingSheet "Portal Feed" tab (ONE read) into the same
 * "dateKey|minutes|Language" -> [bookings] index readBookingsIndex_ builds by
 * scanning six language tabs. The booking script rebuilds this feed every run.
 * Returns null if the feed is missing or empty, so apiTours_ transparently falls
 * back to the per-language-tab read while the feed is still warming up.
 * Columns: Date, Time, Language, Name, Phone, Adults, Children, Source, Income,
 *          Booking ID, Notes, Manager note, Checked-in, Check-in time.
 */
function readPortalFeed_() {
  let sh;
  try { sh = bookingSS_().getSheetByName('Portal Feed'); } catch (e) { return null; }
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues();
  const index = {};
  v.forEach(row => {
    const dateKey = toDateKey_(row[0]);
    const minutes = timeToMinutes_(normTime24_(row[1]));
    const language = String(row[2] || '').trim();
    if (!dateKey || !language) return;
    const note = String(row[10] || '').trim();
    const key = shiftKey_(dateKey, minutes, language);
    (index[key] = index[key] || []).push({
      name: String(row[3] || '').trim(),
      phone: String(row[4] || '').trim(),
      guests: Number(row[5] || 0),
      children: Number(row[6] || 0),
      infants: infantCountFromNote_(note),
      source: String(row[7] || '').trim(),
      income: Number(row[8] || 0),
      bookingId: String(row[9] || '').trim(),
      manualNote: String(row[11] || '').trim(),         // col L = Manager note
      note: note,
      // Phase 2 reads check-ins from here; blank in Phase 1 -> the ledger merge
      // still supplies check-in status, so nothing changes yet.
      feedCheckedIn: (row[12] === '' || row[12] == null) ? null : Number(row[12]),
      feedCheckedAt: String(row[13] || '')
    });
  });
  return index;
}

/**
 * Write a check-in onto the Portal Feed row for a booking — columns M (Checked-in)
 * and N (Check-in time), which are the PORTAL's alone. The 5-minute rebuild only
 * ever writes the reservation columns (A–L) and preserves M/N by Booking ID, so
 * the two writers touch DISJOINT columns and can never clobber each other.
 * Updates an existing row only; a booking not yet in the feed is covered by the
 * ledger + the union read, and the next rebuild gives it a row. Best-effort.
 */
function writeFeedCheckin_(bookingId, checkedIn, checkedAt) {
  try {
    const id = String(bookingId || '').trim();
    if (!id) return false;
    const sh = bookingSS_().getSheetByName('Portal Feed');
    if (!sh || sh.getLastRow() < 2) return false;
    const ids = sh.getRange(2, 10, sh.getLastRow() - 1, 1).getValues();   // col J = Booking ID
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim() === id) {
        sh.getRange(i + 2, 14, 1, 1).setNumberFormat('@');                 // time as text
        sh.getRange(i + 2, 13, 1, 2).setValues([[checkedIn, checkedAt]]);  // M=count, N=time
        return true;
      }
    }
    return false;
  } catch (e) { return false; }
}

/**
 * RECONCILE the Portal Feed check-ins against the ledger (the money record). We
 * write BOTH at save time, so this is a backstop for the rare case where one
 * write missed:
 *   - a check-in in the LEDGER but not the feed  -> push it onto the feed (safe,
 *     disjoint columns);
 *   - a check-in in the FEED but not the ledger  -> FLAG it (the guide may be
 *     unpaid) so a manager can fix it;
 *   - counts that disagree                       -> FLAG.
 * "Checked-in in EITHER = checked in" already holds for the portal via the union
 * read; this keeps the two stores honest. Runs hourly with the queues.
 */
function reconcilePortalFeed_() {
  try {
    const sh = bookingSS_().getSheetByName('Portal Feed');
    if (!sh || sh.getLastRow() < 2) return;
    const feed = sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues();
    const checkins = readAllCheckins_();                 // bookingId|dateKey -> ledger check-in
    const flags = [];
    feed.forEach((row, i) => {
      const dateKey = toDateKey_(row[0]);
      const id = String(row[9] || '').trim();
      if (!id || !dateKey) return;
      const feedCk = (row[12] === '' || row[12] == null) ? null : Number(row[12]);
      const led = checkins[id + '|' + dateKey];
      const ledCk = led ? Number(led.checkedIn) : null;
      if (ledCk != null && feedCk == null) {
        sh.getRange(i + 2, 14, 1, 1).setNumberFormat('@');
        sh.getRange(i + 2, 13, 1, 2).setValues([[ledCk, hhmmFromStamp_(led.updated) || '']]);   // ledger -> feed
      } else if (feedCk != null && ledCk == null) {
        flags.push('feed-only ' + id + ' (' + feedCk + ')');           // ledger write missed
      } else if (feedCk != null && ledCk != null && feedCk !== ledCk) {
        flags.push('differs ' + id + ' feed=' + feedCk + ' ledger=' + ledCk);
      }
    });
    if (flags.length) portalLog_('feed<->ledger', 0, false, 'CHECK: ' + flags.slice(0, 12).join(' | '));
  } catch (e) { portalLog_('reconcilePortalFeed_', 0, false, String(e).slice(0, 200)); }
}

/** Index all active bookings by "dateKey|minutes|Language". */
function readBookingsIndex_() {
  const ss = bookingSS_();
  const index = {};

  ss.getSheets().forEach(sh => {
    const tab = sh.getName();
    // Only the ACTIVE language tabs. "Done Tours" also ends in " Tours" but is
    // an aggregate (no booking ids) and must never be parsed as bookings.
    if (tab.indexOf(PORTAL.BOOKING_TAB_SUFFIX) === -1) return;
    if (/^done\b/i.test(tab)) return;
    const language = tab.replace(PORTAL.BOOKING_TAB_SUFFIX, '').trim();
    if (sh.getLastRow() < 2) return;

    // Read 10 cols so column J (the per-booking management note) comes through.
    // The booking sync only ever writes A:I, so column J is the portal's alone.
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
    values.forEach(row => {
      const dateKey = toDateKey_(row[3]);           // D Tour date
      const minutes = timeToMinutes_(normTime24_(row[4])); // E Time
      if (!dateKey) return;

      const note = String(row[8] || '').trim();
      const key = shiftKey_(dateKey, minutes, language);
      (index[key] = index[key] || []).push({
        name: String(row[0] || '').trim(),
        phone: String(row[1] || '').trim(),
        guests: Number(row[2] || 0),        // ADULTS (paying headcount)
        children: childCountFromNote_(note),// operational info only, never paid
        infants: infantCountFromNote_(note),
        source: String(row[5] || '').trim(),
        income: Number(row[6] || 0),        // OTA income (col G), for the R&R margin
        bookingId: String(row[7] || '').trim(),
        manualNote: String(row[9] || '').trim(),  // J: editable per-booking note
        note                                // "Private" flag comes through here
      });
    });
  });

  return index;
}


/**
 * Surface bookings that have no matching schedule shift yet as live "extra"
 * shifts, so a reservation shows in the portal IMMEDIATELY — before the weekly
 * makeSchedule run materialises it into a Schedule_<Language> grid. Read-only
 * and language-agnostic; mirrors the scheduler's expandOrphanShifts_ so a tour
 * in any language (incl. Italian/French, which may have no grid yet) appears
 * the moment its booking lands. Never duplicates a shift the grid already has.
 */
/**
 * What a guide may see in the shared tour list: every ASSIGNED tour (so they
 * know who is working), plus the UNASSIGNED ones they could actually take —
 * their own language, and no clash with their own shifts (same separation rule
 * the assigner uses). Managers see everything unfiltered.
 */
function visibleShiftsForGuide_(shifts, myShifts, guide, isManager) {
  if (isManager) return shifts;
  const sepMs = ASSIGN_CFG.MIN_SEPARATION_HOURS * 3600000;
  const myShiftMs = (myShifts || []).map(s => shiftStartMs_(s.dateKey, s.minutes));
  const speaks = (guide && guide.languages) || {};
  return (shifts || []).filter(s => {
    if (s.assigned && s.assigned.length) return true;          // someone is on it
    if (speaks[s.language] !== true) return false;             // not a language I run
    const st = shiftStartMs_(s.dateKey, s.minutes);
    return !myShiftMs.some(ms => Math.abs(ms - st) < sepMs);   // no clash with my own
  });
}


/**
 * WEEKLY RECURRING ASSIGNMENT. A Weekly_Schedule rule can name a default guide
 * (col G) who runs that day+time+language every week. Overlay that guide onto any
 * matching shift that has NO assignment yet, so the manager sets the pattern once
 * (e.g. Francesca = Italian 17:00 all week; Albert = English 10:00, Carlos =
 * English 17:00 Mon–Fri) and it carries over automatically. A manual assignment
 * for a specific date is a real grid lock — that shift already has an assigned
 * guide, so it is skipped here and the override wins for that day only.
 */
function applyWeeklyDefaults_(schedule) {
  let rules;
  try { rules = readWeeklySchedule_(control_()); } catch (e) { return; }
  const byKey = {};
  rules.forEach(r => {
    if (!r.guide) return;
    byKey[String(r.day).toLowerCase() + '|' + timeToMinutes_(normTime24_(r.time)) + '|' +
          String(r.language).toLowerCase()] = r.guide;
  });
  if (!Object.keys(byKey).length) return;

  const okCache = {};
  const speaksAndActive = (name, language) => {
    const ck = name + '|' + language;
    if (ck in okCache) return okCache[ck];
    const g = findGuideByName_(name);
    return (okCache[ck] = !!(g && g.active && g.languages[language] === true));
  };

  schedule.forEach(s => {
    if (s.private) return;                              // defaults are for regular slots
    if (s.assigned && s.assigned.length) return;        // a real assignment always wins
    const guide = byKey[String(s.day).toLowerCase() + '|' + s.minutes + '|' + String(s.language).toLowerCase()];
    if (!guide || !speaksAndActive(guide, s.language)) return;
    s.assigned = [guide];
    s.status = 'OK';
    s.weeklyDefault = true;                             // marker: filled by the weekly pattern
  });
}


function appendOrphanBookingShifts_(schedule, bookingsByKey) {
  const today = todayKey_();
  const maxKey = addDaysKey_(today, PORTAL.UPCOMING_DAYS);
  const haveReg = new Set(schedule.filter(s => !s.private).map(s => shiftKey_(s.dateKey, s.minutes, s.language)));
  const havePriv = new Set(schedule.filter(s => s.private).map(s => shiftKey_(s.dateKey, s.minutes, s.language)));

  Object.keys(bookingsByKey).forEach(key => {
    const parts = key.split('|');
    const dateKey = parts[0];
    const minutes = Number(parts[1]);
    const langLower = parts[2] || '';
    if (!dateKey || !Number.isFinite(minutes)) return;
    if (dateKey < today || dateKey > maxKey) return;   // only the upcoming window
    if (shiftIsOver_(dateKey, minutes)) return;        // not tours that already ran

    const language = LANGUAGES.find(l => l.toLowerCase() === langLower) ||
                     (langLower.charAt(0).toUpperCase() + langLower.slice(1));
    const time = Math.floor(minutes / 60) + ':' + String(minutes % 60).padStart(2, '0');
    const base = {
      dateKey, dateText: prettyDate_(dateKey), day: dayNameFromKey_(dateKey),
      time, timeLabel: to12h_(time), minutes, language
    };
    const bs = bookingsByKey[key] || [];
    if (bs.some(b => !/privat/i.test(b.note || '')) && !haveReg.has(key)) {
      haveReg.add(key);
      schedule.push(Object.assign({}, base, { private: false, assigned: [], status: 'Not assigned', extra: true }));
    }
    if (bs.some(b => /privat/i.test(b.note || '')) && !havePriv.has(key)) {
      havePriv.add(key);
      schedule.push(Object.assign({}, base, { private: true, privIndex: 1, assigned: [], status: 'Not assigned', extra: true }));
    }
  });
}


/******************************************************
 * 6. LEDGER  (Guide_Ledger_v1: one tab per guide + Rates)
 ******************************************************/

const LEDGER_HEADERS = [
  'Date', 'Day', 'Time', 'Language', 'Booking', 'Phone', 'Source',
  'Guests', 'Children', 'Checked-in', 'We owe guide (€)', 'Guide owes us (€)', 'R&R makes (€)', 'Type', 'Booking ID', 'Updated', 'Note'
];
const LEDGER_BOOKINGID_COL = 14;   // 0-based index of 'Booking ID' (used when re-reading)
const LEDGER_CHECKEDIN_COL = 9;    // 0-based index of 'Checked-in'
const LEDGER_SOURCE_COL = 6;       // 0-based index of 'Source'
const LEDGER_UPDATED_COL = 15;     // 0-based index of 'Updated' (no longer the last column)
const LEDGER_NOTE_COL = 16;        // 0-based index of 'Note' — the per-booking management note

/**
 * SAFE MIGRATION: insert the Children column into any guide tab still on the
 * old 15-column layout. Run once via setupLedger (idempotent — tabs already
 * migrated are skipped).
 */
function migrateLedgerChildrenColumn_() {
  const ss = ledgerSS_();
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'Rates' || name === 'Unassigned' || /no-shows|check-ins/i.test(name)) return;
    const lastCol = sh.getLastColumn();
    if (lastCol < 8) return;
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    if (header[8] === 'Children') return;                 // already migrated
    if (header[7] !== 'Guests') return;                   // not a guide ledger tab
    sh.insertColumnAfter(8);                              // after 'Guests'
    sh.getRange(1, 9).setValue('Children');
    if (sh.getLastRow() >= 2) sh.getRange(2, 9, sh.getLastRow() - 1, 1).setValue(0);
  });
}

function ledgerSS_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('LEDGER_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate below */ }
  }

  // Create it, move to the Guide Management folder, seed the Rates tab.
  const ss = SpreadsheetApp.create(PORTAL.LEDGER_NAME);
  props.setProperty('LEDGER_ID', ss.getId());
  try {
    const file = DriveApp.getFileById(ss.getId());
    const folder = DriveApp.getFolderById(PORTAL.LEDGER_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) { /* leave in My Drive if folder move fails */ }

  seedRatesTab_(ss);
  return ss;
}

function seedRatesTab_(ss) {
  let sh = ss.getSheetByName('Rates') || ss.insertSheet('Rates', 0);
  sh.clear();
  const rows = [
    ['Setting', 'Value'],
    ['Paid tour — we owe guide (€ per checked-in person)', PORTAL.DEFAULT_PAID_RATE],
    ['Free tour — guide owes us (€ per checked-in person)', PORTAL.DEFAULT_FREE_RATE],
    ['Private tour — we owe guide (flat € per tour)', PORTAL.DEFAULT_PRIVATE_PAY]
  ];
  // One editable commission per free-tour platform (€ per checked-in person).
  Object.keys(PORTAL.DEFAULT_FREE_COMMISSIONS).forEach(k => {
    const label = k === '' ? 'other' : k;
    rows.push(['Free tour commission — ' + label + ' (€ per checked-in person)',
               PORTAL.DEFAULT_FREE_COMMISSIONS[k]]);
  });
  rows.push(['Paid sources (comma separated)', PORTAL.PAID_SOURCES.join(', ')]);
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
  sh.setColumnWidth(1, 360); sh.setColumnWidth(2, 160);
  // Remove the default empty "Sheet1" if present.
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function readRates_() {
  const ss = ledgerSS_();
  const sh = ss.getSheetByName('Rates');
  let paid = PORTAL.DEFAULT_PAID_RATE, free = PORTAL.DEFAULT_FREE_RATE;
  let privatePay = PORTAL.DEFAULT_PRIVATE_PAY;
  let paidSources = PORTAL.PAID_SOURCES.slice();
  // Start from the defaults so a platform without a Rates row still resolves.
  const freeCommissions = {};
  Object.keys(PORTAL.DEFAULT_FREE_COMMISSIONS).forEach(k => { freeCommissions[k] = PORTAL.DEFAULT_FREE_COMMISSIONS[k]; });
  if (sh && sh.getLastRow() >= 2) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    v.forEach(r => {
      const label = String(r[0] || '').toLowerCase();
      if (label.indexOf('free tour commission') === 0) {
        // "Free tour commission — guruwalk (€ …)" -> platform key "guruwalk".
        const m = label.match(/commission\s*[—\-:]\s*([^(]+)/);
        let key = m ? m[1].trim() : '';
        if (key === 'other') key = '';
        freeCommissions[key] = Number(r[1]) || 0;
      }
      else if (label.indexOf('paid tour') === 0) paid = Number(r[1]) || paid;
      else if (label.indexOf('free tour') === 0) free = Number(r[1]) || free;
      else if (label.indexOf('private tour') === 0) privatePay = Number(r[1]) || privatePay;
      else if (label.indexOf('paid sources') === 0 && r[1]) {
        paidSources = String(r[1]).split(',').map(s => s.trim()).filter(Boolean);
      }
    });
  }
  PORTAL._paidSources = paidSources; // cache for isPaidSource_
  return { paid, free, privatePay, freeCommissions, paidSources };
}

/** Free-tour platform commission (€ per checked-in person) for a booking's
 *  source. Matches the Source against the Rates platform keys; falls back to
 *  the '' (other) commission. */
function freeCommissionFor_(source, rates) {
  const s = String(source || '').toLowerCase();
  const map = (rates && rates.freeCommissions) || {};
  const keys = Object.keys(map).filter(k => k);   // named platforms
  const hit = keys.find(k => s.indexOf(k) !== -1);
  return Number((hit != null ? map[hit] : map['']) || 0);
}

function guideTab_(ss, name) {
  const safe = name.substring(0, 90);
  let sh = ss.getSheetByName(safe);
  if (!sh) {
    sh = ss.insertSheet(safe);
    sh.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Read prior check-ins for a guide, keyed "shiftKey|bookingId" -> checkedIn. */
function readGuideCheckins_(name) {
  const ss = ledgerSS_();
  const sh = ss.getSheetByName(name.substring(0, 90));
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;

  const v = sh.getRange(2, 1, sh.getLastRow() - 1, LEDGER_HEADERS.length).getValues();
  v.forEach(r => {
    const dateKey = toDateKey_(r[0]);
    const minutes = timeToMinutes_(normTime24_(r[2]));
    const language = String(r[3] || '').trim();
    const bookingId = String(r[LEDGER_BOOKINGID_COL] || '').trim();
    if (!dateKey || !bookingId) return;
    // { n: checked-in count, at: "HH:mm" when it was saved } — the time is the
    // Updated stamp, so the portal can show WHEN a guest was checked in.
    out[shiftKey_(dateKey, minutes, language) + '|' + bookingId] = {
      n: Number(r[LEDGER_CHECKEDIN_COL] || 0),
      at: hhmmFromStamp_(r[LEDGER_UPDATED_COL])
    };
  });
  return out;
}

/** "2026-08-04 10:03" (or a Date) -> "10:03"; '' if no time is present. */
function hhmmFromStamp_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  }
  const m = String(v || '').match(/(\d{1,2}:\d{2})(?::\d{2})?\s*$/);
  return m ? m[1] : '';
}

/**
 * Full reservation detail from the ledger, keyed by shift, across every guide
 * tab. The ledger is the durable record — name, phone, guests, children,
 * checked-in, booking id — and is NOT cleared when a tour completes. Used so a
 * completed tour keeps showing its guests in the portal after the live booking
 * row has moved to Done. Deduped by booking id per shift.
 */
function readLedgerReservations_() {
  const out = {};
  let names = [];
  try {
    const raw = readGuidesRaw_();
    const cols = guideColumns_(raw.header);
    names = raw.rows.map(r => String(r[cols.nameCol] || '').trim()).filter(Boolean);
  } catch (e) { return out; }
  let ss; try { ss = ledgerSS_(); } catch (e) { return out; }
  names.forEach(name => {
    const sh = ss.getSheetByName(name.substring(0, 90));
    if (!sh || sh.getLastRow() < 2) return;
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, LEDGER_HEADERS.length).getValues();
    v.forEach(r => {
      const dateKey = toDateKey_(r[0]);
      const minutes = timeToMinutes_(normTime24_(r[2]));
      const language = String(r[3] || '').trim();
      const bookingId = String(r[LEDGER_BOOKINGID_COL] || '').trim();
      if (!dateKey || !bookingId) return;
      const key = shiftKey_(dateKey, minutes, language);
      const arr = out[key] = out[key] || [];
      if (arr.some(b => b.bookingId === bookingId)) return;   // dedupe
      arr.push({
        bookingId,
        name: String(r[4] || '').trim(),
        phone: String(r[5] || '').trim(),
        source: String(r[6] || '').trim(),
        guests: Number(r[7] || 0),
        children: Number(r[8] || 0),
        infants: 0, income: 0,
        // "Private" in the note so a private completed tour attaches to its
        // private column, not the regular one (the Type column records it).
        note: /priv/i.test(String(r[13] || '')) ? 'Private' : '',
        manualNote: String(r[LEDGER_NOTE_COL] || ''),
        checkedIn: Number(r[LEDGER_CHECKEDIN_COL] || 0)
      });
    });
  });
  return out;
}

/**
 * Read the ledger tabs of ONLY the named guides, returning BOTH:
 *   reservations: shiftKey -> [booking...]         (durable detail for done tours)
 *   checkins:     shiftKey|bookingId -> { n, at }   (newest 'Updated' wins)
 *
 * A check-in lives in the tab of the guide the tour is ASSIGNED to, and only a
 * tour today-or-earlier can have one — so apiTours_ passes just those few guide
 * names instead of sweeping all twelve tabs on every 8-second poll. Reading two
 * or three small tabs instead of twelve is the difference between a snappy
 * refresh and the 10-60s loads the portal was suffering. If two of the tour's
 * guides both ticked people in, both tabs are passed, so a co-guide's check-in
 * is still seen.
 */
function readLedgerForGuides_(names) {
  const out = { reservations: {}, checkins: {} };
  const list = (names || []).filter(Boolean);
  if (!list.length) return out;
  let ss; try { ss = ledgerSS_(); } catch (e) { return out; }
  list.forEach(name => {
    const sh = ss.getSheetByName(String(name).substring(0, 90));
    if (!sh || sh.getLastRow() < 2) return;
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, LEDGER_HEADERS.length).getValues();
    v.forEach(r => {
      const dateKey = toDateKey_(r[0]);
      const minutes = timeToMinutes_(normTime24_(r[2]));
      const language = String(r[3] || '').trim();
      const bookingId = String(r[LEDGER_BOOKINGID_COL] || '').trim();
      if (!dateKey || !bookingId) return;
      const key = shiftKey_(dateKey, minutes, language);

      // Durable reservation detail (dedupe by booking id per shift).
      const arr = out.reservations[key] || (out.reservations[key] = []);
      if (!arr.some(b => b.bookingId === bookingId)) {
        arr.push({
          bookingId,
          name: String(r[4] || '').trim(),
          phone: String(r[5] || '').trim(),
          source: String(r[6] || '').trim(),
          guests: Number(r[7] || 0),
          children: Number(r[8] || 0),
          infants: 0, income: 0,
          note: /priv/i.test(String(r[13] || '')) ? 'Private' : '',
          manualNote: String(r[LEDGER_NOTE_COL] || ''),
          checkedIn: Number(r[LEDGER_CHECKEDIN_COL] || 0)
        });
      }

      // Check-in across ALL guides — newest 'Updated' wins if two tabs collide.
      const ckKey = key + '|' + bookingId;
      const updated = String(r[LEDGER_UPDATED_COL] || '');
      const prev = out.checkins[ckKey];
      if (!prev || updated >= prev._u) {
        out.checkins[ckKey] = { n: Number(r[LEDGER_CHECKEDIN_COL] || 0),
                                at: hhmmFromStamp_(r[LEDGER_UPDATED_COL]), _u: updated };
      }
    });
  });
  return out;
}

/**
 * Full reservation detail from the BookingSheet's "Completed Log", keyed by
 * shift. The Completed Log records EVERY completed booking with full detail
 * (name, phone, adults, children, source, booking id) whether or not anyone was
 * checked in — so a done tour keeps showing its guests, and un-checked-in guests
 * still appear (their check-in simply reads as 0 / "to check in"). Deduped by
 * booking id per shift.
 */
function readCompletedLogReservations_() {
  const out = {};
  let sh; try { sh = bookingSS_().getSheetByName('Completed Log'); } catch (e) { return out; }
  if (!sh || sh.getLastRow() < 2) return out;
  // Columns: Date, Time, Language, Name, Phone, Adults, Children, Source, Income, Booking ID, Notes, Logged
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
  v.forEach(r => {
    const dateKey = toDateKey_(r[0]);
    const minutes = timeToMinutes_(normTime24_(r[1]));
    const language = String(r[2] || '').trim();
    const bookingId = String(r[9] || '').trim();
    if (!dateKey || !bookingId) return;
    const key = shiftKey_(dateKey, minutes, language);
    const arr = out[key] = out[key] || [];
    if (arr.some(b => b.bookingId === bookingId)) return;
    arr.push({
      bookingId,
      name: String(r[3] || '').trim(),
      phone: String(r[4] || '').trim(),
      guests: Number(r[5] || 0),
      children: Number(r[6] || 0),
      source: String(r[7] || '').trim(),
      infants: 0,
      income: Number(r[8] || 0),
      note: String(r[10] || '')
    });
  });
  return out;
}

/** Upsert this shift's rows for a guide (replace prior rows for the same shift).
 *  Dedupe is belt-and-braces: rows matching the shift key OR carrying one of
 *  the incoming Booking IDs on the same date are removed before the rewrite,
 *  so a repeated save can never stack duplicates even if a time cell was
 *  stored in a weird format by an older version. */
function writeGuideLedger_(name, dateKey, time, language, rows) {
  const ss = ledgerSS_();
  const sh = guideTab_(ss, name);
  const minutes = timeToMinutes_(normTime24_(time));
  const targetKey = shiftKey_(dateKey, minutes, language);

  const incomingIds = new Set();
  rows.forEach(r => {
    const id = String(r[LEDGER_BOOKINGID_COL] || '').trim();
    if (id) incomingIds.add(id + '|' + toDateKey_(r[0]));
  });

  if (sh.getLastRow() >= 2) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, LEDGER_HEADERS.length).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      const rowDate = toDateKey_(v[i][0]);
      const k = shiftKey_(rowDate, timeToMinutes_(normTime24_(v[i][2])), String(v[i][3] || '').trim());
      const idKey = String(v[i][LEDGER_BOOKINGID_COL] || '').trim() + '|' + rowDate;
      if (k === targetKey || (incomingIds.size && incomingIds.has(idKey))) sh.deleteRow(i + 2);
    }
  }

  if (rows.length) {
    const start = sh.getLastRow() + 1;
    // Time column as TEXT first, so Sheets can never coerce "11:00 AM" into a
    // Date (the root cause of the duplicate check-ins).
    sh.getRange(start, 3, rows.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, rows.length, LEDGER_HEADERS.length).setValues(rows);
  }
}


/**
 * RUN ONCE: collapses duplicate check-in rows created before the Time-format
 * fix. Keeps the NEWEST row (by the Updated column) per Booking ID + date,
 * rewrites Time cells as text, and reports what it removed.
 */
function repairLedgerDuplicates() {
  const ss = ledgerSS_();
  let removed = 0;
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'Rates' || name === 'Unassigned') return;
    if (Object.values(QUEUE_TABS).indexOf(name) !== -1) return;
    const lastCol = sh.getLastColumn();
    if (!lastCol) return;
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    if (header[7] !== 'Guests') return;   // not a guide ledger tab
    const last = lastDataRow_(sh);
    if (last < 2) return;

    const v = sh.getRange(2, 1, last - 1, LEDGER_HEADERS.length).getValues();
    const bestByKey = {};   // key -> {idx, updated}
    v.forEach((r, i) => {
      const key = String(r[LEDGER_BOOKINGID_COL] || '').trim() + '|' + toDateKey_(r[0]) +
                  '|' + timeToMinutes_(normTime24_(r[2]));
      const updated = String(r[LEDGER_UPDATED_COL] || '');
      if (!bestByKey[key] || updated >= bestByKey[key].updated) {
        bestByKey[key] = { idx: i, updated };
      }
    });
    const keepIdx = new Set(Object.keys(bestByKey).map(k => bestByKey[k].idx));
    const keepRows = v.filter((r, i) => keepIdx.has(i))
      .map(r => { r[2] = to12h_(normTime24_(r[2])); return r; });   // Time back to text
    removed += v.length - keepRows.length;

    sh.getRange(2, 1, last - 1, LEDGER_HEADERS.length).clearContent();
    if (keepRows.length) {
      sh.getRange(2, 3, keepRows.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, keepRows.length, LEDGER_HEADERS.length).setValues(keepRows);
    }
  });
  Logger.log('Duplicate ledger rows removed: ' + removed);
  return removed;
}

/**
 * All money for one ledger row.
 *   Private tour  -> we owe the guide a flat privatePay; R&R = OTA income - privatePay.
 *   Paid tour     -> we owe the guide 10 €/checked-in; R&R = OTA income - that.
 *   Free tour     -> guide owes us 6 €/checked-in; the platform charges us a
 *                    commission €/checked-in, so R&R keeps (free - commission)
 *                    per checked-in person. Commission is per-platform (Rates).
 */
function computeMoney_(source, checkedIn, isPrivate, income, rates) {
  const paid = isPaidSource_(source);
  const inc = Number(income || 0);
  const ppl = Number(checkedIn || 0);

  if (isPrivate) {
    const weOwe = Number(rates.privatePay || 0);
    return { weOwe, theyOwe: 0, rrMakes: round2_(inc - weOwe), type: 'Private' };
  }
  if (paid) {
    const weOwe = round2_(ppl * rates.paid);
    return { weOwe, theyOwe: 0, rrMakes: round2_(inc - weOwe), type: 'Paid' };
  }
  // Free tour: guide owes us `free`/person; platform commission is per-person.
  const commission = freeCommissionFor_(source, rates);
  const theyOwe = round2_(ppl * rates.free);
  return { weOwe: 0, theyOwe, rrMakes: round2_(ppl * (rates.free - commission)), type: 'Free' };
}

function makeLedgerRow_(o) {
  return [
    o.dateKey, o.day, o.timeLabel, o.language, o.bookingName, o.phone, o.source,
    o.guests, Number(o.children || 0), o.checkedIn, o.weOwe, o.theyOwe, o.rrMakes, o.type, o.bookingId,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    String(o.note || '')
  ];
}


/******************************************************
 * 7. TOKENS  (simple signed session)
 ******************************************************/

/**
 * The signing secret for login tokens. It must be strong (the in-code
 * placeholder sits in a PUBLIC repo, so anyone could forge a token with it) and
 * it must live OUTSIDE the repo. This provisions one AUTOMATICALLY on first use:
 * a random secret is generated and stored in Script Properties, so there is
 * nothing for anyone to run by hand. Guides simply re-login once the first time
 * it replaces the placeholder; after that it is stable.
 */
function tokenSecret_() {
  let props;
  try { props = PropertiesService.getScriptProperties(); }
  catch (e) { return PORTAL.TOKEN_SECRET; }              // no properties -> constant

  let p = props.getProperty('TOKEN_SECRET');
  if (p && p.length >= 24 && p.indexOf('CHANGE_ME') !== 0) return p;   // common path, no lock

  // Provision once, under a lock so two first-hits agree on a single value.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    p = props.getProperty('TOKEN_SECRET');                // re-check inside the lock
    if (!p || p.length < 24 || p.indexOf('CHANGE_ME') === 0) {
      p = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      props.setProperty('TOKEN_SECRET', p);
    }
  } catch (e) {
    return (p && p.length >= 24) ? p : PORTAL.TOKEN_SECRET;  // busy: use constant this request
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
  return p;
}

/** Optional manual override: setTokenSecret('a-long-random-string'). Not required —
 *  tokenSecret_() auto-provisions a strong secret on its own. */
function setTokenSecret(secret) {
  const s = String(secret || '');
  if (s.length < 24) throw new Error('Use at least 24 characters.');
  PropertiesService.getScriptProperties().setProperty('TOKEN_SECRET', s);
  return 'TOKEN_SECRET set (' + s.length + ' chars). Guides will re-login once.';
}

function makeToken_(guideName) {
  const exp = Date.now() + PORTAL.TOKEN_TTL_HOURS * 3600 * 1000;
  const payload = Utilities.base64EncodeWebSafe(guideName + '|' + exp);
  return payload + '.' + sign_(payload);
}

function requireToken_(token) {
  if (!token) return null;
  const parts = String(token).split('.');
  if (parts.length !== 2) return null;
  if (sign_(parts[0]) !== parts[1]) return null;
  const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  const [name, exp] = decoded.split('|');
  if (Number(exp) < Date.now()) return null;
  return name;
}

function sign_(s) {
  const raw = Utilities.computeHmacSha256Signature(s, tokenSecret_());
  return Utilities.base64EncodeWebSafe(raw);
}


/******************************************************
 * 8. HELPERS
 ******************************************************/

function childCountFromNote_(note) {
  const m = String(note || '').match(/(\d+)\s*child/i);
  return m ? Number(m[1]) : 0;
}

function infantCountFromNote_(note) {
  const m = String(note || '').match(/(\d+)\s*infant/i);
  return m ? Number(m[1]) : 0;
}

function isPaidSource_(source) {
  const list = PORTAL._paidSources || PORTAL.PAID_SOURCES;
  return list.some(s => s.toLowerCase() === String(source || '').trim().toLowerCase());
}

function sameName_(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

function shiftKey_(dateKey, minutes, language) {
  return dateKey + '|' + minutes + '|' + String(language || '').trim().toLowerCase();
}


function round2_(n) { return Math.round(Number(n || 0) * 100) / 100; }

/** Accepts a Date or a string, returns "yyyy-MM-dd" (local tz) or ''. */
function toDateKey_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  const d = new Date(s);
  if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return '';
}

/** Normalise "11:00 AM" / "5:00 PM" / "17:00" / "11:00" to 24h "H:MM".
 *  ALSO handles real Date objects: Sheets silently converts a time-looking
 *  string in a non-text cell into a Date, which is what made the ledger's
 *  dedupe-by-shift fail and append duplicate check-in rows. */
function normTime24_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return v.getHours() + ':' + String(v.getMinutes()).padStart(2, '0');
  }
  const s = String(v || '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (m) {
    let h = Number(m[1]);
    if (/pm/i.test(m[3]) && h !== 12) h += 12;
    if (/am/i.test(m[3]) && h === 12) h = 0;
    return h + ':' + m[2];
  }
  m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return Number(m[1]) + ':' + m[2];
  return s;
}

function to12h_(t24) {
  const m = String(t24 || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return String(t24 || '');
  let h = Number(m[1]);
  const suf = h >= 12 ? 'PM' : 'AM';
  const h12 = ((h + 11) % 12) + 1;
  return h12 + ':' + m[2] + ' ' + suf;
}

function timeToMinutes_(t24) {
  const m = String(t24 || '').match(/^(\d{1,2}):(\d{2})$/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : -1;
}

/** True once a shift's tour is over (start + 2h — same rule as the booking
 *  system's Done migration; keeps guides from editing check-ins later without
 *  talking to management). */
function shiftIsOver_(dateKey, minutes) {
  const d = new Date(dateKey + 'T00:00:00');
  if (isNaN(d)) return false;
  // A tour stays on the portal until the EVENING of its own day, so management
  // can still check who was prepaid vs free after it ran. (Guides can only tick
  // check-ins, never untick, so keeping it visible is safe.) It disappears once
  // we pass PORTAL.TOUR_VISIBLE_UNTIL_HOUR on the tour's date.
  const cutoff = d.getTime() + (PORTAL.TOUR_VISIBLE_UNTIL_HOUR || 23) * 3600000;
  return Date.now() > cutoff;
}

/**
 * Read a timestamp cell that may be EITHER a text stamp ("2026-07-20 11:57")
 * or a Date (Sheets silently coerces such strings). Returns {date, text} with
 * date=null when unreadable. Used by the health checks.
 */
function readStampCell_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return { date: v, text: Utilities.formatDate(v, 'Europe/Madrid', 'yyyy-MM-dd HH:mm') };
  }
  const s = String(v || '').trim();
  if (!s) return { date: null, text: '' };
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]),
                       Number(m[4]), Number(m[5]), Number(m[6] || 0));
    return { date: isNaN(d) ? null : d, text: s };
  }
  const d2 = new Date(s);
  return { date: isNaN(d2) ? null : d2, text: s };
}

/** True when two timezone names are currently at the same UTC offset. */
function sameUtcOffset_(tzA, tzB) {
  try {
    const now = new Date();
    return Utilities.formatDate(now, tzA, 'Z') === Utilities.formatDate(now, tzB, 'Z');
  } catch (e) { return false; }
}

/** Epoch ms of a shift's start. */
function shiftStartMs_(dateKey, minutes) {
  const d = new Date(dateKey + 'T00:00:00');
  if (isNaN(d)) return 0;
  return d.getTime() + Math.max(0, Number(minutes) || 0) * 60000;
}

/** Unique key of a shift incl. its private index. */
function shiftKeyFull_(s) {
  return s.dateKey + '|' + s.minutes + '|' + String(s.language || '').toLowerCase() +
         '|' + (s.private ? 'P' + (s.privIndex || 1) : 'R');
}

/** guideNameLower -> [{ms, k}] of every assignment in the schedule. */
function buildBusyMap_(schedule) {
  const busy = {};
  (schedule || []).forEach(s => {
    const ms = shiftStartMs_(s.dateKey, s.minutes);
    const k = shiftKeyFull_(s);
    (s.assigned || []).forEach(n => {
      const nk = String(n).trim().toLowerCase();
      if (!nk) return;
      (busy[nk] = busy[nk] || []).push({ ms, k });
    });
  });
  return busy;
}

/**
 * Guides who can take this shift WITHOUT creating an incompatibility:
 * speak the language, active, and no other assigned tour within
 * MIN_SEPARATION_HOURS. This feeds the portal's assign dropdown, so managers
 * are only offered compatible choices by default.
 */
function eligibleGuidesForShift_(shift, busy, guidesByLanguage) {
  const sepMs = ASSIGN_CFG.MIN_SEPARATION_HOURS * 3600000;
  const myKey = shiftKeyFull_(shift);
  const st = shiftStartMs_(shift.dateKey, shift.minutes);
  return (guidesByLanguage[shift.language] || []).filter(n => {
    const b = busy[String(n).trim().toLowerCase()] || [];
    return !b.some(x => x.k !== myKey && Math.abs(x.ms - st) < sepMs);
  });
}

/** yyyy-MM-dd of the Sunday ending the current week (today..Sunday window). */
function weekEndKey_() {
  const now = new Date();
  const dow = (now.getDay() + 6) % 7;          // 0=Mon..6=Sun
  const sunday = new Date(now.getFullYear(), now.getMonth(), now.getDate() + (6 - dow), 12);
  return Utilities.formatDate(sunday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function todayKey_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function addDaysKey_(dateKey, days) {
  const d = new Date(dateKey + 'T12:00:00');
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function prettyDate_(dateKey) {
  const d = new Date(dateKey + 'T12:00:00');
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE, MMM d');
}

function dayNameFromKey_(dateKey) {
  const d = new Date(dateKey + 'T12:00:00');
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEEE');
}


/******************************************************
 * 8b. UNASSIGNED TAB  (paid + free tours with no guide yet)
 *
 * Lists every active booking whose shift has NO assigned guide in the grids,
 * so management can see what still needs staffing. Once a guide is assigned
 * (grid edited or makeSchedule run), the booking drops off this tab and belongs
 * to that guide (their ledger tab + portal). Zero-people Guruwalk is skipped.
 * Set a time trigger on updateUnassignedLedger (e.g. every 30 min).
 ******************************************************/

function updateUnassignedLedger() { rebuildUnassignedLedger_(); }

/**
 * "Unassigned" = TOURS THAT ALREADY RAN WITH NOBODY ASSIGNED.
 *
 * This is a post-mortem alert — guests had a booking for a tour that has now
 * passed and no guide was ever put on it. It is NOT a list of upcoming
 * bookings awaiting assignment (that is simply the normal pipeline, visible
 * in the portal's All tours, and it was pure noise here).
 *
 * Source of truth: the BookingSheet's "Completed Log" (bookings whose tour is
 * over), cross-checked against the schedule grids READ INCLUDING PAST DATES.
 * Lookback is bounded by UNASSIGNED_LOOKBACK_DAYS; older entries age out on
 * the next rebuild, and the grids themselves only retain a couple of weeks.
 *
 * Rebuilt from scratch on every run (it is a derived view), so it can never
 * accumulate duplicates.
 */
const UNASSIGNED_LOOKBACK_DAYS = 14;

/**
 * Which dates each Schedule_<Language> grid actually contains as ROWS.
 * Returns { language: {dateKey: true} }.
 *
 * Needed because the grids only span the generated window (~2 weeks). For a
 * date OUTSIDE that window there is no evidence either way, so a tour on such
 * a date must NOT be reported as "ran with no guide" — absence of a grid row
 * is not absence of a guide.
 */
function readScheduleDateCoverage_() {
  const cover = {};
  control_().getSheets().forEach(sh => {
    const name = sh.getName();
    if (name.indexOf('Schedule_') !== 0) return;
    const language = name.substring('Schedule_'.length).trim();
    if (!language || sh.getLastRow() < 3) return;
    const vals = sh.getDataRange().getDisplayValues();
    const anchor = gridAnchor_(String((vals[0] && vals[0][0]) || ''));
    cover[language] = cover[language] || {};
    for (let r = 2; r < vals.length; r++) {
      const dk = gridLabelToKey_(String(vals[r][0] || '').trim(), anchor);
      if (dk) cover[language][dk] = true;
    }
  });
  return cover;
}

function rebuildUnassignedLedger_() {
  const rates = readRates_();

  // Who was assigned — including dates already in the past.
  const asg = {};
  try {
    readSchedule_({ includePast: true }).forEach(s => {
      const k = shiftKey_(s.dateKey, s.minutes, s.language) +
                (s.private ? '|P' + (s.privIndex || 1) : '|R');
      asg[k] = (asg[k] || []).concat(s.assigned || []);
      // Regular shifts also answer for private groups at the same slot when a
      // grid predates the private-column layout.
      const loose = shiftKey_(s.dateKey, s.minutes, s.language);
      asg[loose] = (asg[loose] || []).concat(s.assigned || []);
    });
  } catch (e) { /* no grids yet: everything will look unassigned, which is safe */ }

  // Only judge dates the grids actually cover (see readScheduleDateCoverage_).
  let coverage = {};
  try { coverage = readScheduleDateCoverage_(); } catch (e) { /* judge nothing */ }

  const cutoff = addDaysKey_(todayKey_(), -UNASSIGNED_LOOKBACK_DAYS);
  const rows = [];
  let skippedNoGrid = 0;

  readCompletedLog_().forEach(b => {
    if (!b.dateKey || b.dateKey < cutoff) return;             // bounded lookback
    // No grid row for that date+language -> no evidence -> do not accuse.
    if (!(coverage[b.language] && coverage[b.language][b.dateKey])) { skippedNoGrid++; return; }
    const minutes = timeToMinutes_(normTime24_(b.time));
    if (!shiftIsOver_(b.dateKey, minutes)) return;            // paranoia: past only
    const guests = Number(b.adults || 0);
    const source = String(b.source || '');
    if (/guruwalk/i.test(source) && guests <= 0) return;      // zero-people Guruwalk
    const isPriv = /privat/i.test(b.notes || '');

    const exact = shiftKey_(b.dateKey, minutes, b.language) +
                  (isPriv ? '|P1' : '|R');
    const loose = shiftKey_(b.dateKey, minutes, b.language);
    const assigned = (asg[exact] || []).concat(asg[loose] || []);
    if (assigned.length) return;                              // a guide ran it

    const m = computeMoney_(source, guests, isPriv, b.income, rates);
    rows.push([
      b.dateKey, dayNameFromKey_(b.dateKey), to12h_(normTime24_(b.time)), b.language,
      source, b.name || '', guests,
      isPriv ? 'Private' : (isPaidSource_(source) ? 'Paid' : 'Free'),
      Number(b.income || 0), m.rrMakes, b.bookingId || ''
    ]);
  });

  rows.sort((a, b) => (a[0] + a[2]).localeCompare(b[0] + b[2]));
  if (skippedNoGrid) {
    console.log('Unassigned audit: ' + skippedNoGrid +
      ' completed booking(s) skipped — their date is outside the schedule grids.');
  }

  const ss = ledgerSS_();
  const sh = ss.getSheetByName('Unassigned') || ss.insertSheet('Unassigned');
  sh.clear();
  const title = 'TOURS THAT RAN WITH NO GUIDE ASSIGNED — last ' +
                UNASSIGNED_LOOKBACK_DAYS + ' days (rebuilt automatically)';
  sh.getRange(1, 1, 1, 11).merge().setValue(title)
    .setFontWeight('bold').setBackground('#fde68a').setFontColor('#7c2d12');
  const header = ['Date', 'Day', 'Time', 'Language', 'Source', 'Booking', 'Guests',
                  'Type', 'OTA income (€)', 'R&R makes (€)', 'Booking ID'];
  sh.getRange(2, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
  if (rows.length) sh.getRange(3, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(2);
  return rows.length;
}

/** Recover the display language from a booking key by matching the "* Tours" tab. */
function tabLanguageForKey_(key) {
  const lang = key.split('|')[2] || '';
  return lang ? lang.charAt(0).toUpperCase() + lang.slice(1) : '';
}

function minutesToTime_(minutes) {
  const m = Number(minutes);
  if (isNaN(m) || m < 0) return '';
  const h = Math.floor(m / 60), mm = m % 60;
  return h + ':' + (mm < 10 ? '0' + mm : mm);
}


/******************************************************
 * 9. ONE-TIME / MANUAL HELPERS
 ******************************************************/

/** Run once from the editor to create the ledger + Rates tab + one tab per guide. */
function setupLedger() {
  const ss = ledgerSS_();
  migrateLedgerChildrenColumn_();
  repairLedgers();
  ensureGuideTabs_(ss);
  ensureQueueTabs_(ss);
  repairQueueTabs();
  setupLedgerControls();
  Logger.log('Ledger ready: ' + ss.getUrl());
}

/**
 * Create one ledger tab per guide listed in Control_v1 -> Guides (if missing),
 * so management sees every guide immediately, not only after their first check-in.
 * Run setupLedger (or add a guide + re-run) to sync.
 */
function ensureGuideTabs_(ss) {
  ss = ss || ledgerSS_();
  const { header, rows } = readGuidesRaw_();
  const cols = guideColumns_(header);
  rows.forEach(row => {
    const g = parseGuideRow_(row, cols);
    if (g.name) guideTab_(ss, g.name);   // creates the tab + header row if absent
  });
}

/** Quick self-test you can run from the editor. */
function debugPortal() {
  Logger.log('Schedule (upcoming): ' + JSON.stringify(readSchedule_().slice(0, 3), null, 2));
  const idx = readBookingsIndex_();
  Logger.log('Booking shift keys: ' + Object.keys(idx).slice(0, 10).join('\n'));
  Logger.log('Rates: ' + JSON.stringify(readRates_()));
}


/******************************************************
 * 10. MANAGEMENT QUEUES
 *
 * Tabs in Guide_Ledger_v1 (management-only spreadsheet):
 *   "Viator No-shows"    completed Viator bookings never checked in
 *   "GYG No-shows"       completed GetYourGuide bookings never checked in
 *   "GuruWalk Check-ins" guide-checked GuruWalk bookings management must
 *                        report on the GuruWalk platform within 48 h of the
 *                        tour start
 *
 * No-show source of truth: the BookingSheet's hidden "Completed Log" tab
 * (written by bookingList_v2 the moment a finished booking leaves the active
 * tabs — "Done Tours" is aggregated and has no booking ids).
 * A booking is a no-show when no guide ledger tab holds a check-in row for
 * its bookingId + date.
 *
 * Idempotent: every queue entry is keyed (Booking ID | date). Existing rows —
 * including the manager's "Done" checkbox and timestamp — are never
 * recreated or overwritten. Set a time trigger:
 *   updateManagementQueues  — every hour
 *   archiveLedgerMonthly    — monthly, 1st, 02:00-03:00
 ******************************************************/

// Queue tab layout: row 1 = clear button, row 2 = headers, row 3+ = entries.
const QUEUE_BUTTON_ROW = 1;
const QUEUE_HEADER_ROW = 2;
const QUEUE_FIRST_DATA_ROW = 3;

const QUEUE_TABS = {
  VIATOR_NOSHOW: 'Viator No-shows',
  GYG_NOSHOW: 'GYG No-shows',
  GURUWALK: 'GuruWalk Check-ins'
};

const NOSHOW_HEADERS = [
  'Tour date', 'Time', 'Language', 'Source', 'Booking ID', 'Guest', 'Adults',
  'Children', 'Guide', 'Private', 'Portal status', 'OTA action done', 'Done at', 'Notes'
];
const GURUWALK_HEADERS = [
  'Tour date', 'Time', 'Language', 'Booking ID', 'Guest', 'Booked', 'Checked-in',
  'Attendance', 'Children', 'Guide', 'Checked-in at', '48h deadline',
  'Reported in GuruWalk', 'Reported at', 'Notes'
];
// 0-based columns used when reading GuruWalk rows.
const GW_BOOKINGID = 3, GW_BOOKED = 5, GW_CHECKEDIN = 6, GW_ATTEND = 7,
      GW_CHILDREN = 8, GW_GUIDE = 9, GW_DEADLINE = 11, GW_REPORTED = 12;

/** All (everyone came), Some (N of M), or None (no-show) for a guru booking. */
function attendanceLabel_(booked, checkedIn) {
  booked = Number(booked || 0); checkedIn = Number(checkedIn || 0);
  if (checkedIn <= 0) return 'None';
  if (checkedIn >= booked) return 'All';
  return 'Some (' + checkedIn + ' of ' + booked + ')';
}

/** Last row that actually holds DATA in column A (ignores stray checkboxes).
 *  For queue tabs, data starts at QUEUE_FIRST_DATA_ROW; anything above is the
 *  clear button + header. */
function lastDataRow_(sh) {
  const vals = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return i + 1;
  }
  return 0;
}

/** Last data row of a QUEUE tab (never less than the header row). */
function lastQueueRow_(sh) {
  const last = lastDataRow_(sh);
  return Math.max(last, QUEUE_HEADER_ROW);
}

/**
 * RUN ONCE if the queue tabs were created before 2026-07-17: removes the
 * stray full-column checkboxes (they inflated getLastRow() to 1000 and made
 * appends land at row 1001). Safe to re-run.
 */
/**
 * RUN ONCE if a guide ledger tab's columns look misaligned (headers not
 * matching the data, e.g. "R&R makes" missing). For every guide tab:
 *   - if the data is still on the OLD 15-column layout (no Children column),
 *     insert the Children column after Guests so data shifts into place;
 *   - then rewrite row 1 to the canonical 16-column header.
 * All writes/reads are positional, so this only fixes the visible header and
 * the one-time Children insertion — money values are never recomputed. Safe
 * to re-run.
 */
function repairLedgers() {
  const ss = ledgerSS_();
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'Rates' || name === 'Unassigned') return;
    if (Object.values(QUEUE_TABS).indexOf(name) !== -1) return;

    const hdr = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
    const looksLikeGuideTab = hdr[0] === 'Date' || hdr.indexOf('Guests') !== -1;
    if (!looksLikeGuideTab) return;

    // Width of actual data (row 2), to detect the pre-Children 15-col layout.
    let dataWidth = 0;
    const dataRows = Math.max(0, lastDataRow_(sh) - 1);
    if (dataRows > 0) {
      const r2 = sh.getRange(2, 1, 1, sh.getMaxColumns()).getValues()[0];
      for (let i = r2.length - 1; i >= 0; i--) {
        if (String(r2[i]).trim() !== '') { dataWidth = i + 1; break; }
      }
    }

    if (dataWidth === 15) {   // pre-Children 15-col layout -> needs Children inserted
      sh.insertColumnAfter(8);                        // after 'Guests'
      sh.getRange(1, 9).setValue('Children');
      if (dataRows > 0) sh.getRange(2, 9, dataRows, 1).setValue(0);
    }

    // Canonical header (fixes any drift such as a missing 'R&R makes').
    sh.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  Logger.log('Ledgers repaired.');
}

function repairQueueTabs() {
  const ss = ledgerSS_();
  ensureQueueTabs_(ss);
  [QUEUE_TABS.VIATOR_NOSHOW, QUEUE_TABS.GYG_NOSHOW, QUEUE_TABS.GURUWALK].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const last = Math.max(QUEUE_HEADER_ROW, lastQueueRow_(sh));
    const max = sh.getMaxRows();
    if (max > last) {
      // Wipe stray checkboxes/content below the real entries.
      sh.getRange(last + 1, 1, max - last, sh.getMaxColumns())
        .clearContent().clearDataValidations();
    }
  });
  Logger.log('Queue tabs repaired.');
}

/**
 * Queue tabs use a 3-part layout, mirroring the Control sheet's phone
 * controls:
 *   row 1  [ Clear button label | checkbox | status ]
 *   row 2  headers
 *   row 3+ entries
 * Ticking the row-1 checkbox clears every entry (used once the information
 * has been entered on GuruWalk / Viator / GetYourGuide). Safe to re-run:
 * an existing tab is upgraded to this layout without losing entries.
 */
function ensureQueueTabs_(ss) {
  ss = ss || ledgerSS_();
  const mk = (name, headers, label) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
    } else if (String(sh.getRange(QUEUE_HEADER_ROW, 1).getValue() || '') !== headers[0]) {
      // Old layout (headers on row 1): push everything down one row.
      sh.insertRowBefore(1);
    }
    // Button row (always rewritten; cheap and self-healing).
    sh.getRange(QUEUE_BUTTON_ROW, 1, 1, 3)
      .setValues([[label, false, 'Tick the box after entering these on the platform']]);
    sh.getRange(QUEUE_BUTTON_ROW, 2).insertCheckboxes();
    sh.getRange(QUEUE_BUTTON_ROW, 1, 1, 3)
      .setFontWeight('bold').setBackground('#fde68a').setFontColor('#7c2d12');
    sh.getRange(QUEUE_HEADER_ROW, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
    sh.setFrozenRows(QUEUE_HEADER_ROW);
    // NOTE: checkboxes are inserted per appended row, never for a whole
    // column — a full-column checkbox makes getLastRow() = max rows.
    return sh;
  };
  mk(QUEUE_TABS.VIATOR_NOSHOW, NOSHOW_HEADERS, 'CLEAR — after marking these no-shows in Viator');
  mk(QUEUE_TABS.GYG_NOSHOW, NOSHOW_HEADERS, 'CLEAR — after marking these no-shows in GetYourGuide');
  mk(QUEUE_TABS.GURUWALK, GURUWALK_HEADERS, 'CLEAR — after reporting these check-ins in GuruWalk');
}


/**
 * RUN ONCE: installs the on-edit trigger on the LEDGER spreadsheet so the
 * row-1 "CLEAR" checkboxes work. (Triggers are per-spreadsheet, and the
 * ledger is a different file from the Control sheet.)
 */
function setupLedgerControls() {
  const ss = ledgerSS_();
  ensureQueueTabs_(ss);
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'handleLedgerEdit');
  if (!exists) {
    ScriptApp.newTrigger('handleLedgerEdit').forSpreadsheet(ss).onEdit().create();
  }
  Logger.log('Ledger queue controls ready: ' + ss.getUrl());
}

/** Installable on-edit handler for the ledger's queue-tab CLEAR buttons. */
function handleLedgerEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const name = sh.getName();
    if (Object.values(QUEUE_TABS).indexOf(name) === -1) return;
    if (e.range.getRow() !== QUEUE_BUTTON_ROW || e.range.getColumn() !== 2) return;
    if (e.range.getValue() !== true) return;

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
      sh.getRange(QUEUE_BUTTON_ROW, 3).setValue('Busy — try again in a moment');
      e.range.setValue(false);
      return;
    }
    try {
      const last = lastQueueRow_(sh);
      const n = Math.max(0, last - QUEUE_HEADER_ROW);
      if (n > 0) {
        sh.getRange(QUEUE_FIRST_DATA_ROW, 1, n, sh.getMaxColumns())
          .clearContent().clearDataValidations();
      }
      sh.getRange(QUEUE_BUTTON_ROW, 3).setValue(
        'Cleared ' + n + ' entr' + (n === 1 ? 'y' : 'ies') + ' — ' +
        Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm'));
    } catch (err) {
      sh.getRange(QUEUE_BUTTON_ROW, 3).setValue('Error: ' + String(err).slice(0, 80));
      console.error('handleLedgerEdit: ' + err);
    } finally {
      e.range.setValue(false);      // script writes never re-fire onEdit
      lock.releaseLock();
    }
  } catch (outer) {
    console.error('handleLedgerEdit outer: ' + outer);
  }
}

/** Main entry point — run on a time trigger (hourly). */
function updateManagementQueues() {
  safeTriggerRun_('updateManagementQueues', updateManagementQueuesCore_);
}

function updateManagementQueuesCore_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    ensureQueueTabs_();
    reconcilePortalFeed_();          // keep the feed's check-ins and the ledger honest
    updateNoShowQueues_();
    updateGuruwalkCheckinQueue_();
    rebuildUnassignedLedger_();
    markHealthEvent_('HB_QUEUES');
    updateControlHealth_();
  } finally {
    lock.releaseLock();
  }
}

/** All check-in keys across every guide tab: "bookingId|dateKey" -> guide. */
function readAllCheckins_() {
  const ss = ledgerSS_();
  const out = {};
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    if (name === 'Rates' || name === 'Unassigned') return;
    if (Object.values(QUEUE_TABS).indexOf(name) !== -1) return;
    if (sh.getLastRow() < 2) return;
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (header[7] !== 'Guests') return;   // not a guide ledger tab
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, LEDGER_HEADERS.length).getValues();
    v.forEach(r => {
      const bookingId = String(r[LEDGER_BOOKINGID_COL] || '').trim();
      const dateKey = toDateKey_(r[0]);
      if (!bookingId || !dateKey) return;
      out[bookingId + '|' + dateKey] = {
        guide: name,
        source: String(r[LEDGER_SOURCE_COL] || ''),
        checkedIn: Number(r[LEDGER_CHECKEDIN_COL] || 0),
        updated: String(r[LEDGER_UPDATED_COL] || ''),
        time: String(r[2] || ''),
        language: String(r[3] || ''),
        booking: String(r[4] || ''),
        guests: Number(r[7] || 0),
        children: Number(r[8] || 0)
      };
    });
  });
  return out;
}

/** Completed bookings from the BookingSheet's hidden Completed Log tab. */
function readCompletedLog_() {
  const out = [];
  try {
    const sh = bookingSS_().getSheetByName('Completed Log');
    if (!sh || sh.getLastRow() < 2) return out;
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
    v.forEach(r => {
      const bookingId = String(r[9] || '').trim();
      const dateKey = toDateKey_(r[0]);
      if (!bookingId || !dateKey) return;
      out.push({
        // normTime24_ handles a Date-coerced cell; String() would produce
        // "Sat Dec 30 1899 11:00:00 GMT…" and silently break shift matching.
        dateKey, time: normTime24_(r[1]), language: String(r[2] || ''),
        name: String(r[3] || ''), phone: String(r[4] || ''),
        adults: Number(r[5] || 0), children: Number(r[6] || 0),
        source: String(r[7] || ''), income: Number(r[8] || 0),
        bookingId, notes: String(r[10] || '')
      });
    });
  } catch (e) { /* BookingSheet unreachable: skip this cycle */ }
  return out;
}

/** Which guide was assigned to a given completed booking's shift. */
function guideForShift_(schedule, dateKey, time, language, isPrivate) {
  const minutes = timeToMinutes_(normTime24_(time));
  const hit = schedule.find(s =>
    s.dateKey === dateKey && s.minutes === minutes &&
    sameName_(s.language, language) && !!s.private === !!isPrivate);
  return hit && hit.assigned.length ? hit.assigned.join(', ') : '';
}

function updateNoShowQueues_() {
  const ss = ledgerSS_();
  const checkins = readAllCheckins_();
  const completed = readCompletedLog_();
  let schedule = [];
  try { schedule = readSchedule_(); } catch (e) { /* guide column left blank */ }

  const targets = {
    'viator': ss.getSheetByName(QUEUE_TABS.VIATOR_NOSHOW),
    'getyourguide': ss.getSheetByName(QUEUE_TABS.GYG_NOSHOW)
  };

  // Existing keys so completed/pending rows are never recreated.
  const existing = {};
  Object.keys(targets).forEach(k => {
    const sh = targets[k];
    existing[k] = new Set();
    if (!sh) return;
    const last = lastQueueRow_(sh);
    if (last < QUEUE_FIRST_DATA_ROW) return;
    const v = sh.getRange(QUEUE_FIRST_DATA_ROW, 1, last - QUEUE_HEADER_ROW, NOSHOW_HEADERS.length).getValues();
    v.forEach(r => {
      if (String(r[4] || '').trim()) existing[k].add(String(r[4] || '') + '|' + toDateKey_(r[0]));
    });
  });

  const newRows = { 'viator': [], 'getyourguide': [] };
  completed.forEach(b => {
    const srcKey = String(b.source || '').trim().toLowerCase();
    if (!targets[srcKey]) return;                       // only Viator + GYG queues
    const key = b.bookingId + '|' + b.dateKey;
    if (checkins[key]) return;                          // was checked in -> not a no-show
    if (existing[srcKey].has(key)) return;              // already queued
    existing[srcKey].add(key);
    const isPriv = /privat/i.test(b.notes || '');
    newRows[srcKey].push([
      b.dateKey, b.time, b.language, b.source, b.bookingId, b.name,
      b.adults, b.children,
      guideForShift_(schedule, b.dateKey, b.time, b.language, isPriv),
      isPriv ? 'Yes' : '', 'Not checked in', false, '', ''
    ]);
  });

  Object.keys(newRows).forEach(k => {
    const rows = newRows[k], sh = targets[k];
    if (!sh || !rows.length) return;
    const start = Math.max(QUEUE_FIRST_DATA_ROW, lastQueueRow_(sh) + 1);
    sh.getRange(start, 1, rows.length, NOSHOW_HEADERS.length).setValues(rows);
    sh.getRange(start, 12, rows.length, 1).insertCheckboxes();
  });
}

function updateGuruwalkCheckinQueue_() {
  const ss = ledgerSS_();
  ensureQueueTabs_(ss);
  const sh = ss.getSheetByName(QUEUE_TABS.GURUWALK);
  const checkins = readAllCheckins_();
  const completed = readCompletedLog_();

  // Already-queued keys (skip blanks / phantom rows).
  const existing = new Set();
  const lastG = lastQueueRow_(sh);
  if (lastG >= QUEUE_FIRST_DATA_ROW) {
    const v = sh.getRange(QUEUE_FIRST_DATA_ROW, 1, lastG - QUEUE_HEADER_ROW, GURUWALK_HEADERS.length).getValues();
    v.forEach(r => {
      if (String(r[GW_BOOKINGID] || '').trim()) existing.add(String(r[GW_BOOKINGID]) + '|' + toDateKey_(r[0]));
    });
  }

  const rows = [];
  const add = (dateKey, time, language, bookingId, guest, booked, checkedIn, children, guide, checkedAt) => {
    const key = bookingId + '|' + dateKey;
    if (!bookingId || existing.has(key)) return;
    existing.add(key);
    rows.push([
      dateKey, time, language, bookingId, guest, booked, checkedIn,
      attendanceLabel_(booked, checkedIn), children, guide, checkedAt,
      guruwalkDeadline_(dateKey, time), false, '', ''
    ]);
  };

  // 1. Guide-checked GuruWalk bookings -> All / Some.
  Object.keys(checkins).forEach(key => {
    const c = checkins[key];
    if (!/guruwalk/i.test(c.source)) return;
    const parts = key.split('|');
    add(parts[1], c.time, c.language, parts[0], c.booking, c.guests, c.checkedIn, c.children, c.guide, c.updated);
  });

  // 2. Completed GuruWalk bookings with NO check-in -> None (no-show).
  //    Managers still report these (mark as no-show so no commission is owed).
  let schedule = [];
  try { schedule = readSchedule_(); } catch (e) { /* guide left blank */ }
  completed.forEach(b => {
    if (!/guruwalk/i.test(b.source)) return;
    if (checkins[b.bookingId + '|' + b.dateKey]) return;   // handled in pass 1
    const isPriv = /privat/i.test(b.notes || '');
    add(b.dateKey, b.time, b.language, b.bookingId, b.name, b.adults, 0, b.children,
        guideForShift_(schedule, b.dateKey, b.time, b.language, isPriv), '');
  });

  if (rows.length) {
    const start = Math.max(QUEUE_FIRST_DATA_ROW, lastQueueRow_(sh) + 1);
    sh.getRange(start, 1, rows.length, GURUWALK_HEADERS.length).setValues(rows);
    sh.getRange(start, GW_REPORTED + 1, rows.length, 1).insertCheckboxes();
  }
}

/** Tour start + 48 h, formatted for managers. */
function guruwalkDeadline_(dateKey, timeLabel) {
  const minutes = timeToMinutes_(normTime24_(timeLabel));
  const d = new Date(dateKey + 'T12:00:00');
  d.setHours(0, 0, 0, 0);
  const start = new Date(d.getTime() + Math.max(0, minutes) * 60000);
  const deadline = new Date(start.getTime() + 48 * 3600000);
  return Utilities.formatDate(deadline, 'Europe/Madrid', 'yyyy-MM-dd HH:mm');
}

/******************************************************
 * 10B. MANAGER HEALTH DASHBOARD  (Control tab, A1:B14)
 *
 * A phone-glanceable status block. Refreshed by updateManagementQueues
 * (hourly) and by makeSchedule. Timestamps come from script
 * properties written by the functions themselves; counters are recomputed
 * live. The Mobile Controls block lives at N2:P12 on the same tab.
 ******************************************************/

function markHealthEvent_(key) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      key, Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm'));
  } catch (e) { /* ignore */ }
}

function updateControlHealth_() {
  try {
    const control = control_();
    let sh = control.getSheetByName('Control');
    if (!sh) sh = control.insertSheet('Control');
    const props = PropertiesService.getScriptProperties();

    // Booking heartbeat straight from the BookingSheet's Status tab.
    let bookingBeat = '(no Status tab)';
    try {
      const st = bookingSS_().getSheetByName('Status');
      if (st) bookingBeat = readStampCell_(st.getRange(1, 2).getValue()).text || '(empty)';
    } catch (e) { bookingBeat = 'BookingSheet unreachable'; }

    // Live counters.
    let unassigned = 0, pendingGuru = 0, pendingNoShows = 0, openErrors = 0;
    try {
      // Rows on the Unassigned tab = tours that already RAN with no guide.
      const u = ledgerSS_().getSheetByName('Unassigned');
      if (u) unassigned = Math.max(0, lastDataRow_(u) - 2);   // title + header rows
    } catch (e) { /* leave 0 */ }
    try {
      const ss = ledgerSS_();
      const g = ss.getSheetByName(QUEUE_TABS.GURUWALK);
      if (g) {
        const lastG = lastQueueRow_(g);
        if (lastG >= QUEUE_FIRST_DATA_ROW) {
          g.getRange(QUEUE_FIRST_DATA_ROW, 1, lastG - QUEUE_HEADER_ROW, GURUWALK_HEADERS.length).getValues()
            .forEach(r => { if (String(r[GW_BOOKINGID] || '').trim() && r[GW_REPORTED] !== true) pendingGuru++; });
        }
      }
      [QUEUE_TABS.VIATOR_NOSHOW, QUEUE_TABS.GYG_NOSHOW].forEach(name => {
        const q = ss.getSheetByName(name);
        if (!q) return;
        const lastQ = lastQueueRow_(q);
        if (lastQ < QUEUE_FIRST_DATA_ROW) return;
        q.getRange(QUEUE_FIRST_DATA_ROW, 1, lastQ - QUEUE_HEADER_ROW, NOSHOW_HEADERS.length).getValues()
          .forEach(r => { if (String(r[4] || '').trim() && r[11] !== true) pendingNoShows++; });
      });
    } catch (e) { /* leave 0 */ }
    try {
      const errSh = control.getSheetByName('Errors');
      if (errSh && errSh.getLastRow() > 1) {
        const n = Math.min(50, errSh.getLastRow() - 1);
        const cutoff = Date.now() - 48 * 3600000;
        errSh.getRange(errSh.getLastRow() - n + 1, 1, n, 1).getValues()
          .forEach(r => { if (r[0] instanceof Date && r[0].getTime() > cutoff) openErrors++; });
      }
    } catch (e) { /* leave 0 */ }

    const rows = [
      ['SYSTEM HEALTH', 'Updated ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm')],
      ['Booking system last run', bookingBeat],
      ['Last schedule generation', props.getProperty('HB_SCHEDULE') || '(never)'],
      ['Last queue/ledger refresh', props.getProperty('HB_QUEUES') || '(never)'],
      ['Last daily self-test', props.getProperty('HB_SELFTEST') || '(never)'],
      ['Tours that ran with NO guide (14d)', unassigned],
      ['Pending GuruWalk check-ins', pendingGuru],
      ['Pending OTA no-shows', pendingNoShows],
      ['Schedule errors (last 48h)', openErrors],
      ['', ''],
      ['How to read this', 'All timestamps should be recent. Non-zero pending ' +
        'counts = open the Guide_Ledger_v1 queue tabs. Errors = Control sheet ' +
        'Errors tab. Full diagnosis: systemStatus (BookingSheet editor).']
    ];
    // Health block sits BELOW the functions block (which is top-left at A1),
    // separated by one gap row. Position tracks the functions block size.
    const healthFirstRow = MC.FIRST_ACTION_ROW + mcActions_().length + 1;
    // Clear everything from the gap row down to the last used row (A:C), so any
    // stray/duplicate health block left below by an earlier layout is removed.
    // Rows above (the functions block) are never touched.
    const clearTo = Math.max(sh.getLastRow(), healthFirstRow + rows.length + 2);
    sh.getRange(healthFirstRow - 1, 1, clearTo - healthFirstRow + 2, 3).clearContent();
    sh.getRange(healthFirstRow, 1, rows.length, 2).setValues(rows);
    sh.getRange(healthFirstRow, 1, 1, 2).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
    sh.getRange(healthFirstRow + 1, 1, rows.length - 1, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 300);
    sh.setColumnWidth(2, 150);
  } catch (e) { console.log('updateControlHealth_: ' + e); }
}


/******************************************************
 * 11. MONTHLY LEDGER ARCHIVE
 *
 * Time trigger: monthly, day 1, 02:00-03:00.
 * Copies Guide_Ledger_v1 to "YYYY_MM_Guide_Ledger_v1" (previous month) in the
 * same folder, then clears data rows from every tab except Rates so the live
 * ledger starts the month empty. Idempotent via LAST_LEDGER_ARCHIVE property.
 ******************************************************/

function archiveLedgerMonthly() {
  safeTriggerRun_('archiveLedgerMonthly', archiveLedgerMonthlyCore_);
}

function archiveLedgerMonthlyCore_() {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 15);
  const stamp = Utilities.formatDate(prev, 'Europe/Madrid', 'yyyy_MM');
  if (props.getProperty('LAST_LEDGER_ARCHIVE') === stamp) return;   // already archived

  const ss = ledgerSS_();
  const file = DriveApp.getFileById(ss.getId());
  const name = stamp + '_' + PORTAL.LEDGER_NAME;
  let folder;
  try { folder = DriveApp.getFolderById(PORTAL.LEDGER_FOLDER_ID); }
  catch (e) { folder = DriveApp.getRootFolder(); }
  file.makeCopy(name, folder);

  // Clear data rows (keep headers) everywhere except Rates.
  ss.getSheets().forEach(sh => {
    const nm = sh.getName();
    if (nm === 'Rates') return;
    // Queue tabs keep their button row + headers; guide tabs keep headers.
    const firstData = Object.values(QUEUE_TABS).indexOf(nm) !== -1 ? QUEUE_FIRST_DATA_ROW : 2;
    if (sh.getLastRow() >= firstData) {
      sh.getRange(firstData, 1, sh.getLastRow() - firstData + 1, sh.getMaxColumns()).clearContent();
    }
  });

  props.setProperty('LAST_LEDGER_ARCHIVE', stamp);
  // No email: the Control health dashboard shows the last archive instead.
  markHealthEvent_('HB_ARCHIVE');
  Logger.log('Ledger archived to "' + name + '" and cleared for the new month.');
}


/******************************************************
 * 12. QUEUE ACCEPTANCE TEST (idempotency)
 *
 * Runs updateManagementQueues twice and verifies the second run adds no rows.
 ******************************************************/

function testQueueIdempotency() {
  const ss = ledgerSS_();
  ensureQueueTabs_(ss);
  const count = () => [QUEUE_TABS.VIATOR_NOSHOW, QUEUE_TABS.GYG_NOSHOW, QUEUE_TABS.GURUWALK, 'Unassigned']
    .map(n => { const s = ss.getSheetByName(n); return s ? lastDataRow_(s) : 0; });

  updateManagementQueues();
  const first = count();
  updateManagementQueues();
  const second = count();

  const stable = JSON.stringify(first) === JSON.stringify(second);
  console.log('Queue rows after run 1: ' + first + ' | after run 2: ' + second +
              ' -> ' + (stable ? 'PASS (idempotent)' : 'FAIL (duplicates added)'));
  return stable;
}


/******************************************************
 * 13. FULL SYSTEM TEST  (read-only, safe any time)
 *
 * Run fullSystemTest() from the Control project editor. It checks every
 * component end to end and prints a PASS/FAIL/WARN report. It writes
 * NOTHING except the health dashboard, so it is safe to run in production
 * at any moment, including mid-tour.
 ******************************************************/

function fullSystemTest() {
  const R = [];
  const ok   = (m, d) => R.push(['PASS', m, d || '']);
  const bad  = (m, d) => R.push(['FAIL', m, d || '']);
  const warn = (m, d) => R.push(['WARN', m, d || '']);

  /* --- 1. Configuration --- */
  try {
    const tz = Session.getScriptTimeZone();
    if (tz === 'Europe/Madrid') ok('Timezone Europe/Madrid');
    else if (sameUtcOffset_(tz, 'Europe/Madrid')) {
      warn('Timezone is ' + tz + ' (same clock as Madrid, no data impact)',
           'set Project Settings > Time zone to Europe/Madrid for consistency');
    } else {
      bad('Timezone is ' + tz + ' — different offset to Madrid',
          'Project Settings > Time zone must be Europe/Madrid');
    }
  } catch (e) { bad('Timezone unreadable', e); }

  const props = PropertiesService.getScriptProperties();
  ['LEDGER_ID', 'BOOKING_WEBAPP_URL', 'ADMIN_KEY'].forEach(k => {
    props.getProperty(k) ? ok('Script property ' + k + ' set')
                         : warn('Script property ' + k + ' MISSING',
                                k === 'LEDGER_ID' ? 'run setupLedger' : 'phone booking controls will not work');
  });
  if (tokenSecret_().indexOf('CHANGE_ME') === 0) {
    warn('TOKEN_SECRET is still the placeholder — login tokens are forgeable',
         'run setTokenSecret(\'a-long-random-string\') so the key lives outside the public repo');
  } else { ok('TOKEN_SECRET customised (from Script Properties)'); }

  /* --- 2. Spreadsheets + tabs --- */
  let control, booking, ledger;
  try { control = control_(); ok('Control sheet reachable'); } catch (e) { bad('Control sheet unreachable', e); }
  try { booking = bookingSS_(); ok('BookingSheet reachable'); } catch (e) { bad('BookingSheet unreachable', e); }
  try { ledger = ledgerSS_(); ok('Ledger reachable'); } catch (e) { bad('Ledger unreachable', e); }

  if (control) {
    ['Guides', 'Weekly_Schedule', 'Control'].forEach(t =>
      control.getSheetByName(t) ? ok('Control tab "' + t + '"') : bad('Control tab "' + t + '" MISSING'));
    const grids = control.getSheets().filter(s => s.getName().indexOf('Schedule_') === 0);
    grids.length ? ok('Schedule grids: ' + grids.map(s => s.getName().substring(9)).join(', '))
                 : bad('No Schedule_<Language> grids', 'run makeSchedule');
  }
  if (booking) {
    ['English Tours', 'German Tours', 'Spanish Tours', 'Italian Tours', 'French Tours', 'Done Tours', 'Errors', 'Status']
      .forEach(t => booking.getSheetByName(t) ? ok('BookingSheet tab "' + t + '"')
                                              : warn('BookingSheet tab "' + t + '" missing'));
    booking.getSheetByName('Completed Log') ? ok('Completed Log present')
      : warn('Completed Log missing', 'created the first time a tour completes');
  }
  if (ledger) {
    Object.values(QUEUE_TABS).forEach(t => {
      const sh = ledger.getSheetByName(t);
      if (!sh) { bad('Queue tab "' + t + '" missing', 'run setupLedger'); return; }
      const btn = String(sh.getRange(QUEUE_BUTTON_ROW, 1).getValue() || '');
      /^CLEAR/.test(btn) ? ok('Queue tab "' + t + '" has its CLEAR button')
                         : bad('Queue tab "' + t + '" missing CLEAR button', 'run setupLedger');
    });
  }

  /* --- 3. Booking system heartbeat --- */
  try {
    const st = booking && booking.getSheetByName('Status');
    const hb = st ? readStampCell_(st.getRange(1, 2).getValue()) : { date: null, text: '' };
    if (!hb.text) { warn('No booking heartbeat yet', 'has runBookingSystem run?'); }
    else if (!hb.date) { warn('Booking heartbeat unreadable', hb.text); }
    else {
      const mins = Math.round((Date.now() - hb.date.getTime()) / 60000);
      mins <= 15 ? ok('Booking system ran ' + mins + ' min ago')
                 : bad('Booking system last ran ' + mins + ' min ago (' + hb.text + ')',
                       'check its 5-minute trigger / Executions');
    }
  } catch (e) { warn('Heartbeat check failed', e); }

  /* --- 4. Guides + language coverage --- */
  try {
    const raw = readGuidesRaw_();
    const cols = guideColumns_(raw.header);
    const guides = raw.rows.map(r => parseGuideRow_(r, cols)).filter(g => g.name);
    const active = guides.filter(g => g.active);
    active.length ? ok(active.length + ' active guides', active.map(g => g.name).join(', '))
                  : bad('No active guides in the Guides tab');
    guides.filter(g => g.manager).length ? ok('Manager account(s) present')
                                         : warn('No guide flagged as Manager', 'portal assign/move needs one');
    guides.filter(g => g.active && !g.email).forEach(g => warn(g.name + ' has no portal email'));
    cols.languages.forEach(l => {
      const n = active.filter(g => g.languages[l.name] === true).length;
      n ? ok('Language ' + l.name + ': ' + n + ' guide(s)')
        : warn('Language ' + l.name + ': NO active guide');
    });
  } catch (e) { bad('Guides tab unreadable', e); }

  /* --- 5. Schedule integrity --- */
  let schedule = [];
  try {
    schedule = readSchedule_();
    ok('Portal reads ' + schedule.length + ' upcoming shift(s)');
    const dupes = {};
    schedule.forEach(s => { const k = shiftKeyFull_(s); dupes[k] = (dupes[k] || 0) + 1; });
    const dup = Object.keys(dupes).filter(k => dupes[k] > 1);
    dup.length ? bad(dup.length + ' duplicate shift(s) in the grids', dup.join(' | '))
               : ok('No duplicate shifts');

    // Overlaps and language eligibility across the whole live schedule.
    const busy = buildBusyMap_(schedule);
    const sep = ASSIGN_CFG.MIN_SEPARATION_HOURS * 3600000;
    let overlaps = 0;
    Object.keys(busy).forEach(g => {
      const list = busy[g].slice().sort((a, b) => a.ms - b.ms);
      for (let i = 1; i < list.length; i++) {
        if (list[i].ms - list[i - 1].ms < sep) overlaps++;
      }
    });
    overlaps ? bad(overlaps + ' overlapping assignment(s) (<' + ASSIGN_CFG.MIN_SEPARATION_HOURS + 'h)',
                   'run validateScheduleGrids for detail')
             : ok('No overlapping guide assignments');

    let wrongLang = 0;
    schedule.forEach(s => (s.assigned || []).forEach(n => {
      const g = findGuideByName_(n);
      if (g && g.languages[s.language] !== true) {
        wrongLang++;
        bad(n + ' assigned to ' + s.language + ' on ' + s.dateKey, 'does not speak it');
      }
    }));
    if (!wrongLang) ok('Every assigned guide speaks the tour language');

    // Only the CURRENT scheduling window is actionable — tours further out
    // are staffed by the Friday run as the window rolls forward.
    const horizon = formatDate_(endOfScheduleRange_());
    const unassigned = schedule.filter(s => !(s.assigned || []).length && s.dateKey <= horizon);
    const laterUn = schedule.filter(s => !(s.assigned || []).length && s.dateKey > horizon).length;
    unassigned.length ? warn(unassigned.length + ' tour(s) unassigned INSIDE the scheduling window (to ' + horizon + ')',
                             unassigned.slice(0, 5).map(s => s.dateKey + ' ' + s.timeLabel + ' ' + s.language).join(' | '))
                      : ok('All tours in the scheduling window have a guide');
    if (laterUn) ok(laterUn + ' unassigned tour(s) beyond ' + horizon + ' (normal — scheduled on Friday)');
  } catch (e) { bad('Schedule read failed', e); }

  /* --- 6. Bookings vs schedule --- */
  try {
    const idx = readBookingsIndex_();
    const keys = Object.keys(idx);
    const total = keys.reduce((n, k) => n + idx[k].length, 0);
    ok(total + ' active booking(s) across ' + keys.length + ' shift(s)');

    const ids = {};
    let dupes = 0;
    keys.forEach(k => idx[k].forEach(b => {
      if (!b.bookingId) return;
      if (ids[b.bookingId]) { dupes++; bad('Duplicate booking id ' + b.bookingId); }
      ids[b.bookingId] = true;
    }));
    if (!dupes) ok('No duplicate booking ids');

    const schedKeys = {};
    schedule.forEach(s => { schedKeys[shiftKey_(s.dateKey, s.minutes, s.language)] = true; });
    // A booking beyond the scheduling window has no grid row YET, by design.
    const horizon2 = formatDate_(endOfScheduleRange_());
    const orphans = keys.filter(k => !schedKeys[k] && k.split('|')[0] <= horizon2);
    const later = keys.filter(k => !schedKeys[k] && k.split('|')[0] > horizon2).length;
    orphans.length ? warn(orphans.length + ' booked shift(s) INSIDE the window with no schedule row',
                          'run makeSchedule — they should appear as extra tour columns')
                   : ok('Every booked shift inside the scheduling window is in the grids');
    if (later) ok(later + ' booked shift(s) beyond ' + horizon2 + ' (normal — not scheduled yet)');
  } catch (e) { bad('Booking index failed', e); }

  /* --- 7. Ledger + queues --- */
  try {
    const ck = readAllCheckins_();
    ok(Object.keys(ck).length + ' check-in record(s) in the ledger');
    const rates = readRates_();
    ok('Rates: paid ' + rates.paid + '€, free ' + rates.free + '€, private ' + rates.privatePay + '€');
    const u = ledger && ledger.getSheetByName('Unassigned');
    const n = u ? Math.max(0, lastDataRow_(u) - 2) : 0;
    n ? warn(n + ' tour(s) RAN with no guide assigned', 'see the Unassigned tab')
      : ok('No tours ran unassigned');
  } catch (e) { bad('Ledger check failed', e); }

  /* --- 8. Triggers --- */
  try {
    const fns = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
    ['runWeeklyScheduling', 'updateManagementQueues', 'archiveLedgerMonthly',
     'handleMobileControlsEdit', 'handleScheduleEdit', 'handleLedgerEdit'].forEach(f => {
      fns.indexOf(f) !== -1 ? ok('Trigger installed: ' + f)
                            : bad('Trigger MISSING: ' + f, 'see the deployment checklist');
    });
    ['dailyScheduleSelfTest', 'sendGuruwalkCheckinReminder'].forEach(f => {
      if (fns.indexOf(f) !== -1) bad('Obsolete trigger still installed: ' + f, 'delete it');
    });
  } catch (e) { warn('Trigger check failed', e); }

  /* --- report --- */
  try { updateControlHealth_(); } catch (e) { /* best effort */ }
  const fails = R.filter(r => r[0] === 'FAIL');
  const warns = R.filter(r => r[0] === 'WARN');
  console.log('================ FULL SYSTEM TEST ================');
  R.forEach(r => console.log(r[0].padEnd(5) + ' ' + r[1] + (r[2] ? '  — ' + r[2] : '')));
  console.log('=================================================');
  console.log('PASS ' + R.filter(r => r[0] === 'PASS').length +
              ' | WARN ' + warns.length + ' | FAIL ' + fails.length);
  console.log(fails.length ? 'ACTION NEEDED — fix the FAIL lines above.'
                           : (warns.length ? 'Healthy. WARNs are informational.' : 'All green.'));
  return { pass: R.filter(r => r[0] === 'PASS').length, warn: warns.length, fail: fails.length };
}


/******************************************************
 * LIVE SELF-TEST — proves the assign + check-in flow end to end.
 *
 * Run it from the Apps Script editor: Run > runPortalSelfTest. It creates
 * throwaway "Testing <Language>" bookings on a FAKE time slot (13:37, which no
 * real tour uses, so it can never collide with or overwrite real data), then
 * drives the real API exactly as a phone would:
 *   booking surfaces as a shift  ->  manager assigns a guide  ->  the
 *   assignment STICKS on reload with NO duplicate row  ->  a check-in saves
 *   and PERSISTS on reload.
 * Every step is timed. Results land in a "Self Test" tab (Step | ms | Result |
 * Detail) and in the execution log, and it cleans up after itself.
 * Pass { keep: true } to leave the Testing data in place for inspection.
 ******************************************************/
function runPortalSelfTest(opts) {
  const keep = !!(opts && opts.keep);
  const ST = { time24: '13:37', timeLabel: '1:37 PM', langs: ['English', 'Italian'] };
  const R = [];
  const step = (name, fn) => {
    const t0 = Date.now();
    try { const d = fn(); R.push({ name: name, ms: Date.now() - t0, ok: true, detail: d || '' }); return d; }
    catch (e) { R.push({ name: name, ms: Date.now() - t0, ok: false, detail: String(e && e.message ? e.message : e) }); return null; }
  };

  // A real manager + token, so we exercise the exact path the browser uses.
  let managerName = null;
  try {
    const raw = readGuidesRaw_(); const cols = guideColumns_(raw.header);
    raw.rows.forEach(row => {
      const g = parseGuideRow_(row, cols);
      if (!managerName && g.name && g.active && g.manager) managerName = g.name;
    });
  } catch (e) { /* handled below */ }
  if (!managerName) {
    R.push({ name: 'find an active manager in Guides', ms: 0, ok: false, detail: 'none found' });
    return selfTestReport_(R);
  }
  const token = makeToken_(managerName);
  const dateKey = addDaysKey_(todayKey_(), 2);           // upcoming, in-window
  const minutes = timeToMinutes_(ST.time24);
  const guidesByLang = (apiTours_({ token: token }).guidesByLanguage) || {};

  try {
    ST.langs.forEach((lang, li) => {
      const guide = (guidesByLang[lang] || []).find(n => n && n !== managerName) || (guidesByLang[lang] || [])[0];
      const bid = 'TEST-' + lang.toUpperCase() + '-' + Date.now() + '-' + li;

      step('[' + lang + '] create a Testing booking', () => {
        createTestBookingRow_(lang, dateKey, ST.timeLabel, 'Testing ' + lang, bid, 2);
        return bid;
      });

      step('[' + lang + '] the booking surfaces as a shift at the RIGHT date', () => {
        const t = apiTours_({ token: token });
        const ok = (t.allTours || []).some(s => s.dateKey === dateKey && s.time === ST.time24 &&
          s.language === lang && (s.bookings || []).some(b => b.bookingId === bid));
        if (!ok) throw new Error('booking/shift not visible in the portal');
        return dateKey + ' ' + ST.timeLabel;
      });

      if (!guide) { R.push({ name: '[' + lang + '] assign a guide', ms: 0, ok: false, detail: 'no active guide speaks ' + lang }); return; }

      step('[' + lang + '] manager assigns ' + guide, () => {
        const r = apiAssign_({ token: token, dateKey: dateKey, time: ST.time24, language: lang, guide: guide, force: '1' });
        if (!r.ok) throw new Error(r.error || 'assign failed');
        return guide;
      });

      step('[' + lang + '] assignment STICKS on reload, with NO duplicate row', () => {
        const t = apiTours_({ token: token });
        const shifts = (t.allTours || []).filter(s => s.dateKey === dateKey && s.time === ST.time24 && s.language === lang);
        if (!shifts.length) throw new Error('shift vanished after assign');
        if (shifts.length > 1) throw new Error('DUPLICATE shift: ' + shifts.length + ' rows');
        const names = (shifts[0].assigned || []).concat(shifts[0].guide ? [shifts[0].guide] : []);
        if (!names.some(n => n && n.toLowerCase() === guide.toLowerCase())) {
          throw new Error('assigned to ' + JSON.stringify(names) + ', expected ' + guide);
        }
        const gridRows = readSchedule_({ includePast: true })
          .filter(s => s.dateKey === dateKey && s.time === ST.time24 && s.language === lang);
        if (gridRows.length > 1) throw new Error('grid has ' + gridRows.length + ' duplicate rows for the slot');
        return 'assigned=' + guide;
      });

      step('[' + lang + '] a check-in SAVES', () => {
        const data = {
          dateKey: dateKey, time: ST.time24, timeLabel: ST.timeLabel, day: dayNameFromKey_(dateKey),
          language: lang, guide: guide, walkins: [],
          bookings: [{ bookingId: bid, source: 'Website', name: 'Testing ' + lang, phone: '', guests: 2,
                       income: 0, isPrivate: false, manualNote: '', checked: true, checkedIn: 2 }]
        };
        const r = apiSave_({ token: token, data: JSON.stringify(data) });
        if (!r.ok) throw new Error(r.error || 'save failed');
        return 'checkedIn=2';
      });

      step('[' + lang + '] the check-in PERSISTS on reload', () => {
        const ck = readGuideCheckins_(guide);
        const kk = shiftKey_(dateKey, minutes, lang) + '|' + bid;
        const cke = ck[kk];
        if (!cke) throw new Error('no ledger check-in stored');
        if (Number(cke.n) !== 2) throw new Error('checkedIn=' + JSON.stringify(cke) + ', expected 2');
        return 'persisted at ' + (cke.at || '?');
      });
    });
  } finally {
    if (!keep) step('clean up all Testing data', () => deleteSelfTestArtifacts_(ST));
  }

  return selfTestReport_(R);
}

/** Append one throwaway Testing booking row to a BookingSheet language tab. */
function createTestBookingRow_(language, dateKey, timeLabel, name, id, guests) {
  const sh = bookingSS_().getSheetByName(language + PORTAL.BOOKING_TAB_SUFFIX);
  if (!sh) throw new Error(language + ' Tours tab not found');
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 2, 1, 1).setNumberFormat('@');   // phone as text
  sh.getRange(row, 5, 1, 1).setNumberFormat('@');   // time as text
  sh.getRange(row, 8, 1, 1).setNumberFormat('@');   // booking id as text
  sh.getRange(row, 1, 1, 9).setValues([[
    name, '', Number(guests || 0), new Date(dateKey + 'T12:00:00'), timeLabel, 'Website', 0, id, ''
  ]]);
}

/** Remove every Testing artefact: booking rows, ledger rows, and the fake grid column. */
function deleteSelfTestArtifacts_(ST) {
  const removed = { bookings: 0, ledger: 0, gridCols: 0 };

  bookingSS_().getSheets().forEach(sh => {
    const tab = sh.getName();
    if (tab.indexOf(PORTAL.BOOKING_TAB_SUFFIX) === -1 || /^done\b/i.test(tab)) return;
    const last = sh.getLastRow(); if (last < 2) return;
    const ids = sh.getRange(2, 8, last - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (/^TEST-/i.test(String(ids[i][0] || ''))) { sh.deleteRow(i + 2); removed.bookings++; }
    }
  });

  try {
    ledgerSS_().getSheets().forEach(sh => {
      const last = sh.getLastRow(); if (last < 2) return;
      const hdr = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
      if (hdr[0] !== 'Date' || hdr.indexOf('Booking ID') === -1) return;
      const ids = sh.getRange(2, LEDGER_BOOKINGID_COL + 1, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (/^TEST-/i.test(String(ids[i][0] || ''))) { sh.deleteRow(i + 2); removed.ledger++; }
      }
    });
  } catch (e) { /* ledger optional */ }

  control_().getSheets().forEach(sh => {
    if (sh.getName().indexOf('Schedule_') !== 0) return;
    const lastCol = sh.getLastColumn(); if (lastCol < 2) return;
    const hdr = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
    for (let c = lastCol; c >= 2; c--) {
      const h = parseGridTimeHeader_(hdr[c - 1]);
      if (h && h.time === ST.time24) { sh.deleteColumn(c); removed.gridCols++; }
    }
  });

  return 'bookings=' + removed.bookings + ' ledger=' + removed.ledger + ' gridCols=' + removed.gridCols;
}

/** Write the self-test results to a "Self Test" tab and the log; return a summary. */
function selfTestReport_(R) {
  const fails = R.filter(r => !r.ok);
  const slow = R.filter(r => r.ms > PORTAL.SLOW_MS);
  const totalMs = R.reduce((s, r) => s + r.ms, 0);
  const slowest = R.reduce((m, r) => (r.ms > (m ? m.ms : -1) ? r : m), null);
  try {
    const ss = control_();
    let sh = ss.getSheetByName('Self Test') || ss.insertSheet('Self Test');
    sh.clear();
    sh.getRange(1, 1, 1, 4).setValues([['Step', 'ms', 'Result', 'Detail']]).setFontWeight('bold');
    const rows = R.map(r => [r.name, r.ms,
      r.ok ? (r.ms > PORTAL.SLOW_MS ? 'PASS (SLOW)' : 'PASS') : 'FAIL', String(r.detail || '')]);
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    sh.getRange(rows.length + 3, 1, 1, 4).setValues([[
      'TOTAL', totalMs, fails.length ? (fails.length + ' FAILED') : 'ALL PASS',
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    ]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } catch (e) { /* report tab is best-effort */ }

  console.log('================ PORTAL SELF TEST ================');
  R.forEach(r => console.log((r.ok ? 'PASS' : 'FAIL') + ' ' + String(r.ms).padStart(5) + 'ms  ' +
    r.name + (r.detail ? '  — ' + r.detail : '')));
  console.log('TOTAL ' + totalMs + 'ms | ' + (fails.length ? (fails.length + ' FAILED — see the FAIL lines') : 'ALL PASS') +
    (slowest ? (' | slowest: ' + slowest.name + ' ' + slowest.ms + 'ms') : '') +
    (slow.length ? (' | ' + slow.length + ' step(s) OVER ' + PORTAL.SLOW_MS + 'ms') : ''));
  return { pass: R.length - fails.length, fail: fails.length, totalMs: totalMs, slowMs: PORTAL.SLOW_MS,
           slow: slow.map(s => s.name + ' ' + s.ms + 'ms'), fails: fails.map(f => f.name + ': ' + f.detail) };
}
