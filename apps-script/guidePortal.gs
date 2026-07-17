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
  TOKEN_TTL_HOURS: 12,

  // Which sources are "paid" (we owe the guide). Everything else is "free".
  PAID_SOURCES: ['Viator', 'GetYourGuide', 'Airbnb'],

  // Default rates (€ per checked-in person). The Rates tab overrides these.
  DEFAULT_PAID_RATE: 10,       // paid tours: we owe the guide, € per checked-in person
  DEFAULT_FREE_RATE: 6,        // free tours: the guide owes us, € per checked-in person
  DEFAULT_PRIVATE_PAY: 75,     // private tours: flat € we owe the guide who runs it
  DEFAULT_GURUWALK_FEE: 4.70,  // € Guruwalk charges us, per booking (a cost to R&R)

  // Show tours from today up to this many days ahead.
  UPCOMING_DAYS: 45,

  // BookingSheet tab name pattern per language, e.g. "English Tours".
  BOOKING_TAB_SUFFIX: ' Tours',

  // Control_v1 tab that holds the generated assignments.
  SCHEDULE_TAB: 'Schedule',
  GUIDES_TAB: 'Guides'
};


/******************************************************
 * 2. WEB APP ENTRY POINT (JSONP)
 ******************************************************/

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const callback = p.callback || '';

  let out;
  try {
    switch (p.action) {
      case 'login':  out = apiLogin_(p); break;
      case 'tours':  out = apiTours_(p); break;
      case 'save':   out = apiSave_(p);  break;
      case 'ping':   out = { ok: true, pong: true }; break;
      case 'health': out = apiHealth_(); break;
      default:       out = { ok: false, error: 'Unknown action: ' + String(p.action || '(none)') };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }

  return jsonp_(callback, out);
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

  const rates = readRates_();
  const schedule = readSchedule_();                 // all upcoming shifts
  const mine = schedule.filter(s => s.assigned.some(a => sameName_(a, name)));

  const bookingsByKey = readBookingsIndex_();       // "yyyy-mm-dd|minutes|Language" -> [bookings]
  const priorCheckins = readGuideCheckins_(name);   // key|bookingId -> checkedIn

  const tours = mine.map(shift => {
    const key = shiftKey_(shift.dateKey, shift.minutes, shift.language);
    // A private shift shows only its private booking(s); a regular shift only
    // the non-private ones. That's what splits Fake (regular) from Fake 2 (private).
    const bookings = (bookingsByKey[key] || [])
      .filter(b => shift.private ? /privat/i.test(b.note || '') : !/privat/i.test(b.note || ''))
      .map(b => {
        const kk = key + '|' + b.bookingId;
        const isCk = Object.prototype.hasOwnProperty.call(priorCheckins, kk); // ledger row exists = checked in
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
          checked: isCk,
          checkedIn: isCk ? Number(priorCheckins[kk]) : Number(b.guests || 0) // locked count, or booked default to adjust
        };
      });

    const bookedGuests = bookings.reduce((s, b) => s + Number(b.guests || 0), 0);
    const bookedChildren = bookings.reduce((s, b) => s + Number(b.children || 0), 0);
    const checkedGuests = bookings.reduce((s, b) => s + (b.checked ? Number(b.checkedIn || 0) : 0), 0);

    return {
      id: shift.private ? key + '|P' : key,
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

  // "Schedule of other guides": compact, read-only view of all upcoming shifts.
  const scheduleView = schedule.map(s => ({
    dateKey: s.dateKey, dateText: s.dateText, day: s.day,
    time: s.time, language: s.language, assigned: s.assigned, status: s.status
  }));

  // Managers get the full My-tours-style view of EVERY tour (with bookings +
  // check-ins), and can save on a guide's behalf.
  const me = findGuideByName_(name);
  const isManager = !!(me && me.manager);
  let allTours = [];
  if (isManager) {
    const ckCache = {};
    const getCk = g => (ckCache[g] || (ckCache[g] = readGuideCheckins_(g)));
    allTours = schedule.map(shift => {
      const key = shiftKey_(shift.dateKey, shift.minutes, shift.language);
      const primary = shift.assigned[0] || '';
      const ck = primary ? getCk(primary) : {};
      const bookings = (bookingsByKey[key] || [])
        .filter(b => shift.private ? /privat/i.test(b.note || '') : !/privat/i.test(b.note || ''))
        .map(b => {
          const kk = key + '|' + b.bookingId;
          const isCk = Object.prototype.hasOwnProperty.call(ck, kk);
          return {
            bookingId: b.bookingId, name: b.name, phone: b.phone, source: b.source, guests: b.guests,
            children: Number(b.children || 0), infants: Number(b.infants || 0),
            income: Number(b.income || 0),
            paid: isPaidSource_(b.source), isPrivate: /privat/i.test(b.note || ''),
            checked: isCk,
            checkedIn: isCk ? Number(ck[kk]) : Number(b.guests || 0)
          };
        });
      return {
        id: shift.private ? key + '|P' : key, dateKey: shift.dateKey, dateText: shift.dateText, day: shift.day,
        time: shift.time, timeLabel: shift.timeLabel, language: shift.language,
        assigned: shift.assigned, guide: primary, coGuides: shift.assigned, status: shift.status,
        isPrivate: !!shift.private,
        bookedGuests: bookings.reduce((s, b) => s + Number(b.guests || 0), 0),
        bookedChildren: bookings.reduce((s, b) => s + Number(b.children || 0), 0),
        checkedGuests: bookings.reduce((s, b) => s + (b.checked ? Number(b.checkedIn || 0) : 0), 0),
        bookings
      };
    });
  }

  return { ok: true, guide: name, manager: isManager, rates, tours, schedule: scheduleView, allTours };
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
      bookingId: b.bookingId || ''
    }));
  });

  (d.walkins || []).forEach(w => {
    const count = Math.max(0, Number(w.count || 0));
    if (!count) return;
    const m = computeMoney_(w.source, count, false, 0, rates);
    rows.push(makeLedgerRow_({
      dateKey: d.dateKey, day, timeLabel, language: d.language,
      bookingName: 'Walk-in', phone: '', source: w.source || 'Walk-in',
      guests: count, checkedIn: count,
      weOwe: m.weOwe, theyOwe: m.theyOwe, rrMakes: m.rrMakes, type: m.type + ' (walk-in)',
      bookingId: 'WALKIN|' + (w.source || '') // stable key so re-saves overwrite
    }));
  });

  // LockService: two guides saving at once (or a double-tap) must not
  // interleave ledger writes. writeGuideLedger_ itself replaces the shift's
  // rows, so a repeated identical save is a clean overwrite, not a duplicate.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Server busy, try again in a moment' };
  try {
    writeGuideLedger_(targetGuide, d.dateKey, d.time || timeLabel, d.language, rows);
    // Keep the GuruWalk management queue current without waiting for the trigger.
    try { updateGuruwalkCheckinQueue_(); } catch (e) { /* queue refresh is best-effort */ }
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
    deployment: 'portal-v2'
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
    out.bookingTabs = b.getSheets().filter(s => s.getName().indexOf(PORTAL.BOOKING_TAB_SUFFIX) !== -1).length;
  } catch (e) { out.ok = false; out.bookingSheetOk = false; out.error = 'BookingSheet unreachable'; }
  try {
    out.ledgerOk = !!ledgerSS_();
  } catch (e) { out.ledgerOk = false; }
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
           managerCol: idx('Manager'), languages };
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
function readSchedule_() {
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
      const t = normTime24_(timeRow[c]);
      if (t) times.push({ col: c, time: t });
    }

    for (let r = 2; r < vals.length; r++) {
      const label = String(vals[r][0] || '').trim();
      if (!label) continue;
      const dateKey = gridLabelToKey_(label, anchor);
      if (!dateKey) continue;
      if (dateKey < today || dateKey > maxKey) continue;

      times.forEach(t => {
        const raw = String(vals[r][t.col] || '').trim();
        if (!raw) return;
        const lines = raw.split('\n').map(s => s.trim()).filter(Boolean);
        const privLines = lines.filter(l => /🔒|\(private\)/i.test(l));
        const regLines = lines.filter(l => !/🔒|\(private\)/i.test(l));

        const base = {
          dateKey, dateText: prettyDate_(dateKey), day: dayNameFromKey_(dateKey),
          time: t.time, timeLabel: to12h_(t.time), minutes: timeToMinutes_(t.time),
          language
        };

        if (regLines.length) {
          const names = regLines.filter(l => !/not assigned/i.test(l))
            .join(',').split(',').map(s => s.trim()).filter(Boolean);
          out.push(Object.assign({}, base, {
            private: false, assigned: names,
            status: names.length ? 'OK' : 'Not assigned'
          }));
        }

        if (privLines.length) {
          const names = [];
          privLines.forEach(l => {
            let n = l.replace(/🔒/g, '').replace(/\(private\)/ig, '');
            if (/not assigned/i.test(n)) n = n.replace(/not assigned/ig, '');
            n.split(',').forEach(s => { s = s.trim(); if (s) names.push(s); });
          });
          out.push(Object.assign({}, base, {
            private: true, assigned: names,
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
    const k = shiftKey_(s.dateKey, s.minutes, s.language) + (s.private ? '|P' : '|R');
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

/** Anchor {year, month} from a grid title's first yyyy-MM-dd (for year-boundary safety). */
function gridAnchor_(title) {
  const m = String(title).match(/(20\d{2})-(\d{2})-(\d{2})/);
  if (m) return { year: Number(m[1]), month: Number(m[2]) - 1 };
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() };
}

/** "Thu Jul 16" + anchor -> "2026-07-16" (rolls to next year across Dec->Jan). */
function gridLabelToKey_(label, anchor) {
  const m = String(label).match(/([A-Za-z]{3,})\s+(\d{1,2})\s*$/);
  if (!m) return '';
  const months = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  const mon = months[m[1].slice(0, 3).toLowerCase()];
  if (mon == null) return '';
  const day = Number(m[2]);
  const year = mon < anchor.month ? anchor.year + 1 : anchor.year;
  const d = new Date(year, mon, day, 12, 0, 0);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}


/******************************************************
 * 5. BOOKINGSHEET READER  (guests + contacts per shift)
 ******************************************************/

function bookingSS_() { return SpreadsheetApp.openById(PORTAL.BOOKING_SHEET_ID); }

/** Index all active bookings by "dateKey|minutes|Language". */
function readBookingsIndex_() {
  const ss = bookingSS_();
  const index = {};

  ss.getSheets().forEach(sh => {
    const tab = sh.getName();
    if (tab.indexOf(PORTAL.BOOKING_TAB_SUFFIX) === -1) return; // only "* Tours" tabs
    const language = tab.replace(PORTAL.BOOKING_TAB_SUFFIX, '').trim();
    if (sh.getLastRow() < 2) return;

    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
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
        note                                // "Private" flag comes through here
      });
    });
  });

  return index;
}


/******************************************************
 * 6. LEDGER  (Guide_Ledger_v1: one tab per guide + Rates)
 ******************************************************/

const LEDGER_HEADERS = [
  'Date', 'Day', 'Time', 'Language', 'Booking', 'Phone', 'Source',
  'Guests', 'Children', 'Checked-in', 'We owe guide (€)', 'Guide owes us (€)', 'R&R makes (€)', 'Type', 'Booking ID', 'Updated'
];
const LEDGER_BOOKINGID_COL = 14;   // 0-based index of 'Booking ID' (used when re-reading)
const LEDGER_CHECKEDIN_COL = 9;    // 0-based index of 'Checked-in'
const LEDGER_SOURCE_COL = 6;       // 0-based index of 'Source'

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
  sh.getRange(1, 1, 6, 2).setValues([
    ['Setting', 'Value'],
    ['Paid tour — we owe guide (€ per checked-in person)', PORTAL.DEFAULT_PAID_RATE],
    ['Free tour — guide owes us (€ per checked-in person)', PORTAL.DEFAULT_FREE_RATE],
    ['Private tour — we owe guide (flat € per tour)', PORTAL.DEFAULT_PRIVATE_PAY],
    ['Guruwalk fee we pay (€ per booking)', PORTAL.DEFAULT_GURUWALK_FEE],
    ['Paid sources (comma separated)', PORTAL.PAID_SOURCES.join(', ')]
  ]);
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
  let privatePay = PORTAL.DEFAULT_PRIVATE_PAY, guruwalkFee = PORTAL.DEFAULT_GURUWALK_FEE;
  let paidSources = PORTAL.PAID_SOURCES.slice();
  if (sh && sh.getLastRow() >= 2) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    v.forEach(r => {
      const label = String(r[0] || '').toLowerCase();
      if (label.indexOf('paid tour') === 0) paid = Number(r[1]) || paid;
      else if (label.indexOf('free tour') === 0) free = Number(r[1]) || free;
      else if (label.indexOf('private tour') === 0) privatePay = Number(r[1]) || privatePay;
      else if (label.indexOf('guruwalk fee') === 0) guruwalkFee = Number(r[1]) || guruwalkFee;
      else if (label.indexOf('paid sources') === 0 && r[1]) {
        paidSources = String(r[1]).split(',').map(s => s.trim()).filter(Boolean);
      }
    });
  }
  PORTAL._paidSources = paidSources; // cache for isPaidSource_
  return { paid, free, privatePay, guruwalkFee, paidSources };
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
    out[shiftKey_(dateKey, minutes, language) + '|' + bookingId] = Number(r[LEDGER_CHECKEDIN_COL] || 0);
  });
  return out;
}

/** Upsert this shift's rows for a guide (replace prior rows for the same shift). */
function writeGuideLedger_(name, dateKey, time, language, rows) {
  const ss = ledgerSS_();
  const sh = guideTab_(ss, name);
  const minutes = timeToMinutes_(normTime24_(time));
  const targetKey = shiftKey_(dateKey, minutes, language);

  // Delete existing rows for this exact shift (so re-saving overwrites cleanly).
  if (sh.getLastRow() >= 2) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, LEDGER_HEADERS.length).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      const k = shiftKey_(toDateKey_(v[i][0]), timeToMinutes_(normTime24_(v[i][2])), String(v[i][3] || '').trim());
      if (k === targetKey) sh.deleteRow(i + 2);
    }
  }

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, LEDGER_HEADERS.length).setValues(rows);
  }
}

/**
 * All money for one ledger row.
 *   Private tour  -> we owe the guide a flat privatePay; R&R = OTA income - privatePay.
 *   Paid tour     -> we owe the guide 10 €/checked-in; R&R = OTA income - that.
 *   Free tour     -> guide owes us 6 €/checked-in; R&R = that - Guruwalk fee (Guruwalk only).
 */
function computeMoney_(source, checkedIn, isPrivate, income, rates) {
  const paid = isPaidSource_(source);
  const guruFee = /guruwalk/i.test(String(source || '')) ? Number(rates.guruwalkFee || 0) : 0;
  const inc = Number(income || 0);

  if (isPrivate) {
    const weOwe = Number(rates.privatePay || 0);
    return { weOwe, theyOwe: 0, rrMakes: round2_(inc - weOwe), type: 'Private' };
  }
  if (paid) {
    const weOwe = round2_(checkedIn * rates.paid);
    return { weOwe, theyOwe: 0, rrMakes: round2_(inc - weOwe), type: 'Paid' };
  }
  const theyOwe = round2_(checkedIn * rates.free);
  return { weOwe: 0, theyOwe, rrMakes: round2_(theyOwe - guruFee), type: 'Free' };
}

function makeLedgerRow_(o) {
  return [
    o.dateKey, o.day, o.timeLabel, o.language, o.bookingName, o.phone, o.source,
    o.guests, Number(o.children || 0), o.checkedIn, o.weOwe, o.theyOwe, o.rrMakes, o.type, o.bookingId,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  ];
}


/******************************************************
 * 7. TOKENS  (simple signed session)
 ******************************************************/

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
  const raw = Utilities.computeHmacSha256Signature(s, PORTAL.TOKEN_SECRET);
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

function prop_(obj, key, dflt) {
  return Object.prototype.hasOwnProperty.call(obj, key) ? obj[key] : dflt;
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

/** Normalise "11:00 AM" / "5:00 PM" / "17:00" / "11:00" to 24h "H:MM". */
function normTime24_(v) {
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

function rebuildUnassignedLedger_() {
  const schedule = readSchedule_();      // grid-driven: who is assigned where
  const bookings = readBookingsIndex_(); // dateKey|minutes|language -> [bookings]
  const rates = readRates_();

  // Assignment lookup: shiftKey + '|P' (private) or '|R' (regular) -> assigned[]
  const asg = {};
  schedule.forEach(s => {
    const k = shiftKey_(s.dateKey, s.minutes, s.language) + (s.private ? '|P' : '|R');
    asg[k] = (asg[k] || []).concat(s.assigned || []);
  });

  const today = todayKey_();
  const rows = [];
  Object.keys(bookings).forEach(k => {
    const parts = k.split('|');
    const dateKey = parts[0], minutes = Number(parts[1]);
    if (dateKey < today) return;                            // upcoming only, no expired tours
    bookings[k].forEach(b => {
      const guests = Number(b.guests || 0);
      const source = String(b.source || '');
      if (/guruwalk/i.test(source) && guests <= 0) return;    // skip zero-people Guruwalk
      const isPriv = /privat/i.test(b.note || '');
      const assigned = asg[k + (isPriv ? '|P' : '|R')] || [];
      if (assigned.length) return;                            // assigned -> not unassigned
      const language = tabLanguageForKey_(k);
      const m = computeMoney_(source, guests, isPriv, b.income, rates);
      rows.push([
        dateKey, dayNameFromKey_(dateKey), to12h_(minutesToTime_(minutes)), language,
        source, b.name || '', guests,
        isPriv ? 'Private' : (isPaidSource_(source) ? 'Paid' : 'Free'),
        Number(b.income || 0), m.rrMakes, '', b.bookingId || ''
      ]);
    });
  });

  rows.sort((a, b) => (a[0] + a[2]).localeCompare(b[0] + b[2]));

  const ss = ledgerSS_();
  const sh = ss.getSheetByName('Unassigned') || ss.insertSheet('Unassigned');
  sh.clear();
  const header = ['Date', 'Day', 'Time', 'Language', 'Source', 'Booking', 'Guests',
                  'Type', 'OTA income (€)', 'R&R makes (€)', 'Guide', 'Booking ID'];
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
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
  ensureGuideTabs_(ss);
  ensureQueueTabs_(ss);
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
 *   sendGuruwalkCheckinReminder — daily, 16:00-17:00
 *   archiveLedgerMonthly    — monthly, 1st, 02:00-03:00
 ******************************************************/

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
  'Tour date', 'Time', 'Language', 'Booking ID', 'Guest', 'Adults', 'Children',
  'Guide', 'Checked-in at', '48h deadline', 'Reported in GuruWalk', 'Reported at', 'Notes'
];

function ensureQueueTabs_(ss) {
  ss = ss || ledgerSS_();
  const mk = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      // checkbox column
      const ckCol = headers.indexOf('OTA action done') + 1 || headers.indexOf('Reported in GuruWalk') + 1;
      if (ckCol) sh.getRange(2, ckCol, sh.getMaxRows() - 1, 1).insertCheckboxes();
    }
    return sh;
  };
  mk(QUEUE_TABS.VIATOR_NOSHOW, NOSHOW_HEADERS);
  mk(QUEUE_TABS.GYG_NOSHOW, NOSHOW_HEADERS);
  mk(QUEUE_TABS.GURUWALK, GURUWALK_HEADERS);
}

/** Main entry point — run on a time trigger (hourly). */
function updateManagementQueues() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    ensureQueueTabs_();
    updateNoShowQueues_();
    updateGuruwalkCheckinQueue_();
    rebuildUnassignedLedger_();
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
        updated: String(r[LEDGER_HEADERS.length - 1] || ''),
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
        dateKey, time: String(r[1] || ''), language: String(r[2] || ''),
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
    if (!sh || sh.getLastRow() < 2) return;
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, NOSHOW_HEADERS.length).getValues();
    v.forEach(r => existing[k].add(String(r[4] || '') + '|' + toDateKey_(r[0])));
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
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, NOSHOW_HEADERS.length).setValues(rows);
    sh.getRange(start, 12, rows.length, 1).insertCheckboxes();
  });
}

function updateGuruwalkCheckinQueue_() {
  const ss = ledgerSS_();
  ensureQueueTabs_(ss);
  const sh = ss.getSheetByName(QUEUE_TABS.GURUWALK);
  const checkins = readAllCheckins_();

  const existing = new Set();
  if (sh.getLastRow() >= 2) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, GURUWALK_HEADERS.length).getValues();
    v.forEach(r => existing.add(String(r[3] || '') + '|' + toDateKey_(r[0])));
  }

  const rows = [];
  Object.keys(checkins).forEach(key => {
    const c = checkins[key];
    if (!/guruwalk/i.test(c.source)) return;
    if (existing.has(key)) return;
    existing.add(key);
    const parts = key.split('|');
    const bookingId = parts[0], dateKey = parts[1];
    const deadline = guruwalkDeadline_(dateKey, c.time);
    rows.push([
      dateKey, c.time, c.language, bookingId, c.booking, c.guests, c.children,
      c.guide, c.updated, deadline, false, '', ''
    ]);
  });

  if (rows.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, GURUWALK_HEADERS.length).setValues(rows);
    sh.getRange(start, 11, rows.length, 1).insertCheckboxes();
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

/**
 * Daily afternoon reminder (time trigger 16:00-17:00): tells management which
 * GuruWalk check-ins still need to be reported (48 h window) and how many
 * OTA no-shows are pending.
 */
function sendGuruwalkCheckinReminder() {
  const ss = ledgerSS_();
  ensureQueueTabs_(ss);
  const pendingGuru = [];
  const sh = ss.getSheetByName(QUEUE_TABS.GURUWALK);
  if (sh && sh.getLastRow() >= 2) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, GURUWALK_HEADERS.length).getValues();
    v.forEach(r => { if (r[10] !== true) pendingGuru.push(r); });
  }
  let pendingNoShows = 0;
  [QUEUE_TABS.VIATOR_NOSHOW, QUEUE_TABS.GYG_NOSHOW].forEach(name => {
    const q = ss.getSheetByName(name);
    if (!q || q.getLastRow() < 2) return;
    const v = q.getRange(2, 1, q.getLastRow() - 1, NOSHOW_HEADERS.length).getValues();
    v.forEach(r => { if (r[11] !== true) pendingNoShows++; });
  });

  if (!pendingGuru.length && !pendingNoShows) return;   // nothing to nag about

  let body = 'Daily management queue reminder.\n\n';
  if (pendingGuru.length) {
    body += 'GURUWALK CHECK-INS TO REPORT (48h window from tour start):\n';
    pendingGuru.forEach(r => {
      body += `  - ${toDateKey_(r[0])} ${r[1]} ${r[2]} | ${r[4]} (${r[5]} adults` +
              (Number(r[6]) ? `, ${r[6]} children` : '') + `) | guide ${r[7]} | deadline ${r[9]}\n`;
    });
    body += '\n';
  }
  if (pendingNoShows) {
    body += `OTA NO-SHOWS pending action in Viator/GetYourGuide: ${pendingNoShows} ` +
            `(see "Viator No-shows" and "GYG No-shows" tabs in Guide_Ledger_v1).\n`;
  }
  MailApp.sendEmail({
    to: 'rootsandroadstours@gmail.com',
    subject: 'R&R: GuruWalk check-ins / no-shows to process',
    body
  });
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
    if (sh.getName() === 'Rates') return;
    if (sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
  });

  props.setProperty('LAST_LEDGER_ARCHIVE', stamp);
  MailApp.sendEmail({
    to: 'rootsandroadstours@gmail.com',
    subject: 'R&R: ledger archived — ' + name,
    body: 'The guide ledger for ' + stamp.replace('_', '-') +
          ' was copied to "' + name + '" and the live ledger was cleared for the new month.'
  });
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
    .map(n => { const s = ss.getSheetByName(n); return s ? s.getLastRow() : 0; });

  updateManagementQueues();
  const first = count();
  updateManagementQueues();
  const second = count();

  const stable = JSON.stringify(first) === JSON.stringify(second);
  console.log('Queue rows after run 1: ' + first + ' | after run 2: ' + second +
              ' -> ' + (stable ? 'PASS (idempotent)' : 'FAIL (duplicates added)'));
  return stable;
}
