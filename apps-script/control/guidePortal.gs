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
      case 'assign': out = apiAssign_(p); break;
      case 'move':   out = apiMoveBooking_(p); break;
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
  const schedule = readSchedule_();                 // all upcoming shifts (from the grids)
  const bookingsByKey = readBookingsIndex_();       // "yyyy-mm-dd|minutes|Language" -> [bookings]
  appendOrphanBookingShifts_(schedule, bookingsByKey); // bookings with no grid slot yet -> live extra shifts
  // Order every shift by date then start time (so appended orphans slot into
  // their real time position, not at the end of the list).
  schedule.sort((a, b) =>
    (a.dateKey < b.dateKey ? -1 : a.dateKey > b.dateKey ? 1 : 0) ||
    (a.minutes - b.minutes) ||
    String(a.language).localeCompare(String(b.language)));
  const mine = schedule.filter(s => s.assigned.some(a => sameName_(a, name)));

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
          note: String(b.note || ''),
          checked: isCk,
          checkedIn: isCk ? Number(priorCheckins[kk]) : Number(b.guests || 0) // locked count, or booked default to adjust
        };
      });

    const bookedGuests = bookings.reduce((s, b) => s + Number(b.guests || 0), 0);
    const bookedChildren = bookings.reduce((s, b) => s + Number(b.children || 0), 0);
    const checkedGuests = bookings.reduce((s, b) => s + (b.checked ? Number(b.checkedIn || 0) : 0), 0);

    return {
      id: shift.private ? key + '|P' + (shift.privIndex || 1) : key,
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

  // Who is asking — decides what the shared tour list may show.
  const me = findGuideByName_(name);
  const isManager = !!(me && me.manager);

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
    const raw = readGuidesRaw_();
    const cols = guideColumns_(raw.header);
    raw.rows.forEach(row => {
      const g = parseGuideRow_(row, cols);
      if (!g.name || !g.active) return;
      Object.keys(g.languages).forEach(l => {
        if (g.languages[l] === true) {
          (guidesByLanguage[l] = guidesByLanguage[l] || []).push(g.name);
        }
      });
    });
    busyMap = buildBusyMap_(schedule);
  }
  if (isManager) {
    const ckCache = {};
    const getCk = g => (ckCache[g] || (ckCache[g] = readGuideCheckins_(g)));
    // Managers see EVERY upcoming tour (full UPCOMING_DAYS window), not just
    // this week, so they can plan and assign ahead.
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
            note: String(b.note || ''),
            checked: isCk,
            checkedIn: isCk ? Number(ck[kk]) : Number(b.guests || 0)
          };
        });
      return {
        id: shift.private ? key + '|P' + (shift.privIndex || 1) : key,
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
           schedule: scheduleView, allTours, guidesByLanguage };
}


/**
 * action=assign — MANAGER ONLY. Writes a guide into a Schedule_<Language>
 * grid cell from the portal. The name is written in BOLD, i.e. it becomes a
 * management LOCK that makeSchedule preserves. Empty guide = clear to
 * "Not assigned".
 *   params: token, dateKey (yyyy-MM-dd), time (24h "17:00"), language,
 *           isPrivate ("1"/""), privIndex, guide
 */
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
    return writeAssignmentToGrid_(language, dateKey, time, isPriv, privIndex, guide);
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
    return moveBookingRowBetweenTabs_(bookingId, fromLanguage, toLanguage);
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

function writeAssignmentToGrid_(language, dateKey, time, isPriv, privIndex, guide) {
  const sh = control_().getSheetByName('Schedule_' + language);
  if (!sh || sh.getLastRow() < 3) {
    return { ok: false, error: 'Schedule_' + language + ' tab not found or empty' };
  }
  const dv = sh.getDataRange().getDisplayValues();
  const anchor = gridAnchor_(String(dv[0][0] || ''));
  const timeRow = dv[1] || [];

  let col = -1;
  for (let c = 1; c < timeRow.length; c++) {
    const h = parseGridTimeHeader_(timeRow[c]);
    if (h && h.time === time && h.isPrivate === isPriv && (!isPriv || h.index === privIndex)) {
      col = c; break;
    }
  }
  if (col === -1) {
    return { ok: false, error: 'No ' + (isPriv ? 'private ' : '') + time + ' column in Schedule_' + language };
  }

  let row = -1;
  for (let r = 2; r < dv.length; r++) {
    if (gridLabelToKey_(String(dv[r][0] || '').trim(), anchor) === dateKey) { row = r; break; }
  }
  if (row === -1) return { ok: false, error: dateKey + ' is not in Schedule_' + language };

  const cell = sh.getRange(row + 1, col + 1);
  if (!guide) {
    cell.setValue('Not assigned');
    cell.setFontWeight('normal').setFontStyle('italic').setFontColor('#94a3b8');
    return { ok: true, assigned: '' };
  }

  // Manager assignment = LOCK -> bold, so makeSchedule never moves it.
  cell.setFontStyle('normal').setFontColor('#1a2b49');
  cell.setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText(guide)
      .setTextStyle(0, guide.length, SpreadsheetApp.newTextStyle().setBold(true).build())
      .build());
  return { ok: true, assigned: guide };
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
    // Only the ACTIVE language tabs. "Done Tours" also ends in " Tours" but is
    // an aggregate (no booking ids) and must never be parsed as bookings.
    if (tab.indexOf(PORTAL.BOOKING_TAB_SUFFIX) === -1) return;
    if (/^done\b/i.test(tab)) return;
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
      const updated = String(r[LEDGER_HEADERS.length - 1] || '');
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
  const start = d.getTime() + Math.max(0, Number(minutes) || 0) * 60000;
  return Date.now() > start + 2 * 3600000;
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

    if (dataWidth === LEDGER_HEADERS.length - 1) {   // 15 -> needs Children inserted
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
  if (PORTAL.TOKEN_SECRET.indexOf('CHANGE_ME') === 0) {
    warn('TOKEN_SECRET is still the placeholder', 'edit guidePortal.gs section 1');
  } else { ok('TOKEN_SECRET customised'); }

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
