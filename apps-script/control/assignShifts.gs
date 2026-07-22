/******************************************************
 * ROOTS & ROADS — assignShifts.gs  (v2, complete, ready to paste)
 * Bind to: Roots_Roads_Control_v1
 *
 * WHAT CHANGED IN v2
 *  • MANAGER LOCKS: a guide name written in BOLD inside a Schedule_<Language>
 *    grid cell is a management lock. makeSchedule preserves it exactly (still
 *    bold), assigns everyone else around it, and never overwrites it. If a
 *    lock conflicts (wrong language, unavailable, overlapping tour), the lock
 *    is KEPT, the cell is tinted red, and the conflict is written to the
 *    Control sheet's "Errors" tab — never silently resolved.
 *  • OVERLAP RULE: two tours by the same guide must start at least
 *    ASSIGN_CFG.MIN_SEPARATION_HOURS (5) apart. This kills the
 *    "one guide on 10:00 + 11:00" and "two simultaneous 17:00s" bugs the
 *    portal was showing.
 *  • SCORING is centralised in ASSIGN_CFG and documented there.
 *  • Availability advantage: guides who tick more availability get a measured
 *    edge (tiebreaker), to reward giving us more hours.
 *  • Dead code removed (old assignGuides_).
 ******************************************************/

const GUIDE_FILE_ID = "1ZkF0yDVE5Q7V2XUO5051ojJdTXtsvb8FA2sh15rqf-0";
const BOOKING_SHEET_ID = "1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw";

const LANGUAGES = ["English", "German", "Spanish", "French"];
const AVAILABILITY_FIRST_ROW = 5;
const AVAILABILITY_MIN_GUIDE_ROWS = 20;

/* ============================================================
 * ASSIGNMENT CONFIGURATION — every scheduling rule in one place.
 *
 * Assignment is deterministic. Shifts are processed most-valuable-first;
 * for each shift the eligible pool is sorted by, in order:
 *   1. seniority        ascending  (1 beats 2 beats 3 — seniors get first pick)
 *   2. carried value    ascending  (same seniority: whoever has the least
 *                                   accumulated tour value gets the next good
 *                                   tour — balances strong and weak tours)
 *   3. tour count       ascending  (same value: fewest tours first)
 *   4. availability     descending (same count: whoever offered MORE
 *                                   availability wins the tie — the incentive)
 *   5. sheet order      ascending  (final deterministic tiebreak)
 * ============================================================ */
const ASSIGN_CFG = {
  // A guide's two tours must start at least this many hours apart.
  MIN_SEPARATION_HOURS: 5,

  // A private booking's time may not exactly match an availability column
  // (e.g. booked 10:00, availability ticked at 10:30). Accept a guide's
  // availability up to this many minutes away from the booked time. If no
  // one is available within the window the shift stays "Not assigned" and
  // management staffs it manually (by talking to guides).
  PRIVATE_AVAIL_TOLERANCE_MIN: 60,

  // Value of a shift to the guide who runs it (drives balancing).
  VALUE_PAID_PER_GUEST: 10,     // paid tour: guide earns ~10 €/guest
  VALUE_FREE_NET_PER_GUEST: 14, // free tour: ~20 € tip − 6 € commission per guest
  VALUE_PRIVATE_FLAT: 75,       // private tour: flat 75 € to the guide

  PAID_SOURCES: ["Viator", "GetYourGuide", "Airbnb"],

  // Where schedule problems (lock conflicts, missing eligibility) are logged.
  ERRORS_TAB: "Errors",

  // Availability slots for PRIVATE tours, per weekday. These only add tick
  // columns to the availability sheet — they never create tours ("Private"
  // is NOT a language and never appears in Weekly_Schedule). Actual private
  // shifts are booking-driven at the booking's real time, in the booking's
  // language.
  PRIVATE_AVAILABILITY: {
    Monday:    ['10:30', '17:00'],
    Tuesday:   ['10:30', '17:00'],
    Wednesday: ['10:30', '17:00'],
    Thursday:  ['10:30', '17:00'],
    Friday:    ['10:30', '17:00'],
    Saturday:  ['17:00'],
    Sunday:    []
  }
};

/** Pseudo-rules that expose the private availability slots as columns in the
 *  Week tabs. Shaped like Weekly_Schedule rules so the sync code can just
 *  concat them. */
function privateAvailabilityRules_() {
  const out = [];
  Object.keys(ASSIGN_CFG.PRIVATE_AVAILABILITY).forEach(day => {
    ASSIGN_CFG.PRIVATE_AVAILABILITY[day].forEach(time => {
      out.push({ day, time, language: 'Private', guidesNeeded: 0, activeFrom: null, activeUntil: null });
    });
  });
  return out;
}

// Back-compat aliases (used across this file).
const TOUR_DURATION_HOURS = ASSIGN_CFG.MIN_SEPARATION_HOURS;
const VALUE_PAID_PER_GUEST = ASSIGN_CFG.VALUE_PAID_PER_GUEST;
const VALUE_FREE_NET_PER_GUEST = ASSIGN_CFG.VALUE_FREE_NET_PER_GUEST;
const VALUE_PRIVATE_FLAT = ASSIGN_CFG.VALUE_PRIVATE_FLAT;
const PAID_SOURCES_ASSIGN = ASSIGN_CFG.PAID_SOURCES;


function makeSchedule() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  const guideSS = SpreadsheetApp.openById(GUIDE_FILE_ID);

  const guides = readGuides_(controlSS);
  const weeklySchedule = readWeeklySchedule_(controlSS);
  const languageSpeakerCounts = countLanguageSpeakers_(guides);

  const startDate = dateOnly_(new Date());        // today
  const endDate = endOfScheduleRange_();          // Sunday of next full week
  const weekNames = weekTabsToSchedule_(guideSS); // this week's tab + next week's tab

  // Private slots live in Weekly_Schedule as language "Private" so the
  // availability tab shows their times (e.g. 10:00) and guides can tick them.
  // They must NOT generate empty regular tours, so regular shift-building skips
  // them; private shifts are built from actual private bookings below.
  const regularRules = weeklySchedule.filter(r => String(r.language).toLowerCase() !== 'private');

  // Availability index across the scheduled weeks, keyed "dateText|time", so a
  // private booking can be staffed at ITS OWN time even with no regular slot.
  const availIndex = {};
  let shifts = [];
  weekNames.forEach(weekName => {
    const guideCalendar = readGuideCalendar_(guideSS, weekName);
    const availability = readAvailability_(guideSS, weekName);
    availability.forEach(a => {
      const ak = a.dateText + '|' + a.time;
      if (!availIndex[ak]) availIndex[ak] = [];
      if (availIndex[ak].indexOf(a.guideName) === -1) availIndex[ak].push(a.guideName);
    });
    buildShifts_(availability, regularRules, guideCalendar).forEach(s => {
      const d = dateOnly_(s.dateObj);
      if (d >= startDate && d <= endDate) shifts.push(s);
    });
  });

  shifts = expandPrivateShifts_(shifts, availIndex, startDate, endDate);   // private = own shift at its own time
  shifts = expandOrphanShifts_(shifts, availIndex, startDate, endDate);    // bookings outside Weekly_Schedule still get a shift

  // Availability advantage: total slots each guide ticked across the window.
  const availCount = {};
  Object.keys(availIndex).forEach(k => availIndex[k].forEach(n => {
    availCount[n] = (availCount[n] || 0) + 1;
  }));

  // MANAGER LOCKS: bold names in the Schedule_<Language> grids are fixed.
  const locks = readLockedAssignments_(controlSS);

  // Value of each shift to the guide who runs it (private = flat; paid = /guest;
  // free = ~net tip/guest). This drives the balancing.
  const shiftValues = readShiftValues_();
  shifts.forEach(s => {
    const key = s.dateText + "|" + normalizeTime_(s.time) + "|" + s.language;
    s.value = s.isPrivate ? ASSIGN_CFG.VALUE_PRIVATE_FLAT : (shiftValues[key] || 0);
  });

  // Deterministic order: strongest (most valuable) tours first.
  const order = [...shifts].sort((a, b) =>
    b.value - a.value || a.dateTimeObj - b.dateTimeObj || a.language.localeCompare(b.language));

  const accVal = {}, tourCount = {}, assignedByGuide = {};
  guides.forEach(g => { accVal[g.name] = 0; tourCount[g.name] = 0; assignedByGuide[g.name] = []; });
  const byName = {};
  guides.forEach(g => { byName[g.name.toLowerCase()] = g; });
  const conflicts = [];

  // PASS 1 — seat every manager lock first, so auto-assignment works around
  // them. A lock is preserved even when it conflicts; conflicts are flagged.
  order.forEach(shift => {
    const lk = lockKey_(shift.dateText, normalizeTime_(shift.time), shift.language, shift.isPrivate, shift.privIndex);
    const lockedNames = locks[lk] || [];
    shift.lockedGuides = [];
    lockedNames.forEach(nm => {
      const g = byName[nm.toLowerCase()];
      const problems = [];
      if (!g) problems.push('not in Guides tab');
      else {
        if (!g.active) problems.push('inactive');
        if (g.languages[shift.language] !== true) problems.push('does not speak ' + shift.language);
        if (!shift.availableGuides.includes(g.name)) problems.push('not marked available');
        if (hasConflict_(assignedByGuide[g.name] || [], shift.dateTimeObj)) {
          problems.push('overlaps another assigned tour (<' + ASSIGN_CFG.MIN_SEPARATION_HOURS + 'h apart)');
        }
      }
      // Keep the lock regardless; flag any problem. "Not marked available"
      // alone is SOFT (managers assign people all the time after talking to
      // them — that must not paint the grid red); wrong language / overlap /
      // unknown guide are HARD conflicts.
      shift.lockedGuides.push(nm);
      if (g) {
        accVal[g.name] += shift.value;
        tourCount[g.name]++;
        assignedByGuide[g.name].push(shift.dateTimeObj);
      }
      if (problems.length) {
        const hard = problems.some(p => p.indexOf('not marked available') === -1);
        conflicts.push({
          dateText: shift.dateText, time: shift.time, language: shift.language,
          isPrivate: !!shift.isPrivate, privIndex: shift.privIndex, guide: nm,
          problems, hard
        });
      }
    });
  });

  // PASS 2 — auto-assign the remaining seats.
  const assignedShifts = [];
  order.forEach(shift => {
    const eligible = guides.filter(g =>
      g.active &&
      g.languages[shift.language] === true &&
      shift.availableGuides.includes(g.name) &&
      !shift.lockedGuides.some(n => n.toLowerCase() === g.name.toLowerCase()) &&
      !hasConflict_(assignedByGuide[g.name], shift.dateTimeObj)
    );

    const need = Math.max(0, (shift.guidesNeeded || 1) - shift.lockedGuides.length);
    const assigned = [];
    const pool = [...eligible];
    while (assigned.length < need && pool.length) {
      pool.sort((a, b) =>
        a.seniority - b.seniority ||                              // 1. seniors first
        accVal[a.name] - accVal[b.name] ||                        // 2. least value carried
        tourCount[a.name] - tourCount[b.name] ||                  // 3. fewest tours
        (availCount[b.name] || 0) - (availCount[a.name] || 0) ||  // 4. more availability wins
        a.order - b.order                                         // 5. deterministic
      );
      // Re-check overlap: an earlier pick this shift can't collide (same
      // start), but a guide may have gained a lock meanwhile.
      const pick = pool.shift();
      if (hasConflict_(assignedByGuide[pick.name], shift.dateTimeObj)) continue;
      assigned.push(pick);
      accVal[pick.name] += shift.value;
      tourCount[pick.name]++;
      assignedByGuide[pick.name].push(shift.dateTimeObj);
    }

    const totalAssigned = shift.lockedGuides.length + assigned.length;
    const needTotal = shift.guidesNeeded || 1;
    const ok = totalAssigned >= needTotal;
    const hasConflictFlag = conflicts.some(c =>
      c.hard &&
      c.dateText === shift.dateText && normalizeTime_(c.time) === normalizeTime_(shift.time) &&
      c.language === shift.language && !!c.isPrivate === !!shift.isPrivate &&
      (Number(c.privIndex) || 1) === (Number(shift.privIndex) || 1));

    assignedShifts.push({
      ...shift,
      eligibleGuides: eligible.map(g => g.name),
      assignedGuides: shift.lockedGuides.concat(assigned.map(g => g.name)),
      lockedGuides: shift.lockedGuides,
      hasLockConflict: hasConflictFlag,
      status: ok ? "OK" : "Not assigned",
      notes: [
        shift.isPrivate ? "Private" : "",
        shift.extra ? "Extra tour (not in Weekly_Schedule)" : "",
        shift.lockedGuides.length ? "Locked: " + shift.lockedGuides.join(", ") : "",
        hasConflictFlag ? "LOCK CONFLICT (see Errors tab)" : "",
        ok ? "" : `Need ${needTotal}, assigned ${totalAssigned}`
      ].filter(Boolean).join(" · ")
    });
  });

  assignedShifts.sort((a, b) => a.dateTimeObj - b.dateTimeObj || a.language.localeCompare(b.language));

  logScheduleConflicts_(controlSS, conflicts);
  writeDetailedSchedule_(controlSS, assignedShifts);
  makeLanguageScheduleTabs_(controlSS, assignedShifts);

  // Self-check every run: any eligibility/overlap violation in the freshly
  // written grids lands in the Errors tab instead of in a guide's phone.
  let violations = 0;
  try { violations = validateScheduleGrids(); } catch (e) { /* validator must never kill the run */ }

  try { markHealthEvent_('HB_SCHEDULE'); updateControlHealth_(); } catch (e) { /* dashboard is best-effort */ }

  return { shifts: assignedShifts.length, conflicts: conflicts.length, violations };
}


/**
 * RUN ONCE (or whenever the public offer changes): rewrites the
 * English / Spanish / Private rows of Weekly_Schedule to the current offer
 * and PRESERVES every other language's rows (German stays exactly as you
 * maintain it by hand).
 *
 * Current offer (decided 2026-07, GYG + Viator):
 *   Mon/Tue/Thu/Fri: English 11:00 + 17:00 · Spanish 10:30
 *   Wednesday:       English 17:00
 *   Saturday:        English 17:00
 * Weekly_Schedule holds ONLY real tour languages. Private availability slots
 * live in ASSIGN_CFG.PRIVATE_AVAILABILITY (they are not tours and private is
 * not a language). German rows are preserved exactly as you maintain them —
 * including repairing any time cell Sheets had corrupted into 12/30/1899.
 * The Time column is text-formatted so the corruption cannot come back.
 */
function updateWeeklyScheduleToCurrentOffer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Weekly_Schedule');
  if (!sh) throw new Error('Weekly_Schedule tab not found');

  // Preserve rows for languages this function does not manage (e.g. German),
  // normalising the Time cell whether it is a string or a coerced Date.
  const preserved = [];
  if (sh.getLastRow() > 1) {
    const raw = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    const dv = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues();
    raw.forEach((r, i) => {
      const lang = String(dv[i][2] || r[2] || '').trim();
      if (!lang || ['english', 'spanish', 'private'].indexOf(lang.toLowerCase()) !== -1) return;
      const time = normalizeTime_(dv[i][1]) || timeFromCellValue_(r[1]);
      preserved.push([
        String(dv[i][0] || r[0] || '').trim(),
        time || String(dv[i][1] || ''),
        lang,
        Number(r[3]) || 1,
        r[4] || '',
        r[5] || ''
      ]);
    });
  }

  const rows = [];
  // Blank "Active from" = always active. The standard weekly offer must show in
  // EVERY availability week. (Stamping today's date here made the offer look
  // like it started on the run date, hiding tours such as 11:00 from earlier
  // weeks and — on any mid-week re-run — from the current week's early days.)
  const add = (days, time, lang) => days.forEach(d => rows.push([d, time, lang, 1, '', '']));
  const MTThF = ['Monday', 'Tuesday', 'Thursday', 'Friday'];

  add(MTThF, '11:00', 'English');
  add(MTThF, '17:00', 'English');
  add(MTThF, '10:30', 'Spanish');
  add(['Wednesday'], '17:00', 'English');
  add(['Saturday'], '17:00', 'English');

  const all = [['Day', 'Time', 'Language', 'Guides needed', 'Active from', 'Active until']]
    .concat(rows)
    .concat(preserved);

  sh.clear();
  // Time column as TEXT before writing: "11:00" can never again become a Date.
  sh.getRange(1, 2, all.length, 1).setNumberFormat('@');
  sh.getRange(1, 1, all.length, 6).setValues(all);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  Logger.log('Weekly_Schedule updated: ' + rows.length + ' offer rows + ' +
             preserved.length + ' preserved rows (German etc., times repaired).');
}


/* ---------- manager locks (bold names in the grids) ---------- */

function lockKey_(dateText, time, language, isPrivate, privIndex) {
  return dateText + "|" + time + "|" + language + "|" +
         (isPrivate ? "P" + (Number(privIndex) || 1) : "R");
}

/**
 * Parse a Schedule_<Language> column header.
 *   "11:00"           -> { time:"11:00", isPrivate:false, index:1 }
 *   "10:00 · Private" -> { time:"10:00", isPrivate:true,  index:1 }
 *   "17:00 · Private 2" -> { time:"17:00", isPrivate:true, index:2 }
 * Also tolerates display forms like "11:00 AM". Returns null for non-time
 * headers. Shared by the grid writer, the lock reader, and the guide portal
 * (same Apps Script project).
 */
function parseGridTimeHeader_(text) {
  const s = String(text || '').trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2}:\d{2}(?:\s*[AP]M)?)(?:\s*[·\-]\s*Private(?:\s*(\d+))?)?$/i);
  if (!m) return null;
  const t = normalizeTime_(m[1]);
  if (!/^\d{1,2}:\d{2}$/.test(t)) return null;
  return {
    time: t,
    isPrivate: /private/i.test(s),
    index: m[2] ? Number(m[2]) : 1
  };
}

/** Column header text for a shift column. */
function gridHeaderForColumn_(time, isPrivate, index) {
  if (!isPrivate) return time;
  return time + ' · Private' + (Number(index) > 1 ? ' ' + index : '');
}

/**
 * Read every BOLD guide name from every Schedule_<Language> grid.
 * Bold = management lock. Returns { lockKey -> [names] }.
 * A cell can hold a regular line and 🔒-prefixed private lines; a bold run
 * locks the names inside that run only, so a manager can bold one name of two.
 */
function readLockedAssignments_(controlSS) {
  const locks = {};
  controlSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.indexOf("Schedule_") !== 0) return;
    const language = name.substring("Schedule_".length).trim();
    if (!language || sheet.getLastRow() < 3) return;

    const dv = sheet.getDataRange().getDisplayValues();
    const anchorYear = extractYear_(String(dv[0][0] || ""));
    const timeRow = dv[1] || [];
    const rich = sheet.getDataRange().getRichTextValues();

    for (let r = 2; r < dv.length; r++) {
      const label = String(dv[r][0] || "").trim();
      if (!label) continue;
      // label is like "Thu Jul 16": parse month/day directly.
      const m = label.match(/([A-Za-z]{3,})\s+(\d{1,2})$/);
      if (!m) continue;
      const dObj = new Date(`${m[1]} ${m[2]}, ${anchorYear} 12:00:00`);
      if (isNaN(dObj)) continue;
      const dateText = formatDate_(dObj);

      for (let c = 1; c < timeRow.length; c++) {
        const h = parseGridTimeHeader_(timeRow[c]);
        if (!h) continue;
        const rt = rich[r] && rich[r][c];
        if (!rt) continue;
        const runs = rt.getRuns ? rt.getRuns() : [];
        runs.forEach(run => {
          const style = run.getTextStyle();
          if (!style || !style.isBold()) return;
          const txt = String(run.getText() || "");
          // Column header decides private/regular; the legacy 🔒 in-cell
          // marker still wins for grids written before the column redesign.
          const isPriv = h.isPrivate || /🔒|\(private\)/i.test(txt);
          const idx = h.isPrivate ? h.index : 1;
          txt.split(/[\n,]/).forEach(piece => {
            let nm = piece.replace(/🔒/g, "").replace(/\(private\)/ig, "").trim();
            if (!nm || /not assigned|lock conflict|need \d/i.test(nm)) return;
            const k = lockKey_(dateText, h.time, language, isPriv, idx);
            if (!locks[k]) locks[k] = [];
            if (locks[k].indexOf(nm) === -1) locks[k].push(nm);
          });
        });
      }
    }
  });
  return locks;
}

/** Write lock conflicts to the Control sheet's Errors tab (visible flag). */
function logScheduleConflicts_(controlSS, conflicts) {
  if (!conflicts.length) return;
  let sh = controlSS.getSheetByName(ASSIGN_CFG.ERRORS_TAB);
  if (!sh) {
    sh = controlSS.insertSheet(ASSIGN_CFG.ERRORS_TAB);
    sh.getRange(1, 1, 1, 4).setValues([["Timestamp", "Type", "Details", "Raw data"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  const ts = Utilities.formatDate(new Date(), "Europe/Madrid", "yyyy-MM-dd HH:mm");
  const rows = conflicts.map(c => [
    ts, c.hard === false ? "Schedule note" : "Schedule lock conflict",
    `${c.guide} locked on ${c.dateText} ${c.time} ${c.language}${c.isPrivate ? " (private)" : ""}: ${c.problems.join("; ")}`,
    ""
  ]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}


/* ---------- NEW: scheduling window + private tours ---------- */

function endOfScheduleRange_() {
  const today = dateOnly_(new Date());
  const dow = (today.getDay() + 6) % 7;                 // 0=Mon..6=Sun
  const mondayThis = new Date(today); mondayThis.setDate(today.getDate() - dow);
  const sundayNext = new Date(mondayThis); sundayNext.setDate(mondayThis.getDate() + 13);
  return dateOnly_(sundayNext);
}

function weekTabsToSchedule_(guideSS) {
  const today = dateOnly_(new Date());
  const dow = (today.getDay() + 6) % 7;
  const mondayThis = new Date(today); mondayThis.setDate(today.getDate() - dow);
  const mondayNext = new Date(mondayThis); mondayNext.setDate(mondayThis.getDate() + 7);

  const wanted = [formatDate_(mondayThis), formatDate_(mondayNext)];
  const names = [];

  guideSS.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith("Week ")) return;
    const dv = sheet.getDataRange().getDisplayValues();
    const year = extractYear_(String(dv[0][0]));
    const dayRow = dv[2] || [];
    for (let c = 1; c < dayRow.length; c++) {
      if (!dayRow[c]) continue;
      const dk = formatDate_(parseDateFromHeader_(String(dayRow[c]), year));
      if (wanted.indexOf(dk) !== -1 && names.indexOf(sheet.getName()) === -1) names.push(sheet.getName());
    }
  });

  if (!names.length) throw new Error("No availability Week tab found for this week or next week.");
  return names;
}

/**
 * Keep the regular shifts, then add ONE private shift PER PRIVATE BOOKING at
 * its real booked time. Each group is its own shift (own grid column): a
 * private tour is effectively a new shift, never merged into a regular slot.
 * availableGuides come from whoever ticked availability at that date+time —
 * or, if the booked time sits between availability columns (10:00 booking vs
 * 10:30 tick), from the nearest slot within PRIVATE_AVAIL_TOLERANCE_MIN.
 * If nobody matches, the shift stays "Not assigned" and management staffs it
 * manually outside the system.
 */
function expandPrivateShifts_(shifts, availIndex, startDate, endDate) {
  const groups = readPrivateGroups_();   // [{dateText,time,language,bookingId}]
  const out = shifts.map(s => Object.assign({}, s, { isPrivate: false, privIndex: 0 }));

  // Deterministic index per (date,time,language): sort groups by booking id.
  const counters = {};
  groups.sort((a, b) =>
    (a.dateText + a.time + a.language + a.bookingId)
      .localeCompare(b.dateText + b.time + b.language + b.bookingId));

  groups.forEach(g => {
    const dateObj = new Date(g.dateText + 'T12:00:00');
    const d = dateOnly_(dateObj);
    if (d < startDate || d > endDate) return;                 // outside the window
    const slotKey = g.dateText + '|' + g.time + '|' + g.language;
    counters[slotKey] = (counters[slotKey] || 0) + 1;
    out.push({
      week: '', dateObj,
      dateTimeObj: combineDateAndTime_(dateObj, g.time),
      dateText: g.dateText, day: fullDayName_(dateObj),
      time: g.time, language: g.language, guidesNeeded: 1,
      availableGuides: privateAvailability_(availIndex, g.dateText, g.time),
      isPrivate: true,
      privIndex: counters[slotKey],
      privBookingId: g.bookingId
    });
  });

  return out;
}

/**
 * SAFETY NET: any non-private booking sitting at a date/time/language with NO
 * matching shift (a slot missing from Weekly_Schedule, a legacy time, a
 * manually added row...) still becomes a tour: it shows in the grids and the
 * portal, gets auto-assigned from availability like everything else, and
 * comes out "Not assigned" when nobody fits — so management KNOWS to find a
 * guide instead of the tour silently not existing.
 */
function expandOrphanShifts_(shifts, availIndex, startDate, endDate) {
  const existing = new Set(shifts.filter(s => !s.isPrivate)
    .map(s => s.dateText + '|' + normalizeTime_(s.time) + '|' + s.language));
  const out = shifts.slice();

  readActiveBookingSlots_().forEach(slot => {
    const key = slot.dateText + '|' + slot.time + '|' + slot.language;
    if (existing.has(key)) return;
    existing.add(key);
    const dateObj = new Date(slot.dateText + 'T12:00:00');
    const d = dateOnly_(dateObj);
    if (d < startDate || d > endDate) return;
    out.push({
      week: '', dateObj,
      dateTimeObj: combineDateAndTime_(dateObj, slot.time),
      dateText: slot.dateText, day: fullDayName_(dateObj),
      time: slot.time, language: slot.language, guidesNeeded: 1,
      availableGuides: privateAvailability_(availIndex, slot.dateText, slot.time),
      isPrivate: false, privIndex: 0,
      extra: true                       // flagged in the Schedule tab notes
    });
  });
  return out;
}

/** Unique (date, time, language) slots of all NON-private active bookings. */
function readActiveBookingSlots_() {
  const seen = new Set();
  const out = [];
  const ss = SpreadsheetApp.openById(BOOKING_SHEET_ID);
  ss.getSheets().forEach(sh => {
    const tab = sh.getName();
    if (tab.indexOf(" Tours") === -1) return;
    if (/^done\b/i.test(tab)) return;
    const language = tab.replace(" Tours", "").trim();
    if (sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    rows.forEach(row => {
      if (/privat/i.test(String(row[8] || ""))) return;
      const time = timeFromCellValue_(row[4]) || normalizeTime_(row[4]);
      if (!/^\d{1,2}:\d{2}$/.test(time)) return;
      const dateText = row[3] instanceof Date ? formatDate_(row[3]) : formatDate_(new Date(row[3]));
      const key = dateText + '|' + time + '|' + language;
      if (seen.has(key)) return;
      seen.add(key);
      out.push({ dateText, time, language });
    });
  });
  return out;
}

/** Availability for a private booking: exact time, else nearest slot within
 *  the tolerance window (same day). */
function privateAvailability_(availIndex, dateText, time) {
  const exact = availIndex[dateText + '|' + time];
  if (exact && exact.length) return exact;

  const target = timeToMinutes_(time);
  let best = null;
  let bestDiff = ASSIGN_CFG.PRIVATE_AVAIL_TOLERANCE_MIN + 1;
  Object.keys(availIndex).forEach(k => {
    const parts = k.split('|');
    if (parts[0] !== dateText) return;
    const diff = Math.abs(timeToMinutes_(parts[1]) - target);
    if (diff < bestDiff) { bestDiff = diff; best = availIndex[k]; }
  });
  return (best && bestDiff <= ASSIGN_CFG.PRIVATE_AVAIL_TOLERANCE_MIN) ? best : [];
}

/**
 * Value per regular shift = sum over its non-private bookings of the guide's
 * expected earnings (paid: 10/guest, free: ~14/guest). Private groups are valued
 * separately (flat 75) because they're their own shift.
 */
function readShiftValues_() {
  const out = {};
  const ss = SpreadsheetApp.openById(BOOKING_SHEET_ID);
  ss.getSheets().forEach(sh => {
    const tab = sh.getName();
    if (tab.indexOf(" Tours") === -1) return;
    if (/^done\b/i.test(tab)) return;   // "Done Tours" is an aggregate, not bookings
    const language = tab.replace(" Tours", "").trim();
    if (sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    rows.forEach(row => {
      if (/privat/i.test(String(row[8] || ""))) return;   // private valued as its own shift
      const dateKey = row[3] instanceof Date ? formatDate_(row[3]) : formatDate_(new Date(row[3]));
      const time = timeFromCellValue_(row[4]) || normalizeTime_(row[4]);
      const guests = Number(row[2] || 0);
      const source = String(row[5] || "");
      const paid = PAID_SOURCES_ASSIGN.some(s => s.toLowerCase() === source.toLowerCase());
      const val = (paid ? VALUE_PAID_PER_GUEST : VALUE_FREE_NET_PER_GUEST) * guests;
      const key = dateKey + "|" + time + "|" + language;
      out[key] = (out[key] || 0) + val;
    });
  });
  return out;
}

/** One entry per private BOOKING (not per slot), with its real time. */
function readPrivateGroups_() {
  const out = [];
  const ss = SpreadsheetApp.openById(BOOKING_SHEET_ID);
  ss.getSheets().forEach(sh => {
    const tab = sh.getName();
    if (tab.indexOf(" Tours") === -1) return;
    if (/^done\b/i.test(tab)) return;   // "Done Tours" is an aggregate, not bookings
    const language = tab.replace(" Tours", "").trim();
    if (sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    rows.forEach(row => {
      if (!/privat/i.test(String(row[8] || ""))) return;
      const dateText = row[3] instanceof Date ? formatDate_(row[3]) : formatDate_(new Date(row[3]));
      out.push({
        dateText,
        time: timeFromCellValue_(row[4]) || normalizeTime_(row[4]),
        language,
        bookingId: String(row[7] || '').trim() || ('ROW' + out.length)
      });
    });
  });
  return out;
}


/* ---------- unchanged workflow entry points ---------- */

function syncAvailabilityFile() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  const guideSS = SpreadsheetApp.openById(GUIDE_FILE_ID);

  const guideNames = readActiveGuideNames_(controlSS);
  const scheduleRules = readWeeklySchedule_(controlSS).concat(privateAvailabilityRules_());

  guideSS.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith("Week ")) return;
    const weekDates = readWeekDates_(sheet);
    if (weekDates.length === 0) return;
    const existingAvailability = readExistingAvailability_(sheet);
    rebuildAvailabilityWeekSheet_(sheet, weekDates, scheduleRules, guideNames, existingAvailability);
  });
}

function runWeeklyScheduling() {
  ensureWeekTabs_();      // delete past weeks, create the upcoming ones
  syncAvailabilityFile();
  makeSchedule();
  emailWeeklySchedule();  // send the language tables to management
}

/**
 * RUN ONCE (from the editor) to fill Weekly_Schedule with the current offer.
 * Overwrites the tab. Private tours are booking-driven, so they are NOT rows here.
 */
function setupWeeklySchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Weekly_Schedule') || ss.insertSheet('Weekly_Schedule');

  const rows = [['Day', 'Time', 'Language', 'Guides needed', 'Active from', 'Active until']];
  // Blank "Active from" = always active (see updateWeeklyScheduleToCurrentOffer).
  const add = (days, time, lang) => days.forEach(d => rows.push([d, time, lang, 1, '', '']));

  const MTThF = ['Monday', 'Tuesday', 'Thursday', 'Friday'];
  add(MTThF, '11:00', 'English');
  add(MTThF, '17:00', 'English');
  add(MTThF, '10:30', 'Spanish');
  add(MTThF, '17:00', 'Spanish');
  add(['Wednesday'], '17:00', 'English');
  add(['Wednesday'], '17:00', 'Spanish');
  add(['Saturday'], '17:00', 'English');
  add(['Saturday'], '17:00', 'Spanish');

  // Private slots: language "Private" so the availability tab exposes their times
  // (notably 10:00, which has no regular tour). These rows create availability
  // columns only; actual private tours are built from bookings in makeSchedule.
  add(MTThF, '10:00', 'Private');
  add(MTThF, '17:00', 'Private');
  add(['Wednesday'], '10:00', 'Private');
  add(['Wednesday'], '17:00', 'Private');
  add(['Saturday'], '17:00', 'Private');

  sh.clear();
  sh.getRange(1, 1, rows.length, 6).setValues(rows);
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

/**
 * Emails the generated Schedule_<Language> tables to management as HTML, ready
 * to forward to guides. Called at the end of runWeeklyScheduling (Friday run).
 */
function emailWeeklySchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = ss.getSheets().filter(s => s.getName().indexOf('Schedule_') === 0);
  if (!tabs.length) return;

  let html = '<div style="font-family:Arial,sans-serif">';
  tabs.forEach(sh => {
    if (sh.getLastRow() < 2) return;
    const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    html += '<h3 style="color:#1a2b49">' + v[0][0] + '</h3><table cellspacing="0" cellpadding="6" ' +
            'style="border-collapse:collapse;margin-bottom:24px">';
    for (let r = 1; r < v.length; r++) {
      html += '<tr>';
      v[r].forEach((cell, c) => {
        const miss = /not assigned/i.test(cell);
        html += '<td style="border:1px solid #cbd5e1;' + (r === 1 ? 'background:#dbeafe;font-weight:bold;' : '') +
                (miss ? 'color:#94a3b8;' : '') + '">' + String(cell).replace(/\n/g, '<br>') + '</td>';
      });
      html += '</tr>';
    }
    html += '</table>';
  });
  html += '</div>';

  MailApp.sendEmail({
    to: 'rootsandroadstours@gmail.com',
    subject: 'Weekly guide schedule',
    htmlBody: html
  });
}

/** Weeks to keep visible ahead of the current one. */
const WEEKS_AHEAD = 4;

/**
 * Auto-manage the availability Week tabs: delete any week fully in the past,
 * and create the current week + WEEKS_AHEAD upcoming weeks if missing.
 * Safe to run repeatedly.
 */
function ensureWeekTabs_() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  const guideSS = SpreadsheetApp.openById(GUIDE_FILE_ID);
  const guideNames = readActiveGuideNames_(controlSS);
  const scheduleRules = readWeeklySchedule_(controlSS).concat(privateAvailabilityRules_());

  const today = dateOnly_(new Date());
  const dow = (today.getDay() + 6) % 7;
  const mondayThis = new Date(today); mondayThis.setDate(today.getDate() - dow);

  // Names we want to exist (current + next weeks).
  const wanted = {};
  for (let i = 0; i <= WEEKS_AHEAD; i++) {
    const m = new Date(mondayThis); m.setDate(mondayThis.getDate() + 7 * i);
    m.setHours(12, 0, 0, 0);
    wanted["Week " + getISOWeek_(m)] = m;
  }

  // Delete week tabs whose whole week is before this Monday.
  guideSS.getSheets().forEach(sheet => {
    const nm = sheet.getName();
    if (!nm.startsWith("Week ") || wanted[nm]) return;
    try {
      const dv = sheet.getDataRange().getDisplayValues();
      const year = extractYear_(String(dv[0][0]));
      const dayRow = dv[2] || [];
      let minDate = null;
      for (let c = 1; c < dayRow.length; c++) {
        if (!dayRow[c]) continue;
        const d = parseDateFromHeader_(String(dayRow[c]), year);
        if (!minDate || d < minDate) minDate = d;
      }
      if (minDate && dateOnly_(minDate) < mondayThis) guideSS.deleteSheet(sheet);
    } catch (e) { /* leave sheet if unreadable */ }
  });

  // Create any missing wanted week tab.
  Object.keys(wanted).forEach(name => {
    if (guideSS.getSheetByName(name)) return;
    const monday = wanted[name];
    const sheet = guideSS.insertSheet(name);
    const weekDates = [];
    for (let d = 0; d < 7; d++) {
      const dt = new Date(monday); dt.setDate(monday.getDate() + d); dt.setHours(12, 0, 0, 0);
      weekDates.push({ dateObj: dt, dateKey: formatDate_(dt), dayName: fullDayName_(dt), shortLabel: shortDateLabel_(dt) });
    }
    rebuildAvailabilityWeekSheet_(sheet, weekDates, scheduleRules, guideNames, {});
  });
}

/** ISO-8601 week number (matches "Week 27", "Week 28", ...). */
function getISOWeek_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - dayNum + 3);
  const firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((d - firstThursday) / 86400000 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
}


/* ---------- write schedule tabs ---------- */

function writeDetailedSchedule_(controlSS, assignedShifts) {
  const output = [[
    "Week", "Date", "Day", "Time", "Language", "Guides needed",
    "Available guides", "Assigned guide 1", "Assigned guide 2", "Assigned guide 3", "Status", "Notes"
  ]];

  assignedShifts.forEach(shift => {
    output.push([
      shift.week, shift.dateText, shift.day, shift.time, shift.language, shift.guidesNeeded,
      shift.eligibleGuides.join(", "),
      shift.assignedGuides[0] || "", shift.assignedGuides[1] || "", shift.assignedGuides[2] || "",
      shift.status, shift.notes
    ]);
  });

  const sheet = controlSS.getSheetByName("Schedule") || controlSS.insertSheet("Schedule");
  sheet.clear();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  formatScheduleSheet_(sheet, output.length, output[0].length);
}

function makeLanguageScheduleTabs_(controlSS, assignedShifts) {
  // Regenerate EVERY Schedule_<Language> tab, including ones with no shifts
  // in the window. A tab that is not rewritten would keep showing stale
  // assignments (this is exactly how "Albert on German" ghosts survived).
  const withShifts = [...new Set(assignedShifts.map(s => s.language))];
  const existingTabs = controlSS.getSheets()
    .map(sh => sh.getName())
    .filter(n => n.indexOf('Schedule_') === 0)
    .map(n => n.substring('Schedule_'.length).trim());
  const languages = [...new Set(withShifts.concat(existingTabs))].filter(Boolean).sort();

  languages.forEach(language => {
    const languageShifts = assignedShifts.filter(s => s.language === language);
    makeOneLanguageScheduleTab_(controlSS, language, languageShifts);
  });
}

function makeOneLanguageScheduleTab_(controlSS, language, shifts) {
  const tabName = `Schedule_${language}`;
  let sheet = controlSS.getSheetByName(tabName);
  if (!sheet) sheet = controlSS.insertSheet(tabName);

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearFormats();

  if (shifts.length === 0) {
    // Cleared on purpose: no tours in the window -> nothing stale survives.
    sheet.getRange(1, 1).setValue(`${language}: no tours in the scheduling window`)
      .setFontStyle('italic').setFontColor('#94a3b8');
    return;
  }

  /* ---------- columns ----------
   * Regular tours: one column per start time ("11:00").
   * Private tours: one EXTRA column per private group ("10:00 · Private",
   * "10:00 · Private 2", ...) — a private tour is its own shift, never mixed
   * into a regular column. Columns are sorted by time, regular first.
   */
  const colMap = new Map();   // colId -> {time, isPrivate, index, minutes}
  shifts.forEach(s => {
    const t = normalizeTime_(s.time);
    const isPriv = !!s.isPrivate;
    const idx = isPriv ? (Number(s.privIndex) || 1) : 1;
    const colId = isPriv ? t + '|P' + idx : t + '|R';
    if (!colMap.has(colId)) {
      colMap.set(colId, { time: t, isPrivate: isPriv, index: idx, minutes: timeToMinutes_(t) });
    }
  });
  const cols = [...colMap.entries()]
    .sort((a, b) => a[1].minutes - b[1].minutes ||
                    (a[1].isPrivate ? 1 : 0) - (b[1].isPrivate ? 1 : 0) ||
                    a[1].index - b[1].index);

  const dates = [...new Set(shifts.map(s => s.dateText))].sort();   // yyyy-MM-dd
  const title = `${language} schedule (${dates[0]} to ${dates[dates.length - 1]})`;

  // dateText -> colId -> {text, boldNames[], conflict}
  const cells = {};
  shifts.forEach(shift => {
    const t = normalizeTime_(shift.time);
    const colId = shift.isPrivate ? t + '|P' + (Number(shift.privIndex) || 1) : t + '|R';
    const assigned = shift.assignedGuides.join(', ');
    let text = assigned;
    if (shift.status !== 'OK') text = assigned ? `${assigned}\n${shift.status}` : shift.status;
    if (!cells[shift.dateText]) cells[shift.dateText] = {};
    cells[shift.dateText][colId] = {
      text,
      boldNames: (shift.lockedGuides || []).slice(),
      conflict: !!shift.hasLockConflict
    };
  });

  const totalCols = 1 + cols.length;

  sheet.getRange(1, 1, 1, totalCols).merge().setValue(title)
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center')
    .setBackground('#2563eb').setFontColor('#ffffff');

  const headerRow = ['Date'].concat(cols.map(c => gridHeaderForColumn_(c[1].time, c[1].isPrivate, c[1].index)));
  sheet.getRange(2, 1, 1, totalCols).setValues([headerRow])
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#bfdbfe');
  // Tint private column headers so they read as their own shifts.
  cols.forEach((c, i) => {
    if (c[1].isPrivate) sheet.getRange(2, i + 2).setBackground('#fde68a');
  });

  const table = dates.map(dt => {
    const label = Utilities.formatDate(new Date(dt + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE MMM d');
    return [label].concat(cols.map(c => (cells[dt] && cells[dt][c[0]]) ? cells[dt][c[0]].text : ''));
  });

  const tableRange = sheet.getRange(3, 1, table.length, totalCols);
  tableRange.setValues(table).setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle').setWrap(true);

  sheet.getRange(3, 1, table.length, 1).setFontWeight('bold').setBackground('#dbeafe');
  sheet.getRange(3, 2, table.length, totalCols - 1).setHorizontalAlignment('center').setBackground('#f8fbff');

  for (let r = 0; r < table.length; r++) {
    const dt = dates[r];
    for (let ci = 0; ci < cols.length; ci++) {
      const cellInfo = cells[dt] && cells[dt][cols[ci][0]];
      if (!cellInfo) continue;
      const cell = sheet.getRange(3 + r, ci + 2);
      const cellText = cellInfo.text;

      if (cellText.includes('Not assigned')) {
        cell.setFontColor('#94a3b8').setFontStyle('italic');   // muted, no red
      }
      // Re-apply BOLD to manager-locked names so the lock survives the rebuild
      // and stays readable as a lock next run.
      if (cellText && cellInfo.boldNames.length) {
        let builder = SpreadsheetApp.newRichTextValue().setText(cellText);
        cellInfo.boldNames.forEach(nm => {
          let idx = cellText.toLowerCase().indexOf(nm.toLowerCase());
          while (idx !== -1) {
            builder = builder.setTextStyle(idx, idx + nm.length,
              SpreadsheetApp.newTextStyle().setBold(true).build());
            idx = cellText.toLowerCase().indexOf(nm.toLowerCase(), idx + nm.length);
          }
        });
        cell.setRichTextValue(builder.build());
      }
      if (cellInfo.conflict) cell.setBackground('#fde2e2');   // visible conflict flag
    }
  }

  sheet.setRowHeights(3, table.length, 42);
  sheet.autoResizeColumns(1, totalCols);
  sheet.setColumnWidth(1, 120);
  for (let c = 2; c <= totalCols; c++) sheet.setColumnWidth(c, 150);
}


/* ---------- readers ---------- */

function readGuides_(ss) {
  const sheet = ss.getSheetByName("Guides");
  if (!sheet) throw new Error("Guides tab not found.");
  const values = sheet.getDataRange().getValues();
  const guides = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = row[0];
    if (!name) continue;
    guides.push({
      name: String(name).trim(),
      active: row[1] === true,
      seniority: Number(row[2]) || 999,
      order: r,
      languages: { English: row[3] === true, German: row[4] === true, Spanish: row[5] === true, French: row[6] === true }
    });
  }
  return guides;
}

function readActiveGuideNames_(ss) {
  return readGuides_(ss).filter(g => g.active).map(g => g.name);
}

function readWeeklySchedule_(ss) {
  const sheet = ss.getSheetByName("Weekly_Schedule");
  if (!sheet) throw new Error("Weekly_Schedule tab not found.");
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  const rules = [];
  for (let r = 1; r < values.length; r++) {
    const rawRow = values[r], displayRow = displayValues[r];
    const day = String(displayRow[0] || "").trim();
    // Sheets sometimes coerces a "10:00" cell into a Date (epoch 1899, shown
    // as 12/30/1899). Fall back to the raw Date's hours/minutes so the rule
    // survives instead of silently disappearing.
    const time = normalizeTime_(displayRow[1]) || timeFromCellValue_(rawRow[1]);
    const language = String(displayRow[2] || "").trim();
    const guidesNeeded = Number(rawRow[3]) || 1;
    if (!day || !time || !language) continue;
    rules.push({
      day, time, language, guidesNeeded,
      activeFrom: rawRow[4] ? dateOnly_(new Date(rawRow[4])) : null,
      activeUntil: rawRow[5] ? dateOnly_(new Date(rawRow[5])) : null
    });
  }
  return rules;
}

function readAvailability_(guideSS, weekName) {
  const all = [];
  const sheet = guideSS.getSheetByName(weekName);
  if (!sheet) throw new Error(`Availability tab not found: ${weekName}`);
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  const year = extractYear_(String(displayValues[0][0]));
  const dayRow = displayValues[2], timeRow = displayValues[3];
  let currentDayLabel = "";
  for (let c = 1; c < dayRow.length; c++) {
    if (dayRow[c]) currentDayLabel = String(dayRow[c]);
    const time = normalizeTime_(timeRow[c]);
    if (!currentDayLabel || !time) continue;
    const dateObj = parseDateFromHeader_(currentDayLabel, year);
    const dateTimeObj = combineDateAndTime_(dateObj, time);
    const dateText = formatDate_(dateObj);
    const day = fullDayName_(dateObj);
    for (let r = 4; r < values.length; r++) {
      const guideName = values[r][0];
      const checked = values[r][c] === true;
      if (!guideName || !checked) continue;
      all.push({ week: weekName, guideName: String(guideName).trim(), dateObj, dateTimeObj, dateText, day, time });
    }
  }
  return all;
}

function readGuideCalendar_(guideSS, weekName) {
  const slots = [];
  const seen = new Set();
  const sheet = guideSS.getSheetByName(weekName);
  if (!sheet) throw new Error(`Availability tab not found: ${weekName}`);
  const displayValues = sheet.getDataRange().getDisplayValues();
  const year = extractYear_(String(displayValues[0][0]));
  const dayRow = displayValues[2], timeRow = displayValues[3];
  let currentDayLabel = "";
  for (let c = 1; c < dayRow.length; c++) {
    if (dayRow[c]) currentDayLabel = String(dayRow[c]);
    const time = normalizeTime_(timeRow[c]);
    if (!currentDayLabel || !time) continue;
    const dateObj = parseDateFromHeader_(currentDayLabel, year);
    const dateTimeObj = combineDateAndTime_(dateObj, time);
    const dateText = formatDate_(dateObj);
    const day = fullDayName_(dateObj);
    const key = `${dateText}|${day}|${time}`;
    if (seen.has(key)) continue;
    seen.add(key);
    slots.push({ week: weekName, dateObj, dateTimeObj, dateText, day, time });
  }
  return slots;
}

/**
 * Build the tour shifts for a week.
 *
 * IMPORTANT DESIGN RULE: tours come from Weekly_Schedule RULES × DATES, never
 * from the availability sheet's column layout. The old version only created a
 * tour when the Week tab happened to have a matching time column — so when the
 * offer changed (e.g. English 11:00 added) before the tabs were re-synced, the
 * tour silently disappeared from the schedule. Now a tour ALWAYS exists; if
 * nobody ticked availability for it, it simply comes out "Not assigned".
 */
function buildShifts_(availability, weeklySchedule, guideCalendar) {
  const availabilityMap = {};
  availability.forEach(a => {
    const key = `${a.dateText}|${a.day}|${a.time}`;
    if (!availabilityMap[key]) availabilityMap[key] = [];
    if (!availabilityMap[key].includes(a.guideName)) availabilityMap[key].push(a.guideName);
  });

  // Unique DATES covered by this week tab (its columns no longer matter).
  const dates = new Map();
  guideCalendar.forEach(slot => {
    if (!dates.has(slot.dateText)) {
      dates.set(slot.dateText, { dateObj: slot.dateObj, day: slot.day, week: slot.week });
    }
  });

  const shifts = [];
  dates.forEach((info, dateText) => {
    weeklySchedule.forEach(rule => {
      if (rule.day !== info.day) return;
      const slotDateOnly = dateOnly_(info.dateObj);
      if (rule.activeFrom && slotDateOnly < rule.activeFrom) return;
      if (rule.activeUntil && slotDateOnly > rule.activeUntil) return;
      shifts.push({
        week: info.week, dateObj: info.dateObj,
        dateTimeObj: combineDateAndTime_(info.dateObj, rule.time),
        dateText, day: info.day, time: rule.time,
        language: rule.language, guidesNeeded: rule.guidesNeeded,
        availableGuides: availabilityMap[`${dateText}|${info.day}|${rule.time}`] || []
      });
    });
  });
  return shifts;
}


function countLanguageSpeakers_(guides) {
  const counts = {};
  LANGUAGES.forEach(language => counts[language] = 0);
  guides.forEach(guide => {
    if (!guide.active) return;
    LANGUAGES.forEach(language => { if (guide.languages[language] === true) counts[language]++; });
  });
  return counts;
}

function readWeekDates_(sheet) {
  const displayValues = sheet.getDataRange().getDisplayValues();
  const title = String(displayValues[0][0] || "");
  const year = extractYear_(title);
  const dayRow = displayValues[2] || [];
  const dates = [];
  const seen = new Set();
  for (let c = 1; c < dayRow.length; c++) {
    const label = dayRow[c];
    if (!label) continue;
    const dateObj = parseDateFromHeader_(label, year);
    const key = formatDate_(dateObj);
    if (!seen.has(key)) {
      seen.add(key);
      dates.push({ dateObj, dateKey: key, dayName: fullDayName_(dateObj), shortLabel: shortDateLabel_(dateObj) });
    }
  }
  return dates;
}

function readExistingAvailability_(sheet) {
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  const title = String(displayValues[0][0] || "");
  const year = extractYear_(title);
  const dayRow = displayValues[2] || [], timeRow = displayValues[3] || [];
  const existing = {};
  let currentDayLabel = "";
  for (let c = 1; c < dayRow.length; c++) {
    if (dayRow[c]) currentDayLabel = String(dayRow[c]);
    const time = normalizeTime_(timeRow[c]);
    if (!currentDayLabel || !time) continue;
    const dateObj = parseDateFromHeader_(currentDayLabel, year);
    const dateKey = formatDate_(dateObj);
    for (let r = 4; r < values.length; r++) {
      const guideName = values[r][0];
      const checked = values[r][c] === true;
      if (!guideName || !checked) continue;
      existing[`${String(guideName).trim()}|${dateKey}|${time}`] = true;
    }
  }
  return existing;
}

function rebuildAvailabilityWeekSheet_(sheet, weekDates, scheduleRules, guideNames, existingAvailability) {
  const slots = [];
  weekDates.forEach(dateInfo => {
    const dateOnly = dateOnly_(dateInfo.dateObj);
    const times = scheduleRules
      .filter(rule => rule.day === dateInfo.dayName)
      .filter(rule => !rule.activeFrom || dateOnly >= rule.activeFrom)
      .filter(rule => !rule.activeUntil || dateOnly <= rule.activeUntil)
      .map(rule => rule.time);
    const uniqueTimes = [...new Set(times)].sort((a, b) => timeToMinutes_(a) - timeToMinutes_(b));
    uniqueTimes.forEach(time => slots.push({ dateKey: dateInfo.dateKey, dayLabel: dateInfo.shortLabel, time }));
  });

  const totalCols = 1 + slots.length;
  const guideRows = Math.max(AVAILABILITY_MIN_GUIDE_ROWS, guideNames.length + 5);

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.clear();

  const title = `Availability: ${sheet.getName()} (${weekDates[0].shortLabel} - ${weekDates[weekDates.length - 1].shortLabel}, ${weekDates[0].dateObj.getFullYear()})`;

  sheet.getRange(1, 1, 1, totalCols).merge().setValue(title)
    .setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  sheet.getRange(2, 1, 1, totalCols).merge()
    .setValue("Select your name in column A, then tick the shifts you can work. This is availability only.")
    .setBackground("#dbeafe").setFontStyle("italic").setHorizontalAlignment("center");

  sheet.getRange("A3:A4").merge().setValue("Name")
    .setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  let col = 2;
  while (col <= totalCols) {
    const dayLabel = slots[col - 2].dayLabel;
    let width = 1;
    while (col - 2 + width < slots.length && slots[col - 2 + width].dayLabel === dayLabel) width++;
    const dayRange = sheet.getRange(3, col, 1, width);
    if (width > 1) dayRange.merge();
    dayRange.setValue(dayLabel).setBackground("#2563eb").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    col += width;
  }

  const times = slots.map(slot => slot.time);
  if (times.length > 0) {
    sheet.getRange(4, 2, 1, times.length).setValues([times])
      .setBackground("#bfdbfe").setFontWeight("bold").setHorizontalAlignment("center");
  }

  const nameValues = [];
  for (let r = 0; r < guideRows; r++) nameValues.push([guideNames[r] || ""]);
  const nameRange = sheet.getRange(AVAILABILITY_FIRST_ROW, 1, guideRows, 1);
  nameRange.clearDataValidations();
  nameRange.setValues(nameValues);

  const guideRule = SpreadsheetApp.newDataValidation().requireValueInList(guideNames, true).setAllowInvalid(false).build();
  nameRange.setDataValidation(guideRule);

  if (slots.length > 0) {
    const checkboxRange = sheet.getRange(AVAILABILITY_FIRST_ROW, 2, guideRows, slots.length);
    checkboxRange.insertCheckboxes();
    const checkboxValues = [];
    for (let r = 0; r < guideRows; r++) {
      const guideName = guideNames[r] || "";
      const row = [];
      slots.forEach(slot => row.push(existingAvailability[`${guideName}|${slot.dateKey}|${slot.time}`] === true));
      checkboxValues.push(row);
    }
    checkboxRange.setValues(checkboxValues);
  }

  sheet.getRange(3, 1, guideRows + 2, totalCols).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(3, 1, guideRows + 2, 1).setBorder(null, null, null, true, null, null, "#1e3a8a", SpreadsheetApp.BorderStyle.SOLID_THICK);

  let currentCol = 2;
  while (currentCol <= totalCols && slots.length > 0) {
    const dayLabel = slots[currentCol - 2]?.dayLabel;
    let width = 1;
    while (currentCol - 2 + width < slots.length && slots[currentCol - 2 + width].dayLabel === dayLabel) width++;
    const endCol = currentCol + width - 1;
    sheet.getRange(3, endCol, guideRows + 2, 1).setBorder(null, null, null, true, null, null, "#1e3a8a", SpreadsheetApp.BorderStyle.SOLID_THICK);
    currentCol += width;
  }

  sheet.setFrozenRows(4);
  sheet.setColumnWidth(1, 130);
  for (let c = 2; c <= totalCols; c++) sheet.setColumnWidth(c, 85);
}

function formatScheduleSheet_(sheet, numRows, numCols) {
  sheet.getRange(1, 1, 1, numCols).setFontWeight("bold").setBackground("#2563eb").setFontColor("#ffffff").setHorizontalAlignment("center");
  if (numRows > 1) {
    sheet.getRange(2, 1, numRows - 1, numCols).setVerticalAlignment("middle").setWrap(true);
    sheet.getRange(1, 1, numRows, numCols).setBorder(true, true, true, true, true, true);
    const statusRange = sheet.getRange(2, 11, numRows - 1, 1);
    const statusValues = statusRange.getValues();
    for (let r = 0; r < statusValues.length; r++) {
      if (statusValues[r][0] === "Not assigned") {
        sheet.getRange(r + 2, 11, 1, 2).setFontColor("#94a3b8").setFontStyle("italic"); // muted, no red
      }
    }
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, numCols);
}


/* ---------- small helpers ---------- */

function hasConflict_(assignedDates, candidateDate) {
  return assignedDates.some(existing => Math.abs(candidateDate - existing) / (1000 * 60 * 60) < TOUR_DURATION_HOURS);
}

/** "10:00" from a Date-coerced time cell (epoch 1899), else ''. */
function timeFromCellValue_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return v.getHours() + ':' + String(v.getMinutes()).padStart(2, '0');
  }
  return '';
}

function normalizeTime_(value) {
  if (!value) return "";
  const s = String(value).trim();
  const ampm = s.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/i);
  if (ampm) {
    let hour = Number(ampm[1]);
    const minute = Number(ampm[2] || 0);
    const suffix = ampm[3].toUpperCase();
    if (suffix === "PM" && hour !== 12) hour += 12;
    if (suffix === "AM" && hour === 12) hour = 0;
    return `${hour}:${String(minute).padStart(2, "0")}`;
  }
  const simple = s.match(/^(\d{1,2}):(\d{2})$/);
  if (simple) return `${Number(simple[1])}:${simple[2]}`;
  return s;
}

function combineDateAndTime_(dateObj, timeText) {
  const d = new Date(dateObj);
  const parts = String(timeText).split(":");
  d.setHours(Number(parts[0]), Number(parts[1] || 0), 0, 0);
  return d;
}

function dateOnly_(dateObj) {
  const d = new Date(dateObj);
  d.setHours(12, 0, 0, 0);
  return d;
}

function extractYear_(title) {
  const match = String(title).match(/\b(20\d{2})\b/);
  return match ? Number(match[1]) : new Date().getFullYear();
}

function parseDateFromHeader_(label, year) {
  const parts = String(label).trim().split(/\s+/);
  const month = parts[1], day = parts[2];
  return new Date(`${month} ${day}, ${year} 12:00:00`);
}

function fullDayName_(dateObj) { return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "EEEE"); }
function shortDateLabel_(dateObj) { return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "EEE MMM d"); }
function formatDate_(dateObj) { return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
// timeToMinutes_ is defined once for the whole Control project in
// guidePortal.gs (Apps Script shares globals across files).


/* ---------- diagnostics / acceptance checks ---------- */

/**
 * Validate the CURRENT Schedule_<Language> grids against the Guides tab and
 * the overlap rule. Logs every violation and returns the count. Run from the
 * editor after makeSchedule, or any time a manager edits a grid by hand.
 * Catches exactly the class of bugs the portal used to show (Albert on
 * German, one guide on two simultaneous/overlapping tours).
 */
function validateScheduleGrids() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  const guides = readGuides_(controlSS);
  const byName = {};
  guides.forEach(g => { byName[g.name.toLowerCase()] = g; });

  const seen = {};   // guide -> [Date] tour starts
  const problems = [];

  controlSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.indexOf("Schedule_") !== 0) return;
    const language = name.substring("Schedule_".length).trim();
    if (!language || sheet.getLastRow() < 3) return;

    const dv = sheet.getDataRange().getDisplayValues();
    const anchorYear = extractYear_(String(dv[0][0] || ""));
    const timeRow = dv[1] || [];

    for (let r = 2; r < dv.length; r++) {
      const label = String(dv[r][0] || "").trim();
      const m = label.match(/([A-Za-z]{3,})\s+(\d{1,2})$/);
      if (!m) continue;
      const dObj = new Date(`${m[1]} ${m[2]}, ${anchorYear} 12:00:00`);
      if (isNaN(dObj)) continue;

      for (let c = 1; c < timeRow.length; c++) {
        const h = parseGridTimeHeader_(timeRow[c]);
        if (!h) continue;
        const time = h.time;
        const raw = String(dv[r][c] || "").trim();
        if (!raw) continue;
        const start = combineDateAndTime_(dObj, time);

        raw.split("\n").forEach(line => {
          line.replace(/🔒/g, "").replace(/\(private\)/ig, "")
            .split(",").forEach(piece => {
              const nm = piece.trim();
              if (!nm || /not assigned|need \d|lock conflict/i.test(nm)) return;
              const g = byName[nm.toLowerCase()];
              if (!g) { problems.push(`${nm} @ ${formatDate_(dObj)} ${time} ${language}: not in Guides tab`); return; }
              if (g.languages[language] !== true) {
                problems.push(`${nm} @ ${formatDate_(dObj)} ${time} ${language}: does not speak ${language}`);
              }
              if (!seen[nm]) seen[nm] = [];
              seen[nm].forEach(prev => {
                if (Math.abs(start - prev) / 3600000 < ASSIGN_CFG.MIN_SEPARATION_HOURS) {
                  problems.push(`${nm}: overlapping tours ${prev.toISOString().slice(0,16)} and ${start.toISOString().slice(0,16)} (<${ASSIGN_CFG.MIN_SEPARATION_HOURS}h apart)`);
                }
              });
              seen[nm].push(start);
            });
        });
      }
    }
  });

  if (problems.length) {
    problems.forEach(p => console.log("VIOLATION: " + p));
    logScheduleConflicts_(controlSS, problems.map(p => ({
      dateText: "", time: "", language: "", isPrivate: false, guide: "grid check", problems: [p]
    })));
  } else {
    console.log("Schedule grids: no violations found.");
  }
  return problems.length;
}


/* ---------- live validation of MANUAL grid edits ---------- */

/**
 * RUN ONCE: installs the on-edit trigger that watches the Schedule_<Language>
 * grids. From then on, when a manager types a guide name into a grid cell:
 *   1. the cell is automatically made BOLD (= a management lock, so the
 *      Friday regeneration will never wipe a hand edit again), and
 *   2. the assignment is validated immediately — wrong language, unknown
 *      name, or an overlapping tour (<5h) puts a ⚠ note on the cell and a
 *      row in the Errors tab. The edit is KEPT (management decides); the
 *      flag just makes the problem impossible to miss.
 */
function setupScheduleEditValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'handleScheduleEdit');
  if (!exists) {
    ScriptApp.newTrigger('handleScheduleEdit').forSpreadsheet(ss).onEdit().create();
  }
  Logger.log('Schedule edit validation installed.');
}

function handleScheduleEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const name = sh.getName();
    if (name.indexOf('Schedule_') !== 0) return;
    if (e.range.getRow() < 3 || e.range.getColumn() < 2) return;
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

    const language = name.substring('Schedule_'.length).trim();
    const val = String(e.range.getValue() || '').trim();
    if (!val || /not assigned/i.test(val)) { e.range.setNote(''); return; }

    // A typed name IS a management decision -> lock it (bold, like the
    // portal's assign action), so makeSchedule preserves it.
    e.range.setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(val)
        .setTextStyle(0, val.length, SpreadsheetApp.newTextStyle().setBold(true).build())
        .build());
    e.range.setFontStyle('normal').setFontColor('#1a2b49');

    const problems = validateCellAssignment_(sh, e.range, language, val);
    if (problems.length) {
      e.range.setNote('⚠ ' + problems.join('\n⚠ '));
      logScheduleConflicts_(SpreadsheetApp.getActiveSpreadsheet(), problems.map(p => ({
        dateText: '', time: '', language, isPrivate: false, privIndex: 1,
        guide: val, problems: [p], hard: true
      })));
    } else {
      e.range.setNote('');
    }
  } catch (err) {
    console.error('handleScheduleEdit: ' + err);
  }
}

/** Problems (strings) with putting `namesText` into this grid cell. */
function validateCellAssignment_(sh, range, language, namesText) {
  const problems = [];
  const dv = sh.getDataRange().getDisplayValues();
  const anchor = gridAnchor_(String((dv[0] && dv[0][0]) || ''));
  const header = parseGridTimeHeader_(String((dv[1] || [])[range.getColumn() - 1] || ''));
  const dateKey = gridLabelToKey_(String((dv[range.getRow() - 1] || [])[0] || '').trim(), anchor);
  if (!header || !dateKey) return problems;   // not a data cell we understand

  const startMs = shiftStartMs_(dateKey, timeToMinutes_(header.time));
  const sepMs = ASSIGN_CFG.MIN_SEPARATION_HOURS * 3600000;
  const myKey = dateKey + '|' + timeToMinutes_(header.time) + '|' + language.toLowerCase() +
                '|' + (header.isPrivate ? 'P' + header.index : 'R');

  let busy = {};
  try { busy = buildBusyMap_(readSchedule_()); } catch (err) { /* portal reader unavailable */ }

  namesText.split(/[\n,]/).map(s => s.trim()).filter(Boolean).forEach(nm => {
    if (/not assigned|need \d|lock conflict/i.test(nm)) return;
    const g = findGuideByName_(nm);
    if (!g) { problems.push(nm + ': not in the Guides tab'); return; }
    if (!g.active) problems.push(nm + ' is marked inactive');
    if (g.languages[language] !== true) problems.push(nm + ' does not speak ' + language);
    const b = busy[nm.trim().toLowerCase()] || [];
    const clash = b.find(x => x.k !== myKey && Math.abs(x.ms - startMs) < sepMs);
    if (clash) {
      const parts = clash.k.split('|');
      problems.push(nm + ' already has a tour ' + parts[0] + ' at ' +
        to12h_(minutesToTime_(Number(parts[1]))) + ' (' + parts[2] + ') — less than ' +
        ASSIGN_CFG.MIN_SEPARATION_HOURS + 'h apart');
    }
  });
  return problems;
}


