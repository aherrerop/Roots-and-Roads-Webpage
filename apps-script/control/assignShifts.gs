/* ============================================================================
   SCHEDULER — CRISIS PLAYBOOK  (read this first when the schedule/availability is off)
   ----------------------------------------------------------------------------
   THE MAXIM: NEVER lose or silently overwrite a manager's decision.

   SINGLE SOURCE OF TRUTH: Weekly_Schedule is the ONLY source of availability
   columns and offer slots — there are NO hardcoded slots. Whatever Management puts
   there is exactly what the availability sheet, portal and website derive from.
   Its columns (each an independent control):
     G Guide  H "Hide from availability"  I "Private"  J "Hide from website"
   H and J are INDEPENDENT — H hides from the availability sheet only, J from the
   public website only. Set both to hide from both. (readWeeklySchedule_ parses them.)

   MANAGER LOCKS: a guide name typed in BOLD in a Schedule_<Language> grid cell is a
   lock. makeSchedule keeps it exactly, assigns everyone else around it, and NEVER
   overwrites it. A lock that conflicts (wrong language / unavailable / overlap) is
   KEPT, its cell tinted red, and the conflict written to the Errors tab — never
   silently resolved.

   ── WHEN SOMETHING IS WRONG (Apps Script editor) ──
   • A slot is missing / wrong on the       -> edit Weekly_Schedule (the only source),
       availability sheet, portal or website     then runWeeklyScheduling() to rebuild
   • Regenerate the standard offer safely    -> updateWeeklyScheduleToCurrentOffer()
       (preserves German/private/other rows      (10-col aware: never wipes cols H/I/J)
       and every col H/I/J flag)
   • Rebuild from scratch (WIPES manual rows)-> setupWeeklySchedule()   (last resort)

   ── GOLDEN RULES when editing ──
   • Overlaps are ALLOWED on purpose in the portal's manual assign (management may put
     one guide on two nearby tours to share contacts / retime); the batch scheduler
     still spaces guides MIN_SEPARATION_HOURS apart. Don't re-add a hard block.
   • Any writer of Weekly_Schedule must stay 10-column aware so cols H/I/J survive.
   • `node test/run-tests.js` must stay green (reference check + tests2 Weekly_Schedule).
   ============================================================================ */

/******************************************************
 * ROOTS & ROADS — assignShifts.gs
 * Bind to: Roots_Roads_Control_v1
 *
 * WHAT THIS DOES — generate the Schedule_<Language> grids from guide availability.
 *  • MANAGER LOCKS: a guide name written in BOLD inside a Schedule_<Language> grid
 *    cell is a lock. makeSchedule keeps it exactly (still bold), assigns everyone
 *    else around it, and never overwrites it. A lock that conflicts (wrong language,
 *    unavailable, overlapping tour) is KEPT, its cell tinted red, and the conflict
 *    written to the Control sheet's "Errors" tab — never silently resolved.
 *  • OVERLAP RULE: in the batch scheduler two tours by the same guide must start at
 *    least ASSIGN_CFG.MIN_SEPARATION_HOURS (5) apart.
 *  • SCORING is centralised in ASSIGN_CFG and documented there.
 *  • Availability advantage: a guide who ticks more availability gets a measured
 *    edge (a tiebreaker), to reward giving us more hours.
 ******************************************************/

const GUIDE_FILE_ID = "1ZkF0yDVE5Q7V2XUO5051ojJdTXtsvb8FA2sh15rqf-0";
const BOOKING_SHEET_ID = "1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw";

const LANGUAGES = ["English", "German", "Spanish", "French", "Italian"];
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
  // ---- PLANNING HORIZONS (the two knobs to change how far ahead we plan) ----
  // AVAILABILITY_WEEKS_AHEAD: how many weeks BEYOND the current one the
  //   Guide_Availability file keeps tabs for. Guides tick availability there,
  //   and an offer change is applied to EVERY one of these week tabs.
  // SCHEDULE_WEEKS_AHEAD: how many weeks BEYOND the current one the guide
  //   SCHEDULE (Schedule_<Language> grids) is generated for. 1 = this week +
  //   next week, which is what guides are used to seeing.
  AVAILABILITY_WEEKS_AHEAD: 4,
  SCHEDULE_WEEKS_AHEAD: 1,
  // AVAILABILITY_UNTIL: a FIXED seasonal end date. The Guide_Availability file
  //   keeps week tabs out to whichever is FURTHER — the rolling
  //   AVAILABILITY_WEEKS_AHEAD window OR the week containing this date. Lets us
  //   open availability through the end of a season without a huge rolling
  //   window the rest of the year. Blank or a past date = rolling window only.
  //   Bump the year each season (kept as YYYY-MM-DD text).
  AVAILABILITY_UNTIL: '2026-12-31',

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

  // Must mirror the booking project's MODEL paid sources (and the portal's
  // PAID_SOURCES), or the balancer values a paid shift as free. The SF tags are
  // paid too; their real pay is the low SF rate, but SF is scheduled on its own
  // stackable path, not balanced here, so "paid" is the right class for them.
  PAID_SOURCES: ["Viator", "GetYourGuide", "GYG", "GYG-SF", "Viator-SF", "Airbnb"],

  // Where schedule problems (lock conflicts, missing eligibility) are logged.
  ERRORS_TAB: "Errors"
};


// Back-compat aliases (used across this file).
const TOUR_DURATION_HOURS = ASSIGN_CFG.MIN_SEPARATION_HOURS;
const VALUE_PAID_PER_GUEST = ASSIGN_CFG.VALUE_PAID_PER_GUEST;
const VALUE_FREE_NET_PER_GUEST = ASSIGN_CFG.VALUE_FREE_NET_PER_GUEST;
const VALUE_PRIVATE_FLAT = ASSIGN_CFG.VALUE_PRIVATE_FLAT;
const PAID_SOURCES_ASSIGN = ASSIGN_CFG.PAID_SOURCES;


// Master switch for makeSchedule's AVAILABILITY-BASED auto-assignment. Managers flip it
// from the Control tab ("Automatic scheduling: ON/OFF") or via setAutoScheduling(true/false);
// default ON. When OFF, makeSchedule still seats manager LOCKS and the Weekly_Schedule
// recurring guide (col G) — the fixed weekly pattern — but doesn't fill the other seats from
// availability, and runWeeklyScheduling doesn't email the schedule.
function autoScheduleEnabled_() {
  try {
    const v = PropertiesService.getScriptProperties().getProperty('AUTO_SCHEDULE');
    return String(v == null ? 'ON' : v).trim().toUpperCase() !== 'OFF';
  } catch (e) { return true; }
}
function setAutoScheduling(on) {
  try { PropertiesService.getScriptProperties().setProperty('AUTO_SCHEDULE', on ? 'ON' : 'OFF'); } catch (e) {}
  return 'Automatic scheduling is now ' + (on ? 'ON' : 'OFF') + '.';
}
function enableAutoScheduling()  { return setAutoScheduling(true); }
function disableAutoScheduling() { return setAutoScheduling(false); }

// Which guides may be auto-seated into a staffable shift's remaining seats.
//   AUTO ON  -> every guide who ticked availability for the slot (speaks it, not on
//               vacation, not already locked, no overlap) — the balancer picks among them.
//   AUTO OFF -> ONLY the Weekly_Schedule recurring guide (col G) for the slot, if they can
//               do it; availability is NOT required (it's their standing weekly slot).
// The caller has already excluded cleared / vacation-blocked / private shifts.
function pickAutoEligible_(shift, autoOn, guides, byName, assignedByGuide) {
  const locked = shift.lockedGuides || [];
  // Tour-type gate: an SF shift is only for SF-capable guides, a 3h shift only for
  // 3h-capable ones (Guides tab "Tour" column; blank defaults to 3h).
  const runsType = g => {
    const t = (g && g.tourTypes) || { threeH: true, sf: false };
    return shift.sfExt ? !!t.sf : !!t.threeH;
  };
  const freeOf = g => g && g.active && runsType(g) && g.languages[shift.language] === true &&
    !isGuideOnVacation_(g, shift.dateText) &&
    !locked.some(n => n.toLowerCase() === g.name.toLowerCase()) &&
    // SF is a short, stackable add-on: the same SF guide may run consecutive SF
    // slots, so the separation conflict check (for the 3h tour) is skipped for SF.
    (shift.sfExt || !hasConflict_(assignedByGuide[g.name] || [], shift.dateTimeObj));
  if (!autoOn) {
    const wg = byName[String(shift.weeklyGuide || '').toLowerCase()];
    return freeOf(wg) ? [wg] : [];
  }
  return guides.filter(g => freeOf(g) && shift.availableGuides.includes(g.name));
}


// Generate the Schedule_<Language> grids. Seats manager locks first (PASS 1), then auto-
// assigns the remaining seats from availability by the ASSIGN_CFG scoring (PASS 2), keeping
// every lock and every manager clear. Writes the grids and logs any lock conflict.
function makeSchedule() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  const guideSS = SpreadsheetApp.openById(GUIDE_FILE_ID);

  const guides = readGuides_(controlSS);
  const weeklySchedule = readWeeklySchedule_(controlSS);
  const languageSpeakerCounts = countLanguageSpeakers_(guides);

  const startDate = dateOnly_(new Date());        // today
  const endDate = endOfScheduleRange_();          // Sunday of next full week
  const weekNames = weekTabsToSchedule_(guideSS); // this week's tab + next week's tab

  // Private slots live in Weekly_Schedule flagged Private (col I) so the
  // availability tab shows their times (e.g. 10:00) and guides can tick them.
  // They must NOT generate empty regular tours, so regular shift-building skips
  // them; private shifts are built from actual private bookings below.
  const regularRules = weeklySchedule.filter(r => !r.isPrivate);

  // Availability index across the scheduled weeks, keyed "dateText|time", so a
  // private booking can be staffed at ITS OWN time even with no regular slot.
  const availIndex = {};
  let shifts = [];
  weekNames.forEach(weekName => {
    const guideCalendar = readGuideCalendar_(guideSS, weekName);
    let availability = readAvailability_(guideSS, weekName);
    // Merge the SF availability for the SAME week from the SEPARATE SF file: SF
    // guides tick their own sheet, and SF shifts (built from the SF Weekly_Schedule
    // rules below) draw from it. SF and 3h ticks share date|time keys, but the
    // tour-type gate in pickAutoEligible_ keeps each type's guides on its own
    // shifts, so the merge can never cross-staff.
    const sfSS = availabilityFile_(true);
    if (sfSS && sfSS !== guideSS && sfSS.getSheetByName(weekName)) {
      availability = availability.concat(readAvailability_(sfSS, weekName));
    }
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
  // MANAGER CLEARS: slots the manager explicitly emptied stay empty (no default,
  // no auto-assign) — the same authority as a lock, but to nobody.
  const clearedSlots = readClearedSlots_(controlSS);

  // Value of each shift to the guide who runs it (private = flat; paid = /guest;
  // free = ~net tip/guest). This drives the balancing.
  const shiftValues = readShiftValues_();
  shifts.forEach(s => {
    const key = s.dateText + "|" + normalizeTime_(s.time) + "|" + s.language + "|" + (s.sfExt ? 'SF' : 'R');
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
    const lk = lockKey_(shift.dateText, normalizeTime_(shift.time), shift.language, shift.isPrivate, shift.privIndex, shift.sfExt);
    const lockedNames = locks[lk] || [];
    shift.lockedGuides = [];
    lockedNames.forEach(nm => {
      const g = byName[nm.toLowerCase()];
      // VACATION WINS over a standing lock: a guide pinned to their regular slot
      // is NOT staffed on a date they are on holiday. Drop the lock for that date
      // so a substitute is auto-assigned below (PASS 2), and note it so the
      // manager sees why their usual guide isn't on it.
      if (g && isGuideOnVacation_(g, shift.dateText)) {
        (shift.vacationDropped = shift.vacationDropped || []).push(nm);
        return;
      }
      const problems = [];
      if (!g) problems.push('not in Guides tab');
      else {
        if (!g.active) problems.push('inactive');
        if (g.languages[shift.language] !== true) problems.push('does not speak ' + shift.language);
        if (!shift.availableGuides.includes(g.name)) problems.push('not marked available');
        // SF is a short, stackable add-on (consecutive SF slots are fine), so the 5h
        // separation does not apply to an SF lock. 3h locks keep the overlap check.
        if (!shift.sfExt && hasConflict_(assignedByGuide[g.name] || [], shift.dateTimeObj)) {
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
          isPrivate: !!shift.isPrivate, privIndex: shift.privIndex, sfExt: !!shift.sfExt, guide: nm,
          problems, hard
        });
      }
    });
  });

  // PASS 2 — auto-assign the remaining seats (availability-based).
  const autoOn = autoScheduleEnabled_();   // Control-tab master switch (default ON)
  const assignedShifts = [];
  order.forEach(shift => {
    // A manager-cleared slot with no lock stays empty: skip auto-assignment.
    const lkC = lockKey_(shift.dateText, normalizeTime_(shift.time), shift.language, shift.isPrivate, shift.privIndex, shift.sfExt);
    shift.cleared = !!clearedSlots[lkC] && !(shift.lockedGuides && shift.lockedGuides.length);
    // The regular guide is on vacation -> leave THEIR tour UNASSIGNED (do not sub
    // someone else in); the manager restaffs it. Same "no auto-staff" as a clear.
    const vacBlocked = !!(shift.vacationDropped && shift.vacationDropped.length) &&
                       !(shift.lockedGuides && shift.lockedGuides.length);
    // Never auto-assign a PRIVATE tour from availability — a private takes ONLY its
    // Weekly_Schedule default guide, applied by the portal. Cleared / vacation / private
    // never auto-staff at all.
    const blocked = shift.cleared || vacBlocked || shift.isPrivate;
    const eligible = blocked ? [] : pickAutoEligible_(shift, autoOn, guides, byName, assignedByGuide);
    const need = blocked ? 0 : Math.max(0, (shift.guidesNeeded || 1) - shift.lockedGuides.length);
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
      // Re-check overlap: an earlier pick this shift can't collide (same start), but a
      // guide may have gained a lock meanwhile. SF is a short, stackable add-on (one guide
      // runs consecutive SF slots), so the 5h separation is skipped for SF here too —
      // matching pickAutoEligible_.
      const pick = pool.shift();
      if (!shift.sfExt && hasConflict_(assignedByGuide[pick.name], shift.dateTimeObj)) continue;
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
      !!c.sfExt === !!shift.sfExt &&
      (Number(c.privIndex) || 1) === (Number(shift.privIndex) || 1));

    assignedShifts.push({
      ...shift,
      eligibleGuides: eligible.map(g => g.name),
      assignedGuides: shift.lockedGuides.concat(assigned.map(g => g.name)),
      lockedGuides: shift.lockedGuides,
      hasLockConflict: hasConflictFlag,
      // A manager-cleared slot keeps the "(cleared)" marker so it survives this
      // rebuild AND is recognised as cleared on the next run (readClearedSlots_).
      status: shift.cleared ? "Not assigned (cleared)" : (ok ? "OK" : "Not assigned"),
      notes: [
        shift.sfExt ? "SF" : "",
        shift.isPrivate ? "Private" : "",
        (shift.vacationDropped && shift.vacationDropped.length)
          ? (shift.vacationDropped.join(", ") + " on vacation") : "",
        shift.extra ? "Extra tour (not in Weekly_Schedule)" : "",
        shift.lockedGuides.length ? "Locked: " + shift.lockedGuides.join(", ") : "",
        hasConflictFlag ? "LOCK CONFLICT (see Errors tab)" : "",
        shift.cleared ? "Cleared by manager" : (ok ? "" : `Need ${needTotal}, assigned ${totalAssigned}`)
      ].filter(Boolean).join(" · ")
    });
  });

  // Keep manager assignments/clears that fall OUTSIDE this run's window, so the
  // grid rewrite below never wipes a tour staffed weeks ahead.
  preserveManagerGridState_(assignedShifts, locks, clearedSlots, startDate, byName);

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


// Run when the public offer changes: rewrite the English / Spanish / Private rows of
// Weekly_Schedule to the current offer and PRESERVE every other row (German etc. stay
// exactly as hand-maintained). Weekly_Schedule is the ONLY source of availability columns,
// so whatever Management puts here is what the availability sheet shows. The Time column is
// written as text so Sheets can't coerce it into a 12/30/1899 date.
function updateWeeklyScheduleToCurrentOffer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Weekly_Schedule');
  if (!sh) throw new Error('Weekly_Schedule tab not found');

  // This function OWNS only the regular English/Spanish 3h offer. Every other row
  // is preserved: other languages (German, Italian, French …), every private row,
  // AND every SF (Sagrada exterior) row. We keep Guide, Hide from availability,
  // Private, Tour Type and Hide from website, normalise the Time cell (string or
  // coerced Date), and MIGRATE the legacy model where "Private" was written into
  // Language: blank that label and set Private = yes.
  //
  // We ALWAYS write the 11-column layout (Tour Type included), reading the current
  // sheet BY HEADER so a legacy 10-col sheet migrates cleanly and a sheet that
  // already has Tour Type keeps every row's value.
  const HEADER = ['Day', 'Time', 'Language', 'Tour Type', 'Guides needed', 'Active from',
                  'Active until', 'Guide', 'Hide from availability', 'Private', 'Hide from website'];
  const preserved = [];
  if (sh.getLastRow() > 1) {
    const last = sh.getLastRow() - 1, width = sh.getLastColumn();
    const raw = sh.getRange(2, 1, last, width).getValues();
    const dv = sh.getRange(2, 1, last, width).getDisplayValues();
    const C = weeklyScheduleCols_(sh.getRange(1, 1, 1, width).getDisplayValues()[0] || []);
    const at = (row, idx, fb) => (idx > -1 && idx < row.length) ? row[idx] : fb;
    raw.forEach((r, i) => {
      const langRaw = String(at(dv[i], C.language, at(r, C.language, '')) || '').trim();
      const isPriv = /^(1|true|yes|y|x)$/i.test(String(at(dv[i], C.private, '') || '').trim())
                     || /^private$/i.test(langRaw);
      const isSf = C.tourType > -1 && /^\s*sf\b/i.test(String(at(dv[i], C.tourType, '') || '').trim());
      // Regular English/Spanish 3h rows are regenerated below — drop them here.
      // (SF English/Spanish rows are preserved: they are a different tour.)
      if (!isPriv && !isSf && ['english', 'spanish'].indexOf(langRaw.toLowerCase()) !== -1) return;
      if (!langRaw && !isPriv && !isSf) return;              // truly empty row
      const time = normalizeTime_(at(dv[i], C.time, '')) || timeFromCellValue_(at(r, C.time, ''));
      const lang = /^private$/i.test(langRaw) ? '' : langRaw; // strip legacy "Private" label
      preserved.push([
        String(at(dv[i], C.day, at(r, C.day, '')) || '').trim(),
        time || String(at(dv[i], C.time, '') || ''),
        lang,
        isSf ? 'SF' : '3h',                                   // Tour Type (col D)
        Number(at(r, C.needed, '')) || (isPriv ? 0 : 1),
        at(r, C.from, '') || '',
        at(r, C.until, '') || '',
        String(at(dv[i], C.guide, '') || '').trim(),          // Guide
        String(at(dv[i], C.hideAvail, '') || '').trim(),      // Hide from availability
        isPriv ? 'yes' : String(at(dv[i], C.private, '') || '').trim(),  // Private
        String(at(dv[i], C.hideWeb, '') || '').trim()         // Hide from website — preserved as-is
      ]);
    });
  }

  const rows = [];
  // Blank "Active from" = always active. The standard weekly offer must show in
  // EVERY availability week. Managed offer rows are the full "3h" tour and leave
  // Guide/Hide/Private blank.
  const add = (days, time, lang) => days.forEach(d => rows.push([d, time, lang, '3h', 1, '', '', '', '', '', '']));
  const MTThF = ['Monday', 'Tuesday', 'Thursday', 'Friday'];

  add(MTThF, '11:00', 'English');
  add(MTThF, '17:00', 'English');
  add(MTThF, '10:30', 'Spanish');
  add(['Wednesday'], '17:00', 'English');
  add(['Saturday'], '17:00', 'English');

  const all = [HEADER].concat(rows).concat(preserved);

  sh.clear();
  // Time column as TEXT before writing: "11:00" can never again become a Date.
  sh.getRange(1, 2, all.length, 1).setNumberFormat('@');
  sh.getRange(1, 1, all.length, HEADER.length).setValues(all);
  sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  Logger.log('Weekly_Schedule updated: ' + rows.length + ' offer rows + ' +
             preserved.length + ' preserved rows (German/SF etc., times repaired).');
}


/* ---------- manager locks (bold names in the grids) ---------- */

function lockKey_(dateText, time, language, isPrivate, privIndex, isSf) {
  return dateText + "|" + time + "|" + language + "|" +
         (isSf ? "SF" : (isPrivate ? "P" + (Number(privIndex) || 1) : "R"));
}

// Parse a Schedule_<Language> column header into { time, isPrivate, isSf, index }:
//   "11:00"             -> { time:"11:00", isPrivate:false, index:1 }
//   "10:00 · Private"   -> { time:"10:00", isPrivate:true,  index:1 }
//   "17:00 · Private 2" -> { time:"17:00", isPrivate:true,  index:2 }
//   "10:00 · SF"        -> { time:"10:00", isSf:true }
// Tolerates display forms like "11:00 AM"; returns null for a non-time header. Shared by the
// grid writer, the lock reader and the guide portal.
function parseGridTimeHeader_(text) {
  const s = String(text || '').trim();
  if (!s) return null;
  // "11:00" · "10:00 · Private" · "17:00 · Private 2" · "10:00 · SF" (Sagrada exterior).
  const m = s.match(/^(\d{1,2}:\d{2}(?:\s*[AP]M)?)(?:\s*[·\-]\s*(Private|SF)(?:\s*(\d+))?)?$/i);
  if (!m) return null;
  const t = normalizeTime_(m[1]);
  if (!/^\d{1,2}:\d{2}$/.test(t)) return null;
  const variant = (m[2] || '').toUpperCase();
  return {
    time: t,
    isPrivate: variant === 'PRIVATE',
    isSf: variant === 'SF',
    index: m[3] ? Number(m[3]) : 1
  };
}

// The column header text for a shift column. SF = the Sagrada exterior add-on's own column.
function gridHeaderForColumn_(time, isPrivate, index, isSf) {
  if (isSf) return time + ' · SF';
  if (!isPrivate) return time;
  return time + ' · Private' + (Number(index) > 1 ? ' ' + index : '');
}
// The internal column id used by the grid writer (mirrors the header variants).
function gridColId_(time, isPrivate, index, isSf) {
  return isSf ? time + '|SF' : (isPrivate ? time + '|P' + (Number(index) || 1) : time + '|R');
}

// Read every BOLD guide name from every Schedule_<Language> grid — bold = a management lock.
// Returns { lockKey -> [names] }. A cell can hold a regular line and 🔒-prefixed private
// lines; a bold run locks only the names inside that run, so a manager can bold one of two.
function readLockedAssignments_(controlSS) {
  const locks = {};
  controlSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const parsed = parseScheduleTabName_(name);       // Schedule_<lang> + Schedule_SF_<lang>
    if (!parsed) return;
    const language = parsed.language, tabIsSf = parsed.isSf;
    if (!language || sheet.getLastRow() < 3) return;

    const dv = sheet.getDataRange().getDisplayValues();
    // The same window-aware resolver the portal uses, so a locked "Jul 31" row maps to the
    // exact same date the portal shift does (no year drift between the two).
    const anchor = gridAnchor_(String(dv[0][0] || ""));
    const timeRow = dv[1] || [];
    const rich = sheet.getDataRange().getRichTextValues();

    for (let r = 2; r < dv.length; r++) {
      const label = String(dv[r][0] || "").trim();
      if (!label) continue;
      const dateText = gridLabelToKey_(label, anchor);   // "yyyy-MM-dd"
      if (!dateText) continue;

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
          // marker still wins for grids written before the column redesign. An SF
          // grid tab is never private.
          const colSf = tabIsSf || h.isSf;
          const isPriv = !colSf && (h.isPrivate || /🔒|\(private\)/i.test(txt));
          const idx = isPriv ? h.index : 1;
          txt.split(/[\n,]/).forEach(piece => {
            let nm = piece.replace(/🔒/g, "").replace(/\(private\)/ig, "").trim();
            if (!nm || /not assigned|lock conflict|need \d/i.test(nm)) return;
            const k = lockKey_(dateText, h.time, language, isPriv, idx, colSf);
            if (!locks[k]) locks[k] = [];
            if (locks[k].indexOf(nm) === -1) locks[k].push(nm);
          });
        });
      }
    }
  });
  return locks;
}

// Read every MANAGER-CLEARED slot from the Schedule_<Language> grids. A clear is the cell
// text "Not assigned (cleared)" written by the portal; makeSchedule uses this to leave those
// slots empty instead of auto-assigning — the manager's clear wins, like a lock. Returns a
// set of lockKey_ strings. (A plain "Not assigned" cell is NOT cleared — it still gets
// auto-assigned / defaulted.)
function readClearedSlots_(controlSS) {
  const cleared = {};
  controlSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const parsed = parseScheduleTabName_(name);
    if (!parsed) return;
    const language = parsed.language, tabIsSf = parsed.isSf;
    if (!language || sheet.getLastRow() < 3) return;

    const dv = sheet.getDataRange().getDisplayValues();
    const anchor = gridAnchor_(String(dv[0][0] || ''));
    const timeRow = dv[1] || [];
    for (let r = 2; r < dv.length; r++) {
      const label = String(dv[r][0] || '').trim();
      if (!label) continue;
      const dateText = gridLabelToKey_(label, anchor);
      if (!dateText) continue;
      for (let c = 1; c < timeRow.length; c++) {
        const h = parseGridTimeHeader_(timeRow[c]);
        if (!h) continue;
        if (!/cleared/i.test(String(dv[r][c] || ''))) continue;
        const colSf = tabIsSf || h.isSf;
        const isPriv = !colSf && h.isPrivate;
        const idx = isPriv ? h.index : 1;
        cleared[lockKey_(dateText, h.time, language, isPriv, idx, colSf)] = true;
      }
    }
  });
  return cleared;
}

// Keep MANAGER assignments (and clears) that fall OUTSIDE this run's scheduling window.
// makeSchedule only builds shifts for today..next-Sunday, and the grid writer clears each
// Schedule_<Language> tab and rewrites it from those shifts — so an assignment a manager made
// to a tour weeks out would be wiped every run. readLockedAssignments_/readClearedSlots_
// already read EVERY grid row, so re-inject any lock/clear the freshly-built assignedShifts
// don't cover (skipping PAST dates). They carry their guide + lock marker, so the rewrite
// puts them back exactly, bold and all.
function preserveManagerGridState_(assignedShifts, locks, clearedSlots, startDate, byName) {
  const have = {};
  assignedShifts.forEach(s => {
    have[lockKey_(s.dateText, normalizeTime_(s.time), s.language, s.isPrivate, s.privIndex, s.sfExt)] = true;
  });
  const addPreserved = (key, names, cleared) => {
    if (have[key]) return;
    const parts = String(key).split('|');
    if (parts.length < 4) return;
    const dateText = parts[0], time = parts[1], language = parts[2], tag = parts[3];
    if (!dateText || !time || !language) return;
    const isSf = tag === 'SF';
    const isPriv = tag.charAt(0) === 'P';
    const idx = isPriv ? (Number(tag.slice(1)) || 1) : 1;
    const dateObj = dateOnly_(new Date(dateText + 'T12:00:00'));
    if (!dateObj || isNaN(dateObj) || dateObj < startDate) return;   // never resurrect a past assignment
    have[key] = true;
    // VACATION WINS over a preserved (out-of-window) lock too: drop any pinned
    // guide who is on holiday that date, so a far-ahead regular tour is never
    // staffed by someone away. If that empties the slot it shows "Not assigned"
    // (restaff) — never the vacationing guide.
    const kept = [], vac = [];
    (names || []).forEach(nm => {
      const g = byName && byName[String(nm).toLowerCase()];
      if (!cleared && g && isGuideOnVacation_(g, dateText)) vac.push(nm); else kept.push(nm);
    });
    const noteParts = [];
    if (cleared) noteParts.push('Cleared by manager');
    else if (kept.length) noteParts.push('Locked: ' + kept.join(', '));
    if (vac.length) noteParts.push(vac.join(', ') + ' on vacation');
    assignedShifts.push({
      week: 'Week ' + getISOWeek_(dateObj),
      dateText: dateText, day: fullDayName_(dateObj), time: time, language: language,
      guidesNeeded: kept.length ? kept.length : 1,
      isPrivate: isPriv, privIndex: idx, sfExt: isSf, tourName: isSf ? 'SF' : '3h',
      dateTimeObj: combineDateAndTime_(dateObj, time),
      eligibleGuides: [],
      assignedGuides: cleared ? [] : kept.slice(),
      lockedGuides: cleared ? [] : kept.slice(),   // re-bold the lock; a clear (or a dropped vacation guide) has no name
      hasLockConflict: false,
      cleared: !!cleared,
      status: cleared ? 'Not assigned (cleared)' : (kept.length ? 'OK' : 'Not assigned'),
      notes: noteParts.join(' · ')
    });
  };
  Object.keys(locks || {}).forEach(k => addPreserved(k, locks[k], false));
  Object.keys(clearedSlots || {}).forEach(k => addPreserved(k, [], true));
}

// Write HARD lock conflicts to the Control sheet's Errors tab (a visible flag).
function logScheduleConflicts_(controlSS, conflicts) {
  // Only HARD conflicts are errors. "Not marked available" alone is a normal management
  // decision (they assign someone after talking to them), so it is not logged.
  const hard = conflicts.filter(c => c.hard !== false);
  if (!hard.length) return;
  let sh = controlSS.getSheetByName(ASSIGN_CFG.ERRORS_TAB);
  if (!sh) {
    sh = controlSS.insertSheet(ASSIGN_CFG.ERRORS_TAB);
    sh.getRange(1, 1, 1, 4).setValues([["Timestamp", "Type", "Details", "Raw data"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  const ts = Utilities.formatDate(new Date(), "Europe/Madrid", "yyyy-MM-dd HH:mm");
  const rows = hard.map(c => [
    ts, "Schedule lock conflict",
    `${c.guide} locked on ${c.dateText} ${c.time} ${c.language}${c.isPrivate ? " (private)" : ""}: ${c.problems.join("; ")}`,
    ""
  ]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}


/* ---------- Scheduling window + private tours ---------- */

// The last day of the scheduling window: Sunday of the last scheduled week.
function endOfScheduleRange_() {
  const today = dateOnly_(new Date());
  const dow = (today.getDay() + 6) % 7;                 // 0=Mon..6=Sun
  const mondayThis = new Date(today); mondayThis.setDate(today.getDate() - dow);
  // Last day of the scheduling window = Sunday of the last scheduled week.
  const lastSunday = new Date(mondayThis);
  lastSunday.setDate(mondayThis.getDate() + 7 * (1 + ASSIGN_CFG.SCHEDULE_WEEKS_AHEAD) - 1);
  return dateOnly_(lastSunday);
}

// The availability "Week …" tab names that cover the scheduling window (this week +
// SCHEDULE_WEEKS_AHEAD). Throws if none is found, so the caller runs Sync availability first.
function weekTabsToSchedule_(guideSS) {
  const today = dateOnly_(new Date());
  const dow = (today.getDay() + 6) % 7;
  const mondayThis = new Date(today); mondayThis.setDate(today.getDate() - dow);

  // Every week the schedule covers (current + SCHEDULE_WEEKS_AHEAD).
  const wanted = [];
  for (let i = 0; i <= ASSIGN_CFG.SCHEDULE_WEEKS_AHEAD; i++) {
    const m = new Date(mondayThis); m.setDate(mondayThis.getDate() + 7 * i);
    wanted.push(formatDate_(m));
  }
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

  if (!names.length) throw new Error("No availability Week tab found for the scheduling window (current + " +
    ASSIGN_CFG.SCHEDULE_WEEKS_AHEAD + " week(s) ahead). Run Sync availability first.");
  return names;
}

// Keep the regular shifts, then add ONE private shift PER PRIVATE BOOKING at its real booked
// time — each its own shift / grid column, never merged into a regular slot. availableGuides
// come from whoever ticked availability at that date+time, or (if the booked time sits
// between availability columns) the nearest slot within PRIVATE_AVAIL_TOLERANCE_MIN. If
// nobody matches, the shift stays "Not assigned" and management staffs it manually.
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

// Safety net: any non-private booking at a date/time/language with NO matching shift (a slot
// missing from Weekly_Schedule, a legacy time, a manually added row) still becomes a tour —
// it shows in the grids and the portal, gets auto-assigned from availability, and comes out
// "Not assigned" when nobody fits, so management knows to find a guide rather than the tour
// silently not existing.
function expandOrphanShifts_(shifts, availIndex, startDate, endDate) {
  // Variant-aware key: an SF shift and a 3h shift at the same slot are distinct,
  // so an SF booking never fills (or gets swallowed by) the 3h column.
  const vkey = s => s.dateText + '|' + normalizeTime_(s.time) + '|' + s.language + '|' + (s.sfExt ? 'SF' : 'R');
  const existing = new Set(shifts.filter(s => !s.isPrivate).map(vkey));
  const out = shifts.slice();

  readActiveBookingSlots_().forEach(slot => {
    const key = slot.dateText + '|' + slot.time + '|' + slot.language + '|' + (slot.sfExt ? 'SF' : 'R');
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
      isPrivate: false, privIndex: 0, sfExt: !!slot.sfExt,
      extra: true                       // flagged in the Schedule tab notes
    });
  });
  return out;
}

// Unique (date, time, language, variant) slots of all NON-private active bookings. An SF
// booking is its OWN slot (sfExt:true), so a 3h and an SF booking at the same date/time/
// language yield two distinct shifts.
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
      const sfExt = isSfExtSource_(String(row[5] || ""));   // col F = Source (GYG-SF / Viator-SF)
      const time = timeFromCellValue_(row[4]) || normalizeTime_(row[4]);
      if (!/^\d{1,2}:\d{2}$/.test(time)) return;
      const dateText = row[3] instanceof Date ? formatDate_(row[3]) : formatDate_(new Date(row[3]));
      const key = dateText + '|' + time + '|' + language + '|' + (sfExt ? 'SF' : 'R');
      if (seen.has(key)) return;
      seen.add(key);
      out.push({ dateText, time, language, sfExt });
    });
  });
  return out;
}

// Availability for a private booking: the exact time, else the nearest slot within the
// tolerance window (same day).
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
      const sfExt = isSfExtSource_(source);               // SF valued on its OWN shift
      const paid = PAID_SOURCES_ASSIGN.some(s => s.toLowerCase() === source.toLowerCase());
      const val = (paid ? VALUE_PAID_PER_GUEST : VALUE_FREE_NET_PER_GUEST) * guests;
      const key = dateKey + "|" + time + "|" + language + "|" + (sfExt ? 'SF' : 'R');
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


/* ---------- Workflow entry points ---------- */

/* ---------- availability: 3h vs SF (SEPARATE files) ----------
 * The 3-hour tour and the Sagrada-exterior (SF) tour each get their OWN
 * availability SPREADSHEET in the same Drive folder — 3h in GUIDE_FILE_ID
 * (Guide_3h_Availability_v1), SF in SF_GUIDE_FILE_ID (Guide_SF_Availability_v1).
 * Files are opened BY ID, so renaming them in Drive changes nothing. Both use
 * plain "Week <N>" tabs; the FILE determines the tour type. "Sync availability"
 * (Control tab) rebuilds BOTH files. Each lists only the guides who run that tour
 * type and only that type's time slots. */

// Dedicated SF availability file (same Drive folder as the 3h file), opened by ID
// so a Drive rename never breaks it. Empty = SF availability disabled.
var SF_GUIDE_FILE_ID = '1XJXz3tR1aDupSbtsmWKqUd-W6Zepv-VJiAZfGw-649o';

function availabilityFileId_(isSf) { return isSf ? SF_GUIDE_FILE_ID : GUIDE_FILE_ID; }

// Open the availability spreadsheet for a tour type, or null if not configured/reachable.
function availabilityFile_(isSf) {
  const id = availabilityFileId_(isSf);
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}

// True for a plain "Week <N>" availability tab (never the legacy "SF Week <N>").
function isAvailWeekTab_(name) {
  const n = String(name || '');
  return n.indexOf('Week ') === 0;   // "Week 42"; the legacy "SF Week 42" starts with "SF", so it is skipped
}

// Remove legacy "SF Week <N>" tabs from the 3h file — SF lives in its own file.
function cleanupStraySfTabs_() {
  try {
    const g = SpreadsheetApp.openById(GUIDE_FILE_ID);
    g.getSheets().forEach(sh => { if (String(sh.getName()).indexOf('SF Week ') === 0) g.deleteSheet(sh); });
  } catch (e) { /* best-effort cleanup */ }
}

// Which guides are listed on an availability sheet.
//   3h sheet: EVERY active guide (the 3h sheet is never narrowed by tour type).
//   SF sheet: only SF-capable guides (Guides "Tour" column). If none are tagged yet, the SF
//             sheet still lists every active guide so it's never blank — mark SF guides to
//             narrow it.
function availabilityGuideNames_(controlSS, isSf) {
  if (!isSf) return readActiveGuideNames_(controlSS);
  const sf = readGuides_(controlSS).filter(g => g.active && g.tourTypes && !!g.tourTypes.sf).map(g => g.name);
  return sf.length ? sf : readActiveGuideNames_(controlSS);
}

// The Weekly_Schedule rules that belong on this tour type's availability sheet.
//   3h: non-SF rules that are NOT "Hide from availability".
//   SF: SF rules (the "Hide from availability" flag is ignored — SF has its own sheet).
// Private rows keep their tick columns on both (private is staffed from availability too).
function availabilityRulesFor_(controlSS, isSf) {
  return readWeeklySchedule_(controlSS).filter(r =>
    (!!r.sfExt === !!isSf) && (isSf || !r.hideFromAvailability));
}

function syncAvailabilityFile() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  syncAvailabilityInFile_(controlSS, false);                               // 3h file
  if (availabilityRulesFor_(controlSS, true).length) syncAvailabilityInFile_(controlSS, true);   // SF file (only when SF runs)
}

// Rebuild every "Week <N>" tab in one tour type's availability file.
function syncAvailabilityInFile_(controlSS, isSf) {
  const guideSS = availabilityFile_(isSf);
  if (!guideSS) return;
  const rules = availabilityRulesFor_(controlSS, isSf);
  const names = availabilityGuideNames_(controlSS, isSf);
  guideSS.getSheets().forEach(sheet => {
    if (!isAvailWeekTab_(sheet.getName())) return;
    const weekDates = readWeekDates_(sheet);
    if (weekDates.length === 0) return;
    const existingAvailability = readExistingAvailability_(sheet);
    rebuildAvailabilityWeekSheet_(sheet, weekDates, rules, names, existingAvailability);
  });
}

// Trigger entry points must NEVER throw: Apps Script emails the owner a "Summary of failures"
// for any trigger run that ends in an exception, and a transient "Service Spreadsheets timed
// out" is normal under contention. Errors are logged to the Control Errors tab and swallowed;
// the next scheduled run picks the work up again.
function safeTriggerRun_(label, fn) {
  try {
    fn();
  } catch (err) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName(ASSIGN_CFG.ERRORS_TAB);
      if (!sh) {
        sh = ss.insertSheet(ASSIGN_CFG.ERRORS_TAB);
        sh.getRange(1, 1, 1, 4).setValues([["Timestamp", "Type", "Details", "Raw data"]]).setFontWeight("bold");
        sh.setFrozenRows(1);
      }
      sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[
        Utilities.formatDate(new Date(), "Europe/Madrid", "yyyy-MM-dd HH:mm"),
        label + " failed (transient; swallowed so the trigger does not email)",
        String(err && err.stack ? err.stack : err).slice(0, 500), ""
      ]]);
    } catch (e) { /* logging must never throw either */ }
    console.error(label + ' error: ' + (err && err.stack ? err.stack : err));
  }
}

// Run from the editor: create/extend the availability week tabs (out to the seasonal end
// date) and refresh their columns. No email, no scheduling. A runnable wrapper (no trailing
// "_") so it shows in the Run menu.
function syncAvailabilityTabs() {
  ensureWeekTabs_();
  syncAvailabilityFile();
}

// The weekly trigger: refresh the week tabs + availability, build the schedule, and (when
// auto is ON) email it. Wrapped in safeTriggerRun_ so a transient error never emails the owner.
function runWeeklyScheduling() {
  safeTriggerRun_('runWeeklyScheduling', function () {
    ensureWeekTabs_();      // delete past weeks, create the upcoming ones
    syncAvailabilityFile();
    makeSchedule();
    // AUTO OFF: build the schedule (weekly-pattern guides are still seated), but do
    // NOT email it — the manager is scheduling by hand and doesn't want the blast.
    if (autoScheduleEnabled_()) emailWeeklySchedule();  // send the language tables to management
  });
}

// Run once (from the editor) to fill Weekly_Schedule with the current offer. Overwrites the
// whole tab (wipes German too) — normally you run updateWeeklyScheduleToCurrentOffer instead,
// which preserves the other rows.
function setupWeeklySchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Weekly_Schedule') || ss.insertSheet('Weekly_Schedule');

  const rows = [['Day', 'Time', 'Language', 'Tour Type', 'Guides needed', 'Active from', 'Active until', 'Guide', 'Hide from availability', 'Private', 'Hide from website']];
  // Blank "Active from" = always active (see updateWeeklyScheduleToCurrentOffer).
  // Blank "Guide" (col G) = no recurring default; fill it to auto-assign a guide
  // to that weekly slot every week (a manager can still override a single date).
  // Blank "Hide from availability" (col H) = shows in the availability sheet; set
  // TRUE/yes/x to keep the tour on the website + portal but off that sheet.
  // "Private" (col I) = yes/x marks a private slot (any language, or blank Language
  // for language-agnostic): shows on the availability sheet, never staged as a
  // group tour, never sent to the website.
  // Blank "Hide from website" (col J) = shows on the public website; set TRUE/yes/x
  // to keep it OFF the website while still on the availability sheet + portal.
  // Cols H and J are INDEPENDENT — management controls each surface separately.
  const add = (days, time, lang) => days.forEach(d => rows.push([d, time, lang, '3h', 1, '', '', '', '', '', '']));

  const MTThF = ['Monday', 'Tuesday', 'Thursday', 'Friday'];
  add(MTThF, '11:00', 'English');
  add(MTThF, '17:00', 'English');
  add(MTThF, '10:30', 'Spanish');
  add(MTThF, '17:00', 'Spanish');
  add(['Wednesday'], '17:00', 'English');
  add(['Wednesday'], '17:00', 'Spanish');
  add(['Saturday'], '17:00', 'English');
  add(['Saturday'], '17:00', 'Spanish');

  // PRIVATE availability slots (language-agnostic) — availability columns only,
  // never staffed (Guides needed = 0), flagged Private = yes. Mon–Fri 10:00 &
  // 17:00, Sat 17:00. To offer a private tour in a specific language, add a row
  // with that Language and Private = yes.
  const MTWThF = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const addPriv = (days, time) => days.forEach(d => rows.push([d, time, '', '3h', 0, '', '', '', '', 'yes', '']));
  addPriv(MTWThF, '10:00');
  addPriv(MTWThF, '17:00');
  addPriv(['Saturday'], '17:00');

  sh.clear();
  sh.getRange(1, 1, rows.length, 11).setValues(rows);
  sh.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

// Email the generated Schedule_<Language> tables to management as HTML, ready to forward to
// guides. Called at the end of runWeeklyScheduling.
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

// Weeks to keep visible ahead of the current one (see ASSIGN_CFG).
const WEEKS_AHEAD = ASSIGN_CFG.AVAILABILITY_WEEKS_AHEAD;

// How many weeks of availability tabs to project ahead of `mondayThis` (a Monday, midnight-
// stripped): the rolling AVAILABILITY_WEEKS_AHEAD minimum, or further when the fixed seasonal
// end date AVAILABILITY_UNTIL reaches past it — so availability can open through a season's
// end without a giant rolling window the rest of the year. Pure (no sheet access), so unit-
// testable.
function availabilityWeeksAhead_(mondayThis) {
  let weeksAhead = WEEKS_AHEAD;
  const untilRaw = ASSIGN_CFG.AVAILABILITY_UNTIL;
  if (untilRaw) {
    const until = dateOnly_(new Date(untilRaw + 'T12:00:00'));
    if (until && !isNaN(until) && until >= mondayThis) {
      const weeksToUntil = Math.floor((until - mondayThis) / (7 * 86400000));
      if (weeksToUntil > weeksAhead) weeksAhead = weeksToUntil;
    }
  }
  return weeksAhead;
}

// Auto-manage the availability Week tabs: delete any week fully in the past and create the
// current week + the upcoming ones if missing. Safe to run repeatedly.
function ensureWeekTabs_() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  ensureWeekTabsInFile_(controlSS, false);                                 // 3h file
  cleanupStraySfTabs_();                                                   // drop legacy "SF Week N" tabs from the 3h file
  if (availabilityRulesFor_(controlSS, true).length) ensureWeekTabsInFile_(controlSS, true);   // SF file (only when SF runs)
}

// Delete past "Week <N>" tabs and create the current + upcoming ones in one tour type's
// availability file.
function ensureWeekTabsInFile_(controlSS, isSf) {
  const guideSS = availabilityFile_(isSf);
  if (!guideSS) return;
  const rules = availabilityRulesFor_(controlSS, isSf);
  const names = availabilityGuideNames_(controlSS, isSf);

  const today = dateOnly_(new Date());
  const dow = (today.getDay() + 6) % 7;
  const mondayThis = new Date(today); mondayThis.setDate(today.getDate() - dow);
  const weeksAhead = availabilityWeeksAhead_(mondayThis);

  // Weeks we want to exist (current + next weeks), by ISO week number -> Monday.
  const wantedWeeks = {};
  for (let i = 0; i <= weeksAhead; i++) {
    const m = new Date(mondayThis); m.setDate(mondayThis.getDate() + 7 * i);
    m.setHours(12, 0, 0, 0);
    wantedWeeks[getISOWeek_(m)] = m;
  }
  const wantedNames = {};
  Object.keys(wantedWeeks).forEach(w => { wantedNames['Week ' + w] = true; });

  // Delete "Week <N>" tabs whose whole week is before this Monday.
  guideSS.getSheets().forEach(sheet => {
    const nm = sheet.getName();
    if (!isAvailWeekTab_(nm) || wantedNames[nm]) return;
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

  // Create missing wanted week tabs. (For the SF file, the caller only gets here
  // when SF actually runs, so rules is non-empty.)
  Object.keys(wantedWeeks).forEach(w => {
    const name = 'Week ' + w;
    if (guideSS.getSheetByName(name)) return;
    const monday = wantedWeeks[w];
    const weekDates = [];
    for (let d = 0; d < 7; d++) {
      const dt = new Date(monday); dt.setDate(monday.getDate() + d); dt.setHours(12, 0, 0, 0);
      weekDates.push({ dateObj: dt, dateKey: formatDate_(dt), dayName: fullDayName_(dt), shortLabel: shortDateLabel_(dt) });
    }
    rebuildAvailabilityWeekSheet_(guideSS.insertSheet(name), weekDates, rules, names, {});
  });
}

// ISO-8601 week number (matches "Week 27", "Week 28", …).
function getISOWeek_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - dayNum + 3);
  const firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((d - firstThursday) / 86400000 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
}


/* ---------- Write the schedule tabs ---------- */

// Write the flat "Schedule" tab: one row per assigned shift (the master list the grids are
// built from).
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

// Rewrite the per-language Schedule_<Language> grids (3h and SF) from the assigned shifts.
function makeLanguageScheduleTabs_(controlSS, assignedShifts) {
  // Regenerate EVERY Schedule_<Language> tab, including ones with no shifts in the window:
  // a tab that isn't rewritten would keep showing stale assignments.
  const withShifts = [...new Set(assignedShifts.map(s => s.language))];
  const existingTabs = controlSS.getSheets()
    .map(sh => parseScheduleTabName_(sh.getName()))     // Schedule_<lang> AND Schedule_SF_<lang>
    .filter(Boolean)
    .map(p => p.language);
  const languages = [...new Set(withShifts.concat(existingTabs))].filter(Boolean).sort();

  languages.forEach(language => {
    const languageShifts = assignedShifts.filter(s => s.language === language);
    // The 3h/private grid and the SF grid are regenerated SEPARATELY, so an SF
    // assignment never lands in Schedule_<lang> and vice versa.
    makeOneLanguageScheduleTab_(controlSS, language, languageShifts.filter(s => !s.sfExt), false);
    const sfShifts = languageShifts.filter(s => s.sfExt);
    // Only touch the SF grid when this language actually runs SF, or an SF tab
    // already exists (so a stale one still gets cleared) — no empty SF tabs for
    // languages that never run the exterior tour.
    if (sfShifts.length || controlSS.getSheetByName(scheduleTabName_(language, true))) {
      makeOneLanguageScheduleTab_(controlSS, language, sfShifts, true);
    }
  });
}

// Build (or rebuild) one Schedule_<Language> grid tab: title, time-column headers, a row per
// date, each cell holding the assigned guide(s) with manager locks in bold.
function makeOneLanguageScheduleTab_(controlSS, language, shifts, isSf) {
  const tabName = scheduleTabName_(language, isSf);
  const label = isSf ? (language + ' SF') : language;   // "English SF" for the SF grid
  let sheet = controlSS.getSheetByName(tabName);
  if (!sheet) sheet = controlSS.insertSheet(tabName);

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearFormats();

  if (shifts.length === 0) {
    // Cleared on purpose: no tours in the window -> nothing stale survives.
    sheet.getRange(1, 1).setValue(`${label}: no tours in the scheduling window`)
      .setFontStyle('italic').setFontColor('#94a3b8');
    return;
  }

  /* ---------- columns ----------
   * Regular tours: one column per start time ("11:00").
   * Private tours: one EXTRA column per private group ("10:00 · Private",
   * "10:00 · Private 2", ...) — a private tour is its own shift, never mixed
   * into a regular column. Columns are sorted by time, regular first.
   */
  const colMap = new Map();   // colId -> {time, isPrivate, index, minutes, isSf}
  const rank = c => c.isSf ? 2 : (c.isPrivate ? 1 : 0);   // regular, then private, then SF
  shifts.forEach(s => {
    const t = normalizeTime_(s.time);
    const isSf = !!s.sfExt;
    const isPriv = !isSf && !!s.isPrivate;
    const idx = isPriv ? (Number(s.privIndex) || 1) : 1;
    const colId = gridColId_(t, isPriv, idx, isSf);
    if (!colMap.has(colId)) {
      colMap.set(colId, { time: t, isPrivate: isPriv, index: idx, isSf: isSf, minutes: timeToMinutes_(t) });
    }
  });
  const cols = [...colMap.entries()]
    .sort((a, b) => a[1].minutes - b[1].minutes ||
                    rank(a[1]) - rank(b[1]) ||
                    a[1].index - b[1].index);

  const dates = [...new Set(shifts.map(s => s.dateText))].sort();   // yyyy-MM-dd
  const title = `${label} schedule (${dates[0]} to ${dates[dates.length - 1]})`;

  // dateText -> colId -> {text, boldNames[], conflict}
  const cells = {};
  shifts.forEach(shift => {
    const t = normalizeTime_(shift.time);
    const colId = gridColId_(t, !shift.sfExt && shift.isPrivate, Number(shift.privIndex) || 1, !!shift.sfExt);
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

  const headerRow = ['Date'].concat(cols.map(c => gridHeaderForColumn_(c[1].time, c[1].isPrivate, c[1].index, c[1].isSf)));
  sheet.getRange(2, 1, 1, totalCols).setValues([headerRow])
    .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#bfdbfe');
  // Tint private (amber) and SF (indigo) column headers so they read as their own shifts.
  cols.forEach((c, i) => {
    if (c[1].isPrivate) sheet.getRange(2, i + 2).setBackground('#fde68a');
    else if (c[1].isSf) sheet.getRange(2, i + 2).setBackground('#c7d2fe');
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


/* ---------- Readers ---------- */

// Read the Guides tab into guide objects (name, active, seniority, languages, vacations, tour
// types). Columns are resolved BY HEADER NAME, so adding/reordering a column needs no change.
function readGuides_(ss) {
  const sheet = ss.getSheetByName("Guides");
  if (!sheet) throw new Error("Guides tab not found.");
  const values = sheet.getDataRange().getValues();
  // Language columns are resolved BY HEADER NAME, not fixed positions, so adding
  // a language column to the Guides tab needs no code change and never shifts
  // the others. Mirrors guidePortal.gs (guideColumns_).
  const header = (values[0] || []).map(h => String(h).trim());
  const langCol = {};
  LANGUAGES.forEach(l => {
    const c = header.findIndex(h => h.toLowerCase() === l.toLowerCase());
    if (c !== -1) langCol[l] = c;
  });
  const vacCol = header.findIndex(h => /vacation/i.test(h));   // "Vacation dates" column
  const tourCol = header.findIndex(h => /^tour$/i.test(h));    // "Tour" column: 3h / SF / both
  const guides = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = row[0];
    if (!name) continue;
    const languages = {};
    LANGUAGES.forEach(l => { languages[l] = langCol[l] !== undefined && row[langCol[l]] === true; });
    guides.push({
      name: String(name).trim(),
      active: row[1] === true,
      seniority: Number(row[2]) || 999,
      order: r,
      vacations: parseVacationRanges_(vacCol !== -1 ? row[vacCol] : ''),   // blocked date ranges
      // Which tour type(s) the guide runs (shared parser with the portal). Blank
      // defaults to 3h, so the auto-scheduler never puts an SF-only guide on a 3h
      // tour, nor a 3h-only guide on SF.
      tourTypes: parseTourTypes_(tourCol !== -1 ? row[tourCol] : ''),
      languages
    });
  }
  return guides;
}

// Just the names of the active guides.
function readActiveGuideNames_(ss) {
  return readGuides_(ss).filter(g => g.active).map(g => g.name);
}

// Resolve the Weekly_Schedule columns BY HEADER NAME (falling back to the legacy fixed
// positions, accounting for whether the "Tour Type" column was inserted), so a manager can
// add/reorder columns without silently corrupting the reader.
function weeklyScheduleCols_(header) {
  const find = (re, fb) => {
    for (let i = 0; i < header.length; i++) {
      if (re.test(String(header[i] || '').toLowerCase().trim())) return i;
    }
    return fb;
  };
  const tourType = find(/tour\s*type/, -1);
  const s = tourType > -1 ? 1 : 0;   // legacy layout had no Tour Type column
  return {
    day: find(/^day$/, 0), time: find(/^time$/, 1), language: find(/^language$/, 2),
    tourType: tourType,
    needed: find(/guides?\s*needed/, 3 + s),
    from: find(/active\s*from/, 4 + s),
    until: find(/active\s*unt/, 5 + s),
    guide: find(/^guide$/, 6 + s),
    hideAvail: find(/hide.*avail/, 7 + s),
    private: find(/^private$/, 8 + s),
    hideWeb: find(/hide.*(web|site)/, 9 + s)
  };
}

// Read the Weekly_Schedule tab into offer/availability rules (day, time, language, tour type,
// guides needed, active window, recurring guide, and the hide/private flags).
function readWeeklySchedule_(ss) {
  const sheet = ss.getSheetByName("Weekly_Schedule");
  if (!sheet) throw new Error("Weekly_Schedule tab not found.");
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();
  const rules = [];
  // Resolve columns BY HEADER NAME (like the Guides tab) so INSERTING a column —
  // e.g. the new "Tour Type" — never shifts the others out from under the reader.
  const C = weeklyScheduleCols_(displayValues[0] || []);
  for (let r = 1; r < values.length; r++) {
    const rawRow = values[r], displayRow = displayValues[r];
    const day = String(displayRow[C.day] || "").trim();
    // Sheets sometimes coerces a "10:00" cell into a Date (epoch 1899, shown
    // as 12/30/1899). Fall back to the raw Date's hours/minutes so the rule survives.
    const time = normalizeTime_(displayRow[C.time]) || timeFromCellValue_(rawRow[C.time]);
    const language = String(displayRow[C.language] || "").trim();
    // Blank "Guides needed" defaults to 1; an explicit 0 is kept (Private
    // availability slots stage no tour), so never collapse 0 to 1 here.
    const gnStr = String(displayRow[C.needed] || "").trim();
    const guidesNeeded = gnStr === "" ? 1 : (Number.isFinite(Number(gnStr)) ? Number(gnStr) : 1);
    // "Private": a yes/x/TRUE flag, OR the legacy literal "Private" in the Language cell.
    const isPrivate = /^(1|true|yes|y|x)$/i.test(String(C.private > -1 ? displayRow[C.private] : "").trim())
                      || /^private$/i.test(language);
    // "Tour Type": "3h" (the full tour) or "SF" (Sagrada Família exterior add-on) —
    // an SF row builds its OWN tour (card/column/guide), never merged with the 3h.
    const sfExt = C.tourType > -1 && /^\s*sf\b/i.test(String(displayRow[C.tourType] || "").trim());
    if (!day || !time || (!language && !isPrivate)) continue;
    rules.push({
      day, time, language, guidesNeeded, isPrivate, sfExt,
      tourName: sfExt ? 'SF' : '3h',
      activeFrom: (C.from > -1 && rawRow[C.from]) ? dateOnly_(new Date(rawRow[C.from])) : null,
      activeUntil: (C.until > -1 && rawRow[C.until]) ? dateOnly_(new Date(rawRow[C.until])) : null,
      // Optional recurring default guide: who runs this weekly slot unless a manager
      // overrides that date. Read by the portal.
      guide: String(C.guide > -1 ? displayRow[C.guide] : "").trim(),
      // "Hide from availability": hides this slot's column from the guide-availability
      // sheet ONLY. The tour still runs on the website and portal.
      hideFromAvailability: /^(1|true|yes|y|x|hide)$/i.test(String(C.hideAvail > -1 ? displayRow[C.hideAvail] : "").trim()),
      // "Hide from website": hides this slot from the PUBLIC website's bookable
      // availability ONLY. Independent of the other two flags.
      hideFromWebsite: /^(1|true|yes|y|x|hide)$/i.test(String(C.hideWeb > -1 ? displayRow[C.hideWeb] : "").trim())
    });
  }
  return rules;
}

// Read the ticked availability on a week tab into { week, guideName, date, time } rows — one
// per guide × slot they marked.
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

// Read the unique (date, day, time) slots a week's availability tab covers (its columns),
// regardless of who ticked them.
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

// Build the tour shifts for a week. DESIGN RULE: tours come from Weekly_Schedule RULES ×
// DATES, never from the availability sheet's column layout — so a tour ALWAYS exists for each
// rule, and if nobody ticked availability for it, it simply comes out "Not assigned".
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
        sfExt: !!rule.sfExt,                                 // Tour Type = SF -> its own tour
        weeklyGuide: String(rule.guide || '').trim(),        // Weekly_Schedule col G = the recurring guide for this slot
        availableGuides: availabilityMap[`${dateText}|${info.day}|${rule.time}`] || []
      });
    });
  });
  return shifts;
}


// Count how many ACTIVE guides speak each language.
function countLanguageSpeakers_(guides) {
  const counts = {};
  LANGUAGES.forEach(language => counts[language] = 0);
  guides.forEach(guide => {
    if (!guide.active) return;
    LANGUAGES.forEach(language => { if (guide.languages[language] === true) counts[language]++; });
  });
  return counts;
}

// Work out the full Monday–Sunday dates a week tab covers, so a rebuild always spans all 7
// days (Sunday appears as soon as a Sunday slot exists).
function readWeekDates_(sheet) {
  const displayValues = sheet.getDataRange().getDisplayValues();
  const title = String(displayValues[0][0] || "");
  const year = extractYear_(title);
  const dayRow = displayValues[2] || [];
  // Find the earliest dated column, then anchor to its Monday and span the whole week.
  let earliest = null;
  for (let c = 1; c < dayRow.length; c++) {
    const label = dayRow[c];
    if (!label) continue;
    const d = parseDateFromHeader_(label, year);
    if (d && !isNaN(d) && (!earliest || d < earliest)) earliest = d;
  }
  // RECOVERY: a blanked tab has no day headers, so fall back to the date in its
  // TITLE — "Availability: Week 42 (Mon Oct 12 - Sun Oct 18, 2026)". This lets a
  // once-wiped sheet be rebuilt on the next sync instead of being skipped forever.
  if (!earliest) {
    const m = title.match(/\(\s*(?:[A-Za-z]{3,}\s+)?([A-Za-z]{3,})\s+(\d{1,2})\b/);
    if (m) {
      const mi = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[1].slice(0, 3).toLowerCase());
      if (mi >= 0) { const d = new Date(year, mi, Number(m[2]), 12, 0, 0, 0); if (!isNaN(d)) earliest = d; }
    }
  }
  if (!earliest) return [];
  const monday = new Date(earliest);
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));   // snap back to Monday
  monday.setHours(12, 0, 0, 0);
  const dates = [];
  for (let d = 0; d < 7; d++) {
    const dt = new Date(monday); dt.setDate(monday.getDate() + d); dt.setHours(12, 0, 0, 0);
    dates.push({ dateObj: dt, dateKey: formatDate_(dt), dayName: fullDayName_(dt), shortLabel: shortDateLabel_(dt) });
  }
  return dates;
}

// Read which (guide, date, time) ticks already exist on a week tab, so a rebuild can carry
// them over instead of wiping what guides already entered.
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

// Rebuild one availability week tab: a column per (date, time) offered that week, a row per
// guide, carrying over every existing tick.
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

  // SAFETY: never clear + rebuild a sheet when there is nothing to write. If the
  // schedule rules or the guide list came back empty — e.g. a transient
  // "Service Spreadsheets timed out" mid-sync — bail BEFORE touching the sheet,
  // so a good availability tab (names + ticks) is never wiped to a blank shell.
  if (!slots.length || !guideNames || !guideNames.length) return;

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.clear();

  const title = `Availability: ${sheet.getName()} (${weekDates[0].shortLabel} - ${weekDates[weekDates.length - 1].shortLabel}, ${weekDates[0].dateObj.getFullYear()})`;

  // Banner across the whole row, but do NOT merge the title across all columns.
  // A merged cell spanning the full 7-day width makes Google Sheets REPEAT the
  // text when you scroll right (it looked like a corrupted title). Left-aligned
  // text in an unmerged A1/A2 overflows across the coloured banner and reads
  // cleanly with no repeat.
  sheet.getRange(1, 1, 1, totalCols).setBackground("#2563eb");
  sheet.getRange(1, 1).setValue(title)
    .setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");

  sheet.getRange(2, 1, 1, totalCols).setBackground("#dbeafe");
  sheet.getRange(2, 1)
    .setValue("Select your name in column A, then tick the shifts you can work. This is availability only.")
    .setFontStyle("italic").setHorizontalAlignment("left");

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
    const timeRange = sheet.getRange(4, 2, 1, times.length);
    // TEXT format FIRST: otherwise Sheets coerces "10:00" into a Date and the
    // header renders as 12/30/1899 (and readAvailability_ can no longer parse
    // it). Newly created week tabs start on General format, which is why the
    // later weeks were the ones showing 1899.
    timeRange.setNumberFormat('@');
    timeRange.setValues([times])
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

// Style the flat "Schedule" tab (header row, borders, muted "Not assigned" rows).
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


/* ---------- Small helpers ---------- */

// True if candidateDate is within TOUR_DURATION_HOURS of any already-assigned start time.
function hasConflict_(assignedDates, candidateDate) {
  return assignedDates.some(existing => Math.abs(candidateDate - existing) / (1000 * 60 * 60) < TOUR_DURATION_HOURS);
}

// "10:00" from a Date-coerced time cell (Sheets stores a bare time as an epoch-1899 Date), else ''.
function timeFromCellValue_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return v.getHours() + ':' + String(v.getMinutes()).padStart(2, '0');
  }
  return '';
}

// Normalise a time value ("5:00 PM", "17:00", a Date) to 24h "H:MM".
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

// A Date at the given day + "H:MM" time.
function combineDateAndTime_(dateObj, timeText) {
  const d = new Date(dateObj);
  const parts = String(timeText).split(":");
  d.setHours(Number(parts[0]), Number(parts[1] || 0), 0, 0);
  return d;
}

// The same date pinned to noon, so date-only comparisons never trip over a DST/UTC edge.
function dateOnly_(dateObj) {
  const d = new Date(dateObj);
  d.setHours(12, 0, 0, 0);
  return d;
}

// The 4-digit year found in a tab title, else the current year.
function extractYear_(title) {
  const match = String(title).match(/\b(20\d{2})\b/);
  return match ? Number(match[1]) : new Date().getFullYear();
}

// Parse a day-header label like "Mon Oct 12" (+ the year) into a noon Date.
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

// Validate the CURRENT Schedule_<Language> grids against the Guides tab and the overlap rule:
// a guide who doesn't speak the language, doesn't run the tour type, isn't in the Guides tab,
// or has two tours within MIN_SEPARATION_HOURS. Logs every violation and returns the count.
// Run from the editor after makeSchedule, or any time a manager hand-edits a grid.
function validateScheduleGrids() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  const guides = readGuides_(controlSS);
  const byName = {};
  guides.forEach(g => { byName[g.name.toLowerCase()] = g; });

  const seen = {};   // guide -> [Date] tour starts
  const problems = [];

  controlSS.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const parsed = parseScheduleTabName_(name);
    if (!parsed) return;
    const language = parsed.language, tabIsSf = parsed.isSf;
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
              const gt = g.tourTypes || { threeH: true, sf: false };
              if (tabIsSf ? !gt.sf : !gt.threeH) {
                problems.push(`${nm} @ ${formatDate_(dObj)} ${time} ${language}: does not run the ${tabIsSf ? 'SF' : '3h'} tour`);
              }
              if (!seen[nm]) seen[nm] = [];
              seen[nm].forEach(prev => {
                // SF tours are ~1h and stackable: two SF slots never clash. A clash
                // between SF and a long 3h tour still counts.
                if (tabIsSf && prev.sf) return;
                if (Math.abs(start - prev.d) / 3600000 < ASSIGN_CFG.MIN_SEPARATION_HOURS) {
                  problems.push(`${nm}: overlapping tours ${prev.d.toISOString().slice(0,16)} and ${start.toISOString().slice(0,16)} (<${ASSIGN_CFG.MIN_SEPARATION_HOURS}h apart)`);
                }
              });
              seen[nm].push({ d: start, sf: tabIsSf });
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

// Run once: install the on-edit trigger that watches the Schedule_<Language> grids. From then
// on, when a manager types a guide name into a grid cell:
//   1. the cell is made BOLD (= a management lock, so the weekly regeneration never wipes a
//      hand edit), and
//   2. the assignment is validated immediately — wrong language, unknown name, or an
//      overlapping tour (<5h) puts a ⚠ note on the cell and a row in the Errors tab. The edit
//      is KEPT (management decides); the flag just makes the problem impossible to miss.
function setupScheduleEditValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'handleScheduleEdit');
  if (!exists) {
    ScriptApp.newTrigger('handleScheduleEdit').forSpreadsheet(ss).onEdit().create();
  }
  Logger.log('Schedule edit validation installed.');
}

// Installable on-edit handler for the Schedule_<Language> grids: bold a typed name (lock it)
// and validate the cell, flagging any problem with a ⚠ note + an Errors row.
function handleScheduleEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const name = sh.getName();
    const parsed = parseScheduleTabName_(name);
    if (!parsed) return;
    const language = parsed.language, tabIsSf = parsed.isSf;
    if (e.range.getRow() < 3 || e.range.getColumn() < 2) return;
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

    const val = String(e.range.getValue() || '').trim();
    if (!val || /not assigned/i.test(val)) { e.range.setNote(''); return; }

    // A typed name IS a management decision -> lock it (bold, like the
    // portal's assign action), so makeSchedule preserves it.
    e.range.setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText(val)
        .setTextStyle(0, val.length, SpreadsheetApp.newTextStyle().setBold(true).build())
        .build());
    e.range.setFontStyle('normal').setFontColor('#1a2b49');

    const problems = validateCellAssignment_(sh, e.range, language, val, tabIsSf);
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

// The problems (as strings) with putting `namesText` into this grid cell.
function validateCellAssignment_(sh, range, language, namesText, isSf) {
  const problems = [];
  const dv = sh.getDataRange().getDisplayValues();
  const anchor = gridAnchor_(String((dv[0] && dv[0][0]) || ''));
  const header = parseGridTimeHeader_(String((dv[1] || [])[range.getColumn() - 1] || ''));
  const dateKey = gridLabelToKey_(String((dv[range.getRow() - 1] || [])[0] || '').trim(), anchor);
  if (!header || !dateKey) return problems;   // not a data cell we understand

  const sf = !!isSf || !!header.isSf;          // an SF grid tab (or legacy SF column)
  const startMs = shiftStartMs_(dateKey, timeToMinutes_(header.time));
  const sepMs = ASSIGN_CFG.MIN_SEPARATION_HOURS * 3600000;
  const myKey = dateKey + '|' + timeToMinutes_(header.time) + '|' + language.toLowerCase() +
                '|' + (sf ? 'SF' : (header.isPrivate ? 'P' + header.index : 'R'));

  let busy = {};
  try { busy = buildBusyMap_(readSchedule_()); } catch (err) { /* portal reader unavailable */ }

  namesText.split(/[\n,]/).map(s => s.trim()).filter(Boolean).forEach(nm => {
    if (/not assigned|need \d|lock conflict/i.test(nm)) return;
    const g = findGuideByName_(nm);
    if (!g) { problems.push(nm + ': not in the Guides tab'); return; }
    if (!g.active) problems.push(nm + ' is marked inactive');
    if (g.languages[language] !== true) problems.push(nm + ' does not speak ' + language);
    const gt = g.tourTypes || { threeH: true, sf: false };
    if (sf ? !gt.sf : !gt.threeH) problems.push(nm + ' does not run the ' + (sf ? 'SF' : '3h') + ' tour');
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


