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

  // Value of a shift to the guide who runs it (drives balancing).
  VALUE_PAID_PER_GUEST: 10,     // paid tour: guide earns ~10 €/guest
  VALUE_FREE_NET_PER_GUEST: 14, // free tour: ~20 € tip − 6 € commission per guest
  VALUE_PRIVATE_FLAT: 75,       // private tour: flat 75 € to the guide

  PAID_SOURCES: ["Viator", "GetYourGuide", "Airbnb"],

  // Where schedule problems (lock conflicts, missing eligibility) are logged.
  ERRORS_TAB: "Errors"
};

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
    const lk = lockKey_(shift.dateText, normalizeTime_(shift.time), shift.language, shift.isPrivate);
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
      // Keep the lock regardless; flag any problem.
      shift.lockedGuides.push(nm);
      if (g) {
        accVal[g.name] += shift.value;
        tourCount[g.name]++;
        assignedByGuide[g.name].push(shift.dateTimeObj);
      }
      if (problems.length) {
        conflicts.push({
          dateText: shift.dateText, time: shift.time, language: shift.language,
          isPrivate: !!shift.isPrivate, guide: nm, problems
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
      c.dateText === shift.dateText && normalizeTime_(c.time) === normalizeTime_(shift.time) &&
      c.language === shift.language && !!c.isPrivate === !!shift.isPrivate);

    assignedShifts.push({
      ...shift,
      eligibleGuides: eligible.map(g => g.name),
      assignedGuides: shift.lockedGuides.concat(assigned.map(g => g.name)),
      lockedGuides: shift.lockedGuides,
      hasLockConflict: hasConflictFlag,
      status: ok ? "OK" : "Not assigned",
      notes: [
        shift.isPrivate ? "Private" : "",
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
  return { shifts: assignedShifts.length, conflicts: conflicts.length };
}


/* ---------- manager locks (bold names in the grids) ---------- */

function lockKey_(dateText, time, language, isPrivate) {
  return dateText + "|" + time + "|" + language + "|" + (isPrivate ? "P" : "R");
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
        const time = normalizeTime_(timeRow[c]);
        if (!time) continue;
        const rt = rich[r] && rich[r][c];
        if (!rt) continue;
        const runs = rt.getRuns ? rt.getRuns() : [];
        runs.forEach(run => {
          const style = run.getTextStyle();
          if (!style || !style.isBold()) return;
          const txt = String(run.getText() || "");
          const isPriv = /🔒|\(private\)/i.test(txt);
          txt.split(/[\n,]/).forEach(piece => {
            let nm = piece.replace(/🔒/g, "").replace(/\(private\)/ig, "").trim();
            if (!nm || /not assigned|lock conflict/i.test(nm)) return;
            const k = lockKey_(dateText, time, language, isPriv);
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
    ts, "Schedule lock conflict",
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
 * Keep the regular shifts, then add one private shift per private booking group
 * AT ITS OWN booked time, staffed from the availability index (not by cloning a
 * regular slot). This is what lets a 10:00 private tour be scheduled even though
 * no regular tour runs at 10:00. availableGuides come from whoever ticked
 * availability at that date+time.
 */
function expandPrivateShifts_(shifts, availIndex, startDate, endDate) {
  const priv = readPrivateCounts_();   // "dateText|time|language" -> count
  const out = shifts.map(s => Object.assign({}, s, { isPrivate: false }));

  Object.keys(priv).forEach(key => {
    const parts = key.split('|');
    const dateText = parts[0], time = parts[1], language = parts[2];
    const dateObj = new Date(dateText + 'T12:00:00');
    const d = dateOnly_(dateObj);
    if (d < startDate || d > endDate) return;                 // outside the window
    const dateTimeObj = combineDateAndTime_(dateObj, time);
    const availableGuides = availIndex[dateText + '|' + time] || [];
    const n = priv[key];
    for (let i = 0; i < n; i++) {
      out.push({
        week: '', dateObj, dateTimeObj, dateText, day: fullDayName_(dateObj),
        time, language, guidesNeeded: 1, availableGuides, isPrivate: true
      });
    }
  });

  return out;
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
    const language = tab.replace(" Tours", "").trim();
    if (sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    rows.forEach(row => {
      if (/privat/i.test(String(row[8] || ""))) return;   // private valued as its own shift
      const dateKey = row[3] instanceof Date ? formatDate_(row[3]) : formatDate_(new Date(row[3]));
      const time = normalizeTime_(row[4]);
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

function readPrivateCounts_() {
  const out = {};
  const ss = SpreadsheetApp.openById(BOOKING_SHEET_ID);
  ss.getSheets().forEach(sh => {
    const tab = sh.getName();
    if (tab.indexOf(" Tours") === -1) return;
    const language = tab.replace(" Tours", "").trim();
    if (sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    rows.forEach(row => {
      if (!/privat/i.test(String(row[8] || ""))) return;
      const dateKey = row[3] instanceof Date ? formatDate_(row[3]) : formatDate_(new Date(row[3]));
      const time = normalizeTime_(row[4]);
      const key = dateKey + "|" + time + "|" + language;
      out[key] = (out[key] || 0) + 1;
    });
  });
  return out;
}


/* ---------- unchanged workflow entry points ---------- */

function syncAvailabilityFile() {
  const controlSS = SpreadsheetApp.getActiveSpreadsheet();
  const guideSS = SpreadsheetApp.openById(GUIDE_FILE_ID);

  const guideNames = readActiveGuideNames_(controlSS);
  const scheduleRules = readWeeklySchedule_(controlSS);

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
  const today = formatDate_(dateOnly_(new Date()));

  const rows = [['Day', 'Time', 'Language', 'Guides needed', 'Active from', 'Active until']];
  const add = (days, time, lang) => days.forEach(d => rows.push([d, time, lang, 1, today, '']));

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
  const scheduleRules = readWeeklySchedule_(controlSS);

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
  const activeLanguages = [...new Set(assignedShifts.map(s => s.language))].sort();
  activeLanguages.forEach(language => {
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

  if (shifts.length === 0) return;

  const times = [...new Set(shifts.map(s => s.time))].sort((a, b) => timeToMinutes_(a) - timeToMinutes_(b));

  // Rows are ACTUAL DATES (not weekdays), so a multi-week span shows correctly
  // and never overwrites one week's assignment with another's.
  const dates = [...new Set(shifts.map(s => s.dateText))].sort();   // yyyy-MM-dd
  const title = `${language} schedule (${dates[0]} to ${dates[dates.length - 1]})`;

  const byDateTime = {};      // dateText -> time -> cell text
  const boldNames = {};       // dateText|time -> [names to render bold (locks)]
  const conflictCells = {};   // dateText|time -> true (tint red)
  shifts.forEach(shift => {
    if (!byDateTime[shift.dateText]) byDateTime[shift.dateText] = {};
    const assigned = shift.assignedGuides.join(", ");
    let text = assigned;
    if (shift.status !== "OK") text = assigned ? `${assigned}\n${shift.status}` : shift.status;
    if (shift.isPrivate) text = "🔒 " + (text || shift.status) + " (private)";
    // Regular + private share the same date/time cell -> stack them, don't overwrite.
    const prev = byDateTime[shift.dateText][shift.time];
    byDateTime[shift.dateText][shift.time] = prev ? (prev + "\n" + text) : text;
    const bk = shift.dateText + "|" + shift.time;
    (shift.lockedGuides || []).forEach(n => {
      if (!boldNames[bk]) boldNames[bk] = [];
      if (boldNames[bk].indexOf(n) === -1) boldNames[bk].push(n);
    });
    if (shift.hasLockConflict) conflictCells[bk] = true;
  });

  const totalCols = 1 + times.length;

  sheet.getRange(1, 1, 1, totalCols).merge().setValue(title)
    .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center")
    .setBackground("#2563eb").setFontColor("#ffffff");

  sheet.getRange(2, 1, 1, totalCols).setValues([["Date", ...times]])
    .setFontWeight("bold").setHorizontalAlignment("center").setBackground("#bfdbfe");

  const table = dates.map(dt => {
    const label = Utilities.formatDate(new Date(dt + "T12:00:00"), Session.getScriptTimeZone(), "EEE MMM d");
    return [label, ...times.map(time => byDateTime[dt]?.[time] || "")];
  });

  const tableRange = sheet.getRange(3, 1, table.length, totalCols);
  tableRange.setValues(table).setBorder(true, true, true, true, true, true)
    .setVerticalAlignment("middle").setWrap(true);

  sheet.getRange(3, 1, table.length, 1).setFontWeight("bold").setBackground("#dbeafe");
  sheet.getRange(3, 2, table.length, totalCols - 1).setHorizontalAlignment("center").setBackground("#f8fbff");

  for (let r = 0; r < table.length; r++) {
    const dt = dates[r];
    for (let c = 1; c < table[r].length; c++) {
      const cellText = String(table[r][c]);
      const cell = sheet.getRange(3 + r, c + 1);
      if (cellText.includes("Not assigned")) {
        cell.setFontColor("#94a3b8").setFontStyle("italic"); // muted, no red
      }
      const bk = dt + "|" + times[c - 1];
      // Re-apply BOLD to manager-locked names so the lock survives the rebuild
      // and stays visible (and re-readable) as a lock next run.
      const locked = boldNames[bk] || [];
      if (cellText && locked.length) {
        let builder = SpreadsheetApp.newRichTextValue().setText(cellText);
        locked.forEach(nm => {
          let idx = cellText.toLowerCase().indexOf(nm.toLowerCase());
          while (idx !== -1) {
            builder = builder.setTextStyle(idx, idx + nm.length,
              SpreadsheetApp.newTextStyle().setBold(true).build());
            idx = cellText.toLowerCase().indexOf(nm.toLowerCase(), idx + nm.length);
          }
        });
        cell.setRichTextValue(builder.build());
      }
      if (conflictCells[bk]) cell.setBackground("#fde2e2");   // visible conflict flag
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
    const time = normalizeTime_(displayRow[1]);
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

function buildShifts_(availability, weeklySchedule, guideCalendar) {
  const availabilityMap = {};
  availability.forEach(a => {
    const key = `${a.dateText}|${a.day}|${a.time}`;
    if (!availabilityMap[key]) availabilityMap[key] = [];
    if (!availabilityMap[key].includes(a.guideName)) availabilityMap[key].push(a.guideName);
  });
  const shifts = [];
  guideCalendar.forEach(slot => {
    weeklySchedule.forEach(rule => {
      if (rule.day !== slot.day) return;
      if (rule.time !== slot.time) return;
      const slotDateOnly = dateOnly_(slot.dateObj);
      if (rule.activeFrom && slotDateOnly < rule.activeFrom) return;
      if (rule.activeUntil && slotDateOnly > rule.activeUntil) return;
      const key = `${slot.dateText}|${slot.day}|${slot.time}`;
      shifts.push({
        week: slot.week, dateObj: slot.dateObj, dateTimeObj: slot.dateTimeObj,
        dateText: slot.dateText, day: slot.day, time: slot.time,
        language: rule.language, guidesNeeded: rule.guidesNeeded,
        availableGuides: availabilityMap[key] || []
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
        const time = normalizeTime_(timeRow[c]);
        if (!time) continue;
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
