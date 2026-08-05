/* ===== Guide-portal END-TO-END =====
   Drives the REAL portal API — apiTours_ / apiAssign_ / apiSave_ — against the
   stateful Sheets mock, exactly as the phone does over JSONP. Proves the whole
   decision chain that kept breaking in production:
     a booking surfaces as a shift  ->  a manager assigns a guide  ->  the
     assignment STICKS on the next read with a SINGLE grid row (no revert, no
     duplicate)  ->  a check-in saves and PERSISTS with the right count + time
     ->  reassigning swaps the guide, still one row.
   Bundled with mock.js + control/*.gs. */
let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

const BOOK_ID = '1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw';
const dayKey = o => { const d = new Date(); d.setDate(d.getDate() + o); return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); };

// --- Control spreadsheet: guides (a manager + a guide) + an empty offer table.
const control = new __mock.MockSS('control'); SpreadsheetApp._active = control;
control.insertSheet('Guides').getRange(1, 1, 3, 11).setValues([
  ['Guide', 'Active?', 'Seniority', 'English', 'German', 'Spanish', 'French', 'Italian', 'Manager', 'Email', 'Password'],
  ['Albert', true, 2, true, false, false, false, false, true, 'a@x.com', 'pw'],
  ['Carlos', true, 1, true, false, false, false, false, false, 'c@x.com', 'pw']]);
control.insertSheet('Weekly_Schedule').getRange(1, 1, 1, 6).setValues([
  ['Day', 'Time', 'Language', 'Guides needed', 'Active from', 'Active until']]);

// --- BookingSheet (read by the portal via its id): one upcoming English booking.
const booking = new __mock.MockSS('booking'); __mock.SS_BY_ID[BOOK_ID] = booking;
const DATE = dayKey(0);                         // TODAY — the day tours run and check-ins happen
const en = booking.insertSheet('English Tours');
en.getRange(1, 1, 2, 9).setValues([
  ['Name', 'Phone', 'Number of Guests', 'Tour date', 'Time', 'Source', 'Income', 'Booking ID', 'Notes'],
  ['Dana Ortiz', '+34600111222', 2, new Date(DATE + 'T12:00:00'), '10:00 AM', 'GetYourGuide', 30, 'GYGE2E001', '']]);

const token = makeToken_('Albert');
const findShift = r => (r.allTours || []).filter(s => s.dateKey === DATE && s.time === '10:00' && s.language === 'English');

console.log('--- The booking surfaces as an (unassigned) shift with its reservation ---');
let r = apiTours_({ token: token });
check('apiTours_ returns ok for the manager', r && r.ok === true && r.manager === true, r && r.error);
let sh = findShift(r);
check('the booking appears as exactly one shift', sh.length === 1, sh.map(s => s.time));
check('the shift carries the reservation (Dana Ortiz, 2)', sh.length === 1 && sh[0].bookings.some(b => b.bookingId === 'GYGE2E001' && b.guests === 2), sh[0] && sh[0].bookings);
check('the shift starts UNASSIGNED', sh.length === 1 && (sh[0].assigned || []).length === 0, sh[0] && sh[0].assigned);

console.log('--- Manager assigns Carlos; it STICKS on reload, single row, no revert ---');
let a = apiAssign_({ token: token, dateKey: DATE, time: '10:00', language: 'English', guide: 'Carlos', force: '1' });
check('apiAssign_ returns Carlos', a && a.ok === true && a.assigned === 'Carlos', a);
r = apiTours_({ token: token }); sh = findShift(r);
check('after reload there is STILL exactly one shift (no duplicate)', sh.length === 1, sh.length);
check('the shift is assigned to Carlos', sh.length === 1 && (sh[0].assigned || []).map(x => x.toLowerCase()).indexOf('carlos') !== -1, sh[0] && sh[0].assigned);
// The grid itself has a single dated row for this slot.
const gv = control.getSheetByName('Schedule_English').getDataRange().getDisplayValues();
const label = Utilities.formatDate(new Date(DATE + 'T12:00:00'), Session.getScriptTimeZone(), 'EEE MMM d');
check('the grid has a SINGLE row for the date', gv.filter((row, i) => i >= 2 && row[0] === label).length === 1, gv.map(x => x[0]));

console.log('--- A check-in saves and PERSISTS with the right count and a time ---');
const data = {
  dateKey: DATE, time: '10:00', timeLabel: '10:00 AM', day: dayNameFromKey_(DATE),
  language: 'English', guide: 'Carlos', walkins: [],
  bookings: [{ bookingId: 'GYGE2E001', source: 'GetYourGuide', name: 'Dana Ortiz', phone: '+34600111222',
               guests: 2, income: 30, isPrivate: false, manualNote: '', checked: true, checkedIn: 2 }]
};
const s1 = apiSave_({ token: token, data: JSON.stringify(data) });
check('apiSave_ (check-in) returns ok', s1 && s1.ok === true, s1);
r = apiTours_({ token: token }); sh = findShift(r);
const bk = sh.length ? sh[0].bookings.find(b => b.bookingId === 'GYGE2E001') : null;
check('the check-in persisted (checked, 2 in)', bk && bk.checked === true && bk.checkedIn === 2, bk);
check('the check-in TIME is present (HH:mm)', bk && /^\d{1,2}:\d{2}$/.test(bk.checkedAt || ''), bk && bk.checkedAt);
check('the shift total reflects 2 checked in', sh.length === 1 && sh[0].checkedGuests === 2, sh[0] && sh[0].checkedGuests);

console.log('--- Reassigning to Albert swaps the guide, still a single row ---');
apiAssign_({ token: token, dateKey: DATE, time: '10:00', language: 'English', guide: 'Albert', force: '1' });
r = apiTours_({ token: token }); sh = findShift(r);
check('reassigned to Albert', sh.length === 1 && (sh[0].assigned || []).map(x => x.toLowerCase()).indexOf('albert') !== -1, sh[0] && sh[0].assigned);
check('Carlos is gone (not merged in)', sh.length === 1 && (sh[0].assigned || []).map(x => x.toLowerCase()).indexOf('carlos') === -1, sh[0] && sh[0].assigned);
const gv2 = control.getSheetByName('Schedule_English').getDataRange().getDisplayValues();
check('still a single grid row after reassign', gv2.filter((row, i) => i >= 2 && row[0] === label).length === 1, gv2.map(x => x[0]));

console.log('--- Check-in reads only the assigned guide tab, and records children ---');
// A second English booking at 12:00, assigned to Carlos, checked in as Carlos.
// The manager view must reflect it (reading only the ASSIGNED guide tab, which
// is where the check-in lives — not sweeping every guide's tab).
en.getRange(3, 1, 1, 9).setValues([
  ['Ivan Petrov', '+34600333444', 3, new Date(DATE + 'T12:00:00'), '12:00 PM', 'GetYourGuide', 45, 'GYGE2E002', '']]);
apiAssign_({ token: token, dateKey: DATE, time: '12:00', language: 'English', guide: 'Carlos', force: '1' });
const data2 = {
  dateKey: DATE, time: '12:00', timeLabel: '12:00 PM', day: dayNameFromKey_(DATE),
  language: 'English', guide: 'Carlos', walkins: [],
  bookings: [{ bookingId: 'GYGE2E002', source: 'GetYourGuide', name: 'Ivan Petrov', phone: '+34600333444',
               guests: 3, children: 2, income: 45, isPrivate: false, manualNote: '', checked: true, checkedIn: 3 }]
};
apiSave_({ token: token, data: JSON.stringify(data2) });
r = apiTours_({ token: token });
const sh12 = (r.allTours || []).filter(s => s.dateKey === DATE && s.time === '12:00' && s.language === 'English');
const bk12 = sh12.length ? sh12[0].bookings.find(b => b.bookingId === 'GYGE2E002') : null;
check('the 12:00 shift is assigned to Carlos', sh12.length === 1 && (sh12[0].assigned || []).map(x => x.toLowerCase()).indexOf('carlos') !== -1, sh12[0] && sh12[0].assigned);
check('manager sees the check-in (read from the assigned guide tab)', bk12 && bk12.checked === true && bk12.checkedIn === 3, bk12);
check('the check-in carries its time', bk12 && /^\d{1,2}:\d{2}$/.test(bk12.checkedAt || ''), bk12 && bk12.checkedAt);
// The ledger row records the children count sent by the client.
const carlosTab = ledgerSS_().getSheetByName('Carlos');
const carlosRows = carlosTab ? carlosTab.getRange(2, 1, Math.max(0, carlosTab.getLastRow() - 1), LEDGER_HEADERS.length).getValues() : [];
const ivanRow = carlosRows.find(x => String(x[LEDGER_BOOKINGID_COL] || '') === 'GYGE2E002');
check('the ledger records the children count from the save payload', ivanRow && Number(ivanRow[8]) === 2, ivanRow && ivanRow[8]);

console.log('--- Per-phase timings + seniority-ordered eligible list ---');
const rt = apiTours_({ token: token });
check('response reports per-phase timings (sched/book/ledger)', rt.timings && typeof rt.timings.sched === 'number' && typeof rt.timings.book === 'number' && typeof rt.timings.ledger === 'number', rt.timings);
// Carlos is seniority 1, Albert seniority 2 -> Carlos must sort first, even though
// "Albert" is alphabetically earlier. Proves seniority beats alphabetical.
check('English eligible list is most-senior-first (Carlos before Albert)', (rt.guidesByLanguage.English || [])[0] === 'Carlos', rt.guidesByLanguage.English);

console.log('--- Manager window: near tours load first, far tours behind "Load more" ---');
const FAR = dayKey(40);
en.getRange(4, 1, 1, 9).setValues([
  ['Far Future', '+34600555000', 2, new Date(FAR + 'T12:00:00'), '10:00 AM', 'GetYourGuide', 30, 'GYGE2EFAR', '']]);
const rw = apiTours_({ token: token });                       // default 7-day window
check('response carries a freshness timestamp (HH:mm:ss)', /^\d{1,2}:\d{2}:\d{2}$/.test(rw.now || ''), rw.now);
check('default manager window is 5 days', rw.windowDays === 5, rw.windowDays);
check('a 40-day-out tour is NOT in the default window', !(rw.allTours || []).some(s => s.dateKey === FAR), FAR);
check('hasMore flags there are tours beyond the window', rw.hasMore === true, rw.hasMore);
const rw2 = apiTours_({ token: token, days: 45 });            // "Load more"
check('Load more (days=45) brings the far tour in', (rw2.allTours || []).some(s => s.dateKey === FAR), FAR);
check('with the full window there is nothing more to load', rw2.hasMore === false, rw2.hasMore);

console.log('--- Timing: each real portal operation is bounded, and repeats do not grow the grid ---');
const time = fn => { const t = Date.now(); fn(); return Date.now() - t; };
const p95 = a => { const s = a.slice().sort((x, y) => x - y); return s[Math.min(s.length - 1, Math.ceil(0.95 * s.length) - 1)]; };
const toursMs = [], assignMs = [], saveMs = [];
for (let i = 0; i < 6; i++) toursMs.push(time(() => apiTours_({ token: token })));
for (let i = 0; i < 6; i++) assignMs.push(time(() => apiAssign_({ token: token, dateKey: DATE, time: '10:00', language: 'English', guide: (i % 2 ? 'Carlos' : 'Albert'), force: '1' })));
for (let i = 0; i < 6; i++) saveMs.push(time(() => apiSave_({ token: token, data: JSON.stringify(data) })));
console.log('  apiTours_  p95=' + p95(toursMs) + 'ms max=' + Math.max.apply(null, toursMs) + 'ms');
console.log('  apiAssign_ p95=' + p95(assignMs) + 'ms max=' + Math.max.apply(null, assignMs) + 'ms');
console.log('  apiSave_   p95=' + p95(saveMs) + 'ms max=' + Math.max.apply(null, saveMs) + 'ms');
// In node these are ~instant; a hard cap catches a runaway loop / O(n^2) blow-up.
check('apiTours_ does bounded work (no runaway)', Math.max.apply(null, toursMs) < 2000, Math.max.apply(null, toursMs));
check('apiAssign_ does bounded work', Math.max.apply(null, assignMs) < 2000, Math.max.apply(null, assignMs));
check('apiSave_ does bounded work', Math.max.apply(null, saveMs) < 2000, Math.max.apply(null, saveMs));
// The real regression guard: 6 more assigns must NOT accumulate grid rows.
const gv3 = control.getSheetByName('Schedule_English').getDataRange().getDisplayValues();
check('after many repeated assigns the grid STILL has a single row', gv3.filter((row, i) => i >= 2 && row[0] === label).length === 1, gv3.map(x => x[0]));

console.log('--- Weekly recurring default guide fills an unassigned slot ---');
const wsheet = control.getSheetByName('Weekly_Schedule');
wsheet.getRange(1, 7).setValue('Guide');
wsheet.getRange(2, 1, 1, 7).setValues([[dayNameFromKey_(DATE), '17:00', 'English', 1, '', '', 'Carlos']]);
const rWeekly = apiTours_({ token: token, days: 45 });
const wshift = (rWeekly.allTours || []).filter(s => s.dateKey === DATE && s.time === '17:00' && s.language === 'English');
check('the weekly default (Carlos) fills the unassigned 17:00 English slot', wshift.length === 1 && (wshift[0].assigned || []).indexOf('Carlos') !== -1, wshift[0] && wshift[0].assigned);
// A manual assignment on that date must OVERRIDE the weekly default.
apiAssign_({ token: token, dateKey: DATE, time: '17:00', language: 'English', guide: 'Albert', force: '1' });
const rOverride = apiTours_({ token: token, days: 45 });
const oshift = (rOverride.allTours || []).filter(s => s.dateKey === DATE && s.time === '17:00' && s.language === 'English');
check('a manual assignment overrides the weekly default (Albert wins that day)', oshift.length === 1 && (oshift[0].assigned || []).indexOf('Albert') !== -1 && (oshift[0].assigned || []).indexOf('Carlos') === -1, oshift[0] && oshift[0].assigned);

console.log('--- Delete a tour: frees the guide (clears the grid cell) and prunes the empty column ---');
const DEL = dayKey(6);   // an isolated slot nothing else touches
apiAssign_({ token: token, dateKey: DEL, time: '19:00', language: 'English', guide: 'Carlos', force: '1' });
let gEng = control.getSheetByName('Schedule_English').getDataRange().getDisplayValues();
check('assigning created a 19:00 column', gEng[1].some(h => /(^|\D)19:00|7:00 PM/.test(String(h))), gEng[1]);
apiCloseShift_({ token: token, id: shiftKey_(DEL, 19 * 60, 'english') });
gEng = control.getSheetByName('Schedule_English').getDataRange().getDisplayValues();
check('after delete the empty 19:00 column is pruned', !gEng[1].some(h => /(^|\D)19:00|7:00 PM/.test(String(h))), gEng[1]);
const rDel = apiTours_({ token: token, days: 45 });
check('the deleted tour no longer appears (no bookings -> gone)', !(rDel.allTours || []).some(s => s.dateKey === DEL && s.time === '19:00' && s.language === 'English'), null);

console.log('--- Portal Feed: the portal reads reservations from one tab (with check-in cols for Phase 2) ---');
// (All earlier assertions ran with NO feed tab -> they exercised the fallback
//  to the per-language read, proving the fallback works.)
const feed = booking.insertSheet('Portal Feed');
feed.getRange(1, 1, 1, 14).setValues([['Date', 'Time', 'Language', 'Name', 'Phone', 'Adults', 'Children',
  'Source', 'Income', 'Booking ID', 'Notes', 'Manager note', 'Checked-in', 'Check-in time']]);
feed.getRange(2, 2, 1, 1).setNumberFormat('@'); feed.getRange(2, 10, 1, 1).setNumberFormat('@'); feed.getRange(2, 14, 1, 1).setNumberFormat('@');
feed.getRange(2, 1, 1, 14).setValues([[DATE, '10:00 AM', 'English', 'Dana Ortiz', '+34600111222', 2, 0,
  'GetYourGuide', 30, 'GYGE2E001', '', 'hello note', 2, '10:03']]);
const fx = readPortalFeed_();
const fk = shiftKey_(DATE, 600, 'english');
check('readPortalFeed_ indexes the booking by shift', !!(fx && fx[fk] && fx[fk][0].bookingId === 'GYGE2E001'), fx && Object.keys(fx));
check('feed carries the manager note', fx[fk][0].manualNote === 'hello note', fx[fk][0].manualNote);
check('feed carries the check-in for Phase 2 (feedCheckedIn=2, at 10:03)', fx[fk][0].feedCheckedIn === 2 && fx[fk][0].feedCheckedAt === '10:03', [fx[fk][0].feedCheckedIn, fx[fk][0].feedCheckedAt]);

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
