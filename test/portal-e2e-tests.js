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
  ['Albert', true, 1, true, false, false, false, false, true, 'a@x.com', 'pw'],
  ['Carlos', true, 1, true, false, false, false, false, false, 'c@x.com', 'pw']]);
control.insertSheet('Weekly_Schedule').getRange(1, 1, 1, 6).setValues([
  ['Day', 'Time', 'Language', 'Guides needed', 'Active from', 'Active until']]);

// --- BookingSheet (read by the portal via its id): one upcoming English booking.
const booking = new __mock.MockSS('booking'); __mock.SS_BY_ID[BOOK_ID] = booking;
const DATE = dayKey(5);                         // upcoming, inside the portal window
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

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
