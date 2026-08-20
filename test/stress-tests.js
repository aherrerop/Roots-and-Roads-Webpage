/* ===== STRESS TEST — try to BREAK the portal =====
   Adversarial sequences against the REAL functions: double-taps, stale /
   cross-device saves, undo->re-checkin, malformed feed rows, private/regular
   collisions, garbage inputs. Bundled with mock.js + control/*.gs. If any of
   these fails, we found a real bug. */
let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };
const dayKey = o => { const d = new Date(); d.setDate(d.getDate() + o); return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); };
const BOOK_ID = PORTAL.BOOKING_SHEET_ID;
const DATE = dayKey(0);

// --- Control: guides across languages (Albert manager) + empty weekly.
const control = new __mock.MockSS('control'); SpreadsheetApp._active = control;
control.insertSheet('Guides').getRange(1, 1, 5, 11).setValues([
  ['Guide', 'Active?', 'Seniority', 'English', 'German', 'Spanish', 'French', 'Italian', 'Manager', 'Email', 'Password'],
  ['Albert', true, 2, true, false, false, false, false, true, 'a@x.com', 'pw'],
  ['Carlos', true, 1, true, false, false, false, false, false, 'c@x.com', 'pw'],
  ['Francesca', true, 2, false, false, false, false, true, false, 'f@x.com', 'pw'],
  ['Inactive', false, 9, true, false, false, false, false, false, 'i@x.com', 'pw']]);
control.insertSheet('Weekly_Schedule').getRange(1, 1, 1, 7).setValues([
  ['Day', 'Time', 'Language', 'Guides needed', 'Active from', 'Active until', 'Guide']]);

// --- BookingSheet with a Portal Feed (16 cols) + English tab.
const booking = new __mock.MockSS('booking'); __mock.SS_BY_ID[BOOK_ID] = booking;
const FH = ['Date', 'Time', 'Language', 'Name', 'Phone', 'Adults', 'Children', 'Source', 'Income', 'Booking ID', 'Notes', 'Manager note', 'Checked-in', 'Check-in time', 'Guide', 'Type'];
const feed = booking.insertSheet('Portal Feed');
feed.getRange(1, 1, 1, 16).setValues([FH]);
feed.getRange(2, 1, 4, 16).setValues([
  [DATE, '10:00 AM', 'English', 'Alice', '+1', 2, 0, 'GetYourGuide', 30, 'B-A', '', '', '', '', '', 'booking'],
  [DATE, '10:00 AM', 'English', 'Bob', '+1', 1, 0, 'Guruwalk', 6, 'B-B', '', '', '', '', '', 'booking'],
  [DATE, '5:00 PM', 'English', 'Vip', '+1', 4, 0, 'GetYourGuide', 99.98, 'B-P', 'Private', '', '', '', '', 'booking'],
  [DATE, '5:00 PM', 'English', 'Reg', '+1', 2, 0, 'GetYourGuide', 30, 'B-R', '', '', '', '', '', 'booking']]);

// --- Ledger with Rates.
const ledger = new __mock.MockSS('ledger'); __mock.SS_BY_ID['LEDSTRESS'] = ledger; __mock.PROPS['LEDGER_ID'] = 'LEDSTRESS';
ledger.insertSheet('Rates').getRange(1, 1, 6, 2).setValues([
  ['Setting', 'Value'],
  ['Paid tour — we owe guide (€ per checked-in person)', 10],
  ['Free tour — guide owes us (€ per checked-in person)', 6],
  ['Paid sources (comma separated)', 'Viator, GetYourGuide, Airbnb'],
  ['Free tour commission — guruwalk (€ per checked-in person)', 4.7],
  ['Paid private - we owe guide (€ per private tour)', 75]]);

const token = makeToken_('Albert');
const save = (bId, lang, time, guests, guide, priv, source) => apiSave_({ token: token, data: JSON.stringify({
  dateKey: DATE, time: time, timeLabel: time, day: '', language: lang, guide: guide || 'Carlos',
  bookings: [{ bookingId: bId, source: source || 'GetYourGuide', name: bId, phone: '+1', guests: guests, children: 0, income: 30, isPrivate: !!priv, manualNote: '', checked: true, checkedIn: guests }] }) });
const tours = () => apiTours_({ token: token, days: 45 });
const bookingOf = (r, time, id) => { const s = (r.allTours || []).find(x => x.dateKey === DATE && x.time === time && x.language === 'English' && (time === '17:00' ? true : true)); return null; };
const findBk = (r, time, id, priv) => {
  const ss = (r.allTours || []).filter(x => x.dateKey === DATE && x.time === time && x.language === 'English' && (priv === undefined || !!x.isPrivate === !!priv));
  for (const s of ss) { const b = (s.bookings || []).find(b => b.bookingId === id); if (b) return b; }
  return null;
};

console.log('=== STRESS: check-in durability ===');
__RRX = {};
// 1. Double-tap: same booking checked in twice.
save('B-A', 'English', '10:00', 2);
save('B-A', 'English', '10:00', 2);
let led = readLedgerForGuides_(['Carlos']);
check('double-tap check-in makes ONE ledger row (no duplicate)', Object.keys(led.checkins).filter(k => k.indexOf('B-A') !== -1).length === 1, Object.keys(led.checkins));

// 2. Stale cross-device save omits B-A -> must NOT wipe it.
apiSave_({ token: token, data: JSON.stringify({ dateKey: DATE, time: '10:00', language: 'English', guide: 'Carlos',
  bookings: [{ bookingId: 'B-B', source: 'Guruwalk', name: 'Bob', phone: '+1', guests: 1, children: 0, income: 6, isPrivate: false, checked: true, checkedIn: 1 }] }) });
led = readLedgerForGuides_(['Carlos']);
check('a save that omits an earlier check-in does NOT wipe it', Object.keys(led.checkins).some(k => k.indexOf('B-A') !== -1), Object.keys(led.checkins));
check('the later check-in is also there', Object.keys(led.checkins).some(k => k.indexOf('B-B') !== -1), Object.keys(led.checkins));

// 3. The feed-only portal shows both check-ins.
__RRX = {};
let r = tours();
check('portal (feed-only) shows B-A checked in', (findBk(r, '10:00', 'B-A') || {}).checked === true, findBk(r, '10:00', 'B-A'));
check('portal shows B-B checked in', (findBk(r, '10:00', 'B-B') || {}).checked === true, findBk(r, '10:00', 'B-B'));

// 4. Undo B-A, then re-check-in -> works, still one row.
apiUncheckin_({ token: token, bookingId: 'B-A' });
led = readLedgerForGuides_(['Carlos']);
check('undo removes B-A from the ledger', !Object.keys(led.checkins).some(k => k.indexOf('B-A') !== -1), Object.keys(led.checkins));
__RRX = {}; r = tours();
check('portal shows B-A NOT checked after undo', (findBk(r, '10:00', 'B-A') || {}).checked === false, findBk(r, '10:00', 'B-A'));
save('B-A', 'English', '10:00', 2);
led = readLedgerForGuides_(['Carlos']);
check('re-check-in after undo works, still ONE row', Object.keys(led.checkins).filter(k => k.indexOf('B-A') !== -1).length === 1, Object.keys(led.checkins));

// 5. SAFETY NET (cannot-miss): a check-in in the ledger but MISSING from the
//    feed (rebuild race / same-day booking checked in before its feed row) must
//    still show, via the ledger union. Assign the shift so the check-in lands in
//    an assigned guide's tab (exactly as production does).
__RRX = {};
apiAssign_({ token: token, dateKey: DATE, time: '10:00', language: 'English', isPrivate: '', guide: 'Carlos', force: '1' });
save('B-A', 'English', '10:00', 2, 'Carlos');
writeFeedUncheckin_('B-A');                 // simulate the feed transiently losing the check-in
__RRX = {}; r = tours();
check('a check-in in the ledger but MISSING from the feed still shows (cannot-miss safety net)', (findBk(r, '10:00', 'B-A') || {}).checked === true, findBk(r, '10:00', 'B-A'));
// clean up: unassign so later sections start neutral
apiAssign_({ token: token, dateKey: DATE, time: '10:00', language: 'English', isPrivate: '', guide: '', force: '1' });

console.log('=== STRESS: private vs regular at the same slot ===');
__RRX = {};
apiAssign_({ token: token, dateKey: DATE, time: '17:00', language: 'English', isPrivate: '1', guide: 'Albert', force: '1' });
apiAssign_({ token: token, dateKey: DATE, time: '17:00', language: 'English', isPrivate: '', guide: 'Carlos', force: '1' });
__RRX = {}; r = tours();
const priv17 = (r.allTours || []).find(x => x.dateKey === DATE && x.time === '17:00' && x.isPrivate);
const reg17 = (r.allTours || []).find(x => x.dateKey === DATE && x.time === '17:00' && !x.isPrivate);
check('private 17:00 is its own shift assigned to Albert', priv17 && (priv17.assigned || []).indexOf('Albert') !== -1, priv17 && priv17.assigned);
check('regular 17:00 is a separate shift assigned to Carlos', reg17 && (reg17.assigned || []).indexOf('Carlos') !== -1, reg17 && reg17.assigned);
check('private shift shows only the private booking', priv17 && priv17.bookings.length === 1 && priv17.bookings[0].bookingId === 'B-P', priv17 && priv17.bookings.map(b => b.bookingId));
check('regular shift shows only the regular booking', reg17 && reg17.bookings.length === 1 && reg17.bookings[0].bookingId === 'B-R', reg17 && reg17.bookings.map(b => b.bookingId));
// check in the private -> flat 75, separate from regular
save('B-P', 'English', '17:00', 4, 'Albert', true);
save('B-R', 'English', '17:00', 2, 'Carlos', false);
const ledA = readLedgerForGuides_(['Albert']).reservations, ledC = readLedgerForGuides_(['Carlos']).reservations;
check('private + regular check-ins land in different guide tabs (no wipe)',
  Object.keys(ledA).length >= 0 && Object.keys(readLedgerForGuides_(['Albert']).checkins).some(k => k.indexOf('B-P') !== -1) &&
  Object.keys(readLedgerForGuides_(['Carlos']).checkins).some(k => k.indexOf('B-R') !== -1), null);

console.log('=== STRESS: malformed feed rows must not crash the read ===');
__RRX = {};
feed.getRange(6, 1, 3, 16).setValues([
  ['', '10:00 AM', 'English', 'NoDate', '+1', 1, 0, 'GetYourGuide', 10, 'B-BAD1', '', '', '', '', '', 'booking'],
  [DATE, '', '', 'NoTimeLang', '+1', -5, 0, 'GetYourGuide', 10, 'B-BAD2', '', '', '', '', '', 'booking'],
  [DATE, '10:00 AM', 'English', 'HugeGroup', '+1', 9999, 0, 'GetYourGuide', 10, 'B-BAD3', '', '', '', '', '', 'booking']]);
let crashed = false, r2 = null;
try { __RRX = {}; r2 = tours(); } catch (e) { crashed = true; }
check('a malformed feed (blank date/lang, negative/huge guests) does NOT crash the read', !crashed && r2 && r2.ok === true, crashed);
check('the blank-date row does not become a phantom shift', !(r2.allTours || []).some(s => s.bookings.some(b => b.bookingId === 'B-BAD1')), null);
check('the huge-group booking still appears (no silent drop)', !!findBk(r2, '10:00', 'B-BAD3'), null);
// clean up the malformed rows
feed.deleteRow(8); feed.deleteRow(7); feed.deleteRow(6);

console.log('=== STRESS: assign / unassign / delete / reopen lifecycle ===');
__RRX = {};
// bare assign (no booking) -> a placeholder appears
apiAssign_({ token: token, dateKey: DATE, time: '11:00', language: 'English', isPrivate: '', guide: 'Carlos', force: '1' });
__RRX = {}; r = tours();
const bare = (r.allTours || []).find(x => x.dateKey === DATE && x.time === '11:00' && x.language === 'English');
check('a bare assignment shows as an empty shift with the guide', bare && (bare.assigned || []).indexOf('Carlos') !== -1 && bare.bookings.length === 0, bare && [bare && bare.assigned, bare && bare.bookings.length]);
// unassign it -> placeholder gone
apiAssign_({ token: token, dateKey: DATE, time: '11:00', language: 'English', isPrivate: '', guide: '', force: '1' });
__RRX = {}; r = tours();
check('unassigning the bare slot removes the empty shift', !(r.allTours || []).some(x => x.dateKey === DATE && x.time === '11:00' && x.language === 'English'), null);

console.log('=== STRESS: an inactive / wrong-language guide cannot be assigned ===');
__RRX = {};
const badLang = apiAssign_({ token: token, dateKey: DATE, time: '10:00', language: 'English', isPrivate: '', guide: 'Francesca', force: '1' });
check('a guide who does not speak the language is rejected', badLang && badLang.ok === false, badLang);
const inactive = apiAssign_({ token: token, dateKey: DATE, time: '10:00', language: 'English', isPrivate: '', guide: 'Inactive', force: '1' });
check('an inactive guide is rejected', inactive && inactive.ok === false, inactive);

console.log('=== STRESS: a weekly default must NOT block reassigning that guide (flexibility) ===');
__RRX = {};
// Albert is the weekly default for 10:00 English on DATE's weekday.
control.getSheetByName('Weekly_Schedule').getRange(2, 1, 1, 7).setValues([[dayNameFromKey_(DATE), '10:00', 'English', 1, '', '', 'Albert']]);
__RRX = {}; let rw = tours();
let s10 = (rw.allTours || []).find(x => x.dateKey === DATE && x.time === '10:00' && x.language === 'English' && !x.isPrivate);
check('the 10:00 shift shows Albert by weekly default', s10 && (s10.assigned || []).indexOf('Albert') !== -1, s10 && s10.assigned);
// Assign Albert to the adjacent 11:00 WITHOUT force -> the conflict check must
// NOT count his 10:00 DEFAULT as a hard clash.
const asg11 = apiAssign_({ token: token, dateKey: DATE, time: '11:00', language: 'English', isPrivate: '', guide: 'Albert' });
check('a guide defaulted onto 10:00 CAN be assigned to the adjacent 11:00 (no phantom conflict)', asg11 && asg11.ok === true, asg11);
__RRX = {}; rw = tours();
s10 = (rw.allTours || []).find(x => x.dateKey === DATE && x.time === '10:00' && x.language === 'English' && !x.isPrivate);
const s11 = (rw.allTours || []).find(x => x.dateKey === DATE && x.time === '11:00' && x.language === 'English');
check('Albert is really assigned to 11:00', s11 && (s11.assigned || []).indexOf('Albert') !== -1, s11 && s11.assigned);
check('the 10:00 default YIELDS to his real 11:00 (no double-book)', s10 && (s10.assigned || []).indexOf('Albert') === -1, s10 && s10.assigned);

console.log('=== STRESS: a 0-guest check-in flags a PAID tour as a no-show (do not miss reporting) ===');
__RRX = {};
const yday = dayKey(-1);
const cl = booking.insertSheet('Completed Log');
cl.getRange(1, 1, 2, 12).setValues([
  ['Date', 'Time', 'Language', 'Name', 'Phone', 'Adults', 'Children', 'Source', 'Income', 'Booking ID', 'Notes', 'Logged'],
  [yday, '10:00', 'English', 'NoShow Guy', '+1', 2, 0, 'Viator', 23, 'BR-NOSHOW1', '', '']]);
// Guide checked them in with ZERO guests (nobody showed).
apiSave_({ token: token, data: JSON.stringify({ dateKey: yday, time: '10:00', language: 'English', guide: 'Carlos',
  bookings: [{ bookingId: 'BR-NOSHOW1', source: 'Viator', name: 'NoShow Guy', phone: '+1', guests: 2, children: 0, income: 23, isPrivate: false, checked: true, checkedIn: 0 }] }) });
try { ensureQueueTabs_(ledgerSS_()); } catch (e) { try { ensureQueueTabs_(); } catch (e2) {} }
updateNoShowQueues_();
const vsh = ledger.getSheetByName(QUEUE_TABS.VIATOR_NOSHOW);
const vrows = (vsh && vsh.getLastRow() >= 1) ? vsh.getRange(1, 1, vsh.getLastRow(), NOSHOW_HEADERS.length).getValues() : [];
check('a Viator booking checked in with 0 guests IS flagged as a no-show', vrows.some(r => String(r[4] || '') === 'BR-NOSHOW1'), vrows.map(r => r[4]).filter(Boolean));

console.log('=== STRESS: a non-manager cannot assign / undo ===');
const gtok = makeToken_('Carlos');
check('a guide cannot assign', apiAssign_({ token: gtok, dateKey: DATE, time: '10:00', language: 'English', guide: 'Carlos' }).ok === false, null);
check('a guide cannot undo a check-in', apiUncheckin_({ token: gtok, bookingId: 'B-A' }).ok === false, null);
check('a bad token is rejected everywhere', apiTours_({ token: 'garbage.sig' }).ok === false, null);

console.log('=== STRESS: manager moves — time, language, and BOTH at once ===');
__RRX = {};
// Language Tours tabs the move code reads, plus clean bookings dedicated to the
// move tests (so the check-in stress above can't interfere) — in the feed too.
const AH = ['Name', 'Phone', 'Number of Guests', 'Tour date', 'Time', 'Source', 'Income', 'Booking ID', 'Notes'];
const enT = booking.insertSheet('English Tours');
enT.getRange(1, 1, 1, 9).setValues([AH]);
enT.getRange(2, 1, 2, 9).setValues([
  ['MoveA', '+1', 2, DATE, '10:00 AM', 'GetYourGuide', 30, 'B-M1', ''],
  ['MoveB', '+1', 1, DATE, '10:00 AM', 'GetYourGuide', 20, 'B-M2', '']]);
const esT = booking.insertSheet('Spanish Tours');
esT.getRange(1, 1, 1, 9).setValues([AH]);
feed.getRange(feed.getLastRow() + 1, 1, 2, 16).setValues([
  [DATE, '10:00 AM', 'English', 'MoveA', '+1', 2, 0, 'GetYourGuide', 30, 'B-M1', '', '', '', '', '', 'booking'],
  [DATE, '10:00 AM', 'English', 'MoveB', '+1', 1, 0, 'GetYourGuide', 20, 'B-M2', '', '', '', '', '', 'booking']]);
const anyBk = (rr, id) => { for (const s of (rr.allTours || [])) { const b = (s.bookings || []).find(b => b.bookingId === id); if (b) return { tour: s, b }; } return null; };

// -- Time move --------------------------------------------------------------
let mr = apiMoveBookingTime_({ token: token, bookingId: 'B-M1', language: 'English', toTime: '4:30 PM' });
check('time move ok + moved', mr.ok === true && mr.moved === true, mr);
let mt1 = anyBk(tours(), 'B-M1');
check('B-M1 now shows at 16:30 (feed regrouped on the next poll)', mt1 && mt1.tour.time === '16:30' && mt1.tour.language === 'English', mt1 && [mt1.tour.language, mt1.tour.time]);
check('English Tours tab time cell is 4:30 PM (so the rebuild keeps it)', String(enT.getRange(2, 5).getValue()) === '4:30 PM', enT.getRange(2, 5).getValue());
check('re-moving to the same time is a no-op', apiMoveBookingTime_({ token: token, bookingId: 'B-M1', language: 'English', toTime: '16:30' }).moved === false, null);
check('garbage time is rejected', apiMoveBookingTime_({ token: token, bookingId: 'B-M1', language: 'English', toTime: 'nope' }).ok === false, null);
check('B-M1 still at 16:30 after a rejected move', anyBk(tours(), 'B-M1').tour.time === '16:30', null);
check('a non-manager cannot move a time', apiMoveBookingTime_({ token: gtok, bookingId: 'B-M1', language: 'English', toTime: '11:00' }).ok === false, null);
check('an unknown booking id is rejected', apiMoveBookingTime_({ token: token, bookingId: 'GHOST', language: 'English', toTime: '11:00' }).ok === false, null);

// -- Combined move: language + time in ONE call (no intermediate tour) -------
let mc = apiMoveBooking_({ token: token, bookingId: 'B-M2', fromLanguage: 'English', toLanguage: 'Spanish', toTime: '11:00' });
check('combined move ok + moved', mc.ok === true && mc.moved === true, mc);
let mt2 = anyBk(tours(), 'B-M2');
check('B-M2 now Spanish 11:00 in one step', mt2 && mt2.tour.language === 'Spanish' && mt2.tour.time === '11:00', mt2 && [mt2.tour.language, mt2.tour.time]);
check('B-M2 row physically in the Spanish Tours tab at 11:00 AM', esT.getRange(2, 1, Math.max(1, esT.getLastRow() - 1), 9).getValues().some(x => String(x[7]) === 'B-M2' && String(x[4]) === '11:00 AM'), esT.getDataRange().getValues());
check('B-M2 removed from the English Tours tab (not duplicated)', !enT.getRange(2, 1, enT.getLastRow() - 1, 9).getValues().some(x => String(x[7]) === 'B-M2'), null);
check('combined move to the same language + time is a no-op', apiMoveBooking_({ token: token, bookingId: 'B-M2', fromLanguage: 'Spanish', toLanguage: 'Spanish', toTime: '11:00' }).moved === false, null);
check('a non-manager cannot do a combined move', apiMoveBooking_({ token: gtok, bookingId: 'B-M2', fromLanguage: 'Spanish', toLanguage: 'English', toTime: '10:00' }).ok === false, null);

// -- A moved booking must NOT carry its old tour's guide into the new tour ---
enT.getRange(enT.getLastRow() + 1, 1, 1, 9).setValues([['MoveC', '+1', 2, DATE, '10:00 AM', 'GetYourGuide', 30, 'B-M3', '']]);
feed.getRange(feed.getLastRow() + 1, 1, 1, 16).setValues([[DATE, '10:00 AM', 'English', 'MoveC', '+1', 2, 0, 'GetYourGuide', 30, 'B-M3', '', '', '', '', 'Carlos', 'booking']]);
check('sanity: B-M3 starts assigned to Carlos (its feed Guide)', (function () { const x = anyBk(tours(), 'B-M3'); return x && (x.tour.assigned || []).indexOf('Carlos') !== -1; })(), null);
apiMoveBookingTime_({ token: token, bookingId: 'B-M3', language: 'English', toTime: '3:00 PM' });
const m3 = anyBk(tours(), 'B-M3');
check('a moved booking lands in an UNASSIGNED tour (old guide dropped)', m3 && m3.tour.time === '15:00' && (!m3.tour.assigned || m3.tour.assigned.length === 0), m3 && [m3.tour.time, m3.tour.assigned]);

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
