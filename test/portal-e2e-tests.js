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

// DETERMINISTIC CLOCK: this suite seeds "today" tours (e.g. a 10:00 AM tour) and
// checks they surface. Left to the real wall clock the suite passed in the morning
// but FAILED every afternoon, because a 10:00 tour is legitimately "over" by then
// and the portal correctly hides it. Freeze "now" to today at 08:00 so upcoming-
// today tours are always upcoming — the suite is time-of-day independent. Only the
// no-arg `new Date()`/`Date.now()` are frozen; every explicit date still parses.
(function () {
  const RealDate = Date;
  const FIXED = (function () { const d = new RealDate(); d.setHours(8, 0, 0, 0); return d.getTime(); })();
  function FakeDate(...a) { return a.length ? new RealDate(...a) : new RealDate(FIXED); }
  FakeDate.now = () => FIXED;
  FakeDate.parse = RealDate.parse; FakeDate.UTC = RealDate.UTC;
  FakeDate.prototype = RealDate.prototype;
  // eslint-disable-next-line no-global-assign
  Date = FakeDate;
})();

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

console.log('--- FULL (20+) tour: red-outline signal + needs-2nd-guide flag ---');
// Run BEFORE the feed-only stage below (which creates the Portal Feed); here the
// portal still builds from the booking tabs, so a new en-tab booking surfaces.
en.getRange(en.getLastRow() + 1, 1, 1, 9).setValues([
  ['Big Group', '+34600999000', 20, new Date(DATE + 'T12:00:00'), '2:00 PM', 'GetYourGuide', 300, 'GYGFULL20', '']]);
const rFull = apiTours_({ token: token });
const shFull = (rFull.allTours || []).find(s => s.dateKey === DATE && s.time === '14:00' && s.language === 'English');
check('a 20-guest group tour is flagged full (red outline)', shFull && shFull.full === true, shFull && [shFull && shFull.bookedGuests, shFull && shFull.full]);
check('a full tour with <2 guides flags needsSecondGuide', shFull && shFull.needsSecondGuide === true, shFull && shFull.needsSecondGuide);
const shSmall = (rFull.allTours || []).find(s => s.dateKey === DATE && s.time === '12:00' && s.language === 'English');
check('a small (<20) group tour is NOT full', shSmall && shSmall.full === false, shSmall && [shSmall && shSmall.bookedGuests, shSmall && shSmall.full]);
// CHILDREN DON'T COUNT toward the 20: 19 adults + 5 children = 24 people but only
// 19 ADULTS, so the tour is NOT full (the cap is adults one guide can lead).
en.getRange(en.getLastRow() + 1, 1, 1, 9).setValues([
  ['Kids Group', '+34600777000', 19, new Date(DATE + 'T12:00:00'), '9:00 AM', 'GetYourGuide', 285, 'GYGKIDS19', 'Family · 5 children']]);
const rKids = apiTours_({ token: token });
const shKids = (rKids.allTours || []).find(s => s.dateKey === DATE && s.time === '9:00' && s.language === 'English');
check('children are counted and shown (bookedChildren=5)', shKids && shKids.bookedChildren === 5, shKids && [shKids && shKids.bookedGuests, shKids && shKids.bookedChildren]);
check('19 adults + 5 children is NOT full (children excluded from the 20)', shKids && shKids.full === false, shKids && [shKids && shKids.bookedGuests, shKids && shKids.bookedChildren, shKids && shKids.full]);

console.log('--- Second guide (B): a portal overlay, never touches grid/feed/scheduler ---');
const aSg = apiAssign_({ token: token, dateKey: DATE, time: '14:00', language: 'English', guide: 'Albert', slot: 2 });
check('apiAssign_ slot 2 returns ok', aSg && aSg.ok === true && aSg.slot === 2, aSg);
const rSg = apiTours_({ token: token });
const shSg = (rSg.allTours || []).find(s => s.dateKey === DATE && s.time === '14:00' && s.language === 'English');
check('the 2nd guide shows as secondGuide + joins assigned', shSg && shSg.secondGuide === 'Albert' && (shSg.assigned || []).indexOf('Albert') !== -1, shSg && [shSg && shSg.secondGuide, shSg && shSg.assigned]);
check('the 2nd-guide assign is overlay-only (did NOT create the feed; still grid mode)', rSg.timings && rSg.timings.sched !== undefined, rSg.timings);
// The co-guide SEES the tour in their own My-tours (assigned includes them).
const rAlbert = apiTours_({ token: makeToken_('Albert') });
check('the second guide sees the full tour in their own tours', (rAlbert.tours || []).some(s => s.dateKey === DATE && s.time === '14:00' && s.language === 'English'), (rAlbert.tours || []).map(s => s.time));
// Private tours have a single guide — slot 2 is refused.
const rPrivSg = apiAssign_({ token: token, dateKey: DATE, time: '14:00', language: 'English', guide: 'Albert', slot: 2, isPrivate: '1' });
check('a 2nd guide is refused on a private tour', rPrivSg && rPrivSg.ok === false, rPrivSg);
// Clearing the 2nd guide (empty) removes it.
apiAssign_({ token: token, dateKey: DATE, time: '14:00', language: 'English', guide: '', slot: 2 });
const shSg0 = (apiTours_({ token: token }).allTours || []).find(s => s.dateKey === DATE && s.time === '14:00' && s.language === 'English');
check('clearing the 2nd guide removes secondGuide', shSg0 && !shSg0.secondGuide, shSg0 && shSg0.secondGuide);

console.log('--- Split check-in (C): each guide paid for who THEY check in ---');
apiAssign_({ token: token, dateKey: DATE, time: '14:00', language: 'English', guide: 'Albert', slot: 2 });   // 2nd guide = Albert
en.getRange(en.getLastRow() + 1, 1, 1, 9).setValues([
  ['Split Pair', '+34600888000', 2, new Date(DATE + 'T12:00:00'), '2:00 PM', 'GetYourGuide', 30, 'GYGSPLIT2', '']]);
// GYGFULL20 -> guide 1 (Carlos); GYGSPLIT2 -> guide 2 (Albert).
const splitSave = { dateKey: DATE, time: '14:00', timeLabel: '2:00 PM', day: dayNameFromKey_(DATE),
  language: 'English', guide: 'Carlos', secondGuide: 'Albert', bookings: [
    { bookingId: 'GYGFULL20', source: 'GetYourGuide', name: 'Big Group', guests: 20, income: 300, isPrivate: false, checked: true, checkedIn: 20, guideIndex: 1 },
    { bookingId: 'GYGSPLIT2', source: 'GetYourGuide', name: 'Split Pair', guests: 2, income: 30, isPrivate: false, checked: true, checkedIn: 2, guideIndex: 2 } ] };
const rSplit = apiSave_({ token: token, data: JSON.stringify(splitSave) });
check('apiSave_ split returns ok (2 rows written)', rSplit && rSplit.ok === true && rSplit.saved === 2, rSplit);
const idsIn = n => { const sh = ledgerSS_().getSheetByName(n); return sh ? sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), LEDGER_HEADERS.length).getValues().map(r => String(r[LEDGER_BOOKINGID_COL] || '')) : []; };
check('guide 1 (Carlos) ledger has ONLY his guest', idsIn('Carlos').indexOf('GYGFULL20') !== -1 && idsIn('Carlos').indexOf('GYGSPLIT2') === -1, idsIn('Carlos'));
check('guide 2 (Albert) ledger has ONLY his guest', idsIn('Albert').indexOf('GYGSPLIT2') !== -1 && idsIn('Albert').indexOf('GYGFULL20') === -1, idsIn('Albert'));
const shSplit = (apiTours_({ token: token }).allTours || []).find(s => s.dateKey === DATE && s.time === '14:00' && s.language === 'English');
const bFull = shSplit && shSplit.bookings.find(b => b.bookingId === 'GYGFULL20');
const bPair = shSplit && shSplit.bookings.find(b => b.bookingId === 'GYGSPLIT2');
check('the card marks each guest under the right guide (1 vs 2)', bFull && bFull.guideIndex === 1 && bPair && bPair.guideIndex === 2, [bFull && bFull.guideIndex, bPair && bPair.guideIndex]);
// Undo clears the guide-2 marker AND removes the ledger row from guide 2's tab.
apiUncheckin_({ token: token, bookingId: 'GYGSPLIT2' });
const bPair2 = ((apiTours_({ token: token }).allTours || []).find(s => s.dateKey === DATE && s.time === '14:00' && s.language === 'English') || {}).bookings;
const bP2 = bPair2 && bPair2.find(b => b.bookingId === 'GYGSPLIT2');
check('undo clears the guide-2 marker + removes the ledger row', bP2 && bP2.guideIndex === 1 && idsIn('Albert').indexOf('GYGSPLIT2') === -1, [bP2 && bP2.guideIndex, idsIn('Albert')]);
// Co-guide fix: a slot-1 guest is paid to the PRIMARY guide (d.guide), even when a
// NON-manager co-guide (slot 2) is the one saving — not the saver's own tab.
const coSave = { dateKey: DATE, time: '14:00', timeLabel: '2:00 PM', day: dayNameFromKey_(DATE),
  language: 'English', guide: 'Albert', secondGuide: 'Carlos', bookings: [
    { bookingId: 'GYGCO1', source: 'GetYourGuide', name: 'Primary Guest', guests: 2, income: 30, isPrivate: false, checked: true, checkedIn: 2, guideIndex: 1 } ] };
apiSave_({ token: makeToken_('Carlos'), data: JSON.stringify(coSave) });   // Carlos (non-manager co-guide) saves
check('a slot-1 guest is credited to the PRIMARY guide, not the saver', idsIn('Albert').indexOf('GYGCO1') !== -1 && idsIn('Carlos').indexOf('GYGCO1') === -1, [idsIn('Albert').indexOf('GYGCO1'), idsIn('Carlos').indexOf('GYGCO1')]);

console.log('--- Full no-show (D): flat €10 to the guide when nobody shows ---');
en.getRange(en.getLastRow() + 1, 1, 1, 9).setValues([
  ['Ghost Group', '+34600777000', 4, new Date(DATE + 'T12:00:00'), '4:00 PM', 'GetYourGuide', 60, 'GYGNOS1', '']]);
const nsOk = apiNoShow_({ token: token, dateKey: DATE, time: '16:00', language: 'English', guide: 'Carlos' });
check('apiNoShow_ returns ok (paying Carlos)', nsOk && nsOk.ok === true && nsOk.guide === 'Carlos', nsOk);
const nsRows = () => { const sh = ledgerSS_().getSheetByName('Carlos'); return sh ? sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), LEDGER_HEADERS.length).getValues().filter(r => /^NOSHOW\|/.test(String(r[LEDGER_BOOKINGID_COL] || ''))) : []; };
check('a flat €10 "No-show" ledger line is written to the guide', (function () { const rs = nsRows(); return rs.length === 1 && Number(rs[0][10]) === 10 && String(rs[0][13]) === 'No-show'; })(), nsRows().map(r => [r[10], r[13]]));
const nsView = (apiTours_({ token: token }).allTours || []).find(s => s.dateKey === DATE && s.time === '16:00' && s.language === 'English');
check('the card shows the no-show state + guide', nsView && nsView.noShow === true && nsView.noShowGuide === 'Carlos', nsView && [nsView && nsView.noShow, nsView && nsView.noShowGuide]);
apiNoShow_({ token: token, dateKey: DATE, time: '16:00', language: 'English', guide: 'Carlos' });   // re-flag
check('re-flagging does NOT duplicate the €10 line (once per tour)', nsRows().length === 1, nsRows().length);
const nsDeny = apiNoShow_({ token: makeToken_('Carlos'), dateKey: DATE, time: '16:00', language: 'English', guide: 'Albert' });
check('a non-manager cannot flag a no-show paying ANOTHER guide', nsDeny && nsDeny.ok === false, nsDeny);
apiNoShow_({ token: token, dateKey: DATE, time: '16:00', language: 'English', clear: '1' });
const nsCleared = (apiTours_({ token: token }).allTours || []).find(s => s.dateKey === DATE && s.time === '16:00' && s.language === 'English');
check('clearing a no-show removes the line + the state', nsRows().length === 0 && nsCleared && !nsCleared.noShow, [nsRows().length, nsCleared && nsCleared.noShow]);

console.log('--- Per-phase timings + seniority-ordered eligible list ---');
const rt = apiTours_({ token: token });
check('response reports per-phase timings (sched/book/ledger)', rt.timings && typeof rt.timings.sched === 'number' && typeof rt.timings.book === 'number' && typeof rt.timings.ledger === 'number', rt.timings);
// Carlos is seniority 1, Albert seniority 2 -> Carlos must sort first, even though
// "Albert" is alphabetically earlier. Proves seniority beats alphabetical.
check('English eligible list is most-senior-first (Carlos before Albert)', (rt.guidesByLanguage.English || [])[0] === 'Carlos', rt.guidesByLanguage.English);

console.log('--- Assign dropdown: per-guide availability status dots (info only, still assignable) ---');
// Direct unit of guideStatusesForShift_ using the real helpers. Target: DATE 12:00 English.
const gbl = { English: ['Carlos', 'Albert', 'Polina'] };
const bm = {
  carlos: [{ ms: shiftStartMs_(DATE, 720), k: 'someother|12:00' }],   // SAME time (12:00) -> busy
  albert: [{ ms: shiftStartMs_(DATE, 840), k: 'someother|14:00' }]    // 2h later, within 5h -> overlap
  // Polina: no tours -> free
};
const st = guideStatusesForShift_({ dateKey: DATE, minutes: 720, language: 'English' }, gbl, bm);
const byName = {}; st.forEach(s => { byName[s.name] = s.status; });
check('a guide with a SAME-time tour is busy (red)', byName.Carlos === 'busy', byName);
check('a guide with a tour within the 5h window is overlap (yellow)', byName.Albert === 'overlap', byName);
check('a free guide is free (green)', byName.Polina === 'free', byName);
check('EVERY language guide is listed (assignable regardless of status)', st.length === 3, st.map(s => s.name));
// A guide's OWN assignment on THIS shift must not mark them busy against themselves.
const selfKey = shiftKeyFull_({ dateKey: DATE, minutes: 720, language: 'English' });
const selfSt = guideStatusesForShift_({ dateKey: DATE, minutes: 720, language: 'English' },
  { English: ['Carlos'] }, { carlos: [{ ms: shiftStartMs_(DATE, 720), k: selfKey }] });
check('a guide assigned to THIS shift is NOT marked busy by their own assignment', selfSt[0].status === 'free', selfSt);
// The live payload carries guideOptions on manager tours.
const anyShift = (rt.allTours || []).find(s => (s.guideOptions || []).length);
check('apiTours_ attaches guideOptions [{name,status}] to manager tours', !!anyShift && anyShift.guideOptions.every(o => o.name && ['free','overlap','busy'].indexOf(o.status) !== -1), anyShift && anyShift.guideOptions);

console.log('--- Manager window: near tours load first, far tours behind "Load more" ---');
const FAR = dayKey(40);
en.getRange(4, 1, 1, 9).setValues([
  ['Far Future', '+34600555000', 2, new Date(FAR + 'T12:00:00'), '10:00 AM', 'GetYourGuide', 30, 'GYGE2EFAR', '']]);
const rw = apiTours_({ token: token });                       // default near window
check('response carries a freshness timestamp (HH:mm:ss)', /^\d{1,2}:\d{2}:\d{2}$/.test(rw.now || ''), rw.now);
check('default load window is 3 days (both roles)', rw.windowDays === 3, rw.windowDays);
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

console.log('--- Delete a tour: frees the guide (clears the grid cell); column prune is DEFERRED ---');
const DEL = dayKey(6);   // an isolated slot nothing else touches
apiAssign_({ token: token, dateKey: DEL, time: '19:00', language: 'English', guide: 'Carlos', force: '1' });
let gEng = control.getSheetByName('Schedule_English').getDataRange().getDisplayValues();
const delColIdx = gEng[1].findIndex(h => /(^|\D)19:00|7:00 PM/.test(String(h)));
check('assigning created a 19:00 column', delColIdx > -1, gEng[1]);
// The date row for DEL, before the close: Carlos is in the 19:00 cell.
const delRowIdx = gEng.findIndex((row, i) => i >= 2 && gridLabelToKey_(String(row[0] || '').trim(),
  gridAnchor_(String(control.getSheetByName('Schedule_English').getRange(1, 1).getDisplayValue() || ''))) === DEL);
check('the 19:00 cell held the assigned guide before close', delRowIdx > -1 && /Carlos/.test(String(gEng[delRowIdx][delColIdx] || '')), delRowIdx > -1 ? gEng[delRowIdx][delColIdx] : null);
apiCloseShift_({ token: token, id: shiftKey_(DEL, 19 * 60, 'english') });
gEng = control.getSheetByName('Schedule_English').getDataRange().getDisplayValues();
// The ESSENTIAL effect: the guide is cleared from the cell (freed). The empty
// column is NOT pruned synchronously any more — that cosmetic housekeeping was
// moved off closeShift's locked path (it cost ~20s and stalled the whole portal);
// the weekly makeSchedule prunes + restyles instead. So we assert the CELL is
// empty, not that the column is gone.
check('after close the guide is cleared from the 19:00 cell (freed)', delRowIdx === -1 || String(gEng[delRowIdx][delColIdx] || '').trim() === '', delRowIdx > -1 ? gEng[delRowIdx][delColIdx] : '(row gone)');
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

console.log('--- Check-in on the feed: write mirrors to the feed, union read shows it ---');
// writeFeedCheckin_ upserts the feed row's own columns (M/N) by booking id. A
// later write may RAISE the count but must KEEP the first check-in time — the
// same no-reset guarantee as the ledger (so a re-save never collapses times).
const wrote = writeFeedCheckin_('GYGE2E001', 3, '10:09');
const feedRow = feed.getRange(2, 1, 1, 14).getValues()[0];
check('writeFeedCheckin_ raises the count but KEEPS the first time (M=3, N=10:03)', wrote === true && Number(feedRow[12]) === 3 && String(feedRow[13]) === '10:03', [feedRow[12], feedRow[13]]);
// The union read surfaces a check-in that lives ONLY in the feed (no ledger row).
const rUnion = apiTours_({ token: token, days: 45 });
const uShift = (rUnion.allTours || []).filter(s => s.dateKey === DATE && s.time === '10:00' && s.language === 'English');
const uBk = uShift.length ? uShift[0].bookings.find(b => b.bookingId === 'GYGE2E001') : null;
check('the portal shows the feed check-in via the union read', uBk && uBk.checked === true && uBk.checkedIn === 3, uBk);
check('the feed keeps the ORIGINAL check-in time (10:03, not reset to 10:09)', uBk && uBk.checkedAt === '10:03', uBk && uBk.checkedAt);
check('the poll UNIONs the ledger as a cannot-miss safety net (ledger IS read)', rUnion.timings && typeof rUnion.timings.ledger === 'number', rUnion.timings);

console.log('--- Stage 4: the poll builds shifts FROM THE FEED (no schedule-grid read) ---');
writeFeedGuide_(DATE, '10:00', 'English', false, 'Carlos');       // assign via the feed Guide column
ensureFeedShiftRow_(DATE, '15:00', 'English', false, 'Albert');   // an EMPTY assignable shift (no weekly rule at 15:00)
const rs4 = apiTours_({ token: token, days: 45 });
check('feed-only: the poll did NOT read the schedule grid (no sched phase)', rs4.timings && rs4.timings.sched === undefined, rs4.timings);
const s10 = (rs4.allTours || []).find(s => s.dateKey === DATE && s.time === '10:00' && s.language === 'English');
check('a shift with bookings takes its guide from the FEED (Carlos)', s10 && (s10.assigned || []).indexOf('Carlos') !== -1, s10 && s10.assigned);
check('the guest list excludes the shift placeholder (real bookings only)', s10 && s10.bookings.length >= 1 && s10.bookings.every(b => b.bookingId), s10 && s10.bookings.map(b => b.bookingId));
const s15 = (rs4.allTours || []).find(s => s.dateKey === DATE && s.time === '15:00' && s.language === 'English');
check('an EMPTY shift shows from its placeholder row, assigned to Albert', s15 && (s15.assigned || []).indexOf('Albert') !== -1 && s15.bookings.length === 0, s15 && [s15 && s15.assigned, s15 && s15.bookings.length]);
removeFeedShiftRow_(DATE, '15:00', 'English', false);
const rs4b = apiTours_({ token: token, days: 45 });
check('removing the placeholder drops the empty shift', !(rs4b.allTours || []).some(s => s.dateKey === DATE && s.time === '15:00' && s.language === 'English'), null);

console.log('--- Manager UNDO check-in: clears the ledger row + feed M/N atomically ---');
// Seed a check-in in BOTH stores for a fresh booking, then undo it.
apiSave_({ token: token, data: JSON.stringify({ dateKey: DATE, time: '10:00', timeLabel: '10:00 AM', day: '', language: 'English', guide: 'Carlos',
  bookings: [{ bookingId: 'GYGUNDO1', source: 'GetYourGuide', name: 'Undo Me', phone: '+1', guests: 2, children: 0, income: 30, isPrivate: false, manualNote: '', checked: true, checkedIn: 2 }] }) });
const undoBefore = readLedgerForGuides_(['Carlos']).checkins;
check('the seeded check-in is in the ledger', Object.keys(undoBefore).some(k => k.indexOf('GYGUNDO1') !== -1), Object.keys(undoBefore));
const feedVerBeforeUndo = String(__mock.PROPS['PORTAL_FEED_VER'] || '0');
const rUn = apiUncheckin_({ token: token, bookingId: 'GYGUNDO1' });
check('apiUncheckin_ ok + removed a ledger row', rUn && rUn.ok === true && rUn.removed >= 1, rUn);
const undoAfter = readLedgerForGuides_(['Carlos']).checkins;
check('the ledger no longer has that check-in', !Object.keys(undoAfter).some(k => k.indexOf('GYGUNDO1') !== -1), Object.keys(undoAfter));
check('undo bumps the feed version so the portal reflects it next load', String(__mock.PROPS['PORTAL_FEED_VER'] || '0') !== feedVerBeforeUndo, [feedVerBeforeUndo, __mock.PROPS['PORTAL_FEED_VER']]);
check('a non-manager cannot undo (managers only)', (function(){ const r = apiUncheckin_({ token: makeToken_('Carlos'), bookingId: 'X' }); return r && r.ok === false && /manager/i.test(r.error || ''); })(), null);

console.log('--- Save reports write timings + refreshes ONLY the feed cache (freshness) ---');
const feedVerBefore = String(__mock.PROPS['PORTAL_FEED_VER'] || '0');
const globalVerBefore = String(__mock.PROPS['PORTAL_CACHE_VER'] || '0');
const rSave = apiSave_({ token: token, data: JSON.stringify(data) });
check('apiSave_ returns per-phase write timings (ledger/feed/flush)', rSave.ok && rSave.timings && ('ledger' in rSave.timings) && ('feed' in rSave.timings) && ('flush' in rSave.timings), rSave.timings);
check('a check-in bumps the FEED cache version -> next load reads it (Up to date is true)', String(__mock.PROPS['PORTAL_FEED_VER'] || '0') !== feedVerBefore, [feedVerBefore, __mock.PROPS['PORTAL_FEED_VER']]);
check('a check-in does NOT bump the global cache (schedule/guides/weekly stay warm)', String(__mock.PROPS['PORTAL_CACHE_VER'] || '0') === globalVerBefore, [globalVerBefore, __mock.PROPS['PORTAL_CACHE_VER']]);

console.log('--- Private tours: "Paid private" rate + private/regular kept separate in the ledger ---');
const ratesSheet = ledgerSS_().getSheetByName('Rates');
ratesSheet.getRange(ratesSheet.getLastRow() + 1, 1, 1, 2).setValues([['Paid private - we owe guide (€ per private tour)', 80]]);
check('readRates_ reads the "Paid private" label (80)', readRates_().privatePay === 80, readRates_().privatePay);
// The owner's live label style: "Paid SF tour - we owe guide (€ per checked-in
// person)" (no platform named) must set EVERY SF source, even though it says
// neither "exterior" nor "sagrada".
ratesSheet.getRange(ratesSheet.getLastRow() + 1, 1, 1, 2).setValues([['Paid SF tour - we owe guide (€ per checked-in person)', 3]]);
const sfR = readRates_();
check('readRates_ reads a plain "Paid SF tour" label for GYG-SF', sfR.sfPay['GYG-SF'] === 3, sfR.sfPay);
check('… and applies the same generic SF rate to Viator-SF', sfR.sfPay['Viator-SF'] === 3, sfR.sfPay);
// A platform-specific row overrides just that platform.
ratesSheet.getRange(ratesSheet.getLastRow() + 1, 1, 1, 2).setValues([['SF exterior (Viator) - we owe guide (€ per checked-in person)', 5]]);
const sfR2 = readRates_();
check('a "(Viator)" SF row overrides only Viator-SF', sfR2.sfPay['Viator-SF'] === 5 && sfR2.sfPay['GYG-SF'] === 3, sfR2.sfPay);
// A private and a regular tour at the SAME slot must not wipe each other in the ledger.
const lrow = (bid, type, we) => makeLedgerRow_({ dateKey: DATE, day: 'x', timeLabel: '4:00 PM', language: 'English',
  bookingName: bid, phone: '', source: 'GetYourGuide', guests: 1, children: 0, checkedIn: 1, weOwe: we, theyOwe: 0, rrMakes: 0, type: type, bookingId: bid, note: '' });
writeGuideLedger_('Carlos', DATE, '16:00', 'English', [lrow('REG1', 'Paid', 10)]);
writeGuideLedger_('Carlos', DATE, '16:00', 'English', [lrow('PRIV1', 'Private', 80)]);   // same slot, private
const carLedger = ledgerSS_().getSheetByName('Carlos').getDataRange().getValues();
check('the private check-in does NOT wipe the regular one at the same slot',
  carLedger.some(r => String(r[LEDGER_BOOKINGID_COL]) === 'REG1') && carLedger.some(r => String(r[LEDGER_BOOKINGID_COL]) === 'PRIV1'),
  carLedger.map(r => r[LEDGER_BOOKINGID_COL]).filter(Boolean));

console.log('--- Data retention (GDPR): clears OLD guest name+phone, keeps recent + counts ---');
const drs = control.insertSheet('DR_Test');
drs.getRange(1, 1, 3, 4).setValues([
  ['Name', 'Phone', 'Guests', 'Tour date'],
  ['Old Guest', '+34600000001', 2, '2023-01-15'],   // > 12 months ago -> should clear
  ['Recent Guest', '+34600000002', 3, dayKey(-5)]]); // 5 days ago -> should keep
const drCut = dataRetentionCutoffKey_();
const drDry = purgeSheetPII_(drs, 3, [0, 1], drCut, true);
check('dry run flags exactly the 1 old record', drDry === 1, drDry);
check('dry run changes NOTHING (old name still present)', drs.getRange(2, 1).getValue() === 'Old Guest', drs.getRange(2, 1).getValue());
const drPurge = purgeSheetPII_(drs, 3, [0, 1], drCut, false);
check('purge clears exactly the 1 old record', drPurge === 1, drPurge);
check('old guest NAME is cleared', drs.getRange(2, 1).getValue() === '', drs.getRange(2, 1).getValue());
check('old guest PHONE is cleared', drs.getRange(2, 2).getValue() === '', drs.getRange(2, 2).getValue());
check('old row KEEPS its anonymized count (2)', Number(drs.getRange(2, 3).getValue()) === 2, drs.getRange(2, 3).getValue());
check('recent guest is UNTOUCHED', drs.getRange(3, 1).getValue() === 'Recent Guest', drs.getRange(3, 1).getValue());
check('a second purge is a no-op (idempotent)', purgeSheetPII_(drs, 3, [0, 1], drCut, false) === 0, null);

console.log('--- MANAGER HISTORY: last 2 days from Past2Days (mark no-shows, undo check-ins) ---');
const YDAY = dayKey(-1);
const past = booking.insertSheet('Past2Days');
past.getRange(1, 1, 1, 16).setValues([['Date', 'Time', 'Language', 'Name', 'Phone', 'Adults', 'Children',
  'Source', 'Income', 'Booking ID', 'Notes', 'Manager note', 'Checked-in', 'Check-in time', 'Guide', 'Type']]);
past.getRange(2, 2, 3, 1).setNumberFormat('@'); past.getRange(2, 10, 3, 1).setNumberFormat('@'); past.getRange(2, 14, 3, 1).setNumberFormat('@');
past.getRange(2, 1, 3, 16).setValues([
  [YDAY, '10:00 AM', 'English', 'Hist One', '+34600000010', 3, 0, 'GetYourGuide', 45, 'HIST001', '', '', 3, '10:05', 'Carlos', 'booking'],
  [YDAY, '03:00 PM', 'English', 'Hist Two', '+34600000011', 4, 0, 'Guruwalk', 0, 'HIST002', '', '', '', '', 'Carlos', 'booking'],
  [YDAY, '05:00 PM', 'English', 'Hist Three', '+34600000012', 2, 0, 'GetYourGuide', 30, 'HIST003', '', '', 2, '17:05', 'Albert', 'booking']]);
let rHist = apiHistory_({ token: token });
check('apiHistory_ ok for a manager, flagged history, window = 2 days', rHist && rHist.ok === true && rHist.history === true && rHist.days === 2, rHist && rHist.error);
const h1 = (rHist.tours || []).find(t => t.dateKey === YDAY && t.time === '10:00' && t.language === 'English');
const h2 = (rHist.tours || []).find(t => t.dateKey === YDAY && t.time === '15:00' && t.language === 'English');
check('a past tour surfaces with its guide (Carlos) from the snapshot', h1 && (h1.assigned || []).indexOf('Carlos') !== -1, h1 && h1.assigned);
check('the past tour shows its check-in from the snapshot M/N', h1 && h1.bookings.some(b => b.bookingId === 'HIST001' && b.checked === true && b.checkedIn === 3), h1 && h1.bookings);
check('the un-checked past tour shows its guest as not checked in', h2 && h2.bookings.some(b => b.bookingId === 'HIST002' && b.checked === false), h2 && h2.bookings);
check('a MANAGER sees ALL guides\' past tours (incl. Albert\'s)', (rHist.tours || []).some(t => t.dateKey === YDAY && t.time === '17:00' && (t.assigned || []).indexOf('Albert') !== -1), (rHist.tours || []).map(t => t.time));
// GUIDE history: Carlos can open it, flagged non-manager, and sees ONLY his own tours.
const rGuide = apiHistory_({ token: makeToken_('Carlos') });
check('a GUIDE can open history (not managers-only), flagged manager:false', rGuide && rGuide.ok === true && rGuide.history === true && rGuide.manager === false, rGuide && rGuide.error);
check('the guide sees their OWN past tours (HIST001/HIST002)', (rGuide.tours || []).some(t => t.bookings.some(b => b.bookingId === 'HIST001')) && (rGuide.tours || []).some(t => t.bookings.some(b => b.bookingId === 'HIST002')), (rGuide.tours || []).map(t => t.time));
check('the guide does NOT see another guide\'s tour (Albert\'s HIST003) or its guests', !(rGuide.tours || []).some(t => (t.bookings || []).some(b => b.bookingId === 'HIST003')), (rGuide.tours || []).map(t => t.time));
// Mark the un-checked past tour as a total no-show (the guide is still paid).
const nsPast = apiNoShow_({ token: token, dateKey: YDAY, time: '15:00', language: 'English', guide: 'Carlos' });
check('apiNoShow_ accepts a PAST tour (history window)', nsPast && nsPast.ok === true, nsPast);
rHist = apiHistory_({ token: token });
const h2b = (rHist.tours || []).find(t => t.dateKey === YDAY && t.time === '15:00' && t.language === 'English');
check('the past no-show shows on the history card (paid to Carlos)', h2b && h2b.noShow === true && h2b.noShowGuide === 'Carlos', h2b && [h2b && h2b.noShow, h2b && h2b.noShowGuide]);
// MANAGER REMOVES the past no-show -> it disappears from the history card.
const nsClear = apiNoShow_({ token: token, dateKey: YDAY, time: '15:00', language: 'English', guide: 'Carlos', clear: '1' });
check('apiNoShow_ clear ok for a manager on a PAST tour', nsClear && nsClear.ok === true && nsClear.cleared === true, nsClear);
rHist = apiHistory_({ token: token });
const h2c = (rHist.tours || []).find(t => t.dateKey === YDAY && t.time === '15:00' && t.language === 'English');
check('removing the no-show clears it from the history card', h2c && !h2c.noShow, h2c && h2c.noShow);
// Undo the past check-in: clears the snapshot's M/N, so history shows it un-checked.
const unPast = apiUncheckin_({ token: token, bookingId: 'HIST001' });
check('apiUncheckin_ ok for a past-tour check-in', unPast && unPast.ok === true, unPast);
rHist = apiHistory_({ token: token });
const h1b = (rHist.tours || []).find(t => t.dateKey === YDAY && t.time === '10:00' && t.language === 'English');
check('undo clears the past check-in in the history view (snapshot M/N cleared)', h1b && h1b.bookings.some(b => b.bookingId === 'HIST001' && b.checked === false), h1b && h1b.bookings);

// REGRESSION: a past tour whose check-in lives ONLY in the ledger (no Past2Days
// snapshot row) must still show the check-in + the guide who ran it — the whole
// point of History. Seed the Completed Log (roster) + a ledger check-in, NO snapshot.
const clog = booking.insertSheet('Completed Log');
clog.getRange(1, 1, 1, 12).setValues([['Date', 'Time', 'Language', 'Name', 'Phone', 'Adults', 'Children', 'Source', 'Income', 'Booking ID', 'Notes', 'Logged']]);
clog.getRange(2, 2, 1, 1).setNumberFormat('@'); clog.getRange(2, 10, 1, 1).setNumberFormat('@');
clog.getRange(2, 1, 1, 12).setValues([[new Date(YDAY + 'T12:00:00'), '11:00 AM', 'English', 'Ledger Only', '+34600000020', 2, 0, 'Guruwalk', 0, 'HISTLED1', '', '']]);
writeGuideLedger_('Carlos', YDAY, '11:00', 'English', [makeLedgerRow_({ dateKey: YDAY, day: dayNameFromKey_(YDAY), timeLabel: '11:00 AM', language: 'English',
  bookingName: 'Ledger Only', phone: '+34600000020', source: 'Guruwalk', guests: 2, children: 0, checkedIn: 2, weOwe: 0, theyOwe: 10, rrMakes: 10, type: 'Free', bookingId: 'HISTLED1', note: '' })]);
rHist = apiHistory_({ token: token });
const hL = (rHist.tours || []).find(t => t.dateKey === YDAY && t.time === '11:00' && t.language === 'English');
check('a ledger-only past tour surfaces (roster from Completed Log)', hL && hL.bookings.some(b => b.bookingId === 'HISTLED1'), hL && hL.bookings);
check('its check-in shows from the LEDGER even with no snapshot row', hL && hL.bookings.some(b => b.bookingId === 'HISTLED1' && b.checked === true && b.checkedIn === 2), hL && hL.bookings);
check('the tour guide is recovered from the ledger (Carlos)', hL && (hL.assigned || []).indexOf('Carlos') !== -1, hL && hL.assigned);

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
