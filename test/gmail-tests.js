/* ===== Gmail-side integration tests =====
   Drives the REAL confirmation / cancellation / modification passes against a
   stateful Gmail + Sheets mock, and asserts the outcomes that matter:
     - a booking actually lands on the list,
     - "Processed" is applied ONLY when the email was truly handled,
     - a cancelled booking's confirmation is grouped under Cancellations and
       archived,
     - an undeterminable language is registered but flagged UNREAD.
   Bundled with mock.js + gmail-mock.js + booking/*.gs. */
let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

const P = RNR.LABELS.PROCESSED;
const LANGS = ['English Tours', 'German Tours', 'Spanish Tours', 'Italian Tours', 'French Tours'];

function freshBookingSheet() {
  const ss = new __mock.MockSS('booking');
  LANGS.forEach(n => ss.insertSheet(n).getRange(1, 1, 1, 9).setValues([
    ['Name', 'Phone', 'Number of Guests', 'Tour date', 'Time', 'Source', 'Income', 'Booking ID', 'Notes']]));
  __mock.SS_BY_ID['1rGCfe138BeRXrcyvx6H-9y7IGg-BTCi_-N1-AEM0BCw'] = ss;
  SpreadsheetApp._active = ss;
  return ss;
}
function resetWorld() {
  __gmail.reset();
  __gmail.ensure(Object.values(RNR.LABELS));   // labels must exist for safeAddLabel_
  freshBookingSheet();
  resetRunCaches_();
  RNR_SKIP_PROCESSED_ = false;   // audit mode: re-read all threads
  RNR_RUN_STARTED_AT_ = Date.now();
}
const gygBody = (id, name, dateStr, adults, lang) => [
  '¡Hola! Buenas noticias.', 'Se ha reservado tu producto',
  'Barcelona Ultimate Tour', 'Tour en inglés',
  'Número de referencia ' + id,
  'Fecha ' + dateStr,
  'Número de participantes', adults + ' x Adults (Edad 14 - 99)',
  'Cliente principal', name + ' customer-x@reply.getyourguide.com Teléfono: +34600000000 Idioma: ' + lang,
  'Idioma del tour', lang + ' (Live tour guide)'
].join('\n');
const idsOnList = () => activeBookingIdSet_();

console.log('--- Confirmation: registered -> Processed; invalid -> NOT Processed ---');
resetWorld();
const okThread = __gmail.add([RNR.LABELS.GYG_CONFIRM],
  __gmail.msg('Booking - S1 - GYGCONF001', gygBody('GYGCONF001', 'Alice', 'August 4, 2030 10:00 AM', 2, 'Inglés')));
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
check('a valid confirmation lands on the booking list', idsOnList()['GYGCONF001'] === true, Object.keys(idsOnList()));
check('a valid confirmation thread IS marked Processed', okThread.hasLabel(P) === true, [...okThread.labelSet]);

resetWorld();
const badThread = __gmail.add([RNR.LABELS.GYG_CONFIRM],
  __gmail.msg('Booking - S2 - GYGBAD002', ['Se ha reservado tu producto', 'Número de referencia GYGBAD002',
    'Número de participantes', '2 x Adults (Edad 14 - 99)'].join('\n')));   // no name, no date -> invalid
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
check('an invalid confirmation is NOT put on the list', !idsOnList()['GYGBAD002'], Object.keys(idsOnList()));
check('an invalid confirmation is NOT marked Processed (no lying label)', badThread.hasLabel(P) === false, [...badThread.labelSet]);

console.log('--- 2nd GetYourGuide account (destinationstewards) reaches the Portal Feed as "GYG" ---');
// A tour date within the feed window, whenever the suite runs.
const _soon = new Date(Date.now() + 10 * 86400000);
const _dateStr = _soon.toLocaleString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }) + ' 10:00 AM';
resetWorld();
const dsThread = __gmail.add([RNR.LABELS.GYG_CONFIRM],
  __gmail.msg('Reserva - S888888 - GYGDEST01', gygBody('GYGDEST01', 'Dieter', _dateStr, 2, 'Inglés'),
    new Date(), 'destinationstewards@gmail.com'));   // <- forwarded 2nd-account email keeps this To
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
check('2nd-account confirmation lands on the booking list', idsOnList()['GYGDEST01'] === true, Object.keys(idsOnList()));
check('2nd-account confirmation IS marked Processed', dsThread.hasLabel(P) === true, [...dsThread.labelSet]);
rebuildPortalFeed_();   // the guide portal reads exactly this tab
const _feed = SpreadsheetApp._active.getSheetByName(RNR.SHEETS.PORTAL_FEED);
const _rows = _feed.getLastRow() > 1 ? _feed.getRange(2, 1, _feed.getLastRow() - 1, 16).getValues() : [];
const _dsRow = _rows.find(r => String(r[9]).trim() === 'GYGDEST01');   // J = Booking ID
check('the booking reaches the Portal Feed (this is what the portal shows)', !!_dsRow, _rows.map(r => r[9]));
check('its Source is "GYG" on the feed (not GetYourGuide)', _dsRow && String(_dsRow[7]) === 'GYG', _dsRow && _dsRow[7]);
check('the English tour routed it to the English feed row', _dsRow && String(_dsRow[2]) === 'English', _dsRow && _dsRow[2]);

// Control: a normal GYG email (to the main inbox) stays "GetYourGuide".
resetWorld();
__gmail.add([RNR.LABELS.GYG_CONFIRM], __gmail.msg('Reserva - S779080 - GYGMAIN01', gygBody('GYGMAIN01', 'Hans', _dateStr, 2, 'Inglés')));
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
rebuildPortalFeed_();
const _m = SpreadsheetApp._active.getSheetByName(RNR.SHEETS.PORTAL_FEED);
const _mRow = _m.getRange(2, 1, 1, 16).getValues()[0];
check('the main account stays Source = "GetYourGuide"', String(_mRow[7]) === 'GetYourGuide', _mRow[7]);

console.log('--- Confirmation: undeterminable language is registered but flagged UNREAD ---');
resetWorld();
const langThread = __gmail.add([RNR.LABELS.GYG_CONFIRM],
  __gmail.msg('Booking - S3 - GYGLANG003', gygBody('GYGLANG003', 'Bob', 'August 5, 2030 11:00 AM', 2, 'Portuguese')));
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
check('unknown-language booking is still registered (English default)', idsOnList()['GYGLANG003'] === true, Object.keys(idsOnList()));
check('unknown-language thread is Processed (it WAS registered)', langThread.hasLabel(P) === true, null);
check('unknown-language thread is marked UNREAD for management attention', langThread.isUnread() === true, null);

console.log('--- Cancellation: booking removed, and confirmation grouped under Cancellations + archived ---');
resetWorld();
const conf = __gmail.add([RNR.LABELS.GYG_CONFIRM],
  __gmail.msg('Booking - S4 - GYGCANC004', gygBody('GYGCANC004', 'Carol', 'August 6, 2030 10:00 AM', 2, 'Inglés')));
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
check('the booking is on the list before cancellation', idsOnList()['GYGCANC004'] === true, null);
const canc = __gmail.add([RNR.LABELS.GYG_CANCEL],
  __gmail.msg('Se ha cancelado una reserva - S4 - GYGCANC004',
    ['GYGCANC004 ha sido cancelada', 'Número de referencia: GYGCANC004', 'Nombre: Carol', 'Fecha: 6 de agosto de 2030, 10:00'].join('\n')));
resetRunCaches_();
processCancellationLabel_(RNR.LABELS.GYG_CANCEL, RNR.SOURCE.GYG);
check('cancelled booking is removed from the list', !idsOnList()['GYGCANC004'], Object.keys(idsOnList()));
check('the cancellation thread itself is Processed + archived', canc.hasLabel(P) && canc.isInInbox() === false, [...canc.labelSet]);
check('the CONFIRMATION is relabeled to Cancellations (grouped)', conf.hasLabel(RNR.LABELS.GYG_CANCEL) && !conf.hasLabel(RNR.LABELS.GYG_CONFIRM), [...conf.labelSet]);
check('the CONFIRMATION leaves the inbox too', conf.isInInbox() === false, conf.isInInbox());

console.log('--- A cancellation with NO readable id is still FINALISED (stops the retry flood) ---');
resetWorld();
const vagueCancel = __gmail.add([RNR.LABELS.GYG_CANCEL],
  __gmail.msg('Se ha cancelado una reserva', 'Una reserva ha sido cancelada. Gracias.'));
processCancellationLabel_(RNR.LABELS.GYG_CANCEL, RNR.SOURCE.GYG);
check('an id-less cancellation is finalised (Processed) so it never retries forever', vagueCancel.hasLabel(P) === true, [...vagueCancel.labelSet]);

console.log('--- A cancellation the parser MISSES is saved by the label + subject id (the flood fix) ---');
resetWorld();
__gmail.add([RNR.LABELS.GYG_CONFIRM], __gmail.msg('Booking - S1 - GYGREMOVE99', gygBody('GYGREMOVE99', 'Rob', 'August 9, 2030 10:00 AM', 2, 'Inglés')));
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
check('the booking to be cancelled is on the list', !!idsOnList()['GYGREMOVE99'], null);
resetRunCaches_();
// Gmail filed this under Cancellations, but it reads like a confirmation and has
// no parseable cancellation body — the per-source parser yields nothing.
const missed = __gmail.add([RNR.LABELS.GYG_CANCEL], __gmail.msg('Booking - S1 - GYGREMOVE99', ''));
processCancellationLabel_(RNR.LABELS.GYG_CANCEL, RNR.SOURCE.GYG);
check('the missed cancellation STILL removes the booking (via the subject id)', !idsOnList()['GYGREMOVE99'], Object.keys(idsOnList()));
check('the missed cancellation thread is finalised (Processed) — no flood', missed.hasLabel(P) === true, [...missed.labelSet]);

console.log('--- END-TO-END list: 3 bookings + a cancellation, judged by PRODUCTION\'s own invariant oracle ---');
resetWorld();
__gmail.add([RNR.LABELS.GYG_CONFIRM], __gmail.msg('Booking - S10 - GYGAAA010', gygBody('GYGAAA010', 'Ann', 'August 4, 2030 10:00 AM', 2, 'Inglés')));
__gmail.add([RNR.LABELS.GYG_CONFIRM], __gmail.msg('Booking - S11 - GYGBBB011', gygBody('GYGBBB011', 'Bob', 'August 5, 2030 11:00 AM', 3, 'Italiano')));
__gmail.add([RNR.LABELS.GYG_CONFIRM], __gmail.msg('Booking - S12 - GYGCCC012', gygBody('GYGCCC012', 'Cy', 'August 6, 2030 5:00 PM', 1, 'Inglés')));
processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
check('all three bookings are on the list', ['GYGAAA010', 'GYGBBB011', 'GYGCCC012'].every(id => idsOnList()[id]), Object.keys(idsOnList()));
check('the Italian booking routed to Italian Tours', SpreadsheetApp._active.getSheetByName('Italian Tours').getRange(2, 8).getDisplayValue() === 'GYGBBB011', null);
check("PRODUCTION invariants hold (checkInvariants_ = 0)", checkInvariants_() === 0, checkInvariants_());

// Cancel the middle booking.
__gmail.add([RNR.LABELS.GYG_CANCEL], __gmail.msg('Se ha cancelado una reserva - S11 - GYGBBB011',
  ['GYGBBB011 ha sido cancelada', 'Número de referencia: GYGBBB011', 'Nombre: Bob'].join('\n')));
resetRunCaches_();
processCancellationLabel_(RNR.LABELS.GYG_CANCEL, RNR.SOURCE.GYG);
check('cancelled booking is removed', !idsOnList()['GYGBBB011'], Object.keys(idsOnList()));
check('the other two remain', idsOnList()['GYGAAA010'] && idsOnList()['GYGCCC012'], null);
check('invariants STILL hold after the cancellation (nothing cancelled-still-active)', checkInvariants_() === 0, checkInvariants_());

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
