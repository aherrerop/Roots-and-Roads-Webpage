/**
 * Unit tests for the bot's pure matching helpers — the part that maps a Viator
 * departure card to the sheet's guest-count key and computes the time window.
 * No browser / network. Run: node selftest.js
 */
const { normTime, titleLang, lookupGuests, departureMs, cardOnDate } = require('./close-empty-tours.js');

let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

// normTime — must match the sheet's non-zero-padded hour.
check('normTime 16:00', normTime('16:00') === '16:00', normTime('16:00'));
check('normTime 09:00 -> 9:00', normTime('09:00') === '9:00', normTime('09:00'));
check('normTime 10:30', normTime('10:30') === '10:30', normTime('10:30'));

// titleLang — Viator "english tour" -> sheet "English".
check('titleLang english tour', titleLang('english tour') === 'English', titleLang('english tour'));
check('titleLang ITALIAN', titleLang('italian') === 'Italian', titleLang('italian'));
check('titleLang empty', titleLang('') === '', titleLang(''));

// lookupGuests — builds "date|H:mm|Language" and reads the map.
const gm = { '2026-08-28|16:00|Italian': 0, '2026-08-28|17:00|English': 3, '2026-08-28|10:00|English': 2 };
check('lookup Italian 16:00 empty', lookupGuests(gm, '2026-08-28', { time24: '16:00', language: 'italian' }) === 0, null);
check('lookup English 17:00 has 3', lookupGuests(gm, '2026-08-28', { time24: '17:00', language: 'English Tour' }) === 3, null);
check('lookup missing slot -> 0', lookupGuests(gm, '2026-08-28', { time24: '11:00', language: 'french' }) === 0, null);
check('lookup zero-pad tolerated', lookupGuests(gm, '2026-08-28', { time24: '10:00', language: 'english' }) === 2, null);

// departureMs — absolute instant from date + time + tz offset.
const nowIso = '2026-08-28T09:00:00+02:00';
const nowMs = Date.parse(nowIso);
const mins = (t) => Math.round((departureMs('2026-08-28', t, '+02:00') - nowMs) / 60000);
check('16:00 is 420m after 09:00', mins('16:00') === 420, mins('16:00'));
check('10:00 is 60m after 09:00', mins('10:00') === 60, mins('10:00'));
check('09:00 is 0m (already at start)', mins('9:00') === 0, mins('9:00'));

// cardOnDate — day-header "Fri, Aug 28" belongs to 2026-08-28.
check('cardOnDate match Fri Aug 28', cardOnDate({ dayLabel: 'Fri, Aug 28' }, '2026-08-28') === true, null);
check('cardOnDate mismatch Sat Aug 29', cardOnDate({ dayLabel: 'Sat, Aug 29' }, '2026-08-28') === false, null);
check('cardOnDate no label -> trust nav', cardOnDate({ dayLabel: '' }, '2026-08-28') === true, null);

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
