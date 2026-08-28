/* ===== Viator auto-close: booking-side pure logic =====
   Bundled with mock.js + booking sources (incl. viatorAutoClose.gs).
   Covers the 2FA-code extraction (must NOT grab the Supplier ID) and the
   slot-start time math. Sheet-dependent functions are exercised live in prod. */
let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

console.log('--- 2FA code extraction (real email shape) ---');
const realBody =
  "Hi HERRERO PARAREDA,\n" +
  "You requested a two-factor authentication code for Roots and Roads, Supplier ID: 5631527.\n" +
  "Here's your unique code: 572431\n" +
  "Your code expires in 20 minutes.";
const m = realBody.match(VIATOR_CODE_REGEX);
check('extracts the code 572431', m && m[1] === '572431', m && m[1]);
check('does NOT grab Supplier ID 5631527', m && m[1] !== '5631527', m && m[1]);

const noCode = "Supplier ID: 5631527. Someone tried to log in.";
check('no "unique code:" -> no match', !noCode.match(VIATOR_CODE_REGEX), null);

const subjectCase = "your UNIQUE CODE:   99887";
check('case-insensitive + spaces', (subjectCase.match(VIATOR_CODE_REGEX) || [])[1] === '99887', null);

console.log('--- slot start time (Europe/Madrid project TZ) ---');
const s1 = viatorSlotStart_(2026, 7, 28, '16:00');
check('16:00 -> hour 16, min 0', s1 && s1.getHours() === 16 && s1.getMinutes() === 0, s1 && [s1.getHours(), s1.getMinutes()]);
const s2 = viatorSlotStart_(2026, 7, 28, '9:00');
check('9:00 -> hour 9', s2 && s2.getHours() === 9, s2 && s2.getHours());
check('garbage time -> null', viatorSlotStart_(2026, 7, 28, 'xx') === null, null);

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
