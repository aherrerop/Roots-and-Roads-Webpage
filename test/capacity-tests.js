/* ===== Full-tour capacity alerts =====
   A group tour that reaches the people cap emails management ONCE (close it /
   add a 2nd guide), with everything in the subject line. Bundled with mock.js +
   booking sources. Exercises the REAL capacityAlerts_ / readUpcomingTourGroups_. */
let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

// Capture outgoing mail instead of sending it.
let sent = [];
MailApp.sendEmail = (o) => { sent.push(o); };

const booking = new __mock.MockSS('booking'); SpreadsheetApp._active = booking;
const FH = RNR.PORTAL_FEED_HEADERS;
const feed = booking.insertSheet('Portal Feed');
feed.getRange(1, 1, 1, FH.length).setValues([FH]);

const dayKey = o => { const d = new Date(); d.setDate(d.getDate() + o); return Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd'); };
const D = dayKey(3);   // an upcoming date

// One 16-column feed row from a small spec.
let _n = 0;
const frow = (o) => [
  o.date || D, o.time || '10:30 AM', o.lang || 'English', o.name || ('G' + (++_n)), '+1',
  o.adults || 0, o.children || 0, o.source || 'GetYourGuide', 0, o.id || ('B' + (++_n)),
  o.notes || '', '', '', '', o.guide || '', o.type || 'booking'
];

// Rewrite the whole feed body from a list of specs (clears stale rows first).
function setFeed(specs) {
  const last = feed.getLastRow();
  if (last >= 2) feed.getRange(2, 1, last - 1, FH.length).clearContent();
  const rows = specs.map(frow);
  if (rows.length) {
    feed.getRange(2, 2, rows.length, 1).setNumberFormat('@');    // Time as text (no Date coercion)
    feed.getRange(2, 10, rows.length, 1).setNumberFormat('@');   // Booking ID as text
    feed.getRange(2, 1, rows.length, FH.length).setValues(rows);
  }
}
const resetState = () => PropertiesService.getScriptProperties().deleteProperty(RNR.CAPACITY_ALERT_PROP);

console.log('--- grouping: readUpcomingTourGroups_ (per date/time/language/variant, people = adults+children) ---');
resetState();
setFeed([
  { time: '10:30 AM', source: 'GetYourGuide', adults: 14, guide: 'Carlos' },
  { time: '10:30 AM', source: 'Viator', adults: 6, children: 2 },       // slot totals 20 adults (+2 children)
  { time: '5:00 PM', source: 'GetYourGuide', adults: 5 },               // a quiet tour
  { time: '10:30 AM', source: 'GetYourGuide', adults: 25, notes: 'Private' }, // private -> excluded
  { time: '10:30 AM', source: 'GYG-SF', adults: 4 }                     // SF -> its OWN tour
]);
const groups = readUpcomingTourGroups_();
const k3h = D + '|10:30 AM|english|3h';
const kSf = D + '|10:30 AM|english|SF';
const k17 = D + '|5:00 PM|english|3h';
check('3h 10:30 slot sums ADULTS across bookings = 20 (children tracked separately)', groups[k3h] && groups[k3h].adults === 20 && groups[k3h].children === 2 && groups[k3h].people === 22, groups[k3h]);
check('the SF booking is a SEPARATE tour (not merged into the 3h count)', groups[kSf] && groups[kSf].adults === 4, groups[kSf]);
check('the private 25-person booking is EXCLUDED from any group', !Object.values(groups).some(g => g.adults === 25 || g.adults === 45), Object.keys(groups));
check('per-source breakdown is kept (GetYourGuide 14, Viator 8)', groups[k3h] && groups[k3h].sources['GetYourGuide'] === 14 && groups[k3h].sources['Viator'] === 8, groups[k3h] && groups[k3h].sources);
check('the quiet 17:00 tour is tracked but well under the cap', groups[k17] && groups[k17].adults === 5, groups[k17]);

console.log('--- a tour reaching the cap sends ONE alert, all info in the subject ---');
resetState(); sent = [];
capacityAlerts_();
check('exactly one alert email is sent', sent.length === 1, sent.map(s => s.subject));
const sub = (sent[0] || {}).subject || '';
check('subject leads with the red dot then the language', /^🔴 English ·/.test(sub), sub);
check('subject carries the date (no year) then the time', new RegExp(Utilities.formatDate(new Date(D + 'T12:00:00'), 'Europe/Madrid', 'EEE d MMM') + ' 10:30 AM').test(sub), sub);
check('subject date has NO year', !/20\d\d/.test(sub), sub);
check('subject flags it FULL with the ADULT count (+ children noted)', /FULL\s+20 adults \+2k/.test(sub), sub);
check('subject states the action (close or add 2nd guide)', /close or add 2nd guide/i.test(sub), sub);
const body = (sent[0] || {}).body || '';
check('body leads with adults (cap) and notes children + total', /Adults:\s+20\s+\(plus 2 children/.test(body) && /22 people total/.test(body), body);
check('body carries the per-source breakdown', /GetYourGuide: 14/.test(body) && /Viator: 8/.test(body), body);
check('body carries the assigned guide', /Guide:\s+Carlos/.test(body), body);
check('the alert goes to the internal manager address', (sent[0] || {}).to === RNR.INTERNAL_ALERT_TO, sent[0] && sent[0].to);

console.log('--- fires ONCE: no re-alert on the next runs, even as it grows ---');
sent = [];
capacityAlerts_();                                    // same feed, second pass
check('a re-run does NOT re-alert the same full tour', sent.length === 0, sent.map(s => s.subject));
setFeed([
  { time: '10:30 AM', source: 'GetYourGuide', adults: 14, guide: 'Carlos' },  // grew to 22
  { time: '10:30 AM', source: 'Viator', adults: 6, children: 2 }
]);
sent = [];
capacityAlerts_();
check('growing further (22) still does NOT re-alert', sent.length === 0, sent.map(s => s.subject));

console.log('--- re-arms: drops below the cap, then refills -> alerts again ---');
setFeed([{ time: '10:30 AM', source: 'GetYourGuide', adults: 10 }]);   // now 10 people
sent = [];
capacityAlerts_();
check('dropping below the cap sends no alert', sent.length === 0, sent.map(s => s.subject));
setFeed([
  { time: '10:30 AM', source: 'GetYourGuide', adults: 18 },
  { time: '10:30 AM', source: 'Viator', adults: 2 }
]);                                                   // back to 20
sent = [];
capacityAlerts_();
check('refilling to the cap alerts again (re-armed)', sent.length === 1, sent.map(s => s.subject));

console.log('--- children do NOT count toward the cap (adults only, matches the portal) ---');
resetState(); sent = [];
setFeed([{ time: '11:00 AM', source: 'GetYourGuide', adults: 19, children: 5 }]);   // 24 people but only 19 adults
capacityAlerts_();
check('19 adults + 5 children does NOT cross the cap (children excluded)', sent.length === 0, sent.map(s => s.subject));
setFeed([{ time: '11:00 AM', source: 'GetYourGuide', adults: 20, children: 5 }]);   // now 20 adults
sent = [];
capacityAlerts_();
check('20 adults crosses the cap regardless of children', sent.length === 1 && /FULL\s+20 adults \+5k/.test(sent[0].subject), sent.map(s => s.subject));

console.log('--- SF and 3h at the same slot are counted SEPARATELY (neither crosses alone) ---');
resetState(); sent = [];
setFeed([
  { time: '3:30 PM', source: 'GetYourGuide', adults: 15 },   // 3h = 15
  { time: '3:30 PM', source: 'GYG-SF', adults: 15 }          // SF = 15
]);
capacityAlerts_();
check('a 15-person 3h + 15-person SF at one slot do NOT merge to 30 (no alert)', sent.length === 0, sent.map(s => s.subject));

console.log('--- several tours cross in one pass -> a single digest email ---');
resetState(); sent = [];
setFeed([
  { time: '10:30 AM', source: 'GetYourGuide', adults: 20 },
  { time: '5:00 PM', source: 'GetYourGuide', adults: 21 }
]);
capacityAlerts_();
check('two crossings in one pass send ONE digest (not two emails)', sent.length === 1, sent.map(s => s.subject));
check('the digest subject says how many tours are full', /2 tours full/.test((sent[0] || {}).subject || ''), sent[0] && sent[0].subject);
check('the digest body lists both tours', /10:30 AM/.test(sent[0].body) && /5:00 PM/.test(sent[0].body), sent[0] && sent[0].body);

console.log('--- the ">20 means 21" tuning point: threshold constant drives the trigger ---');
check('TOUR_FULL_THRESHOLD is 20 (reaches 20 -> full)', RNR.TOUR_FULL_THRESHOLD === 20, RNR.TOUR_FULL_THRESHOLD);

console.log('--- the alert body carries an auto-removal marker (ISO date + time) ---');
resetState(); sent = [];
setFeed([{ time: '10:30 AM', source: 'GetYourGuide', adults: 20 }]);
capacityAlerts_();
check('body carries "after the tour runs: <ISO date> <time>" so cleanup can read it',
  sent.length === 1 && new RegExp('after the tour runs:\\s*' + D + ' 10:30 AM').test(sent[0].body || ''),
  sent[0] && sent[0].body);

console.log('--- stale alerts are trashed after the tour time; upcoming ones are kept ---');
// Build fake Gmail threads the way GmailApp.search would hand them back.
const mkAlert = (dk, time) => {
  const body = ['This tour has reached the 20-adult cap.', 'Adults: 20',
    'This alert is removed automatically after the tour runs: ' + dk + ' ' + time + '.'].join('\n');
  return { trashed: false,
    getFirstMessageSubject: () => '🔴 English · x ' + time + ' · FULL 20 adults — close or add 2nd guide',
    getMessages: () => [makeFakeMsg_('cap', body)],
    moveToTrash: function () { this.trashed = true; } };
};
const past = mkAlert(dayKey(-1), '10:30 AM');
const future = mkAlert(dayKey(5), '10:30 AM');
// A real email that merely mentions "2nd guide" (no marker / not our subject) must be left alone.
const notOurs = { trashed: false,
  getFirstMessageSubject: () => 'Re: question about the 2nd guide rota',
  getMessages: () => [makeFakeMsg_('re', 'can we talk about the 2nd guide next week?')],
  moveToTrash: function () { this.trashed = true; } };
const _search = GmailApp.search;
GmailApp.search = () => [past, future, notOurs];
cleanupExpiredCapacityAlerts_();
GmailApp.search = _search;
check('a PAST-tour alert is trashed', past.trashed === true, past.trashed);
check('an UPCOMING-tour alert is kept', future.trashed === false, future.trashed);
check('a non-alert email that merely mentions "2nd guide" is NOT trashed', notOurs.trashed === false, notOurs.trashed);

console.log('--- a digest is trashed only after EVERY tour it covers has passed ---');
const mkDigest = (pairs) => {
  const lines = pairs.map(([dk, t]) => 'This alert is removed automatically after the tour runs: ' + dk + ' ' + t + '.');
  return { trashed: false,
    getFirstMessageSubject: () => '🔴 2 tours full (20+) — close or add a 2nd guide',
    getMessages: () => [makeFakeMsg_('digest', '2 tours full\n\n' + lines.join('\n'))],
    moveToTrash: function () { this.trashed = true; } };
};
const digestMixed = mkDigest([[dayKey(-1), '10:30 AM'], [dayKey(5), '5:00 PM']]);   // one past, one still upcoming
const digestAllPast = mkDigest([[dayKey(-2), '10:30 AM'], [dayKey(-1), '5:00 PM']]);
GmailApp.search = () => [digestMixed, digestAllPast];
cleanupExpiredCapacityAlerts_();
GmailApp.search = _search;
check('a digest with a still-upcoming tour is KEPT', digestMixed.trashed === false, digestMixed.trashed);
check('a digest whose tours have ALL passed is trashed', digestAllPast.trashed === true, digestAllPast.trashed);

console.log('--- PRIVATE alerts: ONE email PER private tour, concise subject (source/date/time) ---');
const resetPriv = () => PropertiesService.getScriptProperties().deleteProperty(RNR.PRIVATE_ALERT_PROP);
resetPriv(); sent = [];
setFeed([
  { time: '10:30 AM', lang: 'English', source: 'GetYourGuide', adults: 2, notes: 'Private tour', name: 'Smith', id: 'P-RR' },  // main GYG
  { time: '10:30 AM', lang: 'Spanish', source: 'GYG',          adults: 3, notes: 'Private', name: 'Lopez', id: 'P-SS' },       // 2nd GYG account
  { time: '5:00 PM',  lang: 'English', source: 'Viator',       adults: 2, notes: 'Private', name: 'Jones', id: 'P-V' },
  { time: '10:30 AM', lang: 'English', source: 'GetYourGuide', adults: 14, id: 'REG-1' }                                        // REGULAR -> not a private alert
]);
const pbk = readUpcomingPrivateBookings_();
check('private bookings = one entry PER tour (3 privates, regular excluded)', pbk.length === 3 && !pbk.some(b => b.people === 14), pbk.map(b => b.bookingId));

privateBookingAlerts_();
check('ONE email per private tour (3 privates -> 3 emails, no digest)', sent.length === 3, sent.map(s => s.subject));
const subs = sent.map(s => s.subject);
const dm = D.split('-')[2] + '/' + D.split('-')[1];   // DD/MM from the yyyy-MM-dd test date
check('subject format: "🔴 Close Private Tours, GYG_RR hired on DD/MM at HH:MM" (icon + no timezone)',
  subs.some(x => x === '🔴 Close Private Tours, GYG_RR hired on ' + dm + ' at 10:30') && !subs.some(x => /GMT|\d{4}.*00:00:00/.test(x)), subs);
check('source labels map: GYG_RR (main), GYG_SS (2nd account), Viator',
  subs.some(x => /GYG_RR/.test(x)) && subs.some(x => /GYG_SS/.test(x)) && subs.some(x => /Viator/.test(x)), subs);
check('time shown 24h (5:00 PM -> 17:00)', subs.some(x => /at 17:00$/.test(x)), subs);
check('every private email goes to the internal manager address', sent.every(s => s.to === RNR.INTERNAL_ALERT_TO), sent.map(s => s.to));
check('body is short (two content lines) + a removal marker', sent.every(s => (s.body.match(/after the tour runs:/g) || []).length === 1 && s.body.split('\n').filter(Boolean).length <= 3), sent.map(s => s.body));

sent = [];
privateBookingAlerts_();                                 // same feed, second pass
check('fires ONCE per booking: a re-run does NOT re-alert', sent.length === 0, sent.map(s => s.subject));

// Cancel P-V (remove it); the others stay alerted, nothing new.
setFeed([
  { time: '10:30 AM', lang: 'English', source: 'GetYourGuide', adults: 2, notes: 'Private tour', name: 'Smith', id: 'P-RR' },
  { time: '10:30 AM', lang: 'Spanish', source: 'GYG',          adults: 3, notes: 'Private', name: 'Lopez', id: 'P-SS' }
]);
sent = [];
privateBookingAlerts_();
check('removing a private booking does not re-alert the ones still booked', sent.length === 0, sent.map(s => s.subject));
// A NEW private (new id) re-alerts just that one.
setFeed([
  { time: '10:30 AM', lang: 'English', source: 'GetYourGuide', adults: 2, notes: 'Private tour', name: 'Smith', id: 'P-RR' },
  { time: '10:30 AM', lang: 'Spanish', source: 'GYG',          adults: 3, notes: 'Private', name: 'Lopez', id: 'P-SS' },
  { time: '1:00 PM',  lang: 'English', source: 'Viator',       adults: 2, notes: 'Private', name: 'NewPriv', id: 'P-NEW' }
]);
sent = [];
privateBookingAlerts_();
check('a NEW private booking re-alerts (single email, that booking only)',
  sent.length === 1 && /at 13:00$/.test(sent[0].subject) && /Viator/.test(sent[0].subject), sent.map(s => s.subject));

// PROD bug: Sheets returns the feed DATE cell as a Date OBJECT, which leaked a raw
// "Thu Oct 15 2026 00:00:00 GMT+0200" into the subject. feedYmd_ must normalise it.
resetPriv(); sent = [];
const _fl = feed.getLastRow(); if (_fl >= 2) feed.getRange(2, 1, _fl - 1, FH.length).clearContent();
feed.getRange(2, 1, 1, FH.length).setValues([[
  new Date(D + 'T00:00:00'), '10:30 AM', 'English', 'DateObj', '+1', 2, 0, 'Viator', 0, 'P-DOBJ', 'Private', '', '', '', '', 'booking']]);
privateBookingAlerts_();
check('a Date-OBJECT feed date renders DD/MM in the subject (no raw Date / timezone)',
  sent.length === 1 && /hired on \d{2}\/\d{2} at 10:30$/.test(sent[0].subject) && !/GMT|00:00:00/.test(sent[0].subject), sent.map(s => s.subject));

console.log('--- PRIVATE alert cleanup shares the capacity cleanup (one search) ---');
const mkPriv = (dk, t) => ({ trashed: false,
  getFirstMessageSubject: () => 'Close Private Tours, GYG_RR hired on 27/10 at ' + t,
  getMessages: () => [makeFakeMsg_('priv', 'This alert is removed automatically after the tour runs: ' + dk + ' ' + t + '.')],
  moveToTrash: function () { this.trashed = true; } });
const privPast = mkPriv(dayKey(-1), '10:30 AM'), privFuture = mkPriv(dayKey(5), '10:30 AM');
GmailApp.search = () => [privPast, privFuture];
cleanupExpiredCapacityAlerts_();
GmailApp.search = _search;
check('an expired PRIVATE alert is trashed by the shared cleanup', privPast.trashed === true, privPast.trashed);
check('an upcoming PRIVATE alert is kept', privFuture.trashed === false, privFuture.trashed);

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
