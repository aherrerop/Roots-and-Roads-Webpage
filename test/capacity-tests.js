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
  { time: '10:30 AM', source: 'GetYourGuide', adults: 12, guide: 'Carlos' },
  { time: '10:30 AM', source: 'Viator', adults: 6, children: 2 },       // 8 people -> slot totals 20
  { time: '5:00 PM', source: 'GetYourGuide', adults: 5 },               // a quiet tour
  { time: '10:30 AM', source: 'GetYourGuide', adults: 25, notes: 'Private' }, // private -> excluded
  { time: '10:30 AM', source: 'GYG-SF', adults: 4 }                     // SF -> its OWN tour
]);
const groups = readUpcomingTourGroups_();
const k3h = D + '|10:30 AM|english|3h';
const kSf = D + '|10:30 AM|english|SF';
const k17 = D + '|5:00 PM|english|3h';
check('3h 10:30 slot sums adults + children across bookings = 20', groups[k3h] && groups[k3h].people === 20, groups[k3h]);
check('the SF booking is a SEPARATE tour (not merged into the 3h count)', groups[kSf] && groups[kSf].people === 4, groups[kSf]);
check('the private 25-person booking is EXCLUDED from any group', !Object.values(groups).some(g => g.people === 25 || g.people === 45), Object.keys(groups));
check('per-source breakdown is kept (GetYourGuide 12, Viator 8)', groups[k3h] && groups[k3h].sources['GetYourGuide'] === 12 && groups[k3h].sources['Viator'] === 8, groups[k3h] && groups[k3h].sources);
check('the quiet 17:00 tour is tracked but well under the cap', groups[k17] && groups[k17].people === 5, groups[k17]);

console.log('--- a tour reaching the cap sends ONE alert, all info in the subject ---');
resetState(); sent = [];
capacityAlerts_();
check('exactly one alert email is sent', sent.length === 1, sent.map(s => s.subject));
const sub = (sent[0] || {}).subject || '';
check('subject flags it FULL with the people count', /FULL\s+20p/.test(sub), sub);
check('subject names the day, time and language', /10:30 AM/.test(sub) && /English/.test(sub) && new RegExp(Utilities.formatDate(new Date(D + 'T12:00:00'), 'Europe/Madrid', 'EEE d MMM')).test(sub), sub);
check('subject states the action (close or add 2nd guide)', /close or add 2nd guide/i.test(sub), sub);
const body = (sent[0] || {}).body || '';
check('body carries the adults + children split', /20\s+\(18 adults \+ 2 children\)/.test(body), body);
check('body carries the per-source breakdown', /GetYourGuide: 12/.test(body) && /Viator: 8/.test(body), body);
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

console.log('--- children alone can push a tour to the cap ---');
resetState(); sent = [];
setFeed([{ time: '11:00 AM', source: 'GetYourGuide', adults: 19, children: 1 }]);   // 20 people
capacityAlerts_();
check('adults 19 + children 1 = 20 people crosses the cap', sent.length === 1 && /FULL\s+20p/.test(sent[0].subject), sent.map(s => s.subject));

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

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
