/* ===== Guide portal client logic: pending-assignment overlay =====
   Extracts the REAL functions from guide/index.html (no copy, no drift) and
   proves the "reverts to the old guide" bug is fixed: a stale server read that
   still shows the old guide can never undo the assignment you just made. */
const fs = require('fs');
const path = require('path');

let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

const html = fs.readFileSync(path.join(__dirname, '..', 'guide', 'index.html'), 'utf8');
const start = html.indexOf('const PENDING_TTL = 90000;');
const end = html.indexOf('function renderAll(r){');
if (start < 0 || end < 0) { console.error('Could not locate the pending-assignment block in guide/index.html'); process.exit(1); }
const block = html.slice(start, end);
// A minimal localStorage so the persist-across-reload logic is exercised.
const LS = {}; global.localStorage = { getItem: k => (k in LS ? LS[k] : null), setItem: (k, v) => { LS[k] = String(v); }, removeItem: k => { delete LS[k]; } };
// Run the extracted block in its own scope and expose the functions for testing.
const A = (new Function(block + '\n;return { shiftIdKey, setPendingAssign, applyPendingAssigns, loadPending, pending: () => PENDING_ASSIGN };'))();

console.log('--- Pending assignment SURVIVES a full page reload (persisted) ---');
A.setPendingAssign({ dateKey: '2026-08-10', time: '17:00', language: 'English', isPrivate: '' }, 'Carlos');
const reloaded = A.loadPending();   // what a freshly reloaded page would read back
check('the just-assigned guide is still remembered after a reload', reloaded['2026-08-10|17:00|English|R'] && reloaded['2026-08-10|17:00|English|R'].guide === 'Carlos', reloaded);
check('an expired pending is dropped on reload', (function(){ LS['rr_pending'] = JSON.stringify({ x: { guide: 'Old', until: Date.now() - 1 } }); return Object.keys(A.loadPending()).length === 0; })(), null);
delete A.pending()['2026-08-10|17:00|English|R']; LS['rr_pending'] = '{}';   // reset for the flip-flop tests below

console.log('--- The save payload carries children (so the ledger records them) ---');
check('the save payload sends children in each booking', /children:\s*b\.children/.test(html), null);

console.log('--- Freshness indicator + manager window controls exist ---');
check('freshness element is in the page', /id="fresh"/.test(html), null);
check('setFresh runs on load (loading/ok/stale states)', /setFresh\("loading"\)/.test(html) && /setFresh\("ok"/.test(html) && /setFresh\("stale"\)/.test(html), null);
check('loadTours passes the manager window days', /params\.days\s*=\s*MANAGER_DAYS/.test(html), null);
check('a "Load more" button appears when the server says hasMore', /HAS_MORE/.test(html) && /loadMoreBtn/.test(html), null);
check('the overlay holds through the settle window (no clear-on-match flip-back)', !/cur===p\.guide/.test(html), null);

console.log('--- Check-in only confirms "✓" once the ledger write succeeds ---');
check('saveTour reports whether the write succeeded', /return ok;/.test(html) && /ok=!!\(r&&r\.ok\)/.test(html), null);
console.log('--- Offline check-in queue: held on the phone, retried until it lands ---');
check('a check-in that fails to save is QUEUED (not lost)', /function ckEnqueue/.test(html) && /ckEnqueue\(buildTourData\(tid\)\)/.test(html), null);
check('the queue flushes on boot, poll, focus and reconnect', /addEventListener\("online", ?ckFlush\)/.test(html) && /ckFlush\(\);\s*\/\/ push any check-ins queued/.test(html), null);
check('a queued check-in shows a syncing state, not a revert', /✓ ⏳/.test(html), null);

console.log('--- Pending assignment survives a stale read (the Albert->Carlos bug) ---');
const shift = g => ({ dateKey: '2026-07-31', time: '17:00', language: 'English', isPrivate: false, assigned: g ? [g] : [], guide: g || '', status: g ? 'OK' : 'Not assigned' });
const info = { dateKey: '2026-07-31', time: '17:00', language: 'English', isPrivate: '' };

A.setPendingAssign(info, 'Carlos');                 // manager just picked Carlos
const stale = { allTours: [shift('Albert')], tours: [], schedule: [] };   // server replica still says Albert
A.applyPendingAssigns(stale);
check('a stale "Albert" read is overridden by the pending Carlos', stale.allTours[0].assigned.join() === 'Carlos', stale.allTours[0]);
check('the old guide is gone, not merged', stale.allTours[0].assigned.indexOf('Albert') === -1, stale.allTours[0].assigned);
check('pending is still held', Object.keys(A.pending()).length === 1, A.pending());

// A read that MATCHES must NOT clear the pending — a later lagging replica could
// still return the old value, and we must keep overriding it.
const matched = { allTours: [shift('Carlos')], tours: [], schedule: [] };
A.applyPendingAssigns(matched);
check('a matching read still shows Carlos', matched.allTours[0].assigned.join() === 'Carlos', matched.allTours[0]);
check('pending is STILL held after a match (guards a lagging replica)', Object.keys(A.pending()).length === 1, A.pending());
const laggard = { allTours: [shift('Albert')], tours: [], schedule: [] };  // replica flips back to old
A.applyPendingAssigns(laggard);
check('a LATER lagging "Albert" read is still overridden to Carlos (no flip-back)', laggard.allTours[0].assigned.join() === 'Carlos', laggard.allTours[0]);

// Once the settle window passes the pending expires and the real value flows.
Object.keys(A.pending()).forEach(k => { A.pending()[k].until = Date.now() - 1; });
const settled = { allTours: [shift('Carlos')], tours: [], schedule: [] };
A.applyPendingAssigns(settled);
check('after the settle window the pending clears', Object.keys(A.pending()).length === 0, A.pending());

console.log('--- A later stale poll cannot resurrect the old guide mid-flight ---');
A.setPendingAssign(info, 'Carlos');
[shift('Albert'), shift('Albert'), shift('Carlos')].forEach((s, i) => {
  const r = { allTours: [s], tours: [], schedule: [] };
  A.applyPendingAssigns(r);
  check('poll #' + (i + 1) + ' shows Carlos', r.allTours[0].assigned.join() === 'Carlos', r.allTours[0]);
});

console.log('--- Unassign is honoured the same way ---');
A.setPendingAssign(info, '');                        // cleared to Not assigned
const stillAssigned = { allTours: [shift('Albert')], tours: [], schedule: [] };
A.applyPendingAssigns(stillAssigned);
check('a stale read showing Albert is overridden to unassigned', stillAssigned.allTours[0].assigned.length === 0, stillAssigned.allTours[0]);

console.log('--- A different shift is untouched by an unrelated pending ---');
A.setPendingAssign(info, 'Carlos');
const other = { allTours: [{ dateKey: '2026-08-01', time: '11:00', language: 'English', isPrivate: false, assigned: ['Polina'], guide: 'Polina', status: 'OK' }], tours: [], schedule: [] };
A.applyPendingAssigns(other);
check('an unrelated shift keeps its own guide', other.allTours[0].assigned.join() === 'Polina', other.allTours[0]);

console.log('--- Manager stays on their tab across reloads ---');
check('the active tab is persisted (rr_tab)', /store\.tab/.test(html) && /rr_tab/.test(html), null);
check('switchTab records the choice', /function switchTab\(t\)\{[\s\S]*store\.tab = t;/.test(html), null);
check('renderAll restores the last tab (no jump back to My tours)', /function restoreTab/.test(html) && /restoreTab\(\);/.test(html), null);

console.log('--- Resilient loads: one miss must not show "Can\'t reach the server" ---');
check('a single failed load is gated (banner only after 2 in a row)', /LOAD_FAILS\s*<\s*2/.test(html), null);
check('a miss schedules a quiet auto-retry instead of an error', /RETRY_TIMER=setTimeout\(\(\)=>\{ RETRY_TIMER=null; loadTours\("retry"\)/.test(html), null);
check('a cached screen is kept on failure (error UI only when nothing cached)', /if\(!store\.cache\) showLoadError/.test(html), null);

console.log('--- A queued (offline) check-in survives a reload on screen ---');
check('a pending-check-in overlay exists and runs in renderAll', /function applyPendingCheckins\(r\)/.test(html) && /applyPendingCheckins\(r\);/.test(html), null);
check('the overlay reads the offline queue and marks bookings pending', /ckQueueLoad\(\)/.test(html) && /b\.pending=true/.test(html), null);
check('a pending check-in renders as "✓ ⏳", not "Check in"', /b\.pending\?'✓ ⏳'/.test(html), null);

console.log('--- Manager can undo a check-in from the portal ---');
check('a confirmed check-in is tappable ONLY for a manager (guides/pending stay locked)', /!\(MANAGER && !b\.pending\)/.test(html), null);
check('the undo binds on confirmed ✓ buttons for managers and calls action=uncheckin', /if\(MANAGER\) document\.querySelectorAll\("\.ckin\.done:not\(\[disabled\]\)"\)/.test(html) && /"uncheckin"/.test(html), null);
check('undo confirms before deleting and refreshes after', /Undo the check-in for/.test(html) && /loadTours\("undo"\)/.test(html), null);

console.log('--- Load analytics: reload-early vs real slowness is reported ---');
check('a client telemetry beacon exists', /function beacon\(ev, ms, nav\)/.test(html) && /"clientlog"/.test(html), null);
check('an early reload mid-load reports an abort', /pagehide[\s\S]*beacon\("abort"/.test(html), null);
check('the initial load is tagged so completed-load times are measurable', /loadTours\("initial"\)/.test(html), null);

console.log('--- Resilient writes: one quiet retry, and reconcile after any move ---');
check('api() retries once on a non-auth failure', /async function api\(action, params\)\{[\s\S]*catch\(err\)\{[\s\S]*if\(err && err\.kind==="auth"\) throw err;[\s\S]*return await apiCall\(action, params\);/.test(html), null);
check('a move reconciles with loadTours whether it succeeds OR fails', /\.mvbox[\s\S]*?loadTours\(\);\s+\/\/ reconcile either way/.test(html), null);
check('one Move applies BOTH language and time in a single call', /api\("move",\{[\s\S]*?toLanguage:toLanguage,[\s\S]*?toTime: timeChange/.test(html), null);
check('the Move button shows only after a dropdown is changed', /go\.hidden = \(langSel\.value===info\.fromLanguage && timeSel\.value===info\.fromTime && \(!dateSel \|\| dateSel\.value===info\.fromDate\)\)/.test(html), null);
check('a day move is offered (date input) and passed to the move call', /class="mvsel mv-date"/.test(html) && /toDate: dateChange\? toDate : ""/.test(html), null);
check('guides (not managers) get a green tour with guests, red when empty', /const peopleCls = MANAGER \? '' : \(Number\(t\.bookedGuests\)>0 \? ' has-people' : ' no-people'\)/.test(html) && /\.tour\.has-people\{/.test(html) && /\.tour\.no-people\{/.test(html), null);

console.log('--- Auto-refresh never interrupts a mid-action user (no wiped forms) ---');
check('the poll pauses while the "Open a schedule" form is open', /details\.openform\[open\]"\)\) return false;/.test(html), null);
check('choosing in a SELECT also holds the poll', /INPUT\|TEXTAREA\|SELECT/.test(html), null);
check('the create-tour form fields count as editing', /\.of-date,\.of-time,\.of-lang,\.of-guide/.test(html), null);
check('a version-check hard-reload will not fire mid-action', /openform\[open\]"\) \|\| Date\.now\(\)-LAST_EDIT<8000\) return;/.test(html), null);
check('a move holds the poll (MUTATING) so it cannot race the write', /go\.textContent="Moving…"; MUTATING\+\+/.test(html), null);

console.log('--- Typed tour time: bare afternoon hours read as PM (4:30 = 16:30) ---');
const ptStart = html.indexOf('function parseTypedTime(raw){');
const parseTypedTime = (new Function(html.slice(ptStart, html.indexOf('\n}', ptStart) + 2) + '\n;return parseTypedTime;'))();
check('"4:30" -> 16:30 (PM), labelled 4:30 PM', (function(){ const r=parseTypedTime('4:30'); return r && r.t24==='16:30' && r.label==='4:30 PM'; })(), parseTypedTime('4:30'));
check('"16:30" -> 16:30 PM (explicit 24h respected)', (function(){ const r=parseTypedTime('16:30'); return r && r.t24==='16:30' && r.label==='4:30 PM'; })(), parseTypedTime('16:30'));
check('"10:00" -> 10:00 AM (morning stays AM)', (function(){ const r=parseTypedTime('10:00'); return r && r.t24==='10:00' && r.label==='10:00 AM'; })(), parseTypedTime('10:00'));
check('"11:00" -> 11:00 AM', (function(){ const r=parseTypedTime('11:00'); return r && r.t24==='11:00' && r.label==='11:00 AM'; })(), parseTypedTime('11:00'));
check('"12:30" -> 12:30 PM (noon stays PM)', (function(){ const r=parseTypedTime('12:30'); return r && r.t24==='12:30' && r.label==='12:30 PM'; })(), parseTypedTime('12:30'));
check('"4:30 PM" explicit is honoured', (function(){ const r=parseTypedTime('4:30 PM'); return r && r.t24==='16:30'; })(), parseTypedTime('4:30 PM'));
check('"9 am" and "5pm" shorthand parse', parseTypedTime('9 am').t24==='9:00' && parseTypedTime('5pm').t24==='17:00', [parseTypedTime('9 am'), parseTypedTime('5pm')]);
check('garbage input is rejected (null), never a wrong time', parseTypedTime('later')===null && parseTypedTime('25:99')===null, [parseTypedTime('later'), parseTypedTime('25:99')]);

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
