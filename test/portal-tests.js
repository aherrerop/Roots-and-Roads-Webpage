/* ===== Guide portal client logic: pending-assignment overlay =====
   Extracts the REAL functions from guide/index.html (no copy, no drift) and
   proves the "reverts to the old guide" bug is fixed: a stale server read that
   still shows the old guide can never undo the assignment you just made. */
const fs = require('fs');
const path = require('path');

let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

const html = fs.readFileSync(path.join(__dirname, '..', 'guide', 'index.html'), 'utf8');
const start = html.indexOf('let PENDING_ASSIGN');
const end = html.indexOf('function renderAll(r){');
if (start < 0 || end < 0) { console.error('Could not locate the pending-assignment block in guide/index.html'); process.exit(1); }
const block = html.slice(start, end);
// Run the extracted block in its own scope and expose the functions for testing.
const A = (new Function(block + '\n;return { shiftIdKey, setPendingAssign, applyPendingAssigns, pending: () => PENDING_ASSIGN };'))();

console.log('--- The save payload carries children (so the ledger records them) ---');
const saveFn = html.slice(html.indexOf('async function saveTour('), html.indexOf('/* ================== ALL GUIDES'));
check('saveTour sends children in each booking', /children:\s*b\.children/.test(saveFn), saveFn.slice(0, 0));

console.log('--- Freshness indicator + manager window controls exist ---');
check('freshness element is in the page', /id="fresh"/.test(html), null);
check('setFresh runs on load (loading/ok/stale states)', /setFresh\("loading"\)/.test(html) && /setFresh\("ok"/.test(html) && /setFresh\("stale"\)/.test(html), null);
check('loadTours passes the manager window days', /params\.days\s*=\s*MANAGER_DAYS/.test(html), null);
check('a "Load more" button appears when the server says hasMore', /HAS_MORE/.test(html) && /loadMoreBtn/.test(html), null);

console.log('--- Pending assignment survives a stale read (the Albert->Carlos bug) ---');
const shift = g => ({ dateKey: '2026-07-31', time: '17:00', language: 'English', isPrivate: false, assigned: g ? [g] : [], guide: g || '', status: g ? 'OK' : 'Not assigned' });
const info = { dateKey: '2026-07-31', time: '17:00', language: 'English', isPrivate: '' };

A.setPendingAssign(info, 'Carlos');                 // manager just picked Carlos
const stale = { allTours: [shift('Albert')], tours: [], schedule: [] };   // server replica still says Albert
A.applyPendingAssigns(stale);
check('a stale "Albert" read is overridden by the pending Carlos', stale.allTours[0].assigned.join() === 'Carlos', stale.allTours[0]);
check('the old guide is gone, not merged', stale.allTours[0].assigned.indexOf('Albert') === -1, stale.allTours[0].assigned);
check('pending is still held until the server confirms', Object.keys(A.pending()).length === 1, A.pending());

const caughtUp = { allTours: [shift('Carlos')], tours: [], schedule: [] };   // server caught up
A.applyPendingAssigns(caughtUp);
check('once the server shows Carlos, it stays Carlos', caughtUp.allTours[0].assigned.join() === 'Carlos', caughtUp.allTours[0]);
check('pending is cleared after confirmation', Object.keys(A.pending()).length === 0, A.pending());

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

console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
