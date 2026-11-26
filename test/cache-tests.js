/* Cache behaviour — cachedRead_ hit/miss + version-bump freshness. Uses the REAL
   in-memory cache (other suites use the no-op) so the actual caching runs. This
   is what the speed pass relies on: a warm poll skips reads, and any write bumps
   a version so the very next read is fresh. */
let pass = 0, fail = 0;
const check = (l, c, g) => { if (c) { pass++; console.log('PASS  ' + l); } else { fail++; console.log('FAIL  ' + l + '  (got: ' + JSON.stringify(g) + ')'); } };

__mock.installRealCache();
__mock.PROPS['PORTAL_CACHE_VER'] = '0';
__mock.PROPS['PORTAL_FEED_VER'] = '0';

console.log('--- cachedRead_ actually caches, and a value survives a repeat read ---');
let calls = 0; const fn = () => { calls++; return { v: calls }; };
const a = cachedRead_('x', 60, fn);
const b = cachedRead_('x', 60, fn);
check('fn runs ONCE across two reads (cache hit)', calls === 1 && a.v === 1 && b.v === 1, [calls, a, b]);

console.log('--- bumping the global version invalidates (assign/move/note/close path) ---');
bumpCacheVersion_();
const c = cachedRead_('x', 60, fn);
check('bumpCacheVersion_ -> fn runs again, fresh value', calls === 2 && c.v === 2, [calls, c]);

console.log('--- the ledger backstop (cachedReadStable_) survives BOTH an assign burst and a check-in burst ---');
// The feed is the PRIMARY check-in source; the ledger is only a backstop, and a
// displayed booking always has a feed row. The backstop is keyed on the guide set + TTL
// alone (cachedReadStable_), independent of BOTH versions: a future assign/move can't
// change a PAST/today check-in, and a check-in already shows from the feed, so NEITHER an
// assign burst nor a check-in burst must re-open the ledger + re-read every guide tab (3-9s
// of I/O) on the next poll and starve guide loads. It refreshes on the TTL + hourly reconcile.
let lc = 0; const lfn = () => { lc++; return { n: lc }; };
const lkey = () => 'led:carlos';
cachedReadStable_(lkey(), 60, lfn); cachedReadStable_(lkey(), 60, lfn);
check('repeat poll with no change hits the cache (ledger read skipped)', lc === 1, lc);
bumpFeedCacheVersion_();                                  // a check-in / undo
cachedReadStable_(lkey(), 60, lfn);
check('a check-in does NOT re-open the ledger (feed carries the check-in)', lc === 1, lc);
bumpCacheVersion_();                                     // an assign/move burst
cachedReadStable_(lkey(), 60, lfn);
check('an assign/move does NOT re-open the ledger (past check-ins are unaffected)', lc === 1, lc);

console.log('--- an oversized value is never cached (stays a live read) ---');
let big = 0; const bigfn = () => { big++; return { s: 'x'.repeat(96000) }; };
cachedRead_('big', 60, bigfn); cachedRead_('big', 60, bigfn);
check('a >95KB value reads live every time (never cached)', big === 2, big);

console.log('--- a DIFFERENT guide set is a different key (reassignment safety) ---');
let g2 = 0; const g2fn = () => { g2++; return { n: g2 }; };
cachedReadStable_('led:setA', 60, g2fn);
cachedReadStable_('led:setB', 60, g2fn);
check('a changed guide set misses (reads fresh, not another set\'s check-ins)', g2 === 2, g2);

console.log('--- the ASSEMBLE cache (asm:<horizon>) skips the rebuild on a warm poll AND across check-ins ---');
// The assemble phase (buildScheduleFromFeed_ + weekly shifts + defaults + sort) was
// pure CPU re-run on EVERY poll (1.4-13s), the cause of the multi-minute timeouts.
// It caches the shift STRUCTURE. A check-in changes only M/N counts (applied
// per-request from the fresh feed), NOT the structure, so it is NOT keyed on the
// feed version — a check-in must not rebuild the whole schedule. An assign/move/
// close bumps the GLOBAL version (added inside cachedRead_) and DOES rebuild it, so
// it can never serve a stale assignment. Maxim: cannot miss an assignment.
__mock.PROPS['PORTAL_CACHE_VER'] = '0';
__mock.PROPS['PORTAL_FEED_VER'] = '0';
let ab = 0; const abfn = () => { ab++; return { shifts: ab }; };
const akey = (h) => 'asm:' + h;
cachedRead_(akey(90), 60, abfn); cachedRead_(akey(90), 60, abfn);
check('a warm poll (same horizon) skips the rebuild', ab === 1, ab);
bumpFeedCacheVersion_();                                  // a check-in
cachedRead_(akey(90), 60, abfn);
check('a check-in does NOT rebuild the schedule (structure unchanged)', ab === 1, ab);
bumpCacheVersion_();                                      // an assign / move / close
cachedRead_(akey(90), 60, abfn);
check('an assign/move/close rebuilds the schedule (no stale assignment)', ab === 2, ab);
cachedRead_(akey(30), 60, abfn);                          // guide window vs manager window
check('a different offer horizon is a different key (guide vs manager window)', ab === 3, ab);

console.log('--- CONFIG reads (guides/closed) survive an assign burst, refresh on a config change ---');
// Guides + Closed_Shifts don't change on an assign/move/check-in, so they are
// keyed on the CONFIG version, not the global one. An assign burst must NOT throw
// them away (that forced every watching guide into a cold load); only a close/
// reopen or a password change (bumpConfigVersion_) refreshes them.
__mock.PROPS['PORTAL_CACHE_VER'] = '0';
__mock.PROPS['PORTAL_CONFIG_VER'] = '0';
let cf = 0; const cffn = () => { cf++; return { g: cf }; };
cachedReadConfig_('guides', 300, cffn); cachedReadConfig_('guides', 300, cffn);
check('a repeat poll hits the config cache (guides read skipped)', cf === 1, cf);
bumpCacheVersion_(); bumpCacheVersion_();                 // an assign burst
cachedReadConfig_('guides', 300, cffn);
check('an assign does NOT invalidate a config read (stays warm through the burst)', cf === 1, cf);
bumpConfigVersion_();                                     // a close/reopen or password change
cachedReadConfig_('guides', 300, cffn);
check('a config change (bumpConfigVersion_) refreshes the config read', cf === 2, cf);

console.log('--- Adaptive backpressure: contention level, EWMA escalation + recovery ---');
check('fast loads -> normal / 30s poll', contentionLevelFor_(1000).level==='normal' && contentionLevelFor_(1000).pollSec===30, contentionLevelFor_(1000));
check('at busy threshold -> busy / 60s poll', contentionLevelFor_(4000).level==='busy' && contentionLevelFor_(4000).pollSec===60, contentionLevelFor_(4000));
check('at overloaded threshold -> overloaded / 120s poll', contentionLevelFor_(9000).level==='overloaded' && contentionLevelFor_(9000).pollSec===120, contentionLevelFor_(9000));
// A sustained run of slow loads (a queue forming) escalates; then fast loads recover.
let lvl; recordLoadSignal_(1000);
for (let i = 0; i < 5; i++) lvl = recordLoadSignal_(20000);
check('sustained slow loads escalate to overloaded (poll backs off automatically)', lvl.level === 'overloaded' && lvl.pollSec === 120, lvl);
check('readContention_ reflects the live level', readContention_().level === 'overloaded', readContention_());
for (let i = 0; i < 8; i++) lvl = recordLoadSignal_(500);
check('when load falls the level returns to normal (auto-recovery)', lvl.level === 'normal' && lvl.pollSec === 30, lvl);

console.log('--- Snapshot self-correction: a just-made assignment shows at once (no cold rebuild, no 5-min stale window) ---');
// reapplySnapshotMutations_ re-derives each shift's guide from the FRESH feed over the
// snapshot's structure, so a manager's assign/move/clear shows on the very next load.
__RRX = {};
const SK = (d, m, l) => d + '|' + m + '|' + String(l).toLowerCase();
const mkShift = (over) => ({ dateKey: '2027-06-15', minutes: 930, time: '15:30', timeLabel: '3:30 PM',
  language: 'Italian', day: 'Tuesday', dateText: 'Jun 15', private: false, sfExt: false, privIndex: 1,
  assigned: over ? [over] : [], status: over ? 'OK' : 'Not assigned' });

// 1. The snapshot still shows the OLD guide; the fresh feed now names a NEW guide -> reapply flips it.
let sched = [mkShift('OldGuide')];
let feed = { [SK('2027-06-15', 930, 'Italian')]: [{ source: 'Website', note: '', feedGuide: 'Francesca', rowType: 'booking', bookingId: 'B1' }] };
let ok = reapplySnapshotMutations_(sched, feed, '2027-06-15', '2027-06-18');
check('a stale snapshot assignment is replaced by the fresh feed guide', ok === true && sched[0].assigned.join() === 'Francesca', sched[0].assigned);

// 2. A BARE assigned slot carries its guide on a 0-person 'shift' placeholder row -> still picked up.
sched = [mkShift('')];
feed = { [SK('2027-06-15', 930, 'Italian')]: [{ source: '', note: '', feedGuide: 'Carlos', rowType: 'shift' }] };
reapplySnapshotMutations_(sched, feed, '2027-06-15', '2027-06-18');
check('a bare assignment (placeholder row) is reflected, not dropped', sched[0].assigned.join() === 'Carlos', sched[0].assigned);

// 3. Unassign: the feed no longer names a guide -> the shift clears.
sched = [mkShift('OldGuide')];
feed = { [SK('2027-06-15', 930, 'Italian')]: [{ source: 'Website', note: '', feedGuide: '', rowType: 'booking', bookingId: 'B1' }] };
reapplySnapshotMutations_(sched, feed, '2027-06-15', '2027-06-18');
check('an unassigned shift (no feed guide) clears to Not assigned', sched[0].assigned.length === 0 && sched[0].status === 'Not assigned', sched[0]);

// 4. Stage 2b: a NEW BOOKING on a slot the snapshot lacks is BUILT IN (reapply returns true and
//    the slot appears) instead of forcing a cold rebuild. The booking is never hidden.
sched = [mkShift('Francesca')];
feed = {
  [SK('2027-06-15', 930, 'Italian')]: [{ source: 'Website', note: '', feedGuide: 'Francesca', rowType: 'booking', bookingId: 'B1' }],
  [SK('2027-06-15', 660, 'English')]: [{ source: 'Website', note: '', feedGuide: '', rowType: 'booking', bookingId: 'B2' }]
};
let built = reapplySnapshotMutations_(sched, feed, '2027-06-15', '2027-06-18');
let newSlot = sched.find(s => s.language === 'English' && s.minutes === 660);
check('Stage 2b: a new booking slot the snapshot lacks is BUILT IN (returns true, slot appears)', built === true && !!newSlot && newSlot.status === 'Not assigned', [built, newSlot && newSlot.status]);

// 4b. KILL-SWITCH: PLANA_BUILDIN=OFF reverts to the safe cold-build bail (returns false) on a new
//     booking slot — the instant, no-deploy rollback. The caller then cold-builds (never misses it).
__mock.PROPS['PLANA_BUILDIN'] = 'OFF';
sched = [mkShift('Francesca')];
feed = {
  [SK('2027-06-15', 930, 'Italian')]: [{ source: 'Website', note: '', feedGuide: 'Francesca', rowType: 'booking', bookingId: 'B1' }],
  [SK('2027-06-15', 660, 'English')]: [{ source: 'Website', note: '', feedGuide: '', rowType: 'booking', bookingId: 'B2' }]
};
check('kill-switch OFF: a new booking slot forces the cold build again (returns false)', reapplySnapshotMutations_(sched, feed, '2027-06-15', '2027-06-18') === false, null);
delete __mock.PROPS['PLANA_BUILDIN'];

// 5. ...but a bare placeholder on a missing slot is NOT a booking, so it does not force a cold build.
sched = [mkShift('Francesca')];
feed = {
  [SK('2027-06-15', 930, 'Italian')]: [{ source: 'Website', note: '', feedGuide: 'Francesca', rowType: 'booking', bookingId: 'B1' }],
  [SK('2027-06-15', 660, 'English')]: [{ source: '', note: '', feedGuide: 'Hans', rowType: 'shift' }]
};
check('a bare placeholder on a missing slot does NOT force a cold build (true)', reapplySnapshotMutations_(sched, feed, '2027-06-15', '2027-06-18') === true, null);

console.log('--- Static read path (Stage 1): the PUBLIC read-model must be PII-FREE ---');
let pmSched = [mkShift('Carlos')];
let pmFeed = { [SK('2027-06-15', 930, 'Italian')]: [
  { source: 'Website', note: '', feedGuide: 'Carlos', rowType: 'booking', bookingId: 'B9',
    name: 'Jane Secret', phone: '+34600123456', guests: 3, children: 1, feedCheckedIn: 2 }
] };
let pmModel = portalPublicReadModel_(pmSched, pmFeed, '2027-06-14', 3);
let pmJson = JSON.stringify(pmModel);
check('read-model: the tour appears with correct head-counts (counts only, no who)', pmModel.tours.length === 1 && pmModel.tours[0].bookedGuests === 3 && pmModel.tours[0].bookedChildren === 1 && pmModel.tours[0].checkedGuests === 2 && pmModel.tours[0].booked === 1, pmModel.tours[0]);
check('read-model: NO guest NAME leaks to the public blob', pmJson.indexOf('Jane Secret') === -1, pmJson.slice(0, 140));
check('read-model: NO guest PHONE leaks to the public blob', pmJson.indexOf('600123456') === -1, null);
check('read-model: no name/phone/note/bookings KEYS on any tour', pmModel.tours.every(t => !('name' in t) && !('phone' in t) && !('bookings' in t) && !('note' in t)), Object.keys(pmModel.tours[0]));
check('read-model: still carries the assignment + flags needed to render', pmModel.tours[0].guide === 'Carlos' && pmModel.tours[0].status === 'OK' && pmModel.tours[0].isPrivate === false && pmModel.tours[0].sfExt === false, pmModel.tours[0]);
check('read-model: carries a guidesByLanguage roster object (guide names only, not guest PII)', pmModel.guidesByLanguage && typeof pmModel.guidesByLanguage === 'object' && JSON.stringify(pmModel.guidesByLanguage).indexOf('Jane Secret') === -1, Object.keys(pmModel.guidesByLanguage || {}));

console.log('--- Fail-closed staging guard: a non-prod copy must override every sheet id or REFUSE ---');
delete __mock.PROPS['ENV']; delete __mock.PROPS['TESTSHEETKEY'];
check('production (ENV unset): extSheetId_ falls back to the live default', extSheetId_('TESTSHEETKEY', 'LIVE_DEFAULT') === 'LIVE_DEFAULT', null);
__mock.PROPS['TESTSHEETKEY'] = 'COPY_ID';
check('an override Script Property always wins', extSheetId_('TESTSHEETKEY', 'LIVE_DEFAULT') === 'COPY_ID', null);
delete __mock.PROPS['TESTSHEETKEY']; __mock.PROPS['ENV'] = 'staging';
check('isCopyEnv_ true when ENV=staging', isCopyEnv_() === true, null);
let guardThrew = false; try { extSheetId_('TESTSHEETKEY', 'LIVE_DEFAULT'); } catch (e) { guardThrew = true; }
check('ENV=staging + MISSING override THROWS (never silently uses the live sheet)', guardThrew === true, null);
__mock.PROPS['TESTSHEETKEY'] = 'COPY_ID3';
check('ENV=staging + override present -> returns the copy id (no throw)', extSheetId_('TESTSHEETKEY', 'LIVE_DEFAULT') === 'COPY_ID3', null);
delete __mock.PROPS['ENV']; delete __mock.PROPS['TESTSHEETKEY'];
check('production again after cleanup (ENV unset): isCopyEnv_ false', isCopyEnv_() === false, null);

__mock.removeRealCache();
console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
