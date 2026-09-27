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

console.log('--- the ledger backstop is keyed on the guide set, NOT feedCacheVersion_ ---');
// The feed is the PRIMARY check-in source; the ledger is only a backstop, and a
// displayed booking always has a feed row. So a check-in must NOT re-open the
// ledger file + re-read every guide tab on the next poll (the moment a guide is
// tapping check-ins). The ledger stays keyed on the guide set + the GLOBAL version
// (an assign/move, added by cachedRead_) and refreshes on the TTL.
let lc = 0; const lfn = () => { lc++; return { n: lc }; };
const lkey = () => 'led:carlos';
cachedRead_(lkey(), 60, lfn); cachedRead_(lkey(), 60, lfn);
check('repeat poll with no change hits the cache (ledger read skipped)', lc === 1, lc);
bumpFeedCacheVersion_();                                  // a check-in / undo
cachedRead_(lkey(), 60, lfn);
check('a check-in does NOT re-open the ledger (feed carries the check-in)', lc === 1, lc);
bumpCacheVersion_();                                     // an assign/move DOES refresh it
cachedRead_(lkey(), 60, lfn);
check('an assign/move refreshes the ledger backstop (global version)', lc === 2, lc);

console.log('--- an oversized value is never cached (stays a live read) ---');
let big = 0; const bigfn = () => { big++; return { s: 'x'.repeat(96000) }; };
cachedRead_('big', 60, bigfn); cachedRead_('big', 60, bigfn);
check('a >95KB value reads live every time (never cached)', big === 2, big);

console.log('--- a DIFFERENT guide set is a different key (reassignment safety) ---');
let g2 = 0; const g2fn = () => { g2++; return { n: g2 }; };
cachedRead_('led:setA', 60, g2fn);
cachedRead_('led:setB', 60, g2fn);
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

__mock.removeRealCache();
console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
