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

console.log('--- the ledger/feed pattern (key includes feedCacheVersion_) refreshes on a check-in ---');
let lc = 0; const lfn = () => { lc++; return { n: lc }; };
const lkey = () => 'led:' + feedCacheVersion_() + ':carlos';
cachedRead_(lkey(), 60, lfn); cachedRead_(lkey(), 60, lfn);
check('repeat poll with no change hits the cache (ledger read skipped)', lc === 1, lc);
bumpFeedCacheVersion_();                                  // a check-in / undo does this
cachedRead_(lkey(), 60, lfn);
check('a check-in (bumpFeedCacheVersion_) refreshes the ledger read', lc === 2, lc);

console.log('--- an oversized value is never cached (stays a live read) ---');
let big = 0; const bigfn = () => { big++; return { s: 'x'.repeat(96000) }; };
cachedRead_('big', 60, bigfn); cachedRead_('big', 60, bigfn);
check('a >95KB value reads live every time (never cached)', big === 2, big);

console.log('--- a DIFFERENT guide set is a different key (reassignment safety) ---');
let g2 = 0; const g2fn = () => { g2++; return { n: g2 }; };
cachedRead_('led:' + feedCacheVersion_() + ':setA', 60, g2fn);
cachedRead_('led:' + feedCacheVersion_() + ':setB', 60, g2fn);
check('a changed guide set misses (reads fresh, not another set\'s check-ins)', g2 === 2, g2);

console.log('--- the ASSEMBLE cache (asm:<feedVer>:<horizon>) skips the rebuild on a warm poll ---');
// The assemble phase (buildScheduleFromFeed_ + weekly shifts + defaults + sort) was
// pure CPU re-run on EVERY poll (1.4-13s), the cause of the multi-minute timeouts.
// It is now cachedRead_ under the feed version + offer horizon. It MUST invalidate on
// a check-in (feed) AND on an assign/move/close (global, added inside cachedRead_) so
// it can never serve a stale assignment. Maxim: cannot miss an assignment.
__mock.PROPS['PORTAL_CACHE_VER'] = '0';
__mock.PROPS['PORTAL_FEED_VER'] = '0';
let ab = 0; const abfn = () => { ab++; return { shifts: ab }; };
const akey = (h) => 'asm:' + feedCacheVersion_() + ':' + h;
cachedRead_(akey(90), 60, abfn); cachedRead_(akey(90), 60, abfn);
check('a warm poll (same feed+horizon) skips the rebuild', ab === 1, ab);
bumpCacheVersion_();                                      // an assign / move / close
cachedRead_(akey(90), 60, abfn);
check('an assign/move/close rebuilds the schedule (no stale assignment)', ab === 2, ab);
bumpFeedCacheVersion_();                                  // a check-in
cachedRead_(akey(90), 60, abfn);
check('a check-in rebuilds the schedule (feed version in the key)', ab === 3, ab);
cachedRead_(akey(30), 60, abfn);                          // guide window vs manager window
check('a different offer horizon is a different key (guide vs manager window)', ab === 4, ab);

__mock.removeRealCache();
console.log('=================================');
console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
