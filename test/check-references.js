/**
 * STATIC CHECK — every internal function call must resolve.
 *
 * Catches the failure mode that shipped a ReferenceError to production:
 * a function deleted (or renamed) while call sites remained. Apps Script
 * only fails at RUNTIME, so this must be checked before deploying.
 *
 * Each Apps Script PROJECT shares one global scope, so files are grouped
 * per project rather than checked individually.
 *
 * Usage: node test/check-references.js
 */
const fs = require('fs');
const path = require('path');
const AS = path.join(__dirname, '..', 'apps-script');

const PROJECTS = {
  'BookingSheet': ['booking/bookingList_v2.gs', 'booking/websiteAvailabilityUpdate.gs'],
  'Control': ['control/guidePortal.gs', 'control/assignShifts.gs', 'control/mobileControls.gs']
};

// Built-ins and Apps Script services that legitimately look like calls.
const IGNORE = new Set(['function_', 'if_', 'for_', 'while_', 'catch_', 'switch_', 'return_']);

let failed = 0;

for (const [project, files] of Object.entries(PROJECTS)) {
  const src = files.map(f => fs.readFileSync(path.join(AS, f), 'utf8')).join('\n');

  const defined = new Set();
  let m;
  const defRe = /^function ([A-Za-z0-9_]+)\s*\(/gm;
  while ((m = defRe.exec(src))) defined.add(m[1]);

  // Project-internal helpers are the trailing-underscore convention.
  const called = new Map();
  const callRe = /\b([A-Za-z_][A-Za-z0-9_]*_)\s*\(/g;
  while ((m = callRe.exec(src))) {
    if (!called.has(m[1])) {
      const line = src.slice(0, m.index).split('\n').length;
      called.set(m[1], line);
    }
  }

  const missing = [...called.keys()].filter(n => !defined.has(n) && !IGNORE.has(n));

  // Duplicate definitions inside one project break silently (last wins).
  const counts = {};
  const defRe2 = /^function ([A-Za-z0-9_]+)\s*\(/gm;
  while ((m = defRe2.exec(src))) counts[m[1]] = (counts[m[1]] || 0) + 1;
  const dupes = Object.keys(counts).filter(k => counts[k] > 1);

  if (missing.length || dupes.length) {
    failed = 1;
    console.log('FAIL  ' + project);
    missing.forEach(n => console.log('        undefined call: ' + n + '()  (first used line ~' + called.get(n) + ')'));
    dupes.forEach(n => console.log('        defined twice : ' + n + '()'));
  } else {
    console.log('PASS  ' + project + '  (' + defined.size + ' functions, all calls resolve, no duplicates)');
  }
}

console.log(failed ? 'REFERENCE CHECK FAILED — do not deploy.' : 'REFERENCE CHECK PASSED.');
process.exit(failed);
