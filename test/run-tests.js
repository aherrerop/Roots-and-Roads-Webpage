/**
 * Cross-platform test runner (no bash/WSL needed — runs on plain Windows too).
 *
 * Mirrors run-tests.sh: the reference check, then each Sheets-logic suite run
 * in a fresh Node process over one shared global scope (mocks + the Control
 * project sources + the suite), exactly like Apps Script loads a project.
 *
 * Usage: node test/run-tests.js
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const testDir = __dirname;
const AS = path.join(testDir, '..', 'apps-script');
const read = f => fs.readFileSync(f, 'utf8');
const node = process.execPath;

// 1. Static reference check — hard gate (undefined calls / duplicate defs).
try {
  execFileSync(node, [path.join(testDir, 'check-references.js')], { stdio: 'inherit' });
} catch (e) {
  process.exit(1);
}

// 2. Sheets-logic suites.
const control = [
  path.join(AS, 'control', 'assignShifts.gs'),
  path.join(AS, 'control', 'guidePortal.gs')
];
const suites = ['tests.js', 'tests2.js', 'tests3.js', 'tests4.js', 'tests5.js'];

let failed = 0;
for (const suite of suites) {
  const bundle = [path.join(testDir, 'mock.js'), ...control, path.join(testDir, suite)]
    .map(read).join('\n');
  const tmp = path.join(os.tmpdir(), 'rr_' + suite.replace(/\W/g, '_') + '.js');
  fs.writeFileSync(tmp, bundle);
  try {
    execFileSync(node, [tmp], { stdio: 'inherit' });
  } catch (e) {
    failed = 1;
  }
}
process.exit(failed);
