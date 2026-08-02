/**
 * Cross-platform test runner (no bash/WSL needed — runs on plain Windows too).
 *
 * Mirrors the Apps Script model: each suite runs in a fresh Node process over
 * one shared global scope (mocks + a project's sources + the suite), exactly
 * like Apps Script loads a project.
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

// Sources per Apps Script project (shared global scope, as on the server).
const CONTROL = [
  path.join(AS, 'control', 'assignShifts.gs'),
  path.join(AS, 'control', 'guidePortal.gs')
];
const BOOKING = [
  path.join(AS, 'booking', 'bookingList_v2.gs'),
  path.join(AS, 'booking', 'websiteAvailabilityUpdate.gs')
];

// Each suite = [project sources, suite file].
const SUITES = [
  [CONTROL, 'tests.js'],
  [CONTROL, 'tests2.js'],
  [CONTROL, 'tests3.js'],
  [CONTROL, 'tests4.js'],
  [CONTROL, 'tests5.js'],
  [CONTROL, 'tests6.js'],          // Italian + French: control side
  [BOOKING, 'booking-tests.js']    // Italian + French: booking side
];

// Gmail-side integration suite needs the stateful Gmail mock layered on top.
const GMAIL_SUITE = [BOOKING, 'gmail-tests.js', 'gmail-mock.js'];
// Portal end-to-end runs the real apiTours_/apiAssign_/apiSave_ chain.
const PORTAL_E2E = [CONTROL, 'portal-e2e-tests.js'];

let failed = 0;
for (const [sources, suite] of SUITES) {
  const bundle = [path.join(testDir, 'mock.js'), ...sources, path.join(testDir, suite)]
    .map(read).join('\n');
  const tmp = path.join(os.tmpdir(), 'rr_' + suite.replace(/\W/g, '_') + '.js');
  fs.writeFileSync(tmp, bundle);
  try {
    execFileSync(node, [tmp], { stdio: 'inherit' });
  } catch (e) {
    failed = 1;
  }
}

// Gmail-side integration suite: mock.js + gmail-mock.js (overrides the Gmail
// stub) + booking sources + the suite.
{
  const [sources, suite, gmailMock] = GMAIL_SUITE;
  const bundle = [path.join(testDir, 'mock.js'), path.join(testDir, gmailMock), ...sources, path.join(testDir, suite)]
    .map(read).join('\n');
  const tmp = path.join(os.tmpdir(), 'rr_gmail.js');
  fs.writeFileSync(tmp, bundle);
  try { execFileSync(node, [tmp], { stdio: 'inherit' }); } catch (e) { failed = 1; }
}

// Portal end-to-end: control sources + the real API chain.
{
  const [sources, suite] = PORTAL_E2E;
  const bundle = [path.join(testDir, 'mock.js'), ...sources, path.join(testDir, suite)].map(read).join('\n');
  const tmp = path.join(os.tmpdir(), 'rr_portal_e2e.js');
  fs.writeFileSync(tmp, bundle);
  try { execFileSync(node, [tmp], { stdio: 'inherit' }); } catch (e) { failed = 1; }
}

// Standalone suite: the portal's own client JS (extracted from guide/index.html),
// no mock/Apps Script bundle needed.
try {
  execFileSync(node, [path.join(testDir, 'portal-tests.js')], { stdio: 'inherit' });
} catch (e) {
  failed = 1;
}

process.exit(failed);
