#!/bin/bash
# Runs all system tests. Requires node. Usage: bash test/run-tests.sh
cd "$(dirname "$0")"
node check-references.js || exit 1
cat mock.js ../apps-script/control/assignShifts.gs ../apps-script/control/guidePortal.gs tests.js > /tmp/rr_run1.js && node /tmp/rr_run1.js
cat mock.js ../apps-script/control/assignShifts.gs ../apps-script/control/guidePortal.gs tests2.js > /tmp/rr_run2.js && node /tmp/rr_run2.js
cat mock.js ../apps-script/control/assignShifts.gs ../apps-script/control/guidePortal.gs tests3.js > /tmp/rr_run3.js && node /tmp/rr_run3.js
cat mock.js ../apps-script/control/assignShifts.gs ../apps-script/control/guidePortal.gs tests4.js > /tmp/rr_run4.js && node /tmp/rr_run4.js
cat mock.js ../apps-script/control/assignShifts.gs ../apps-script/control/guidePortal.gs tests5.js > /tmp/rr_run5.js && node /tmp/rr_run5.js
