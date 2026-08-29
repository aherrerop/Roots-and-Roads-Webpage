/**
 * One-time: capture a VERIFIED Viator session so CI can reuse it (avoids the
 * ~monthly device challenge most runs). Run locally on your own machine:
 *
 *   cd viator-autoclose
 *   npm install
 *   npx playwright install chromium
 *   node capture-session.js
 *
 * A browser opens. Log in by hand (enter the emailed 2FA code if asked). Once
 * you're on the Availability page, come back here and press Enter. It writes
 * state.json AND prints a base64 blob — paste that into the GitHub secret
 * VIATOR_STORAGE_STATE_B64 so the first CI run starts already trusted.
 */
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { chromium } = require('playwright');

const STATE = path.join(__dirname, 'state.json');

(async () => {
  const browser = await chromium.launch({ headless: false });
  // Match the CI bot's context so the captured session replays cleanly.
  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
    viewport: { width: 1366, height: 768 },
    locale: 'en-US',
    timezoneId: 'Europe/Madrid'
  });
  const page = await context.newPage();
  await page.goto('https://supplier.viator.com/login');

  console.log('\n>>> Log in manually in the browser window (do the 2FA email code if asked).');
  console.log('>>> When you can see the Availability page, press Enter here.\n');

  await new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question('Press Enter once logged in… ', () => { rl.close(); resolve(); });
  });

  await context.storageState({ path: STATE });
  const b64 = fs.readFileSync(STATE).toString('base64');
  console.log(`\nSaved ${STATE}`);
  console.log('\n--- VIATOR_STORAGE_STATE_B64 (paste into the GitHub secret) ---\n');
  console.log(b64);
  console.log('\n--- end ---\n');

  await browser.close();
  process.exit(0);
})();
