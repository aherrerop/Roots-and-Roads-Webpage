/**
 * Viator auto-close bot
 * ---------------------
 * Closes ("Sold out") Viator departures that are inside their per-product
 * close-window AND have ZERO guests across ALL booking platforms.
 *
 * The decision data comes from the BookingSheet web app (?closecheck) which
 * already aggregates every platform. This bot only ACTUATES on supplier.viator.com.
 *
 * Safety: DRY_RUN defaults to true — it reports what it WOULD close without
 * clicking. Set DRY_RUN=false to arm it. It only ever clicks "Sold out"
 * (reversible, no cancellation), never "Not operating" or "Available".
 *
 * Env (see README):
 *   BOOKING_WEBAPP_URL, ADMIN_KEY        — decision endpoint + auth
 *   VIATOR_EMAIL, VIATOR_PASSWORD        — supplier portal login
 *   VIATOR_STORAGE_STATE_B64 (optional)  — seed a pre-verified session
 *   DRY_RUN            (default "true")
 *   HEADLESS          (default "true")
 *   ALERT_ON_BLOCK    (default "true")   — email management if blocked (CAPTCHA)
 */

const fs = require('fs');
const path = require('path');
// playwright is required lazily inside main() so the pure helpers stay
// importable for unit tests without the browser dependency installed.

// Env is read lazily (no throw at import time) so the pure helpers below can be
// unit-tested via require(). main() validates required vars before doing work.
const CFG = {
  get bookingUrl() { return process.env.BOOKING_WEBAPP_URL; },
  get adminKey() { return process.env.ADMIN_KEY; },
  get email() { return process.env.VIATOR_EMAIL; },
  get password() { return process.env.VIATOR_PASSWORD; },
  dryRun: (process.env.DRY_RUN ?? 'true') !== 'false',
  headless: (process.env.HEADLESS ?? 'true') !== 'false',
  alertOnBlock: (process.env.ALERT_ON_BLOCK ?? 'true') !== 'false',
  statePath: path.join(__dirname, 'state.json'),
  availabilityUrl: 'https://supplier.viator.com/availability',
  loginUrl: 'https://supplier.viator.com/login',
  codePollAttempts: 6,     // ~6 * 10s = 1 min waiting for the 2FA email
  codePollDelayMs: 10000
};

// Thrown when the bot cannot proceed unattended (CAPTCHA). Triggers the
// email-management fallback instead of a hard failure.
class BlockedError extends Error {}

function requireEnv() {
  const missing = ['BOOKING_WEBAPP_URL', 'ADMIN_KEY', 'VIATOR_EMAIL', 'VIATOR_PASSWORD']
    .filter(n => !process.env[n]);
  if (missing.length) throw new Error(`Missing env: ${missing.join(', ')}`);
}

function log(...a) { console.log(new Date().toISOString(), ...a); }

async function main() {
  requireEnv();
  // 1) Ask the sheet for the all-platform guest counts + per-product windows.
  const data = await fetchJson(
    `${CFG.bookingUrl}?closecheck=1&key=${encodeURIComponent(CFG.adminKey)}`
  );
  if (!data.ok) throw new Error(`closecheck failed: ${data.error || 'unknown'}`);

  const productHours = data.productCloseHours || {};   // { "5631527P3": {hours,enabled,name} }
  const guestMap = data.guestMap || {};                // "date|H:mm|Language" -> guests
  const scanDates = data.scanDates || [];              // Madrid ISO days to scan
  const nowIso = data.now;
  const tzOffset = (String(nowIso).match(/([+-]\d{2}:\d{2})$/) || [])[1] || 'Z';
  const nowMs = Date.parse(nowIso);
  log(`closecheck ok. now=${nowIso} scanDates=${scanDates.join(',')} products=${Object.keys(productHours).join(',')}`);

  const { chromium } = require('playwright');
  const browser = await chromium.launch({
    headless: CFG.headless,
    args: ['--disable-blink-features=AutomationControlled']
  });
  // Present as a normal desktop Chrome so a legitimately-captured session is
  // less likely to be re-challenged. (This does not defeat CAPTCHAs — a fresh
  // login from a flagged IP will still be blocked; the saved session is what
  // avoids the login page entirely.)
  const context = await browser.newContext({
    ...(await loadState()),
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
    viewport: { width: 1366, height: 768 },
    locale: 'en-US',
    timezoneId: 'Europe/Madrid'
  });
  const page = await context.newPage();

  const closed = [], wouldClose = [], skipped = [];
  let blocked = null;

  try {
    await ensureLoggedIn(page);

    for (const date of scanDates) {
      await gotoAvailabilityForDate(page, date);
      // 2) Enumerate straight from Viator's page — covers group P3 AND private P5.
      const cards = (await readDepartureCards(page)).filter(c => cardOnDate(c, date));
      log(`date ${date}: ${cards.length} departure cards on page`);

      for (const card of cards) {
        if (!card.canSellOut) continue;   // already Sold out / not sellable — nothing to do
        const tag = { date, time: card.time24, language: card.language, product: card.productCode };
        const cfg = productHours[card.productCode];

        if (!cfg) { skipped.push({ ...tag, why: `product not in Control table` }); continue; }
        if (!cfg.enabled) { skipped.push({ ...tag, why: `product disabled in Control` }); continue; }

        const guests = lookupGuests(guestMap, date, card);
        if (guests > 0) { skipped.push({ ...tag, why: `has ${guests} guest(s) across platforms` }); continue; }

        const minutes = Math.round((departureMs(date, card.time24, tzOffset) - nowMs) / 60000);
        if (minutes <= 0) { skipped.push({ ...tag, why: `already started/passed` }); continue; }
        if (minutes > cfg.hours * 60) { skipped.push({ ...tag, why: `outside ${cfg.hours}h window (${minutes}m out)` }); continue; }

        const label = `${date} ${card.time24} ${titleLang(card.language)} [${card.productCode}] (${minutes}m out, window ${cfg.hours}h)`;
        if (CFG.dryRun) { log('WOULD CLOSE', label); wouldClose.push(label); }
        else { await closeCard(page, card); log('CLOSED', label); closed.push(label); }
      }
    }
  } catch (err) {
    if (err instanceof BlockedError) { blocked = err.message; log('BLOCKED:', err.message); }
    else throw err;
  } finally {
    await saveState(context).catch(() => {});
    await browser.close().catch(() => {});
  }

  // If blocked, hand the empty list (schedule view) to management to close by hand.
  if (blocked) {
    const list = (data.departures || []).filter(d => d.empty)
      .map(d => `• ${d.date} ${d.time} ${d.language} (${d.minutesUntilStart}m out)`)
      .join('\n') || '(check Viator availability for empty departures)';
    await sendAlert(blocked, list);
  }

  log('SUMMARY', JSON.stringify({
    dryRun: CFG.dryRun, closed: closed.length, wouldClose: wouldClose.length,
    skipped: skipped.length, blocked: Boolean(blocked)
  }));
  if (skipped.length) log('skipped:', JSON.stringify(skipped, null, 2));
}

/* ------------------------------- login ------------------------------- */

// Viator supplier login is TWO-STEP: email + "Next" -> password + submit ->
// (sometimes) a 2FA code emailed to Gmail. Selectors confirmed against the real
// page (email is type=text, placeholder "Email address").
async function ensureLoggedIn(page) {
  await page.goto(CFG.availabilityUrl, { waitUntil: 'domcontentloaded' });
  if (await isLoggedIn(page)) { log('Session reused — already logged in.'); return; }

  log('Not logged in — signing in.');
  await page.goto(CFG.loginUrl, { waitUntil: 'domcontentloaded' });

  // Step 1 — email, then "Next".
  const emailField = page.locator(
    'input[placeholder="Email address" i], input[type="email"], input[name="email"], input[name="username"]'
  ).first();
  try {
    await emailField.waitFor({ state: 'visible', timeout: 15000 });
  } catch (e) {
    log('LOGIN inputs on page:', JSON.stringify(await describeInputs(page)));
    if (await hasCaptcha(page)) throw new BlockedError('CAPTCHA on Viator login — manual close required.');
    throw new Error('Email field not found on login page.');
  }
  await emailField.fill(CFG.email);
  await clickFirst(page, ['button[type="submit"]', 'button:has-text("Next")', 'button:has-text("Continue")']);
  await page.waitForTimeout(1500);

  // Step 2 — password, then submit.
  const pwField = page.locator('input[type="password"], input[name="password"]').first();
  try {
    await pwField.waitFor({ state: 'visible', timeout: 15000 });
  } catch (e) {
    log('PASSWORD-STEP inputs on page:', JSON.stringify(await describeInputs(page)));
    if (await hasCaptcha(page)) throw new BlockedError('CAPTCHA on Viator login — manual close required.');
    throw new Error('Password field not found after email step.');
  }
  await pwField.fill(CFG.password);
  await clickFirst(page, ['button[type="submit"]', 'button:has-text("Log in")', 'button:has-text("Sign in")', 'button:has-text("Next")', 'button:has-text("Continue")']);
  await page.waitForTimeout(2500);

  if (await hasCaptcha(page)) throw new BlockedError('CAPTCHA on Viator login — manual close required.');

  // Step 3 — 2FA code (only sometimes).
  if (await needsCode(page)) {
    log('2FA code requested — fetching from Gmail.');
    const code = await waitForLoginCode();
    if (!code) throw new BlockedError('2FA code needed but none arrived in Gmail in time.');
    const codeField = page.locator(
      'input[autocomplete="one-time-code"], input[name*="code" i], input[placeholder*="code" i], input[type="tel"]'
    ).first();
    try {
      await codeField.waitFor({ state: 'visible', timeout: 8000 });
    } catch (e) {
      log('CODE-STEP inputs on page:', JSON.stringify(await describeInputs(page)));
      throw new Error('2FA code field not found.');
    }
    await codeField.fill(code);
    await clickFirst(page, ['button[type="submit"]', 'button:has-text("Verify")', 'button:has-text("Submit")', 'button:has-text("Continue")', 'button:has-text("Next")']);
    await page.waitForTimeout(2500);
  }

  await page.goto(CFG.availabilityUrl, { waitUntil: 'domcontentloaded' });
  if (!(await isLoggedIn(page))) {
    if (await hasCaptcha(page)) throw new BlockedError('CAPTCHA after code — manual close required.');
    log('POST-LOGIN inputs on page:', JSON.stringify(await describeInputs(page)));
    throw new Error('Login did not stick (check credentials / selectors).');
  }
  log('Logged in.');
}

// Click the first selector that exists — resilient to label/DOM variations.
async function clickFirst(page, selectors) {
  for (const sel of selectors) {
    const el = page.locator(sel).first();
    if (await el.count().catch(() => 0)) {
      await el.click().catch(() => {});
      return true;
    }
  }
  return false;
}

// Dump every input's identifying attributes — logged on any login-step failure
// so the exact selector can be fixed from one run, no guessing.
async function describeInputs(page) {
  return page.evaluate(() =>
    Array.from(document.querySelectorAll('input')).map(i => ({
      type: i.type, name: i.name || null, id: i.id || null,
      placeholder: i.placeholder || null, autocomplete: i.autocomplete || null,
      visible: !!(i.offsetWidth || i.offsetHeight)
    }))
  ).catch(() => []);
}

async function isLoggedIn(page) {
  // Availability nav visible == authenticated area reached.
  return page.getByText(/Availability/i).first().isVisible().catch(() => false);
}
async function hasCaptcha(page) {
  const f = page.locator('iframe[src*="recaptcha"], iframe[src*="hcaptcha"], iframe[title*="captcha" i]');
  if (await f.count().catch(() => 0)) return true;
  return page.getByText(/not a robot|verify you are human|captcha/i).first().isVisible().catch(() => false);
}
async function needsCode(page) {
  return page.getByText(/two-factor|verification code|enter the code|unique code/i)
    .first().isVisible().catch(() => false);
}

async function waitForLoginCode() {
  for (let i = 0; i < CFG.codePollAttempts; i++) {
    const r = await fetchJson(
      `${CFG.bookingUrl}?viatorcode=1&key=${encodeURIComponent(CFG.adminKey)}`
    ).catch(() => ({}));
    if (r && r.ok && r.found && r.code) { log(`Got 2FA code from Gmail (age ${r.ageMinutes}m).`); return r.code; }
    log(`Waiting for 2FA email… (${i + 1}/${CFG.codePollAttempts})`);
    await sleep(CFG.codePollDelayMs);
  }
  return null;
}

/* --------------------------- availability ---------------------------- */

async function gotoAvailabilityForDate(page, isoDate) {
  // TUNING POINT #2: date navigation. Try a query param first, then fall back
  // to the on-page date field. Our window is <=24h so today/tomorrow only.
  await page.goto(`${CFG.availabilityUrl}?date=${isoDate}`, { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(1500);
  // (If ?date is ignored, the default view still shows the next days incl.
  //  today+tomorrow; readDepartureCards tags every visible card and we match
  //  by date via the day header captured below.)
}

/**
 * Tag and read every departure card on the page in one pass. Returns
 * [{ id, productCode, time24, language, dayLabel, canSellOut }]. Each card is
 * tagged with data-vac-id so we can click its "Sold out" link precisely.
 */
async function readDepartureCards(page) {
  return page.evaluate(() => {
    const codeRe = /\((\d{5,}P\d+)\)/;
    const tgRe = /TG\d+~(\d{1,2}:\d{2})/;
    const langRe = /(English|German|Spanish|Italian|French)\s*Tour/i;
    const dayRe = /^[A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2}/;

    // Track the most recent day header while walking the DOM in order.
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
    let currentDay = '';
    const out = [];
    let idc = 0;
    const seen = new Set();

    while (walker.nextNode()) {
      const el = walker.currentNode;
      const own = (el.childElementCount === 0 ? (el.textContent || '') : '').trim();
      if (own && dayRe.test(own)) currentDay = own;

      const txt = el.innerText || '';
      if (!codeRe.test(txt) || !tgRe.test(txt)) continue;
      if (!/Sold out|Available|Not operating/i.test(txt)) continue;
      // smallest qualifying element = the card (skip if a child also qualifies)
      const childQualifies = Array.from(el.children).some(c => {
        const t = c.innerText || '';
        return codeRe.test(t) && tgRe.test(t) && /Sold out|Available|Not operating/i.test(t);
      });
      if (childQualifies) continue;

      const key = txt.slice(0, 120);
      if (seen.has(key)) continue;
      seen.add(key);

      const id = 'vac-' + (idc++);
      el.setAttribute('data-vac-id', id);
      out.push({
        id,
        productCode: (txt.match(codeRe) || [])[1] || '',
        time24: (txt.match(tgRe) || [])[1] || '',
        language: ((txt.match(langRe) || [])[1] || '').toLowerCase(),
        dayLabel: currentDay,
        // AVAILABLE rows expose a "Sold out" action; SOLD OUT rows expose "Available".
        canSellOut: /\bSold out\b/i.test(txt)
      });
    }
    return out;
  });
}

async function closeCard(page, card) {
  const scope = page.locator(`[data-vac-id="${card.id}"]`);
  const soldOut = scope.getByText(/^Sold out$/i).first();
  await soldOut.click();
  // Confirm dialog, if any.
  const confirm = page.getByRole('button', { name: /confirm|yes|sold out|ok/i }).first();
  if (await confirm.isVisible().catch(() => false)) await confirm.click();
  await page.waitForTimeout(800);
}

/* ------------------------------ matching ----------------------------- */

// Match the sheet's key normalisation: hour is NOT zero-padded ("16:00","9:00").
function normTime(t) {
  const m = String(t).match(/^(\d{1,2}):(\d{2})$/);
  return m ? `${Number(m[1])}:${m[2]}` : String(t);
}
function titleLang(s) {
  const w = String(s || '').toLowerCase().replace(/\s*tour$/, '').trim();
  return w ? w[0].toUpperCase() + w.slice(1) : '';
}
function lookupGuests(guestMap, date, card) {
  return Number(guestMap[`${date}|${normTime(card.time24)}|${titleLang(card.language)}`] || 0);
}
function departureMs(date, time24, offset) {
  const m = String(time24).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return NaN;
  const hh = String(Number(m[1])).padStart(2, '0');
  return Date.parse(`${date}T${hh}:${m[2]}:00${offset}`);
}
// Keep only cards belonging to `date` (the availability page may show several
// days). dayLabel like "Fri, Aug 28"; if it's missing, trust the navigation.
function cardOnDate(card, date) {
  if (!card.dayLabel) return true;
  const exp = new Date(`${date}T12:00:00`)
    .toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
  return card.dayLabel.replace(/\s+/g, ' ').trim().startsWith(exp);
}

/* ------------------------------ plumbing ----------------------------- */

async function sendAlert(reason, list) {
  if (!CFG.alertOnBlock) return;
  const url = `${CFG.bookingUrl}?viatoralert=1&key=${encodeURIComponent(CFG.adminKey)}` +
    `&reason=${encodeURIComponent(reason)}&tours=${encodeURIComponent(list)}`;
  const r = await fetchJson(url).catch(e => ({ ok: false, error: String(e) }));
  log('alert email:', r.ok ? 'sent' : `FAILED ${r.error}`);
}

async function loadState() {
  if (process.env.VIATOR_STORAGE_STATE_B64 && !fs.existsSync(CFG.statePath)) {
    fs.writeFileSync(CFG.statePath, Buffer.from(process.env.VIATOR_STORAGE_STATE_B64, 'base64'));
    log('Seeded session from VIATOR_STORAGE_STATE_B64.');
  }
  if (fs.existsSync(CFG.statePath)) return { storageState: CFG.statePath };
  return {};
}
async function saveState(context) {
  await context.storageState({ path: CFG.statePath });
}

async function fetchJson(url) {
  const res = await fetch(url, { redirect: 'follow' });
  const text = await res.text();
  try { return JSON.parse(text); }
  catch { throw new Error(`Non-JSON from ${url.split('?')[0]}: ${text.slice(0, 200)}`); }
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

if (require.main === module) {
  main().then(() => process.exit(0)).catch(err => { console.error('FATAL', err); process.exit(1); });
}

// Exported for unit tests (viator-autoclose/selftest.js).
module.exports = { normTime, titleLang, lookupGuests, departureMs, cardOnDate };
