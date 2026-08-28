/******************************************************
 * VIATOR AUTO-CLOSE — data endpoints
 *
 * Lives in the BookingSheet Apps Script project. Exposes two read-only,
 * ADMIN_KEY-guarded payloads that the external GitHub Action bot consumes:
 *
 *   doGet(?closecheck=1&key=ADMIN_KEY[&hours=10])
 *     -> which upcoming departures are inside the final N-hour window AND
 *        have ZERO guests across ALL platforms (Website/Viator/GYG/Airbnb/
 *        Guruwalk/Free Tour). The bot closes exactly the empty ones on Viator.
 *
 *   doGet(?viatorcode=1&key=ADMIN_KEY)
 *     -> the most recent Viator supplier-portal login verification code from
 *        Gmail, so the bot can clear the ~monthly device challenge unattended.
 *
 * The branches are wired into the single doGet() in websiteAvailabilityUpdate.gs
 * (Apps Script allows only one doGet per project).
 *
 * WHY the decision lives here (not in the bot): the schedule + the all-platform
 * guest counts already exist in this project (websiteReadCombinedSchedule_ /
 * websiteReadBookedGuestsMap_). The bot stays dumb — it only actuates.
 ******************************************************/

// Fallback close window (hours before start) for a product with no valid row in
// the Control tab. The real per-product values live in the Control sheet — see
// viatorReadProductCloseHours_() and setupViatorControls() (Control project).
const VIATOR_DEFAULT_WINDOW_HOURS = 10;

// Horizon used when the Control tab yields no enabled products at all.
const VIATOR_DEFAULT_MAX_WINDOW_HOURS = 24;

// Per-product close-hours control, read from the Control spreadsheet's "Control"
// tab. Editable block written by setupViatorControls() at columns E..H:
//   E: Product name | F: Product code | G: Close hours before (if empty) | H: Enabled
const VIATOR_CONTROL_TAB = 'Control';
const VIATOR_CONTROL_FIRST_ROW = 2;   // row 1 is the header
const VIATOR_CONTROL_COL = 5;         // column E
const VIATOR_CONTROL_WIDTH = 4;       // E..H

/**
 * Gmail search + extraction config for the supplier-portal login code.
 * Confirmed against a real message (2026-08): the 2FA code email comes from
 * account@t1.viator.com, subject "Your two-factor authentication code request",
 * body: "Here's your unique code: 572431" (also prints "Supplier ID: 5631527",
 * so the regex MUST anchor on "unique code:" — never a bare digit run).
 */
const VIATOR_CODE_QUERY =
  'from:account@t1.viator.com subject:"two-factor authentication code" newer_than:1d';

// Anchored on the label so it can't pick up the Supplier ID or any other number.
const VIATOR_CODE_REGEX = /unique code:\s*(\d{4,8})/i;

// Ignore any candidate code email older than this — a stale code is useless and
// risks typing an expired one.
const VIATOR_CODE_MAX_AGE_MIN = 15;

// Where the "please close these manually" alert goes when the bot is blocked
// (CAPTCHA, or a code was needed but couldn't be read). Falls back to the
// booking account's own address if the script property is unset.
const VIATOR_ALERT_PROP = 'VIATOR_ALERT_EMAIL';


/**
 * Read the per-product close-hours from the Control tab.
 * @return {{products:Array, map:Object, maxHours:number}}
 */
function viatorReadProductCloseHours_() {
  const control = SpreadsheetApp.openById(WEBSITE_CONTROL_SPREADSHEET_ID);
  const sh = control.getSheetByName(VIATOR_CONTROL_TAB);
  const products = [];

  if (sh && sh.getLastRow() >= VIATOR_CONTROL_FIRST_ROW) {
    const n = sh.getLastRow() - VIATOR_CONTROL_FIRST_ROW + 1;
    const rows = sh.getRange(VIATOR_CONTROL_FIRST_ROW, VIATOR_CONTROL_COL, n, VIATOR_CONTROL_WIDTH).getValues();
    rows.forEach(r => {
      const name = websiteClean_(r[0]);
      const code = websiteClean_(r[1]).toUpperCase();
      if (!/^\d+P\d+$/.test(code)) return;   // only genuine Viator product codes
      const hours = Number(r[2]) > 0 ? Number(r[2]) : VIATOR_DEFAULT_WINDOW_HOURS;
      const enabledRaw = r[3];
      const enabled = enabledRaw === '' || enabledRaw === true ||
        String(enabledRaw).trim().toUpperCase() === 'TRUE';
      products.push({ code, name, hours, enabled });
    });
  }

  const map = {};
  products.forEach(p => { map[p.code] = { hours: p.hours, enabled: p.enabled, name: p.name }; });
  const maxHours = products
    .filter(p => p.enabled)
    .reduce((m, p) => Math.max(m, p.hours), 0) || VIATOR_DEFAULT_MAX_WINDOW_HOURS;

  return { products, map, maxHours };
}


/**
 * Build the close-check payload.
 *
 * The bot enumerates the actual departures from Viator's availability page (so
 * BOTH group P3 and private P5 are covered, with no dependency on the website
 * schedule and no risk of private tours leaking onto the public site). This
 * endpoint therefore returns the raw all-platform guest counts as `guestMap`
 * (keyed "date|H:mm|Language") for the days the bot will scan, plus the
 * per-product windows. The bot marks a Viator card "Sold out" when its
 * guestMap count is 0 and it's inside that product's window.
 *
 * `departures` (schedule-derived, group tours) is also returned, purely for
 * human-readable logging/debugging — the bot's decision uses guestMap.
 *
 * @param {(number|string)} windowHoursOverride
 */
function viatorClosecheckPayload_(windowHoursOverride) {
  const control = viatorReadProductCloseHours_();
  const hours = Number(windowHoursOverride) > 0 ? Number(windowHoursOverride) : control.maxHours;
  const now = new Date();
  const horizon = new Date(now.getTime() + hours * 60 * 60 * 1000);

  const slots = websiteReadCombinedSchedule_();
  const booked = websiteReadBookedGuestsMap_();

  // Scan enough calendar days (Madrid) to cover the largest window, +1 for a
  // window that crosses midnight.
  const daysToScan = Math.max(1, Math.ceil(hours / 24)) + 1;
  const scanDates = [];
  for (let offset = 0; offset < daysToScan; offset++) {
    const dayNoon = new Date(now.getFullYear(), now.getMonth(), now.getDate() + offset, 12, 0, 0);
    scanDates.push(websiteDateKey_(dayNoon));
  }
  const scanSet = new Set(scanDates);

  // guestMap: all-platform counts for the scan days only (keeps payload small).
  const guestMap = {};
  booked.forEach((v, k) => {
    if (scanSet.has(String(k).split('|')[0])) guestMap[k] = Number(v || 0);
  });

  // Schedule-derived group departures inside the horizon — logging aid only.
  const departures = [];
  for (let offset = 0; offset < daysToScan; offset++) {
    const dayNoon = new Date(now.getFullYear(), now.getMonth(), now.getDate() + offset, 12, 0, 0);
    const iso = websiteDateKey_(dayNoon);
    const weekday = Utilities.formatDate(dayNoon, WEBSITE_TZ, 'EEEE');

    slots.forEach(slot => {
      if (slot.day !== weekday) return;
      if (slot.activeFrom && dayNoon < slot.activeFrom) return;
      if (slot.activeUntil && dayNoon > slot.activeUntil) return;

      const start = viatorSlotStart_(now.getFullYear(), now.getMonth(), now.getDate() + offset, slot.time);
      if (!start) return;
      if (start <= now || start > horizon) return;

      const key = websiteAvailabilityKey_(iso, slot.time, slot.language);
      const guests = Number(booked.get(key) || 0);

      departures.push({
        date: iso,
        time: slot.time,
        displayTime: slot.displayTime,
        language: slot.language,
        guests,
        empty: guests === 0,
        minutesUntilStart: Math.round((start - now) / 60000)
      });
    });
  }

  departures.sort((a, b) => a.minutesUntilStart - b.minutesUntilStart);

  return {
    scanDates,
    guestMap,
    now: Utilities.formatDate(now, WEBSITE_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    tz: WEBSITE_TZ,
    horizonHours: hours,
    // The bot applies these per Viator row (keyed by productCode from the page).
    productCloseHours: control.map,
    products: control.products,
    departures
  };
}


/**
 * Absolute instant for a slot's wall-clock start. The project timezone is
 * Europe/Madrid (appsscript.json), so a local Date constructor yields the
 * correct Madrid instant, DST included.
 */
function viatorSlotStart_(year, monthIndex, day, timeHHmm) {
  const t = websiteNormalizeTime_(timeHHmm);
  const m = t.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return new Date(year, monthIndex, day, Number(m[1]), Number(m[2]), 0, 0);
}


/**
 * Latest Viator login verification code from Gmail (most recent matching thread,
 * newest message). Returns null-ish payload when none is fresh enough.
 *
 * @return {{code:(string|null), ageMinutes:(number|null), subject:(string|null),
 *           from:(string|null), found:boolean}}
 */
function viatorLatestLoginCode_() {
  const threads = GmailApp.search(VIATOR_CODE_QUERY, 0, 5);
  let best = null;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const when = msg.getDate();
      if (!best || when > best.when) {
        best = { when, msg };
      }
    });
  });

  if (!best) {
    return { found: false, code: null, ageMinutes: null, subject: null, from: null };
  }

  const ageMinutes = Math.round((Date.now() - best.when.getTime()) / 60000);
  const haystack = (best.msg.getSubject() || '') + '\n' + (best.msg.getPlainBody() || '');
  const match = haystack.match(VIATOR_CODE_REGEX);
  const fresh = ageMinutes <= VIATOR_CODE_MAX_AGE_MIN;

  return {
    found: Boolean(match) && fresh,
    code: match && fresh ? match[1] : null,
    ageMinutes,
    subject: best.msg.getSubject() || null,
    from: best.msg.getFrom() || null
  };
}


/**
 * Email management the list of tours the bot could NOT close itself (CAPTCHA, or
 * a code was needed but unreadable), so they can be closed by hand. The list is
 * passed as the `tours` parameter (newline- or "; "-separated). Recipient comes
 * from the VIATOR_ALERT_EMAIL script property, else the booking account.
 *
 * @return {boolean} whether an email was sent
 */
function viatorSendAlert_(e) {
  const tours = String(e?.parameter?.tours || '').trim();
  const reason = String(e?.parameter?.reason || 'The Viator auto-close bot was blocked.');
  const to = PropertiesService.getScriptProperties().getProperty(VIATOR_ALERT_PROP) ||
    Session.getEffectiveUser().getEmail();
  if (!to) return false;

  const body =
    reason + '\n\n' +
    'These empty Viator departures still need to be marked "Sold out" MANUALLY ' +
    'in supplier.viator.com (Availability):\n\n' +
    (tours || '(no list provided)') + '\n\n' +
    '— Roots & Roads Viator auto-close';

  MailApp.sendEmail(to, '⚠️ Viator auto-close needs a manual hand', body);
  return true;
}


/**
 * Shared auth + JSON responder for the Viator branches of doGet().
 * Called from websiteAdminRun_-style routing in websiteAvailabilityUpdate.gs.
 */
function viatorRespond_(e) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  const given = String(e?.parameter?.key || '');

  let out;
  try {
    if (!key) {
      out = { ok: false, error: 'ADMIN_KEY not configured' };
    } else if (given !== key) {
      out = { ok: false, error: 'Bad key' };
    } else if (String(e?.parameter?.closecheck || '') === '1') {
      out = { ok: true, ...viatorClosecheckPayload_(e?.parameter?.hours) };
    } else if (String(e?.parameter?.viatorcode || '') === '1') {
      out = { ok: true, ...viatorLatestLoginCode_() };
    } else if (String(e?.parameter?.viatoralert || '') === '1') {
      out = { ok: true, sent: viatorSendAlert_(e) };
    } else {
      out = { ok: false, error: 'No Viator action' };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }

  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
