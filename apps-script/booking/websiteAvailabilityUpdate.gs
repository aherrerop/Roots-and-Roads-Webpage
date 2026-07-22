/******************************************************
 * WEBSITE AVAILABILITY
 * Put this file in the BookingSheet Apps Script project.
 ******************************************************/

const WEBSITE_CONTROL_SPREADSHEET_ID = '1A8RrqIoWw-HpxCLDRGVJLcflja8tSGEOgaXd-2sSkLs';

const WEBSITE_TZ = 'Europe/Madrid';
const WEBSITE_CAPACITY = 20;
const WEBSITE_CUTOFF_HOUR = 21;

const WEBSITE_ACTIVE_BOOKING_SHEETS = [
  'English Tours',
  'German Tours',
  'Spanish Tours',
  'Italian Tours',
  'French Tours'
];

const WEBSITE_CONTROL_WEEKLY_TAB = 'Weekly_Schedule';
const WEBSITE_CONTROL_EXTRA_TAB = 'Website_Schedule';


function doGet(e) {
  const callback = String(e?.parameter?.callback || '').trim();
  const ym = String(e?.parameter?.ym || '').trim();

  // ---- ADMIN: remote-run hook for the Mobile Controls block in the ----
  // ---- Control sheet. Requires the ADMIN_KEY script property.       ----
  // GET ?admin=1&key=<ADMIN_KEY>&fn=runBookingSystem|runBookingAudit
  if (String(e?.parameter?.admin || '') === '1') {
    return websiteAdminRun_(e);
  }

  let payload;

  try {
    if (!/^\d{4}-\d{2}$/.test(ym)) {
      payload = { error: 'Bad ym. Use YYYY-MM' };
    } else {
      payload = {
        ym,
        days: websiteBuildMonthAvailability_(ym)
      };
    }
  } catch (err) {
    payload = {
      error: String(err?.message || err)
    };
  }

  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}


function websiteBuildMonthAvailability_(ym) {
  const scheduleSlots = websiteReadCombinedSchedule_();
  const bookedMap = websiteReadBookedGuestsMap_();

  const [year, month] = ym.split('-').map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();
  const out = {};

  for (let day = 1; day <= daysInMonth; day++) {
    const dateObj = new Date(year, month - 1, day, 12, 0, 0);
    const iso = websiteDateKey_(dateObj);
    const weekday = Utilities.formatDate(dateObj, WEBSITE_TZ, 'EEEE');

    scheduleSlots.forEach(slot => {
      if (slot.day !== weekday) return;
      if (slot.activeFrom && dateObj < slot.activeFrom) return;
      if (slot.activeUntil && dateObj > slot.activeUntil) return;

      const key = websiteAvailabilityKey_(iso, slot.time, slot.language);
      const booked = Number(bookedMap.get(key) || 0);
      const spotsLeft = Math.max(0, WEBSITE_CAPACITY - booked);
      const closed = websiteClosedByCutoff_(dateObj);

      if (!out[iso]) out[iso] = [];

      out[iso].push({
        language: slot.language,
        time: slot.displayTime,
        spotsLeft,
        available: !closed && spotsLeft > 0,
        status: (!closed && spotsLeft > 0) ? 'AVAILABLE' : 'CLOSED'
      });
    });
  }

  Object.keys(out).forEach(iso => {
    out[iso].sort((a, b) => {
      const lang = a.language.localeCompare(b.language);
      if (lang !== 0) return lang;
      return websiteTimeToMinutes_(a.time) - websiteTimeToMinutes_(b.time);
    });
  });

  return out;
}


function websiteReadCombinedSchedule_() {
  const control = SpreadsheetApp.openById(WEBSITE_CONTROL_SPREADSHEET_ID);
  const slots = new Map();

  websiteReadWeeklyScheduleIntoMap_(control, slots);
  websiteReadWebsiteScheduleIntoMap_(control, slots);

  return Array.from(slots.values());
}


function websiteReadWeeklyScheduleIntoMap_(control, slots) {
  const sh = control.getSheetByName(WEBSITE_CONTROL_WEEKLY_TAB);
  if (!sh || sh.getLastRow() < 2) return;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  const display = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues();

  values.forEach((row, i) => {
    const day = websiteClean_(display[i][0]);
    const time = websiteNormalizeTime_(display[i][1]);
    const language = websiteNormalizeLanguage_(display[i][2]);

    if (!day || !time || !language) return;

    const activeFrom = row[4] ? websiteDateOnly_(new Date(row[4])) : null;
    const activeUntil = row[5] ? websiteDateOnly_(new Date(row[5])) : null;

    const key = websiteScheduleKey_(day, time, language);

    if (!slots.has(key)) {
      slots.set(key, {
        day,
        time,
        displayTime: websiteDisplayTime_(time),
        language,
        activeFrom,
        activeUntil
      });
    }
  });
}


function websiteReadWebsiteScheduleIntoMap_(control, slots) {
  const sh = control.getSheetByName(WEBSITE_CONTROL_EXTRA_TAB);
  if (!sh || sh.getLastRow() < 2) return;

  const display = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues();

  display.forEach(row => {
    const day = websiteClean_(row[0]);
    const time = websiteNormalizeTime_(row[1]);
    const language = websiteNormalizeLanguage_(row[2]);

    if (!day || !time || !language) return;

    const key = websiteScheduleKey_(day, time, language);

    if (!slots.has(key)) {
      slots.set(key, {
        day,
        time,
        displayTime: websiteDisplayTime_(time),
        language,
        activeFrom: null,
        activeUntil: null
      });
    }
  });
}


function websiteReadBookedGuestsMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = new Map();

  WEBSITE_ACTIVE_BOOKING_SHEETS.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const display = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getDisplayValues();

    values.forEach((row, i) => {
      const guests = Number(row[2] || 0);
      const date = websiteNormalizeDate_(row[3]);
      const time = websiteNormalizeTime_(display[i][4]);
      const language = websiteSheetToLanguage_(sheetName);

      if (!guests || !date || !time || !language) return;

      const iso = websiteDateKey_(date);
      const key = websiteAvailabilityKey_(iso, time, language);

      map.set(key, Number(map.get(key) || 0) + guests);
    });
  });

  return map;
}


function websiteClosedByCutoff_(dateObj) {
  const d = websiteDateOnly_(dateObj);
  const cutoff = new Date(d);

  cutoff.setDate(cutoff.getDate() - 1);
  cutoff.setHours(WEBSITE_CUTOFF_HOUR, 0, 0, 0);

  return new Date() >= cutoff;
}


function websiteScheduleKey_(day, time, language) {
  return [
    websiteClean_(day),
    websiteNormalizeTime_(time),
    websiteNormalizeLanguage_(language)
  ].join('|');
}


function websiteAvailabilityKey_(iso, time, language) {
  return [
    iso,
    websiteNormalizeTime_(time),
    websiteNormalizeLanguage_(language)
  ].join('|');
}


function websiteClean_(v) {
  return String(v || '').replace(/\s+/g, ' ').trim();
}


function websiteNormalizeLanguage_(v) {
  const s = websiteClean_(v).toLowerCase();

  if (s.includes('german') || s.includes('deutsch') || s.includes('aleman') || s.includes('alemán')) return 'German';
  if (s.includes('spanish') || s.includes('espanol') || s.includes('español') || s.includes('castellano')) return 'Spanish';
  if (s.includes('italian') || s.includes('italiano') || s.includes('italiana') || s.includes('italien')) return 'Italian';
  if (s.includes('french') || s.includes('français') || s.includes('francais') || s.includes('francese') || s.includes('frances')) return 'French';

  return 'English';
}


function websiteNormalizeTime_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, WEBSITE_TZ, 'H:mm');
  }

  const s = websiteClean_(value);
  if (!s) return '';

  const ampm = s.match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)$/i);
  if (ampm) {
    let hour = Number(ampm[1]);
    const minute = Number(ampm[2] || 0);
    const suffix = ampm[3].toUpperCase();

    if (suffix === 'PM' && hour !== 12) hour += 12;
    if (suffix === 'AM' && hour === 12) hour = 0;

    return `${hour}:${String(minute).padStart(2, '0')}`;
  }

  const simple = s.match(/^(\d{1,2}):(\d{2})$/);
  if (simple) return `${Number(simple[1])}:${simple[2]}`;

  return s;
}


function websiteDisplayTime_(time) {
  const t = websiteNormalizeTime_(time);
  const m = t.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return t;

  let hour = Number(m[1]);
  const minute = m[2];
  const suffix = hour >= 12 ? 'PM' : 'AM';

  let h12 = hour % 12;
  if (h12 === 0) h12 = 12;

  return `${h12}:${minute} ${suffix}`;
}


function websiteTimeToMinutes_(time) {
  const t = websiteNormalizeTime_(time);
  const m = t.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return 99999;

  return Number(m[1]) * 60 + Number(m[2]);
}


function websiteNormalizeDate_(value) {
  if (value instanceof Date && !isNaN(value)) return websiteDateOnly_(value);

  const s = websiteClean_(value);
  if (!s) return null;

  const d = new Date(s);
  if (isNaN(d)) return null;

  return websiteDateOnly_(d);
}


function websiteDateOnly_(dateObj) {
  const d = new Date(dateObj);
  d.setHours(12, 0, 0, 0);
  return d;
}


function websiteDateKey_(dateObj) {
  return Utilities.formatDate(dateObj, WEBSITE_TZ, 'yyyy-MM-dd');
}


function websiteSheetToLanguage_(sheetName) {
  if (sheetName === 'German Tours') return 'German';
  if (sheetName === 'Spanish Tours') return 'Spanish';
  if (sheetName === 'Italian Tours') return 'Italian';
  if (sheetName === 'French Tours') return 'French';
  return 'English';
}


function debugWebsiteAvailabilityJuly2026() {
  const result = websiteBuildMonthAvailability_('2026-07');
  console.log(JSON.stringify(result, null, 2));
}

/**
 * Remote-run endpoint used by the Control sheet's Mobile Controls block.
 * Only the two whitelisted booking functions can run, and only with the
 * correct ADMIN_KEY (Script Property in THIS project — set it to the same
 * value as the Control project's ADMIN_KEY).
 */
function websiteAdminRun_(e) {
  let out;
  try {
    const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
    const given = String(e?.parameter?.key || '');
    const fn = String(e?.parameter?.fn || '');
    const allowed = { runBookingSystem: runBookingSystem, runBookingAudit: runBookingAudit };

    if (!key) out = { ok: false, error: 'ADMIN_KEY not configured' };
    else if (given !== key) out = { ok: false, error: 'Bad key' };
    else if (!allowed[fn]) out = { ok: false, error: 'Function not allowed: ' + fn };
    else { allowed[fn](); out = { ok: true, ran: fn }; }
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
