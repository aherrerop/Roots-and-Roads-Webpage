/******************************************************
 * ROOTS & ROADS BOOKING SYSTEM  (v2)
 *
 * WHAT THIS DOES
 *   Reads Gmail messages that your Gmail filters have already labelled per
 *   source (Confirmations / Cancellations / Modifications), parses each
 *   booking, and keeps the language tabs of the BookingSheet in sync.
 *   Completed tours are aggregated into the "Done Tours" tab.
 *
 * SOURCES SUPPORTED
 *   Viator, Guruwalk, Airbnb, Website, GetYourGuide, Free Tour
 *
 * SHEET TABS
 *   English / German / Spanish / Italian / French Tours  -> active bookings
 *   Done Tours                                     -> aggregated past tours
 *   Errors                                         -> run-time error log
 *
 * TRIGGERS (set these in Apps Script > Triggers)
 *   runBookingSystem()  -> every 5 min. Fast pass, skips threads already
 *                          tagged Processed.
 *   runBookingAudit()   -> 2-3x/day. Identical pipeline but ignores the
 *                          Processed marker, so every labelled thread gets
 *                          fully re-read for consistency.
 *
 * WEBSITE ENDPOINT
 *   doPost(e)
 *
 * ---------------------------------------------------------------
 * WHAT CHANGED FROM v1 (see chat for full rationale)
 *
 *   1) LABEL LIFECYCLE, now consistent across all three email types:
 *      - "Publishing Pages/Processed" is now added to confirmation,
 *        modification AND cancellation threads once each has been read and
 *        acted on (v1 only added it to confirmations). This is what lets the
 *        fast (5-min) run skip a thread's body on every later pass; the
 *        twice-daily audit ignores the marker and re-reads everything. This
 *        is also the main fix for the "Service invoked too many times for
 *        one day: gmail" errors — v1's cancellation pass and its supporting
 *        cache force-read every Cancellations-labelled thread's body on
 *        EVERY 5-minute run forever, and Viator's old "Amended" pipeline and
 *        the GYG modification handler did the same for Modifications. Those
 *        are now Processed-gated like everything else on fast runs.
 *      - When a booking is cancelled, the matching Confirmations AND
 *        Modifications threads for that booking id have their
 *        Confirm/Modify label REMOVED and the Cancel label ADDED (v1 only
 *        ever relabelled the Confirm thread). All three are archived.
 *      - When a booking reaches Done, EVERY other label this system uses
 *        (Confirm/Modify/Cancel/Processed) is stripped from the thread and
 *        only the Done label is left. This keeps the live
 *        Confirmations/Modifications labels showing ONLY upcoming tours, so
 *        a manager scanning Gmail can spot anything that looks "lost" (i.e.
 *        present in Gmail but missing from the sheet) at a glance.
 *      - Cancelled bookings are terminal: once labelled Cancel they are
 *        archived and left alone forever (never touched again, never moved
 *        to Done).
 *
 *   2) VIATOR "AMENDED" LABEL REMOVED. Your Gmail filter now sends every
 *      Viator "Amended Booking:" email straight to
 *      Publishing Pages/Viator/Modifications. Those emails are
 *      delta-style ("1 traveler added" / "1 traveler removed"), not a full
 *      new booking state, so the Viator branch of the modification pipeline
 *      still sums deltas per booking id across a thread's full message
 *      history (same math as v1's dedicated Amended handler) — it just now
 *      reads from the Modifications label and is Processed-gated like every
 *      other modification thread.
 *
 *   3) GUIDE PORTAL: the portal reads the BookingSheet directly through its
 *      own web app, so rows written here are visible to guides as soon as
 *      the portal refreshes — no webhook needed.
 *
 *   4) Read/unread and inbox state are still NEVER used as a processing
 *      signal anywhere (unchanged from v1) — only labels are.
 *
 *   5) Fast-run thread caps are now smaller than audit-run caps
 *      (MAX_THREADS_FAST / MAX_THREADS_AUDIT) so a 5-minute pass reads less
 *      per label, and the twice-daily audit is allowed to clear a bigger
 *      backlog.
 * ---------------------------------------------------------------
 ******************************************************/


/******************************************************
 * 1. CONFIGURATION
 ******************************************************/

const RNR = {

  SHEETS: {
    ENGLISH: 'English Tours',
    GERMAN: 'German Tours',
    SPANISH: 'Spanish Tours',
    ITALIAN: 'Italian Tours',
    FRENCH: 'French Tours',
    DONE: 'Done Tours',
    ERRORS: 'Errors'
  },

  ACTIVE_HEADERS: [
    'Name',
    'Phone',
    'Number of Guests',
    'Tour date',
    'Time',
    'Source',
    'Income',
    'Booking ID',
    'Notes'
  ],

  DONE_HEADERS: [
    'Date',
    'Time',
    'Language',
    'Number of Free Guests',
    'Number of Paid Guests',
    'Income',
    'Guide',
    'Comments'
  ],

  // Structured + deduplicated: a repeating unresolved error updates its
  // existing row's Count/Last seen instead of adding a row every 5 minutes.
  ERROR_HEADERS: [
    'Timestamp',
    'Type',
    'Details',
    'Raw data',
    'Count',
    'Last seen'
  ],

  /**
   * Gmail label names. These MUST match the labels your Gmail filters apply.
   * NOTE: Viator/Ammended is gone in v2 — "Amended Booking:" mail now files
   * straight into Viator/Modifications per your updated filter.
   */
  LABELS: {
    VIATOR_CONFIRM: 'Publishing Pages/Viator/Confirmations',
    VIATOR_CANCEL: 'Publishing Pages/Viator/Cancellations',
    VIATOR_MODIFY: 'Publishing Pages/Viator/Modifications',
    VIATOR_DONE: 'Publishing Pages/Viator/Done',
    VIATOR_CONVERSATIONS: 'Viator/Conversations',

    GURUWALK_CONFIRM: 'Publishing Pages/Guruwalk/Confirmations',
    GURUWALK_CANCEL: 'Publishing Pages/Guruwalk/Cancellations',
    GURUWALK_MODIFY: 'Publishing Pages/Guruwalk/Modifications',
    GURUWALK_DONE: 'Publishing Pages/Guruwalk/Done',

    AIRBNB_CONFIRM: 'Publishing Pages/Airbnb/Confirmations',
    AIRBNB_CANCEL: 'Publishing Pages/Airbnb/Cancellations',
    AIRBNB_DONE: 'Publishing Pages/Airbnb/Done',

    GYG_CONFIRM: 'Publishing Pages/GetYourGuide/Confirmations',
    GYG_CANCEL: 'Publishing Pages/GetYourGuide/Cancellations',
    GYG_MODIFY: 'Publishing Pages/GetYourGuide/Modifications',
    GYG_DONE: 'Publishing Pages/GetYourGuide/Done',

    FREETOUR_CONFIRM: 'Publishing Pages/FreeTour/Confirmations',
    FREETOUR_CANCEL: 'Publishing Pages/FreeTour/Cancellations',
    FREETOUR_MODIFY: 'Publishing Pages/FreeTour/Modifications',
    FREETOUR_DONE: 'Publishing Pages/FreeTour/Done',

    WEB_CONFIRM: 'Webpage/Confirmations',
    WEB_DONE: 'Webpage/Done',

    // Applied to a Confirmation, Modification OR Cancellation thread once it
    // has been read and fully acted on this run. Frequent (5-min) runs skip
    // any thread already carrying this label; the twice-daily audit ignores
    // it and re-reads everything. Never used to mean read/unread or archived
    // — it is purely "the algorithm has already considered this thread".
    PROCESSED: 'Publishing Pages/Processed'
  },

  SOURCE: {
    VIATOR: 'Viator',
    GURUWALK: 'Guruwalk',
    WEBSITE: 'Website',
    AIRBNB: 'Airbnb',
    GYG: 'GetYourGuide',
    FREETOUR: 'Free Tour'
  },

  /**
   * Guest accounting model per source.
   *   'paid' -> counted in "Number of Paid Guests" in Done Tours.
   *   'free' -> counted in "Number of Free Guests" in Done Tours.
   */
  MODEL: {
    Viator: 'paid',
    Airbnb: 'paid',
    GetYourGuide: 'paid',
    Guruwalk: 'free',
    'Free Tour': 'free',
    Website: 'free'
  },

  LANGUAGE: {
    ENGLISH: 'English',
    GERMAN: 'German',
    SPANISH: 'Spanish',
    ITALIAN: 'Italian',
    FRENCH: 'French'
  },

  DEFAULT_TIME: '11:00 AM',

  WEBSITE_WEEKDAY_TIME: '10:30 AM',
  WEBSITE_WEEKEND_TIME: '4:30 PM',

  DONE_AFTER_HOURS: 2,

  // Fast (5-min) runs only ever see Processed-labelled threads on repeat
  // passes, so they can afford a smaller cap. The audit run needs a bigger
  // cap to clear whatever backlog piled up. Raise MAX_THREADS_AUDIT further
  // if your Errors tab still shows quota errors after a few days on v2.
  MAX_THREADS_FAST: 20,
  MAX_THREADS_AUDIT: 60,

  MAX_RUN_MS: 180000,

  /* ============================================================
   * COMMISSIONS / INCOME MODEL  — edit these numbers only.
   * ============================================================ */

  GYG_COMMISSION: 0.25,
  FREE_COMMISSION_PER_GUEST: 6,
  GURUWALK_FEE_PER_BOOKING: 4.70,

  /* ============================================================
   * PRIVATE TOURS
   * ============================================================ */
  PRIVATE_TOUR_KEYWORDS: /\bprivate\b|tour privado|grupo privado|group up to|grupo de hasta/i,
  PRIVATE_TOUR_NOTE: 'Private',

  INTERNAL_ALERT_TO: 'rootsandroadstours@gmail.com',

  // INBOX POLICY
  //   A processed booking email STAYS IN THE INBOX until its tour is over.
  //   The inbox is therefore the live "upcoming tours" list: anything sitting
  //   there is a tour that still has to happen. It leaves the inbox only when
  //   the tour completes (start + DONE_AFTER_HOURS) or when it is cancelled.
  //   Cancelled bookings are archived immediately — that tour will never run,
  //   so leaving it in the inbox would be noise. Set this to false if you
  //   would rather keep cancellations in the inbox too.
  ARCHIVE_CANCELLED_IMMEDIATELY: true,

  // Hidden handoff tab: full detail of every completed booking, written just
  // before its row leaves the active tabs for the aggregated "Done Tours".
  // The guide portal's no-show queues read this (Done Tours has no booking
  // ids, so this is the only place a completed booking's id survives).
  COMPLETED_LOG_TAB: 'Completed Log',
  COMPLETED_LOG_KEEP_DAYS: 21
};


/**
 * Spanish month names -> 0-based month index. Used to parse GYG modification
 * dates like "16 de julio de 2026 a las 11:00" (confirmations use English dates).
 */
const RNR_SPANISH_MONTHS_ = {
  enero: 0, febrero: 1, marzo: 2, abril: 3, mayo: 4, junio: 5,
  julio: 6, agosto: 7, septiembre: 8, setiembre: 8, octubre: 9,
  noviembre: 10, diciembre: 11
};


/**
 * Central per-source configuration table.
 * Adding another OTA later = add one entry here + one parser function.
 * Loops below iterate this table instead of hard-coding each source.
 */
function sourceConfigs_() {
  return [
    {
      source: RNR.SOURCE.VIATOR,
      confirm: RNR.LABELS.VIATOR_CONFIRM,
      cancel: RNR.LABELS.VIATOR_CANCEL,
      modify: RNR.LABELS.VIATOR_MODIFY,
      done: RNR.LABELS.VIATOR_DONE
    },
    {
      source: RNR.SOURCE.GURUWALK,
      confirm: RNR.LABELS.GURUWALK_CONFIRM,
      cancel: RNR.LABELS.GURUWALK_CANCEL,
      modify: RNR.LABELS.GURUWALK_MODIFY,
      done: RNR.LABELS.GURUWALK_DONE
    },
    {
      source: RNR.SOURCE.AIRBNB,
      confirm: RNR.LABELS.AIRBNB_CONFIRM,
      cancel: RNR.LABELS.AIRBNB_CANCEL,
      done: RNR.LABELS.AIRBNB_DONE
    },
    {
      source: RNR.SOURCE.GYG,
      confirm: RNR.LABELS.GYG_CONFIRM,
      cancel: RNR.LABELS.GYG_CANCEL,
      modify: RNR.LABELS.GYG_MODIFY,
      done: RNR.LABELS.GYG_DONE
    },
    {
      source: RNR.SOURCE.FREETOUR,
      confirm: RNR.LABELS.FREETOUR_CONFIRM,
      cancel: RNR.LABELS.FREETOUR_CANCEL,
      modify: RNR.LABELS.FREETOUR_MODIFY,
      done: RNR.LABELS.FREETOUR_DONE
    }
  ];
}


/******************************************************
 * 2. RUN STATE + PER-RUN CACHES
 *
 * All caches are reset at the top of every run so one execution never reads
 * a given Gmail message more than once.
 ******************************************************/

let RNR_RUN_STARTED_AT_ = 0;
let RNR_RUN_STATS_ = { processed: 0, upserts: 0, errors: 0 };
let RNR_LABEL_CACHE_ = null;     // labelName -> GmailLabel (1 lookup per run)
let RNR_THREADS_CACHE_ = null;   // mode|labelName -> threads[] (1 fetch per run)
let RNR_TEXT_CACHE_ = null;              // messageId -> best text
let RNR_CANCEL_CACHE_ = null;            // array of cancellation bookings
let RNR_CONFIRMATION_CACHE_ = null;      // "source|id" -> confirmation booking
let RNR_CONFIRM_THREAD_INDEX_ = null;    // "source|id" -> Gmail thread
let RNR_MODIFY_THREAD_INDEX_ = null;     // "source|id" -> Gmail thread
let RNR_REINSTATED_IDS_ = null;          // Set of "source|id" rebooked this run

// When true, getThreadsSafe_ skips threads already carrying the PROCESSED
// label. Set true by the frequent run, false by the twice-daily audit.
var RNR_SKIP_PROCESSED_ = false;


function resetRunCaches_() {
  RNR_RUN_STATS_ = { processed: 0, upserts: 0, errors: 0 };
  RNR_LABEL_CACHE_ = new Map();
  RNR_THREADS_CACHE_ = new Map();
  RNR_TEXT_CACHE_ = new Map();
  RNR_CANCEL_CACHE_ = null;
  RNR_CONFIRMATION_CACHE_ = null;
  RNR_CONFIRM_THREAD_INDEX_ = null;
  RNR_MODIFY_THREAD_INDEX_ = null;
  RNR_REINSTATED_IDS_ = new Set();
}


/**
 * Mark a booking as reinstated for this run. A GetYourGuide rebooking
 * ("reprogramada") reuses the ORIGINAL booking code, so a stale cancellation
 * email for that same code must NOT delete the rebooked tour, and must NOT
 * cause its Confirm/Modify Gmail labels to be stripped to Cancel. Any id in
 * this set is treated as active for the rest of the run.
 */
function markReinstated_(source, bookingId) {
  const id = normalizeId_(bookingId);
  if (source && id && RNR_REINSTATED_IDS_) RNR_REINSTATED_IDS_.add(source + '|' + id);
}

function isReinstated_(source, bookingId) {
  const id = normalizeId_(bookingId);
  return Boolean(RNR_REINSTATED_IDS_ && source && id && RNR_REINSTATED_IDS_.has(source + '|' + id));
}


function runHasTimeLeft_() {
  return !RNR_RUN_STARTED_AT_ || (Date.now() - RNR_RUN_STARTED_AT_ < RNR.MAX_RUN_MS);
}


/******************************************************
 * 3. MAIN FUNCTIONS
 ******************************************************/

/**
 * FULL setup. Run once from the editor, or automatically on audit runs.
 * ensureLabels_ + formatSheets_ are Gmail/Spreadsheet-heavy, so the fast
 * 5-minute run does NOT call this (see runBookingCore_).
 */
function setupBookingSystem() {
  ensureSheets_();
  ensureLabels_();          // ~24 Gmail lookups — audit/manual only
  formatSheets_();
}

/** Cheap: create the tabs if missing. No Gmail calls. */
function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, RNR.SHEETS.ENGLISH, RNR.ACTIVE_HEADERS);
  ensureSheet_(ss, RNR.SHEETS.GERMAN, RNR.ACTIVE_HEADERS);
  ensureSheet_(ss, RNR.SHEETS.SPANISH, RNR.ACTIVE_HEADERS);
  ensureSheet_(ss, RNR.SHEETS.ITALIAN, RNR.ACTIVE_HEADERS);
  ensureSheet_(ss, RNR.SHEETS.FRENCH, RNR.ACTIVE_HEADERS);
  ensureSheet_(ss, RNR.SHEETS.DONE, RNR.DONE_HEADERS);
  ensureSheet_(ss, RNR.SHEETS.ERRORS, RNR.ERROR_HEADERS);
}


/**
 * Trigger entry points must NEVER throw. Apps Script emails the owner a
 * "Summary of failures" for every trigger run that ends in an exception, and a
 * transient "Service Spreadsheets timed out while accessing document" is normal
 * under contention (someone editing the sheet, or the Control project reading
 * it at the same moment). Those are logged to the Errors tab and swallowed —
 * the next run, 5 minutes later, simply picks the work up again.
 */
function safeTriggerRun_(label, fn) {
  try {
    fn();
  } catch (err) {
    try { logError_(label + ' failed (transient; swallowed so the trigger does not email)', err, ''); } catch (e) { /* logging must never throw either */ }
    console.log(label + ' error: ' + (err && err.stack ? err.stack : err));
  }
}

// Frequent trigger (every 5 min): skips already-processed threads.
function runBookingSystem() { safeTriggerRun_('runBookingSystem', function () { runBookingCore_(true); }); }

// Twice-daily audit trigger: identical pipeline, ignores PROCESSED so every
// labelled thread is re-read for full consistency.
function runBookingAudit() { safeTriggerRun_('runBookingAudit', function () { runBookingCore_(false); }); }


function runBookingCore_(skipProcessed) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    console.log('Skipped: another run is already active.');
    return;
  }

  RNR_RUN_STARTED_AT_ = Date.now();
  resetRunCaches_();
  RNR_SKIP_PROCESSED_ = !!skipProcessed;

  try {
    // Fast (5-min) runs must be Gmail-cheap: only create tabs. Labels already
    // exist, and getThreadsSafe_ finds unprocessed mail via Gmail SEARCH
    // (label:… -label:…-processed) instead of listing+inspecting every thread.
    // The audit does the full, heavier setup a few times a day.
    if (skipProcessed) {
      ensureSheets_();
    } else {
      setupBookingSystem();
      ensureProcessedLabel_();
    }

    // ONE complete, idempotent pass. Every labelled thread is read every run
    // (read/unread and inbox state are NEVER used as a signal). Re-processing
    // is safe because upsertActiveBooking_ updates in place instead of
    // duplicating, and every Gmail mutation is itself idempotent.

    // 1. Cancellations first (nothing cancelled should survive). This pass
    //    only removes sheet rows and finalises the CANCELLATION thread
    //    itself (Processed + archive); it deliberately does NOT relabel the
    //    matching Confirm/Modify threads yet — see step 4, which runs after
    //    modifications so a same-run GYG reinstatement is already known and
    //    can't be wrongly cancelled.
    if (runHasTimeLeft_()) processCancellations_();

    // 2. Confirmations + modifications.
    if (runHasTimeLeft_()) processConfirmations_();
    if (runHasTimeLeft_()) processModifications_();

    // 3. Consistency: relabel any Confirm/Modify thread whose booking id is
    //    now known-cancelled (and not reinstated this run) to Cancel, drop
    //    any active row that has a cancellation email, and run cancellations
    //    once more to catch anything newly exposed.
    if (runHasTimeLeft_()) reconcileCancelledThreadLabels_();
    if (runHasTimeLeft_()) removeActiveBookingsThatHaveCancellationEmails_();
    if (runHasTimeLeft_()) processCancellations_();

    // 4. Finished tours -> Done. Sheet rows move every run (cheap, Sheets
    //    only). The Gmail side (archive + relabel to Done) runs every run too
    //    but only inspects the small "still in inbox" candidate set on fast
    //    runs (see getThreadsForCompletionSweep_).
    let completedNow = [];
    if (runHasTimeLeft_()) completedNow = moveCompletedBookingRowsToDone_() || [];
    if (runHasTimeLeft_()) moveCompletedGmailThreadsToDone_(completedNow);

    // 4b. INVARIANT CHECKS (audits only): duplicates, cancelled-still-active,
    //     completed-still-active, invalid rows. Findings land in Errors
    //     (deduped) — the daily self-test email surfaces them.
    if (!RNR_SKIP_PROCESSED_ && runHasTimeLeft_()) checkInvariants_();

    // 5. Tidy the sheets — a full rewrite of every tab, so only when
    //    something actually changed this run, or on audits.
    const dirty = RNR_RUN_STATS_.upserts > 0 || !RNR_SKIP_PROCESSED_;
    if (dirty && runHasTimeLeft_()) dedupeActiveSheets_();
    if (dirty && runHasTimeLeft_()) sortActiveSheets_();
    if (dirty && runHasTimeLeft_()) sortDoneSheet_();

  } catch (err) {
    // Errors are logged, never rethrown, so Google does not email failure alerts.
    logError_('runBookingSystem error', err, '');
    console.log(String(err && err.stack ? err.stack : err));

  } finally {
    writeRunStatus_(skipProcessed ? 'fast (5-min)' : 'audit (full re-read)');
    lock.releaseLock();
  }
}


/**
 * Heartbeat: rewrites the small "Status" tab after EVERY run so a manager can
 * see at a glance that the system is alive and what the last run did.
 */
function writeRunStatus_(mode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('Status');
    if (!sh) sh = ss.insertSheet('Status');
    const rows = [
      ['Last run finished', Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm:ss')],
      ['Mode', mode],
      ['Duration (seconds)', Math.round((Date.now() - RNR_RUN_STARTED_AT_) / 1000)],
      ['Threads marked Processed this run', RNR_RUN_STATS_.processed],
      ['Booking rows written/updated this run', RNR_RUN_STATS_.upserts],
      ['Errors logged this run', RNR_RUN_STATS_.errors],
      ['', ''],
      ['How to read this', 'This tab refreshes after every run (every 5 min + audits). ' +
        'If "Last run finished" is more than ~10 minutes old, the trigger is not running: ' +
        'open Apps Script > Triggers and > Executions. Error details: Errors tab. ' +
        'Full diagnosis: run systemStatus() in the editor.']
    ];
    sh.clear();
    // Column B as TEXT first: otherwise Sheets turns the timestamp into a Date
    // object and anything reading it back gets "Mon Jul 20 2026 11:57:44 GMT…"
    // instead of a parseable stamp (this is what made the health check read
    // "NaN min ago").
    sh.getRange(1, 2, rows.length, 1).setNumberFormat('@');
    sh.getRange(1, 1, rows.length, 2).setValues(rows);
    sh.getRange(1, 1, rows.length, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 280);
    sh.setColumnWidth(2, 520);
  } catch (e) { console.log('writeRunStatus_: ' + e); }
}


/** Create the PROCESSED label once so safeAddLabel_ can apply it. */
function ensureProcessedLabel_() {
  try {
    if (!GmailApp.getUserLabelByName(RNR.LABELS.PROCESSED)) {
      GmailApp.createLabel(RNR.LABELS.PROCESSED);
    }
  } catch (e) { console.log('ensureProcessedLabel_: ' + e); }
}


/**
 * Website booking endpoint.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    // A long booking run holds the lock. NEVER lose the reservation: log it
    // and email the raw details to management for manual entry.
    try {
      const raw = JSON.stringify(e && e.parameter ? e.parameter : {});
      logError_('Website booking arrived while system busy — NOT saved automatically', 'manual entry needed', raw);
      MailApp.sendEmail({
        to: RNR.INTERNAL_ALERT_TO,
        subject: 'R&R: WEBSITE BOOKING NEEDS MANUAL ENTRY (system was busy)',
        body: 'Add this reservation to the BookingSheet by hand:\n\n' +
              JSON.stringify(e && e.parameter ? e.parameter : {}, null, 2)
      });
    } catch (err2) { /* nothing more we can do */ }
    return textResponse_('BUSY');
  }

  RNR_RUN_STARTED_AT_ = Date.now();
  resetRunCaches_();

  try {
    ensureSheets_();   // cheap; full setup/formatting belongs to the audit run

    const params = e && e.parameter ? e.parameter : {};
    const booking = websiteParamsToBooking_(params);

    if (!isValidBooking_(booking)) {
      throw new Error('Invalid website booking');
    }

    upsertActiveBooking_(booking);   // inserts in sorted position

    sendWebsiteReservationAlert_(booking, params);       // internal alert
    sendWebsiteConfirmationViaBrevo_(booking, params);   // customer confirmation

    return textResponse_('OK');

  } catch (err) {
    logError_(
      'Website booking error',
      err,
      JSON.stringify(e && e.parameter ? e.parameter : {})
    );
    return textResponse_('ERROR');

  } finally {
    lock.releaseLock();
  }
}


/******************************************************
 * 4. SHEET SETUP AND FORMATTING
 ******************************************************/

function activeSheetNames_() {
  return [RNR.SHEETS.ENGLISH, RNR.SHEETS.GERMAN, RNR.SHEETS.SPANISH,
          RNR.SHEETS.ITALIAN, RNR.SHEETS.FRENCH];
}


function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(name);

  // Only touch the header row when it is actually wrong. This runs for EVERY
  // active tab on EVERY 5-minute run; rewriting headers unconditionally was
  // pure Spreadsheets load (and load is what triggers "Service Spreadsheets
  // timed out" under contention).
  if (!isNew) {
    const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    if (headers.every((h, i) => String(current[i]) === String(h))) return;
  }
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
}


function ensureLabels_() {
  Object.values(RNR.LABELS).forEach(labelName => {
    try {
      if (!GmailApp.getUserLabelByName(labelName)) {
        GmailApp.createLabel(labelName);
      }
    } catch (e) {
      // Label creation can fail transiently; do not abort setup.
      console.log('ensureLabels_ skip ' + labelName + ': ' + e);
    }
  });
}


function formatSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  activeSheetNames_().forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    sh.getRange('A:I').setFontFamily('Arial').setFontSize(10);
    sh.getRange('A1:I1')
      .setFontWeight('bold')
      .setWrap(true)
      .setVerticalAlignment('middle');

    sh.setRowHeight(1, 36);

    // Phone column must be text, otherwise +phone is read as a formula.
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('C:C').setNumberFormat('0');
    sh.getRange('D:D').setNumberFormat('ddd, mmm d');
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('0.##');
    sh.getRange('H:H').setNumberFormat('@');
    sh.getRange('I:I').setWrap(true);

    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 140);
    sh.setColumnWidth(3, 85);
    sh.setColumnWidth(4, 125);
    sh.setColumnWidth(5, 85);
    sh.setColumnWidth(6, 120);
    sh.setColumnWidth(7, 85);
    sh.setColumnWidth(8, 150);
    sh.setColumnWidth(9, 260);
  });

  const done = ss.getSheetByName(RNR.SHEETS.DONE);
  if (done) {
    done.getRange('A:H').setFontFamily('Arial').setFontSize(10);
    done.getRange('A1:H1')
      .setFontWeight('bold')
      .setWrap(true)
      .setVerticalAlignment('middle');

    done.setRowHeight(1, 36);

    done.getRange('A:A').setNumberFormat('ddd, mmm d');
    done.getRange('B:B').setNumberFormat('@');
    done.getRange('D:E').setNumberFormat('0');
    done.getRange('F:F').setNumberFormat('0.##');

    done.setColumnWidth(1, 125);
    done.setColumnWidth(2, 85);
    done.setColumnWidth(3, 95);
    done.setColumnWidth(4, 95);
    done.setColumnWidth(5, 95);
    done.setColumnWidth(6, 90);
    done.setColumnWidth(7, 130);
    done.setColumnWidth(8, 300);
  }

  const errors = ss.getSheetByName(RNR.SHEETS.ERRORS);
  if (errors) {
    errors.getRange('A:D').setFontFamily('Arial').setFontSize(10);
    errors.getRange('A1:D1')
      .setFontWeight('bold')
      .setWrap(true)
      .setVerticalAlignment('middle');

    errors.setFrozenRows(1);
    errors.setColumnWidth(1, 150);
    errors.setColumnWidth(2, 180);
    errors.setColumnWidth(3, 500);
    errors.setColumnWidth(4, 400);
  }
}


/******************************************************
 * 5. SAFE GMAIL WRAPPERS  (the "operation not allowed" fix)
 *
 * A single problem thread must never abort an entire run. Every Gmail
 * mutation and every batch read goes through these guards.
 ******************************************************/

/**
 * QUOTA MODEL (the fix for "Service invoked too many times for one day: gmail")
 *
 * The old approach fetched every thread under every label and then read each
 * thread's labels to decide whether to skip it — hundreds of Gmail calls per
 * run, ~288 runs/day, quota dead by afternoon.
 *
 * Now:
 *   FAST run  -> ONE GmailApp.search per label:
 *                  label:<x> -label:publishing-pages-processed
 *                Gmail itself returns only the not-yet-processed threads.
 *                An idle 5-minute run costs ~15 Gmail calls total.
 *   AUDIT run -> direct label.getThreads (complete, no filtering needed
 *                because audits deliberately re-read everything).
 *   Every label object and every thread list is cached per run.
 */
function getThreadsSafe_(labelName) {
  try {
    if (!RNR_THREADS_CACHE_) RNR_THREADS_CACHE_ = new Map();
    const key = (RNR_SKIP_PROCESSED_ ? 'F|' : 'A|') + labelName;
    if (RNR_THREADS_CACHE_.has(key)) return RNR_THREADS_CACHE_.get(key);

    let threads;
    if (RNR_SKIP_PROCESSED_) {
      const q = searchTokenForLabel_(labelName) + ' -' + searchTokenForLabel_(RNR.LABELS.PROCESSED);
      threads = GmailApp.search(q, 0, RNR.MAX_THREADS_FAST) || [];
    } else {
      const label = getLabel_(labelName);
      threads = label ? (label.getThreads(0, RNR.MAX_THREADS_AUDIT) || []) : [];
    }
    RNR_THREADS_CACHE_.set(key, threads);
    return threads;
  } catch (e) {
    console.log('getThreadsSafe_ ' + labelName + ': ' + e);
    return [];
  }
}

/**
 * Threads to consider for "tour finished -> archive + Done label".
 *
 * Since processed bookings now STAY in the inbox until their tour is over,
 * the set of candidates is exactly "still in the inbox under this label" —
 * a small, cheap search (one call per label) that keeps completion timely on
 * the 5-minute run. The audit still walks the whole label so nothing that was
 * archived by hand or by an older version is missed.
 */
function getThreadsForCompletionSweep_(labelName) {
  try {
    if (!RNR_THREADS_CACHE_) RNR_THREADS_CACHE_ = new Map();
    const key = (RNR_SKIP_PROCESSED_ ? 'CF|' : 'CA|') + labelName;
    if (RNR_THREADS_CACHE_.has(key)) return RNR_THREADS_CACHE_.get(key);

    let threads;
    if (RNR_SKIP_PROCESSED_) {
      threads = GmailApp.search(searchTokenForLabel_(labelName) + ' in:inbox', 0, RNR.MAX_THREADS_FAST) || [];
    } else {
      const label = getLabel_(labelName);
      threads = label ? (label.getThreads(0, RNR.MAX_THREADS_AUDIT) || []) : [];
    }
    RNR_THREADS_CACHE_.set(key, threads);
    return threads;
  } catch (e) {
    console.log('getThreadsForCompletionSweep_ ' + labelName + ': ' + e);
    return [];
  }
}

/** Gmail search canonical form: "Publishing Pages/Viator/Confirmations"
 *  -> label:publishing-pages-viator-confirmations  (spaces and / become -). */
function searchTokenForLabel_(labelName) {
  return 'label:' + String(labelName || '').toLowerCase().replace(/[\s/]+/g, '-');
}

/** Cached label lookup: each label name costs at most ONE Gmail call per run. */
function getLabel_(labelName) {
  if (!RNR_LABEL_CACHE_) RNR_LABEL_CACHE_ = new Map();
  if (RNR_LABEL_CACHE_.has(labelName)) return RNR_LABEL_CACHE_.get(labelName);
  let label = null;
  try { label = GmailApp.getUserLabelByName(labelName); } catch (e) { /* ignore */ }
  RNR_LABEL_CACHE_.set(labelName, label);
  return label;
}

function threadHasProcessedLabel_(thread) {
  try {
    const labels = thread.getLabels() || [];
    for (let i = 0; i < labels.length; i++) {
      if (labels[i].getName() === RNR.LABELS.PROCESSED) return true;
    }
  } catch (e) { /* ignore */ }
  return false;
}

function threadHasLabel_(thread, labelName) {
  try {
    const labels = thread.getLabels() || [];
    for (let i = 0; i < labels.length; i++) {
      if (labels[i].getName() === labelName) return true;
    }
  } catch (e) { /* ignore */ }
  return false;
}

/**
 * Unconditional read of every thread under a label, ignoring PROCESSED.
 * Cached per run. Only used where full visibility matters (GYG Done lookup).
 */
function getThreadsForce_(labelName) {
  try {
    if (!RNR_THREADS_CACHE_) RNR_THREADS_CACHE_ = new Map();
    const key = 'X|' + labelName;
    if (RNR_THREADS_CACHE_.has(key)) return RNR_THREADS_CACHE_.get(key);
    const cap = RNR_SKIP_PROCESSED_ ? RNR.MAX_THREADS_FAST : RNR.MAX_THREADS_AUDIT;
    const label = getLabel_(labelName);
    const threads = label ? (label.getThreads(0, cap) || []) : [];
    RNR_THREADS_CACHE_.set(key, threads);
    return threads;
  } catch (e) {
    console.log('getThreadsForce_ ' + labelName + ': ' + e);
    return [];
  }
}

function safeMarkRead_(thread) {
  try { if (thread) thread.markRead(); } catch (e) { /* ignore */ }
}

function safeArchive_(thread) {
  try { if (thread) thread.moveToArchive(); } catch (e) { /* ignore */ }
}

function safeAddLabel_(thread, labelName) {
  try {
    const label = getLabel_(labelName);
    if (thread && label) thread.addLabel(label);
  } catch (e) { /* ignore */ }
}

function safeRemoveLabel_(thread, labelName) {
  try {
    const label = getLabel_(labelName);
    if (thread && label) thread.removeLabel(label);
  } catch (e) { /* ignore */ }
}

function moveThreadOutOfInbox_(thread) {
  safeMarkRead_(thread);
  safeArchive_(thread);
}

/**
 * Finalise a thread as Processed: add the marker label, archive it. Used
 * after a Confirmation/Modification/Cancellation thread has been fully read
 * and acted on this run, regardless of which label it keeps.
 */
/**
 * Mark a thread as fully handled by the algorithm.
 *
 * IMPORTANT: this does NOT archive. A processed booking email stays in the
 * inbox until its tour is over (see RNR.ARCHIVE_CANCELLED_IMMEDIATELY and
 * moveCompletedPlatformThreadsToDone_). The inbox is the live list of
 * upcoming tours; the Processed label is what tells the algorithm — and you —
 * that the booking is already in the sheet.
 */
function finalizeThreadProcessed_(thread) {
  safeAddLabel_(thread, RNR.LABELS.PROCESSED);
  safeMarkRead_(thread);
  if (RNR_RUN_STATS_) RNR_RUN_STATS_.processed++;
}

/**
 * Same as above, but for a thread whose booking is CANCELLED: that tour will
 * never run, so it also leaves the inbox straight away.
 */
function finalizeThreadCancelled_(thread) {
  safeAddLabel_(thread, RNR.LABELS.PROCESSED);
  safeMarkRead_(thread);
  if (RNR.ARCHIVE_CANCELLED_IMMEDIATELY) safeArchive_(thread);
  if (RNR_RUN_STATS_) RNR_RUN_STATS_.processed++;
}

/**
 * Strip every RNR-managed label this source uses from a thread except the
 * ones listed in keep[]. Used when a booking is cancelled (keep = [cancel])
 * or completed (keep = [done], i.e. strip everything including Processed).
 */
function stripSourceLabelsExcept_(thread, cfg, keep) {
  const all = [cfg.confirm, cfg.modify, cfg.cancel, RNR.LABELS.PROCESSED].filter(Boolean);
  all.forEach(labelName => {
    if (keep.indexOf(labelName) === -1) safeRemoveLabel_(thread, labelName);
  });
  keep.forEach(labelName => { if (labelName) safeAddLabel_(thread, labelName); });
}


/******************************************************
 * 6. WEBSITE BOOKINGS
 ******************************************************/

function websiteParamsToBooking_(p) {
  const date = normalizeDate_(p.tour_date);
  const time = normalizeTime_(p.tour_time);
  const language = normalizeLanguage_(p.language);

  return normalizeBooking_({
    name: p.name,
    phone: cleanPhone_(p.phone),
    guests: Number(p.guests || 1),
    date,
    time,
    language,
    source: RNR.SOURCE.WEBSITE,
    income: 0,
    bookingId: generateWebsiteBookingId_(date, p.name, p.phone, p.guests),
    notes: p.message,
    hasExplicitGuests: true,
    hasExplicitDate: true,
    hasExplicitTime: true,
    hasExplicitIncome: true
  });
}




function generateWebsiteBookingId_(date, name, phone, guests) {
  const d = dateKey_(date) || Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');

  const base = [
    d,
    cleanText_(name),
    cleanPhoneKey_(phone),
    String(guests || 1),
    String(new Date().getTime())
  ].join('|');

  const hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, base)
  ).replace(/[^A-Z0-9]/gi, '').slice(0, 9).toUpperCase();

  return 'RR' + hash;
}


/**
 * Deterministic fallback booking id (no timestamp) for sources whose emails
 * may not carry a stable platform reference (e.g. some Free Tour notifications).
 * Because it is stable, re-processing the same booking dedupes cleanly.
 */
function generateFallbackBookingId_(prefix, date, name, phone, guests) {
  const base = [
    prefix,
    dateKey_(date),
    normalizeNameKey_(name),
    cleanPhoneKey_(phone),
    String(guests || 1)
  ].join('|');

  const hash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, base)
  ).replace(/[^A-Z0-9]/gi, '').slice(0, 10).toUpperCase();

  return prefix + hash;
}


function sendWebsiteConfirmationViaBrevo_(booking, params) {
  const email = cleanText_(params.email);
  if (!email) throw new Error('Missing website customer email');

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('BREVO_API_KEY');
  const templateId = Number(props.getProperty('BREVO_TEMPLATE_ID'));

  if (!apiKey || !templateId) throw new Error('Brevo API key or template ID missing');

  const payload = {
    to: [{ email: email, name: booking.name }],
    templateId: templateId,
    params: {
      NAME: booking.name || '',
      TOUR_DATE: booking.date
        ? Utilities.formatDate(booking.date, 'Europe/Madrid', 'EEE, d MMM yyyy')
        : '',
      TOUR_TIME: normalizeTime_(booking.time),
      BOOKING_ID: booking.bookingId || '',
      GUESTS: String(booking.guests || ''),
      MESSAGE: String(params.message || '')
    }
  };

  const res = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() >= 300) {
    throw new Error('Brevo error: ' + res.getContentText());
  }
}


function sendWebsiteReservationAlert_(booking, params) {
  const subject = 'NEW WEBSITE RESERVATION - Roots & Roads';

  const body =
    'A new website reservation has been received.\n\n' +
    'Language: ' + safeString_(booking.language) + '\n' +
    'Name: ' + safeString_(booking.name) + '\n' +
    'Email: ' + safeString_(params.email) + '\n' +
    'Phone: ' + safeString_(booking.phone) + '\n' +
    'Guests: ' + safeString_(booking.guests) + '\n' +
    'Tour date: ' + (
      booking.date ? Utilities.formatDate(booking.date, 'Europe/Madrid', 'EEE, d MMM yyyy') : ''
    ) + '\n' +
    'Time: ' + safeString_(booking.time) + '\n' +
    'Source: ' + safeString_(booking.source) + '\n' +
    'Booking ID: ' + safeString_(booking.bookingId) + '\n\n' +
    'Message:\n' + safeString_(params.message);

  MailApp.sendEmail({ to: RNR.INTERNAL_ALERT_TO, subject, body });
}


/******************************************************
 * 7. CONFIRMATIONS
 ******************************************************/

function processConfirmations_() {
  processConfirmationLabel_(RNR.LABELS.VIATOR_CONFIRM, RNR.SOURCE.VIATOR);
  processConfirmationLabel_(RNR.LABELS.GURUWALK_CONFIRM, RNR.SOURCE.GURUWALK);
  processConfirmationLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG);
  processConfirmationLabel_(RNR.LABELS.FREETOUR_CONFIRM, RNR.SOURCE.FREETOUR);

  // Airbnb forwarded emails may already be read, so process them separately
  // (harmless: this system never branches on read/unread anyway).
  processConfirmationLabel_(RNR.LABELS.AIRBNB_CONFIRM, RNR.SOURCE.AIRBNB);
}


/**
 * True when a thread filed under Confirmations is really a modification or a
 * cancellation. Gmail groups a booking's confirmation, modification and
 * cancellation into ONE conversation, and OTA subjects overlap ("Booking detail
 * change: - S779080 - GYG…" also matches the broad "Booking -" filter).
 *
 * Classified from the TEXT, so it still works when the message alone cannot
 * produce a complete booking — resolving it is the modification pass's job, not
 * the confirmation pass's. Same wording the parsers route on.
 */
function threadIsModifyOrCancel_(thread) {
  try {
    return thread.getMessages().some(m => {
      const s = (m.getSubject() || '') + '\n' + getBestMessageText_(m);
      return /booking detail change|cambio en los datos|reprogramad|rebooked|ha vuelto a reservar|se ha modificado|modificad|amend|updated booking/i.test(s) ||
             /cancel(?:led|ed|lation)?|cancelad|cancelaci[oó]n|anulad/i.test(s);
    });
  } catch (e) {
    return false;   // unreadable thread -> fall back to the parse-based check
  }
}


function processConfirmationLabel_(labelName, source) {
  const threads = getThreadsSafe_(labelName);

  for (const thread of threads) {
    if (!runHasTimeLeft_()) break;

    try {
      const bookings = uniqueBookings_(parseThread_(thread, source, 'confirm'));

      if (!bookings.length) {
        // A thread under Confirmations can legitimately contain NO
        // confirmation:
        //   - Gmail groups a booking's confirmation, modification and
        //     cancellation into ONE conversation, so the thread carries
        //     several of our labels at once;
        //   - OTA subjects overlap ("Booking detail change: - S779080 - GYG…"
        //     also matches the broad "Booking -" confirmation filter).
        // Those threads are owned by the modification / cancellation passes.
        // Skip them SILENTLY and, critically, do NOT mark them Processed here
        // — those passes run after this one and select threads with the same
        // "not yet Processed" search, so finalising now would make them skip
        // the thread entirely.
        // Ownership is decided by CLASSIFICATION as well as by a full parse: a
        // "Booking detail change" carries only the reference + the new date, so
        // it can never yield a complete booking on its own (the modification
        // pass looks the original up by id). Without this it was logged as a
        // bogus "Confirmation parse failed".
        const ownedElsewhere = parseThread_(thread, source, 'any').length > 0 ||
                               threadIsModifyOrCancel_(thread);
        if (ownedElsewhere) continue;

        // Nothing parsed in ANY mode -> a genuine parser/misfile problem.
        // Not marked Processed, so the audit keeps retrying it. Logged on
        // audits only (a fast-run log would add a row every 5 minutes), and
        // logError_ deduplicates repeats.
        if (!RNR_SKIP_PROCESSED_) {
          logError_('Confirmation parse failed (no booking found)',
            'Subject: ' + (thread.getFirstMessageSubject() || ''), labelName);
        }
        continue;
      }

      for (const booking of bookings) {
        if (!isValidBooking_(booking)) continue;
        if (isCompleted_(booking)) continue;

        // Fail-safe: if this booking also appears in a cancellation email,
        // do not (re)add it to the active sheet.
        if (isBookingCancelledByEmail_(booking)) continue;

        // Confirmations INSERT only. Once the row exists, management edits
        // are authoritative; audits re-reading this email will not revert
        // them. Changes flow in through modification emails only.
        upsertActiveBooking_(booking, false);
      }

      // Parsed OK -> mark considered. Confirm label itself is left in
      // place; it only changes on cancel or done.
      finalizeThreadProcessed_(thread);

    } catch (e) {
      // Log-and-continue: never let one thread kill the whole label pass.
      logError_('processConfirmationLabel_ ' + source, e, labelName);
    }
  }
}


/******************************************************
 * 8. CANCELLATIONS
 ******************************************************/

function processCancellations_() {
  sourceConfigs_().forEach(cfg => {
    if (!cfg.cancel) return;
    if (!runHasTimeLeft_()) return;
    processCancellationLabel_(cfg.cancel, cfg.source);
  });
}


function processCancellationLabel_(labelName, source) {
  // Fast runs skip Cancel threads already marked Processed (nothing left to
  // do for a cancellation that was already fully applied last run); the
  // audit re-reads all of them. New cancellation emails are never Processed
  // yet, so they are always seen on the very next run of either kind.
  const threads = getThreadsSafe_(labelName);

  for (const thread of threads) {
    if (!runHasTimeLeft_()) break;

    try {
      const cancellations = cancellationBookingsFromThread_(thread, source);

      for (const cancelled of cancellations) {
        if (!cancelled.bookingId && !cancelled.name) continue;

        // A booking rebooked THIS run (GYG "reprogramada" reuses the same
        // code) must survive a stale cancellation email for that code.
        if (isReinstated_(source, cancelled.bookingId)) continue;

        removeActiveBooking_(cancelled);
        if (cancelled.bookingId) {
          removeActiveBookingBySourceAndId_(source, cancelled.bookingId);
        }
      }

      // Even if the thread carried no usable cancellation signal, finalise
      // it: a cancellation-labelled thread should not create manual work.
      // Cancelled => leaves the inbox now (that tour will never run).
      // NOTE: this only finalises the CANCELLATION thread itself (mark
      // Processed + archive). Relabelling the matching Confirm/Modify
      // threads to Cancel happens later in reconcileCancelledThreadLabels_,
      // after modifications have run, so a same-run GYG reinstatement is
      // already known and can't be wrongly cancelled.
      finalizeThreadCancelled_(thread);

    } catch (e) {
      logError_('processCancellationLabel_ ' + source, e, labelName);
    }
  }
}


/******************************************************
 * 8B. AUTOMATIC CANCELLATION FAIL-SAFES
 *
 * No new tabs, no new labels, no unread dependency, no manual review queue.
 ******************************************************/



/** Sources that have a cancellation label. */
function cancellationConfigs_() {
  return sourceConfigs_().filter(cfg => cfg.cancel);
}


/**
 * Build (once per run) the list of every booking that appears under any
 * cancellation label. Reused by all fail-safe checks below. Respects the
 * same Processed-skip as everything else on fast runs — once a cancellation
 * has been fully applied (sheet row gone, Confirm/Modify threads relabelled)
 * there is nothing left for it to protect against until something changes,
 * and the audit run re-scans everything regardless.
 */
function getCancellationCache_() {
  if (RNR_CANCEL_CACHE_) return RNR_CANCEL_CACHE_;

  const items = [];

  cancellationConfigs_().forEach(cfg => {
    const threads = getThreadsSafe_(cfg.cancel);

    threads.forEach(thread => {
      if (!runHasTimeLeft_()) return;
      try {
        cancellationBookingsFromThread_(thread, cfg.source).forEach(b => {
          items.push(normalizeBooking_({
            ...b,
            source: cfg.source,
            emailSubject: thread.getFirstMessageSubject()
          }));
        });
      } catch (e) {
        logError_('getCancellationCache_ ' + cfg.source, e, cfg.cancel);
      }
    });
  });

  RNR_CANCEL_CACHE_ = items;
  return RNR_CANCEL_CACHE_;
}


function cancellationBookingsFromThread_(thread, source) {
  const parsed = uniqueBookings_(parseThread_(thread, source, 'cancel'));
  const out = [];

  parsed.forEach(b => {
    const n = normalizeBooking_({ ...b, source });
    if (n.bookingId || n.name) out.push(n);
  });

  if (out.length) return out;

  // Fallback for Gmail-threaded emails where the visible subject still says
  // "New Booking" but the thread carries a cancellation-labelled message.
  // Only trust the id if the thread actually contains cancellation language.
  if (threadHasCancellationSignal_(thread)) {
    const bookingId = extractBookingIdFromThread_(thread, source);

    if (bookingId) {
      out.push(normalizeBooking_({
        bookingId,
        source,
        name: '',
        phone: '',
        guests: 1,
        date: new Date(),
        time: RNR.DEFAULT_TIME,
        language: RNR.LANGUAGE.ENGLISH,
        hasExplicitGuests: false,
        hasExplicitDate: false,
        hasExplicitTime: false,
        hasExplicitIncome: false
      }));
    }
  }

  return out;
}


/**
 * Cancellation keywords, English + Spanish (GYG/Free Tour arrive in Spanish here).
 */
function threadHasCancellationSignal_(thread) {
  const subject = thread.getFirstMessageSubject() || '';
  const body = thread.getMessages()
    .map(m => (m.getSubject() || '') + '\n' + getBestMessageText_(m))
    .join('\n');

  return /cancelled|canceled|cancellation|acknowledge this cancellation|has cancelled a booking|has canceled a booking|cancelad|cancelaci[oó]n|anulad|reserva anulada/i
    .test(subject + '\n' + body);
}


function isBookingCancelledByEmail_(booking) {
  const b = normalizeBooking_(booking);
  if (!b.source) return false;

  // A rebooked (reinstated) booking is active again despite any old
  // cancellation email that reused the same booking code.
  if (isReinstated_(b.source, b.bookingId)) return false;

  return getCancellationCache_().some(c => cancellationMatchesBooking_(c, b));
}


function cancellationMatchesBooking_(cancelled, booking) {
  const C = normalizeBooking_(cancelled);
  const B = normalizeBooking_(booking);

  if (!C.source || !B.source || C.source !== B.source) return false;

  const cId = normalizeId_(C.bookingId);
  const bId = normalizeId_(B.bookingId);

  // Strongest signal: same platform id.
  if (cId && bId && cId === bId) return true;

  // Conservative fallback on human data.
  const sameDate = C.dateKey && B.dateKey && C.dateKey === B.dateKey;
  const sameTime = normalizeTime_(C.time) && normalizeTime_(B.time) &&
                   normalizeTime_(C.time) === normalizeTime_(B.time);
  const samePhone = C.phoneKey && B.phoneKey && C.phoneKey === B.phoneKey;
  const compatibleName = namesCompatible_(C.name, B.name);

  return sameDate && sameTime && (samePhone || compatibleName);
}


function removeActiveBookingsThatHaveCancellationEmails_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  activeSheetNames_().forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const deleteRows = [];

    rows.forEach((row, i) => {
      const b = rowToBooking_(row, sheetName);
      if (isBookingCancelledByEmail_(b)) deleteRows.push(i + 2);
    });

    deleteRows.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
  });
}


/**
 * For every source, look at every Confirm AND Modify thread. If a thread's
 * booking id is known-cancelled (present in the cancellation cache) AND has
 * NOT been reinstated this run, strip its Confirm/Modify label, add Cancel,
 * mark Processed, and archive. This is the step that makes "once cancelled,
 * pull the Confirm/Modify labels and file everything under Cancel" actually
 * happen — it runs AFTER processModifications_ so a same-run GYG
 * reschedule (markReinstated_) is already known and is correctly skipped.
 */
function reconcileCancelledThreadLabels_() {
  cancellationConfigs_().forEach(cfg => {
    const seen = new Set();

    [cfg.confirm, cfg.modify].filter(Boolean).forEach(labelName => {
      const threads = getThreadsSafe_(labelName);

      threads.forEach(thread => {
        if (!runHasTimeLeft_()) return;

        try {
          const threadKey = thread.getId();
          if (seen.has(threadKey)) return; // a thread could carry both labels

          const bookings = uniqueBookings_(parseThread_(thread, cfg.source, 'any'));
          const ids = bookings.map(b => normalizeId_(b.bookingId)).filter(Boolean);

          const threadId = normalizeId_(extractBookingIdFromThread_(thread, cfg.source));
          if (threadId) ids.push(threadId);

          const cancelledId = ids.find(id =>
            !isReinstated_(cfg.source, id) &&
            getCancellationCache_().some(c => c.source === cfg.source && normalizeId_(c.bookingId) === id)
          );

          if (cancelledId) {
            seen.add(threadKey);
            stripSourceLabelsExcept_(thread, cfg, [cfg.cancel]);
            finalizeThreadCancelled_(thread);
          }
        } catch (e) {
          logError_('reconcileCancelledThreadLabels_ ' + cfg.source, e, labelName);
        }
      });
    });
  });
}


function removeActiveBookingBySourceAndId_(source, bookingId) {
  const id = normalizeId_(bookingId);
  if (!id) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  activeSheetNames_().forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const deleteRows = [];

    rows.forEach((row, i) => {
      const b = rowToBooking_(row, sheetName);
      if (b.source === source && normalizeId_(b.bookingId) === id) deleteRows.push(i + 2);
    });

    deleteRows.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
  });
}


/**
 * Move the Confirm and/or Modify thread(s) for a given booking id straight
 * to Cancel, outside the main reconcile sweep (used when a modification pass
 * discovers, mid-run, that the booking it just computed is actually
 * cancelled).
 */
function moveMatchingThreadsToCancellationById_(source, bookingId) {
  const cfg = sourceConfigs_().find(c => c.source === source);
  if (!cfg || !cfg.cancel) return;

  const id = normalizeId_(bookingId);
  if (!id) return;

  const key = source + '|' + id;
  const confirmThread = getConfirmationThreadIndex_().get(key);
  const modifyThread = cfg.modify ? getModificationThreadIndex_().get(key) : null;

  [confirmThread, modifyThread].forEach(thread => {
    if (!thread) return;
    stripSourceLabelsExcept_(thread, cfg, [cfg.cancel]);
    finalizeThreadCancelled_(thread);
  });
}


/**
 * A booking was modified but is still active (not cancelled, not done), so its
 * confirmation thread stays in the inbox as an upcoming tour. We only mark it
 * read + Processed. (Before the inbox policy change this archived the thread,
 * which hid upcoming tours.)
 */
function moveMatchingConfirmationOutOfInbox_(booking) {
  const b = normalizeBooking_(booking);
  const id = normalizeId_(b.bookingId);
  if (!b.source || !id) return;

  const thread = getConfirmationThreadIndex_().get(b.source + '|' + id);
  if (thread) finalizeThreadProcessed_(thread);
}


/******************************************************
 * 9. MODIFICATIONS
 ******************************************************/

function processModifications_() {
  // Viator: dedicated cumulative-delta handler (see function doc below).
  processViatorModificationsLabel_();

  // Guruwalk delivers a full "new state" + "previous booking" pair per
  // message, so the generic single-pass handler is enough.
  processModificationLabel_(RNR.LABELS.GURUWALK_MODIFY, RNR.SOURCE.GURUWALK);

  // GetYourGuide: dedicated handler (an activity can be rescheduled several
  // times, each email carrying both the NEW and the struck-through OLD
  // date; we must apply only the most recent change per booking).
  processGygModificationsLabel_();

  // Free Tour delivers a full "new state" message.
  processModificationLabel_(RNR.LABELS.FREETOUR_MODIFY, RNR.SOURCE.FREETOUR);
}


/**
 * Viator modification handler.
 *
 * Your Gmail filter files every Viator "Amended Booking:" email straight
 * into Viator/Modifications. These are delta-style notifications ("1
 * traveler added" / "1 traveler removed"), not a full restated booking, so
 * final guest count = ORIGINAL CONFIRMATION guests + sum of every delta in
 * the thread's message history (a thread accumulates one message per
 * amendment over the life of the booking).
 *
 * IMPORTANT correctness fix vs the old logic: the delta is always anchored
 * to the ORIGINAL CONFIRMATION booking (findConfirmationBookingBySourceAndId_),
 * never to the current active-sheet row. Anchoring to the active row would
 * double-apply the delta every time the thread is re-processed (each re-run
 * would add totalDelta on top of a row that already reflects totalDelta from
 * the previous run), causing the guest count to drift upward forever. Anchoring
 * to the stable original confirmation makes re-processing idempotent no
 * matter how many times this function runs against the same thread.
 */
function processViatorModificationsLabel_() {
  const threads = getThreadsSafe_(RNR.LABELS.VIATOR_MODIFY);

  threads.forEach(thread => {
    if (!runHasTimeLeft_()) return;

    try {
      const messages = thread.getMessages();
      const byBookingId = new Map();

      messages.forEach(msg => {
        const parsed = parseViatorMessage_(msg, 'modify');
        if (!parsed || !parsed.bookingId) return;

        const id = normalizeId_(parsed.bookingId);
        if (!byBookingId.has(id)) {
          byBookingId.set(id, { bookingId: parsed.bookingId, totalDelta: 0, latestParsed: parsed });
        }

        const item = byBookingId.get(id);
        item.totalDelta += Number(parsed.guestDelta || 0);
        item.latestParsed = parsed; // last message wins for date/time/name/etc.
      });

      byBookingId.forEach(item => {
        const anchor =
          findConfirmationBookingBySourceAndId_(RNR.SOURCE.VIATOR, item.bookingId) ||
          findActiveBookingById_(RNR.SOURCE.VIATOR, item.bookingId);

        if (!anchor) return;              // nothing to amend against
        if (isCompleted_(anchor)) return; // tour already ran

        const anchorGuests = Number(anchor.guests || 1);
        const newGuests = Math.max(1, anchorGuests + Number(item.totalDelta || 0));

        const anchorIncome = Number(anchor.income || 0);
        const newIncome = anchorGuests > 0
          ? Math.round((anchorIncome * newGuests / anchorGuests) * 100) / 100
          : anchorIncome;

        const p = normalizeBooking_(item.latestParsed);

        const finalBooking = normalizeBooking_({
          name: p.name || anchor.name,
          phone: p.phone || anchor.phone,
          guests: newGuests,
          date: p.hasExplicitDate ? p.date : anchor.date,
          time: p.hasExplicitTime ? p.time : anchor.time,
          language: p.language || anchor.language,
          source: RNR.SOURCE.VIATOR,
          income: newIncome,
          bookingId: item.bookingId,
          notes: anchor.notes || '',
          hasExplicitGuests: true,
          hasExplicitDate: true,
          hasExplicitTime: true,
          hasExplicitIncome: true
        });

        if (isBookingCancelledByEmail_(finalBooking)) {
          removeActiveBookingBySourceAndId_(RNR.SOURCE.VIATOR, item.bookingId);
          moveMatchingThreadsToCancellationById_(RNR.SOURCE.VIATOR, item.bookingId);
          return;
        }

        if (isValidBooking_(finalBooking) && !isCompleted_(finalBooking)) {
          upsertActiveBooking_(finalBooking);
          moveMatchingConfirmationOutOfInbox_(finalBooking);
        }
      });

      // Whole thread considered: fast runs will not re-read it until a new
      // message arrives to a NEW thread, or the twice-daily audit re-scans it.
      finalizeThreadProcessed_(thread);

    } catch (e) {
      logError_('processViatorModificationsLabel_', e, RNR.LABELS.VIATOR_MODIFY);
    }
  });
}


/**
 * Dedicated GetYourGuide modification handler (fixes "rescheduled twice but
 * not updated").
 *
 * For every GYG/Modifications thread it:
 *   1. Groups messages by booking code (GYG........).
 *   2. Keeps the LATEST message per code (by email date) for the new
 *      date/time/guests/language, and the latest message that carries a PRICE.
 *   3. Fills missing fields (name, phone, ...) from the existing active row or
 *      the confirmation, recomputes income = price * (1 - GYG_COMMISSION), and
 *      writes the final state.
 *   4. Marks the code reinstated so a stale cancellation cannot delete it.
 */
function processGygModificationsLabel_() {
  const threads = getThreadsSafe_(RNR.LABELS.GYG_MODIFY);

  // code -> { at:Date, text:String, price:Number|null, priceAt:Date|null }
  const latest = new Map();
  const threadsByCode = new Map(); // code -> Set(thread) so we can finalize them all

  threads.forEach(thread => {
    if (!runHasTimeLeft_()) return;
    try {
      thread.getMessages().forEach(msg => {
        const text = (msg.getSubject() || '') + '\n' + getBestMessageText_(msg);
        const id = normalizeId_(extractFirst_(text, [/\b(GYG[A-Z0-9]{5,})\b/i, /\b(S\d{5,})\b/i]));
        if (!id) return;

        const at = msg.getDate();
        const price = gygPrice_(text);

        let item = latest.get(id);
        if (!item) { item = { at: null, text: '', price: null, priceAt: null }; latest.set(id, item); }

        if (!item.at || at > item.at) { item.at = at; item.text = text; }
        if (price > 0 && (!item.priceAt || at > item.priceAt)) { item.price = price; item.priceAt = at; }

        if (!threadsByCode.has(id)) threadsByCode.set(id, new Set());
        threadsByCode.get(id).add(thread);
      });
    } catch (e) {
      logError_('processGygModificationsLabel_ read', e, RNR.LABELS.GYG_MODIFY);
    }
  });

  latest.forEach((item, id) => {
    try {
      const f = gygFields_(item.text);

      // A booking rescheduled from a past date to a future one may already have
      // its confirmation archived to the Done label. Look there too, so we can
      // still recover the customer name/phone the modification email omits.
      const base =
        findActiveBookingById_(RNR.SOURCE.GYG, id) ||
        findConfirmationBookingBySourceAndId_(RNR.SOURCE.GYG, id) ||
        findGygBookingInLabelById_(RNR.LABELS.GYG_DONE, id);

      // Choose the NEW date. The email holds both new and old (struck) dates;
      // prefer the one that differs from the booking's current date, else first.
      let dateTok = f.dateTokens[0] || '';
      if (base && f.dateTokens.length > 1) {
        const changed = f.dateTokens.find(t => dateKey_(normalizeDate_(t)) !== base.dateKey);
        if (changed) dateTok = changed;
      }

      const date = normalizeDate_(dateTok) || (base ? base.date : null);
      const time = extractGygTime_(dateTok) || (base ? base.time : '');
      const mpp = f.participants || null;
      const guests = (mpp && mpp.adults) ? mpp.adults : (f.guests || (base ? base.guests : 1));
      const children = (mpp && mpp.children) ? mpp.children : (base ? Number(base.children || 0) : 0);
      const infants = (mpp && mpp.infants) ? mpp.infants : (base ? Number(base.infants || 0) : 0);
      const language = f.langRaw ? normalizeLanguage_(f.langRaw) : (base ? base.language : RNR.LANGUAGE.ENGLISH);
      const name = (base && base.name) ? base.name : f.name;
      const phone = f.phone || (base ? base.phone : '');
      const income = (item.price != null && item.price > 0)
        ? gygNetIncome_(item.price)
        : (base ? base.income : 0);
      const notes = composeNotes_(f.isPrivate || /privat/i.test(String(base && base.notes || '')), children, infants, '');

      const finalBooking = normalizeBooking_({
        name, phone, guests, children, infants, date, time, language,
        source: RNR.SOURCE.GYG,
        income,
        bookingId: id,
        notes,
        hasExplicitGuests: true,
        hasExplicitDate: true,
        hasExplicitTime: true,
        hasExplicitIncome: true
      });

      if (!isValidBooking_(finalBooking)) return;   // no existing row + no name -> skip
      if (isCompleted_(finalBooking)) return;

      // A rescheduled booking is active. Neutralise any stale cancellation for
      // this code for the rest of the run.
      markReinstated_(RNR.SOURCE.GYG, id);

      upsertActiveBooking_(finalBooking);
      moveMatchingConfirmationOutOfInbox_(finalBooking);
    } catch (e) {
      logError_('processGygModificationsLabel_ apply', e, String(id));
    } finally {
      const ts = threadsByCode.get(id);
      if (ts) ts.forEach(t => finalizeThreadProcessed_(t));
    }
  });

  // Any thread that carried no recognisable booking code still needs to be
  // considered handled, or it will be re-read every fast run forever.
  threads.forEach(thread => {
    if (!threadHasProcessedLabel_(thread)) finalizeThreadProcessed_(thread);
  });
}


/**
 * Generic single-message modification handler (Guruwalk, Free Tour).
 * Guruwalk's message carries both the new state and a "previous booking"
 * block; Free Tour carries only the new state.
 */
function processModificationLabel_(labelName, source) {
  const threads = getThreadsSafe_(labelName);

  for (const thread of threads) {
    if (!runHasTimeLeft_()) break;

    try {
      processModificationThread_(thread, source);
      finalizeThreadProcessed_(thread);
    } catch (e) {
      logError_('processModificationLabel_ ' + source, e, labelName);
    }
  }
}


function processModificationThread_(thread, source) {
  const messages = thread.getMessages();

  for (const msg of messages) {
    const parsed = parseModificationMessage_(msg, source);

    if (parsed.oldBooking) {
      removeActiveBooking_(parsed.oldBooking);
      moveMatchingConfirmationOutOfInbox_(parsed.oldBooking);
    }

    if (parsed.newBooking) {
      const completedBooking = completeModifiedBookingFromExisting_(parsed.newBooking);

      if (isValidBooking_(completedBooking)) {
        if (!isBookingCancelledByEmail_(completedBooking)) {
          if (!isCompleted_(completedBooking)) upsertActiveBooking_(completedBooking);
          moveMatchingConfirmationOutOfInbox_(completedBooking);
        } else {
          removeActiveBookingBySourceAndId_(completedBooking.source, completedBooking.bookingId);
          moveMatchingThreadsToCancellationById_(completedBooking.source, completedBooking.bookingId);
        }
      }
    }
  }
}


/**
 * Route a single modification message to old/new booking objects.
 * Free Tour delivers a single "new state" message; Guruwalk delivers new +
 * previous in one message. Viator is handled entirely by
 * processViatorModificationsLabel_ and never reaches this router.
 */
function parseModificationMessage_(msg, source) {
  if (source === RNR.SOURCE.FREETOUR) {
    return { oldBooking: null, newBooking: parseFreetourMessage_(msg, 'modify') };
  }

  if (source === RNR.SOURCE.GURUWALK) {
    const bookings = parseGuruwalkMessage_(msg, 'modify');
    return { newBooking: bookings[0] || null, oldBooking: bookings[1] || null };
  }

  return { oldBooking: null, newBooking: null };
}


/**
 * Modification emails often omit fields. Fill the gaps from the existing
 * active row (or its confirmation) and recompute income from guest count.
 */
function completeModifiedBookingFromExisting_(booking) {
  const b = normalizeBooking_(booking);
  const existing =
    findActiveBookingById_(b.source, b.bookingId) ||
    findConfirmationBookingBySourceAndId_(b.source, b.bookingId);

  if (!existing) return b;

  const oldGuests = Number(existing.guests || 1);
  let newGuests = oldGuests;

  if (b.hasExplicitGuests) {
    newGuests = Number(b.guests || oldGuests);
  } else if (b.guestDelta) {
    newGuests = Math.max(1, oldGuests + Number(b.guestDelta || 0));
  }

  const oldIncome = Number(existing.income || 0);
  let newIncome = oldIncome;

  if (b.hasExplicitIncome) {
    newIncome = Number(b.income || 0);
  } else if (RNR.MODEL[b.source] === 'free') {
    // Free channels price income purely from guest count.
    newIncome = incomeForFreeSource_(b.source, newGuests);
  } else if (oldGuests > 0 && newGuests !== oldGuests) {
    // Paid channels: scale the net payout proportionally.
    newIncome = Math.round((oldIncome * newGuests / oldGuests) * 100) / 100;
  }

  return normalizeBooking_({
    name: b.name || existing.name,
    phone: b.phone || existing.phone,
    guests: newGuests,
    date: b.hasExplicitDate ? b.date : existing.date,
    time: b.hasExplicitTime ? b.time : existing.time,
    language: b.language || existing.language,
    source: b.source || existing.source,
    income: newIncome,
    bookingId: b.bookingId || existing.bookingId,
    hasExplicitGuests: true,
    hasExplicitDate: true,
    hasExplicitTime: true,
    hasExplicitIncome: true
  });
}


function findActiveBookingById_(source, bookingId) {
  const ref = findActiveBookingRowRef_(source, bookingId);
  return ref ? ref.booking : null;
}

/** Locate a booking's row ACROSS ALL language tabs: {sh, row, sheetName, booking}. */
function findActiveBookingRowRef_(source, bookingId) {
  const id = normalizeId_(bookingId);
  if (!id) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const sheetName of activeSheetNames_()) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) continue;

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    for (let i = 0; i < rows.length; i++) {
      const b = rowToBooking_(rows[i], sheetName);
      if (b.source === source && normalizeId_(b.bookingId) === id) {
        return { sh, row: i + 2, sheetName, booking: b };
      }
    }
  }
  return null;
}


/******************************************************
 * 10. MOVE COMPLETED TO DONE
 ******************************************************/



function moveCompletedBookingRowsToDone_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const completed = [];

  activeSheetNames_().forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const rowsToDelete = [];

    for (let i = 0; i < rows.length; i++) {
      const booking = rowToBooking_(rows[i], sheetName);
      if (isValidBooking_(booking) && isCompleted_(booking)) {
        completed.push(booking);
        rowsToDelete.push(i + 2);
      }
    }

    rowsToDelete.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
  });

  if (completed.length) {
    appendCompletedLog_(completed);   // full detail FIRST (Done Tours aggregates)
    mergeCompletedIntoDone_(completed);
  }
  return completed;                   // drives the targeted Gmail archive
}


/**
 * Append completed bookings (full detail incl. booking id and children) to
 * the hidden Completed Log tab. Idempotent: keyed on bookingId|dateKey, an
 * existing entry is never duplicated. Entries older than
 * COMPLETED_LOG_KEEP_DAYS are pruned. The guide portal's no-show queues
 * consume this tab.
 */
function appendCompletedLog_(completed) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(RNR.COMPLETED_LOG_TAB);
    const header = ['Date', 'Time', 'Language', 'Name', 'Phone', 'Adults',
                    'Children', 'Source', 'Income', 'Booking ID', 'Notes', 'Logged'];
    if (!sh) {
      sh = ss.insertSheet(RNR.COMPLETED_LOG_TAB);
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      sh.setFrozenRows(1);
      try { sh.hideSheet(); } catch (e) { /* fine if it stays visible */ }
    }

    // Existing keys for dedupe + prune of old rows.
    const existing = new Set();
    const cutoff = new Date(Date.now() - RNR.COMPLETED_LOG_KEEP_DAYS * 86400000);
    if (sh.getLastRow() >= 2) {
      const v = sh.getRange(2, 1, sh.getLastRow() - 1, header.length).getValues();
      const del = [];
      v.forEach((r, i) => {
        const dk = dateKey_(r[0]);
        existing.add(String(r[9] || '') + '|' + dk);
        const d = normalizeDate_(r[0]);
        if (d && d < cutoff) del.push(i + 2);
      });
      del.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
    }

    const rows = [];
    completed.forEach(b => {
      const key = String(b.bookingId || '') + '|' + b.dateKey;
      if (!b.bookingId || existing.has(key)) return;
      existing.add(key);
      rows.push([
        stripTime_(b.date), normalizeTime_(b.time), b.language, b.name, b.phone,
        Number(b.guests || 0), Number(b.children || childrenFromNotes_(b.notes) || 0),
        b.source, Number(b.income || 0), b.bookingId, b.notes || '',
        Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm')
      ]);
    });
    if (rows.length) {
      const start = sh.getLastRow() + 1;
      // Time as TEXT so Sheets cannot coerce "11:00 AM" into a Date (that
      // corruption broke shift matching for the Unassigned audit).
      sh.getRange(start, 2, rows.length, 1).setNumberFormat('@');
      sh.getRange(start, 1, rows.length, header.length).setValues(rows);
    }
  } catch (e) {
    logError_('appendCompletedLog_', e, '');
  }
}


function mergeCompletedIntoDone_(completed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const done = ss.getSheetByName(RNR.SHEETS.DONE);
  if (!done) return;

  const map = readDoneMap_();

  completed.forEach(booking => {
    const key = doneKey_(booking.date, booking.time, booking.language);

    if (!map.has(key)) {
      map.set(key, {
        date: stripTime_(booking.date),
        time: normalizeTime_(booking.time),
        language: booking.language,
        freeGuests: 0,
        paidGuests: 0,
        income: 0,
        guide: '',
        comments: ''
      });
    }

    const item = map.get(key);

    // Guest bucket is driven by the source's MODEL (paid vs free).
    if (RNR.MODEL[booking.source] === 'free') {
      item.freeGuests += Number(booking.guests || 0);
    } else {
      item.paidGuests += Number(booking.guests || 0);
    }

    item.income += Number(booking.income || 0);
  });

  const rows = Array.from(map.values())
    .sort((a, b) => combineDateTime_(b.date, b.time) - combineDateTime_(a.date, a.time))
    .map(x => [
      stripTime_(x.date),
      normalizeTime_(x.time),
      x.language,
      Number(x.freeGuests || 0),
      Number(x.paidGuests || 0),
      Number(x.income || 0),
      x.guide || '',
      x.comments || ''
    ]);

  if (done.getLastRow() > 1) {
    done.getRange(2, 1, done.getLastRow() - 1, 8).clearContent();
  }
  if (rows.length) {
    done.getRange(2, 1, rows.length, 8).setValues(rows);
  }
}


function readDoneMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const done = ss.getSheetByName(RNR.SHEETS.DONE);
  const map = new Map();

  if (!done || done.getLastRow() < 2) return map;

  const rows = done.getRange(2, 1, done.getLastRow() - 1, 8).getValues();

  rows.forEach(row => {
    const date = normalizeDate_(row[0]);
    const time = normalizeTime_(row[1]);
    const language = normalizeLanguage_(row[2]);
    if (!date || !time || !language) return;

    map.set(doneKey_(date, time, language), {
      date,
      time,
      language,
      freeGuests: Number(row[3] || 0),
      paidGuests: Number(row[4] || 0),
      income: Number(row[5] || 0),
      guide: String(row[6] || ''),
      comments: String(row[7] || '')
    });
  });

  return map;
}


/**
 * QUOTA-CRITICAL. Archive the Gmail threads of tours that are over.
 *
 * WHY THIS CHANGED (the "Service invoked too many times for one day: gmail"
 * loop): the inbox policy keeps every UPCOMING booking in the inbox, so the
 * old "scan label:x in:inbox and parse each thread" sweep read ~200 message
 * bodies EVERY five minutes — 288 runs/day blew the daily Gmail read quota
 * by mid-morning. The Processed label cannot fix that, because those inbox
 * threads are all legitimately Processed already.
 *
 * Now: we already KNOW which bookings just completed (they were moved to
 * Done Tours this very run). We look their threads up by booking id with a
 * Gmail SEARCH — which does not read message bodies — and archive only
 * those. A run with no completions costs ZERO Gmail calls here.
 *
 * The exhaustive body-reading sweep still exists as a safety net for
 * threads whose id could not be matched, but it runs on the AUDIT only
 * (a few times a day, where the cost is affordable).
 */
function moveCompletedGmailThreadsToDone_(completedBookings) {
  archiveCompletedThreadsByIds_(completedBookings || []);

  if (!RNR_SKIP_PROCESSED_) {          // audit only
    moveCompletedPlatformThreadsToDone_();
    moveCompletedWebsiteThreadsToDone_();
    moveCompletedViatorConversationThreads_();
  }
}

/**
 * Targeted archive: for each just-completed booking, find its thread(s) by
 * booking id (Gmail search, no body reads), strip the working labels, apply
 * <Source>/Done and archive. Idempotent: a thread already carrying Done is
 * skipped by the search filter.
 */
function archiveCompletedThreadsByIds_(completedBookings) {
  if (!completedBookings.length) return;

  const byId = {};
  completedBookings.forEach(b => {
    const id = normalizeId_(b.bookingId);
    if (id && b.source) byId[b.source + '|' + id] = b;
  });

  Object.keys(byId).forEach(k => {
    if (!runHasTimeLeft_()) return;
    const parts = k.split('|');
    const source = parts[0], id = parts[1];
    const cfg = sourceConfigs_().find(c => c.source === source);
    if (!cfg || !cfg.done) return;

    try {
      // Search the working labels for this id; exclude anything already Done.
      const labels = [cfg.confirm, cfg.modify].filter(Boolean)
        .map(l => searchTokenForLabel_(l)).join(' OR ');
      const q = '(' + labels + ') -' + searchTokenForLabel_(cfg.done) + ' "' + id + '"';
      const threads = GmailApp.search(q, 0, 5) || [];
      threads.forEach(t => {
        stripSourceLabelsExcept_(t, cfg, [cfg.done]);
        moveThreadOutOfInbox_(t);
      });
    } catch (e) {
      logError_('archiveCompletedThreadsByIds_ ' + source, e, id);
    }
  });
}


/**
 * Once a booking is completed, its Confirm/Modify/Cancel/Processed labels
 * are ALL stripped and only Done is left, so the live Confirmations and
 * Modifications labels show exclusively upcoming tours — anything that
 * looks "stuck" there but is missing from the sheet is genuinely lost, not
 * just old and completed.
 */
function moveCompletedPlatformThreadsToDone_() {
  sourceConfigs_().filter(cfg => cfg.done).forEach(cfg => {
    if (!GmailApp.getUserLabelByName(cfg.done)) return;

    const activeLabels = [cfg.confirm, cfg.modify].filter(Boolean);
    const seen = new Set();

    activeLabels.forEach(activeLabelName => {
      const threads = getThreadsForCompletionSweep_(activeLabelName);

      for (const thread of threads) {
        if (!runHasTimeLeft_()) break;

        try {
          const threadKey = thread.getId();
          if (seen.has(threadKey)) continue;
          if (!completedThreadShouldMoveToDone_(thread, cfg.source)) continue;

          seen.add(threadKey);
          stripSourceLabelsExcept_(thread, cfg, [cfg.done]);
          moveThreadOutOfInbox_(thread);
        } catch (e) {
          logError_('moveCompletedPlatformThreadsToDone_ ' + cfg.source, e, activeLabelName);
        }
      }
    });
  });
}


function completedThreadShouldMoveToDone_(thread, source) {
  // 1) Normal parse (ordinary confirmations).
  const bookings = uniqueBookings_(parseThread_(thread, source, 'any')).filter(isValidBooking_);
  if (bookings.length && bookings.every(isCompleted_)) return true;

  // 2) Platform booking id (modification threads lacking full fields).
  const bookingId = extractBookingIdFromThread_(thread, source);
  if (bookingId) {
    const active = findActiveBookingById_(source, bookingId);
    if (active && isCompleted_(active)) return true;
  }

  // 3) Subject/body date fallback (old modification threads after the tour date).
  const fallbackDate = extractTourDateFromThread_(thread);
  if (!fallbackDate) return false;

  const fallbackTime = extractTourTimeFromThread_(thread) || RNR.DEFAULT_TIME;

  return isCompleted_(normalizeBooking_({
    name: 'Thread fallback',
    phone: '',
    guests: 1,
    date: fallbackDate,
    time: fallbackTime,
    language: RNR.LANGUAGE.ENGLISH,
    source,
    income: 0,
    bookingId: bookingId || 'THREAD-FALLBACK',
    hasExplicitDate: true,
    hasExplicitTime: Boolean(fallbackTime),
    hasExplicitGuests: false,
    hasExplicitIncome: false
  }));
}




/**
 * Confirmation bookings indexed by "source|id" (built once per run).
 * Used to anchor modifications to the ORIGINAL confirmation, and to fill
 * fields a modification email omits.
 */
function getConfirmationCache_() {
  if (RNR_CONFIRMATION_CACHE_) return RNR_CONFIRMATION_CACHE_;

  const cache = new Map();

  sourceConfigs_().forEach(cfg => {
    if (!runHasTimeLeft_()) return;
    const threads = getThreadsSafe_(cfg.confirm);

    for (const thread of threads) {
      if (!runHasTimeLeft_()) break;
      try {
        parseThread_(thread, cfg.source, 'confirm').forEach(b => {
          const n = normalizeBooking_(b);
          const id = normalizeId_(n.bookingId);
          if (id) cache.set(cfg.source + '|' + id, n);
        });
      } catch (e) {
        logError_('getConfirmationCache_ ' + cfg.source, e, cfg.confirm);
      }
    }
  });

  RNR_CONFIRMATION_CACHE_ = cache;
  return RNR_CONFIRMATION_CACHE_;
}


/**
 * Confirmation THREADS indexed by "source|id" (built once per run). Avoids
 * per-booking full-label rescans, which is what originally blew the Gmail
 * read quota.
 */
function getConfirmationThreadIndex_() {
  if (RNR_CONFIRM_THREAD_INDEX_) return RNR_CONFIRM_THREAD_INDEX_;
  RNR_CONFIRM_THREAD_INDEX_ = buildThreadIndex_(cfg => cfg.confirm, 'confirm');
  return RNR_CONFIRM_THREAD_INDEX_;
}

/**
 * Modification THREADS indexed by "source|id" (built once per run). Used by
 * the cancellation reconcile step to strip the Modify label without
 * rescanning the whole label per booking.
 */
function getModificationThreadIndex_() {
  if (RNR_MODIFY_THREAD_INDEX_) return RNR_MODIFY_THREAD_INDEX_;
  RNR_MODIFY_THREAD_INDEX_ = buildThreadIndex_(cfg => cfg.modify, 'any');
  return RNR_MODIFY_THREAD_INDEX_;
}

function buildThreadIndex_(labelPicker, parseMode) {
  const index = new Map();

  sourceConfigs_().forEach(cfg => {
    const labelName = labelPicker(cfg);
    if (!labelName) return;
    if (!runHasTimeLeft_()) return;

    const threads = getThreadsSafe_(labelName);

    for (const thread of threads) {
      if (!runHasTimeLeft_()) break;
      try {
        const ids = new Set();

        parseThread_(thread, cfg.source, parseMode).forEach(b => {
          const id = normalizeId_(b.bookingId);
          if (id) ids.add(id);
        });

        const threadId = normalizeId_(extractBookingIdFromThread_(thread, cfg.source));
        if (threadId) ids.add(threadId);

        ids.forEach(id => {
          const key = cfg.source + '|' + id;
          if (!index.has(key)) index.set(key, thread);
        });
      } catch (e) {
        logError_('buildThreadIndex_ ' + cfg.source, e, labelName);
      }
    }
  });

  return index;
}


function findConfirmationBookingBySourceAndId_(source, bookingId) {
  const id = normalizeId_(bookingId);
  if (!id) return null;
  return getConfirmationCache_().get(source + '|' + id) || null;
}


function extractTourDateFromThread_(thread) {
  const subject = thread.getFirstMessageSubject() || '';
  const text = subject + '\n' + thread.getMessages()
    .map(m => (m.getSubject() || '') + '\n' + getBestMessageText_(m)).join('\n');

  const dateText = extractFirst_(text, [
    /Amended Booking:\s*([A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4})/i,
    /Cancelled Booking:\s*([A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4})/i,
    /New Booking for\s+([A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4})/i,
    /Travel Date:\s*([^\n\r]+)/i,
    /Tour Date:\s*([^\n\r]+)/i,
    /Fecha(?:\s+del?\s+tour)?:\s*([^\n\r]+)/i,   // ES (GYG / Free Tour)
    /Date:\s*([^\n\r]+)/i
  ]);

  return normalizeDate_(dateText);
}


function extractTourTimeFromThread_(thread) {
  const text = thread.getMessages()
    .map(m => (m.getSubject() || '') + '\n' + getBestMessageText_(m)).join('\n');

  if (/Viator/i.test(thread.getFirstMessageSubject() || '') || /\bBR-\d+\b/i.test(text)) {
    return normalizeTime_(extractViatorTime_(text));
  }

  const timeText = extractFirst_(text, [
    /Time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /Hora:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,   // ES
    /Start time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /(\d{1,2}:\d{2}\s*(?:AM|PM))\s*[–-]/i
  ]);

  return normalizeTime_(timeText);
}


function moveCompletedWebsiteThreadsToDone_() {
  const threads = getThreadsForCompletionSweep_(RNR.LABELS.WEB_CONFIRM);
  if (!GmailApp.getUserLabelByName(RNR.LABELS.WEB_DONE)) return;

  for (const thread of threads) {
    if (!runHasTimeLeft_()) break;

    try {
      const bookings = uniqueBookings_(parseThread_(thread, RNR.SOURCE.WEBSITE, 'any'))
        .filter(isValidBooking_);

      if (!bookings.length) continue;
      if (!bookings.every(isCompleted_)) continue;

      safeAddLabel_(thread, RNR.LABELS.WEB_DONE);
      safeRemoveLabel_(thread, RNR.LABELS.WEB_CONFIRM);
      safeRemoveLabel_(thread, RNR.LABELS.PROCESSED);
      moveThreadOutOfInbox_(thread);
    } catch (e) {
      logError_('moveCompletedWebsiteThreadsToDone_', e, RNR.LABELS.WEB_CONFIRM);
    }
  }
}


function moveCompletedViatorConversationThreads_() {
  const threads = getThreadsForCompletionSweep_(RNR.LABELS.VIATOR_CONVERSATIONS);

  for (const thread of threads) {
    if (!runHasTimeLeft_()) break;

    try {
      const bookingId = extractBookingIdFromThread_(thread, RNR.SOURCE.VIATOR);
      if (!bookingId) continue;

      const base =
        findActiveBookingById_(RNR.SOURCE.VIATOR, bookingId) ||
        findConfirmationBookingBySourceAndId_(RNR.SOURCE.VIATOR, bookingId);

      if (!base || !isCompleted_(base)) continue;

      // Keep the Viator/Conversations label, but archive it.
      moveThreadOutOfInbox_(thread);
    } catch (e) {
      logError_('moveCompletedViatorConversationThreads_', e, RNR.LABELS.VIATOR_CONVERSATIONS);
    }
  }
}


/**
 * Pull a platform booking id from a thread. Each source uses its own reference
 * shape: Viator BR-#######, GetYourGuide GYG######## / S######, Guruwalk /
 * Free Tour alphanumeric codes.
 */
function extractBookingIdFromThread_(thread, source) {
  const subject = thread.getFirstMessageSubject() || '';
  const body = thread.getMessages().map(m => getBestMessageText_(m)).join('\n');
  const text = subject + '\n' + body;

  if (source === RNR.SOURCE.GYG) {
    return extractFirst_(text, [
      /\b(GYG[A-Z0-9]{5,})\b/i,
      /\b(S\d{5,})\b/i,
      /Booking(?:\s*(?:code|reference|ID|number))?\s*[:#-]?\s*([A-Z0-9-]{6,})/i
    ]);
  }

  if (source === RNR.SOURCE.FREETOUR) {
    return extractFirst_(text, [
      /Booking\s*(?:code|reference|ID|number)\s*[:#]?\s*([A-Z0-9-]{4,})/i,
      /Reference\s*[:#]?\s*([A-Z0-9-]{4,})/i,
      /Reserva\s*(?:n[ºo\.]*|c[oó]digo)?\s*[:#]?\s*([A-Z0-9-]{4,})/i
    ]);
  }

  // Default (Viator + generic).
  return extractFirst_(text, [
    /\b(BR-\d+)\b/i,
    /Viator booking\s*(BR-\d+)/i
  ]);
}


/******************************************************
 * 11. ACTIVE SHEET WRITE / DELETE
 ******************************************************/

function upsertActiveBooking_(booking, allowUpdate) {
  if (allowUpdate === undefined) allowUpdate = true;
  const b = normalizeBooking_(booking);

  // Final guard: no confirmation/modification can restore a cancelled booking.
  if (isBookingCancelledByEmail_(b)) {
    if (b.bookingId) {
      removeActiveBookingBySourceAndId_(b.source, b.bookingId);
      moveMatchingThreadsToCancellationById_(b.source, b.bookingId);
    }
    return;
  }

  // CROSS-TAB AUTHORITY: if this booking already lives on ANY language tab,
  // that tab wins — management may have moved the guest to another language
  // (portal "move"), and an email must never move it back or duplicate it
  // into the original language.
  if (b.bookingId) {
    const ref = findActiveBookingRowRef_(b.source, b.bookingId);
    if (ref) {
      if (allowUpdate) {
        if (RNR_RUN_STATS_) RNR_RUN_STATS_.upserts++;
        writeBookingRow_(ref.sh, ref.row, b);   // update fields, keep the tab
      }
      return;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = languageToSheet_(b.language);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  if (RNR_RUN_STATS_) RNR_RUN_STATS_.upserts++;

  // Protect/restore the header row; a booking must never overwrite row 1.
  sh.getRange(1, 1, 1, RNR.ACTIVE_HEADERS.length).setValues([RNR.ACTIVE_HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('@');

  const lastRow = sh.getLastRow();
  const newWhen = combineDateTime_(b.date, b.time);
  let insertAt = -1;   // first existing row that sorts AFTER the new booking
  let isNewRow = true;

  if (lastRow >= 2) {
    const rows = sh.getRange(2, 1, lastRow - 1, 9).getValues();
    for (let i = 0; i < rows.length; i++) {
      const existing = rowToBooking_(rows[i], sheetName);

      // Same booking already on the sheet.
      if (sameBooking_(existing, b)) {
        // MANUAL EDIT AUTHORITY: confirmation re-reads (audits) pass
        // allowUpdate=false — an existing row is never touched, so whatever a
        // manager typed in any column STAYS. Only modification/cancellation
        // emails (allowUpdate=true) or reparseActiveRowsFromEmail change it.
        if (allowUpdate) writeBookingRow_(sh, i + 2, b);
        return;
      }

      // Remember where a new row should slide in to keep the sheet sorted,
      // so we never have to re-sort the whole tab (fewer writes = quota-safe).
      if (insertAt === -1 && combineDateTime_(existing.date, existing.time) > newWhen) {
        insertAt = i + 2;
      }
    }
  }

  if (insertAt > -1) {
    sh.insertRowBefore(insertAt);
    writeBookingRow_(sh, insertAt, b);
  } else {
    writeBookingRow_(sh, Math.max(sh.getLastRow() + 1, 2), b);
  }
}


/** One retry after a pause for Google's transient "Service Spreadsheets
 *  failed" errors (seen in the Errors tab on heavy write passes). */
function withRetry_(fn) {
  try { return fn(); }
  catch (e) { Utilities.sleep(1500); return fn(); }
}

/**
 * BULK write of many booking rows in ONE setValues call (plus one number-
 * format pass per column). The old per-cell writer cost 9 API calls per row,
 * which is what made dedupe/sort passes slow and prone to transient
 * Spreadsheets-service failures.
 */
function writeBookingRowsBulk_(sh, startRow, bookings) {
  if (!bookings.length) return;
  const n = bookings.length;
  withRetry_(() => {
    sh.getRange(startRow, 2, n, 1).setNumberFormat('@');        // phone as text
    sh.getRange(startRow, 3, n, 1).setNumberFormat('0');
    sh.getRange(startRow, 4, n, 1).setNumberFormat('ddd, mmm d');
    sh.getRange(startRow, 5, n, 1).setNumberFormat('@');        // time as text
    sh.getRange(startRow, 7, n, 1).setNumberFormat('0.##');
    sh.getRange(startRow, 8, n, 1).setNumberFormat('@');        // booking id as text
  });
  const values = bookings.map(b => {
    const bb = normalizeBooking_(b);
    return [bb.name, cleanPhone_(bb.phone), Number(bb.guests || 0), stripTime_(bb.date),
            normalizeTime_(bb.time), bb.source, Number(bb.income || 0), bb.bookingId, bb.notes || ''];
  });
  withRetry_(() => sh.getRange(startRow, 1, n, 9).setValues(values));
}


/** Per-column formats of an active booking row (matches ACTIVE_HEADERS). */
const RNR_ROW_FORMATS_ = [['@', '@', '0', 'ddd, mmm d', '@', '@', '0.##', '@', '@']];

function writeBookingRow_(sh, rowNumber, booking) {
  const b = normalizeBooking_(booking);
  const phone = cleanPhone_(b.phone);

  // ONE format call + ONE value call instead of ~24 per-cell calls. The
  // per-cell writer was the single biggest Spreadsheets cost on every run and
  // is what made the 5-minute trigger hit "Service Spreadsheets timed out".
  const r = sh.getRange(rowNumber, 1, 1, 9);
  withRetry_(() => r.setNumberFormats(RNR_ROW_FORMATS_));
  withRetry_(() => r.setValues([[
    b.name,
    phone ? "'" + phone : "",
    Number(b.guests || 0),
    stripTime_(b.date),
    normalizeTime_(b.time),
    b.source,
    Number(b.income || 0),
    b.bookingId,
    b.notes || ''
  ]]));
}


function removeActiveBooking_(booking) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = normalizeBooking_(booking);

  activeSheetNames_().forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const deleteRows = [];

    for (let i = 0; i < rows.length; i++) {
      if (sameBooking_(rowToBooking_(rows[i], sheetName), target)) deleteRows.push(i + 2);
    }

    deleteRows.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
  });
}


function bookingToRow_(b) {
  return [
    b.name,
    cleanPhone_(b.phone),
    Number(b.guests || 0),
    stripTime_(b.date),
    normalizeTime_(b.time),
    b.source,
    Number(b.income || 0),
    b.bookingId,
    b.notes || ''
  ];
}


function rowToBooking_(row, sheetName) {
  return normalizeBooking_({
    name: row[0],
    phone: row[1],
    guests: row[2],
    date: row[3],
    time: row[4],
    language: sheetToLanguage_(sheetName),
    source: row[5],
    income: row[6],
    bookingId: row[7],
    notes: row[8],
    children: childrenFromNotes_(row[8]),
    infants: infantsFromNotes_(row[8]),
    hasExplicitGuests: true,
    hasExplicitDate: true,
    hasExplicitTime: true,
    hasExplicitIncome: true
  });
}


/******************************************************
 * 12. DEDUPLICATION
 ******************************************************/

function dedupeActiveSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  activeSheetNames_().forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 3) return;

    sh.getRange('B:B').setNumberFormat('@');

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const keep = [];

    rows.forEach(row => {
      const b = rowToBooking_(row, sheetName);
      const idx = keep.findIndex(existing => sameBooking_(existing, b));
      if (idx === -1) keep.push(b);
      else keep[idx] = chooseBetterBooking_(keep[idx], b);
    });

    withRetry_(() => sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent());
    writeBookingRowsBulk_(sh, 2, keep);
  });
}


function uniqueBookings_(bookings) {
  const out = [];

  bookings.map(normalizeBooking_).forEach(b => {
    if (!isValidBooking_(b)) return;
    const idx = out.findIndex(x => sameBooking_(x, b));
    if (idx === -1) out.push(b);
    else out[idx] = chooseBetterBooking_(out[idx], b);
  });

  return out;
}


function chooseBetterBooking_(a, b) {
  const A = normalizeBooking_(a);
  const B = normalizeBooking_(b);

  const score = x =>
    String(x.name || '').length +
    String(x.phone || '').length +
    (Number(x.income || 0) > 0 ? 30 : 0) +
    (x.hasExplicitTime ? 10 : 0) +
    (x.hasExplicitGuests ? 10 : 0);

  return score(B) > score(A) ? B : A;
}


/******************************************************
 * 13. MATCHING
 ******************************************************/

function sameBooking_(a, b) {
  const A = normalizeBooking_(a);
  const B = normalizeBooking_(b);

  if (!A.source || !B.source || A.source !== B.source) return false;

  const sameId = A.bookingId && B.bookingId &&
                 normalizeId_(A.bookingId) === normalizeId_(B.bookingId);
  const compatibleName = namesCompatible_(A.name, B.name);

  if (sameId && compatibleName) return true;

  // Viator and GetYourGuide ids are globally unique per booking, so an id
  // match alone is authoritative even if the name text differs slightly.
  if (sameId && (A.source === RNR.SOURCE.VIATOR || A.source === RNR.SOURCE.GYG)) return true;

  const samePhone = A.phoneKey && B.phoneKey && A.phoneKey === B.phoneKey;
  const sameDate = A.dateKey && B.dateKey && A.dateKey === B.dateKey;
  const sameTime = normalizeTime_(A.time) === normalizeTime_(B.time);

  return compatibleName && samePhone && sameDate && sameTime;
}


function namesCompatible_(a, b) {
  const A = normalizeNameKey_(a);
  const B = normalizeNameKey_(b);

  if (!A || !B) return false;
  if (A === B) return true;
  if (A.length >= 4 && B.startsWith(A + ' ')) return true;
  if (B.length >= 4 && A.startsWith(B + ' ')) return true;
  return false;
}


/******************************************************
 * 14. SORTING
 ******************************************************/

function sortActiveSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  activeSheetNames_().forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 3) return;

    sh.getRange('B:B').setNumberFormat('@');

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    const bookings = rows.map(row => rowToBooking_(row, sheetName));

    bookings.sort((a, b) => combineDateTime_(a.date, a.time) - combineDateTime_(b.date, b.time));

    withRetry_(() => sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent());
    writeBookingRowsBulk_(sh, 2, bookings);
  });
}


function sortDoneSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(RNR.SHEETS.DONE);
  if (!sh || sh.getLastRow() < 3) return;

  // Done Tours is an 8-column tab — reading/writing 9 padded a ghost column.
  const range = sh.getRange(2, 1, sh.getLastRow() - 1, 8);
  const rows = range.getValues();
  rows.sort((a, b) => combineDateTime_(b[0], b[1]) - combineDateTime_(a[0], a[1]));
  range.setValues(rows);
}


/******************************************************
 * 15. PARSING ROUTER
 ******************************************************/

function parseThread_(thread, source, mode) {
  const out = [];

  thread.getMessages().forEach(msg => {
    if (source === RNR.SOURCE.VIATOR) {
      const b = parseViatorMessage_(msg, mode);
      if (b) out.push(b);
    } else if (source === RNR.SOURCE.GURUWALK) {
      out.push(...parseGuruwalkMessage_(msg, mode));
    } else if (source === RNR.SOURCE.AIRBNB) {
      const b = parseAirbnbMessage_(msg, mode);
      if (b) out.push(b);
    } else if (source === RNR.SOURCE.GYG) {
      const b = parseGygMessage_(msg, mode);
      if (b) out.push(b);
    } else if (source === RNR.SOURCE.FREETOUR) {
      const b = parseFreetourMessage_(msg, mode);
      if (b) out.push(b);
    } else if (source === RNR.SOURCE.WEBSITE) {
      const b = parseWebsiteAlertMessage_(msg, mode);
      if (b) out.push(b);
    }
  });

  return out;
}


/******************************************************
 * 16. VIATOR PARSER
 ******************************************************/

function parseViatorMessage_(msg, mode) {
  const subject = msg.getSubject() || '';
  const text = subject + '\n' + getBestMessageText_(msg);

  const isCancel = /cancelled|canceled|cancellation|acknowledge this cancellation/i.test(text);
  const isModify = /amended|amendment|booking amended|has been amended|travell?er .* removed|travell?er .* added|passenger .* removed|passenger .* added/i.test(text);

  if (mode === 'confirm' && (isCancel || isModify)) return null;
  if (mode === 'cancel' && !isCancel) return null;
  if (mode === 'modify' && !isModify) return null;

  const bookingId = extractFirst_(text, [
    /Booking Reference:\s*#?(BR-\d+)/i,
    /Booking reference:\s*#?(BR-\d+)/i,
    /Booking ID:\s*#?(BR-\d+)/i,
    /#(BR-\d+)/i,
    /\b(BR-\d+)\b/i
  ]);

  const name = extractFirst_(text, [
    /Lead Traveler Name:\s*([^\n\r]+)/i,
    /Lead traveller name:\s*([^\n\r]+)/i,
    /Lead traveler name\s*([^\n\r]+)/i,
    /Lead traveller name\s*([^\n\r]+)/i,
    /Traveler Name:\s*([^\n\r]+)/i,
    /Traveller Name:\s*([^\n\r]+)/i
  ]);

  const vp = viatorParticipants_(text);
  const explicitGuests = vp.adults > 0 ? vp.adults : extractViatorGuestCount_(text);
  const guestDelta = extractViatorGuestDelta_(text);

  const dateText = extractFirst_(text, [
    /Travel Date:\s*([^\n\r]+)/i,
    /Tour Date:\s*([^\n\r]+)/i,
    /The following booking for[\s\S]{0,160}?on\s+([A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4})/i,
    /for\s+([A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4})\s+\(#/i
  ]);

  const rawTime = extractViatorTime_(text);
  const rawPhone = extractViatorPhone_(text);

  const languageText = extractFirst_(text, [
    /Tour Language:\s*([^\n\r]+)/i,
    /Language:\s*([^\n\r]+)/i
  ]);

  const income = extractViatorNetRate_(text);
  const isPrivate = RNR.PRIVATE_TOUR_KEYWORDS.test(text);   // Viator private tours

  return normalizeBooking_({
    name,
    phone: cleanPhone_(rawPhone),
    guests: explicitGuests || 1,
    date: normalizeDate_(dateText),
    time: rawTime ? normalizeTime_(rawTime) : '',
    language: normalizeLanguage_(languageText),
    source: RNR.SOURCE.VIATOR,
    income,
    bookingId,
    children: vp.children,
    infants: vp.infants,
    notes: composeNotes_(isPrivate, vp.children, vp.infants, ''),
    isCancellation: isCancel,
    guestDelta,
    hasExplicitGuests: explicitGuests !== null,
    hasExplicitDate: Boolean(dateText),
    hasExplicitTime: Boolean(rawTime),
    hasExplicitIncome: Number(income || 0) > 0
  });
}


function extractViatorGuestCount_(text) {
  const s = String(text || '');
  const direct = extractFirst_(s, [
    /Travelers:\s*(\d+)/i,
    /Travellers:\s*(\d+)/i,
    /Passengers:\s*(\d+)/i,
    /Participants:\s*(\d+)/i,
    /Adults?:\s*(\d+)/i,
    /Number of Travelers:\s*(\d+)/i,
    /Number of Travellers:\s*(\d+)/i
  ]);
  return direct ? Number(direct) : null;
}


/**
 * Adults/children from the Viator "Travelers:" line, e.g.
 * "Travelers: 4 Adults" or "Travelers: 2 Adults, 3 Children".
 * A bare number ("Travelers: 4") counts as adults.
 */
function viatorParticipants_(text) {
  const line = extractFirst_(text, [
    /Travelers?:\s*([^\n\r]+)/i,
    /Travellers?:\s*([^\n\r]+)/i,
    /Passengers?:\s*([^\n\r]+)/i
  ]);
  if (!line) return { adults: 0, children: 0, infants: 0 };
  const c = participantCounts_(line);
  if (!c.adults && !c.children && !c.infants) {
    const n = line.match(/(\d+)/);
    if (n) c.adults = Number(n[1]);
  }
  return c;
}


function extractViatorGuestDelta_(text) {
  const s = String(text || '');
  const removed = [...s.matchAll(/(?:travell?er|passenger|participant)[^.\n\r]*removed/gi)].length;
  const added = [...s.matchAll(/(?:travell?er|passenger|participant)[^.\n\r]*added/gi)].length;
  return added - removed;
}


function extractViatorTime_(text) {
  const s = String(text || '');
  const candidates = [
    /Tour Grade Code:\s*[A-Z0-9]+\s*~\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /TG\d+\s*~\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /Start Time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /Start time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /Departure Time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /Time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i
  ];
  for (const r of candidates) {
    const m = s.match(r);
    if (m && isTimeLike_(m[1])) return m[1].trim();
  }
  return '';
}


function isTimeLike_(x) {
  return /^\d{1,2}:\d{2}\s*(AM|PM)?$/i.test(String(x || '').trim());
}


function extractViatorPhone_(text) {
  const raw = extractFirst_(text, [
    /Phone:\s*([^\n\r]+)/i,
    /Phone\s*\(Alternate Phone\)\s*([^\n\r]+)/i,
    /Phone:\s*\(Alternate Phone\)\s*([^\n\r]+)/i,
    /Mobile:\s*([^\n\r]+)/i,
    /Telephone:\s*([^\n\r]+)/i
  ]);
  return String(raw || '')
    .replace(/Send the customer.*$/i, '')
    .replace(/\(Alternate Phone\)/i, '')
    .trim();
}


function extractViatorNetRate_(text) {
  const lines = String(text || '').split(/\n/);
  for (const line of lines) {
    if (!/Net Rate/i.test(line)) continue;
    const afterLabel = line.replace(/^.*?Net Rate\s*:\s*/i, '');
    const matches = afterLabel.match(/[0-9]+(?:[.,][0-9]+)?/g);
    if (matches && matches.length) return parseMoney_(matches[matches.length - 1]);
  }
  const fallback = String(text || '').match(/Net Rate[\s\S]{0,80}?([0-9]+(?:[.,][0-9]+)?)/i);
  return fallback ? parseMoney_(fallback[1]) : 0;
}


/******************************************************
 * 17. GURUWALK PARSER
 ******************************************************/

function parseGuruwalkMessage_(msg, mode) {
  const subject = msg.getSubject() || '';
  const text = subject + '\n' + getBestMessageText_(msg);

  const isCancel = /cancelled|canceled|cancellation|has cancelled a booking/i.test(text);
  const isModify = /modification|modified a booking|previous booking/i.test(text);

  if (mode === 'confirm' && (isCancel || isModify)) return [];
  if (mode === 'cancel' && !isCancel) return [];
  if (mode === 'modify' && !isModify) return [];

  const bookings = parseGuruwalkBlocks_(text);

  return uniqueBookings_(bookings.map(b => {
    b.source = RNR.SOURCE.GURUWALK;
    b.income = guruwalkIncome_(b.guests);
    return normalizeBooking_(b);
  }));
}


function parseGuruwalkBlocks_(text) {
  const s = String(text || '')
    .replace(/ /g, ' ')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n');

  const out = [];
  const matches = [...s.matchAll(/Booking code:\s*([A-Z0-9-]+)/gi)];

  for (let i = 0; i < matches.length; i++) {
    const start = matches[i].index;
    const end = i + 1 < matches.length ? matches[i + 1].index : s.length;

    const walkerStart = Math.max(
      s.lastIndexOf('Walker:', start),
      s.lastIndexOf('Name:', start),
      s.lastIndexOf('Traveler:', start),
      s.lastIndexOf('Traveller:', start)
    );

    const block = s.slice(walkerStart >= 0 ? walkerStart : start, end);

    const bookingId = extractFirst_(block, [
      /Booking code:\s*([A-Z0-9-]+)/i,
      /Booking ID:\s*([A-Z0-9-]+)/i,
      /Reservation ID:\s*([A-Z0-9-]+)/i
    ]);

    const name = extractFirst_(block, [
      /Walker:\s*([^\n\r]+)/i,
      /Name:\s*([^\n\r]+)/i,
      /Traveler:\s*([^\n\r]+)/i,
      /Traveller:\s*([^\n\r]+)/i
    ]);

    const rawPhone = extractFirst_(block, [
      /Phone\s*:?\s*([+\d][+\d\s().-]*)/i,
      /Mobile\s*:?\s*([+\d][+\d\s().-]*)/i,
      /Telephone\s*:?\s*([+\d][+\d\s().-]*)/i
    ]);

    const guestsText = extractFirst_(block, [
      /Attendees:\s*(\d+)/i,
      /People:\s*(\d+)/i,
      /Guests:\s*(\d+)/i,
      /Participants:\s*(\d+)/i
    ]);

    const languageText = extractFirst_(block, [/Language:\s*([^\n\r]+)/i]);

    const dateText = extractFirst_(block, [
      /Date:\s*([^\n\r]+)/i,
      /Tour date:\s*([^\n\r]+)/i
    ]);

    const timeText = extractFirst_(block, [
      /Time:\s*(\d{1,2}:\d{2})/i,
      /Start time:\s*(\d{1,2}:\d{2})/i
    ]);

    out.push({
      bookingId,
      name: cleanText_(name),
      phone: cleanPhone_(rawPhone),
      guests: guestsText ? Number(guestsText) : 1,
      date: normalizeDate_(dateText),
      time: normalizeTime_(timeText || RNR.DEFAULT_TIME),
      language: normalizeLanguage_(languageText),
      hasExplicitGuests: Boolean(guestsText),
      hasExplicitDate: Boolean(dateText),
      hasExplicitTime: Boolean(timeText),
      hasExplicitIncome: true
    });
  }

  return out;
}


/******************************************************
 * 18. AIRBNB PARSER
 ******************************************************/

function parseAirbnbMessage_(msg, mode) {
  const subject = msg.getSubject() || '';
  const text = subject + '\n' + getBestMessageText_(msg);

  // Airbnb confirmations contain a "Cancellations" policy section; don't misread it.
  const isCancel =
    /cancelled your experience/i.test(subject) ||
    /canceled your experience/i.test(subject) ||
    /reservation (was )?cancelled/i.test(text) ||
    /reservation (was )?canceled/i.test(text) ||
    /guest (has )?cancelled/i.test(text) ||
    /guest (has )?canceled/i.test(text);

  const isConfirm =
    /booked your experience/i.test(subject) ||
    /booked your experience/i.test(text) ||
    /Confirmed:/i.test(subject);

  if (mode === 'confirm' && (!isConfirm || isCancel)) return null;
  if (mode === 'cancel' && !isCancel) return null;

  const bookingId = extractFirst_(text, [
    /Confirmation code\s*[:\n\r ]+\s*([A-Z0-9]{6,})/i,
    /\bConfirmation code\b[\s\S]{0,120}?\b([A-Z0-9]{6,})\b/i
  ]);

  const name = extractFirst_(text, [
    /Confirmed:\s*([A-Za-zÀ-ÿ' -]+?)\s+booked your experience/i,
    /Fwd:\s*Confirmed:\s*([A-Za-zÀ-ÿ' -]+?)\s+booked your experience/i,
    /^([A-Za-zÀ-ÿ' -]+?)\s+booked your experience/im,
    /\n([A-Za-zÀ-ÿ' -]+)\n\s*Identity verified/i
  ]);

  const guestsText = extractFirst_(text, [
    /Guests\s*\n\s*(\d+)\s+adults?/i,
    /Guests\s+(\d+)\s+adults?/i,
    /Guests\s*\n\s*(\d+)\s+guests?/i,
    /Guests\s+(\d+)\s+guests?/i,
    /(\d+)\s+adults?/i
  ]);

  const dateText = extractFirst_(text, [
    /Date and time[\s\S]{0,220}?([A-Za-z]{3},\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})/i,
    /Date and time[\s\S]{0,220}?([A-Za-z]+\s+\d{1,2},\s+\d{4})/i,
    /booked your experience for\s+([A-Za-z]+\s+\d{1,2})/i
  ]);

  const timeText = extractFirst_(text, [
    /Date and time[\s\S]{0,260}?(\d{1,2}:\d{2}\s*(?:AM|PM))/i,
    /(\d{1,2}:\d{2}\s*(?:AM|PM))\s*[–-]\s*\d{1,2}:\d{2}/i
  ]);

  const incomeText = extractFirst_(text, [
    /Total\s*\(EUR\)\s*€\s*([0-9]+(?:[.,][0-9]+)?)/i,
    /Total\s*\(EUR\)[\s\S]{0,60}?([0-9]+(?:[.,][0-9]+)?)/i,
    /Total\s*€\s*([0-9]+(?:[.,][0-9]+)?)/i
  ]);

  const date = normalizeAirbnbDate_(dateText, msg.getDate());
  const kidsM = text.match(/(\d+)\s+child(?:ren)?/i);
  const infM = text.match(/(\d+)\s+infants?/i);
  const abChildren = kidsM ? Number(kidsM[1]) : 0;
  const abInfants = infM ? Number(infM[1]) : 0;

  return normalizeBooking_({
    name,
    phone: '',
    guests: guestsText ? Number(guestsText) : 1,   // "N adults" pattern = adults only
    children: abChildren,
    infants: abInfants,
    date,
    time: normalizeTime_(timeText),
    language: RNR.LANGUAGE.ENGLISH,
    source: RNR.SOURCE.AIRBNB,
    income: parseMoney_(incomeText),
    bookingId,
    notes: composeNotes_(false, abChildren, abInfants, 'Airbnb'),
    isCancellation: isCancel,
    hasExplicitGuests: Boolean(guestsText),
    hasExplicitDate: Boolean(date),
    hasExplicitTime: Boolean(timeText),
    hasExplicitIncome: Boolean(incomeText)
  });
}


function normalizeAirbnbDate_(dateText, fallbackMsgDate) {
  const s = cleanText_(dateText);
  const full = normalizeDate_(s);
  if (full) return full;

  const m = s.match(/^([A-Za-z]+)\s+(\d{1,2})$/);
  if (m) {
    const y = fallbackMsgDate instanceof Date ? fallbackMsgDate.getFullYear() : new Date().getFullYear();
    const d = new Date(`${m[1]} ${m[2]}, ${y}`);
    if (!isNaN(d)) return stripTime_(d);
  }
  return null;
}


/******************************************************
 * 19. GETYOURGUIDE PARSER
 *
 * Booking id: the GYG reference (GYG........) is used as the stable id, with
 * the short "S######" number as fallback.
 * Income: GYG is a PAID channel, so we read the supplier payout / net amount.
 ******************************************************/

function parseGygMessage_(msg, mode) {
  const subject = msg.getSubject() || '';
  const text = subject + '\n' + getBestMessageText_(msg);

  // Classification: the SUBJECT is authoritative — Gmail filters route on it,
  // and GYG bodies contain harmless words like "actualizaciones" (footer) that
  // would misclassify a confirmation as a modification. Body keywords are only
  // a fallback for unknown subjects. A rebooking/modification still wins over
  // a cancellation ("reprogramada" emails contain the word "cancelada").
  let isModify, isCancel;
  if (/cancelad|cancelaci[oó]n|cancelled|canceled|anulad/i.test(subject)) {
    isModify = false; isCancel = true;
  } else if (/booking detail change|cambio en los datos|reprogramad|rebooked/i.test(subject)) {
    isModify = true; isCancel = false;
  } else if (/nueva reserva recibida|new booking received|^\s*(?:re:\s*|fwd?:\s*)?booking\s*-/i.test(subject)) {
    isModify = false; isCancel = false;
  } else {
    isModify = /reprogramad|ha vuelto a reservar|se ha modificado|modificad|amend|updated booking|rebooked|actualizad/i.test(text);
    isCancel = !isModify &&
      /cancel(?:led|ed|lation)?|cancelad|cancelaci[oó]n|anulad|reserva anulada/i.test(text);
  }

  if (mode === 'confirm' && (isCancel || isModify)) return null;
  if (mode === 'cancel' && !isCancel) return null;
  if (mode === 'modify' && !isModify) return null;

  const f = gygFields_(text);

  // Confirmations carry one date; if several are present (modifications) the
  // first is the current/new one.
  const dateTok = f.dateTokens[0] || valueAfterLabel_(text, [/^Fecha\b/i, /^Date\b/i]);
  const time = extractGygTime_(dateTok);
  const income = gygNetIncome_(f.price);
  const pp = f.participants || { adults: f.guests || 1, children: 0, infants: 0 };

  return normalizeBooking_({
    name: f.name,
    phone: cleanPhone_(f.phone),
    guests: pp.adults || 1,               // ADULTS only; children live in Notes
    children: pp.children,
    infants: pp.infants,
    date: normalizeDate_(dateTok),
    time,
    language: f.langRaw ? normalizeLanguage_(f.langRaw) : RNR.LANGUAGE.ENGLISH,
    source: RNR.SOURCE.GYG,
    income,
    notes: composeNotes_(f.isPrivate, pp.children, pp.infants, ''),
    isCancellation: isCancel,
    hasExplicitGuests: Boolean(pp.adults || f.guests),
    hasExplicitDate: Boolean(dateTok),
    hasExplicitTime: Boolean(time),
    hasExplicitIncome: income > 0,
    bookingId: f.bookingId
  });
}


/**
 * Find a GYG booking (with a real name) by id inside a given label's threads.
 * Used to recover a rescheduled booking whose confirmation was archived to Done.
 */
function findGygBookingInLabelById_(labelName, bookingId) {
  const id = normalizeId_(bookingId);
  if (!id) return null;

  // Targeted search (bounded): never scan the whole Done label.
  let threads = [];
  try {
    threads = GmailApp.search(searchTokenForLabel_(labelName) + ' "' + id + '"', 0, 3) || [];
  } catch (e) { return null; }

  for (const thread of threads) {
    try {
      const bookings = parseThread_(thread, RNR.SOURCE.GYG, 'any');
      for (const b of bookings) {
        if (normalizeId_(b.bookingId) === id && b.name) return normalizeBooking_(b);
      }
    } catch (e) { /* skip bad thread */ }
  }
  return null;
}


/**
 * DEBUG: run from the editor to see where a booking id lives and how it parses.
 * Edit the id below, run, then read the Execution log.
 */
function debugWhereIsBooking() {
  const id = 'GYG996ZWGNXL';   // <-- change to the booking id you're chasing

  const labels = Object.values(RNR.LABELS);
  labels.forEach(labelName => {
    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) return;
    label.getThreads(0, RNR.MAX_THREADS_AUDIT).forEach(thread => {
      const subj = thread.getFirstMessageSubject() || '';
      if (subj.indexOf(id) !== -1 ||
          thread.getMessages().some(m => getBestMessageText_(m).indexOf(id) !== -1)) {
        console.log('FOUND under label: ' + labelName + '  | subject: ' + subj +
                    '  | unread: ' + thread.isUnread());
      }
    });
  });
  console.log('If nothing printed, that id is under NO Publishing label -> fix the Gmail filter.');
}


/**
 * Extract every GYG field from the (subject + body) text in one place, so the
 * confirmation parser and the modification handler stay in sync.
 */
function gygFields_(text) {
  return {
    bookingId: extractFirst_(text, [/\b(GYG[A-Z0-9]{5,})\b/i, /\b(S\d{5,})\b/i]),
    name: cleanPersonName_(valueAfterLabel_(text, [/^Cliente principal\b/i, /^Lead customer\b/i, /^Main customer\b/i])),
    phone: extractFirst_(text, [
      /Tel[eé]fono:\s*([+\d][+\d\s().-]*)/i,
      /Phone(?:\s*number)?:\s*([+\d][+\d\s().-]*)/i,
      /M[oó]vil:\s*([+\d][+\d\s().-]*)/i
    ]),
    dateTokens: gygDateTokens_(text),
    guests: gygGuests_(text),          // legacy total (fallback)
    participants: gygParticipants_(text),
    isPrivate: RNR.PRIVATE_TOUR_KEYWORDS.test(text),
    langRaw: gygLanguageRaw_(text),
    price: gygPrice_(text)
  };
}


/**
 * Keep only the human name from a value an OTA gave us. Some templates render
 * an email address and/or field labels on the SAME line as the name — e.g. GYG
 * anonymises the customer as customer-xxxx@reply.getyourguide.com:
 *   "Daleska Miclos customer-raawfecnv52myobv@reply.getyourguide.com Teléfono: +34632276997 Idioma: English"
 * Everything from the email (or a trailing field label) onward is dropped, so
 * only "Daleska Miclos" survives. A clean name is returned unchanged.
 *
 * Applied centrally in normalizeBooking_, so EVERY source (GYG, Viator,
 * Website, Free Tour, …) is protected against this class of contamination.
 */
function cleanPersonName_(raw) {
  let s = String(raw || '').replace(/\s+/g, ' ').trim();
  s = s.replace(/\s*\S*@\S+.*$/i, '');                               // email + anything after it
  s = s.replace(/\s*(?:Tel[eé]fono|Phone|M[oó]vil|Idioma|Language|Lengua)\s*:.*$/i, ''); // trailing labels
  return s.replace(/[\s,;:_-]+$/, '').trim();
}


/**
 * All date tokens in order of appearance. Confirmations use an English date
 * ("July 14, 2026 11:00 AM"); modifications use a Spanish one
 * ("14 de julio de 2026 a las 17:00") and list both the new and the old date.
 */
function gygDateTokens_(text) {
  const s = String(text || '');
  const re = /([A-Z][a-z]+ \d{1,2}, \d{4}(?:\s+\d{1,2}:\d{2}\s*(?:AM|PM))?)|(\d{1,2}\s+de\s+[a-zà-ÿ]+\s+de\s+\d{4}(?:\s+a\s+las\s+\d{1,2}:\d{2})?)/gi;
  const out = [];
  let m;
  while ((m = re.exec(s))) out.push((m[1] || m[2]).trim());
  return out;
}


/**
 * Time from a GYG date token: "11:00 AM", "a las 17:00" (24h) or bare "17:00".
 */
function extractGygTime_(token) {
  const s = String(token || '');
  let m = s.match(/(\d{1,2}:\d{2})\s*(AM|PM)/i);
  if (m) return normalizeTime_(m[1] + ' ' + m[2].toUpperCase());
  m = s.match(/a\s+las\s+(\d{1,2}:\d{2})/i);
  if (m) return normalizeTime_(m[1]);
  m = s.match(/(\d{1,2}:\d{2})/);
  if (m) return normalizeTime_(m[1]);
  return '';
}


/**
 * Guest count. Private tours show "1 x Group up to 10 (6 Personas)" -> take the
 * "(N Personas)" figure. Otherwise sum every "N x" ticket line.
 */
function gygGuests_(text) {
  // "Group up to 10 ( 2 Personas)" — GYG wraps the count in <strong>, which
  // htmlToText turns into a space after "(", so allow whitespace inside the parens.
  const priv = String(text || '').match(/\([\s*_]*(\d+)[\s*_]*(?:personas?|personen|people|guests?|pax)[\s*_]*\)/i);
  if (priv) return Number(priv[1]);

  const val = valueAfterLabel_(text, [
    /^N[uú]mero de participantes\b/i,
    /^Participantes\b/i,
    /^Participants?\b/i,
    /^Number of participants\b/i
  ]);
  const xs = [...String(val).matchAll(/(\d+)\s*x\b/gi)];
  if (xs.length) return xs.reduce((sum, mm) => sum + Number(mm[1]), 0);

  const g = extractFirst_(val, [/(\d+)/]);
  return g ? Number(g) : 1;
}


/**
 * Adults/children/infants for a GYG email. The participants block renders as
 * a label line followed by one line per ticket type, e.g.
 *   "Número de participantes"
 *   "1 x Child (Edad 0 - 13)"
 *   "3 x Adults (Edad 14 - 99)"
 * Private tours show "1 x Group up to 10 (5 Personas)" -> 5 adults, children
 * unknown (GYG does not break a private group down).
 */
function gygParticipants_(text) {
  const s = String(text || '');

  // Private: the "(N Personas)" figure is the whole group.
  const priv = s.match(/\([\s*_]*(\d+)[\s*_]*(?:personas?|personen|people|guests?|pax)[\s*_]*\)/i);
  if (priv) return { adults: Number(priv[1]), children: 0, infants: 0 };

  // Window of lines after the participants label (label + up to 5 lines).
  const lines = s.split('\n').map(l => l.trim());
  let win = '';
  for (let i = 0; i < lines.length; i++) {
    if (/^(N[uú]mero de participantes|Participantes|Participants?|Number of participants)\b/i.test(lines[i])) {
      win = lines.slice(i, i + 6).join('\n');
      break;
    }
  }

  const c = participantCounts_(win || s);
  if (!c.adults && !c.children && !c.infants) {
    const total = gygGuests_(s);
    return { adults: total || 1, children: 0, infants: 0 };
  }
  if (!c.adults) c.adults = 1;   // never write a 0-guest tour row
  return c;
}


/**
 * Tour language. Prefer "Idioma del tour" (confirmations); fall back to plain
 * "Idioma" (modifications). Never grab the customer contact "Idioma:" line when
 * a tour-language line exists.
 */
function gygLanguageRaw_(text) {
  const tour = valueAfterLabel_(text, [/^Idioma del tour\b/i, /^Tour language\b/i, /^Language of the tour\b/i]);
  if (tour) return tour;
  return valueAfterLabel_(text, [/^Idioma\b/i, /^Language\b/i]);
}


/**
 * Price shown in the email ("24,00 €", "133,30 €", "Precio abonado 36,00 €").
 */
function gygPrice_(text) {
  const m = String(text || '').match(/([0-9]+[.,][0-9]{2})\s*€/);
  if (m) return parseMoney_(m[1]);
  const v = valueAfterLabel_(text, [/^Precio abonado\b/i, /^Precio\b/i, /^Price\b/i]);
  return parseMoney_(extractFirst_(v, [/([0-9]+[.,]?[0-9]*)/]));
}


/**
 * Income we keep from a GYG price after their commission.
 * income = price * (1 - GYG_COMMISSION).
 */
function gygNetIncome_(price) {
  const p = Number(price || 0);
  if (!p) return 0;
  return Math.round(p * (1 - Number(RNR.GYG_COMMISSION || 0)) * 100) / 100;
}


/**
 * Table-layout helper: find a label line, return its value.
 * GYG / Free Tour render each field as two table cells, so after htmlToText_ the
 * value is usually on the line AFTER the label. Handles "Label: value" too.
 */
function valueAfterLabel_(text, labelRegexes) {
  const lines = String(text || '').split('\n').map(l => l.trim());
  for (let i = 0; i < lines.length; i++) {
    for (const lab of labelRegexes) {
      if (!lab.test(lines[i])) continue;
      const sameLine = lines[i].replace(lab, '').replace(/^[:\s\-–]+/, '').trim();
      if (sameLine) return sameLine;
      for (let j = i + 1; j <= i + 4 && j < lines.length; j++) {
        if (lines[j]) return lines[j];
      }
    }
  }
  return '';
}


/******************************************************
 * 20. FREE TOUR PARSER
 *
 * Free Tour is treated as a FREE channel (tips-based). If no booking
 * reference is present, a stable id is generated so re-runs dedupe.
 ******************************************************/

function parseFreetourMessage_(msg, mode) {
  const subject = msg.getSubject() || '';
  const text = subject + '\n' + getBestMessageText_(msg);

  const isCancel = /cancel(?:led|ed|lation)?|cancelad|cancelaci[oó]n|anulad|reserva anulada/i.test(text);
  const isModify = /modif(?:y|ied|ication)?|modificad|cambio|updated|actualizad|reprogramad|rescheduled/i.test(text);

  if (mode === 'confirm' && (isCancel || isModify)) return null;
  if (mode === 'cancel' && !isCancel) return null;
  if (mode === 'modify' && !isModify) return null;

  let bookingId = extractFirst_(text, [
    /Booking\s*(?:code|reference|ID|number)\s*[:#]?\s*([A-Z0-9-]{4,})/i,
    /Reference\s*[:#]?\s*([A-Z0-9-]{4,})/i,
    /Reserva\s*(?:n[ºo\.]*|c[oó]digo|ID)?\s*[:#]?\s*([A-Z0-9-]{4,})/i,
    /Localizador\s*[:#]?\s*([A-Z0-9-]{4,})/i
  ]);

  const name = extractFirst_(text, [
    /Name:\s*([^\n\r]+)/i,
    /Guest(?:'s)? name:\s*([^\n\r]+)/i,
    /Customer:\s*([^\n\r]+)/i,
    /Traveler:\s*([^\n\r]+)/i,
    /Nombre:\s*([^\n\r]+)/i,
    /Cliente:\s*([^\n\r]+)/i
  ]);

  const rawPhone = extractFirst_(text, [
    /Phone(?:\s*number)?:\s*([+\d][+\d\s().-]*)/i,
    /Mobile:\s*([+\d][+\d\s().-]*)/i,
    /Tel[eé]fono:\s*([+\d][+\d\s().-]*)/i,
    /M[oó]vil:\s*([+\d][+\d\s().-]*)/i
  ]);

  const guestsText = extractFirst_(text, [
    /Guests?:\s*(\d+)/i,
    /People:\s*(\d+)/i,
    /Participants?:\s*(\d+)/i,
    /Pax:\s*(\d+)/i,
    /Personas?:\s*(\d+)/i,
    /Asistentes?:\s*(\d+)/i,
    /Plazas?:\s*(\d+)/i,
    /(\d+)\s+(?:guest|person|people|persona|plaza)s?/i
  ]);

  const dateText = extractFirst_(text, [
    /Date:\s*([^\n\r]+)/i,
    /Tour date:\s*([^\n\r]+)/i,
    /Fecha(?:\s+(?:del?\s+tour|de\s+la\s+reserva))?:\s*([^\n\r]+)/i,
    /D[ií]a:\s*([^\n\r]+)/i
  ]);

  const timeText = extractFirst_(text, [
    /Time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /Start time:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /Hora:\s*(\d{1,2}:\d{2}\s*(?:AM|PM)?)/i,
    /(\d{1,2}:\d{2}\s*(?:AM|PM))/i
  ]);

  const languageText = extractFirst_(text, [
    /Language:\s*([^\n\r]+)/i,
    /Idioma:\s*([^\n\r]+)/i
  ]);

  const guests = guestsText ? Number(guestsText) : 1;
  const ftKidsText = extractFirst_(text, [/Child(?:ren)?:\s*(\d+)/i, /Ni[ñn]os?:\s*(\d+)/i]);
  const ftChildren = ftKidsText ? Number(ftKidsText) : 0;
  const date = normalizeDate_(dateText);

  // Stable fallback id so re-processing dedupes cleanly.
  if (!bookingId) bookingId = generateFallbackBookingId_('FT', date, name, rawPhone, guests);

  return normalizeBooking_({
    name,
    phone: cleanPhone_(rawPhone),
    guests,
    children: ftChildren,
    date,
    time: timeText ? normalizeTime_(timeText) : '',
    language: normalizeLanguage_(languageText),
    source: RNR.SOURCE.FREETOUR,
    income: incomeForFreeSource_(RNR.SOURCE.FREETOUR, guests),
    bookingId,
    notes: composeNotes_(false, ftChildren, 0, 'Free Tour'),
    isCancellation: isCancel,
    hasExplicitGuests: Boolean(guestsText),
    hasExplicitDate: Boolean(dateText),
    hasExplicitTime: Boolean(timeText),
    hasExplicitIncome: true
  });
}


/******************************************************
 * 21. WEBSITE ALERT PARSER
 ******************************************************/

function parseWebsiteAlertMessage_(msg, mode) {
  const subject = msg.getSubject() || '';
  const text = subject + '\n' + getBestMessageText_(msg);

  if (!/NEW WEBSITE RESERVATION/i.test(text)) return null;

  const name = extractFirst_(text, [/Name:\s*([^\n\r]+)/i]);
  const phone = extractFirst_(text, [/Phone:\s*([^\n\r]+)/i]);
  const guests = extractFirst_(text, [/Guests:\s*(\d+)/i]);
  const dateText = extractFirst_(text, [/Tour date:\s*([^\n\r]+)/i]);
  const timeText = extractFirst_(text, [/Time:\s*([^\n\r]+)/i]);
  const languageText = extractFirst_(text, [/Language:\s*([^\n\r]+)/i]);
  const bookingId = extractFirst_(text, [/Booking ID:\s*([A-Z0-9-]+)/i]);
  const message = extractFirst_(text, [/Message:\s*([\s\S]*)/i]);

  return normalizeBooking_({
    name,
    phone: cleanPhone_(phone),
    guests: guests ? Number(guests) : 1,
    date: normalizeDate_(dateText),
    time: normalizeTime_(timeText),
    language: normalizeLanguage_(languageText),
    source: RNR.SOURCE.WEBSITE,
    income: 0,
    bookingId,
    notes: message,
    hasExplicitGuests: Boolean(guests),
    hasExplicitDate: Boolean(dateText),
    hasExplicitTime: Boolean(timeText),
    hasExplicitIncome: true
  });
}


/******************************************************
 * 21B. PARTICIPANTS (adults / children / infants)
 *
 * Children never count toward "Number of Guests" (adults only), never change
 * income, and are recorded in Notes ("3 children") so the guide portal can
 * show "2+3". Infants are kept separately when the source supplies them.
 ******************************************************/

/**
 * Count adults/children/infants in a participant text fragment.
 * Matches "3 x Adults (Edad 14 - 99)", "1 x Child (Edad 0 - 13)",
 * "2 adults", "1 Niño", "1 Infant", "2 Erwachsene" etc.
 */
function participantCounts_(text) {
  const s = String(text || '');
  let adults = 0, children = 0, infants = 0;
  // [\s*_·-]* between the parts so bold markers / stray punctuation from any
  // body flavour ("*3 x* Adults", "3 x  Adults") still match.
  const re = /(\d+)[\s*_·-]*(?:x[\s*_·-]*)?(adult[oe]?s?|adults?|child(?:ren)?|ni[ñn][oa]s?|kids?|infants?|beb[eé]s?|kinder|erwachsene[rn]?)\b/gi;
  let m;
  while ((m = re.exec(s))) {
    const n = Number(m[1]);
    const w = m[2].toLowerCase();
    if (/^(child|ni[ñn]|kid)/.test(w)) children += n;
    else if (/^(infant|beb)/.test(w)) infants += n;
    else adults += n;
  }
  return { adults, children, infants };
}


/**
 * Compose the Notes cell. Order: Private flag, children, infants, extra tag.
 * The guide portal parses "Private" and "N children" back out of this string.
 */
function composeNotes_(isPrivate, children, infants, extra) {
  const parts = [];
  if (isPrivate) parts.push(RNR.PRIVATE_TOUR_NOTE);
  if (Number(children) > 0) parts.push(children + (Number(children) === 1 ? ' child' : ' children'));
  if (Number(infants) > 0) parts.push(infants + (Number(infants) === 1 ? ' infant' : ' infants'));
  if (extra) parts.push(String(extra));
  return parts.join(' · ');
}


/** Recover a child/infant count from a Notes cell ("Private · 3 children"). */
function childrenFromNotes_(notes) {
  const m = String(notes || '').match(/(\d+)\s*child/i);
  return m ? Number(m[1]) : 0;
}

function infantsFromNotes_(notes) {
  const m = String(notes || '').match(/(\d+)\s*infant/i);
  return m ? Number(m[1]) : 0;
}


/******************************************************
 * 22. NORMALIZATION HELPERS
 ******************************************************/

function normalizeBooking_(x) {
  const date = normalizeDate_(x && x.date);
  const time = normalizeTime_(x && x.time);
  const language = normalizeLanguage_(x && x.language);
  const phone = cleanPhone_(x && x.phone);
  const name = cleanPersonName_(x && x.name);

  return {
    name,
    phone,
    guests: Number((x && x.guests) || 1),
    date,
    time,
    language,
    source: cleanText_(x && x.source),
    income: Number((x && x.income) || 0),
    bookingId: cleanText_(x && x.bookingId),
    notes: cleanText_(x && x.notes),
    isCancellation: Boolean(x && x.isCancellation),
    children: Number((x && x.children) || 0),
    infants: Number((x && x.infants) || 0),
    guestDelta: Number((x && x.guestDelta) || 0),
    hasExplicitGuests: Boolean(x && x.hasExplicitGuests),
    hasExplicitDate: Boolean(x && x.hasExplicitDate),
    hasExplicitTime: Boolean(x && x.hasExplicitTime),
    hasExplicitIncome: Boolean(x && x.hasExplicitIncome),
    nameKey: normalizeNameKey_(name),
    phoneKey: cleanPhoneKey_(phone),
    dateKey: dateKey_(date)
  };
}


function isValidBooking_(x) {
  const b = normalizeBooking_(x);
  return Boolean(b.name && b.bookingId && b.date && b.time && b.language && b.source);
}


function cleanText_(x) {
  return String(x || '').replace(/\s+/g, ' ').trim();
}


function cleanPhone_(x) {
  let s = String(x || '').trim();
  s = s
    .replace(/^=/, '')
    .replace(/^"|"$/g, '')
    .replace(/^'/, '')
    .replace(/Phone.*?:/i, '')
    .replace(/Mobile.*?:/i, '')
    .replace(/Telephone.*?:/i, '')
    .replace(/Tel[eé]fono.*?:/i, '')
    .replace(/M[oó]vil.*?:/i, '')
    .replace(/\(Alternate Phone\)/i, '')
    .replace(/Send the customer.*$/i, '')
    .trim();

  s = s.replace(/[^\d+]/g, '');
  if (!s) return '';

  if (s.includes('+')) s = '+' + s.replace(/\+/g, '');
  else s = '+' + s;

  return s;
}


function cleanPhoneKey_(x) {
  return String(x || '').replace(/[^\d]/g, '');
}


function normalizeNameKey_(x) {
  return String(x || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^\p{L}\p{N}\s]/gu, '')
    .replace(/\s+/g, ' ')
    .trim();
}


function normalizeId_(x) {
  return String(x || '').toUpperCase().replace(/\s+/g, '').replace(/^#/, '').trim();
}


function normalizeLanguage_(x) {
  const s = String(x || '').toLowerCase();

  if (s.includes('spanish') || s.includes('español') || s.includes('espanol') ||
      s.includes('castellano') || s.includes('spanisch') || s.includes('espagnol')) {
    return RNR.LANGUAGE.SPANISH;
  }
  if (s.includes('german') || s.includes('deutsch') || s.includes('alemán') ||
      s.includes('aleman') || s.includes('alemao') || s.includes('allemand')) {
    return RNR.LANGUAGE.GERMAN;
  }
  if (s.includes('italian') || s.includes('italiano') || s.includes('italiana') ||
      s.includes('italien') || s.includes('italienisch')) {
    return RNR.LANGUAGE.ITALIAN;
  }
  if (s.includes('french') || s.includes('français') || s.includes('francais') ||
      s.includes('francese') || s.includes('francés') || s.includes('frances') ||
      s.includes('französisch') || s.includes('franzosisch')) {
    return RNR.LANGUAGE.FRENCH;
  }
  // Includes GYG Spanish wording: "Inglés (Live tour guide)".
  if (s.includes('english') || s.includes('inglés') || s.includes('ingles') ||
      s.includes('anglais') || s.includes('englisch')) {
    return RNR.LANGUAGE.ENGLISH;
  }
  // Genuinely unknown languages default to English (unchanged legacy behaviour).
  return RNR.LANGUAGE.ENGLISH;
}


function languageToSheet_(language) {
  const lang = normalizeLanguage_(language);
  if (lang === RNR.LANGUAGE.SPANISH) return RNR.SHEETS.SPANISH;
  if (lang === RNR.LANGUAGE.GERMAN) return RNR.SHEETS.GERMAN;
  if (lang === RNR.LANGUAGE.ITALIAN) return RNR.SHEETS.ITALIAN;
  if (lang === RNR.LANGUAGE.FRENCH) return RNR.SHEETS.FRENCH;
  return RNR.SHEETS.ENGLISH;
}


function sheetToLanguage_(sheetName) {
  if (sheetName === RNR.SHEETS.SPANISH) return RNR.LANGUAGE.SPANISH;
  if (sheetName === RNR.SHEETS.GERMAN) return RNR.LANGUAGE.GERMAN;
  if (sheetName === RNR.SHEETS.ITALIAN) return RNR.LANGUAGE.ITALIAN;
  if (sheetName === RNR.SHEETS.FRENCH) return RNR.LANGUAGE.FRENCH;
  return RNR.LANGUAGE.ENGLISH;
}


function parseMoney_(x) {
  const s = String(x || '').replace(/[^\d.,]/g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}




/** Guruwalk income = flat commission we charge the guide, 6 €/guest. */
function guruwalkIncome_(guests) {
  return Number(guests || 0) * Number(RNR.FREE_COMMISSION_PER_GUEST || 0);
}


/**
 * Income for free-model sources = flat 6 €/guest commission the guide owes us
 * (Guruwalk, Free Tour and Website alike).
 */
function incomeForFreeSource_(source, guests) {
  return Math.round(Number(guests || 0) * Number(RNR.FREE_COMMISSION_PER_GUEST || 0) * 100) / 100;
}


/******************************************************
 * 23. DATE AND TIME HELPERS
 ******************************************************/

function normalizeDate_(x) {
  if (x instanceof Date && !isNaN(x)) return stripTime_(x);

  let s = String(x || '').trim();
  if (!s) return null;

  // Drop a leading weekday word in any language, e.g. "Sat," / "Sáb," / "Fri".
  s = s.replace(/^[A-Za-zÀ-ÿ]{2,10},\s*/i, '').replace(/\s+/g, ' ').trim();

  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));

  const euro = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (euro) {
    let year = Number(euro[3]);
    if (year < 100) year += 2000;
    return new Date(year, Number(euro[2]) - 1, Number(euro[1]));
  }

  // Spanish long date, e.g. "14 de julio de 2026" (with optional " a las 17:00").
  // GYG modification emails use this format; confirmations use English dates.
  const es = s.match(/(\d{1,2})\s+de\s+([a-zà-ÿ]+)\s+de\s+(\d{4})/i);
  if (es) {
    const mo = RNR_SPANISH_MONTHS_[es[2].toLowerCase()];
    if (mo !== undefined) return new Date(Number(es[3]), mo, Number(es[1]));
  }

  const d = new Date(s);
  if (isNaN(d)) return null;
  return stripTime_(d);
}


function normalizeTime_(x) {
  const s = String(x || '').trim();
  if (!s) return '';

  let m = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if (m) return `${Number(m[1])}:${String(m[2]).padStart(2, '0')} ${m[3].toUpperCase()}`;

  m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return timeFrom24_(Number(m[1]), Number(m[2]));

  return '';
}


function timeFrom24_(hour, minute) {
  const suffix = hour >= 12 ? 'PM' : 'AM';
  const h12 = ((hour + 11) % 12) + 1;
  return `${h12}:${String(minute).padStart(2, '0')} ${suffix}`;
}


function combineDateTime_(date, time) {
  const d = normalizeDate_(date);
  if (!d) return new Date(9999, 0, 1);

  const t = normalizeTime_(time);
  const m = t.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);

  let h = 0, min = 0;
  if (m) {
    h = Number(m[1]);
    min = Number(m[2]);
    if (m[3].toUpperCase() === 'PM' && h < 12) h += 12;
    if (m[3].toUpperCase() === 'AM' && h === 12) h = 0;
  }

  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), h, min, 0);
}


function stripTime_(d) {
  if (!(d instanceof Date) || isNaN(d)) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


function dateKey_(date) {
  const d = normalizeDate_(date);
  if (!d) return '';
  return [
    d.getFullYear(),
    String(d.getMonth() + 1).padStart(2, '0'),
    String(d.getDate()).padStart(2, '0')
  ].join('-');
}


function isCompleted_(booking) {
  const b = normalizeBooking_(booking);
  const start = combineDateTime_(b.date, b.time);
  const doneAt = new Date(start.getTime() + RNR.DONE_AFTER_HOURS * 60 * 60 * 1000);
  return new Date() >= doneAt;
}


function doneKey_(date, time, language) {
  return [dateKey_(date), normalizeTime_(time), normalizeLanguage_(language)].join('|');
}


/******************************************************
 * 24. EMAIL TEXT HELPERS
 *
 * getBestMessageText_ is the single choke point for reading a message body.
 * It (a) caches per message id so a message is fetched from Gmail at most once
 * per run, and (b) swallows body-read failures so one bad message cannot
 * abort the run ("operation not allowed" fix).
 ******************************************************/

/**
 * Choose which rendering of a message body to parse, then normalise it.
 *
 * Gmail gives two flavours and they are NOT equivalent:
 *   getBody()       -> real HTML. htmlToText_ turns the booking table into
 *                      "label \n value" lines. This is the good one.
 *   getPlainBody()  -> Gmail's own text conversion. It renders <strong> as
 *                      *asterisks* ("*1 x* Child", "(*5* Personas)") and can
 *                      flatten the table, which is what made a 3-adult
 *                      booking store as 1 guest and a 5-person private group
 *                      store as 1.
 *
 * Rule: prefer the flavour that actually contains the participants block; if
 * both or neither do, prefer HTML when it looks substantial. Then strip the
 * asterisk bold markers so either flavour parses identically.
 */
function pickBestBody_(html, plain) {
  const norm = t => String(t || '').replace(/\*+/g, '');
  const h = norm(html), p = norm(plain);
  const hasParts = t => /N[uú]mero de participantes|Participantes|Number of participants|Participants?\b|Travelers?:|Travellers?:/i.test(t);

  const hOk = h && hasParts(h);
  const pOk = p && hasParts(p);

  if (hOk && !pOk) return h;
  if (pOk && !hOk) return p;
  if (h && h.length >= 200) return h;      // both or neither: HTML is richer
  return p || h;
}


function getBestMessageText_(msg) {
  try {
    if (!RNR_TEXT_CACHE_) RNR_TEXT_CACHE_ = new Map();

    const id = msg.getId();
    if (RNR_TEXT_CACHE_.has(id)) return RNR_TEXT_CACHE_.get(id);

    let text = '';
    try {
      const plain = String(msg.getPlainBody() || '').trim();
      const html = htmlToText_(msg.getBody() || '');
      text = pickBestBody_(html, plain);
    } catch (bodyErr) {
      // Rare Gmail states throw on body read. Fall back to the subject.
      text = String(msg.getSubject() || '');
    }

    RNR_TEXT_CACHE_.set(id, text);
    return text;

  } catch (e) {
    return '';
  }
}


function htmlToText_(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/td>/gi, '\n')
    .replace(/<\/th>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&#39;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/\r/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}


function extractFirst_(text, regexes) {
  const s = String(text || '');
  for (const r of regexes) {
    const m = s.match(r);
    if (m && m[1] !== undefined) return String(m[1]).trim();
  }
  return '';
}


/******************************************************
 * 26. RESPONSE AND ERROR HELPERS
 ******************************************************/

function textResponse_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.TEXT);
}


function safeString_(v) {
  return v === null || v === undefined ? '' : String(v);
}


function logError_(type, err, rawData) {
  try {
    if (RNR_RUN_STATS_) RNR_RUN_STATS_.errors++;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(RNR.SHEETS.ERRORS) || ss.insertSheet(RNR.SHEETS.ERRORS);

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, RNR.ERROR_HEADERS.length).setValues([RNR.ERROR_HEADERS]);
    }

    const details = String(err && err.stack ? err.stack : err);
    const now = new Date();
    const nowText = Utilities.formatDate(now, 'Europe/Madrid', 'yyyy-MM-dd HH:mm');

    // DEDUP: same Type + same first 120 chars of Details within the last 24h
    // -> bump Count + Last seen on the existing row instead of appending.
    const sig = type + '|' + details.slice(0, 120);
    const last = sh.getLastRow();
    if (last >= 2) {
      const n = Math.min(30, last - 1);
      const v = sh.getRange(last - n + 1, 1, n, RNR.ERROR_HEADERS.length).getValues();
      for (let i = v.length - 1; i >= 0; i--) {
        const ts = v[i][0];
        if (!(ts instanceof Date) || now - ts > 24 * 3600000) continue;
        if (String(v[i][1]) + '|' + String(v[i][2]).slice(0, 120) === sig) {
          const row = last - n + 1 + i;
          sh.getRange(row, 5).setValue((Number(v[i][4]) || 1) + 1);
          sh.getRange(row, 6).setValue(nowText);
          return;
        }
      }
    }

    sh.appendRow([now, type, details, rawData || '', 1, nowText]);
  } catch (e) {
    console.log('logError_ failed: ' + e + ' | original: ' + type + ' ' + err);
  }
}


/******************************************************
 * 27. DEBUG / PARSER TUNING HELPERS
 *
 * Run these from the Apps Script editor. They print exactly what each
 * parser extracts so you can confirm the field regexes match your real
 * emails before trusting the automation.
 ******************************************************/

function debugGygParsingOnly() {
  debugParseLabel_(RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG, parseGygMessage_);
}

function debugFreetourParsingOnly() {
  debugParseLabel_(RNR.LABELS.FREETOUR_CONFIRM, RNR.SOURCE.FREETOUR, parseFreetourMessage_);
}

function debugParseLabel_(labelName, source, parserFn) {
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) { console.log('Missing label: ' + labelName); return; }

  const threads = label.getThreads(0, 10);
  console.log('Threads found under ' + labelName + ': ' + threads.length);

  threads.forEach((thread, ti) => {
    console.log('\n===== THREAD ' + (ti + 1) + ' =====');
    console.log('Subject: ' + thread.getFirstMessageSubject());

    thread.getMessages().forEach((msg, mi) => {
      const parsed = parserFn(msg, 'confirm');
      console.log('\n-- message ' + (mi + 1) + ' --');
      console.log('Parsed: ' + JSON.stringify(parsed, null, 2));
      if (parsed) console.log('Valid booking: ' + isValidBooking_(parsed));

      const text = (msg.getSubject() || '') + '\n' + getBestMessageText_(msg);
      console.log('Text sample:\n' + text.slice(0, 2500));
    });
  });
}


/** Retained multi-source parser dump. */
function debugRecentParsedEmails() {
  const tests = [
    [RNR.LABELS.VIATOR_CONFIRM, RNR.SOURCE.VIATOR, 'confirm'],
    [RNR.LABELS.VIATOR_CANCEL, RNR.SOURCE.VIATOR, 'cancel'],
    [RNR.LABELS.VIATOR_MODIFY, RNR.SOURCE.VIATOR, 'modify'],
    [RNR.LABELS.GURUWALK_CONFIRM, RNR.SOURCE.GURUWALK, 'confirm'],
    [RNR.LABELS.GURUWALK_CANCEL, RNR.SOURCE.GURUWALK, 'cancel'],
    [RNR.LABELS.GURUWALK_MODIFY, RNR.SOURCE.GURUWALK, 'modify'],
    [RNR.LABELS.GYG_CONFIRM, RNR.SOURCE.GYG, 'confirm'],
    [RNR.LABELS.GYG_CANCEL, RNR.SOURCE.GYG, 'cancel'],
    [RNR.LABELS.GYG_MODIFY, RNR.SOURCE.GYG, 'modify'],
    [RNR.LABELS.FREETOUR_CONFIRM, RNR.SOURCE.FREETOUR, 'confirm'],
    [RNR.LABELS.FREETOUR_CANCEL, RNR.SOURCE.FREETOUR, 'cancel'],
    [RNR.LABELS.FREETOUR_MODIFY, RNR.SOURCE.FREETOUR, 'modify']
  ];

  for (const [labelName, source, mode] of tests) {
    const label = GmailApp.getUserLabelByName(labelName);
    console.log('\n\n===== ' + labelName + ' =====');
    if (!label) { console.log('Missing label.'); continue; }

    label.getThreads(0, 5).forEach(thread => {
      console.log(JSON.stringify(parseThread_(thread, source, mode), null, 2));
    });
  }
}


/**
 * DEBUG: print which labels are currently on every thread under every
 * RNR label, and whether each carries Processed. Useful for sanity-checking
 * the new label lifecycle (cancel should strip confirm/modify; done should
 * strip everything except done) without waiting for a live run.
 */
function debugLabelLifecycleAudit() {
  Object.entries(RNR.LABELS).forEach(([key, labelName]) => {
    const label = GmailApp.getUserLabelByName(labelName);
    if (!label) { console.log(key + ' (' + labelName + '): MISSING'); return; }

    const threads = label.getThreads(0, 5);
    console.log('\n===== ' + key + ' — ' + labelName + ' (' + threads.length + ' shown) =====');

    threads.forEach(thread => {
      const names = (thread.getLabels() || []).map(l => l.getName()).join(', ');
      console.log(thread.getFirstMessageSubject() + '  ->  [' + names + ']');
    });
  });
}


/******************************************************
 * 28. PARSER FIXTURES + ACCEPTANCE TESTS
 *
 * Run testBookingParsers() from the Apps Script editor. It parses embedded
 * copies of real emails (no Gmail reads) and logs PASS/FAIL per assertion.
 * Add a new fixture whenever an OTA changes its template.
 ******************************************************/

function makeFakeMsg_(subject, body) {
  const id = 'FIXTURE_' + Math.random().toString(36).slice(2);
  return {
    getId: function () { return id; },
    getSubject: function () { return subject; },
    getPlainBody: function () { return body; },
    getBody: function () { return ''; },
    getDate: function () { return new Date(); }
  };
}

const RNR_FIXTURES_ = {
  // GYG Spanish confirmation with children (from real email, 2026-07)
  gygEsChildren: {
    subject: 'Booking - S779080 - GYGWZARHZ63W',
    body: [
      '¡Hola! Buenas noticias.', 'Se ha reservado tu producto', '',
      'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town',
      'Número de referencia', 'GYGWZARHZ63W',
      'Fecha', 'July 20, 2026 11:00 AM',
      'Número de participantes', '1 x Child (Edad 0 - 13)', '3 x Adults (Edad 14 - 99)',
      'Cliente principal', 'Fabiola Gardi',
      'customer-3o35yxxqhbthbdbn@reply.getyourguide.com',
      'Teléfono: +393334666605', 'Idioma: Italian',
      'Idioma del tour', 'Inglés (Live tour guide)',
      'Precio', '54,90 €'
    ].join('\n')
  },
  // GYG Spanish private confirmation (Tour privado / Group up to N)
  gygEsPrivate: {
    subject: 'Urgente: nueva reserva recibida - S779080 - GYGG45QBNGLN',
    body: [
      '¡Hola! Buenas noticias.', 'Ha recibido una reserva de última hora:', '',
      'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town', 'Tour privado',
      'Número de referencia', 'GYGG45QBNGLN',
      'Fecha', 'July 15, 2026 10:00 AM',
      'Número de participantes', '1 x Group up to 10 (5 Personas)',
      'Cliente principal', 'Daniel Morban',
      'customer-ytsxu45y6p5ru2gk@reply.getyourguide.com',
      'Teléfono: +18099921237', 'Idioma: Spanish',
      'Idioma del tour', 'Inglés (Live tour guide)',
      'Precio', '133,30 €'
    ].join('\n')
  },
  // Viator private product confirmation (real email, 2026-07)
  viatorPrivate: {
    subject: 'New Booking for Fri, Jul 24, 2026 (#BR-1422956849)',
    body: [
      'You have a new reservation for Private Complete Barcelona Walking Tour with Local Guide.',
      'Booking Details',
      'Booking Reference: BR-1422956849',
      'Tour Name: Private Complete Barcelona Walking Tour with Local Guide',
      'Travel Date: Fri, Jul 24, 2026',
      'Lead Traveler Name: Detlef Bludau',
      'Traveler Names: Detlef Bludau, Passenger Two, Passenger Three, Passenger Four',
      'Travelers: 4 Adults',
      'Product Code: 5631527P5',
      'Tour Grade: German Tour 11:00',
      'Tour Grade Code: TG3~11:00',
      'Tour Language: German - Guide',
      'Net Rate: EUR €93,50',
      'Phone: (Alternate Phone)DE+49 1797747255 Send the customer a message.'
    ].join('\n')
  },
  // REGRESSION (2026-07-20): real bookings that were stored as 1 guest.
  // Gmail's getPlainBody() renders <strong> as *...*; these are that exact
  // flavour. GYG996ZFRKQ9 = 3 adults + 1 child, GYG48YMW6KAB = 5-person private.
  gygPlainBodyChildren: {
    subject: 'Booking - S779080 - GYG996ZFRKQ9',
    body: [
      '¡Hola! Buenas noticias.', 'Ha recibido una reserva de última hora:',
      'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town',
      'Número de referencia', 'GYG996ZFRKQ9',
      'Fecha', 'July 20, 2026 11:00 AM',
      'Número de participantes', '',
      '*1 x* Child (Edad 0 - 13)', '*3 x* Adults (Edad 14 - 99)',
      'Cliente principal', 'Katarzyna Świstak',
      'Teléfono: +48607788771', 'Idioma: Polish',
      'Idioma del tour', 'Inglés (Live tour guide)',
      'Precio', '54,90 €'
    ].join('\n')
  },
  gygPlainBodyPrivate: {
    subject: 'Booking - S779080 - GYG48YMW6KAB',
    body: [
      '¡Hola! Buenas noticias.', 'Se ha reservado tu producto',
      'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town', 'Tour privado',
      'Número de referencia', 'GYG48YMW6KAB',
      'Fecha', 'July 21, 2026 10:00 AM',
      'Número de participantes', '',
      '*1 x* Group up to 8 (*5* Personas)',
      'Cliente principal', 'Angelika Mbanza',
      'Teléfono: +41763166313', 'Idioma: French',
      'Idioma del tour', 'Inglés (Live tour guide)',
      'Precio', '133,30 €'
    ].join('\n')
  },
  // Viator mixed adults + children
  viatorChildren: {
    subject: 'New Booking for Mon, Aug 3, 2026 (#BR-1409320087)',
    body: [
      'Booking Reference: BR-1409320087',
      'Travel Date: Mon, Aug 3, 2026',
      'Lead Traveler Name: Hang Nguyen',
      'Travelers: 2 Adults, 3 Children',
      'Tour Grade Code: TG1~11:00',
      'Tour Language: English - Guide',
      'Net Rate: EUR €95,60',
      'Phone: +17032093574'
    ].join('\n')
  },
  // REGRESSION (2026-07): GYG rendered the anonymised customer email (and the
  // Teléfono/Idioma labels) on the SAME line as "Cliente principal", so the
  // whole "Name customer-...@... Teléfono ... Idioma ..." string landed in the
  // Name column. Only the human name must survive.
  gygInlineCustomerEmail: {
    subject: 'Booking - S779080 - GYG996ZAK7R7',
    body: [
      '¡Hola! Buenas noticias.', 'Se ha reservado tu producto',
      'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town',
      'Número de referencia', 'GYG996ZAK7R7',
      'Fecha', 'July 23, 2026 11:00 AM',
      'Número de participantes', '1 x Adults (Edad 14 - 99)',
      'Cliente principal Daleska Miclos customer-raawfecnv52myobv@reply.getyourguide.com Teléfono: +34632276997 Idioma: English',
      'Idioma del tour', 'Inglés (Live tour guide)',
      'Precio', '18,30 €'
    ].join('\n')
  },
  // GYG Italian-language tour confirmation.
  gygItalian: {
    subject: 'Booking - S779080 - GYGITALIAN1',
    body: [
      '¡Hola! Buenas noticias.', 'Se ha reservado tu producto',
      'Barcelona Ultimate Tour: Sagrada Familia, Gaudi & Old Town',
      'Número de referencia', 'GYGITALIAN1',
      'Fecha', 'July 20, 2026 11:00 AM',
      'Número de participantes', '2 x Adults (Edad 14 - 99)',
      'Cliente principal', 'Marco Rossi',
      'Teléfono: +390612345678', 'Idioma: Italian',
      'Idioma del tour', 'Italiano (Live tour guide)',
      'Precio', '54,90 €'
    ].join('\n')
  },
  // GYG Italian cancellation (must override / be rejected in confirm mode).
  gygItalianCancel: {
    subject: 'Reserva cancelada - S779080 - GYGITALIAN1',
    body: [
      'Tu reserva ha sido cancelada.',
      'Número de referencia', 'GYGITALIAN1',
      'Idioma del tour', 'Italiano (Live tour guide)'
    ].join('\n')
  },
  // Viator French-language tour confirmation.
  viatorFrench: {
    subject: 'New Booking for Fri, Jul 24, 2026 (#BR-1422956850)',
    body: [
      'You have a new reservation for Complete Barcelona Walking Tour.',
      'Booking Reference: BR-1422956850',
      'Travel Date: Fri, Jul 24, 2026',
      'Lead Traveler Name: Pierre Dupont',
      'Travelers: 3 Adults',
      'Tour Grade: French Tour 17:00',
      'Tour Grade Code: TG3~17:00',
      'Tour Language: French - Guide',
      'Net Rate: EUR €70,00',
      'Phone: +33612345678'
    ].join('\n')
  }
};

function testBookingParsers() {
  resetRunCaches_();
  let pass = 0, fail = 0;
  const check = (label, cond, got) => {
    if (cond) { pass++; console.log('PASS  ' + label); }
    else { fail++; console.log('FAIL  ' + label + '  (got: ' + JSON.stringify(got) + ')'); }
  };

  // GYG Spanish confirmation with children
  let f = RNR_FIXTURES_.gygEsChildren;
  let b = parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('GYG ES: parses', !!b, b);
  if (b) {
    check('GYG ES: booking id', b.bookingId === 'GYGWZARHZ63W', b.bookingId);
    check('GYG ES: 3 adults (not 4)', b.guests === 3, b.guests);
    check('GYG ES: 1 child parsed', b.children === 1, b.children);
    check('GYG ES: notes say child', /1 child/.test(b.notes), b.notes);
    check('GYG ES: name', b.name === 'Fabiola Gardi', b.name);
    check('GYG ES: tour language English (not customer Italian)', b.language === 'English', b.language);
    check('GYG ES: income 54.90*0.75', Math.abs(b.income - 41.18) < 0.02, b.income);
    check('GYG ES: date Jul 20 2026', b.dateKey === '2026-07-20', b.dateKey);
    check('GYG ES: time 11:00 AM', b.time === '11:00 AM', b.time);
    check('GYG ES: valid', isValidBooking_(b), b);
  }

  // GYG private
  f = RNR_FIXTURES_.gygEsPrivate;
  b = parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('GYG priv: parses', !!b, b);
  if (b) {
    check('GYG priv: 5 personas -> 5 adults', b.guests === 5, b.guests);
    check('GYG priv: Private note', /Private/.test(b.notes), b.notes);
    check('GYG priv: no children', b.children === 0, b.children);
    check('GYG priv: income 133.30*0.75', Math.abs(b.income - 99.98) < 0.02, b.income);
  }

  // Viator private product
  f = RNR_FIXTURES_.viatorPrivate;
  b = parseViatorMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('Viator priv: parses', !!b, b);
  if (b) {
    check('Viator priv: id', b.bookingId === 'BR-1422956849', b.bookingId);
    check('Viator priv: 4 adults', b.guests === 4, b.guests);
    check('Viator priv: Private note', /Private/.test(b.notes), b.notes);
    check('Viator priv: German', b.language === 'German', b.language);
    check('Viator priv: time 11:00 AM', b.time === '11:00 AM', b.time);
    check('Viator priv: net 93.50', Math.abs(b.income - 93.5) < 0.01, b.income);
  }

  // Viator adults + children
  f = RNR_FIXTURES_.viatorChildren;
  b = parseViatorMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('Viator kids: parses', !!b, b);
  if (b) {
    check('Viator kids: 2 adults', b.guests === 2, b.guests);
    check('Viator kids: 3 children', b.children === 3, b.children);
    check('Viator kids: notes', /3 children/.test(b.notes), b.notes);
    check('Viator kids: income unchanged by children', Math.abs(b.income - 95.6) < 0.01, b.income);
  }

  // REGRESSION: Gmail plain-body (*bold*) flavour must parse identically.
  f = RNR_FIXTURES_.gygPlainBodyChildren;
  b = parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('GYG plain-body: parses', !!b, b);
  if (b) {
    check('GYG plain-body: 3 adults (was 1 — the bug)', b.guests === 3, b.guests);
    check('GYG plain-body: 1 child recorded', b.children === 1, b.children);
    check('GYG plain-body: notes mention child', /1 child/.test(b.notes), b.notes);
    check('GYG plain-body: name with diacritics', b.name === 'Katarzyna Świstak', b.name);
    check('GYG plain-body: income 54.90*0.75', Math.abs(b.income - 41.18) < 0.02, b.income);
  }

  f = RNR_FIXTURES_.gygPlainBodyPrivate;
  b = parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('GYG plain-body private: parses', !!b, b);
  if (b) {
    check('GYG plain-body private: 5 guests (was 1 — the bug)', b.guests === 5, b.guests);
    check('GYG plain-body private: Private note', /Private/.test(b.notes), b.notes);
    check('GYG plain-body private: no children', b.children === 0, b.children);
    check('GYG plain-body private: income 133.30*0.75', Math.abs(b.income - 99.98) < 0.02, b.income);
  }

  // REGRESSION: customer email inline with the name must be stripped so the
  // Name column holds only the person's name.
  f = RNR_FIXTURES_.gygInlineCustomerEmail;
  b = parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('GYG inline email: parses', !!b, b);
  if (b) {
    check('GYG inline email: name only (no email/labels)', b.name === 'Daleska Miclos', b.name);
    check('GYG inline email: name carries no @ or customer-', !/@|customer-/i.test(b.name), b.name);
    check('GYG inline email: booking id', b.bookingId === 'GYG996ZAK7R7', b.bookingId);
  }

  // Italian tour: parses, routes to Italian Tours, does NOT fall back to English.
  f = RNR_FIXTURES_.gygItalian;
  b = parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('GYG Italian: parses', !!b, b);
  if (b) {
    check('GYG Italian: language Italian (not English)', b.language === 'Italian', b.language);
    check('GYG Italian: routes to Italian Tours', languageToSheet_(b.language) === RNR.SHEETS.ITALIAN, languageToSheet_(b.language));
    check('GYG Italian: 2 adults', b.guests === 2, b.guests);
  }
  // Italian cancellation overrides / is rejected in confirm mode.
  f = RNR_FIXTURES_.gygItalianCancel;
  check('GYG Italian cancel: rejected in confirm mode',
    parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'confirm') === null, 'not null');
  const bItCancel = parseGygMessage_(makeFakeMsg_(f.subject, f.body), 'cancel');
  check('GYG Italian cancel: parsed in cancel mode with isCancellation', !!bItCancel && bItCancel.isCancellation === true, bItCancel);

  // French tour (Viator): parses, routes to French Tours, no English fallback.
  f = RNR_FIXTURES_.viatorFrench;
  b = parseViatorMessage_(makeFakeMsg_(f.subject, f.body), 'confirm');
  check('Viator French: parses', !!b, b);
  if (b) {
    check('Viator French: language French (not English)', b.language === 'French', b.language);
    check('Viator French: routes to French Tours', languageToSheet_(b.language) === RNR.SHEETS.FRENCH, languageToSheet_(b.language));
    check('Viator French: 3 adults', b.guests === 3, b.guests);
  }

  // Language routing units (recognition, regression, unsupported, no fallback).
  check('lang: normalizeLanguage_ Italiano -> Italian', normalizeLanguage_('Italiano (Live tour guide)') === 'Italian', normalizeLanguage_('Italiano (Live tour guide)'));
  check('lang: normalizeLanguage_ Français -> French', normalizeLanguage_('Français') === 'French', normalizeLanguage_('Français'));
  check('lang: normalizeLanguage_ Italian NOT English (no fallback)', normalizeLanguage_('Italian') !== 'English', normalizeLanguage_('Italian'));
  check('lang: normalizeLanguage_ French NOT English (no fallback)', normalizeLanguage_('French') !== 'English', normalizeLanguage_('French'));
  check('lang: regression EN/DE/ES', normalizeLanguage_('Inglés') === 'English' && normalizeLanguage_('Deutsch') === 'German' && normalizeLanguage_('Español') === 'Spanish', 'regression');
  check('lang: unsupported -> English (documented default)', normalizeLanguage_('Klingon') === 'English', normalizeLanguage_('Klingon'));
  check('lang: sheetToLanguage_ round-trips IT/FR', sheetToLanguage_(RNR.SHEETS.ITALIAN) === 'Italian' && sheetToLanguage_(RNR.SHEETS.FRENCH) === 'French', 'roundtrip');

  // Footer noise ("actualizaciones") must NOT flip a confirmation to modify
  f = RNR_FIXTURES_.gygEsChildren;
  b = parseGygMessage_(makeFakeMsg_(f.subject,
        f.body + '\nPuede recibir actualizaciones y novedades en su correo.'), 'confirm');
  check('GYG ES: footer "actualizaciones" still parses as confirmation', !!b && b.bookingId === 'GYGWZARHZ63W', b);

  // Cancellation always wins over confirmation (classification level)
  const cancelTxt = f.body + '\nThe following booking has been cancelled.';
  const bc = parseViatorMessage_(makeFakeMsg_('Cancelled Booking: Mon, Aug 3, 2026', cancelTxt), 'confirm');
  check('Viator: cancel text rejected in confirm mode', bc === null, bc);

  // cleanPersonName_ unit checks — the shared cleaner used by every parser.
  check('cleanName: strips GYG inline email',
    cleanPersonName_('Georgi Nikolov Gitsov customer-zft@reply.getyourguide.com Teléfono: +359 Idioma: English') === 'Georgi Nikolov Gitsov',
    cleanPersonName_('Georgi Nikolov Gitsov customer-zft@reply.getyourguide.com Teléfono: +359 Idioma: English'));
  check('cleanName: strips trailing labels without email',
    cleanPersonName_('Jean-Pierre Dupont Phone: +33 1') === 'Jean-Pierre Dupont',
    cleanPersonName_('Jean-Pierre Dupont Phone: +33 1'));
  check('cleanName: leaves a clean name untouched',
    cleanPersonName_('Katarzyna Świstak') === 'Katarzyna Świstak',
    cleanPersonName_('Katarzyna Świstak'));

  console.log('---------------------------------');
  console.log('RESULT: ' + pass + ' passed, ' + fail + ' failed');
  return fail === 0;
}


/******************************************************
 * 28B. INVARIANT CHECKS
 *
 * Enforced facts, verified on every audit run:
 *   I1  one booking id -> at most one active row (across ALL language tabs)
 *   I2  a booking with a cancellation email is not active
 *   I3  a completed tour (start + 2h) is not in the active tabs
 *   I4  every active row has name, date, time, source, booking id
 * Violations are logged (deduped) — they never throw.
 ******************************************************/

function checkInvariants_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const seenIds = {};
    const problems = [];

    activeSheetNames_().forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
      rows.forEach((row, i) => {
        const b = rowToBooking_(row, sheetName);
        const where = sheetName + ' row ' + (i + 2);

        if (!b.name || !b.bookingId || !b.date || !b.time || !b.source) {
          problems.push('I4 invalid row (' + where + '): missing ' +
            [!b.name && 'name', !b.bookingId && 'booking id', !b.date && 'date',
             !b.time && 'time', !b.source && 'source'].filter(Boolean).join(', '));
          return;
        }
        // I5: a name that still holds an email / "customer-" token / URL means a
        // parser let contamination through (checked on the RAW cell, since
        // rowToBooking_ cleans on read). Catches the class of bug regardless of
        // which OTA template caused it.
        const rawName = String(row[0] || '');
        if (/@|customer-|https?:\/\//i.test(rawName)) {
          problems.push('I5 contaminated name (' + where + '): ' + rawName.slice(0, 60));
        }
        const idKey = b.source + '|' + normalizeId_(b.bookingId);
        if (seenIds[idKey]) {
          problems.push('I1 duplicate active booking ' + b.bookingId + ' (' +
            seenIds[idKey] + ' AND ' + where + ')');
        } else {
          seenIds[idKey] = where;
        }
        if (isCompleted_(b)) {
          problems.push('I3 completed tour still active (' + where + '): ' +
            b.bookingId + ' ' + b.dateKey + ' ' + b.time);
        }
        if (isBookingCancelledByEmail_(b)) {
          problems.push('I2 cancelled booking still active (' + where + '): ' + b.bookingId);
        }
      });
    });

    problems.forEach(p => logError_('INVARIANT ' + p.slice(0, 2), p, ''));
    if (problems.length) console.log('Invariant violations: ' + problems.length);
    return problems.length;
  } catch (e) {
    console.log('checkInvariants_: ' + e);
    return -1;
  }
}


/******************************************************
 * 29. MANAGER DIAGNOSIS + FORCE TOOLS
 *
 * systemStatus()            -> full health report in the execution log
 * forceProcessEverythingNow() -> full audit re-read + status report
 * testInternalAlertEmail()  -> proves this script can send email
 ******************************************************/

/**
 * One-stop diagnosis. Run from the editor, read the Execution log.
 * Checks: Gmail read access (quota), email quota, unprocessed threads per
 * label, installed triggers, newest Errors rows.
 */
function systemStatus() {
  const lines = [];
  resetRunCaches_();
  RNR_SKIP_PROCESSED_ = true;

  // 1. Gmail read probe (fails loudly when the daily read quota is exhausted)
  try {
    const label = GmailApp.getUserLabelByName(RNR.LABELS.GYG_CONFIRM);
    const t = label ? label.getThreads(0, 1) : [];
    if (t.length) getBestMessageText_(t[0].getMessages()[0]);
    lines.push('Gmail read access: OK');
  } catch (e) {
    lines.push('Gmail read access: FAILING -> ' + e +
      '  (usually the daily Gmail quota; it recovers within 24h and unprocessed mail is retried automatically)');
  }

  // 1b. Project timezone (all date math assumes Europe/Madrid)
  try {
    const tz = Session.getScriptTimeZone();
    if (tz !== 'Europe/Madrid') {
      lines.push('WARNING: script timezone is ' + tz + ' — must be Europe/Madrid ' +
        '(Apps Script > Project Settings). Dates/cutoffs will be wrong until fixed.');
    } else {
      lines.push('Timezone: Europe/Madrid OK');
    }
  } catch (e) { /* ignore */ }

  // 2. Email quota
  try { lines.push('Emails this script can still send today: ' + MailApp.getRemainingDailyQuota()); }
  catch (e) { lines.push('Email quota check failed: ' + e); }

  // 3. Unprocessed threads per label (newest 30 per label)
  sourceConfigs_().forEach(cfg => {
    [['confirm', cfg.confirm], ['modify', cfg.modify], ['cancel', cfg.cancel]].forEach(pair => {
      const labelName = pair[1];
      if (!labelName) return;
      try {
        const label = GmailApp.getUserLabelByName(labelName);
        if (!label) { lines.push(labelName + ': LABEL MISSING IN GMAIL'); return; }
        const threads = label.getThreads(0, 30) || [];
        const un = threads.filter(t => !threadHasProcessedLabel_(t)).length;
        if (un) lines.push(labelName + ': ' + un + ' unprocessed (of newest ' + threads.length + ')');
      } catch (e) {
        lines.push(labelName + ': read failed -> ' + e);
      }
    });
  });

  // 4. Triggers
  try {
    const fns = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
    lines.push('Triggers installed here: ' + (fns.join(', ') || 'NONE'));
    if (fns.indexOf('runBookingSystem') === -1) lines.push('WARNING: runBookingSystem has NO trigger');
    if (fns.indexOf('runBookingAudit') === -1) lines.push('WARNING: runBookingAudit has NO trigger');
  } catch (e) { lines.push('Trigger check failed: ' + e); }

  // 5. Newest error rows
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RNR.SHEETS.ERRORS);
    if (sh && sh.getLastRow() > 1) {
      const n = Math.min(5, sh.getLastRow() - 1);
      const v = sh.getRange(sh.getLastRow() - n + 1, 1, n, 3).getDisplayValues();
      lines.push('Newest Errors rows:');
      v.forEach(r => lines.push('   ' + r[0] + ' | ' + r[1] + ' | ' + String(r[2]).slice(0, 140)));
    } else {
      lines.push('Errors tab: empty');
    }
  } catch (e) { lines.push('Errors read failed: ' + e); }

  lines.forEach(l => console.log(l));
  return lines.join('\n');
}

/**
 * FORCE button: re-reads EVERY labelled thread (Processed included), repairs
 * sheet state, then prints the status report. Safe to run any time —
 * everything is idempotent. Use when you see mail without Processed or a
 * missing/duplicated booking.
 */
function forceProcessEverythingNow() {
  runBookingAudit();
  return systemStatus();
}

/**
 * RECOVERY: reprocess ONE booking. Set DEBUG_BOOKING_ID (section 31), run
 * this. It strips the Processed label from every thread of that booking's
 * source labels that mentions the id (bounded search), then runs one fast
 * pass so the normal pipeline re-reads exactly those threads. Never touches
 * unread state, never mass-restores mail.
 */
function reprocessBookingById() {
  const wanted = normalizeId_(DEBUG_BOOKING_ID);
  if (!wanted) { console.log('Set DEBUG_BOOKING_ID first (section 31).'); return; }

  resetRunCaches_();
  let stripped = 0;
  sourceConfigs_().forEach(cfg => {
    [cfg.confirm, cfg.modify, cfg.cancel].forEach(labelName => {
      if (!labelName) return;
      try {
        const threads = GmailApp.search(searchTokenForLabel_(labelName) + ' "' + wanted + '"', 0, 10) || [];
        threads.forEach(t => { safeRemoveLabel_(t, RNR.LABELS.PROCESSED); stripped++; });
      } catch (e) { console.log('reprocess search ' + labelName + ': ' + e); }
    });
  });
  console.log('Processed label removed from ' + stripped + ' thread(s) for ' + wanted +
              '. Running one fast pass now...');
  runBookingSystem();
  console.log('Done. Check the result with debugBooking if needed.');
}


/** Sends a test email to management so you can verify MailApp works. */
function testInternalAlertEmail() {
  MailApp.sendEmail({
    to: RNR.INTERNAL_ALERT_TO,
    subject: 'R&R test email — booking system',
    body: 'If you can read this, the booking script CAN send email. Sent: ' + new Date()
  });
  console.log('Sent. Remaining daily email quota: ' + MailApp.getRemainingDailyQuota());
}


/******************************************************
 * 30. INBOX BACKFILL
 *
 * The system used to archive a booking email as soon as it was processed, so
 * confirmations/modifications for tours that have NOT happened yet ended up
 * out of the inbox. Under the current policy the inbox is the live list of
 * upcoming tours, so those threads must come back.
 *
 * restoreUpcomingThreadsToInbox() walks every Confirmations and Modifications
 * label, works out each thread's tour date, and moves it back to the inbox if
 * the tour has not finished yet. Cancelled threads are never restored.
 * Idempotent: a thread already in the inbox is left alone. Run it once from
 * the editor after deploying; safe to run again any time.
 ******************************************************/

function restoreUpcomingThreadsToInbox() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { console.log('Another run is active; try again in a minute.'); return; }

  RNR_RUN_STARTED_AT_ = Date.now();
  resetRunCaches_();
  RNR_SKIP_PROCESSED_ = false;      // walk EVERYTHING, processed included

  let restored = 0, checked = 0, skippedDone = 0, errors = 0;
  const restoredList = [];

  try {
    sourceConfigs_().forEach(cfg => {
      [cfg.confirm, cfg.modify].filter(Boolean).forEach(labelName => {
        if (!runHasTimeLeft_()) return;
        const label = getLabel_(labelName);
        if (!label) return;

        const threads = label.getThreads(0, RNR.MAX_THREADS_AUDIT) || [];
        threads.forEach(thread => {
          if (!runHasTimeLeft_()) return;
          checked++;
          try {
            // Never resurrect a cancelled booking's thread.
            if (threadHasLabel_(thread, cfg.cancel)) return;
            if (thread.isInInbox()) return;                 // already visible

            if (threadTourIsUpcoming_(thread, cfg.source)) {
              thread.moveToInbox();
              restored++;
              if (restoredList.length < 40) {
                restoredList.push(thread.getFirstMessageSubject() || '(no subject)');
              }
            } else {
              skippedDone++;
            }
          } catch (e) {
            errors++;
            console.log('restore skip: ' + e);
          }
        });
      });
    });
  } catch (err) {
    logError_('restoreUpcomingThreadsToInbox', err, '');
  } finally {
    lock.releaseLock();
  }

  console.log('Threads checked: ' + checked);
  console.log('Restored to inbox (upcoming tours): ' + restored);
  console.log('Left archived (tour already finished): ' + skippedDone);
  console.log('Errors: ' + errors);
  if (restoredList.length) {
    console.log('--- restored ---');
    restoredList.forEach(s => console.log('  ' + s));
    if (restored > restoredList.length) console.log('  ...and ' + (restored - restoredList.length) + ' more');
  }
  return restored;
}

/**
 * True when a thread's tour has NOT finished yet. Uses the parsed bookings
 * first, then the subject/body date fallback. Unknown date -> treated as
 * upcoming, so we never hide something we could not read.
 */
function threadTourIsUpcoming_(thread, source) {
  try {
    const bookings = uniqueBookings_(parseThread_(thread, source, 'any')).filter(isValidBooking_);
    if (bookings.length) return bookings.some(b => !isCompleted_(b));

    const d = extractTourDateFromThread_(thread);
    if (!d) return true;                       // unreadable -> show it
    const t = extractTourTimeFromThread_(thread) || RNR.DEFAULT_TIME;
    return !isCompleted_(normalizeBooking_({
      name: 'x', bookingId: 'x', date: d, time: t,
      language: RNR.LANGUAGE.ENGLISH, source,
      hasExplicitDate: true, hasExplicitTime: true
    }));
  } catch (e) {
    return true;                               // on doubt, show it
  }
}


/******************************************************
 * 31. SINGLE-BOOKING DIAGNOSIS + FORCE REPAIR
 *
 * Use these when a row on the sheet does not match the email.
 *
 *   debugBooking('GYG996ZFRKQ9')   -> prints EXACTLY what the script reads
 *                                     and parses for that booking.
 *   reparseActiveRowsFromEmail()   -> re-reads every active row's own
 *                                     confirmation email and rewrites the
 *                                     row. Fixes historical rows in one go.
 ******************************************************/

/**
 * Print the full parsing story for ONE booking id. Read the Execution log
 * top to bottom: it shows which label the thread sits under, which body
 * flavour was chosen and why, the exact participants window the parser saw,
 * and the final parsed booking. This is the fastest way to prove whether a
 * wrong guest count is a parser problem or a "never re-processed" problem.
 */
// Edit this id, then run debugBooking from the editor (the editor cannot
// pass arguments to a function).
const DEBUG_BOOKING_ID = 'GYG996ZFRKQ9';

function debugBooking(bookingId) {
  const wanted = normalizeId_(bookingId || DEBUG_BOOKING_ID);
  if (!wanted) { console.log('Set DEBUG_BOOKING_ID at the top of section 31, then run again.'); return; }

  resetRunCaches_();
  RNR_SKIP_PROCESSED_ = false;
  let found = 0;

  sourceConfigs_().forEach(cfg => {
    [['CONFIRM', cfg.confirm], ['MODIFY', cfg.modify], ['CANCEL', cfg.cancel]].forEach(pair => {
      const kind = pair[0], labelName = pair[1];
      if (!labelName) return;
      const label = getLabel_(labelName);
      if (!label) return;

      (label.getThreads(0, RNR.MAX_THREADS_AUDIT) || []).forEach(thread => {
        const subject = thread.getFirstMessageSubject() || '';
        thread.getMessages().forEach(msg => {
          let html = '', plain = '';
          try { html = htmlToText_(msg.getBody() || ''); } catch (e) { html = '(html read failed: ' + e + ')'; }
          try { plain = String(msg.getPlainBody() || '').trim(); } catch (e) { plain = '(plain read failed: ' + e + ')'; }
          const chosen = pickBestBody_(html, plain);
          const full = subject + '\n' + chosen;
          if (full.toUpperCase().indexOf(wanted) === -1) return;

          found++;
          console.log('\n==================================================');
          console.log('FOUND under: ' + labelName + '  (' + kind + ')');
          console.log('Subject    : ' + subject);
          console.log('In inbox   : ' + thread.isInInbox() + '   | Processed label: ' + threadHasProcessedLabel_(thread));
          console.log('Body sizes : html=' + html.length + ' chars, plain=' + plain.length + ' chars');
          console.log('Body used  : ' + (chosen === html.replace(/\*+/g, '') ? 'HTML' : 'PLAIN') +
                      ' (whichever contained the participants block)');

          // The exact window the participant parser works on.
          const lines = String(chosen).split('\n').map(l => l.trim());
          let idx = -1;
          for (let i = 0; i < lines.length; i++) {
            if (/^(N[uú]mero de participantes|Participantes|Participants?|Number of participants|Travelers?:|Travellers?:)/i.test(lines[i])) { idx = i; break; }
          }
          if (idx === -1) {
            console.log('PARTICIPANTS BLOCK: *** NOT FOUND *** — this is why the guest count is wrong.');
            console.log('First 40 lines of the text actually parsed:');
            lines.slice(0, 40).forEach((l, i) => console.log('   ' + i + ': [' + l + ']'));
          } else {
            console.log('Participants window (what the parser sees):');
            lines.slice(idx, idx + 6).forEach((l, i) => console.log('   ' + (idx + i) + ': [' + l + ']'));
            console.log('participantCounts_ -> ' + JSON.stringify(participantCounts_(lines.slice(idx, idx + 6).join('\n'))));
          }

          if (cfg.source === RNR.SOURCE.GYG) {
            console.log('gygParticipants_   -> ' + JSON.stringify(gygParticipants_(full)));
            console.log('gygGuests_ (legacy)-> ' + gygGuests_(full));
          }

          const parsed = parseThread_(thread, cfg.source, 'any')
            .filter(b => normalizeId_(b.bookingId) === wanted)[0];
          console.log('PARSED BOOKING     -> ' + JSON.stringify(parsed, null, 2));

          const row = findActiveBookingById_(cfg.source, wanted);
          console.log('CURRENT SHEET ROW  -> ' + (row ? JSON.stringify({
            name: row.name, guests: row.guests, children: row.children,
            notes: row.notes, date: row.dateKey, time: row.time
          }) : '(no active row)'));
          if (parsed && row && Number(parsed.guests) !== Number(row.guests)) {
            console.log('>>> MISMATCH: email says ' + parsed.guests + ' guests, sheet says ' + row.guests +
                        '. Run reparseActiveRowsFromEmail() to rewrite the sheet from the emails.');
          }
        });
      });
    });
  });

  if (!found) console.log('No thread found containing ' + wanted + ' under any RNR label.');
}


/**
 * Re-read every ACTIVE booking row's own confirmation email and rewrite the
 * row from it. Use after a parser fix so historical rows pick up the
 * correction without waiting for anything. Idempotent: rows are updated in
 * place (never duplicated), and cancelled/completed bookings are untouched.
 */
function reparseActiveRowsFromEmail() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { console.log('Another run is active; try again in a minute.'); return; }

  RNR_RUN_STARTED_AT_ = Date.now();
  resetRunCaches_();
  RNR_SKIP_PROCESSED_ = false;   // read everything, Processed included

  let checked = 0, fixed = 0, unchanged = 0, noEmail = 0;
  const changes = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Index every confirmation booking once (source|id -> parsed booking).
    const cache = getConfirmationCache_();

    activeSheetNames_().forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;

      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
      rows.forEach((row, i) => {
        const current = rowToBooking_(row, sheetName);
        if (!current.bookingId || !current.source) return;
        checked++;

        const parsed = cache.get(current.source + '|' + normalizeId_(current.bookingId));
        if (!parsed) { noEmail++; return; }

        const newGuests = Number(parsed.guests || 0);
        const newChildren = Number(parsed.children || 0);
        const newNotes = String(parsed.notes || '');

        // Only rewrite when the email genuinely disagrees with the sheet.
        const differs = newGuests !== Number(current.guests) ||
                        newChildren !== Number(current.children) ||
                        (newNotes && newNotes !== String(current.notes || ''));
        if (!differs) { unchanged++; return; }

        // Preserve any manual note a manager typed (e.g. "Moved to 21 ...").
        const manualNote = String(current.notes || '')
          .split(' · ')
          .filter(part => !/^Private$/i.test(part) && !/\d+\s*(child|infant)/i.test(part))
          .join(' · ');
        const mergedNotes = [newNotes, manualNote].filter(Boolean).join(' · ');

        const updated = normalizeBooking_({
          name: current.name || parsed.name,
          phone: current.phone || parsed.phone,
          guests: newGuests,
          children: newChildren,
          infants: Number(parsed.infants || 0),
          date: current.date,           // keep the sheet's date/time: a
          time: current.time,           // modification may have moved it
          language: current.language,
          source: current.source,
          income: Number(parsed.income || current.income || 0),
          bookingId: current.bookingId,
          notes: mergedNotes,
          hasExplicitGuests: true, hasExplicitDate: true,
          hasExplicitTime: true, hasExplicitIncome: true
        });

        writeBookingRow_(sh, i + 2, updated);
        fixed++;
        changes.push(current.bookingId + ' (' + current.name + '): guests ' +
                     current.guests + ' -> ' + newGuests +
                     (newChildren ? ', children ' + newChildren : '') );
      });
    });
  } catch (err) {
    logError_('reparseActiveRowsFromEmail', err, '');
    console.log(String(err && err.stack ? err.stack : err));
  } finally {
    lock.releaseLock();
  }

  console.log('Active rows checked      : ' + checked);
  console.log('Rows corrected from email: ' + fixed);
  console.log('Already correct          : ' + unchanged);
  console.log('No confirmation email found (left alone): ' + noEmail);
  if (changes.length) {
    console.log('--- corrections ---');
    changes.forEach(c => console.log('  ' + c));
  }
  return fixed;
}


