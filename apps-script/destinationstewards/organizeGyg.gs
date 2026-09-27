/*  ============================================================================
 *  GetYourGuide inbox organiser — for the SECOND GYG account
 *  (destinationstewards@gmail.com, the "3-in-1" listing).
 *
 *  This is a STANDALONE Apps Script that lives IN the destinationstewards
 *  account. It mirrors how the main rootsandroadstours booking script tidies
 *  Gmail: it labels each GetYourGuide BOOKING email by type (Confirmation /
 *  Cancellation / Modification), marks it read, tags it Processed, moves a
 *  cancelled booking's confirmation into Cancellations, and files past tours in
 *  Done. It does NOT touch any spreadsheet: destinationstewards keeps forwarding
 *  bookings to rootsandroadstours, where the real booking system records them.
 *
 *  WHAT WAS WRONG BEFORE (and is fixed here):
 *   1. It matched the broad sender "getyourguide.com", so NON-booking GYG mail
 *      (payouts, reviews, marketing) fell through to the "confirm" default and
 *      got the Confirmations label. Now it only touches the booking-notification
 *      sender, and only acts on emails that actually look like a booking.
 *   2. It only ever looked at threads NOT yet marked Processed. Anything a naive
 *      Gmail filter had already labelled Confirmations + marked read was never
 *      re-examined, so cancellations/modifications stayed stuck under
 *      Confirmations. It now ALSO self-heals: every run it re-checks the
 *      Confirmations label and moves any cancel/modify email to the right place.
 *
 *  HOW TO INSTALL (while logged into destinationstewards@gmail.com):
 *    1. script.google.com -> New project.
 *    2. Delete the sample code, paste THIS whole file, Save.
 *    3. Run  installTrigger  once (▶). Approve the Gmail permission when asked.
 *    4. Done — it runs every 5 minutes. To tidy the whole mailbox right now,
 *       also Run  organizeGygInbox  once. (Run  runSelfTest  to sanity-check the
 *       classifier without touching any mail.)
 *
 *  It never throws (so Google never emails a failure notice) and is idempotent
 *  (safe to run as often as you like).
 *  ==========================================================================*/

var CFG = {
  // ONLY GetYourGuide BOOKING notifications are touched. They all come from
  // do-not-reply@notification.getyourguide.com. Matching this exact sub-domain
  // (not the whole "getyourguide.com") keeps payout / review / marketing mail out.
  SENDER_MATCH: 'notification.getyourguide.com',

  // Label names — the SAME structure this account already uses, so nothing is
  // duplicated. Nested labels use "/".
  LABELS: {
    CONFIRM:   'Publishing Pages/GetYourGuide/Confirmations',
    CANCEL:    'Publishing Pages/GetYourGuide/Cancellations',
    MODIFY:    'Publishing Pages/GetYourGuide/Modifications',
    DONE:      'Publishing Pages/GetYourGuide/Done',
    PROCESSED: 'Publishing Pages/Processed'   // marker so a thread is handled once
  },

  // How many threads to scan per run (Gmail search cap; well above a day's volume).
  MAX_THREADS: 250,

  // A confirmation whose tour date is already past is moved to Done + archived.
  MOVE_PAST_TO_DONE: true
};

// Gmail sub-query that finds cancel/modify emails hiding under the Confirmations
// label (a naive filter's mistake). Kept in sync with classifyGygSubject_.
var MISLABEL_Q_ = 'subject:(canceled OR cancelled OR cancelada OR anulada OR ' +
                  '"detail change" OR change OR changed OR modif OR modificaci OR cambio)';

/* ----------------------------------------------------------------------------
 *  MAIN — runs on the 5-min trigger.
 * --------------------------------------------------------------------------*/
function organizeGygInbox() {
  try {
    var L = ensureLabels_();

    // (1) NEW booking notifications not yet handled -> classify, label, read, Processed.
    handleThreads_(GmailApp.search(
      'from:(' + CFG.SENDER_MATCH + ') -label:"' + CFG.LABELS.PROCESSED + '"', 0, CFG.MAX_THREADS), L);

    // (2) SELF-HEAL — the fix for "almost everything is a confirmation": re-check
    //     every thread currently under Confirmations whose subject looks like a
    //     cancel/modify, and move it to the correct label. Catches emails a Gmail
    //     filter mislabelled (even if it already marked them read / Processed).
    handleThreads_(GmailApp.search(
      'label:"' + CFG.LABELS.CONFIRM + '" ' + MISLABEL_Q_, 0, CFG.MAX_THREADS), L);

    // (3) Past-tour confirmations -> Done + archive, so the inbox holds only
    //     UPCOMING bookings (re-scanned each run, so a booking is filed on the
    //     first run AFTER its tour date).
    if (CFG.MOVE_PAST_TO_DONE) sweepPastConfirmations_(L);

  } catch (e) {
    // Trigger entry points must never throw (or Google emails a failure summary).
    console.log('organizeGygInbox error (swallowed): ' + (e && e.stack ? e.stack : e));
  }
}

/* ----------------------------------------------------------------------------
 *  Classify + apply labels to a batch of threads.
 * --------------------------------------------------------------------------*/
function handleThreads_(threads, L) {
  (threads || []).forEach(function (thread) {
    try {
      var subj = String(thread.getFirstMessageSubject() || '');
      // Not a booking notification (a stray payout/review/etc.) -> leave it alone.
      if (!isBookingNotification_(subj)) return;
      applyClassification_(thread, classifyGygSubject_(subj), L);
    } catch (eThread) {
      // One bad thread must never stop the rest.
      console.log('thread skipped: ' + (eThread && eThread.message));
    }
  });
}

/** Set the ONE correct label, remove the other two, mark read + Processed. This
 *  is the authority — it corrects whatever a filter put on the thread. */
function applyClassification_(thread, type, L) {
  thread.markRead();
  thread.addLabel(L.PROCESSED);

  if (type === 'cancel') {
    // A cancellation is done: file under Cancellations, strip any wrong labels,
    // archive it, AND move the ORIGINAL confirmation for the same booking too.
    thread.addLabel(L.CANCEL);
    thread.removeLabel(L.CONFIRM);
    thread.removeLabel(L.MODIFY);
    thread.moveToArchive();
    moveConfirmationToCancellations_(thread, L);

  } else if (type === 'modify') {
    // A modification is its own email — label Modifications, drop any stray
    // Confirmations. Stays in the inbox (the booking is still live).
    thread.addLabel(L.MODIFY);
    thread.removeLabel(L.CONFIRM);
    thread.removeLabel(L.CANCEL);

  } else {
    // A new booking. Kept VISIBLE in the inbox as the live upcoming list; the
    // Done sweep files it once its tour date passes.
    thread.addLabel(L.CONFIRM);
    thread.removeLabel(L.CANCEL);
    thread.removeLabel(L.MODIFY);
  }
}

/* ----------------------------------------------------------------------------
 *  Classification — from the SUBJECT only (a confirmation's BODY can say
 *  "cancel" in boilerplate). GetYourGuide's real subjects:
 *     confirm : "Booking - S809442 - GYG..."
 *     cancel  : "A booking has been canceled - S779080 - GYG..."
 *     modify  : "Booking detail change: - S779080 - GYG..."
 *  Order matters: cancel, then modify, else confirm. Extra ES/DE/FR/IT stems are
 *  belt-and-braces in case GYG ever sends this account in another language.
 * --------------------------------------------------------------------------*/
function classifyGygSubject_(subject) {
  var s = String(subject || '');
  if (/cancel|cancelad|cancelaci[oó]n|anulad|storn|annul/i.test(s)) return 'cancel';
  if (/detail change|has changed|\bchanged?\b|modif|cambio|reprogramad|updated|ge[aä]ndert/i.test(s)) return 'modify';
  return 'confirm';
}

/** Does this subject look like a GYG BOOKING notification at all (vs a payout,
 *  review, marketing blast)? Only these are labelled. */
function isBookingNotification_(subject) {
  return /booking|reserva|buchung|prenotazione|r[ée]servation|has been cancel|detail change|cambio|modif/i.test(String(subject || ''));
}

/* ----------------------------------------------------------------------------
 *  When a cancellation arrives, move the matching CONFIRMATION (same GYG
 *  reference) out of the inbox / Confirmations and into Cancellations.
 * --------------------------------------------------------------------------*/
function moveConfirmationToCancellations_(cancelThread, L) {
  var ref = extractGygRef_(threadText_(cancelThread));
  if (!ref) return;
  var hits = GmailApp.search('from:(' + CFG.SENDER_MATCH + ') "' + ref + '"', 0, 10) || [];
  hits.forEach(function (t) {
    try {
      if (t.getId() === cancelThread.getId()) return;                 // the cancellation itself
      if (classifyGygSubject_(t.getFirstMessageSubject() || '') !== 'confirm') return;  // only the confirmation
      t.removeLabel(L.CONFIRM);
      t.removeLabel(L.MODIFY);
      t.addLabel(L.CANCEL);
      t.addLabel(L.PROCESSED);
      t.markRead();
      t.moveToArchive();
    } catch (e) { /* skip a bad match */ }
  });
}

/* ----------------------------------------------------------------------------
 *  Sweep: any CONFIRMATION whose tour date has passed -> Done + archive.
 * --------------------------------------------------------------------------*/
function sweepPastConfirmations_(L) {
  var conf = GmailApp.search('label:"' + CFG.LABELS.CONFIRM + '"', 0, CFG.MAX_THREADS) || [];
  conf.forEach(function (t) {
    try {
      if (tourIsPast_(threadText_(t))) { t.removeLabel(L.CONFIRM); t.addLabel(L.DONE); t.moveToArchive(); }
    } catch (e) { /* skip */ }
  });
}

/* ----------------------------------------------------------------------------
 *  Helpers
 * --------------------------------------------------------------------------*/

/** Create the labels if they don't exist yet; return them keyed like CFG.LABELS. */
function ensureLabels_() {
  var out = {};
  Object.keys(CFG.LABELS).forEach(function (k) {
    var name = CFG.LABELS[k];
    out[k] = GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
  });
  return out;
}

/** All text of a thread (subject + bodies), for reference/date extraction. */
function threadText_(thread) {
  return thread.getMessages().map(function (m) {
    return (m.getSubject() || '') + '\n' + (m.getPlainBody() || '');
  }).join('\n');
}

/** The GetYourGuide booking reference, e.g. "GYGRFQFF38ZH". */
function extractGygRef_(text) {
  var m = String(text || '').match(/\bGYG[0-9A-Z]{6,}\b/i);
  return m ? m[0].toUpperCase() : '';
}

/** True if the tour date in the confirmation text is before now. Best-effort:
 *  GYG renders the date in English ("October 1, 2026, 3:30 PM"). */
function tourIsPast_(text) {
  var m = String(text || '').match(
    /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s*(\d{4})/i);
  if (!m) return false;                       // can't tell -> leave it in the inbox
  var d = new Date(m[1] + ' ' + m[2] + ', ' + m[3] + ' 23:59:59');
  if (isNaN(d)) return false;
  return d.getTime() < Date.now();
}

/* ----------------------------------------------------------------------------
 *  Trigger setup — run ONCE from the editor.
 * --------------------------------------------------------------------------*/
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'organizeGygInbox') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('organizeGygInbox').timeBased().everyMinutes(5).create();
  organizeGygInbox();   // and do a first pass right now
  return 'Installed: organizeGygInbox runs every 5 minutes.';
}

/* ----------------------------------------------------------------------------
 *  Self-test — run from the editor to verify the classifier against real GYG
 *  subjects WITHOUT touching any mail. Returns "ALL PASS" or the failures.
 * --------------------------------------------------------------------------*/
function runSelfTest() {
  var cases = [
    ['Booking - S809442 - GYGRFQFF38ZH', 'confirm'],
    ['Booking - S779080 - GYG32LXXYFFF', 'confirm'],
    ['A booking has been canceled - S779080 - GYGG45K447X5', 'cancel'],
    ['A booking has been canceled - S809442 - GYGZGZQ942QB', 'cancel'],
    ['Booking detail change: - S779080 - GYG83W7WQG48', 'modify'],
    ['Booking detail change: - S779080 - GYGMX38R9X85', 'modify']
  ];
  var fails = [];
  cases.forEach(function (c) {
    var got = classifyGygSubject_(c[0]);
    if (got !== c[1] || !isBookingNotification_(c[0])) fails.push(c[0] + ' -> ' + got + ' (want ' + c[1] + ')');
  });
  // A non-booking notification must be skipped (not labelled).
  if (isBookingNotification_('You have received a payout')) fails.push('payout wrongly treated as a booking');
  var msg = fails.length ? ('FAIL:\n' + fails.join('\n')) : 'ALL PASS';
  console.log(msg);
  return msg;
}
