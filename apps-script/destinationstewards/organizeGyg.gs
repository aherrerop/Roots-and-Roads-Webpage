/*  ============================================================================
 *  GetYourGuide inbox organiser — for the SECOND GYG account
 *  (destinationstewards@gmail.com, the "3-in-1" listing).
 *
 *  This is a STANDALONE Apps Script that lives IN the destinationstewards
 *  account. It mirrors how the main rootsandroadstours booking script tidies
 *  Gmail — it labels each GetYourGuide email by type (Confirmation /
 *  Cancellation / Modification), marks it read, tags it Processed, and MOVES a
 *  cancelled booking's confirmation into Cancellations. It does NOT touch any
 *  spreadsheet: destinationstewards keeps forwarding bookings to
 *  rootsandroadstours, where the real booking system still records them.
 *
 *  HOW TO INSTALL (do this while logged into destinationstewards@gmail.com):
 *    1. Go to  script.google.com  →  New project.
 *    2. Delete the sample code, paste THIS whole file, and Save.
 *    3. Run  installTrigger  once (top menu ▶). Approve the Gmail permission
 *       when asked — that authorises the script to read/label your mail.
 *    4. Done. It now runs every 5 minutes. To tidy everything already sitting
 *       in the mailbox right now, also Run  organizeGygInbox  once.
 *
 *  It never throws (so Google never emails you a failure notice) and is
 *  idempotent (safe to run as often as you like).
 *  ==========================================================================*/

var CFG = {
  // Only emails from GetYourGuide are touched. (Booking notifications come from
  // do-not-reply@notification.getyourguide.com.)
  SENDER_MATCH: 'getyourguide.com',

  // Label names. Change these if you prefer a different structure — nested
  // labels use "/" (e.g. "Bookings/Confirmations" shows as Bookings ▸ Confirmations).
  LABELS: {
    CONFIRM: 'Bookings/Confirmations',
    CANCEL:  'Bookings/Cancellations',
    MODIFY:  'Bookings/Modifications',
    DONE:    'Bookings/Done',
    PROCESSED: 'Bookings/Processed'   // marker so a thread is handled only once
  },

  // How many threads to scan per run (Gmail search cap; well above a day's volume).
  MAX_THREADS: 200,

  // A confirmation whose tour date is already past is moved to Done + archived.
  MOVE_PAST_TO_DONE: true
};

/* ----------------------------------------------------------------------------
 *  MAIN — classify + label + mark read + Processed. Runs on the 5-min trigger.
 * --------------------------------------------------------------------------*/
function organizeGygInbox() {
  try {
    var L = ensureLabels_();
    // Every GetYourGuide thread not yet marked Processed.
    var query = 'from:(' + CFG.SENDER_MATCH + ') -label:"' + CFG.LABELS.PROCESSED + '"';
    var threads = GmailApp.search(query, 0, CFG.MAX_THREADS) || [];

    threads.forEach(function (thread) {
      try {
        var type = classifyThread_(thread);         // 'cancel' | 'modify' | 'confirm'
        var typeLabel = type === 'cancel' ? L.CANCEL : (type === 'modify' ? L.MODIFY : L.CONFIRM);

        thread.addLabel(typeLabel);
        thread.addLabel(L.PROCESSED);
        thread.markRead();

        if (type === 'cancel') {
          // A cancellation is done: file it under Cancellations and get it out of
          // the inbox — AND move the ORIGINAL confirmation for the same booking to
          // Cancellations too, exactly like the main script does.
          thread.moveToArchive();
          moveConfirmationToCancellations_(thread, L);
        }
        // Confirmations: kept VISIBLE in the inbox as your live upcoming list (the
        // Done sweep below files them once the tour date passes). Modifications:
        // labelled + read + Processed, left in the inbox (the booking is still live).
      } catch (eThread) {
        // One bad thread must never stop the rest.
        console.log('thread skipped: ' + (eThread && eThread.message));
      }
    });

    // SWEEP: any CONFIRMATION whose tour date has now passed -> Done + archive, so
    // the inbox only ever holds UPCOMING bookings. Re-scans the label each run so a
    // future booking is moved on the FIRST run after its tour date, not just when
    // it first arrived.
    if (CFG.MOVE_PAST_TO_DONE) {
      var conf = GmailApp.search('label:"' + CFG.LABELS.CONFIRM + '"', 0, CFG.MAX_THREADS) || [];
      conf.forEach(function (t) {
        try {
          if (tourIsPast_(threadText_(t))) {
            t.removeLabel(L.CONFIRM); t.addLabel(L.DONE); t.moveToArchive();
          }
        } catch (e) { /* skip */ }
      });
    }
  } catch (e) {
    // Trigger entry points must never throw (or Google emails a failure summary).
    console.log('organizeGygInbox error (swallowed): ' + (e && e.stack ? e.stack : e));
  }
}

/* ----------------------------------------------------------------------------
 *  Classify a thread from its SUBJECT (never the body: a confirmation's body can
 *  mention "cancel" in boilerplate). Order matters: cancel, then modify, else confirm.
 * --------------------------------------------------------------------------*/
function classifyThread_(thread) {
  var subj = String(thread.getFirstMessageSubject() || '');
  if (/\bcancel/i.test(subj) || /cancelad|cancelaci[oó]n|anulad|storn/i.test(subj)) return 'cancel';
  if (/detail change|\bchange\b|modif|cambio|reprogramad|updated|ge[aä]ndert/i.test(subj)) return 'modify';
  return 'confirm';
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
      if (t.getId() === cancelThread.getId()) return;                 // that's the cancellation itself
      if (classifyThread_(t) !== 'confirm') return;                   // only the confirmation
      t.removeLabel(L.CONFIRM);
      t.addLabel(L.CANCEL);
      t.addLabel(L.PROCESSED);
      t.markRead();
      t.moveToArchive();
    } catch (e) { /* skip a bad match */ }
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

/** The GetYourGuide booking reference, e.g. "GYG83W7LWAGA". */
function extractGygRef_(text) {
  var m = String(text || '').match(/\bGYG[0-9A-Z]{6,}\b/i);
  return m ? m[0].toUpperCase() : '';
}

/** True if the tour date in the confirmation text is before today. Best-effort:
 *  GYG confirmations render the date in English ("October 2, 2026, 3:30 PM"). */
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
  // Remove any existing trigger for this function, then add a fresh 5-min one.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'organizeGygInbox') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('organizeGygInbox').timeBased().everyMinutes(5).create();
  organizeGygInbox();   // and do a first pass right now
  return 'Installed: organizeGygInbox runs every 5 minutes.';
}
