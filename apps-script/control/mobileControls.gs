/******************************************************
 * ROOTS & ROADS — mobileControls.gs
 * Bind to: Roots_Roads_Control_v1 (same project as assignShifts.gs +
 * guidePortal.gs).
 *
 * WHAT THIS DOES
 *   Lets Albert and Carlos run the main operations from the Google Sheets
 *   PHONE APP by ticking a checkbox — no laptop, no Apps Script editor.
 *
 * SETUP (once, from a computer)
 *   1. Run setupMobileControls() from the Apps Script editor. It:
 *        - creates the "Control" tab (if missing),
 *        - writes the Mobile Controls block at Control!N2:P12,
 *        - installs the installable "On edit" trigger.
 *   2. For the two booking actions, set Script Properties (Project Settings):
 *        BOOKING_WEBAPP_URL = the BookingSheet web app /exec URL
 *        ADMIN_KEY          = any long random string (same value must be set
 *                             as ADMIN_KEY in the BookingSheet project too)
 *      Without them, those two rows show "Set BOOKING_WEBAPP_URL first".
 *
 * HOW IT WORKS
 *   Tick a checkbox in column O -> status cell (column P) shows "Running…",
 *   then "Done: <time>" or a short error. The checkbox resets itself.
 *   Script-driven resets do NOT retrigger anything (Apps Script only fires
 *   onEdit for HUMAN edits), and LockService blocks two managers launching
 *   the same run simultaneously.
 ******************************************************/

const MC = {
  TAB: 'Control',
  // Functions block sits at the TOP-LEFT for easy access; the SYSTEM HEALTH
  // block is written below it (see updateControlHealth_ / HEALTH_FIRST_ROW).
  RANGE: 'A1:C12',
  HEADER_ROW: 1,     // A1:C1 = header
  FIRST_ACTION_ROW: 2,
  COL_ACTION: 1,     // A
  COL_RUN: 2,        // B (checkbox)
  COL_STATUS: 3      // C
};

/**
 * Action rows, in display order. `fn` runs in THIS project; `remote` calls
 * the BookingSheet web app (where runBookingSystem/runBookingAudit live).
 */
function mcActions_() {
  return [
    { label: 'Run booking update', remote: 'runBookingSystem' },
    { label: 'Run booking audit', remote: 'runBookingAudit' },
    { label: 'Sync availability', fn: syncAvailabilityFile },
    { label: 'Generate schedules', fn: makeSchedule },
    { label: 'Refresh ledger & queues', fn: updateManagementQueues },
    { label: 'Weekly full run (avail + schedule + email)', fn: runWeeklyScheduling },
    { label: 'Full operational refresh', fn: mcFullRefresh_ }
  ];
}

/**
 * Full refresh that stays inside execution limits: booking update runs
 * REMOTELY (its own execution in the BookingSheet project), and the local
 * part is sync + schedule + queues.
 */
function mcFullRefresh_() {
  mcCallRemote_('runBookingSystem');   // fire-and-report; separate quota
  syncAvailabilityFile();
  makeSchedule();
  updateManagementQueues();
}

/** One-time setup: tab, block, checkboxes, trigger. Safe to re-run. */
function setupMobileControls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MC.TAB);
  if (!sh) sh = ss.insertSheet(MC.TAB);

  // Migration: the functions block used to live at N2:P12. Clear that old
  // location (values + leftover checkboxes) now that it lives top-left at A1.
  sh.getRange(1, 14, 13, 3).clearContent().clearDataValidations();

  const actions = mcActions_();
  sh.getRange(MC.HEADER_ROW, MC.COL_ACTION, 1, 3)
    .setValues([['Action', 'Run', 'Status / last result']])
    .setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');

  const rows = actions.map(a => [a.label, false, '']);
  const block = sh.getRange(MC.FIRST_ACTION_ROW, MC.COL_ACTION, rows.length, 3);
  // Preserve existing status text if the block already exists.
  const existing = block.getValues();
  rows.forEach((r, i) => { if (existing[i] && existing[i][2]) r[2] = existing[i][2]; });
  block.setValues(rows);
  sh.getRange(MC.FIRST_ACTION_ROW, MC.COL_RUN, rows.length, 1).insertCheckboxes();
  // Widths shared with the health block below (same columns A/B/C).
  sh.setColumnWidth(MC.COL_ACTION, 300);
  sh.setColumnWidth(MC.COL_RUN, 150);
  sh.setColumnWidth(MC.COL_STATUS, 240);

  // Installable ON EDIT trigger (a simple onEdit cannot call other services).
  const exists = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'handleMobileControlsEdit');
  if (!exists) {
    ScriptApp.newTrigger('handleMobileControlsEdit').forSpreadsheet(ss).onEdit().create();
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Mobile Controls ready at ' + MC.TAB + '!' + MC.RANGE);
}

/** Installable On-edit handler. Only reacts to the Run checkboxes. */
function handleMobileControlsEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== MC.TAB) return;
    if (e.range.getColumn() !== MC.COL_RUN || e.range.getNumColumns() !== 1) return;
    const row = e.range.getRow();
    const actions = mcActions_();
    const idx = row - MC.FIRST_ACTION_ROW;
    if (idx < 0 || idx >= actions.length) return;
    if (e.value !== 'TRUE' && e.range.getValue() !== true) return;  // only tick -> TRUE

    const action = actions[idx];
    const statusCell = sh.getRange(row, MC.COL_STATUS);

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(2000)) {
      statusCell.setValue('Busy: another operation is running');
      e.range.setValue(false);
      return;
    }

    statusCell.setValue('Running…');
    SpreadsheetApp.flush();

    try {
      let note = '';
      if (action.remote) note = mcCallRemote_(action.remote);
      else action.fn();
      statusCell.setValue('Done: ' +
        Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm') +
        (note ? ' — ' + note : ''));
    } catch (err) {
      const msg = String(err && err.message ? err.message : err);
      statusCell.setValue('Error: ' + msg.slice(0, 90));
      console.error('Mobile control "' + action.label + '" failed: ' +
        (err && err.stack ? err.stack : err));
    } finally {
      // Script-driven reset: does NOT fire onEdit again (only human edits do).
      e.range.setValue(false);
      lock.releaseLock();
    }
  } catch (outer) {
    console.error('handleMobileControlsEdit: ' + outer);
  }
}

/** Call runBookingSystem / runBookingAudit in the BookingSheet project. */
function mcCallRemote_(fnName) {
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('BOOKING_WEBAPP_URL');
  const key = props.getProperty('ADMIN_KEY');
  if (!url || !key) throw new Error('Set BOOKING_WEBAPP_URL + ADMIN_KEY first (Project Settings)');
  const res = UrlFetchApp.fetch(
    url + '?admin=1&key=' + encodeURIComponent(key) + '&fn=' + encodeURIComponent(fnName),
    { muteHttpExceptions: true, followRedirects: true });
  const code = res.getResponseCode();
  const body = res.getContentText() || '';
  if (code >= 300) throw new Error('Remote HTTP ' + code);
  let parsed = null;
  try { parsed = JSON.parse(body); } catch (err) { /* non-JSON */ }
  if (parsed && parsed.ok === false) throw new Error(parsed.error || 'Remote error');
  return 'remote ok';
}

/** Diagnostic: verifies trigger, block, and remote configuration. */
function validateMobileControls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const problems = [];

  const sh = ss.getSheetByName(MC.TAB);
  if (!sh) problems.push('Tab "' + MC.TAB + '" missing — run setupMobileControls()');
  else {
    const header = sh.getRange(MC.HEADER_ROW, MC.COL_ACTION, 1, 3).getValues()[0];
    if (header[0] !== 'Action') problems.push('Header not found at A1 — run setupMobileControls()');
    const n = mcActions_().length;
    const checks = sh.getRange(MC.FIRST_ACTION_ROW, MC.COL_RUN, n, 1).getValues();
    if (checks.some(r => typeof r[0] !== 'boolean')) problems.push('Run column is missing checkboxes');
  }

  const hasTrigger = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === 'handleMobileControlsEdit');
  if (!hasTrigger) problems.push('Installable onEdit trigger missing — run setupMobileControls()');

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('BOOKING_WEBAPP_URL')) problems.push('BOOKING_WEBAPP_URL property not set (booking rows will not work)');
  if (!props.getProperty('ADMIN_KEY')) problems.push('ADMIN_KEY property not set (booking rows will not work)');

  if (problems.length) problems.forEach(p => console.log('PROBLEM: ' + p));
  else console.log('Mobile Controls: everything configured correctly.');
  return problems;
}


/**
 * RECOVERY: clears a stuck "Running…" status (e.g. after an execution
 * timeout) and unticks any stuck checkbox. Safe any time — it only resets
 * the display cells; it never launches or kills a run.
 */
function clearStaleControls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(MC.TAB);
  if (!sh) { Logger.log('No ' + MC.TAB + ' tab.'); return; }
  const n = mcActions_().length;
  const status = sh.getRange(MC.FIRST_ACTION_ROW, MC.COL_STATUS, n, 1);
  const vals = status.getValues();
  let cleared = 0;
  vals.forEach((r, i) => {
    if (/^Running/.test(String(r[0] || ''))) {
      sh.getRange(MC.FIRST_ACTION_ROW + i, MC.COL_STATUS)
        .setValue('Reset (was stuck) ' +
          Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH:mm'));
      cleared++;
    }
  });
  sh.getRange(MC.FIRST_ACTION_ROW, MC.COL_RUN, n, 1).setValue(false);
  Logger.log('Stuck statuses cleared: ' + cleared + '. All checkboxes reset.');
}
