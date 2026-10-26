/******************************************************
 * ROOTS & ROADS — viatorControls.gs
 * Bind to: Roots_Roads_Control_v1 (same project as mobileControls.gs).
 *
 * WHAT THIS DOES
 *   Writes an editable table in the "Control" tab where management sets, PER
 *   VIATOR PRODUCT, how many hours before start an EMPTY departure should be
 *   auto-closed ("Sold out") on Viator — and whether that product is managed at
 *   all. The BookingSheet project's ?closecheck endpoint reads this table; the
 *   external GitHub Action bot applies each product's hours to its departures.
 *
 * SETUP (once, from a computer)
 *   Run setupViatorControls() from the Apps Script editor. Safe to re-run — it
 *   rewrites only its own block (columns E..H) and leaves the Mobile Controls
 *   block (A..C) and System Health untouched.
 *
 * LAYOUT  (Control!E1:H)
 *   E: Product name | F: Product code | G: Close hrs before (empty) | H: Enabled
 ******************************************************/

const VC = {
  TAB: 'Control',
  FIRST_COL: 5,     // E
  WIDTH: 4,         // E..H
  HEADER_ROW: 1,
  FIRST_DATA_ROW: 2
};

// The default product rows. Edit hours/enabled later directly in the sheet; add products by
// copying the format. The product code MUST match Viator (e.g. 5631527P3).
function vcSeedRows_() {
  return [
    ['Barcelona Walking Tour: Sagrada Familia, Gaudi and Gothic Quarter', '5631527P3', 10, true],
    ['Private Complete Barcelona Walking Tour with Local Guide', '5631527P5', 24, true]
  ];
}

// One-time setup (run from the editor): write the header, seed rows and checkboxes for the
// Viator block. Safe to re-run — it only touches columns E..H and reseeds the defaults.
function setupViatorControls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(VC.TAB);
  if (!sh) sh = ss.insertSheet(VC.TAB);

  const rows = vcSeedRows_();

  // Clear just this block (E..H) before reseeding; a re-run is explicit setup.
  const clearRows = Math.max(sh.getLastRow(), VC.FIRST_DATA_ROW + rows.length + 5);
  sh.getRange(VC.HEADER_ROW, VC.FIRST_COL, clearRows, VC.WIDTH)
    .clearContent().clearDataValidations();

  // Header row (E..H).
  sh.getRange(VC.HEADER_ROW, VC.FIRST_COL, 1, VC.WIDTH)
    .setValues([['Viator product', 'Product code', 'Close hrs before (empty)', 'Enabled']])
    .setFontWeight('bold').setBackground('#0e7c66').setFontColor('#ffffff');

  // Data rows, then make Enabled (H) checkboxes and Hours (G) plain integers.
  sh.getRange(VC.FIRST_DATA_ROW, VC.FIRST_COL, rows.length, VC.WIDTH).setValues(rows);
  sh.getRange(VC.FIRST_DATA_ROW, VC.FIRST_COL + 3, rows.length, 1).insertCheckboxes();
  sh.getRange(VC.FIRST_DATA_ROW, VC.FIRST_COL + 2, rows.length, 1).setNumberFormat('0');

  sh.setColumnWidth(VC.FIRST_COL, 320);       // E name
  sh.setColumnWidth(VC.FIRST_COL + 1, 120);   // F code
  sh.setColumnWidth(VC.FIRST_COL + 2, 180);   // G hours
  sh.setColumnWidth(VC.FIRST_COL + 3, 90);    // H enabled

  ss.toast('Viator close-hours control ready at ' + VC.TAB + '!E1');
}
