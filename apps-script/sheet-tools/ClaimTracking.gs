/**
 * NUSA — Claim tracking columns  (script bound to "NUSA Expense Claims — Log")
 * ------------------------------------------------------------------
 * Adds three dropdown columns at the END of the "Expense Ledger" tab, for the paper side
 * of each claim that the claim system can't see:
 *   Submitted                — Received by NUSA / Submitted to Accounts Payable / Sent back for changes
 *   Cheque Received from AP  — Yes / No
 *   Cheque Picked Up         — Yes / No
 *
 * Nothing else writes to these columns: the claim web app (Code.gs V3/V4/V5) only writes
 * columns A–T, and the budget tracker only reads A–T, so your choices are never overwritten.
 *
 * A claim with several receipts has one ledger row per receipt. Pick a value on ANY row of a
 * claim and the other rows of the same claim (same Claim ID) get the same value automatically.
 * Manual rows (no Claim ID) are left as you set them.
 *
 * SETUP (once): in the claim log, Extensions > Apps Script. Add a new script file
 * (+ > Script editor) named "Claim Tracking", paste this file, Save.
 * Pick addClaimTrackingColumns in the function list and press Run.
 * Safe to run again: it never adds a column twice and never changes what you've picked.
 */

const TRACK_TAB = 'Expense Ledger';
const TRACK_ROWS = 3000; // rows 2–3001, same range the budget tracker reads
const TRACK_COLUMNS = [
  { head: 'Submitted', values: ['Received by NUSA', 'Submitted to Accounts Payable', 'Sent back for changes'], width: 210 },
  { head: 'Cheque Received from AP', values: ['Yes', 'No'], width: 170 },
  { head: 'Cheque Picked Up', values: ['Yes', 'No'], width: 130 },
];

function trackHeaders_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
}

// RUN ONCE: adds the three columns after the last column (if they aren't there yet) and their dropdowns.
function addClaimTrackingColumns() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TRACK_TAB);
  if (!sh) throw new Error('No "' + TRACK_TAB + '" tab yet. Run setupExpenseLedger in the claim web app\'s Apps Script first.');
  if (sh.getMaxRows() < TRACK_ROWS + 1) sh.insertRowsAfter(sh.getMaxRows(), TRACK_ROWS + 1 - sh.getMaxRows());
  const head = trackHeaders_(sh);
  let last = 0;
  head.forEach(function (h, i) { if (h !== '') last = i + 1; });
  const added = [];
  TRACK_COLUMNS.forEach(function (t) {
    let c = head.indexOf(t.head) + 1;
    if (!c) {
      c = ++last;
      if (c > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
      sh.getRange(1, c).setValue(t.head).setFontWeight('bold');
      sh.setColumnWidth(c, t.width);
      added.push(t.head);
    }
    sh.getRange(2, c, TRACK_ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(t.values, true).setAllowInvalid(false).build());
  });
  console.log(added.length ? 'Added columns: ' + added.join(', ') + '.' : 'Columns were already there; dropdowns refreshed.');
}

// Simple trigger: when you pick a value in one of the three columns, copy it to every row of the same claim.
function onEdit(e) {
  const r = e && e.range;
  if (!r || r.getNumColumns() !== 1 || r.getRow() < 2) return;
  const sh = r.getSheet();
  if (sh.getName() !== TRACK_TAB) return;
  const head = trackHeaders_(sh);
  const col = r.getColumn();
  if (!TRACK_COLUMNS.some(function (t) { return t.head === head[col - 1]; })) return;
  const idCol = head.indexOf('Claim ID') + 1;
  const n = sh.getLastRow() - 1;
  if (!idCol || n < 1) return;
  const ids = sh.getRange(2, idCol, n, 1).getValues().map(function (x) { return String(x[0]).trim(); });
  const vals = sh.getRange(2, col, n, 1).getValues();
  const pick = {};
  for (let i = r.getRow() - 2; i < Math.min(r.getLastRow() - 1, n); i++) { if (ids[i]) pick[ids[i]] = vals[i][0]; }
  let changed = false;
  ids.forEach(function (id, i) {
    if (id && pick.hasOwnProperty(id) && vals[i][0] !== pick[id]) { vals[i][0] = pick[id]; changed = true; }
  });
  if (changed) sh.getRange(2, col, n, 1).setValues(vals);
}
