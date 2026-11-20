/**
 * NUSA budget tracker inside the claim log  (same Apps Script project as "Add manual expense")
 * ------------------------------------------------------------------
 * Builds the budget tracker as tabs of "NUSA Expense Claims — Log", reading the Expense Ledger live:
 *   Dashboard       totals, departments, GL lines ranked closest-to-budget first (department picker)
 *   Charts          graphs of spending: by department, % used vs % of year, by month, top budget lines
 *   GL Budget      every GL line: budget, spent, pending, remaining, % used, status, Fall/Spring halves
 *   Line Items      the AGM line items (change a budget here)
 *   Tracker Ledger  the Expense Ledger with budget line, department and semester worked out (read-only)
 *   GL Map          which budget line each GL code is charged to
 *   Dept Summary    tie-out to the AGM totals
 *   Tracker Notes   how it works
 * Claims, Expense Ledger and GL List are never touched.
 * Revenue Ledger: money NUSA receives (student fees, events revenue, transfers in, investment gains).
 *   Created the first time the tracker is (re)built, or by running setupRevenueLedger; a rebuild never touches it.
 *   GL Budget shows each revenue line's "received" (Status Received) next to its budget; the Dashboard has a
 *   "Revenue received" tile. Refunds of something NUSA paid for (e.g. a software refund) are NOT revenue: add them
 *   to Expense Ledger as a negative amount on the same GL code, so that expense line goes down.
 * Extra GL Map keys (TRACKER_EXTRA below), from VP Finance decisions: e.g. payments the university booked on
 *   NUSA Funding Clubs that are really the loan repayment to the university (Capital Projects).
 *
 * SETUP: add this file and "Tracker Data" to the project, Save, then choose buildBudgetTracker next to
 * Run and run it once. After a new AGM budget: paste a new "Tracker Data" file and run
 * rebuildBudgetTracker. It replaces the tracker tabs but keeps your edits: budgets typed on Line Items and
 * budget lines changed on GL Map are put back, unless the AGM budget changed that same line (then the new AGM
 * value is used and the log lists it). A copy of the old Line Items is kept on "Line Items (before rebuild)".
 * Expenses live on Expense Ledger, which a rebuild never touches.
 */
const TR = { dash: 'Dashboard', charts: 'Charts', gb:'GL Budget', li: 'Line Items', led: 'Tracker Ledger', map: 'GL Map', sum: 'Dept Summary', notes: 'Tracker Notes' };
const TR_ROWS = 3000;                        // Expense Ledger rows read (2..3001)
const LED_LAST = 4 + TR_ROWS;                // Tracker Ledger rows 5..3004
const MONEY = '$#,##0.00;($#,##0.00);"-"';
const PCT = '0.0%;-0.0%;"-"';
const C_NAVY = '#1F3864', C_CALC = '#595959', C_SUB = '#595959', C_TILE = '#EEF2F8', C_INPUT = '#FFF2CC';
const STATUS_FILL = [['Over budget', '#F8CBAD'], ['Unbudgeted spend', '#F8CBAD'], ['Almost out', '#FCE4D6'], ['Watch', '#FFF2CC'], ['On track', '#E2EFDA']];

// Revenue Ledger tab (not a tracker tab: a rebuild never deletes it). Column L is a helper ("g" + GL code digits)
// so SUMIFS compares text: a 28-digit code read as a number would lose digits.
const RL = 'Revenue Ledger', RL_ROWS = 1000;
const RL_HEADS = ['Entry ID', 'Date', 'GL Code', 'GL Name', 'Group', 'Description', 'From', 'Amount', 'Status', 'NU Ref', 'Notes', 'Match Key'];

// Extra GL Map keys (set by VP Finance): a GL code the university uses, sent to an existing budget line when the ledger row's
// "Charge to Department" is dept. Key = "<code digits>|<dept>".
const TRACKER_EXTRA = [
  { code: '410  7211156025100003907010000', dept: 'Capital Projects', line: 'Loan repayment to the university',
    why: 'Payments the university booked on NUSA Funding Clubs for the loan repayment. On Expense Ledger those rows have Charge to Department = Capital Projects.' },
];

// TRACKER_DATA plus the TRACKER_EXTRA map keys.
function trackerData_() {
  const D = TRACKER_DATA;
  const map = D.map.slice();
  TRACKER_EXTRA.forEach(function (x) {
    const g = D.gl.filter(function (r) { return r[1] === x.dept && r[3] === 'Expenditure' && r[5] === x.line; })[0];
    if (!g) throw new Error('TRACKER_EXTRA: no budget line "' + x.line + '" in ' + x.dept + ' (check the AGM budget / Tracker Data).');
    map.push([x.code.replace(/\D/g, '') + '|' + x.dept, g[10], x.code, x.dept, x.why]);
  });
  return { built: D.built, depts: D.depts, gl: D.gl, li: D.li, map: map, agm: D.agm };
}

function buildBudgetTracker() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName('Expense Ledger')) throw new Error('No "Expense Ledger" tab yet. Run setupExpenseLedger in the claim web app first.');
  const existing = Object.keys(TR).map(function (k) { return TR[k]; }).filter(function (n) { return ss.getSheetByName(n); });
  if (existing.length) throw new Error('The tracker already exists (' + existing.join(', ') + '). To replace it, run rebuildBudgetTracker (it keeps your Line Items and GL Map edits).');
  build_(ss);
}

function rebuildBudgetTracker() {
  const ss = SpreadsheetApp.getActive();
  const ledger = ss.getSheetByName('Expense Ledger');
  if (!ledger) throw new Error('No "Expense Ledger" tab yet. Run setupExpenseLedger in the claim web app first.');
  const ledgerRows = ledger.getLastRow();
  const saved = saveTrackerEdits_(ss);
  const keep = ss.getSheetByName('Claims') || ss.getSheets()[0];
  ss.setActiveSheet(keep);
  Object.keys(TR).forEach(function (k) { const sh = ss.getSheetByName(TR[k]); if (sh) ss.deleteSheet(sh); });
  build_(ss);
  const res = restoreTrackerEdits_(ss, saved);
  const lines = ['Expense Ledger untouched: ' + Math.max(0, ledgerRows - 1) + ' expense rows (manual + claims), still counted live.'];
  if (!saved.checked) lines.push('Your old Line Items could not be checked for edits (it was built before this feature). Compare with the tab "' + LI_BACKUP + '" if you had typed budgets there.');
  lines.push('Your edits put back: ' + res.kept.length + (res.kept.length ? ' - ' + res.kept.join('; ') : ''));
  if (res.notKept.length) lines.push('Edits NOT put back (the AGM budget changed that line, or it is gone): ' + res.notKept.join('; '));
  if (saved.backup) {
    ss.setActiveSheet(saved.backup);
    ss.moveActiveSheet(ss.getNumSheets()); // last tab, out of the way
    ss.setActiveSheet(ss.getSheetByName(TR.dash));
    saved.backup.getRange('A1:A3').setValues([['Line Items before the rebuild of ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a')],
      ['A copy kept so nothing is lost. Nothing reads this tab; delete it whenever you like (the next rebuild replaces it).'], [lines.slice(1).join('   ')]]);
  }
  lines.forEach(function (l) { console.log(l); });
}

/* ---------------------------------------------------------------- keep edits across a rebuild */
// Line Items column P and GL Map column J hold the values the tracker was built with, so a rebuild can tell
// which budgets / budget lines were changed by hand.
const LI_ORIG = 'AGM Budget (as built)', MAP_ORIG = 'Budget Line (as built)', LI_BACKUP = 'Line Items (before rebuild)';
const cents_ = function (v) { return Math.round((Number(v) || 0) * 100); };
// A Line Items row's identity: department | section | GL code | line item | semester (+ #n for repeated names).
function liKey_(r, seen) {
  const k = [r[1], r[3], String(r[6]).replace(/\D/g, ''), String(r[8]).trim(), r[14]].join('|');
  seen[k] = (seen[k] || 0) + 1;
  return k + '#' + seen[k];
}
function liName_(r) { return r[1] + ' > ' + String(r[8]).trim() + (r[14] && r[14] !== 'Full year' ? ' (' + r[14] + ')' : ''); }

function saveTrackerEdits_(ss) {
  const out = { li: {}, map: {}, checked: false, backup: null };
  const li = ss.getSheetByName(TR.li);
  if (li && li.getLastRow() >= 5) {
    const w = Math.min(16, li.getLastColumn());
    const has = w >= 16 && li.getRange(4, 16).getValue() === LI_ORIG;
    out.checked = has;
    const seen = {};
    li.getRange(5, 1, li.getLastRow() - 4, w).getValues().forEach(function (r) {
      if (typeof r[0] !== 'number') return; // header/total rows
      const k = liKey_(r, seen);
      if (has && cents_(r[10]) !== cents_(r[15])) out.li[k] = { budget: Number(r[10]) || 0, orig: Number(r[15]) || 0, name: liName_(r) };
    });
    const old = ss.getSheetByName(LI_BACKUP);
    if (old) ss.deleteSheet(old);
    out.backup = li.copyTo(ss).setName(LI_BACKUP);
  }
  const mp = ss.getSheetByName(TR.map);
  if (mp && mp.getLastRow() >= 5 && mp.getLastColumn() >= 10 && mp.getRange(4, 10).getValue() === MAP_ORIG) {
    mp.getRange(5, 1, mp.getLastRow() - 4, 10).getValues().forEach(function (r) {
      if (r[0] && String(r[1]) !== String(r[9])) out.map[String(r[0])] = { line: String(r[1]), orig: String(r[9]) };
    });
  }
  return out;
}

function restoreTrackerEdits_(ss, saved) {
  const D = trackerData_(), kept = [], notKept = [];
  const money = function (v) { return '$' + (Number(v) || 0).toFixed(2); };
  const li = ss.getSheetByName(TR.li);
  const rows = li.getRange(5, 1, D.li.length, 16).getValues(), at = {}, seen = {};
  rows.forEach(function (r, i) { at[liKey_(r, seen)] = { row: i + 5, orig: r[15] }; });
  Object.keys(saved.li).forEach(function (k) {
    const e = saved.li[k], n = at[k];
    if (!n) { notKept.push(e.name + ' (' + money(e.budget) + '; line no longer in the budget)'); return; }
    if (cents_(n.orig) !== cents_(e.orig)) {
      notKept.push(e.name + ' (you had ' + money(e.budget) + '; the AGM changed it from ' + money(e.orig) + ' to ' + money(n.orig) + ', so ' + money(n.orig) + ' is used)');
      return;
    }
    li.getRange(n.row, 11).setValue(e.budget).setBackground(C_INPUT);
    kept.push(e.name + ' ' + money(e.budget));
  });
  const mp = ss.getSheetByName(TR.map);
  const labels = {};
  D.gl.forEach(function (g) { if (g[3] === 'Expenditure') labels[g[10]] = true; });
  const mrow = {};
  D.map.forEach(function (m, i) { mrow[m[0]] = { row: i + 5, orig: m[1] }; });
  Object.keys(saved.map).forEach(function (k) {
    const e = saved.map[k], n = mrow[k];
    const nm = 'GL Map ' + k.replace(/\|$/, '') + ' -> ' + e.line;
    if (!n || n.orig !== e.orig || !labels[e.line]) { notKept.push(nm + ' (' + (!n ? 'code no longer on the map' : !labels[e.line] ? 'that budget line no longer exists' : 'the budget changed this code\'s default line') + ')'); return; }
    mp.getRange(n.row, 2).setValue(e.line).setBackground(C_INPUT);
    kept.push(nm);
  });
  return { kept: kept, notKept: notKept };
}

// Puts the Dashboard's department picker (C23) back to "All departments" with its dropdown and colour,
// e.g. after something was pasted over it. Touches nothing else.
function fixDashboardPicker() {
  const s = SpreadsheetApp.getActive().getSheetByName(TR.dash);
  if (!s) throw new Error('No "' + TR.dash + '" tab. Run buildBudgetTracker first.');
  s.getRange('B23').setValue('Department:').setFontWeight('bold');
  s.getRange('C23').clearDataValidations().setValue('All departments').setFontWeight('bold').setBackground(C_INPUT)
    .setBorder(true, true, true, true, null, null)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['All departments'].concat(TRACKER_DATA.depts), true).setAllowInvalid(false).build());
  console.log('Dashboard department picker (C23) restored.');
}

/* ---------------------------------------------------------------- Revenue Ledger */
// Run once to add the Revenue Ledger tab before the next rebuild (the build also adds it if it's missing).
function setupRevenueLedger() {
  const ss = SpreadsheetApp.getActive();
  const had = !!ss.getSheetByName(RL);
  ensureRevenueLedger_(ss, trackerData_());
  console.log((had ? 'The "' + RL + '" tab already existed: headers, formula columns and dropdowns put back (your rows unchanged).' : 'Added the "' + RL + '" tab.') +
    ' Run rebuildBudgetTracker so GL Budget and the Dashboard read it.');
}

// Revenue GL codes for the dropdown: Type = Revenue on GL List (or, on an older GL List without Type, account
// numbers starting 4, 6 or 8), plus every revenue code in the budget. Other codes can still be typed.
function revenueCodes_(ss, D) {
  const out = [], seen = {};
  const add = function (c) { c = String(c || '').trim(); if (c && !seen[c]) { seen[c] = 1; out.push(c); } };
  const gl = ss.getSheetByName('GL List');
  if (gl && gl.getLastRow() > 1) {
    const v = gl.getRange(1, 1, gl.getLastRow(), gl.getLastColumn()).getValues();
    const t = v[0].map(String).indexOf('Type');
    v.slice(1).forEach(function (r) {
      const acct = String(r[0]).replace(/\D/g, '').slice(8, 13);
      if (t >= 0 ? r[t] === 'Revenue' : /^[468]/.test(acct)) add(r[0]);
    });
  }
  D.gl.forEach(function (g) {
    const d = String(g[6]).replace(/\D/g, '');
    if (g[3] === 'Revenue' && d.length === 28) add(d.slice(0, 3) + '-' + d.slice(3, 8) + '-' + d.slice(8));
  });
  return out;
}

// Adds the tab if it's missing, then (every build) puts back its headers, formula columns (D, E, L), formats and
// dropdowns. The rows typed in the other columns are never changed.
function ensureRevenueLedger_(ss, D) {
  let s = ss.getSheetByName(RL);
  if (!s) {
    const names = ss.getSheets().map(function (x) { return x.getName(); });
    const at = names.indexOf('Expense Ledger');
    s = ss.insertSheet(RL, at >= 0 ? at + 1 : names.length);
  }
  rowsNeeded_(s, RL_ROWS + 1, RL_HEADS.length);
  header_(s, 1, 1, RL_HEADS, [80, 90, 230, 230, 150, 300, 160, 100, 90, 110, 300, 60], ['GL Name', 'Group', 'Match Key']);
  const col = function (c) { return c + '2:' + c + (RL_ROWS + 1); };
  s.getRange('D2:E' + (RL_ROWS + 1)).clearContent(); // formula columns: anything typed there would block the formulas
  s.getRange(col('L')).clearContent();
  s.getRange('D2').setFormula('=ARRAYFORMULA(IF(' + col('C') + '="","",IFERROR(VLOOKUP(' + col('C') + ",'GL List'!$A:$B,2,FALSE),\"(not on GL List)\")))");
  s.getRange('E2').setFormula('=ARRAYFORMULA(IF(' + col('C') + '="","",IFERROR(VLOOKUP(' + col('C') + ",'GL List'!$A:$C,3,FALSE),\"\")))");
  s.getRange('L2').setFormula('=ARRAYFORMULA(IF(' + col('C') + '="","","g"&REGEXREPLACE(TO_TEXT(' + col('C') + '),"\\D","")))');
  s.getRange(col('B')).setNumberFormat('yyyy-mm-dd');
  s.getRange(col('C')).setNumberFormat('@')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(revenueCodes_(ss, D), true).setAllowInvalid(true).build());
  s.getRange(col('H')).setNumberFormat(MONEY);
  s.getRange(col('I')).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Received', 'Expected', 'Void'], true).setAllowInvalid(false).build());
  s.hideColumns(12, 1);
  s.setFrozenRows(1);
  return s;
}

/* ---------------------------------------------------------------- build */
function build_(ss) {
  const D = trackerData_();
  ensureRevenueLedger_(ss, D); // GL Budget reads it, so it must exist first
  // Create every tab first so formulas that point at each other resolve.
  const n = ss.getSheets().length;
  const sh = {
    dash: ss.insertSheet(TR.dash, 0), charts: ss.insertSheet(TR.charts, 1),
    gb: ss.insertSheet(TR.gb, n + 2), li: ss.insertSheet(TR.li, n + 3), led: ss.insertSheet(TR.led, n + 4),
    map: ss.insertSheet(TR.map, n + 5), sum: ss.insertSheet(TR.sum, n + 6), notes: ss.insertSheet(TR.notes, n + 7),
  };
  const byKey = {}; D.gl.forEach(function (g) { byKey[g[0]] = g; });
  const liLast = 4 + D.li.length, gbLast = 4 + D.gl.length;

  buildLineItems_(sh.li, D, byKey, liLast);
  buildMap_(sh.map, D);
  buildLedger_(sh.led);
  buildGlBudget_(sh.gb, D, liLast, gbLast);
  buildDashboard_(sh.dash, D, gbLast);
  buildCharts_(sh.charts, D, gbLast);
  buildSummary_(sh.sum, D, gbLast);
  buildNotes_(sh.notes);

  [sh.charts, sh.gb, sh.led, sh.map, sh.sum].forEach(function (s) { s.protect().setDescription('Tracker formulas').setWarningOnly(true); });
  const p = sh.dash.protect().setDescription('Tracker dashboard').setWarningOnly(true);
  p.setUnprotectedRanges([sh.dash.getRange('C23')]);
  SpreadsheetApp.flush();
  ss.setActiveSheet(sh.dash);
  console.log('Budget tracker built: ' + D.gl.length + ' GL lines, ' + D.li.length + ' line items (data from ' + D.built + ').');
}

/* ---------------------------------------------------------------- helpers */
function title_(s, title, sub, sub2) {
  s.getRange('A1').setValue(title).setFontSize(16).setFontWeight('bold').setFontColor(C_NAVY);
  if (sub) s.getRange('A2').setValue(sub).setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);
  if (sub2) s.getRange('A3').setValue(sub2).setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);
}
function header_(s, row, col, heads, widths, calc) {
  const r = s.getRange(row, col, 1, heads.length);
  r.setValues([heads]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(C_NAVY).setWrap(true).setVerticalAlignment('middle');
  heads.forEach(function (h, i) {
    if (calc && calc.indexOf(h) >= 0) s.getRange(row, col + i).setBackground(C_CALC);
    if (widths) s.setColumnWidth(col + i, widths[i]);
  });
  s.setRowHeight(row, 36);
}
function rowsNeeded_(s, rows, cols) {
  if (s.getMaxRows() < rows) s.insertRowsAfter(s.getMaxRows(), rows - s.getMaxRows());
  if (cols && s.getMaxColumns() < cols) s.insertColumnsAfter(s.getMaxColumns(), cols - s.getMaxColumns());
}
function deptArray_(D) { return '{' + D.depts.map(function (d) { return '"' + d + '"'; }).join(';') + '}'; }

/* ---------------------------------------------------------------- Line Items */
function buildLineItems_(s, D, byKey, last) {
  rowsNeeded_(s, last + 5, 16);
  title_(s, '2026-27 Budget: Line Items by GL Code', 'Every budget line from the AGM budget, tied to the GL code it rolls up to. To change a budget, edit the 2026-27 Budget column (blue).',
    'Your edits are kept when the tracker is rebuilt (column P = the AGM amount it was built with; if the AGM changes that line, the new AGM amount wins).');
  const heads = ['Line ID', 'Department (Tab)', 'Dept #', 'Section', 'Category', 'GL Description', 'GL Code', 'Sub-group', 'Line Item',
    '2025-26 Actual', '2026-27 Budget', 'Counted in AGM Total?', 'Source Cell (AGM file)', 'GL Key', 'Semester', LI_ORIG];
  header_(s, 4, 1, heads, [55, 150, 60, 85, 170, 230, 230, 100, 260, 100, 100, 90, 160, 55, 80, 110], [LI_ORIG]);
  const vals = D.li.map(function (l) {
    const p = byKey[l[1]];
    return [l[0], p[1], p[2], p[3], p[4], p[5], p[6], l[2], l[3], l[4], l[5], l[6] ? 'Yes' : 'No', l[7], l[1], l[8], l[5]];
  });
  s.getRange(5, 7, vals.length, 1).setNumberFormat('@');
  s.getRange(5, 1, vals.length, heads.length).setValues(vals);
  s.getRange(5, 10, vals.length, 2).setNumberFormat(MONEY);
  s.getRange(5, 16, vals.length, 1).setNumberFormat(MONEY).setFontColor(C_CALC);
  s.getRange(5, 11, vals.length, 1).setFontColor('#0000FF');
  s.getRange(last + 2, 9).setValue('Total of visible rows:').setFontWeight('bold');
  s.getRange(last + 2, 10, 1, 2).setFormulas([['=SUBTOTAL(109,J5:J' + last + ')', '=SUBTOTAL(109,K5:K' + last + ')']]).setNumberFormat(MONEY).setFontWeight('bold');
  s.getRange(4, 1, vals.length + 1, heads.length).createFilter();
  s.setFrozenRows(4); s.setFrozenColumns(2);
}

/* ---------------------------------------------------------------- GL Map */
function buildMap_(s, D) {
  const lines = []; const seen = {};
  D.gl.forEach(function (g) { if (g[3] === 'Expenditure' && !seen[g[10]]) { seen[g[10]] = 1; lines.push([g[10], g[1]]); } });
  rowsNeeded_(s, Math.max(D.map.length, lines.length) + 10, 10);
  title_(s, 'GL Map: which budget line each GL code is charged to',
    'Key = GL code digits | department. The row with an empty department is the default when the ledger row has no "Charge to Department".',
    'To send a GL code to a different budget line, change its Budget Line (copy the label from column G). Columns G:H list every budget line and its department. Your changes are kept when the tracker is rebuilt (column J = the line it was built with).');
  header_(s, 4, 1, ['Key', 'Budget Line', 'GL Code', 'Department', 'Why'], [260, 380, 230, 160, 360]);
  s.getRange(5, 1, D.map.length, 5).setNumberFormat('@').setValues(D.map);
  s.getRange(5, 2, D.map.length, 1).setFontColor('#0000FF');
  header_(s, 4, 7, ['Budget Line', 'Department'], [380, 160]);
  s.getRange(5, 7, lines.length, 2).setValues(lines);
  header_(s, 4, 10, [MAP_ORIG], [380], [MAP_ORIG]);
  s.getRange(5, 10, D.map.length, 1).setNumberFormat('@').setFontColor(C_CALC).setValues(D.map.map(function (m) { return [m[1]]; }));
  s.setFrozenRows(4);
}

/* ---------------------------------------------------------------- Tracker Ledger (live formulas over Expense Ledger) */
function buildLedger_(s) {
  rowsNeeded_(s, LED_LAST + 5, 20);
  title_(s, 'Tracker Ledger',
    'Read-only: the Expense Ledger tab with each row\'s budget line, department and semester worked out. Edit expenses on Expense Ledger (or use NUSA > Add manual expense).',
    'Counted: Approved and Posted = Spent, Pending = Pending, Void = ignored (a blank Status counts as Approved). Matched? = No means the row isn\'t counted: its GL code isn\'t on the GL Map or its Status isn\'t Pending/Approved/Posted/Void.');
  const heads = ['Entry ID', 'Source', 'Date', 'Department', 'Budget Line', 'GL Code', 'GL Name', 'Description', 'Payee',
    'Amount (incl. tax)', 'GST', 'HST', 'Net Expense', 'Status', 'Semester', 'Month', 'Claim ID', 'Claim Status', 'Notes', 'Matched?'];
  header_(s, 4, 1, heads, [85, 65, 85, 160, 330, 230, 190, 260, 150, 110, 70, 70, 95, 80, 75, 70, 85, 190, 200, 75]);
  const E = function (c) { return "'Expense Ledger'!" + c + '2:' + c + (TR_ROWS + 1); };
  const own = function (c) { return c + '5:' + c + LED_LAST; };
  const blank = 'LEN(' + E('A') + '&' + E('D') + '&' + E('J') + ')=0';
  const wrap = function (x) { return '=ARRAYFORMULA(IF(' + blank + ',"",' + x + '))'; };
  const digits = 'REGEXREPLACE(TO_TEXT(' + E('D') + '),"\\D","")';
  const f = [
    wrap(E('A')),                                                                      // A Entry ID
    wrap(E('B')),                                                                      // B Source
    wrap('IF(' + E('C') + '="","",' + E('C') + ')'),                                   // C Date
    wrap('IF(' + own('E') + '<>"",IFERROR(VLOOKUP(' + own('E') + ",'GL Map'!$G$5:$H,2,FALSE),\"\"),IF(TRIM(" + E('G') + ')<>"",TRIM(' + E('G') + '),' + E('F') + '))'), // D Department
    wrap('IF(' + digits + '="","",IFERROR(VLOOKUP(' + digits + '&"|"&TRIM(' + E('G') + "),'GL Map'!$A$5:$B,2,FALSE),IFERROR(VLOOKUP(" + digits + "&\"|\",'GL Map'!$A$5:$B,2,FALSE),\"\")))"), // E Budget Line
    wrap(E('D')), wrap(E('E')), wrap(E('H')), wrap(E('I')),                          // F GL Code, G GL Name, H Description, I Payee
    wrap('IFERROR(VALUE(' + E('J') + '),0)'), wrap('IFERROR(VALUE(' + E('K') + '),0)'), // J Amount, K GST
    wrap('IFERROR(VALUE(' + E('L') + '),0)'), wrap('IFERROR(VALUE(' + E('M') + '),0)'), // L HST, M Net Expense
    wrap('IF(TRIM(' + E('N') + ')="","Approved",PROPER(TRIM(' + E('N') + ')))'),        // N Status
    wrap('IF(ISNUMBER(' + E('C') + '),IF(MONTH(' + E('C') + ')<=4,"Spring","Fall"),IF(MONTH(TODAY())<=4,"Spring","Fall"))'), // O Semester
    wrap('IF(ISNUMBER(' + E('C') + '),TEXT(' + E('C') + ',"yyyy-mm"),"")'),            // P Month
    wrap(E('P')), wrap(E('R')), wrap(E('T')),                                          // Q Claim ID, R Claim Status, S Notes
    wrap('IF((' + own('E') + '="")+ISNA(MATCH(' + own('N') + ',{"Pending";"Approved";"Posted";"Void"},0)),"No","Yes")'), // T Matched?
  ];
  s.getRange(5, 1, 1, f.length).setFormulas([f]);
  s.getRange(5, 3, TR_ROWS, 1).setNumberFormat('yyyy-mm-dd');
  s.getRange(5, 10, TR_ROWS, 4).setNumberFormat(MONEY);
  s.setFrozenRows(4); s.setFrozenColumns(3);
}

/* ---------------------------------------------------------------- GL Budget */
function buildGlBudget_(s, D, liLast, last) {
  rowsNeeded_(s, last + 5, 36);
  title_(s, '2026-27 Budget by GL Code',
    'One row per GL line on each department tab. Budget = sum of the line items. Spent/Pending come from the Expense Ledger (live). Use the filter buttons to filter by Department, Category, GL Code or Has Budget?.',
    'Revenue lines: Spent = revenue received and Pending = revenue expected, from the Revenue Ledger tab (by GL code).');
  const heads = ['GL Key', 'Department (Tab)', 'Dept #', 'Section', 'Category', 'GL Description', 'GL Code',
    '2026-27 Budget', 'Spent (Approved + Posted)', 'Pending', 'Remaining', '% Used (this period)', 'Status (this period)', 'Has Budget?',
    '2025-26 Actual', '# Line Items', 'GL Dept #', 'Account #', 'Same GL Code: Total Across All Tabs',
    'Tracker Label (used in Transactions)', 'Source Cell (AGM file)',
    'Split by Semester?', 'Fall Budget', 'Fall Spent', 'Spring Budget', 'Spring Spent',
    'Tracking Period', 'Period Budget', 'Period Spent', 'Period Pending', 'Period Remaining',
    'Dept Order', 'Rank Key', 'Show on Dashboard', '% Used (number)', 'Revenue Match Key'];
  const calc = ['2026-27 Budget', 'Spent (Approved + Posted)', 'Pending', 'Remaining', '% Used (this period)', 'Status (this period)', 'Has Budget?',
    '# Line Items', 'Same GL Code: Total Across All Tabs', 'Fall Budget', 'Fall Spent', 'Spring Budget', 'Spring Spent', 'Tracking Period',
    'Period Budget', 'Period Spent', 'Period Pending', 'Period Remaining'];
  header_(s, 4, 1, heads, [50, 150, 60, 85, 170, 230, 230, 100, 100, 85, 100, 75, 120, 70, 100, 60, 65, 65, 110, 330, 150,
    70, 90, 90, 90, 90, 80, 95, 95, 95, 95, 60, 60, 60, 60, 60], calc);
  const LI = function (c) { return "'Line Items'!$" + c + '$5:$' + c + '$' + liLast; };
  const TL = function (c) { return "'Tracker Ledger'!$" + c + '$5:$' + c + '$' + LED_LAST; };
  const RV = function (c) { return "'" + RL + "'!$" + c + '$2:$' + c + '$' + (RL_ROWS + 1); };
  const received = function (r, status) { return '=IF($AJ' + r + '="",0,SUMIFS(' + RV('H') + ',' + RV('L') + ',$AJ' + r + ',' + RV('I') + ',"' + status + '"))'; };
  const spent = function (r, extra) {
    return 'SUMIFS(' + TL('J') + ',' + TL('E') + ',$T' + r + ',' + TL('N') + ',"Approved"' + (extra || '') + ')+SUMIFS(' + TL('J') + ',' + TL('E') + ',$T' + r + ',' + TL('N') + ',"Posted"' + (extra || '') + ')';
  };
  const depts = deptArray_(D);
  const vals = D.gl.map(function (g, i) {
    const r = 5 + i;
    const liSum = function (sem) { return '=SUMIFS(' + LI('K') + ',' + LI('N') + ',A' + r + ',' + LI('L') + ',"Yes"' + (sem ? ',' + LI('O') + ',"' + sem + '"' : '') + ')'; };
    const rev = g[3] === 'Revenue';
    return [
      g[0], g[1], g[2], g[3], g[4], g[5], g[6],
      liSum(''),                                                                        // H budget
      rev ? received(r, 'Received') : '=' + spent(r),                                   // I spent (revenue: received)
      rev ? received(r, 'Expected') : '=SUMIFS(' + TL('J') + ',' + TL('E') + ',$T' + r + ',' + TL('N') + ',"Pending")', // J pending (revenue: expected)
      '=H' + r + '-I' + r,                                                              // K remaining
      '=IF(AB' + r + '>0,AC' + r + '/AB' + r + ',"")',                                  // L % used
      rev ? '=IF(AND(AB' + r + '=0,AC' + r + '=0),"-",IF(AB' + r + '=0,"Not in budget",IF(AC' + r + '>=AB' + r + ',"Budget reached","Receiving")))'
        : '=IF(AND(AB' + r + '=0,AC' + r + '=0),"-",IF(AB' + r + '=0,"Unbudgeted spend",IF(AC' + r + '>AB' + r + ',"Over budget",IF(AC' + r + '/AB' + r + '>=0.9,"Almost out",IF(AC' + r + '/AB' + r + '>=0.75,"Watch","On track")))))',
      '=IF(H' + r + '<>0,"Yes","No")',                                                  // N has budget
      g[7],                                                                             // O actual
      '=COUNTIFS(' + LI('N') + ',A' + r + ')',                                          // P # line items
      g[8], g[9],                                                                       // Q GL dept #, R account #
      '=IF(LEFT(G' + r + ',1)="(","",SUMIFS(H$5:H$' + last + ',G$5:G$' + last + ',G' + r + ',D$5:D$' + last + ',D' + r + '))',
      g[10], g[11], g[12] ? 'Yes' : 'No',                                               // T label, U source, V split
      liSum('Fall'), '=' + spent(r, ',' + TL('O') + ',"Fall"'),                         // W, X
      liSum('Spring'), '=' + spent(r, ',' + TL('O') + ',"Spring"'),                     // Y, Z
      '=IF(V' + r + '="Yes",Dashboard!$K$5,"Full year")',                               // AA period (Dashboard's Current semester)
      '=IF(V' + r + '="Yes",IF(AA' + r + '="Fall",W' + r + ',Y' + r + '),H' + r + ')',  // AB period budget
      '=IF(V' + r + '="Yes",IF(AA' + r + '="Fall",X' + r + ',Z' + r + '),I' + r + ')',  // AC period spent
      '=IF(V' + r + '="Yes",SUMIFS(' + TL('J') + ',' + TL('E') + ',$T' + r + ',' + TL('N') + ',"Pending",' + TL('O') + ',AA' + r + '),J' + r + ')', // AD
      '=AB' + r + '-AC' + r,                                                            // AE period remaining
      '=IFERROR(MATCH(B' + r + ',' + depts + ',0),99)',                                 // AF dept order
      '=IF(AB' + r + '>0,AC' + r + '/AB' + r + ',IF(AC' + r + '>0,1000,-1))',           // AG rank key
      '=IF(AND(D' + r + '="Expenditure",OR(H' + r + '<>0,I' + r + '<>0,J' + r + '<>0)),1,0)', // AH show
      '=IF(AB' + r + '>0,AC' + r + '/AB' + r + ',0)',                                   // AI % used (number, for QUERY)
      rev && String(g[6]).replace(/\D/g, '') ? 'g' + String(g[6]).replace(/\D/g, '') : '', // AJ revenue match key
    ];
  });
  // Text columns must stay text (GL codes, labels); set formats first so values aren't re-interpreted.
  s.getRange(5, 7, vals.length, 1).setNumberFormat('@');
  s.getRange(5, 20, vals.length, 1).setNumberFormat('@');
  s.getRange(5, 36, vals.length, 1).setNumberFormat('@');
  s.getRange(5, 1, vals.length, heads.length).setValues(vals);
  [8, 9, 10, 11, 15, 19, 23, 24, 25, 26, 28, 29, 30, 31].forEach(function (c) { s.getRange(5, c, vals.length, 1).setNumberFormat(MONEY); });
  s.getRange(5, 12, vals.length, 1).setNumberFormat(PCT);
  s.getRange(5, 15, vals.length, 1).setFontColor('#0000FF');
  s.getRange(last + 2, 6).setValue('Total of visible rows (updates with filters):').setFontWeight('bold');
  [8, 9, 10, 11, 15].forEach(function (c) {
    const L = columnLetter_(c);
    s.getRange(last + 2, c).setFormula('=SUBTOTAL(109,' + L + '5:' + L + last + ')').setNumberFormat(MONEY).setFontWeight('bold');
  });
  s.getRange(4, 1, vals.length + 1, 31).createFilter();
  s.hideColumns(32, 5);
  s.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$N5="No"').setFontColor('#A6A6A6')
    .setRanges([s.getRange(5, 1, vals.length, 31)]).build()]);
  s.setFrozenRows(4); s.setFrozenColumns(2);
}

function columnLetter_(c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - m) / 26); } return s; }

/* ---------------------------------------------------------------- Dashboard */
function buildDashboard_(s, D, gbLast) {
  rowsNeeded_(s, 360, 13);
  const GB = function (c) { return "'GL Budget'!$" + c + '$5:$' + c + '$' + gbLast; };
  [18, 190, 240, 280, 130, 110, 110, 120, 110, 110, 110, 110, 120].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  s.getRange('B1').setValue('NUSA 2026-27 Budget Tracker').setFontSize(18).setFontWeight('bold').setFontColor(C_NAVY);
  s.getRange('B2').setFormula('="As of "&TEXT(TODAY(),"mmmm d, yyyy")&"   |   Fiscal year May 1, 2026 to April 30, 2027   |   Updates live from the Expense Ledger"')
    .setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);

  // Same split as the AGM budget's P&L: revenue and operating expenses (giving the surplus), with Capital
  // Projects on its own (paid from retained earnings, not from the year's revenue). Spent / Remaining / % used
  // cover all spending, Capital Projects included. Current semester must stay in K5 (GL Budget reads it).
  const SEC = function (col, sec, capital) {
    return '=SUMIFS(' + GB(col) + ',' + GB('D') + ',"' + sec + '",' + GB('B') + ',"' + (capital ? '' : '<>') + 'Capital Projects")';
  };
  const kpis = [
    ['Revenue budget', SEC('H', 'Revenue', false), '$#,##0'],
    ['Operating expense budget', SEC('H', 'Expenditure', false), '$#,##0'],
    ['Budgeted surplus', '=B5-C5', '$#,##0;($#,##0)'],
    ['Capital Projects (from savings)', SEC('H', 'Expenditure', true), '$#,##0'],
    ['Spent', '=SUMIFS(' + GB('I') + ',' + GB('D') + ',"Expenditure")', '$#,##0'],
    ['Pending', '=SUMIFS(' + GB('J') + ',' + GB('D') + ',"Expenditure")', '$#,##0'],
    ['Remaining (all spending)', '=C5+E5-F5', '$#,##0'],
    ['% of budget used', '=IF(C5+E5>0,F5/(C5+E5),"")', PCT],
    ['% of year gone', '=MAX(0,MIN(1,(TODAY()-DATE(2026,5,1))/(DATE(2027,4,30)-DATE(2026,5,1)+1)))', PCT],
    ['Current semester', '=IF(MONTH(TODAY())<=4,"Spring","Fall")', '@'],
    ['Unmatched entries', "=COUNTIFS('Tracker Ledger'!$T$5:$T$" + LED_LAST + ",\"No\",'Tracker Ledger'!$N$5:$N$" + LED_LAST + ',"<>Void")', '0'],
    ['Revenue received (Revenue Ledger)', '=SUMIFS(' + GB('I') + ',' + GB('D') + ',"Revenue")', '$#,##0'],
  ];
  kpis.forEach(function (k, i) {
    s.getRange(4, 2 + i).setValue(k[0]).setFontSize(9).setFontColor(C_SUB).setBackground(C_TILE).setWrap(true);
    s.getRange(5, 2 + i).setFormula(k[1]).setNumberFormat(k[2]).setFontSize(14).setFontWeight('bold').setFontColor(C_NAVY).setBackground(C_TILE).setHorizontalAlignment('left');
  });
  s.setRowHeight(5, 30);
  s.getRange('B6').setValue('Revenue - operating expenses = surplus, as in the AGM budget\'s P&L. Capital Projects are paid from NUSA\'s savings (retained earnings), so the department table below adds them: operating expenses + Capital Projects = total spending budget. Revenue received comes from the Revenue Ledger tab.')
    .setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);

  s.getRange('B7').setValue('By department (expenses)').setFontSize(12).setFontWeight('bold').setFontColor(C_NAVY);
  s.getRange('I7').setValue('Graphs: see the Charts tab').setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);
  header_(s, 8, 2, ['Department', 'Budget', 'Spent', 'Pending', 'Remaining', '% Used']);
  const rows = D.depts.map(function (d, i) {
    const r = 9 + i;
    return [d,
      '=SUMIFS(' + GB('H') + ',' + GB('B') + ',$B' + r + ',' + GB('D') + ',"Expenditure")',
      '=SUMIFS(' + GB('I') + ',' + GB('B') + ',$B' + r + ',' + GB('D') + ',"Expenditure")',
      '=SUMIFS(' + GB('J') + ',' + GB('B') + ',$B' + r + ',' + GB('D') + ',"Expenditure")',
      '=C' + r + '-D' + r, '=IF(C' + r + '>0,D' + r + '/C' + r + ',"")'];
  });
  const rt = 9 + rows.length;
  s.getRange(9, 2, rows.length, 6).setValues(rows);
  s.getRange(rt, 2, 1, 6).setValues([['Total (incl. Capital Projects)', '=SUM(C9:C' + (rt - 1) + ')', '=SUM(D9:D' + (rt - 1) + ')', '=SUM(E9:E' + (rt - 1) + ')', '=SUM(F9:F' + (rt - 1) + ')', '=IF(C' + rt + '>0,D' + rt + '/C' + rt + ',"")']])
    .setFontWeight('bold').setBorder(true, null, null, null, null, null, '#BFBFBF', SpreadsheetApp.BorderStyle.SOLID);
  s.getRange(9, 3, rows.length + 1, 4).setNumberFormat(MONEY);
  s.getRange(9, 7, rows.length + 1, 1).setNumberFormat(PCT);

  const r0 = rt + 3;     // section title row
  const sel = r0 + 1;    // department picker row
  const hr = r0 + 2;     // header row
  const first = hr + 1;  // first ranked row
  if (sel !== 23) throw new Error('Dashboard layout changed: the department picker must stay in C23 (update the protection range).');
  s.getRange(r0, 2).setValue('GL codes ranked: closest to budget first').setFontSize(12).setFontWeight('bold').setFontColor(C_NAVY);
  s.getRange(sel, 2).setValue('Department:').setFontWeight('bold');
  const pick = s.getRange(sel, 3);
  pick.setValue('All departments').setFontWeight('bold').setBackground(C_INPUT).setBorder(true, true, true, true, null, null)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['All departments'].concat(D.depts), true).setAllowInvalid(false).build());
  s.getRange(sel, 4).setValue('Pick one department or All departments. Each department is ranked by % used, then by money left. Faculty lines show the current semester\'s half (Period column); other lines show the full year.')
    .setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);
  header_(s, hr, 2, ['Department', 'GL Code', 'GL Description', 'Period', 'Budget', 'Spent', 'Pending', 'Remaining', '% Used', 'Status']);
  s.getRange(first, 2).setFormula(
    "=IFERROR(QUERY('GL Budget'!$A$5:$AI$" + gbLast + ',"select B,G,F,AA,AB,AC,AD,AE,AI,M where AH=1"&IF($C$' + sel +
    '="All departments",""," and B="""&$C$' + sel + '&"""")&" order by AF asc, AG desc, AE asc",0),"No lines to show")');
  const n = 330;
  s.getRange(first, 6, n, 4).setNumberFormat(MONEY);
  s.getRange(first, 10, n, 1).setNumberFormat(PCT);
  const body = s.getRange(first, 2, n, 10);
  const rules = [SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B' + first + '<>"",$B' + first + '<>$B' + (first - 1) + ')')
    .setBold(true).setBackground('#DCE3EE').setRanges([body]).build()];
  STATUS_FILL.forEach(function (sf) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(sf[0]).setBackground(sf[1]).setRanges([s.getRange(first, 11, n, 1)]).build());
  });
  [s.getRange(first, 10, n, 1), s.getRange(9, 7, rows.length, 1)].forEach(function (rg) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMaxpointWithValue('#5B9BD5', SpreadsheetApp.InterpolationType.NUMBER, '1').setRanges([rg]).build());
  });
  s.setConditionalFormatRules(rules);
  s.setFrozenRows(5);
  s.setHiddenGridlines(true);
}

/* ---------------------------------------------------------------- Charts */
// Charts read small live tables in columns N:S of the same tab (built from GL Budget and Tracker Ledger).
function buildCharts_(s, D, gbLast) {
  rowsNeeded_(s, 70, 20);
  const GB = function (c) { return "'GL Budget'!$" + c + '$5:$' + c + '$' + gbLast; };
  const TL = function (c) { return "'Tracker Ledger'!$" + c + '$5:$' + c + '$' + LED_LAST; };
  for (let c = 1; c <= 12; c++) s.setColumnWidth(c, 100);
  s.setColumnWidth(13, 30);
  s.getRange('A1').setValue('Expense Charts').setFontSize(18).setFontWeight('bold').setFontColor(C_NAVY);
  s.getRange('A2').setValue('Update live from the Expense Ledger. Amounts include tax; Spent = Approved + Posted. Full-year budgets (Faculty is not halved here). The charts read the tables in columns N:S; don\'t edit them.')
    .setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);
  const sub = function (a1, text) { s.getRange(a1).setValue(text).setFontWeight('bold').setFontColor(C_NAVY); };

  // Table 1: departments (N4:S14)
  const dRow = 4, dLast = dRow + D.depts.length;
  sub('N3', 'Chart data: by department');
  header_(s, dRow, 14, ['Department', 'Spent', 'Pending', 'Left to spend', '% of budget used', '% of year gone'], [190, 95, 95, 95, 95, 95]);
  const yearGone = 'MAX(0,MIN(1,(TODAY()-DATE(2026,5,1))/(DATE(2027,4,30)-DATE(2026,5,1)+1)))';
  const bud = function (r) { return 'SUMIFS(' + GB('H') + ',' + GB('B') + ',$N' + r + ',' + GB('D') + ',"Expenditure")'; };
  s.getRange(dRow + 1, 14, D.depts.length, 6).setValues(D.depts.map(function (d, i) {
    const r = dRow + 1 + i;
    return [d,
      '=SUMIFS(' + GB('I') + ',' + GB('B') + ',$N' + r + ',' + GB('D') + ',"Expenditure")',
      '=SUMIFS(' + GB('J') + ',' + GB('B') + ',$N' + r + ',' + GB('D') + ',"Expenditure")',
      '=MAX(0,' + bud(r) + '-O' + r + '-P' + r + ')',
      '=IFERROR(O' + r + '/' + bud(r) + ',0)',
      '=' + yearGone];
  }));
  s.getRange(dRow + 1, 15, D.depts.length, 3).setNumberFormat('$#,##0');
  s.getRange(dRow + 1, 18, D.depts.length, 2).setNumberFormat('0%');

  // Table 2: fiscal-year months (N18:R30)
  const mRow = 18, months = [];
  for (let i = 0; i < 12; i++) {
    const y = i < 8 ? 2026 : 2027, m = (i + 4) % 12 + 1;
    months.push([['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][m - 1] + ' ' + y, y + '-' + (m < 10 ? '0' : '') + m]);
  }
  sub('N17', 'Chart data: by month (expenses with a date in the fiscal year)');
  header_(s, mRow, 14, ['Month', 'Spent', 'Pending', 'Spent so far', 'Even pace (budget ÷ 12)']);
  const byMonth = function (key, status) { return 'SUMIFS(' + TL('J') + ',' + TL('P') + ',"' + key + '",' + TL('N') + ',"' + status + '",' + TL('T') + ',"Yes")'; };
  s.getRange(mRow + 1, 14, 12, 1).setNumberFormat('@');
  s.getRange(mRow + 1, 14, 12, 5).setValues(months.map(function (mo, i) {
    const r = mRow + 1 + i;
    return [mo[0], '=' + byMonth(mo[1], 'Approved') + '+' + byMonth(mo[1], 'Posted'), '=' + byMonth(mo[1], 'Pending'),
      '=SUM(O$' + (mRow + 1) + ':O' + r + ')', '=SUMIFS(' + GB('H') + ',' + GB('D') + ',"Expenditure")*' + (i + 1) + '/12'];
  }));
  s.getRange(mRow + 1, 15, 12, 4).setNumberFormat('$#,##0');

  // Table 3: top 10 budget lines by spending (N34:P44)
  const tRow = 34;
  sub('N33', 'Chart data: top 10 budget lines by spending');
  header_(s, tRow, 14, ['Budget line', 'Spent', 'Budget']);
  s.getRange(tRow + 1, 14).setFormula('=IFERROR(ARRAYFORMULA(SORTN(FILTER({' + GB('F') + '&" ("&' + GB('B') + '&")",' + GB('I') + ',' + GB('H') + '},' +
    GB('D') + '="Expenditure",' + GB('I') + '>0),10,0,2,FALSE)),"No spending yet")');
  s.getRange(tRow + 1, 15, 10, 2).setNumberFormat('$#,##0');

  const W = 590, H = 380;
  const chart = function (type, ranges, row, col, w, h, opts) {
    const b = s.newChart().setChartType(type).setNumHeaders(1).setPosition(row, col, 5, 0)
      .setOption('width', w).setOption('height', h).setOption('legend', { position: 'top' });
    ranges.forEach(function (a1) { b.addRange(s.getRange(a1)); });
    Object.keys(opts).forEach(function (k) { b.setOption(k, opts[k]); });
    s.insertChart(b.build());
  };
  chart(Charts.ChartType.BAR, ['N4:Q' + dLast], 4, 1, W, H,
    { title: 'Budget used by department', isStacked: true, colors: [C_NAVY, '#F4B183', '#D9D9D9'] });
  chart(Charts.ChartType.COMBO, ['N4:N' + dLast, 'R4:S' + dLast], 4, 7, W, H,
    { title: '% of budget used vs % of the year gone', seriesType: 'bars', colors: ['#5B9BD5', '#C00000'], series: { 1: { type: 'line' } } });
  chart(Charts.ChartType.COMBO, ['N' + mRow + ':R' + (mRow + 12)], 24, 1, W, H,
    { title: 'Spending by month', seriesType: 'bars', isStacked: true, colors: [C_NAVY, '#F4B183', '#70AD47', '#A6A6A6'],
      series: { 2: { type: 'line', targetAxisIndex: 1 }, 3: { type: 'line', targetAxisIndex: 1, lineDashStyle: [4, 4] } } });
  chart(Charts.ChartType.PIE, ['N4:O' + dLast], 24, 7, W, H,
    { title: 'Where the money went (spent by department)', pieHole: 0.4, legend: { position: 'right' } });
  chart(Charts.ChartType.BAR, ['N' + tRow + ':P' + (tRow + 10)], 44, 1, W * 2 + 10, 420,
    { title: 'Top 10 budget lines by spending', colors: [C_NAVY, '#D9D9D9'] });
  s.setHiddenGridlines(true);
}

/* ---------------------------------------------------------------- Dept Summary */
function buildSummary_(s, D, gbLast) {
  rowsNeeded_(s, 30, 13);
  const GB = function (c) { return "'GL Budget'!$" + c + '$5:$' + c + '$' + gbLast; };
  title_(s, 'Department Totals: Tie-out to AGM Budget', 'Revenue/Expenditure = sum of GL Budget. AGM figures (blue) are from the AGM budget\'s P&L tab. Diff should be zero. Revenue Received = Revenue Ledger rows with Status Received.');
  header_(s, 4, 1, ['Department (Tab)', 'Dept #', 'Revenue Budget', 'AGM Revenue', 'Diff', 'Expenditure Budget', 'AGM Expenditure', 'Diff', 'Net (Rev - Exp)', '# GL Lines', '# GL Lines with Budget', 'Revenue Received'],
    [180, 60, 110, 110, 80, 120, 120, 80, 110, 80, 100, 120]);
  const vals = D.agm.map(function (a, i) {
    const r = 5 + i;
    return [a[0], a[1],
      '=SUMIFS(' + GB('H') + ',' + GB('B') + ',A' + r + ',' + GB('D') + ',"Revenue")', a[2], '=C' + r + '-D' + r,
      '=SUMIFS(' + GB('H') + ',' + GB('B') + ',A' + r + ',' + GB('D') + ',"Expenditure")', a[3], '=F' + r + '-G' + r, '=C' + r + '-F' + r,
      '=COUNTIFS(' + GB('B') + ',A' + r + ')', '=COUNTIFS(' + GB('B') + ',A' + r + ',' + GB('N') + ',"Yes")',
      '=SUMIFS(' + GB('I') + ',' + GB('B') + ',A' + r + ',' + GB('D') + ',"Revenue")'];
  });
  s.getRange(5, 1, vals.length, 12).setValues(vals);
  const t = 5 + vals.length;
  const tot = ['Total', ''];
  for (let c = 3; c <= 12; c++) { const L = columnLetter_(c); tot.push('=SUM(' + L + '5:' + L + (t - 1) + ')'); }
  s.getRange(t, 1, 1, 12).setValues([tot]).setFontWeight('bold');
  s.getRange(5, 3, vals.length + 1, 7).setNumberFormat(MONEY);
  s.getRange(5, 12, vals.length + 1, 1).setNumberFormat(MONEY);
  s.getRange(5, 4, vals.length, 1).setFontColor('#0000FF'); s.getRange(5, 7, vals.length, 1).setFontColor('#0000FF');
  s.getRange(t + 1, 1, 1, 12).setValues([['Revenue Ledger: received but not on any budget line', '', '', '', '', '', '', '', '', '', '',
    "=SUMIFS('" + RL + "'!$H$2:$H$" + (RL_ROWS + 1) + ",'" + RL + "'!$I$2:$I$" + (RL_ROWS + 1) + ',"Received")-L' + t]]).setFontStyle('italic');
  s.getRange(t + 1, 12).setNumberFormat(MONEY);
  s.getRange(t + 2, 1).setValue('Note: the AGM P&L total leaves out Capital Projects. It is listed here as its own line. Revenue Ledger rows whose GL code has no revenue line in the budget (e.g. an investment fund) are counted in the line above.').setFontSize(9).setFontStyle('italic').setFontColor(C_SUB);
  s.setFrozenRows(4);
}

/* ---------------------------------------------------------------- Notes */
function buildNotes_(s) {
  s.setColumnWidth(1, 900);
  const notes = ['How the budget tracker works',
    '• Everything updates live: claims arrive on the Expense Ledger tab when VP Finance assigns GL codes and signs, and manual purchases are added with NUSA > Add manual expense (or typed on Expense Ledger). The tracker tabs read them instantly; nothing to refresh.',
    '• Status: Pending = waiting on approvals. Approved and Posted count as Spent. Void is ignored. A row with no Status counts as Approved. Amounts count including tax.',
    '• To correct a claim row on Expense Ledger, edit it and set Locked = Yes so the claim system never overwrites it.',
    '• Each GL code is charged to a budget line through the GL Map tab. If a code sits on several budget lines, the one with the largest budget is the default; set "Charge to Department" on the ledger row to pick another department. "Unmatched entries" on the Dashboard counts rows that aren\'t counted anywhere (GL code not on the map, or an unknown Status).',
    '• Faculty: every Faculty expense line is split in half, one half per semester (Fall = May-Dec, Spring = Jan-Apr), with no rollover. The semester comes from the expense date. The Dashboard\'s Current semester follows today\'s date.',
    '• GL codes that aren\'t in the AGM budget appear on GL Budget as $0 lines (Category "(Not in AGM budget)" or "Clubs (not in AGM budget)"), so spending on them shows as Unbudgeted spend.',
    '• To change a budget: edit the 2026-27 Budget column on Line Items. Rebuilding the tracker (rebuildBudgetTracker, after a new AGM budget) keeps your Line Items and GL Map edits, unless the AGM changed that same line; it also saves a copy of the old Line Items on "Line Items (before rebuild)". Expenses are on Expense Ledger, which a rebuild never touches.',
    '• Charts tab: budget used by department, % of budget used vs % of the year gone, spending by month (with a running total and an even-pace line), where the money went, and the top 10 budget lines. They update live; expenses without a date are left out of the by-month chart only.',
    '• The formula tabs are protected with a warning so they are not changed by accident. The Dashboard\'s department picker (C23) is free to change.',
    '• Web dashboard: File > Download > Microsoft Excel (.xlsx), then load that file on the NUSA Budget Tracker page.',
    '• Revenue Ledger tab: one row per amount NUSA receives (student fees, events revenue, sponsorships, transfers in, investment gains). Pick the revenue GL code, type the amount as a positive number and set Status (Received, Expected or Void). GL Budget shows each revenue line\'s received amount (Spent column) and expected amount (Pending column); the Dashboard shows "Revenue received"; Dept Summary shows revenue received by department and any received revenue whose GL code has no budget line. A rebuild never touches this tab.',
    '• Refunds are not revenue: when money comes back for something NUSA paid (e.g. a software refund or a deposit returned), add it on Expense Ledger as a negative amount on the same GL code, so that budget line\'s spending goes down.',
    '• Loan repayment to the university (Capital Projects, no GL in the budget): the university booked the payment on NUSA Funding Clubs (410-72111-56025-10000-39070-10000). On Expense Ledger those rows have Charge to Department = Capital Projects, so they count on the loan repayment line, not on the Proposal Committee. The university also shows each club receiving the same amount (club "Transfers In"); that is not NUSA revenue, so it is not on the Revenue Ledger.',
    '',
    'Things to know about the budget file',
    '• Capital Projects: revenue is the "Fund Balance at the beginning of the period", the figure the budget P&L uses. "Loan repayment to the university" has GL "N/A" in the budget; the repayment booked on NUSA Funding Clubs counts there (see above).',
    '• Proposals Committee Fund lines on the Faculty tab share their GL codes with the Proposal Committee tab; spending on a shared code counts against the line with the largest budget unless "Charge to Department" says otherwise.',
    '• Every km claim is charged to Travel-Mileage (Personal Vehicle) 410-72103-56302, budgeted on the Student Lounge (The Commons) tab. Km claims signed earlier keep the code they were signed with.'];
  s.getRange(1, 1, notes.length, 1).setValues(notes.map(function (t) { return [t]; })).setWrap(true).setVerticalAlignment('top');
  s.getRange(1, 1).setFontWeight('bold').setFontSize(13).setFontColor(C_NAVY);
  s.getRange(notes.indexOf('Things to know about the budget file') + 1, 1).setFontWeight('bold').setFontSize(13).setFontColor(C_NAVY);
}
