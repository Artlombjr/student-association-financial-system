/**
 * NUSA — Add manual expense  (script bound to "NUSA Expense Claims — Log")
 * ------------------------------------------------------------------
 * Adds a "NUSA" menu to the claim log with "Add manual expense", which opens
 * a sidebar form and writes one row to the "Expense Ledger" tab in the same
 * format the claim system (Code.gs V3) and the budget tracker use.
 *
 * Only people who can EDIT this spreadsheet can open the menu or run it, so
 * with the log's sharing set to Restricted it is only you. It is a separate
 * project from the claim web app: nothing here is reachable from the public link.
 *
 * SETUP (once): in the claim log, Extensions > Apps Script. Paste this file into
 * Code.gs, add an HTML file named exactly "Sidebar" with "ManualExpenseSidebar.html",
 * Save, then reload the spreadsheet. The first time you use the menu, Google asks
 * you to authorize it (your own script, your own sheet).
 * Needs the Expense Ledger + GL List tabs, created by setupExpenseLedger() in V3.
 */

const LEDGER_TAB = 'Expense Ledger';
const GL_TAB = 'GL List';
const STATUSES = ['Approved', 'Posted', 'Pending'];
// Department names exactly as the Excel tracker uses them, with each one's department number.
const DEPTS = [
  ['Administration', '72103'], ['Faculty', '72104'], ['Communications', '72105'], ['Events', '72106'],
  ["The Courier", '72107'], ['Yearbook', '72108'], ['Student Lounge (The Commons)', '72109'],
  ['Special Initiatives', '72110'], ['Proposal Committee', '72111'], ['Capital Projects', '73001'],
];
// GL List groups that belong to a different tracker department.
const GROUP_TO_DEPT = { 'Clubs (Fund 413)': 'Faculty' };
// Proposal Committee (72111) codes are budgeted both on the Proposal Committee tab and in Faculty's
// "Proposals Committee Fund", so they can be charged to either department.
const SHARED_DEPTS = { '72111': ['Faculty', 'Proposal Committee'] };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('NUSA')
    .addItem('Add manual expense', 'showManualExpense')
    .addItem('Email me claims waiting for my signature', 'emailMeUnsignedClaims') // file "Unsigned Claims"
    .addItem('Remind the President to sign claims', 'remindPresidentClaims') // file "President Reminders"
    .addItem('Remind people with sent-back claims', 'remindSentBackClaims') // file "Sent Back Reminders"
    .addToUi();
}

function showManualExpense() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Add manual expense');
  SpreadsheetApp.getUi().showSidebar(html);
}

function ledgerSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(LEDGER_TAB);
  if (!sh) throw new Error('No "' + LEDGER_TAB + '" tab yet. Run setupExpenseLedger in the claim web app\'s Apps Script first.');
  return sh;
}

// Column index (0-based) of every ledger header, so the form keeps working if columns move.
function ledgerColumns_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = {};
  head.forEach(function (h, i) { if (h !== '') col[String(h).trim()] = i; });
  ['Entry ID', 'Source', 'Date', 'GL Code', 'Amount (incl. tax)', 'Status'].forEach(function (h) {
    if (col[h] === undefined) throw new Error('The Expense Ledger tab is missing its "' + h + '" column.');
  });
  return { col: col, width: head.length };
}

// Every expense GL code offered, and the tracker departments it can be charged to: the department of its
// GL List group (falling back to the department its number belongs to), plus SHARED_DEPTS.
// Revenue codes on the GL List (column D "Type" = Revenue) are left out: this form records expenses.
function glOptions_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(GL_TAB);
  if (!sh) throw new Error('No "' + GL_TAB + '" tab yet. Run setupExpenseLedger in the claim web app\'s Apps Script first.');
  const rows = sh.getDataRange().getValues().slice(1).filter(function (r) { return r[0] && String(r[3] || '').trim() !== 'Revenue'; });
  const byNum = {}; DEPTS.forEach(function (d) { byNum[d[1]] = d[0]; });
  const isDept = function (name) { return DEPTS.some(function (d) { return d[0] === name; }); };
  const seen = {};
  const out = [];
  rows.forEach(function (r) {
    const code = String(r[0]).trim(), name = String(r[1] || '').trim(), group = String(r[2] || '').trim();
    if (seen[code]) return; seen[code] = true;
    const num = deptNumber_(code);
    const g = GROUP_TO_DEPT[group] || group;
    const depts = [];
    if (isDept(g)) depts.push(g); else if (byNum[num]) depts.push(byNum[num]);
    (SHARED_DEPTS[num] || []).forEach(function (d) { if (depts.indexOf(d) < 0) depts.push(d); });
    out.push({ code: code, name: name, group: group, depts: depts });
  });
  return out;
}

function deptNumber_(code) { const d = String(code).replace(/\D/g, ''); return d.slice(3, 8); }

// Data for the sidebar.
function getManualExpenseSetup() {
  const sh = ledgerSheet_();
  ledgerColumns_(sh);
  const tz = Session.getScriptTimeZone();
  return {
    depts: DEPTS.map(function (d) { return d[0]; }),
    gls: glOptions_(),
    statuses: STATUSES,
    today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    recent: recentManual_(sh, 5),
  };
}

function recentManual_(sh, n) {
  const c = ledgerColumns_(sh).col;
  const data = sh.getDataRange().getValues().slice(1);
  const tz = Session.getScriptTimeZone();
  return data.filter(function (r) { return String(r[c['Source']]).toLowerCase() === 'manual'; }).slice(-n).reverse().map(function (r) {
    const d = r[c['Date']];
    return {
      id: String(r[c['Entry ID']] || ''),
      date: d instanceof Date ? Utilities.formatDate(d, tz, 'MMM d') : String(d || ''),
      desc: c['Description'] !== undefined ? String(r[c['Description']] || '') : '',
      amount: Number(r[c['Amount (incl. tax)']]) || 0,
      status: String(r[c['Status']] || ''),
    };
  });
}

function nextEntryId_(sh, c) {
  const ids = sh.getRange(2, c['Entry ID'] + 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  let max = 0;
  ids.forEach(function (r) { const m = String(r[0]).match(/^M-(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  return 'M-' + String(max + 1).padStart(3, '0');
}

function round2_(n) { return Math.round((parseFloat(n) || 0) * 100) / 100; }

// Called by the sidebar. Validates everything again on the server, then appends one ledger row.
function addManualExpense(f) {
  f = f || {};
  const dept = String(f.dept || '').trim();
  const code = String(f.gl || '').trim();
  const desc = String(f.desc || '').trim();
  const status = String(f.status || 'Approved').trim();
  const amount = round2_(f.amount), gst = round2_(f.gst), hst = round2_(f.hst);
  const m = String(f.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);

  const problems = [];
  if (!DEPTS.some(function (d) { return d[0] === dept; })) problems.push('a department');
  const gl = glOptions_().filter(function (g) { return g.code === code; })[0];
  if (!gl) problems.push('a GL code from the list');
  else if (gl.depts.indexOf(dept) < 0) problems.push('a GL code that belongs to ' + dept);
  if (!m) problems.push('a date');
  if (!desc) problems.push('a description');
  if (!amount) problems.push('an amount (use a negative amount for a refund)');
  if (STATUSES.indexOf(status) < 0) problems.push('a status');
  if (problems.length) throw new Error('Please add ' + problems.join(', ') + '.');

  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = ledgerSheet_();
    const L = ledgerColumns_(sh);
    const c = L.col;
    const id = nextEntryId_(sh, c);
    const row = new Array(L.width).fill('');
    const set = function (h, v) { if (c[h] !== undefined) row[c[h]] = v; };
    set('Entry ID', id);
    set('Source', 'Manual');
    set('Date', new Date(+m[1], +m[2] - 1, +m[3]));
    set('GL Code', code);
    set('GL Name', gl.name);
    set('GL Group', gl.group);
    set('Charge to Department (optional)', dept);
    set('Description', desc);
    set('Payee', String(f.payee || '').trim());
    set('Amount (incl. tax)', amount);
    set('GST', gst);
    set('HST', hst);
    set('Net Expense', round2_(amount - gst * 0.67 - hst * 0.7377));
    set('Status', status);
    set('Last Updated', new Date());
    set('Notes', String(f.notes || '').trim());
    sh.appendRow(row);
    SpreadsheetApp.flush();
    return { id: id, recent: recentManual_(sh, 5) };
  } finally {
    lock.releaseLock();
  }
}
