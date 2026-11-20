// Runs the real "Manual Expense - Code.gs" against a simulated claim log (Expense Ledger + GL List built by V3's own code).
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const P = require("./paths").file;  // file locations (one folder per Apps Script project)
const manual = fs.readFileSync(P("ManualExpense.gs"), "utf8");
const v3 = fs.readFileSync(P("Code.gs"), "utf8");

function makeSheet(name) {
  return { rows: [], name,
    getDataRange() { const rows = this.rows; return { getValues: () => rows.map(r => r.slice()) }; },
    getLastRow() { return this.rows.length; }, getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); },
    appendRow(r) { this.rows.push(r.slice()); },
    getRange(r, c, nr = 1, nc = 1) { const self = this; return {
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => ((self.rows[r - 1 + i] || [])[c - 1 + j] ?? ""))),
      setValues(v) { for (let i = 0; i < nr; i++) { self.rows[r - 1 + i] = self.rows[r - 1 + i] || []; for (let j = 0; j < nc; j++) self.rows[r - 1 + i][c - 1 + j] = v[i][j]; } return this; },
      setNumberFormat() { return this; }, setDataValidation() { return this; }, setFontWeight() { return this; } }; },
    setFrozenRows() {}, clearContents() { this.rows = []; } };
}
const sheets = {};
const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = makeSheet(n)), getId: () => "SS" };

// Build the real Expense Ledger header + GL List using V3's own functions.
const v3ctx = { console, Math, Date, JSON, String, Object, Array, parseFloat,
  SpreadsheetApp: { newDataValidation: () => { const b = { requireValueInList: () => b, requireValueInRange: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; } } };
vm.createContext(v3ctx);
vm.runInContext(v3 + "\n__o.LEDGER_HEADERS = LEDGER_HEADERS; __o.list = ss_glListRange_;", Object.assign(v3ctx, { __o: v3ctx }));
const led = ss.insertSheet("Expense Ledger"); led.rows.push(v3ctx.LEDGER_HEADERS.slice());
v3ctx.list(ss);
// a claim row already in the ledger (must stay untouched)
const claimRow = v3ctx.LEDGER_HEADERS.map(() => ""); claimRow[0] = "abcd1234-1"; claimRow[1] = "Claim"; led.rows.push(claimRow);

const ctx = { console, Math, Date, JSON, String, Object, Array, Number, parseFloat,
  SpreadsheetApp: { getActive: () => ss, flush() {}, getUi: () => ({}) },
  LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Session: { getScriptTimeZone: () => "America/Vancouver" },
  Utilities: { formatDate: (d, tz, f) => f === "yyyy-MM-dd" ? d.toISOString().slice(0, 10) : "Sep " + d.getDate() } };
vm.createContext(ctx);
vm.runInContext(manual, ctx);

const setup = ctx.getManualExpenseSetup();
assert.strictEqual(setup.depts.length, 10);
const byCode = c => setup.gls.find(g => g.code === c);
const sp = byCode("410-72111-56025100003907010000");
assert.strictEqual(JSON.stringify(sp.depts.slice().sort()), JSON.stringify(["Faculty", "Proposal Committee"]), "72111 Funding Clubs offered under both");
assert.strictEqual(JSON.stringify(byCode("413-73006-56025100003904410000").depts.slice().sort()), JSON.stringify(["Faculty"]), "club code under Faculty");
assert.strictEqual(JSON.stringify(byCode("410-72109-55600100001000010000").depts.slice().sort()), JSON.stringify(["Student Lounge (The Commons)"]), "Lounge Office Misc");
assert.ok(setup.gls.every(g => g.depts.length >= 1), "every GL code reachable from some department");
assert.ok(!setup.gls.some(g => g.group === "Student Lounge (The Commons)" && g.depts.includes("Administration")), "Lounge codes not offered under Administration");
const mix = byCode("410-72111-56025100001000010000");
assert.strictEqual(JSON.stringify(mix.depts.slice().sort()), JSON.stringify(["Faculty", "Proposal Committee"]), "72111 MixEx under both");

const H = v3ctx.LEDGER_HEADERS, col = h => H.indexOf(h);
const r1 = ctx.addManualExpense({ dept: "Student Lounge (The Commons)", gl: "410-72109-55612100003902110000", date: "2026-09-20", amount: "250", desc: "Supply run", payee: "Wholesale Store", status: "Approved" });
assert.strictEqual(r1.id, "M-001");
const row = led.rows[led.rows.length - 1];
assert.strictEqual(row.length, H.length);
assert.strictEqual(row[col("Source")], "Manual");
assert.strictEqual(row[col("GL Name")], "Food Costs");
assert.strictEqual(row[col("Charge to Department (optional)")], "Student Lounge (The Commons)");
assert.strictEqual(row[col("Amount (incl. tax)")], 250);
assert.strictEqual(row[col("Net Expense")], 250);
assert.strictEqual(row[col("Status")], "Approved");
assert.strictEqual(row[col("Claim ID")], "", "no claim id on manual rows");
assert.strictEqual(row[col("Date")].getDate(), 20);
const r2 = ctx.addManualExpense({ dept: "Proposal Committee", gl: "410-72111-56025100003907010000", date: "2027-02-01", amount: "-40.5", gst: "2", desc: "Refund", status: "Posted" });
assert.strictEqual(r2.id, "M-002");
assert.strictEqual(led.rows[led.rows.length - 1][col("Net Expense")], -41.84);
assert.strictEqual(r2.recent[0].id, "M-002");
assert.strictEqual(led.rows[1][0], "abcd1234-1", "claim row untouched");

const bad = (f, re) => assert.throws(() => ctx.addManualExpense(f), re);
bad({ dept: "Events", gl: "410-72109-55612100003902110000", date: "2026-09-20", amount: 5, desc: "x", status: "Approved" }, /belongs to Events/);
bad({ dept: "Events", gl: "999", date: "2026-09-20", amount: 5, desc: "x", status: "Approved" }, /GL code from the list/);
bad({ dept: "Nope", gl: "", date: "", amount: 0, desc: "", status: "Void" }, /department.*date.*description.*amount.*status/);
assert.strictEqual(led.rows.length, 4, "rejected entries add nothing");
console.log("ALL MANUAL EXPENSE TESTS PASSED (" + setup.gls.length + " GL codes)");
