// Runs the real "Claim Tracking - Code.gs" against a simulated Expense Ledger (header from V5's own LEDGER_HEADERS).
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const P = require("./paths").file;  // file locations (one folder per Apps Script project)
const code = fs.readFileSync(P("ClaimTracking.gs"), "utf8");
const v5 = fs.readFileSync(P("Code.gs"), "utf8");
const H = v5.match(/const LEDGER_HEADERS = (\[[\s\S]*?\]);/)[1];
const HEAD = eval(H);

const validations = {};
function makeSheet(name) {
  const sh = { rows: [], maxRows: 1000, maxCols: HEAD.length, widths: {}, writes: 0,
    getName: () => name,
    getLastRow() { return this.rows.length; }, getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); },
    getMaxRows() { return this.maxRows; }, getMaxColumns() { return this.maxCols; },
    insertRowsAfter(at, n) { this.maxRows += n; }, insertColumnsAfter(at, n) { this.maxCols += n; },
    setColumnWidth(c, w) { this.widths[c] = w; },
    getRange(r, c, nr = 1, nc = 1) {
      const self = this;
      assert.ok(r + nr - 1 <= self.maxRows && c + nc - 1 <= self.maxCols, `range ${r},${c},${nr},${nc} outside the sheet`);
      const api = {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => ((self.rows[r - 1 + i] || [])[c - 1 + j] ?? ""))),
        setValues(v) { self.writes++; for (let i = 0; i < nr; i++) { self.rows[r - 1 + i] = self.rows[r - 1 + i] || []; for (let j = 0; j < nc; j++) self.rows[r - 1 + i][c - 1 + j] = v[i][j]; } return api; },
        setValue(v) { return api.setValues([[v]]); },
        setFontWeight() { return api; },
        setDataValidation(dv) { validations[c] = { from: r, rows: nr, values: dv.values }; return api; },
        getRow: () => r, getLastRow: () => r + nr - 1, getColumn: () => c, getNumColumns: () => nc, getSheet: () => self,
      };
      return api;
    } };
  return sh;
}
const led = makeSheet("Expense Ledger");
led.rows.push(HEAD.slice());
const C = h => HEAD.indexOf(h);
const row = (id, claim) => { const r = HEAD.map(() => ""); r[C("Entry ID")] = id; r[C("Claim ID")] = claim; return r; };
led.rows.push(row("aaaa1111-1", "aaaa1111"), row("aaaa1111-2", "aaaa1111"), row("aaaa1111-3", "aaaa1111"),
  row("M-001", ""), row("bbbb2222-1", "bbbb2222"), row("M-002", ""));
const other = makeSheet("Claims"); other.maxCols = 30;
const ss = { getSheetByName: n => ({ "Expense Ledger": led, "Claims": other })[n] || null };
const dv = () => { const b = { values: null, requireValueInList(v) { b.values = v; return b; }, setAllowInvalid: () => b, build: () => b }; return b; };
const logs = [];
const ctx = { console: { log: s => logs.push(s) }, Math, String, Object, Array,
  SpreadsheetApp: { getActive: () => ss, newDataValidation: dv } };
vm.createContext(ctx);
vm.runInContext(code, ctx);

// ---- setup adds U, V, W after Notes, with their dropdowns down to row 3001
ctx.addClaimTrackingColumns();
assert.strictEqual(HEAD.length, 20, "V5 ledger is A–T");
assert.deepStrictEqual(led.rows[0].slice(20), ["Submitted", "Cheque Received from AP", "Cheque Picked Up"]);
assert.deepStrictEqual(led.rows[0].slice(0, 20), HEAD, "existing headers untouched");
assert.strictEqual(JSON.stringify(validations[21]), JSON.stringify({ from: 2, rows: 3000, values: ["Received by NUSA", "Submitted to Accounts Payable", "Sent back for changes"] }));
assert.strictEqual(JSON.stringify(validations[22].values), JSON.stringify(["Yes", "No"]));
assert.strictEqual(JSON.stringify(validations[23].values), JSON.stringify(["Yes", "No"]));
assert.ok(led.maxRows >= 3001 && led.maxCols === 23);
assert.match(logs.pop(), /Added columns: Submitted, Cheque Received from AP, Cheque Picked Up/);
// running again adds nothing
ctx.addClaimTrackingColumns();
assert.strictEqual(led.rows[0].length, 23);
assert.match(logs.pop(), /already there/);

// ---- picking a value on one line of a claim fills every line of that claim, nothing else
const edit = (rowNum, col, v, nr = 1) => { for (let i = 0; i < nr; i++) led.rows[rowNum - 1 + i][col - 1] = v; ctx.onEdit({ range: led.getRange(rowNum, col, nr, 1) }); };
edit(3, 21, "Submitted to Accounts Payable"); // aaaa1111 line 2
assert.deepStrictEqual([2, 3, 4].map(r => led.rows[r - 1][20]), Array(3).fill("Submitted to Accounts Payable"));
assert.strictEqual(led.rows[5][20] ?? "", "", "manual row untouched");
assert.strictEqual(led.rows[4][20] ?? "", "", "other manual row untouched");
assert.strictEqual(led.rows[5][20] ?? "", "");
assert.strictEqual(led.rows[6 - 1][20] ?? "", "", "other claim untouched");
edit(6, 22, "Yes"); // bbbb2222 single line
assert.strictEqual(led.rows[5][21], "Yes");
assert.ok(![1, 2, 3, 4].some(r => led.rows[r][21] === "Yes"));
edit(5, 23, "Yes"); // manual row: only itself
assert.strictEqual(led.rows[4][22], "Yes");
assert.ok(![1, 2, 3, 5, 6].some(r => led.rows[r][22] === "Yes"));
edit(2, 21, ""); // clearing one line clears the claim
assert.deepStrictEqual([2, 3, 4].map(r => led.rows[r - 1][20]), ["", "", ""]);
// edits elsewhere do nothing
const before = led.writes;
edit(2, C("Notes") + 1, "note"); edit(1, 21, "Submitted"); ctx.onEdit({ range: other.getRange(2, 21) });
ctx.onEdit({ range: led.getRange(2, 21, 1, 2) }); ctx.onEdit({});
assert.strictEqual(led.writes, before, "edits outside the three columns copy nothing");
assert.strictEqual(led.rows[2][C("Notes")] ?? "", "", "other lines' notes not copied");
console.log("ALL CLAIM TRACKING TESTS PASSED");
