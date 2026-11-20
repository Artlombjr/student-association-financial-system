// Full GL List with a Type column ("GL Codes" script): the claim log's "GL Codes" script writes it, the claim web app 
// keeps it, the manual expense form and the ledger dropdown only offer Expense codes, and the tracker has a line for
// every expense code.
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const P = require("./paths").file;  // file locations (one folder per Apps Script project)
const glc = fs.readFileSync(P("GlCodes.gs"), "utf8");
const manual = fs.readFileSync(P("ManualExpense.gs"), "utf8");
const J = (x) => JSON.parse(JSON.stringify(x));

function makeSheet(name) {
  return { rows: [], name, validation: null, filter: null,
    getDataRange() { const rows = this.rows; return { getValues: () => (rows.length ? rows : [[""]]).map(r => r.slice()) }; },
    getLastRow() { return this.rows.length; }, getLastColumn() { return Math.max(1, ...this.rows.map(r => r.length)); },
    appendRow(r) { this.rows.push(r.slice()); },
    getRange(r, c, nr = 1, nc = 1) { const self = this; const rg = {
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => ((self.rows[r - 1 + i] || [])[c - 1 + j] ?? ""))),
      setValues(v) { assert.strictEqual(v.length, nr); for (let i = 0; i < nr; i++) { assert.strictEqual(v[i].length, nc, "row width"); self.rows[r - 1 + i] = self.rows[r - 1 + i] || []; for (let j = 0; j < nc; j++) self.rows[r - 1 + i][c - 1 + j] = v[i][j]; } return rg; },
      setNumberFormat: () => rg, setFontWeight: () => rg, setBackground: () => rg, setFontColor: () => rg,
      setDataValidation(v) { self.validation = { col: c, rows: nr, rule: v }; return rg; },
      createFilter() { self.filter = { remove() { self.filter = null; } }; return self.filter; } }; return rg; },
    getFilter() { return this.filter; }, setFrozenRows() {}, setColumnWidth() {},
    clear() { this.rows = []; }, clearContents() { this.rows = []; } };
}
const rule = () => { const b = { list: null, range: null, requireValueInList(v) { b.list = v.slice(); return b; }, requireValueInRange(r) { b.range = r; return b; }, setAllowInvalid(x) { b.allowInvalid = x; return b; }, build: () => ({ list: b.list, range: b.range, allowInvalid: b.allowInvalid }) }; return b; };
function newSS() {
  const sheets = {};
  return { sheets, getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = makeSheet(n)), getId: () => "SS" };
}

// ---- 1. the claim log's GL Codes script
const ss = newSS();
const led = ss.insertSheet("Expense Ledger");
const v5 = fs.readFileSync(P("Code.gs"), "utf8");
const v5ctx = { console: { log() {} }, Math, Date, JSON, String, Object, Array, parseFloat, SpreadsheetApp: { newDataValidation: rule } };
vm.createContext(v5ctx);
vm.runInContext(v5 + "\n__o.LEDGER_HEADERS = LEDGER_HEADERS; __o.GL = GL; __o.MILEAGE_GL = MILEAGE_GL; __o.list = ss_glListRange_;", Object.assign(v5ctx, { __o: v5ctx }));
led.rows.push(v5ctx.LEDGER_HEADERS.slice());
const old = ss.insertSheet("GL List");
old.rows = [["GL Code", "GL Name", "Group"], ["410-72103-50340100003802310000", "Wages - Student Work (Grants)", "Administration"], ["410-99999-55600100001000010000", "Some old code", "Administration"]];
const ctx = { console: { log() {} }, Math, Date, JSON, String, Object, Array, Number, parseFloat,
  SpreadsheetApp: { getActive: () => ss, newDataValidation: rule, flush() {} },
  LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Session: { getScriptTimeZone: () => "America/Vancouver" },
  Utilities: { formatDate: (d) => d.toISOString().slice(0, 10) } };
vm.createContext(ctx);
vm.runInContext(glc + "\n" + manual + "\n__o.GL_CODES = GL_CODES; __o.GLC_HEAD = GLC_HEAD;", Object.assign(ctx, { __o: ctx }));
const codes = J(ctx.GL_CODES);
const n = ctx.updateGlList();
const gl = ss.sheets["GL List"];
assert.deepStrictEqual(gl.rows[0], J(ctx.GLC_HEAD));
assert.deepStrictEqual(gl.rows[0].slice(0, 4), ["GL Code", "GL Name", "Group", "Type"]);
assert.strictEqual(n, codes.length + 1, "the old unknown code is kept");
assert.deepStrictEqual(gl.rows[gl.rows.length - 1].slice(0, 4), ["410-99999-55600100001000010000", "Some old code", "Administration", "Expense"]);
assert.ok(gl.filter, "filter on the header row");
const body = gl.rows.slice(1);
assert.ok(body.every(r => r[3] === "Revenue" || r[3] === "Expense"), "every row is Revenue or Expense");
assert.ok(body.every(r => /^41[03]-\d{5}-\d{20}$/.test(r[0])), "codes in the claim form's format");
assert.strictEqual(new Set(body.map(r => r[0])).size, body.length, "no duplicate codes");
assert.ok(body.some(r => r[3] === "Revenue") && body.filter(r => r[3] === "Expense").length > 40);
// type follows the account number: 4xxxx / 65075 / 83200 = Revenue, 5xxxx / 66000 = Expense
body.forEach(r => { const a = r[0].slice(10, 15); const rev = a[0] === "4" || a === "65075" || a === "83200"; assert.strictEqual(r[3], rev ? "Revenue" : "Expense", r[0] + " " + r[3]); });
// every claim-form code (and mileage) is on the list as Expense, with the claim form's name
const claim = []; v5ctx.GL.forEach(g => g.items.forEach(it => claim.push(it)));
claim.forEach(([name, code]) => { const row = body.find(r => r[0] === code); assert.ok(row, "claim code on the list: " + code); assert.strictEqual(row[3], "Expense"); assert.strictEqual(row[5], "Yes"); });
assert.ok(body.find(r => r[0] === v5ctx.MILEAGE_GL), "mileage code");
assert.ok(!claim.some(([, c]) => c.split("-")[2].startsWith("4")), "no income (4xxxx) codes in the claim form");
assert.ok(claim.some(([, c]) => c === "410-72111-56025100003907010000") && claim.some(([, c]) => c === "410-72105-55610100001000010000"));
// ledger dropdown = expense codes only
const v = led.validation;
assert.strictEqual(v.col, v5ctx.LEDGER_HEADERS.indexOf("GL Code") + 1);
assert.ok(v.rule.list.length < 500, "Google's dropdown list limit");
assert.ok(v.rule.list.every(c => body.find(r => r[0] === c)[3] === "Expense"));
assert.ok(v.rule.allowInvalid, "old codes on past rows stay allowed");

// ---- 2. manual expense form: expense codes only, every one reachable from a department
const setup = ctx.getManualExpenseSetup();
const rev = body.filter(r => r[3] === "Revenue").map(r => r[0]);
assert.ok(!setup.gls.some(g => rev.includes(g.code)), "no revenue codes in the manual expense form");
assert.ok(setup.gls.every(g => g.depts.length >= 1), "every code reachable from a department");
const club = setup.gls.find(g => g.code === "413-73041-56025100001000010000");
assert.ok(club && club.depts.includes("Faculty"), "club code under Faculty");
assert.ok(setup.gls.find(g => g.code === "410-72110-56025100001000010000").depts.includes("Special Initiatives"));
const r1 = ctx.addManualExpense({ dept: "Communications", gl: "410-72105-55610100001000010000", date: "2026-10-15", amount: "120", desc: "Light stand", status: "Approved" });
assert.strictEqual(r1.id, "M-001");
assert.throws(() => ctx.addManualExpense({ dept: "Administration", gl: rev.find(c => c.startsWith("410-72103")), date: "2026-10-15", amount: "5", desc: "x", status: "Approved" }), /GL code from the list/);

// ---- 3. the claim web app keeps the full list  and still rewrites an old 3-column list
for (const file of ["Code.gs"]) {
  const src = fs.readFileSync(P(file), "utf8");
  const c2 = { console: { log() {} }, Math, Date, JSON, String, Object, Array, parseFloat, SpreadsheetApp: { newDataValidation: rule } };
  vm.createContext(c2);
  vm.runInContext(src + "\n__o.list = ss_glListRange_; __o.GL = GL;", Object.assign(c2, { __o: c2 }));
  const s2 = newSS(); const t = s2.insertSheet("GL List");
  t.rows = gl.rows.filter(r => r[0] !== "410-72105-55610100001000010000").map(r => r.slice()); // one claim code missing
  const before = t.rows.length;
  const list = c2.list(s2);
  assert.strictEqual(t.rows.length, before + 1, file + ": missing claim code appended, nothing removed");
  assert.deepStrictEqual(t.rows[t.rows.length - 1].slice(0, 4), ["410-72105-55610100001000010000", "Office - Small Equipment", "Communications", "Expense"]);
  assert.strictEqual(t.rows[1][8], gl.rows[1][8], file + ": other columns untouched");
  assert.ok(!list.some(c => rev.includes(c)) && list.includes("410-72105-55610100001000010000"), file + ": dropdown = expense codes");
  const s3 = newSS(); const t3 = s3.insertSheet("GL List");
  t3.rows = [["GL Code", "GL Name", "Group"], ["x", "y", "z"]];
  const l3 = c2.list(s3);
  assert.deepStrictEqual(t3.rows[0], ["GL Code", "GL Name", "Group", "Type"], file + ": old list rewritten with a Type column");
  assert.ok(t3.rows.slice(1).every(r => r[3] === "Expense") && !t3.rows.some(r => r[0] === "x"));
  assert.strictEqual(l3.length, t3.rows.length - 1);
}

// ---- 4. the budget tracker has a line for every expense code on the list
const td = fs.readFileSync(P("TrackerData.gs"), "utf8");
const T = JSON.parse(td.slice(td.indexOf("{"), td.lastIndexOf("}") + 1));
const mapped = new Set(T.map.map(m => String(m[0]).split("|")[0]));
const missing = body.filter(r => r[3] === "Expense" && !r[8].startsWith("Kept")).map(r => r[0]).filter(c => !mapped.has(c.replace(/\D/g, "")));
assert.deepStrictEqual(missing, [], "expense codes without a tracker line");
console.log("ALL GL LIST TESTS PASSED (" + body.length + " codes: " + rev.length + " revenue, " + (body.length - rev.length) + " expense)");
