// Runs "Tracker - Code.gs" + "Tracker Data.gs" against a strict simulated spreadsheet:
// only real Apps Script method names exist, so a typo'd call throws. Then checks sizes and every formula written.
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const P = require("./paths").file;  // file locations (one folder per Apps Script project)
const code = fs.readFileSync(P("Tracker.gs"), "utf8");
const data = fs.readFileSync(P("TrackerData.gs"), "utf8");

const colNum = L => L.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
function parseA1(a1) { const m = a1.match(/^([A-Z]+)(\d+)$/); return [+m[2], colNum(m[1])]; }

function makeSheet(ss, name) {
  const sh = { name, cells: {}, formats: {}, maxRows: 1000, maxCols: 26, frozen: [0, 0], hidden: [], cf: [], filters: [], prot: null, widths: {}, charts: [] };
  const api = {
    // EmbeddedChartBuilder: only real method names
    newChart: () => {
      const c = { ranges: [], opts: {}, type: null, pos: null, headers: null };
      const b = {
        setChartType: t => { assert.ok(t, "chart type"); c.type = t; return b; },
        addRange: rg => { assert.ok(rg && rg.a1, "chart range"); c.ranges.push(rg); return b; },
        setNumHeaders: n => { c.headers = n; return b; },
        setOption: (k, v) => { assert.ok(v !== undefined, "option " + k); c.opts[k] = v; return b; },
        setPosition: (r, col, x, y) => { assert.ok(r >= 1 && col >= 1 && x >= 0 && y >= 0, "chart position"); c.pos = [r, col]; return b; },
        build: () => c,
      };
      return b;
    },
    insertChart: c => { sh.charts.push(c); },
    getName: () => name,
    setName: n => { sh.name = n; return api; },
    copyTo: (target) => { const c = makeSheet(ss, "Copy of " + sh.name); Object.assign(c.cells, sh.cells); sheets.push(c); return c.api; },
    getLastRow: () => Object.keys(sh.cells).reduce((m, k) => Math.max(m, +k.split(",")[0]), 0),
    getLastColumn: () => Object.keys(sh.cells).reduce((m, k) => Math.max(m, +k.split(",")[1]), 0),
    getMaxRows: () => sh.maxRows, getMaxColumns: () => sh.maxCols,
    insertRowsAfter: (a, n) => { sh.maxRows += n; }, insertColumnsAfter: (a, n) => { sh.maxCols += n; },
    setColumnWidth: (c, w) => { sh.widths[c] = w; }, setRowHeight: () => {},
    setFrozenRows: n => { sh.frozen[0] = n; }, setFrozenColumns: n => { sh.frozen[1] = n; },
    hideColumns: (c, n) => { sh.hidden.push([c, n]); }, setHiddenGridlines: () => {},
    setConditionalFormatRules: r => { sh.cf = r; },
    protect: () => { const p = { setDescription: () => p, setWarningOnly: () => p, setUnprotectedRanges: r => { sh.unprot = r.map(x => x.a1); return p; } }; sh.prot = p; return p; },
    getRange: (r, c, nr = 1, nc = 1) => {
      if (typeof r === "string") {
        const [a, b] = r.split(":"); [r, c] = parseA1(a);
        if (b) { const [r2, c2] = parseA1(b); assert.ok(r2 >= r && c2 >= c, "range order"); nr = r2 - r + 1; nc = c2 - c + 1; }
      }
      assert.ok(r >= 1 && c >= 1, "bad range");
      if (r + nr - 1 > sh.maxRows || c + nc - 1 > sh.maxCols) throw new Error(`${name}: range R${r}C${c}+${nr}x${nc} outside sheet ${sh.maxRows}x${sh.maxCols}`);
      const R = { a1: `${name}!R${r}C${c}`, r, c, nr, nc };
      const each = fn => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) fn(r + i, c + j, i, j); return rg; };
      const rg = {
        a1: R.a1,
        setValue: v => { sh.cells[`${r},${c}`] = v; return rg; },
        getValue: () => sh.cells[`${r},${c}`] === undefined ? "" : sh.cells[`${r},${c}`],
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const v = sh.cells[`${r + i},${c + j}`]; return v === undefined ? "" : v; })),
        setFormula: v => { assert.ok(String(v).startsWith("="), "formula"); sh.cells[`${r},${c}`] = v; return rg; },
        setValues: v => { assert.strictEqual(v.length, nr, `${name} rows`); v.forEach(row => assert.strictEqual(row.length, nc, `${name} cols`)); return each((a, b, i, j) => { sh.cells[`${a},${b}`] = v[i][j]; }); },
        setFormulas: v => { assert.strictEqual(v.length, nr); return each((a, b, i, j) => { sh.cells[`${a},${b}`] = v[i][j]; }); },
        setNumberFormat: f => each((a, b) => { sh.formats[`${a},${b}`] = f; }),
        setFontSize: () => rg, setFontWeight: () => rg, setFontColor: () => rg, setFontStyle: () => rg, setBackground: () => rg,
        setWrap: () => rg, setVerticalAlignment: () => rg, setHorizontalAlignment: () => rg, setBorder: () => rg,
        setDataValidation: () => rg, createFilter: () => { sh.filters.push(R); return {}; },
        clearContent: () => each((a, b) => { delete sh.cells[`${a},${b}`]; }),
      };
      return rg;
    },
  };
  sh.api = api;
  return sh;
}
const sheets = [];
const ss = {
  getSheetByName: n => (sheets.find(s => s.name === n) || {}).api || null,
  getSheets: () => sheets.map(s => s.api),
  insertSheet: (n, i) => { assert.ok(!sheets.some(s => s.name === n), "duplicate sheet " + n); const s = makeSheet(ss, n); sheets.splice(i, 0, s); return s.api; },
  deleteSheet: a => { const i = sheets.findIndex(s => s.api === a); sheets.splice(i, 1); },
  active: null,
  setActiveSheet: a => { ss.active = a; return a; },
  getNumSheets: () => sheets.length,
  moveActiveSheet: pos => { const i = sheets.findIndex(s => s.api === ss.active); const [x] = sheets.splice(i, 1); sheets.splice(pos - 1, 0, x); },
};
["Claims", "Expense Ledger", "GL List"].forEach((n, i) => ss.insertSheet(n, i));
const builder = () => { const b = {}; ["whenFormulaSatisfied", "setFontColor", "setRanges", "setBold", "setBackground", "whenTextEqualTo",
  "setGradientMinpointWithValue", "setGradientMaxpointWithValue", "requireValueInList", "setAllowInvalid"].forEach(m => b[m] = () => b); b.build = () => ({}); return b; };
const logs = [];
const ctx = { console: { log: (m) => logs.push(m), warn() {} }, Math, Date, JSON, String, Object, Array, Number,
  Utilities: { formatDate: (d) => d.toISOString() }, Session: { getScriptTimeZone: () => "America/Vancouver" },
  Charts: { ChartType: { BAR: "BAR", COMBO: "COMBO", PIE: "PIE", COLUMN: "COLUMN", LINE: "LINE" } },
  SpreadsheetApp: { getActive: () => ss, flush() {}, newConditionalFormatRule: builder, newDataValidation: builder,
    InterpolationType: { NUMBER: "NUMBER" }, BorderStyle: { SOLID: "SOLID" } } };
vm.createContext(ctx);
vm.runInContext(data + "\n" + code + "\n__o.TRACKER_DATA = TRACKER_DATA; __o.trackerData_ = trackerData_;", Object.assign(ctx, { __o: ctx }));

ctx.buildBudgetTracker();
const names = sheets.map(s => s.name);
assert.deepStrictEqual(names, ["Dashboard", "Charts", "Claims", "Expense Ledger", "Revenue Ledger", "GL List", "GL Budget", "Line Items", "Tracker Ledger", "GL Map", "Dept Summary", "Tracker Notes"]);
assert.throws(() => ctx.buildBudgetTracker(), /already exists/);
const S0 = n => sheets.find(s => s.name === n);
S0("Expense Ledger").cells["1,1"] = "Entry ID"; S0("Expense Ledger").cells["73,1"] = "M-017";   // 72 expense rows
S0("Revenue Ledger").cells["2,1"] = "R-001"; S0("Revenue Ledger").cells["2,8"] = 1234;           // a revenue row
S0("Revenue Ledger").cells["3,4"] = "typed in a formula column";                                 // cleared by the next build
ctx.rebuildBudgetTracker();
assert.deepStrictEqual(sheets.map(s => s.name), names.concat(["Line Items (before rebuild)"]), "rebuild recreates the same tabs + a backup of Line Items");
assert.ok(logs.some(l => /Expense Ledger untouched: 72 expense rows/.test(l)), logs.join(" | "));
assert.strictEqual(S0("Expense Ledger").cells["73,1"], "M-017", "Expense Ledger not touched");
assert.strictEqual(S0("Revenue Ledger").cells["2,1"], "R-001", "Revenue Ledger rows not touched by a rebuild");
assert.strictEqual(S0("Revenue Ledger").cells["2,8"], 1234);
assert.strictEqual(S0("Revenue Ledger").cells["3,4"], undefined, "formula column cleared");
assert.ok(String(S0("Revenue Ledger").cells["2,4"]).startsWith("=ARRAYFORMULA"), "formulas put back on an existing tab");
assert.strictEqual(sheets.filter(s => s.name === "Revenue Ledger").length, 1, "only one Revenue Ledger");

// ---- edits on Line Items + GL Map survive a rebuild (unless the AGM changed that line)
{
  const D0 = ctx.trackerData_(), li0 = S0("Line Items"), mp0 = S0("GL Map");
  assert.strictEqual(li0.cells["4,16"], "AGM Budget (as built)");
  assert.strictEqual(mp0.cells["4,10"], "Budget Line (as built)");
  const rowOf = (pred) => { for (let r = 5; r < 5 + D0.li.length; r++) if (pred(r)) return r; throw new Error("row not found"); };
  const rA = rowOf(r => li0.cells[`${r},11`] > 0), rB = rowOf(r => r > rA && li0.cells[`${r},11`] > 0);
  const nameA = li0.cells[`${rA},9`], nameB = li0.cells[`${rB},9`], origB = li0.cells[`${rB},11`];
  li0.cells[`${rA},11`] = 1234.5;          // hand edit, AGM line unchanged -> kept
  li0.cells[`${rB},11`] = 99;              // hand edit, but the new AGM data changes this line -> AGM wins
  const liB = D0.li.find(l => l[0] === li0.cells[`${rB},1`]); liB[5] = origB + 10;
  const expRows = D0.gl.filter(g => g[3] === "Expenditure");
  const mi = D0.map.findIndex(m => m[0].endsWith("|")), other = expRows.find(g => g[10] !== D0.map[mi][1])[10];
  mp0.cells[`${5 + mi},2`] = other;        // GL code sent to another budget line -> kept
  ctx.rebuildBudgetTracker();
  const li1 = S0("Line Items"), mp1 = S0("GL Map");
  assert.notStrictEqual(li1, li0, "tab rebuilt");
  const r1A = rowOf.call(null, r => li1.cells[`${r},1`] === li0.cells[`${rA},1`]);
  assert.strictEqual(li1.cells[`${r1A},9`], nameA); assert.strictEqual(li1.cells[`${r1A},11`], 1234.5, "edit kept");
  const r1B = rowOf.call(null, r => li1.cells[`${r},1`] === li0.cells[`${rB},1`]);
  assert.strictEqual(li1.cells[`${r1B},9`], nameB); assert.strictEqual(li1.cells[`${r1B},11`], origB + 10, "AGM change wins");
  assert.strictEqual(li1.cells[`${r1B},16`], origB + 10);
  assert.strictEqual(mp1.cells[`${5 + mi},2`], other, "GL Map edit kept"); assert.strictEqual(mp1.cells[`${5 + mi},10`], D0.map[mi][1]);
  assert.ok(logs.some(l => /Your edits put back: 2 /.test(l)), logs.join(" | "));
  assert.ok(logs.some(l => /NOT put back.*the AGM changed it from/.test(l)));
  const bk = S0("Line Items (before rebuild)");
  assert.strictEqual(sheets.filter(s => s.name === "Line Items (before rebuild)").length, 1, "only one backup tab");
  assert.strictEqual(bk.cells[`${rA},11`], 1234.5, "backup holds the old values");
  assert.ok(/^Line Items before the rebuild/.test(bk.cells["1,1"]));
  liB[5] = origB; ctx.rebuildBudgetTracker();  // back to normal for the checks below (edit A is carried again)
}
// a Line Items tab built before this feature (no column P): nothing to compare, the backup tab is the safety net
delete S0("Line Items").cells["4,16"]; logs.length = 0;
ctx.rebuildBudgetTracker();
assert.ok(logs.some(l => /could not be checked/.test(l)) && logs.some(l => /Your edits put back: 1 - GL Map/.test(l)), logs.join(" | "));
assert.strictEqual(sheets[sheets.length - 1].name, "Line Items (before rebuild)", "backup is the last tab");

// ---- formula checks
const S = n => sheets.find(s => s.name === n);
const known = new Set(sheets.map(s => s.name));
let count = 0;
for (const s of sheets) for (const [k, v] of Object.entries(s.cells)) {
  if (typeof v !== "string" || !v.startsWith("=")) continue;
  count++;
  const where = `${s.name}!${k}: ${v.slice(0, 120)}`;
  let depth = 0, inStr = false;
  for (const ch of v) { if (ch === '"') inStr = !inStr; else if (!inStr) { if (ch === "(") depth++; if (ch === ")") depth--; assert.ok(depth >= 0, "paren " + where); } }
  assert.ok(!inStr, "unclosed quote " + where); assert.strictEqual(depth, 0, "parens " + where);
  for (const m of v.matchAll(/'([^']+)'!/g)) assert.ok(known.has(m[1]), "unknown sheet " + m[1] + " in " + where);
  for (const m of v.matchAll(/(?<![A-Za-z'])(Dashboard)!/g)) assert.ok(known.has(m[1]));
  assert.ok(!/undefined|NaN|\[object/.test(v), "bad text " + where);
}

// ---- layout checks
const gb = S("GL Budget"), li = S("Line Items"), led = S("Tracker Ledger"), dash = S("Dashboard"), mp = S("GL Map");
const D = ctx.trackerData_();
const gbLast = 4 + D.gl.length, liLast = 4 + D.li.length;
assert.strictEqual(gb.cells["4,20"], "Tracker Label (used in Transactions)");
assert.strictEqual(gb.cells[`${gbLast},1`], D.gl[D.gl.length - 1][0]);
assert.ok(String(gb.cells["5,8"]).includes(`'Line Items'!$K$5:$K$${liLast}`), "budget formula covers all line items");
assert.ok(String(gb.cells[`${5 + D.gl.findIndex(g => g[3] === "Expenditure")},9`]).includes("'Tracker Ledger'!$J$5:$J$3004"), "expense spent reads Tracker Ledger");
assert.strictEqual(gb.formats["5,7"], "@"); assert.strictEqual(li.formats["5,7"], "@");
assert.strictEqual(led.cells["4,5"], "Budget Line");
assert.ok(led.maxRows >= 3004, "ledger has room for 3000 rows");
for (let c = 1; c <= 20; c++) assert.ok(String(led.cells[`5,${c}`]).startsWith("=ARRAYFORMULA(IF(LEN('Expense Ledger'!A2:A3001"), "ledger col " + c);
assert.ok(led.cells["5,5"].includes('REGEXREPLACE(TO_TEXT(\'Expense Ledger\'!D2:D3001),"\\D","")'), "digits regex");
assert.strictEqual(dash.cells["23,3"], "All departments");
// top tiles: revenue / operating expenses / surplus / Capital Projects, like the AGM P&L; semester stays in K5
assert.deepStrictEqual([2,3,4,5,6,11].map(c => dash.cells[`4,${c}`]), ["Revenue budget", "Operating expense budget", "Budgeted surplus", "Capital Projects (from savings)", "Spent", "Current semester"]);
assert.ok(dash.cells["5,2"].includes('"Revenue"') && dash.cells["5,2"].includes('"<>Capital Projects"'));
assert.ok(dash.cells["5,5"].includes('"Expenditure"') && dash.cells["5,5"].includes(',"Capital Projects")'));
assert.strictEqual(dash.cells["5,8"], "=C5+E5-F5"); assert.strictEqual(dash.cells["5,11"], '=IF(MONTH(TODAY())<=4,"Spring","Fall")');
assert.ok(String(gb.cells["5,27"]).includes("Dashboard!$K$5"), "GL Budget period reads the semester tile");
{ const D1 = ctx.TRACKER_DATA, g = {}; D1.gl.forEach(x => g[x[0]] = x); const t = { r: 0, e: 0, c: 0 };
  D1.li.forEach(l => { const p = g[l[1]]; if (p[3] === "Revenue" && p[1] !== "Capital Projects") t.r += l[5]; if (p[3] === "Expenditure") { if (p[1] === "Capital Projects") t.c += l[5]; else t.e += l[5]; } });
  const pl = D1.agm.filter(a => a[0] !== "Capital Projects");
  assert.ok(Math.abs(t.r - pl.reduce((s, a) => s + a[2], 0)) < 0.005 && Math.abs(t.e - pl.reduce((s, a) => s + a[3], 0)) < 0.005, "tiles = AGM P&L revenue and expenses");
  console.log("Dashboard tiles: revenue", t.r.toFixed(2), "operating expenses", t.e.toFixed(2), "surplus", (t.r - t.e).toFixed(2), "Capital Projects", t.c.toFixed(2)); }
assert.strictEqual(JSON.stringify(dash.unprot), JSON.stringify(["Dashboard!R23C3"]));
assert.ok(dash.cells["25,2"].startsWith("=IFERROR(QUERY('GL Budget'!$A$5:$AI$" + gbLast), "ranked query");
assert.ok(dash.cells["25,2"].includes('" and B="""&$C$23&""""'), "department filter quoting");
assert.strictEqual(dash.cells["24,2"], "Department");
assert.strictEqual(mp.cells["4,2"], "Budget Line"); assert.strictEqual(mp.cells["4,7"], "Budget Line");
// charts: 5 charts, every range inside the chart data tables, headers on row 1 of each range
const ch = S("Charts");
assert.strictEqual(ch.charts.length, 5, "five charts");
assert.strictEqual(sheets.reduce((n, s) => n + s.charts.length, 0), 5, "charts only on the Charts tab (rebuild removes old ones)");
const nDept = D.depts.length;
assert.strictEqual(ch.cells["4,14"], "Department"); assert.strictEqual(ch.cells[`${4 + nDept},14`], D.depts[nDept - 1]);
assert.strictEqual(ch.cells["18,14"], "Month"); assert.strictEqual(ch.cells["19,14"], "May 2026"); assert.strictEqual(ch.cells["30,14"], "Apr 2027");
assert.strictEqual(ch.formats["19,14"], "@", "month labels stay text");
assert.ok(ch.cells["19,15"].includes('"2026-05"') && ch.cells["30,15"].includes('"2027-04"'), "month keys match Tracker Ledger's yyyy-mm");
assert.ok(ch.cells["19,15"].includes("'Tracker Ledger'!$T$5:$T$3004,\"Yes\""), "monthly spent counts matched rows only");
assert.strictEqual(ch.cells["34,14"], "Budget line");
assert.ok(ch.cells["35,14"].startsWith("=IFERROR(ARRAYFORMULA(SORTN(FILTER({'GL Budget'!$F$5:$F$" + gbLast), "top 10 formula");
for (const c of ch.charts) {
  assert.strictEqual(c.headers, 1, "chart headers");
  for (const rg of c.ranges) {
    const [, r, col] = rg.a1.match(/R(\d+)C(\d+)$/).map(Number);
    assert.ok(col >= 14 && col <= 19 && [4, 18, 34].includes(r), "chart range starts on a data table header: " + rg.a1);
  }
  assert.ok(c.pos[1] < 14, "charts sit left of the data tables");
}
assert.strictEqual(dash.cells["7,9"], "Graphs: see the Charts tab");
// ---- revenue ledger + lines outside the budget
{
  const rl = S("Revenue Ledger"), ds = S("Dept Summary");
  assert.deepStrictEqual([1, 2, 3, 8, 9, 12].map(c => rl.cells[`1,${c}`]), ["Entry ID", "Date", "GL Code", "Amount", "Status", "Match Key"]);
  assert.ok(rl.cells["2,12"].includes('"g"&REGEXREPLACE(TO_TEXT(C2:C1001),"\\D","")'), "revenue match key is text");
  assert.ok(rl.cells["2,4"].includes("'GL List'!$A:$B"), "GL name lookup");
  assert.ok(!Object.values(rl.cells).some(v => typeof v === "string" && /'(GL Budget|Dashboard|Tracker Ledger|GL Map)'!/.test(v)), "Revenue Ledger never points at a tracker tab (deleted on rebuild)");
  const revRow = D.gl.findIndex(g => g[3] === "Revenue" && g[5] === "Student Association Fees"), r = 5 + revRow;
  assert.strictEqual(gb.cells[`${r},36`], "g4107210341035100003300610000", "revenue match key on GL Budget");
  assert.strictEqual(gb.formats[`${r},36`], "@");
  assert.strictEqual(gb.cells[`${r},9`], `=IF($AJ${r}="",0,SUMIFS('Revenue Ledger'!$H$2:$H$1001,'Revenue Ledger'!$L$2:$L$1001,$AJ${r},'Revenue Ledger'!$I$2:$I$1001,"Received"))`);
  assert.ok(gb.cells[`${r},10`].includes('"Expected"') && gb.cells[`${r},13`].includes("Receiving"));
  const noGl = D.gl.findIndex(g => g[3] === "Revenue" && /No GL/.test(g[6]));
  assert.strictEqual(gb.cells[`${5 + noGl},36`], "", "revenue line without a GL code has no key");
  const expRow = D.gl.findIndex(g => g[3] === "Expenditure");
  assert.ok(gb.cells[`${5 + expRow},9`].includes("'Tracker Ledger'"), "expense lines unchanged");
  assert.strictEqual(gb.cells[`${5 + expRow},36`], "");
  assert.strictEqual(dash.cells["4,13"], "Revenue received (Revenue Ledger)");
  assert.ok(dash.cells["5,13"].includes('"Revenue"'));
  assert.strictEqual(dash.cells["4,11"], "Current semester", "semester tile stays in K");
  assert.strictEqual(ds.cells["4,12"], "Revenue Received");
  const t = 5 + D.agm.length;
  assert.ok(String(ds.cells[`${t + 1},12`]).includes(`"Received")-L${t}`), "received revenue not on a budget line");
  // the club-payment code, charged to Capital Projects, counts on "Loan repayment to the university"
  const loan = D.gl.find(g => g[1] === "Capital Projects" && g[5] === "Loan repayment to the university");
  assert.ok(loan, "loan repayment line");
  const mk = D.map.find(m => m[0] === "4107211156025100003907010000|Capital Projects");
  assert.ok(mk && mk[1] === loan[10] && mk[3] === "Capital Projects", "map key sends Charge to Department = Capital Projects to the loan line");
  assert.strictEqual(D.map.find(m => m[0] === "4107211156025100003907010000|")[1].startsWith("Proposal Committee"), true, "default for that code unchanged");
  assert.strictEqual(D.gl.length, ctx.TRACKER_DATA.gl.length, "no budget line added");
  assert.strictEqual(ctx.TRACKER_DATA.map.length + 1, D.map.length, "TRACKER_DATA itself is not changed");
  assert.strictEqual(mp.cells[`${4 + D.map.length},1`], mk[0], "extra key written on the GL Map tab");
  assert.ok(Object.values(S("Tracker Notes").cells).some(v => /Revenue Ledger tab/.test(v)));
}
// every GL Map budget line exists as a GL Budget label, and map keys are unique
const labels = new Set(D.gl.map(g => g[10]));
D.map.forEach(m => assert.ok(labels.has(m[1]), "map target missing: " + m[1]));
assert.strictEqual(new Set(D.map.map(m => m[0])).size, D.map.length, "unique map keys");
// data tie-out: budgets by department from line items == AGM figures
const byKey = {}; D.gl.forEach(g => byKey[g[0]] = g);
const tot = {}; D.li.forEach(l => { if (!l[6]) return; const g = byKey[l[1]]; const k = g[1] + "|" + g[3]; tot[k] = (tot[k] || 0) + l[5]; });
D.agm.forEach(a => {
  assert.ok(Math.abs((tot[a[0] + "|Revenue"] || 0) - a[2]) < 0.005, "revenue tie " + a[0]);
  assert.ok(Math.abs((tot[a[0] + "|Expenditure"] || 0) - a[3]) < 0.005, "expense tie " + a[0]);
});
console.log(`ALL TRACKER BUILD TESTS PASSED: ${count} formulas, ${D.gl.length} GL lines, ${D.li.length} line items, tie-out OK`);
