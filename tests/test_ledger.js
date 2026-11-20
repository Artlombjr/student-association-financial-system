// Simulates the Google Sheets pieces the ledger code touches (V4 = V3 ledger + owner-only setup), then runs the real functions.
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const code = fs.readFileSync(require("./paths").file("Code.gs"), "utf8");

function makeSheet(name, id, ss) {
  const sh = { name, id, rows: [], frozen: 0,
    getName: () => name, getSheetId: () => id, getParent: () => ss, setFrozenRows(n) { this.frozen = n; },
    clearContents() { this.rows = []; },
    appendRow(r) { this.rows.push(r.slice()); },
    getDataRange() { const rows = this.rows; return { getValues: () => rows.map(r => r.slice()) }; },
    getRange(r, c, nr = 1, nc = 1) {
      const self = this;
      const ensure = () => { while (self.rows.length < r + nr - 1) self.rows.push([]); };
      const api = {
        setValues(v) { assert.ok(v.length === nr && v.every(x => x.length === nc), `setValues: data ${v.length}x${v[0].length} but range ${nr}x${nc}`); ensure(); for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) self.rows[r - 1 + i][c - 1 + j] = v[i][j]; return api; },
        setValue(v) { ensure(); self.rows[r - 1][c - 1] = v; return api; },
        setNumberFormat() { return api; }, setDataValidation() { return api; }, setFontWeight() { return api; },
      };
      return api;
    } };
  return sh;
}
const ss = { sheets: {}, getId: () => "SSID",
  getSheetByName(n) { return this.sheets[n] || null; },
  insertSheet(n) { this.sheets[n] = makeSheet(n, Object.keys(this.sheets).length * 111, ss); return this.sheets[n]; } };
const log = ss.insertSheet("Claims");
log.rows.push(["ID","Timestamp","Status","Payable to","Name","Email","Total","GST","HST","Folder","PDF file id","Token","Decided by","Decided at","JSON file id","Form"]);
const payloads = {};
const dv = () => { const b = { requireValueInList: () => b, requireValueInRange: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; };
const props = {};
let uuidN = 0;
const pad = n => String(n).padStart(2, "0");
const ctx = { console, Math, Date, JSON, String, Object, Array, parseFloat, SpreadsheetApp: { newDataValidation: dv },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
  Utilities: { getUuid: () => "1234abcd-0000-4000-8000-" + String(++uuidN).padStart(12, "0"),
    formatDate: (d, tz, f) => f.replace("yyyy", d.getFullYear()).replace("MM", pad(d.getMonth() + 1)).replace("dd", pad(d.getDate())).replace("HH", pad(d.getHours())).replace("mm", pad(d.getMinutes())) },
  Session: { getScriptTimeZone: () => "America/Vancouver", getActiveUser: () => ({ getEmail: () => "owner@x" }), getEffectiveUser: () => ({ getEmail: () => "owner@x" }) },
  ContentService: { MimeType: { TEXT: "text", CSV: "csv" }, createTextOutput: s => ({ body: s, setMimeType(m) { this.mime = m; return this; } }) },
  ScriptApp: { getService: () => ({ getUrl: () => "https://dev" }) } };
vm.createContext(ctx);
vm.runInContext(code + `
getLogSheet_ = function () { return __log; };
loadPayload_ = function (id) { return JSON.parse(JSON.stringify(__payloads[id])); };
__out.LC = LC; __out.GL = GL;`, Object.assign(ctx, { __log: log, __payloads: payloads, __out: ctx }));

const L = ctx.LC;
const led = () => ss.getSheetByName("Expense Ledger").rows.slice(1);
const byId = id => led().find(r => r[L["Entry ID"]] === id);

// Claim A: expense claim with 2 receipts, VP assigns GL codes and signs.
log.rows.push(["aaaa1111", new Date(), "Pending VP Finance signature", "Student", "Ana Test", "a@x", 150, 5, 0, "f", "p", "t", "", "", "JA", "expense-claim"]);
payloads.JA = { form: "expense-claim", lines: [
  { date: "2026-09-18", desc: "Pizza for nursing event", amount: "100", gst: "5", hst: "0", net: 96.65, gl: "410-72104-56025100003902710000", glName: "Events - Nursing Representative" },
  { date: "2027-01-10", desc: "Decorations for spring mixer", amount: "50", gst: "0", hst: "0", net: 50, gl: "410-72111-56025100001000010000", glName: "MixEx/EventHosting" } ] };
ctx.writeLedgerRows_({ id: "aaaa1111", name: "Ana Test" }, ctx.loadPayload_("JA"), "Pending President signature");
assert.strictEqual(led().length, 2);
assert.strictEqual(byId("aaaa1111-1")[L["Status"]], "Pending");
assert.strictEqual(byId("aaaa1111-1")[L["Amount (incl. tax)"]], 100);
assert.strictEqual(byId("aaaa1111-1")[L["GL Group"]], "Faculty");
assert.strictEqual(byId("aaaa1111-2")[L["Date"]].getMonth(), 0);

// Status moves on through the real updateClaimStatus_ -> rows follow.
ctx.updateClaimStatus_("aaaa1111", "Pending VP Finance final approval", "pres");
assert.strictEqual(byId("aaaa1111-2")[L["Status"]], "Pending");
// User edits: locks line 1 with a corrected amount, charges line 2 to Proposal Committee, adds a manual purchase.
const sh = ss.getSheetByName("Expense Ledger");
let r1 = sh.rows.findIndex(r => r[L["Entry ID"]] === "aaaa1111-1");
sh.rows[r1][L["Amount (incl. tax)"]] = 90; sh.rows[r1][L["Locked"]] = "Yes"; sh.rows[r1][L["Status"]] = "Posted";
let r2 = sh.rows.findIndex(r => r[L["Entry ID"]] === "aaaa1111-2");
sh.rows[r2][L["Charge to Department (optional)"]] = "Proposal Committee"; sh.rows[r2][L["Notes"]] = "PC event";
sh.rows.push(["M-001", "Manual", new Date(2026, 8, 20), "410-72109-55612100003902110000", "Food Costs", "", "", "Warehouse Store run", "Warehouse Store", 250, 0, 0, 250, "Approved", "", "", "", "", "", ""]);
// Claim Tracking columns (U, V, W, added by the bound "Claim Tracking" script after Notes) on locked + unlocked rows.
sh.rows[0].push("Submitted", "Cheque Received from AP", "Cheque Picked Up");
sh.rows[r1].push("Submitted to Accounts Payable", "Yes", "No");
sh.rows[r2].push("Received by NUSA", "No", "No");
const tracked = () => [r1, r2].map(i => sh.rows[i].slice(20, 23).join("|")).join(" / ");
const trackedBefore = tracked();
ctx.updateClaimStatus_("aaaa1111", "Approved & sent to AP", "vp");
assert.strictEqual(byId("aaaa1111-1")[L["Status"]], "Posted", "locked row keeps its status");
assert.strictEqual(byId("aaaa1111-1")[L["Amount (incl. tax)"]], 90, "locked row keeps its amount");
assert.strictEqual(byId("aaaa1111-1")[L["Claim Status"]], "Approved & sent to AP", "claim status still refreshed");
assert.strictEqual(byId("aaaa1111-2")[L["Status"]], "Approved");
// Re-writing the claim (e.g. VP re-assigns GL) keeps Charge to Department + Notes on unlocked rows.
ctx.writeLedgerRows_({ id: "aaaa1111", name: "Ana Test" }, ctx.loadPayload_("JA"), "Approved & sent to AP");
assert.strictEqual(byId("aaaa1111-2")[L["Charge to Department (optional)"]], "Proposal Committee");
assert.strictEqual(byId("aaaa1111-2")[L["Notes"]], "PC event");
assert.strictEqual(byId("aaaa1111-1")[L["Amount (incl. tax)"]], 90);
assert.strictEqual(byId("M-001")[L["Amount (incl. tax)"]], 250, "manual row untouched");
assert.strictEqual(led().length, 3);

// Claim B: mileage, already approved before V3 -> back-fill creates it. Claim C still waiting on VP -> skipped.
log.rows.push(["bbbb2222", new Date(), "Approved & sent to AP", "Student", "Ben Drive", "b@x", 25, 0, 0, "f", "p", "t", "", "", "JB", "mileage"]);
payloads.JB = { form: "mileage", lines: [{ date: "2026-09-05", km: "40", destination: "Northfield", purpose: "Supplies", parking: "1" }] };
log.rows.push(["cccc3333", new Date(), "Pending VP Finance signature", "Student", "Cy", "c@x", 10, 0, 0, "f", "p", "t", "", "", "JC", "expense-claim"]);
payloads.JC = { form: "expense-claim", lines: [{ date: "2026-09-20", desc: "x y z", amount: "10", gst: "0", hst: "0" }] };
const url = ctx.setupExpenseLedger();
const m = byId("bbbb2222-1");
assert.strictEqual(m[L["Amount (incl. tax)"]], 25, "40 km x $0.60 + $1 parking");
assert.strictEqual(m[L["GL Code"]], "410-72103-56302100001000010000", "km claims go to Travel-Mileage");
assert.strictEqual(m[L["GL Name"]], "Travel-Mileage (Personal Vehicle)");
assert.strictEqual(m[L["GL Group"]], "Student Lounge (The Commons)");
// A km claim signed earlier keeps the code it was signed with.
log.rows.push(["kkkk9999", new Date(), "Approved & sent to AP", "Student", "Old Drive", "d@x", 6, 0, 0, "f", "p", "t", "", "", "JK", "mileage"]);
payloads.JK = { form: "mileage", lines: [{ date: "2026-09-05", km: "10", destination: "Warehouse Store", purpose: "Supplies", parking: "0", gl: "410-72109-55600100001000010000", glName: "Office - Miscellaneous (Mileage)" }] };
ctx.setupExpenseLedger();
assert.strictEqual(byId("kkkk9999-1")[L["GL Code"]], "410-72109-55600100001000010000", "old km claim keeps its code");
assert.strictEqual(byId("kkkk9999-1")[L["GL Name"]], "Office - Miscellaneous (Mileage)");
assert.strictEqual(byId("kkkk9999-1")[L["GL Group"]], "Student Lounge (The Commons)");
assert.strictEqual(m[L["Status"]], "Approved");
assert.ok(!byId("cccc3333-1"), "claim still waiting on VP signature is not in the ledger");
assert.strictEqual(byId("aaaa1111-1")[L["Amount (incl. tax)"]], 90, "back-fill respects the lock");
assert.strictEqual(tracked(), trackedBefore, "claim system never touches the Claim Tracking columns");
// Sent back at final approval -> unlocked rows become Void.
ctx.updateClaimStatus_("bbbb2222", "Sent back for changes", "vp");
assert.strictEqual(byId("bbbb2222-1")[L["Status"]], "Void");
// Status change for a claim with GL codes but no rows yet -> rows are created on the fly.
log.rows.push(["dddd4444", new Date(), "Pending President signature", "Student", "Di", "d@x", 20, 0, 0, "f", "p", "t", "", "", "JD", "expense-claim"]);
payloads.JD = { form: "expense-claim", lines: [{ date: "2026-09-21", desc: "Tape and markers", amount: "20", gst: "1", hst: "0", gl: "410-72105-55600100001000010000" }] };
ctx.updateClaimStatus_("dddd4444", "Pending VP Finance final approval", "pres");
assert.strictEqual(byId("dddd4444-1")[L["Net Expense"]], 19.33);
// GL list includes the 4 added codes
assert.ok(ctx.GL.some(g => g.items.some(i => i[1] === "410-72109-55612100003902110000")));
// Mileage trip with a blank purpose -> no stray separator (blank purpose test)
log.rows.push(["eeee5555", new Date(), "Pending President signature", "Student", "Ed", "e@x", 6, 0, 0, "f", "p", "t", "", "", "JE", "mileage"]);
payloads.JE = { form: "mileage", lines: [{ date: "2026-09-10", km: "10", destination: "Riverside", purpose: "", parking: "" }] };
ctx.updateClaimStatus_("eeee5555", "Pending VP Finance final approval", "pres");
assert.strictEqual(byId("eeee5555-1")[L["Description"]], "Mileage: Riverside (10 km)");
assert.strictEqual(byId("eeee5555-1")[L["Amount (incl. tax)"]], 6);
// ---- private feed
const feedUrl = ctx.ledgerFeedUrl_();
assert.ok(feedUrl.startsWith("https://script.google.com/macros/s/") && feedUrl.includes("/exec?feed=ledger&key="), feedUrl);
const key = feedUrl.split("key=")[1];
assert.ok(key.length >= 60, "long key");
const bad = ctx.doGet({ parameter: { feed: "ledger", key: "nope" } });
assert.ok(bad.body.startsWith("Forbidden"), "wrong key refused");
assert.ok(ctx.doGet({ parameter: { feed: "ledger" } }).body.startsWith("Forbidden"), "missing key refused");
const good = ctx.doGet({ parameter: { feed: "ledger", key } });
assert.strictEqual(good.mime, "csv");
const lines = good.body.split(String.fromCharCode(13, 10));
assert.strictEqual(lines[0].split(",")[0], '"Entry ID"');
assert.strictEqual(lines.length, led().length + 1);
assert.ok(!good.body.includes("@x"), "no emails in feed");
assert.ok(!/Token/.test(good.body), "no tokens in feed");
assert.ok(good.body.includes('"aaaa1111-1","Claim",2026-09-18,'), "date as yyyy-MM-dd, text quoted");
require("fs").writeFileSync(require("os").tmpdir() + "/feed_sample.csv", good.body);
const url2 = ctx.rotateLedgerFeedKey();
assert.notStrictEqual(url2, feedUrl);
assert.ok(ctx.doGet({ parameter: { feed: "ledger", key } }).body.startsWith("Forbidden"), "old key dead after rotation");
assert.strictEqual(ctx.doGet({ parameter: { feed: "ledger", key: url2.split("key=")[1] } }).mime, "csv");
console.log("ALL LEDGER TESTS PASSED\n" + url + "\nrows:", led().length);
