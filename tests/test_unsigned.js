// Runs the real "Unsigned Claims - Code.gs" against a simulated claim log (Claims tab), Drive and Mail.
// Checks: which claims are picked, order, links, attachments + size cap, the confirm dialog, the "nothing waiting" case,
// and that it fits with the other files of the bound project (no clashing names, menu item points at it).
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const P = require("./paths").file;  // file locations (one folder per Apps Script project)
const code = fs.readFileSync(P("UnsignedClaims.gs"), "utf8");
const manual = fs.readFileSync(P("ManualExpense.gs"), "utf8");
const v5 = fs.readFileSync(P("Code.gs"), "utf8");

const HEAD = ["ID", "Timestamp", "Status", "Payable to", "Name", "Email", "Total", "GST", "HST", "Folder", "PDF file id", "Token", "Decided by", "Decided at", "JSON file id", "Form"];
// The column order must match the claim web app's LOG_HEADERS.
const lh = v5.match(/const LOG_HEADERS = \[([\s\S]*?)\];/)[1].match(/'([^']*)'/g).map(s => s.slice(1, -1));
assert.deepStrictEqual(lh, HEAD, "Claims columns changed in Code.gs_V5");

const day = 86400000, now = Date.now();
const rows = [
  HEAD,
  ["a1", new Date(now - 10 * day), "Pending VP Finance signature", "Ann Lee", "Ann Lee", "a@x", 120.5, 0, 0, "https://drive/f1", "pdf1", "tokA", "", "", "j1", "expense"],
  ["b2", new Date(now - 3 * day), "Pending VP Finance signature", "Bo <b>", "Bo <b>", "b@x", 1234.56, 0, 0, "", "pdf2", "tokB", "", "", "j2", "mileage"],
  ["c3", new Date(now - 20 * day), "Pending President signature", "Cy", "Cy", "c@x", 50, 0, 0, "", "pdf3", "tokC", "", "", "", "expense"],
  ["d4", new Date(now - 30 * day), "Returned by President: fix receipt", "Di", "Di", "d@x", 75, 0, 0, "", "pdf4", "tokD", "", new Date(now - 2 * day), "", "expense"],
  ["e5", new Date(now - 5 * day), "Approved - signed PDF from NU Staff saved", "Ed", "Ed", "e@x", 10, 0, 0, "", "pdf5", "tokE", "", "", "", "expense"],
  [1234567, new Date(now - 40 * day), "Pending VP Finance final approval", "Fay", "Fay", "f@x", 300, 0, 0, "", "big", "tokF", "", new Date(now - 1 * day), "", "expense"],
  ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["g7", new Date(now - 1 * day), "Pending NU Staff signed PDF", "Gil", "Gil", "g@x", 20, 0, 0, "", "", "tokG", "", "", "", "expense"],
  ["h8", new Date(now - 1 * day), "Sent back to requester", "Hal", "Hal", "h@x", 20, 0, 0, "", "pdf8", "tokH", "", "", "", "expense"],
];
let claimsRows = rows;
const sent = [], alerts = [];
let answer = "YES";
const blob = (n) => ({ getBytes: () => ({ length: n }), name: "b" + n });
const ctx = {
  console: { log() {}, warn() {} },
  SpreadsheetApp: {
    getActive: () => ({ getSheetByName: (n) => n === "Claims" && claimsRows ? { getDataRange: () => ({ getValues: () => claimsRows.map(r => r.slice()) }) } : null }),
    getUi: () => ({
      ButtonSet: { OK: "OK", YES_NO: "YES_NO" }, Button: { YES: "YES", NO: "NO" },
      alert: (t, m, b) => { alerts.push({ t, m, b }); return b === "YES_NO" ? answer : "OK"; },
    }),
  },
  DriveApp: { getFileById: (id) => { if (id === "pdf4") throw new Error("gone"); return { getBlob: () => blob(id === "big" ? 19500000 : 400000) }; } },
  MailApp: { sendEmail: (o) => sent.push(o) },
  Session: { getScriptTimeZone: () => "America/Vancouver" },
  Utilities: { formatDate: (d, tz, f) => d.toISOString().slice(0, 10) },
};
vm.createContext(ctx);
vm.runInContext(code, ctx);

// 1. picks only claims waiting on VP Finance, signature ones first, oldest first
const J = (x) => JSON.parse(JSON.stringify(x)); // arrays from the vm context have another prototype
const list = ctx.ucWaitingClaims_();
assert.deepStrictEqual(J(list.map(c => c.id)), ["a1", "b2", "1234567", "d4", "g7"]);
assert.deepStrictEqual(J(list.map(c => c.sign)), [true, true, true, false, false]);
assert.strictEqual(list[2].days, 1, "waiting since = Decided at when present");
assert.strictEqual(list[0].days, 10, "else the submission time");
assert.strictEqual(list[1].mileage, true);

// 2. sends one email to VP Finance with links, escaping, attachments within the cap
ctx.emailMeUnsignedClaims();
assert.strictEqual(sent.length, 1);
const m = sent[0];
assert.strictEqual(m.to, "vp.finance@example.org");
assert.ok(/^Claims waiting for your signature: 3 \(\+2 other\)/.test(m.subject), m.subject);
const url = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
assert.ok(v5.includes("WEB_APP_URL: '" + url + "'"), "web app URL matches Code.gs_V5");
assert.ok(m.htmlBody.includes(url + "?action=decision&amp;id=a1&amp;token=tokA"));
assert.ok(m.htmlBody.includes(url + "?action=decision&amp;id=1234567&amp;token=tokF"));
assert.ok(m.body.includes(url + "?action=decision&id=b2&token=tokB"));
assert.ok(m.htmlBody.includes("Bo &lt;b&gt;") && !m.htmlBody.includes("Bo <b>"), "names are escaped");
assert.ok(m.htmlBody.includes("$1,234.56") && m.htmlBody.includes("$1,655.06"), "amounts + signature total");
assert.ok(m.htmlBody.includes("Kilometer reimbursement"));
assert.ok(m.htmlBody.includes("Also waiting on you"));
assert.ok(!m.htmlBody.includes("tokC") && !m.htmlBody.includes("tokE") && !m.htmlBody.includes("tokH"), "other stages left out");
// pdf1 + pdf2 (0.4 MB each) fit, big (19.5 MB) goes over the 20 MB cap, pdf4 is missing, g7 has no PDF
assert.strictEqual(m.attachments.length, 2);
assert.ok(/2 PDFs attached/.test(alerts[alerts.length - 1].m) && /3 PDFs only linked/.test(alerts[alerts.length - 1].m), alerts[alerts.length - 1].m);
assert.ok(alerts[0].m.includes("3 claims waiting for your signature and 2 other waiting on you"), alerts[0].m);

// 3. "No" in the confirm dialog sends nothing
answer = "NO"; ctx.emailMeUnsignedClaims();
assert.strictEqual(sent.length, 1);

// 4. nothing waiting -> a message, no email
claimsRows = [HEAD, rows[3], rows[5]];
answer = "YES"; ctx.emailMeUnsignedClaims();
assert.strictEqual(sent.length, 1);
assert.strictEqual(alerts[alerts.length - 1].t, "Nothing waiting on you");

// 5. no Claims tab -> clear error
claimsRows = null;
assert.throws(() => ctx.ucWaitingClaims_(), /No "Claims" tab/);

// 6. fits the bound project: menu item, no top-level name clashes with the other bound files
assert.ok(manual.includes(".addItem('Email me claims waiting for my signature', 'emailMeUnsignedClaims')"));
const top = (s) => new Set([...s.matchAll(/^(?:function\s+([A-Za-z_$][\w$]*)|(?:const|let|var)\s+([A-Za-z_$][\w$]*))/gm)].map(x => x[1] || x[2]));
const mine = top(code);
["ManualExpense.gs", "Tracker.gs", "TrackerData.gs", "ClaimTracking.gs"].forEach(f => {
  const other = top(fs.readFileSync(P(f), "utf8"));
  const clash = [...mine].filter(n => other.has(n));
  assert.deepStrictEqual(clash, [], "name clash with " + f);
});
console.log("test_unsigned: all checks passed");
