// Runs the real "President Reminders - Code.gs" against a simulated claim log (Claims tab), Mail and the menu dialog.
// Checks: which claims are picked, order, links, escaping, the optional note, Cancel, "nothing waiting", last-reminder
// date, and that it fits with the other files of the bound project (no clashing names, menu item points at it).
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const P = require("./paths").file;  // file locations (one folder per Apps Script project)
const code = fs.readFileSync(P("PresidentReminders.gs"), "utf8");
const manual = fs.readFileSync(P("ManualExpense.gs"), "utf8");
const v5 = fs.readFileSync(P("Code.gs"), "utf8");

const HEAD = ["ID", "Timestamp", "Status", "Payable to", "Name", "Email", "Total", "GST", "HST", "Folder", "PDF file id", "Token", "Decided by", "Decided at", "JSON file id", "Form"];
const lh = v5.match(/const LOG_HEADERS = \[([\s\S]*?)\];/)[1].match(/'([^']*)'/g).map(s => s.slice(1, -1));
assert.deepStrictEqual(lh, HEAD, "Claims columns changed in Code.gs_V5");
assert.ok(v5.includes("PRESIDENT_EMAIL: 'president@example.org'"), "President email matches Code.gs_V5");
const url = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
assert.ok(v5.includes("WEB_APP_URL: '" + url + "'"), "web app URL matches Code.gs_V5");

const day = 86400000, now = Date.now();
const rows = [
  HEAD,
  ["a1", new Date(now - 10 * day), "Pending VP Finance signature", "Ann", "Ann", "a@x", 120.5, 0, 0, "", "", "tokA", "", "", "", "expense"],
  ["b2", new Date(now - 9 * day), "Pending President signature", "Bo <b>", "Bo <b>", "b@x", 1234.56, 0, 0, "", "", "tokB", "", new Date(now - 3 * day), "", "mileage"],
  [1234567, new Date(now - 20 * day), "Pending President signature", "Cy", "Cy", "c@x", 50, 0, 0, "", "", "tokC", "", new Date(now - 8 * day), "", "expense"],
  ["d4", new Date(now - 30 * day), "Returned by President: fix receipt", "Di", "Di", "d@x", 75, 0, 0, "", "", "tokD", "", "", "", "expense"],
  ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["e5", new Date(now - 1 * day), "Sent back for changes", "Ed", "Ed", "e@x", 20, 0, 0, "", "", "tokE", "", "", "", "expense"],
];
let claimsRows = rows;
const sent = [], alerts = [], prompts = [], props = {};
let button = "OK", text = "";
const ctx = {
  console: { log() {}, warn() {} },
  SpreadsheetApp: {
    getActive: () => ({ getSheetByName: (n) => n === "Claims" && claimsRows ? { getDataRange: () => ({ getValues: () => claimsRows.map(r => r.slice()) }) } : null }),
    getUi: () => ({
      ButtonSet: { OK: "OK", OK_CANCEL: "OK_CANCEL" }, Button: { OK: "OK", CANCEL: "CANCEL" },
      alert: (t, m, b) => { alerts.push({ t, m, b }); return "OK"; },
      prompt: (t, m, b) => { prompts.push({ t, m, b }); return { getSelectedButton: () => button, getResponseText: () => text }; },
    }),
  },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  MailApp: { sendEmail: (o) => sent.push(o) },
  Session: { getScriptTimeZone: () => "America/Vancouver" },
  Utilities: { formatDate: (d, tz, f) => d.toISOString().slice(0, 10) },
};
vm.createContext(ctx);
vm.runInContext(code, ctx);
const J = (x) => JSON.parse(JSON.stringify(x));

// 1. picks only "Pending President signature", longest waiting first (since Decided at, else submitted)
const list = ctx.prWaitingClaims_();
assert.deepStrictEqual(J(list.map(c => c.id)), ["1234567", "b2"]);
assert.deepStrictEqual(J(list.map(c => c.days)), [8, 3]);
assert.strictEqual(list[1].mileage, true);

// 2. Cancel sends nothing and records nothing
button = "CANCEL"; ctx.remindPresidentClaims();
assert.strictEqual(sent.length, 0);
assert.ok(prompts[0].m.includes("2 claims waiting") && prompts[0].m.includes("$1,284.56") && prompts[0].m.includes("never"), prompts[0].m);
assert.ok(!props.PR_LAST_REMINDER);

// 3. OK sends one email to the President, VP in copy, with links, escaping, note, total
button = "OK"; text = "  Please sign <these> before Friday.  ";
ctx.remindPresidentClaims();
assert.strictEqual(sent.length, 1);
const m = sent[0];
assert.strictEqual(m.to, "president@example.org");
assert.strictEqual(m.cc, "vp.finance@example.org");
assert.strictEqual(m.replyTo, "vp.finance@example.org");
assert.strictEqual(m.subject, "Reminder: 2 NUSA claims waiting for your signature");
assert.ok(m.htmlBody.includes(url + "?action=decision&amp;id=1234567&amp;token=tokC"));
assert.ok(m.htmlBody.includes(url + "?action=decision&amp;id=b2&amp;token=tokB"));
assert.ok(m.body.includes(url + "?action=decision&id=b2&token=tokB"));
assert.ok(m.htmlBody.includes("Bo &lt;b&gt;") && !m.htmlBody.includes("Bo <b>"), "names escaped");
assert.ok(m.htmlBody.includes("Please sign &lt;these&gt; before Friday.") && m.body.startsWith("Please sign <these> before Friday."), "note included, trimmed");
assert.ok(m.htmlBody.includes("$1,284.56") && m.htmlBody.includes("Kilometer reimbursement"));
["tokA", "tokD", "tokE"].forEach(t => assert.ok(!m.htmlBody.includes(t), "other stages left out: " + t));
assert.ok(props.PR_LAST_REMINDER, "last reminder remembered");
assert.strictEqual(alerts[alerts.length - 1].t, "Sent");

// 4. next time the dialog shows the last reminder date
button = "CANCEL"; ctx.remindPresidentClaims();
assert.ok(!prompts[prompts.length - 1].m.includes("never"));

// 5. no note -> no note block
button = "OK"; text = ""; ctx.remindPresidentClaims();
assert.ok(!sent[1].htmlBody.includes("border-left:3px"));
assert.ok(sent[1].body.startsWith("Claims waiting"));

// 6. nothing waiting -> a message, no email
claimsRows = [HEAD, rows[1], rows[4]];
ctx.remindPresidentClaims();
assert.strictEqual(sent.length, 2);
assert.strictEqual(alerts[alerts.length - 1].t, "Nothing to remind");

// 7. no Claims tab -> clear error
claimsRows = null;
assert.throws(() => ctx.prWaitingClaims_(), /No "Claims" tab/);

// 8. fits the bound project: menu items, no top-level name clashes with the other bound files
assert.ok(manual.includes(".addItem('Remind the President to sign claims', 'remindPresidentClaims')"));
assert.ok(manual.includes(".addItem('Remind people with sent-back claims', 'remindSentBackClaims')"));
const top = (s) => new Set([...s.matchAll(/^(?:function\s+([A-Za-z_$][\w$]*)|(?:const|let|var)\s+([A-Za-z_$][\w$]*))/gm)].map(x => x[1] || x[2]));
const mine = top(code);
["ManualExpense.gs", "Tracker.gs", "TrackerData.gs", "ClaimTracking.gs",
  "UnsignedClaims.gs", "SentBackReminders.gs", "GlCodes.gs"].forEach(f => {
  const other = top(fs.readFileSync(P(f), "utf8"));
  const clash = [...mine].filter(n => other.has(n));
  assert.deepStrictEqual(clash, [], "name clash with " + f);
});
console.log("test_president: all checks passed");
