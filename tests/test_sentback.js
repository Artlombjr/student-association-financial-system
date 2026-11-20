// Runs the real "Sent Back Reminders - Code.gs" against a simulated claim log (Claims tab), Drive, Mail and document
// properties. Checks: which claims are listed, the dialog (rows, disabled rows, its script parses), the reminder emails
// (recipient, copy, note, message, edit link = the one the web app accepts), skipping, "last reminder", and that it fits
// with the other files of the bound project (menu item, no clashing names).
const fs = require("fs"), vm = require("vm"), assert = require("assert");
const P = require("./paths").file;  // file locations (one folder per Apps Script project)
const code = fs.readFileSync(P("SentBackReminders.gs"), "utf8");
const manual = fs.readFileSync(P("ManualExpense.gs"), "utf8");
const v5 = fs.readFileSync(P("Code.gs"), "utf8");

const HEAD = ["ID", "Timestamp", "Status", "Payable to", "Name", "Email", "Total", "GST", "HST", "Folder", "PDF file id", "Token", "Decided by", "Decided at", "JSON file id", "Form"];
const lh = v5.match(/const LOG_HEADERS = \[([\s\S]*?)\];/)[1].match(/'([^']*)'/g).map(s => s.slice(1, -1));
assert.deepStrictEqual(lh, HEAD, "Claims columns changed in Code.gs_V5");

const day = 86400000, now = Date.now();
const rows = [
  HEAD,
  ["a1", new Date(now - 10 * day), "Sent back for changes", "Ann Lee", "Ann Lee", "ann@x.ca", 120.5, 0, 0, "", "pdf1", "tokA", "VP", new Date(now - 4 * day), "j1", "expense-claim"],
  ["b2", new Date(now - 30 * day), "Sent back for changes", "Bo <b>", "Bo <b>", "bo@x.ca", 1234.56, 0, 0, "", "pdf2", "tokB", "VP", new Date(now - 20 * day), "j2", "mileage"],
  ["c3", new Date(now - 5 * day), "Pending VP Finance signature", "Cy", "Cy", "cy@x.ca", 50, 0, 0, "", "pdf3", "tokC", "", "", "j3", "expense-claim"],
  [1234567, new Date(now - 9 * day), "Sent back for changes", "Di", "Di", "", 75, 0, 0, "", "pdf4", "tokD", "VP", new Date(now - 2 * day), "j4", "expense-claim"],
  ["e5", new Date(now - 60 * day), "Sent back to requester", "Ed", "Ed", "ed@x.ca", 10, 0, 0, "", "pdf5", "tokE", "VP", new Date(now - 50 * day), "", "expense-claim"],
  ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
];
let claimsRows = rows;
const json = { j1: { sentBackNote: "Receipt 2 is missing.\nPlease add it." }, j2: { sentBackNote: "Add the parking receipt <now>" } };
const sent = [], alerts = [], dialogs = [], props = {};
const ctx = {
  console: { log() {}, warn() {} },
  SpreadsheetApp: {
    getActive: () => ({ getSheetByName: (n) => n === "Claims" && claimsRows ? { getDataRange: () => ({ getValues: () => claimsRows.map(r => r.slice()) }) } : null }),
    getUi: () => ({
      ButtonSet: { OK: "OK" },
      alert: (t, m, b) => { alerts.push({ t, m }); return "OK"; },
      showModalDialog: (h, t) => dialogs.push({ html: h.html, t }),
    }),
  },
  HtmlService: { createHtmlOutput: (html) => { const o = { html, setWidth: () => o, setHeight: () => o }; return o; } },
  DriveApp: { getFileById: (id) => { if (!json[id]) throw new Error("gone"); return { getBlob: () => ({ getDataAsString: () => JSON.stringify(json[id]) }) }; } },
  MailApp: { sendEmail: (o) => sent.push(o) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  Session: { getScriptTimeZone: () => "America/Vancouver" },
  Utilities: { formatDate: (d, tz, f) => d.toISOString().slice(0, 10) },
};
vm.createContext(ctx);
vm.runInContext(code, ctx);
const J = (x) => JSON.parse(JSON.stringify(x));

// 1. lists every "Sent back..." claim, sent back longest ago first, with the note and why one can't be reminded
const list = ctx.sbrSentBackClaims_(true);
assert.deepStrictEqual(J(list.map(c => c.id)), ["e5", "b2", "a1", "1234567"]);
assert.deepStrictEqual(J(list.map(c => c.canEdit)), [false, true, true, false]);
assert.ok(/older claim/.test(list[0].why) && /no email/.test(list[3].why));
assert.strictEqual(list[2].note, "Receipt 2 is missing.\nPlease add it.");
assert.strictEqual(list[2].days, 4, "days since Decided at");
assert.strictEqual(list[1].mileage, true);

// 2. menu item opens the dialog; rows escaped; claims that can't be reminded are disabled; its script parses
ctx.remindSentBackClaims();
assert.strictEqual(dialogs.length, 1);
const h = dialogs[0].html;
assert.ok(h.includes('value="a1" checked') && h.includes('value="b2" checked'));
assert.ok(h.includes('value="e5" disabled') && h.includes('value="1234567" disabled'));
assert.ok(h.includes("Bo &lt;b&gt;") && !h.includes("Bo <b>"), "names escaped");
assert.ok(h.includes("Add the parking receipt &lt;now&gt;"), "notes escaped");
assert.ok(!h.includes("tokA"), "tokens are not put in the dialog");
new Function(h.match(/<script>([\s\S]*)<\/script>/)[1]); // client script is valid JavaScript
assert.ok(h.includes(".sbrSendReminders(ids,"));

// 3. sending: one email per ticked claim, to the requester, VP Finance in copy, note + message + edit link
const url = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
assert.ok(v5.includes("WEB_APP_URL: '" + url + "'"), "web app URL matches Code.gs_V5");
// the web app's edit link format and the payload key the note is saved under
assert.ok(v5.includes("webAppUrl_() + '?action=edit&id=' + encodeURIComponent(id) + '&token=' + encodeURIComponent(token)"));
assert.ok(v5.includes("p.sentBackNote = note.trim()"));
assert.ok(/function editRow_[\s\S]*?row\.token !== token[\s\S]*?isSentBack_\(row\.status\)/.test(v5), "edit page checks token + Sent back status");
const res = ctx.sbrSendReminders(["a1", "b2", "e5", "zz"], "  Please do it by Friday.  ");
assert.deepStrictEqual(J(res.sent), ["Ann Lee", "Bo <b>"]);
assert.strictEqual(res.skipped.length, 2);
assert.ok(/Ed \(older claim/.test(res.skipped[0]) && /zz \(not "Sent back/.test(res.skipped[1]), res.skipped.join(" | "));
assert.strictEqual(sent.length, 2);
const [ma, mb] = sent;
assert.strictEqual(ma.to, "ann@x.ca"); assert.strictEqual(ma.cc, "vp.finance@example.org"); assert.strictEqual(ma.replyTo, "vp.finance@example.org");
assert.strictEqual(ma.subject, "Reminder: your NUSA expense claim still needs a change");
assert.strictEqual(mb.subject, "Reminder: your NUSA kilometer reimbursement still needs a change");
assert.ok(ma.htmlBody.includes(url + "?action=edit&amp;id=a1&amp;token=tokA"));
assert.ok(ma.body.includes(url + "?action=edit&id=a1&token=tokA"));
assert.ok(ma.htmlBody.includes("Receipt 2 is missing.<br>Please add it."));
assert.ok(ma.htmlBody.includes("Please do it by Friday.") && !ma.htmlBody.includes("  Please"), "message trimmed + included");
assert.ok(ma.htmlBody.includes("Hi Ann,") && ma.htmlBody.includes("$120.50"));
assert.ok(mb.htmlBody.includes("Hi Bo,") && mb.htmlBody.includes("$1,234.56"));
assert.ok(mb.htmlBody.includes("Add the parking receipt &lt;now&gt;"));
assert.ok(props["SBR_LAST_a1"] && props["SBR_LAST_b2"] && !props["SBR_LAST_e5"], "last reminder recorded");

// 4. the dialog shows the last reminder; the token is not rotated (claims tab untouched)
ctx.remindSentBackClaims();
const rowA1 = dialogs[1].html.split('value="a1"')[1].split("</tr>")[0];
assert.ok(rowA1.includes(props["SBR_LAST_a1"].slice(0, 10)) && !rowA1.includes("Never"), "a1 shows its last reminder date");
assert.strictEqual(J(ctx.sbrSentBackClaims_(false).find(c => c.id === "a1")).last, props["SBR_LAST_a1"]);
assert.strictEqual(claimsRows[1][11], "tokA");

// 5. a claim with no note says the note is in the first email; empty message adds nothing
json.j1 = {};
sent.length = 0;
ctx.sbrSendReminders(["a1"], "");
assert.ok(sent[0].htmlBody.includes("What needs fixing is in our earlier email."));
assert.throws(() => ctx.sbrSendReminders([], ""), /Tick at least one/);

// 6. nothing sent back -> a message, no dialog; no Claims tab -> clear error
claimsRows = [HEAD, rows[3]];
ctx.remindSentBackClaims();
assert.strictEqual(alerts[alerts.length - 1].t, "Nothing to remind");
assert.strictEqual(dialogs.length, 2);
claimsRows = null;
assert.throws(() => ctx.sbrSentBackClaims_(false), /No "Claims" tab/);

// 7. fits the bound project: menu item, no top-level name clashes with the other bound files
assert.ok(manual.includes(".addItem('Remind people with sent-back claims', 'remindSentBackClaims')"));
const top = (s) => new Set([...s.matchAll(/^(?:function\s+([A-Za-z_$][\w$]*)|(?:const|let|var)\s+([A-Za-z_$][\w$]*))/gm)].map(x => x[1] || x[2]));
const mine = top(code);
["ManualExpense.gs", "Tracker.gs", "TrackerData.gs", "ClaimTracking.gs",
 "UnsignedClaims.gs", "GlCodes.gs"].forEach(f => {
  const other = top(fs.readFileSync(P(f), "utf8"));
  const clash = [...mine].filter(n => other.has(n));
  assert.deepStrictEqual(clash, [], "name clash with " + f);
});
console.log("test_sentback: all checks passed");
