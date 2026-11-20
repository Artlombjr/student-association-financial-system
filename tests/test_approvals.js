// V5 approval chain: requester -> VP Finance (link) -> President (link) -> NU Staff gets the PDF
// by email, signs it and replies -> signed PDF saved in the claim folder + emailed to VP Finance -> AP.
// Runs the real Code.gs_V5 functions against simulated Sheets/Drive/Mail/Gmail and checks every step,
// the return/resend paths, claims routed before V5 (NU Staff sign link, VP final approval),
// reminders, and the signature block in both PDFs. (test_approvals_v4.js keeps testing V4.)
// Optional: node test_approvals.js <outDir>  writes sample packet HTML files for each stage.
const fs = require("fs"), vm = require("vm"), assert = require("assert"), path = require("path");
const SRC = require("./paths").file("Code.gs");
const outDir = process.argv[2] || null;

function makeSheet(name, ss) {
  return { name, rows: [], getParent: () => ss, setFrozenRows() {}, clearContents() { this.rows = []; },
    appendRow(r) { this.rows.push(r.slice()); }, deleteRow(i) { this.rows.splice(i - 1, 1); },
    getDataRange() { const rows = this.rows; return { getValues: () => rows.map(r => r.slice()) }; },
    getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); },
    getLastRow() { return this.rows.length; }, getMaxRows() { return Math.max(1000, this.rows.length); },
    getRange(r, c, nr = 1, nc = 1) {
      const self = this, ensure = () => { while (self.rows.length < r + nr - 1) self.rows.push([]); };
      const api = {
        setValues(v) { ensure(); for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) self.rows[r - 1 + i][c - 1 + j] = v[i][j]; return api; },
        setValue(v) { ensure(); self.rows[r - 1][c - 1] = v; return api; }, clearContent() { return api; },
        setNumberFormat() { return api; }, setDataValidation() { return api; }, setFontWeight() { return api; },
        getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = self.rows[r - 1 + i] || []; out.push(Array.from({ length: nc }, (_, j) => row[c - 1 + j])); } return out; },
      };
      return api;
    } };
}

function setup(staffEmail, opts) {
  opts = opts || {};
  const props = {};
  const pad = n => String(n).padStart(2, "0");
  const ss = { sheets: {}, getSheetByName(n) { return this.sheets[n] || null; }, insertSheet(n) { return (this.sheets[n] = makeSheet(n, ss)); } };
  const log = ss.insertSheet("Claims");
  const files = {}, mail = [], fetches = [], triggers = [];
  let fid = 0, uuid = 0;
  const folder = { viewers: [], addViewer(e) { this.viewers.push(e); }, getUrl: () => "https://drive/folder", createFile(blob) { return newFile(blob.html, blob.name, blob.type); }, moveTo() {},
    getFiles() { const list = Object.values(files).filter(f => !f.trashed); let i = 0; return { hasNext: () => i < list.length, next: () => list[i++] }; } };
  function newFile(content, name, mime) {
    const id = "F" + (++fid);
    files[id] = { id, content, name, mime: mime || "", viewers: [], trashed: false,
      getId: () => id, getName() { return this.name; }, getMimeType() { return this.mime; },
      getBlob() { const f = this; return { getDataAsString: () => f.content, getBytes: () => Array.isArray(f.content) ? f.content : [], getContentType: () => f.mime, setName(n) { return this; } }; },
      setContent(s) { this.content = s; }, setName(n) { this.name = n; return this; }, setTrashed(t) { this.trashed = t; }, addViewer(e) { this.viewers.push(e); },
      getOwner: () => ({ getEmail: () => opts.owner || "vp.finance@example.org" }),
      getParents: () => { let done = false; return { hasNext: () => !done, next: () => { done = true; return folder; } }; } };
    return files[id];
  }
  const dv = () => { const b = { requireValueInList: () => b, requireValueInRange: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; };
  // Simulated Gmail of the script's account: threads with messages {from, subject, pdf}.
  const threads = [], labels = {}, searches = [];
  function addThread(msgs) {
    const t = { labels: [], msgs, addLabel(l) { this.labels.push(l.name); },
      getMessages: () => msgs.map(m => ({ getFrom: () => m.from, getSubject: () => m.subject, getPlainBody: () => m.body || "",
        getAttachments: () => (m.files || (m.pdf === false ? [] : [{ name: "Signed.pdf", type: "application/pdf", content: m.content || "SIGNED-PDF" }])).map(a => ({
          getName: () => a.name, getContentType: () => a.type,
          copyBlob: () => ({ html: a.content, name: a.name, type: a.type, setName(n) { this.name = n; return this; } }) })) })) };
    threads.push(t);
    return t;
  }
  const ctx = { console: { log() {}, warn() {} }, Math, Date, JSON, String, Object, Array, parseFloat, encodeURIComponent,
    SpreadsheetApp: { newDataValidation: dv, flush() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; }, getKeys: () => Object.keys(props) }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true }) },
    Utilities: { getUuid: () => "0000abcd-0000-4000-8000-" + String(++uuid).padStart(12, "0"),
      formatDate: (d, tz, f) => f.replace("yyyy", d.getFullYear()).replace("yy", String(d.getFullYear()).slice(2)).replace("MM", pad(d.getMonth() + 1))
        .replace("dd", pad(d.getDate())).replace("HH", pad(d.getHours())).replace("mm", pad(d.getMinutes())),
      base64Decode: b64 => Array.from(Buffer.from(b64, "base64")), base64Encode: bytes => Buffer.from(bytes).toString("base64"), sleep() {},
      newBlob: (html, type, name) => ({ html, type, name, getBytes: () => [], getContentType: () => type, getAs() { return this; }, setName(n) { this.name = n; return this; } }) },
    Session: { getScriptTimeZone: () => "America/Vancouver",
      getActiveUser: () => { if (opts.noScope) throw new Error("no permission"); return { getEmail: () => opts.activeUser || "" }; },
      getEffectiveUser: () => ({ getEmail: () => opts.me || "vp.finance@example.org" }) },
    DriveApp: { getFileById: id => { if (!files[id]) throw new Error("no file " + id); return files[id]; } },
    MailApp: { sendEmail: m => { if ((opts.badAddress && m.to === opts.badAddress) || (opts.failTo && opts.failTo.has(m.to))) throw new Error("Invalid email: " + m.to); mail.push(m); } },
    GmailApp: { search: q => { searches.push(q); return threads.filter(t => !t.labels.includes("NUSA-signed-PDF")); },
      getUserLabelByName: n => labels[n] || null, createLabel: n => (labels[n] = { name: n }) },
    ScriptApp: { getService: () => ({ getUrl: () => "https://dev" }), getProjectTriggers: () => triggers.slice(), deleteTrigger(t) { triggers.splice(triggers.indexOf(t), 1); },
      getOAuthToken: () => "TOKEN", newTrigger: fn => { const b = { timeBased: () => b, everyMinutes: n => { b.every = n; return b; }, everyDays: () => b, atHour: () => b, inTimezone: () => b,
        create: () => { const t = { getHandlerFunction: () => fn, every: b.every }; triggers.push(t); return t; } }; return b; } },
    UrlFetchApp: { fetch: (url, o) => { fetches.push({ url, auth: o && o.headers && o.headers.Authorization }); return (opts.fetch || (() => ({ getResponseCode: () => 404, getContentText: () => "{}" })))(url, files); } } };
  vm.createContext(ctx);
  const src = fs.readFileSync(SRC, "utf8");
  assert.match(src, /STAFF_EMAIL: 'Sam\.Patel@university\.example\.edu'/, "configured NU Staff email");
  assert.match(src, /VP_FINANCE_NU_EMAIL: 'vpfinance@university\.example\.edu'/, "configured VP Finance NU email");
  assert.match(src, /PRESIDENT_EMAIL: 'president@example\.org'/, "configured President email");
  assert.match(src, /VP_FINANCE_EMAIL: 'vp\.finance@example\.org'/, "configured VP Finance email");
  const code = src.replace(/STAFF_EMAIL: '[^']*'/, "STAFF_EMAIL: " + JSON.stringify(staffEmail));
  vm.runInContext(code + "\ngetLogSheet_ = function () { return __log; }; __out.LC = LC; __out.LOG_HEADERS = LOG_HEADERS; __out.CONFIG = CONFIG;",
    Object.assign(ctx, { __log: log, __out: ctx }));
  log.appendRow(ctx.LOG_HEADERS);
  // Run the signed-PDF check now (clears the 4-minute throttle first).
  const check = () => { delete props.LAST_SIGNED_PDF_CHECK; return ctx.checkStaffSignedPdfs(); };
  return { ctx, log, files, mail, newFile, props, fetches, folder, addThread, threads, searches, triggers, check };
}

const SIG = n => "data:image/png;base64,SIG" + n;
function addClaim(env, id, form, extra, logTime) {
  const payload = Object.assign({ form, first: "Ana", last: "Test", email: "ana@student.university.example.edu", payTo: "Student",
    printName: "Ana Test", signature: SIG("req"), totals: { amt: 100, gst: 5, hst: 0, net: 96.65, km: 50, parking: 5, kmReimb: 30 },
    lines: form === "mileage"
      ? [{ date: "2026-09-18", km: 50, destination: "Northfield - Vancouver", purpose: "Conference", parking: 5 }]
      : [{ date: "2026-09-18", amount: 100, gst: 5, hst: 0, net: 96.65, desc: "Pizza for nursing event" }] }, extra || {});
  const pdf = env.newFile("<old pdf>", "packet.pdf"), json = env.newFile(JSON.stringify(payload), "claim.json");
  env.log.appendRow([id, logTime || new Date(), "Pending VP Finance signature", "Student", "Ana Test", "ana@student.university.example.edu", 100, 5, 0,
    "https://drive/folder", pdf.id, "tok-" + id, "", "", json.id, form]);
}
const row = (env, id) => env.ctx.findClaim_(id);
const payloadOf = (env, id) => JSON.parse(env.files[row(env, id).jsonId].content);
const pdfHtml = (env, id) => env.files[row(env, id).pdfId].content;
const save = (name, html) => { if (outDir) fs.writeFileSync(path.join(outDir, name), html); };
const STAFF = "Sam.Patel@university.example.edu", VPF = "vp.finance@example.org", VPNU = "vpfinance@university.example.edu", ME = "owner@example.org";
// A claim routed to NU Staff before V5 (sign link): same data, older status.
const legacyStaff = (env, id) => { env.log.rows.find(x => env.ctx.sameId_(x[0], id))[2] = "Pending NU Staff signature"; };
// VP Finance forwards NU Staff's reply from their NU Outlook (subject kept, "FW: RE:" added).
const reply = (env, id, from, extra) => env.addThread([Object.assign({ from: from || "Alex Rivera <" + VPNU + ">", subject: "FW: RE: " + env.ctx.staffPdfSubject_(row(env, id)) }, extra || {})]);

// ---------- 1. V5 happy path: VP Finance relays the PDF to NU Staff by NU email and forwards the signed reply ----------
{
  const env = setup(STAFF, { me: ME, owner: ME }), c = env.ctx, id = "exp00001";
  addClaim(env, id, "expense-claim");
  save("expense_0_submitted.html", c.buildPacketHtml_(payloadOf(env, id), id, []));
  assert.strictEqual(c.stageOf_(row(env, id).status), "vp1");
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  assert.strictEqual(row(env, id).status, "Pending President signature");
  assert.match(env.mail.at(-1).to, /president@example\.org/);
  assert.match(env.mail.at(-1).htmlBody, /action=decision/, "the President still gets a sign link");
  save("expense_1_vp_signed.html", pdfHtml(env, id));

  const before = env.mail.length;
  c.presidentSign(id, row(env, id).token, SIG("pres"));
  let r = row(env, id);
  assert.strictEqual(r.status, "Pending NU Staff signed PDF");
  assert.strictEqual(c.stageOf_(r.status), "staffpdf");
  const sent = env.mail.slice(before);
  assert.strictEqual(sent.length, 1, "one hand-off email, to VP Finance only");
  assert.ok(sent.every(m => !/university\.example\.edu/i.test(m.to + (m.cc || ""))), "the system never emails a university.example.edu address");
  const toVp = sent[0];
  assert.strictEqual(toVp.to, VPF);
  assert.strictEqual(toVp.subject, "[Expense Claim] Ana Test - $100.00 - send to Sam Patel to sign (Claim exp00001)");
  assert.strictEqual(c.staffPdfSubject_(r), "Expense Claim - Ana Test - $100.00 - please sign (Claim exp00001)");
  assert.ok(toVp.htmlBody.includes(">Expense Claim - Ana Test - $100.00 - please sign (Claim exp00001)</div>"), "subject to copy for Sam");
  assert.match(toVp.htmlBody, /write a new email to <b>Sam\.Patel@university\.example\.edu<\/b>/);
  assert.match(toVp.htmlBody, /In Outlook \(vpfinance@university\.example\.edu\)/);
  assert.match(toVp.htmlBody, /Hi Sam,<br><br>Attached is a NUSA expense claim/, "message to copy for Sam");
  assert.match(toVp.htmlBody, /Director of Student Engagement/);
  assert.match(toVp.htmlBody, /forward the reply from Outlook to <b>owner@example\.org<\/b>/);
  assert.match(toVp.htmlBody, /Keep <b>\(Claim exp00001\)<\/b> in the subject/);
  assert.match(toVp.htmlBody, /action=decision/, "VP keeps a link to re-send or send back");
  assert.strictEqual(toVp.attachments[0].getDataAsString(), pdfHtml(env, id), "the current packet (VP + President signatures) is attached");
  assert.deepStrictEqual(env.files[r.pdfId].viewers, [], "no Drive sharing with the university.example.edu address");
  save("expense_2_president_signed.html", pdfHtml(env, id));
  // Old links can't act; the VP's page offers re-send / send back only
  assert.throws(() => c.presidentSign(id, r.token, SIG("x")), /already been completed/);
  assert.throws(() => c.staffSign(id, r.token, SIG("x")), /already been completed/, "no link signing in V5");
  assert.throws(() => c.finalApprove(id, r.token), /already been completed/);
  const waitPage = c.decisionPageHtml_(r, "staffpdf", payloadOf(env, id));
  assert.match(waitPage, /Waiting for NU Staff&#39;s signed PDF/);
  assert.match(waitPage, /Email me the PDF and message again/);
  assert.match(waitPage, /doResendStaff\(\)/); assert.match(waitPage, /doSendBack\(\)/);
  assert.ok(!/doSign\(\)"|doFinal\(\)"/.test(waitPage));

  // Sam replies with the signed PDF
  const unsignedId = r.pdfId, t = reply(env, id), before2 = env.mail.length;
  assert.strictEqual(env.check(), 1);
  r = row(env, id);
  assert.strictEqual(r.status, "Approved - signed PDF from NU Staff saved");
  assert.strictEqual(c.stageOf_(r.status), "done");
  assert.strictEqual(env.files[r.pdfId].content, "SIGNED-PDF", "log points at the signed PDF");
  assert.strictEqual(env.files[r.pdfId].name, "packet - signed by NU Staff.pdf");
  assert.ok(!env.files[unsignedId].trashed, "the unsigned packet stays in the folder");
  assert.deepStrictEqual(t.labels, ["NUSA-signed-PDF"], "thread labelled as handled");
  assert.match(env.searches[0], /-label:NUSA-signed-PDF/);
  const p = payloadOf(env, id);
  assert.strictEqual(p.staffName, "Sam Patel"); assert.strictEqual(p.staffSignedFrom, VPNU); assert.strictEqual(p.staffSignedPdfId, r.pdfId);
  const after = env.mail.slice(before2);
  assert.strictEqual(after.length, 2, "VP Finance + requester");
  const fin = after.find(m => m.to === VPF);
  assert.match(fin.subject, /FINAL, ready for Accounts Payable/);
  assert.match(fin.htmlBody, /Sam Patel \(NU Staff\) signed the PDF and sent it back/);
  assert.match(fin.htmlBody, /packet - signed by NU Staff\.pdf/);
  assert.match(fin.htmlBody, /forward this email \(with its attachments\) to NU Accounts Payable/);
  assert.strictEqual(fin.attachments[0].getDataAsString(), "SIGNED-PDF", "signed PDF attached for VP Finance");
  assert.ok(after.every(m => !/accountspayable/i.test(m.to + (m.cc || ""))), "system never emails Accounts Payable");
  const toStudent = after.find(m => m.to === "ana@student.university.example.edu");
  assert.match(toStudent.htmlBody, /VP Finance, President and NU Staff/);
  assert.strictEqual(toStudent.attachments[0].getDataAsString(), "SIGNED-PDF");
  assert.strictEqual(c.getLedgerSheet_().rows[1][c.LC["Status"]], "Approved", "ledger counts it as spent");
  assert.strictEqual(c.reminderRecipient_(c.stageOf_(r.status)), null, "no reminders once approved");
  // Nothing happens twice
  const n = env.mail.length;
  assert.strictEqual(env.check(), 0);
  assert.strictEqual(env.mail.length, n);
}

// ---------- 1b. Only real replies count; strays are reported once; email failure retries ----------
{
  const env = setup(STAFF, { me: ME, owner: ME }), c = env.ctx;
  ["a0000001", "a0000002", "a0000003"].forEach(id => {
    addClaim(env, id, "expense-claim");
    c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
    c.presidentSign(id, row(env, id).token, SIG("pres"));
  });
  const our = env.addThread([{ from: "NUSA Finance <" + ME + ">", subject: c.staffPdfSubject_(row(env, "a0000001")) }]); // our own email (unsigned PDF)
  const stranger = reply(env, "a0000001", "someone@gmail.com");
  const noPdf = reply(env, "a0000001", STAFF, { pdf: false });
  const noId = env.addThread([{ from: VPNU, subject: "FW: signed", body: "here you go" }]);
  const mails = env.mail.length;
  assert.strictEqual(env.check(), 0);
  assert.strictEqual(row(env, "a0000001").status, "Pending NU Staff signed PDF", "none of those approve the claim");
  assert.strictEqual(env.mail.length, mails, "and nothing is emailed");
  [our, stranger, noPdf, noId].forEach(t => assert.deepStrictEqual(t.labels, []));
  // Subject retyped: the tag in the forwarded message text is enough
  reply(env, "a0000001", undefined, { subject: "FW: signed claim", body: "Hi Alex, signed.\n> Thank you,\n> (Claim a0000001)" });
  assert.strictEqual(env.check(), 1);
  assert.strictEqual(row(env, "a0000001").status, "Approved - signed PDF from NU Staff saved");
  // A second PDF for an approved claim: not saved, VP told once
  const late = reply(env, "a0000001");
  const m2 = env.mail.length;
  assert.strictEqual(env.check(), 0);
  assert.strictEqual(env.mail.length - m2, 1);
  assert.match(env.mail.at(-1).subject, /Signed PDF for claim a0000001 was not saved/);
  assert.match(env.mail.at(-1).htmlBody, /not waiting for NU Staff/);
  assert.deepStrictEqual(late.labels, ["NUSA-signed-PDF"]);
  assert.strictEqual(env.check(), 0); assert.strictEqual(env.mail.length - m2, 1, "reported only once");
  // Unknown claim ID
  env.addThread([{ from: STAFF, subject: "Re: [Expense Claim] X - please sign the PDF and reply (Claim zzzz9999)" }]);
  env.check();
  assert.match(env.mail.at(-1).htmlBody, /no claim with this ID was found/);
  // VP Finance email fails -> claim not marked, file removed, retried next run
  env.ctx.MailApp.sendEmail = (orig => m => { if (m.to === VPF) throw new Error("quota"); orig(m); })(env.ctx.MailApp.sendEmail);
  const t2 = reply(env, "a0000002");
  assert.strictEqual(env.check(), 0);
  assert.strictEqual(row(env, "a0000002").status, "Pending NU Staff signed PDF");
  assert.deepStrictEqual(t2.labels, [], "left for the next run");
  assert.ok(Object.values(env.files).filter(f => /signed by NU Staff/.test(f.name) && !f.trashed).length === 1, "failed copy removed (only a0000001's is kept)");
  env.ctx.MailApp.sendEmail = m => env.mail.push(m);
  assert.strictEqual(env.check(), 1);
  assert.strictEqual(row(env, "a0000002").status, "Approved - signed PDF from NU Staff saved");
  // Throttle: a web visitor can't hammer it
  assert.strictEqual(c.checkStaffSignedPdfs(), 0);
  // Other accepted senders: VP Finance's Gmail, Sam himself, and extra addresses in STAFF_REPLY_FROM
  ["a0000003"].forEach(id => { reply(env, id, "Alex <" + VPF + ">"); assert.strictEqual(env.check(), 1); });
  const env2 = setup(STAFF, { me: ME });
  env2.ctx.CONFIG.STAFF_REPLY_FROM = "sam.patel@example.org";
  addClaim(env2, "b0000001", "expense-claim");
  env2.ctx.vpSignApprove("b0000001", row(env2, "b0000001").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  env2.ctx.presidentSign("b0000001", row(env2, "b0000001").token, SIG("pres"));
  reply(env2, "b0000001", "Sam Patel <Sam.Patel@example.org>");
  assert.strictEqual(env2.check(), 1);
  assert.match(env2.searches[0], /from:\(sam\.patel@university\.example\.edu OR vpfinance@university\.example\.edu OR sam\.patel@example\.org OR vp\.finance@example\.org\)/);
  // Watcher trigger: owner only, every 10 minutes, not stacked
  const own = setup(STAFF, { me: ME, activeUser: ME });
  own.ctx.installSignedPdfWatcher(); own.ctx.installSignedPdfWatcher();
  assert.strictEqual(own.triggers.length, 1); assert.strictEqual(own.triggers[0].every, 10);
  assert.strictEqual(own.triggers[0].getHandlerFunction(), "checkStaffSignedPdfs");
  own.ctx.removeSignedPdfWatcher(); assert.strictEqual(own.triggers.length, 0);
}

// ---------- 1c. A claim routed before V5 (NU Staff sign link) still finishes the old way ----------
{
  const env = setup(STAFF), c = env.ctx, id = "old00009";
  addClaim(env, id, "expense-claim");
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c.presidentSign(id, row(env, id).token, SIG("pres"));
  legacyStaff(env, id);
  let r = row(env, id);
  assert.strictEqual(c.stageOf_(r.status), "staff");
  assert.match(c.decisionPageHtml_(r, "staff", payloadOf(env, id)), /Open the Drive folder \(all receipts\)/, "sign page links the folder");

  // Staff decision page
  const page = c.decisionPageHtml_(r, "staff", payloadOf(env, id));
  assert.match(page, /Sign &amp; approve &mdash; NU Staff/);
  assert.match(page, /run\.staffSign\(ID,TK,sig\)/);
  assert.match(page, /rr\.staffReturn\(ID,TK,n\)/);

  c.staffSign(id, r.token, SIG("staff"));
  r = row(env, id);
  assert.strictEqual(r.status, "Pending VP Finance final approval");
  assert.strictEqual(env.mail.at(-1).to, "vp.finance@example.org");
  assert.match(env.mail.at(-1).htmlBody, /has all signatures/);
  const p = payloadOf(env, id);
  assert.strictEqual(p.staffName, "Sam Patel");
  const html = pdfHtml(env, id);
  save("expense_3_staff_signed.html", html);
  ["Alex Rivera", "Jordan Blake", "Sam Patel", "Ana Test"].forEach(n => assert.ok(html.includes(n), n + " printed"));
  const order = ["Requester", "NUSA Vice President of Finance", "NUSA President", "Director of Student Engagement"].map(s => html.indexOf(">" + s + "</td>"));
  assert.ok(order.every((v, i) => v > 0 && (i === 0 || v > order[i - 1])), "signature rows in order");
  ["SIGreq", "SIGvp", "SIGpres", "SIGstaff"].forEach(s => assert.ok(html.includes(s), s + " image on PDF"));
  assert.match(c.decisionPageHtml_(r, "vpfinal", p), /All signatures are in/);

  const finalPage = c.decisionPageHtml_(r, "vpfinal", p);
  assert.match(finalPage, /Final approval &amp; email me the final PDF/);
  assert.ok(!/send to Accounts Payable/i.test(finalPage), "no 'send to AP' wording on the final page");
  const before = env.mail.length;
  c.finalApprove(id, r.token);
  assert.strictEqual(row(env, id).status, "Approved - final PDF sent to VP Finance");
  const sent = env.mail.slice(before);
  // Only two emails: the final packet to VP Finance, and the requester's copy. Nothing to any university.example.edu AP address.
  assert.strictEqual(sent.length, 2);
  assert.ok(sent.every(m => !/accountspayable/i.test(m.to + (m.cc || ""))), "system never emails Accounts Payable");
  const toVp = sent.find(m => m.to === "vp.finance@example.org");
  assert.match(toVp.subject, /FINAL, ready for Accounts Payable/);
  assert.match(toVp.htmlBody, /forward this email \(with its attachments\) to NU Accounts Payable/);
  assert.match(toVp.htmlBody, /Sam Patel \(NU Staff\)/);
  assert.ok(toVp.attachments.length >= 1, "final PDF attached for VP Finance");
  const toStudent = sent.find(m => m.to === "ana@student.university.example.edu");
  assert.match(toStudent.htmlBody, /VP Finance, President and NU Staff/);
  assert.match(toStudent.htmlBody, /NUSA Finance will now submit it to NU Accounts Payable/);
  // Ledger: pending through the chain, approved at the end
  const L = c.LC, led = c.getLedgerSheet_().rows.slice(1);
  assert.strictEqual(led.length, 1);
  assert.strictEqual(led[0][L["Status"]], "Approved");
}

// ---------- 2. Mileage routed before V5: NU Staff returns it by link, VP re-sends (now as a PDF), staff replies ----------
{
  const env = setup(STAFF, { me: ME }), c = env.ctx, id = "mil00001";
  addClaim(env, id, "mileage");
  c.vpSignApprove(id, row(env, id).token, SIG("vp"));
  c.presidentSign(id, row(env, id).token, SIG("pres"));
  assert.strictEqual(c.ledgerStatus_(row(env, id).status), "Pending");
  const ledRow = () => c.getLedgerSheet_().rows[1];
  assert.strictEqual(ledRow()[c.LC["Status"]], "Pending");
  assert.strictEqual(ledRow()[c.LC["Claim Status"]], "Pending NU Staff signed PDF");
  legacyStaff(env, id);

  assert.throws(() => c.staffReturn(id, row(env, id).token, "  "), /add a note/);
  c.staffReturn(id, row(env, id).token, "Please attach the parking receipt.");
  let r = row(env, id);
  assert.strictEqual(r.status, "Returned by NU Staff");
  assert.strictEqual(c.stageOf_(r.status), "staffreturned");
  assert.match(env.mail.at(-1).subject, /returned by NU Staff/);
  const page = c.decisionPageHtml_(r, "staffreturned", payloadOf(env, id));
  assert.match(page, /NU Staff note:<\/b><br>Please attach the parking receipt\./);
  assert.match(page, /Re-send to NU Staff/);
  assert.strictEqual(c.reminderRecipient_("staffreturned"), "vp.finance@example.org");

  c.resendToStaff(id, r.token);
  assert.strictEqual(row(env, id).status, "Pending NU Staff signed PDF", "re-sent the V5 way");
  assert.strictEqual(env.mail.at(-1).to, VPF);
  assert.match(env.mail.at(-1).subject, /\[Kilometer Reimbursement\].*\(Claim mil00001\)$/);
  assert.strictEqual(c.reminderRecipient_("staffpdf"), VPF);
  const html = pdfHtml(env, id);
  save("mileage_2_sent_to_staff.html", html);
  assert.match(html, /KILOMETER REIMBURSEMENT FORM/);
  ["SIGreq", "SIGvp", "SIGpres", ">Director of Student Engagement</td>", ">NUSA Vice President of Finance</td>", ">NUSA President</td>"].forEach(s => assert.ok(html.includes(s), s));
  assert.ok(!html.includes("SIGstaff"), "NU Staff line left blank for the hand signature");
  assert.ok(!html.includes("Please print name"), "no print-name label");
  // the resend can be repeated while waiting (e.g. Sam lost the email)
  c.resendToStaff(id, row(env, id).token);
  reply(env, id);
  assert.strictEqual(env.check(), 1);
  assert.strictEqual(row(env, id).status, "Approved - signed PDF from NU Staff saved");
  assert.strictEqual(c.getLedgerSheet_().rows[1][c.LC["Status"]], "Approved");
  assert.strictEqual(c.ledgerStatus_("Approved & sent to AP"), "Approved", "older approved claims still count");
  assert.strictEqual(c.reminderRecipient_(c.stageOf_("Approved - final PDF sent to VP Finance")), null, "no reminders once approved");
}

// ---------- 3. VP can send back to the student after a return (was blocked in V3) ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx, id = "exp00002";
  addClaim(env, id, "expense-claim");
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c.presidentSign(id, row(env, id).token, SIG("pres"));
  legacyStaff(env, id);
  c.staffReturn(id, row(env, id).token, "Wrong GL");
  c.sendBackClaim(id, row(env, id).token, "Please fix the receipt");
  assert.strictEqual(row(env, id).status, "Sent back for changes");
  assert.strictEqual(c.getLedgerSheet_().rows[1][c.LC["Status"]], "Void");
  // V5: VP can also send it back while waiting for NU Staff's PDF; a late PDF is then not saved
  const id2 = "exp00012";
  addClaim(env, id2, "expense-claim");
  c.vpSignApprove(id2, row(env, id2).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c.presidentSign(id2, row(env, id2).token, SIG("pres"));
  c.sendBackClaim(id2, row(env, id2).token, "Missing receipt");
  assert.strictEqual(row(env, id2).status, "Sent back for changes");
  reply(env, id2, STAFF);
  env.check();
  assert.strictEqual(row(env, id2).status, "Sent back for changes");
}

// ---------- 4. A claim already at final approval before V4 (no NU Staff signature) ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx, id = "old00001";
  addClaim(env, id, "expense-claim", { vpSignature: SIG("vp"), vpName: "Alex Rivera", presidentSignature: SIG("pres"), presidentName: "Jordan Blake" });
  env.log.rows[1][2] = "Pending VP Finance final approval";
  let r = row(env, id);
  assert.throws(() => c.finalApprove(id, r.token), /still needs the NU Staff signature/);
  const page = c.decisionPageHtml_(r, "vpfinal", payloadOf(env, id));
  assert.match(page, /NU Staff signature needed/);
  assert.ok(!/doFinal\(\)"/.test(page), "no final-approve button yet");
  c.resendToStaff(id, r.token);
  assert.strictEqual(row(env, id).status, "Pending NU Staff signed PDF");
  // once the signed PDF is saved, resend is no longer allowed
  reply(env, id);
  assert.strictEqual(env.check(), 1);
  assert.throws(() => c.resendToStaff(id, row(env, id).token), /already been completed/);
}

// ---------- 5. Staff email not filled in: President's signature is refused cleanly, nothing saved ----------
{
  const env = setup("STAFF_EMAIL_HERE"), c = env.ctx, id = "exp00003";
  addClaim(env, id, "expense-claim");
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  assert.throws(() => c.presidentSign(id, row(env, id).token, SIG("pres")), /NU Staff email is not set up/);
  assert.strictEqual(row(env, id).status, "Pending President signature");
  assert.ok(!payloadOf(env, id).presidentSignature);
}

// ================= Audit fixes =================

// ---------- 6. A tampered signature can't inject content into the PDF ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx;
  const evil = 'x" onerror="alert(1)"><div>Approved by: Alex Rivera</div><img src="y';
  const base = { form: "expense-claim", first: "A", last: "B", email: "a@b.ca", payTo: "Employee", address: "1 St", city: "Northfield",
    printName: "A B", lines: [{ amount: 5, fileData: "data:image/png;base64,AAAA", desc: "pens for the office" }] };
  assert.ok(c.validatePayload(Object.assign({}, base, { signature: evil })).some(m => /signature/.test(m)), "bad signature rejected");
  assert.strictEqual(c.validatePayload(Object.assign({}, base, { signature: "data:image/png;base64,iVBORw0KGgo=" })).length, 0, "real canvas PNG accepted");
  const html = c.signaturesHtml_({ printName: "A B", signature: evil, vpSignature: evil });
  assert.ok(!html.includes("onerror") && !html.includes("<div>Approved"), "PDF drops non-image signatures");
  assert.ok(!c.buildPacketHtml_({ lines: [] }, "id", [{ name: "r", dataUrl: evil }]).includes("onerror"), "receipt images sanitized");
  // an approver's sign page can't post junk either
  addClaim(env, "sig00001", "expense-claim");
  assert.throws(() => c.vpSignApprove("sig00001", "tok-sig00001", evil, ["410-72104-56025100003902710000"]), /could not be read/);
  assert.strictEqual(row(env, "sig00001").status, "Pending VP Finance signature");
}

// ---------- 7. The PDF keeps the submission date through every signature ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx, id = "date0001";
  addClaim(env, id, "expense-claim", { submittedAt: new Date(2026, 8, 3, 14, 5).toISOString() });
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c.presidentSign(id, row(env, id).token, SIG("pres"));
  legacyStaff(env, id);
  c.staffSign(id, row(env, id).token, SIG("staff"));
  const html = pdfHtml(env, id);
  assert.match(html, /Date \(mm\/dd\/yy\): <b>09\/03\/26<\/b>/, "form date = submission date");
  assert.match(html, /Submitted 2026-09-03 14:05/);
  // A claim submitted before V4 (no submittedAt) takes the log's Timestamp at its next re-signing.
  const env2 = setup("claims.office@example.org"), id2 = "date0002";
  addClaim(env2, id2, "mileage", {}, new Date(2026, 8, 10, 9, 30));
  env2.ctx.vpSignApprove(id2, row(env2, id2).token, SIG("vp"));
  assert.match(pdfHtml(env2, id2), /Date \(mm\/dd\/yy\): <b>09\/10\/26<\/b>/);
  assert.ok(payloadOf(env2, id2).submittedAt, "saved for later rebuilds");
  // submitClaim-style new claims get submittedAt stamped (checked via the helper fallback)
  assert.ok(c.submittedDate_({}) instanceof Date);
}

// ---------- 8. Final approval: requester email failure doesn't leave the claim pending ----------
{
  const env = setup("claims.office@example.org", { badAddress: "ana@student.university.example.edu" }), c = env.ctx, id = "fail0001";
  addClaim(env, id, "expense-claim");
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c.presidentSign(id, row(env, id).token, SIG("pres"));
  legacyStaff(env, id);
  c.staffSign(id, row(env, id).token, SIG("staff"));
  const tok = row(env, id).token, before = env.mail.length;
  const res = c.finalApprove(id, tok);
  assert.strictEqual(res.requesterEmailFailed, true);
  assert.strictEqual(row(env, id).status, "Approved - final PDF sent to VP Finance");
  assert.strictEqual(env.mail.length - before, 1, "VP Finance got the final packet exactly once");
  assert.throws(() => c.finalApprove(id, tok), /already been completed/, "second click can't re-send");
  assert.match(env.mail.at(-1).htmlBody, /Open the claim&#39;s Drive folder/);
  // the page tells VP Finance to forward the copy themselves
  assert.match(c.decisionPageHtml_(row(env, id), "vpfinal", payloadOf(env, id)), /requesterEmailFailed/);
  // V5: the same for a signed PDF that came back by email
  const env2 = setup(STAFF, { me: ME, badAddress: "ana@student.university.example.edu" }), c2 = env2.ctx, id2 = "fail0002";
  addClaim(env2, id2, "expense-claim");
  c2.vpSignApprove(id2, row(env2, id2).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c2.presidentSign(id2, row(env2, id2).token, SIG("pres"));
  reply(env2, id2);
  const b2 = env2.mail.length;
  assert.strictEqual(env2.check(), 1);
  assert.strictEqual(row(env2, id2).status, "Approved - signed PDF from NU Staff saved");
  assert.strictEqual(env2.mail.length - b2, 1, "VP Finance got it exactly once");
}

// ---------- 9. Reminders: every signer gets the Drive-folder link ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx;
  const claim = { id: "r1", token: "t", name: "Ana", total: 5, form: "expense-claim", folderUrl: "https://drive/folder" };
  assert.ok(/Drive folder/.test(c.reminderEmailHtml_(claim, "staff", 2)));
  assert.ok(/Drive folder/.test(c.reminderEmailHtml_(claim, "president", 2)));
  assert.ok(/Drive folder/.test(c.reminderEmailHtml_(claim, "vpfinal", 2)));
  // Reminders reach NU Staff, and can't be triggered twice within an hour (anti-spam)
  addClaim(env, "rem00001", "expense-claim");
  c.vpSignApprove("rem00001", row(env, "rem00001").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c.presidentSign("rem00001", row(env, "rem00001").token, SIG("pres"));
  const before = env.mail.length;
  assert.strictEqual(c.sendDailyReminders(), 0, "just signed: no reminder yet");
  delete env.props.LAST_REMINDER_RUN;
  env.log.rows.find(x => x[0] === "rem00001")[13] = new Date(Date.now() - 5.1 * 86400000);
  assert.strictEqual(c.sendDailyReminders(), 1);
  const rem = env.mail.at(-1);
  assert.strictEqual(rem.to, VPF);
  assert.strictEqual(rem.subject, "[Reminder] [Expense Claim] Ana Test - $100.00 - not signed by Sam Patel for the past 5 days - send to Sam Patel to sign (Claim rem00001)");
  assert.match(rem.htmlBody, /Still waiting for Sam&#39;s signed PDF/);
  assert.strictEqual(rem.attachments[0].getDataAsString(), pdfHtml(env, "rem00001"), "the PDF comes with the reminder");
  assert.strictEqual(c.sendDailyReminders(), 0, "second run within the hour is skipped");
  assert.strictEqual(env.mail.length - before, 1);
  // the signed reply is matched as usual
  reply(env, "rem00001");
  assert.strictEqual(env.check(), 1);
}

// ---------- 10. Maintenance functions: owner only ----------
{
  const admin = ["installDailyReminderTrigger", "removeReminderTriggers", "refreshGlList", "setupExpenseLedger",
    "rotateLedgerFeedKey", "backfillExpenseLedger", "migrateLegacyRows", "shareCurrentPresidentPackets", "shareCurrentClaimFolders", "listHeicReceipts", "convertHeicReceipts", "listSentBackClaims", "resendEditLinks", "installSignedPdfWatcher", "removeSignedPdfWatcher", "remindPresident", "remindStaff", "remindVpFinance", "pdfReceiptsToFix", "savePdfReceiptPages", "voidDuplicateClaim", "fixClaimGlCodes"];
  const anon = setup("claims.office@example.org", { activeUser: "" });
  admin.forEach(fn => assert.throws(() => anon.ctx[fn](), /only be run from the Apps Script editor/, fn + " refused for web visitors"));
  assert.ok(!anon.props.LEDGER_FEED_KEY, "no feed key leaked or created");
  const stranger = setup("claims.office@example.org", { activeUser: "someone@gmail.com" });
  assert.throws(() => stranger.ctx.setupExpenseLedger(), /only be run/);
  const noScope = setup("claims.office@example.org", { noScope: true });
  assert.throws(() => noScope.ctx.refreshGlList(), /userinfo\.email/);
  const owner = setup("claims.office@example.org", { activeUser: "vp.finance@example.org" });
  owner.ctx.refreshGlList();
  assert.match(owner.ctx.setupExpenseLedger(), /feed=ledger&key=/, "owner in the editor still works");
  // remindPresident (run by hand from the editor): only claims waiting on the President, only to the President
  addClaim(owner, "presrem1", "expense-claim");
  owner.ctx.vpSignApprove("presrem1", row(owner, "presrem1").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  addClaim(owner, "presrem2", "expense-claim"); // still waiting on VP Finance -> not included
  const m0 = owner.mail.length;
  assert.strictEqual(owner.ctx.remindPresident(), 1);
  assert.strictEqual(owner.mail.length - m0, 1);
  const pm = owner.mail.at(-1);
  assert.strictEqual(pm.to, "president@example.org");
  assert.match(pm.subject, /^\[Reminder\] The expense claim from Ana Test \(\$100\.00\) has not been signed/);
  assert.ok(pm.htmlBody.includes(row(owner, "presrem1").token), "carries the current, working link");
  assert.match(pm.htmlBody, /Sent by NUSA VP Finance/);
  assert.strictEqual(owner.ctx.remindPresident(), 1, "can be re-run (no hourly limit)");
  // remindStaff: V5 signed-PDF claims -> VP Finance gets the hand-off email again; old sign-link claims -> NU Staff
  assert.strictEqual(owner.ctx.remindStaff(), 0, "nothing waiting on NU Staff yet");
  owner.ctx.presidentSign("presrem1", row(owner, "presrem1").token, SIG("pres"));
  addClaim(owner, "staffrem", "expense-claim");
  owner.ctx.vpSignApprove("staffrem", row(owner, "staffrem").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  owner.ctx.presidentSign("staffrem", row(owner, "staffrem").token, SIG("pres"));
  legacyStaff(owner, "staffrem");
  const m1 = owner.mail.length;
  assert.strictEqual(owner.ctx.remindStaff(), 2);
  const [h, l] = owner.mail.slice(m1);
  assert.strictEqual(h.to, VPF);
  assert.match(h.subject, /\(Claim presrem1\)$/);
  assert.match(h.htmlBody, /Still waiting for Sam&#39;s signed PDF/);
  assert.ok(h.attachments.length >= 1, "PDF attached");
  assert.strictEqual(l.to, owner.ctx.CONFIG.STAFF_EMAIL);
  assert.match(l.subject, /^\[Reminder\] The expense claim from Ana Test \(\$100\.00\) has not been signed/);
  assert.ok(l.htmlBody.includes(row(owner, "staffrem").token));
  assert.strictEqual(owner.ctx.remindPresident(), 0, "President no longer has anything waiting");
  // Every browser-callable function is either part of the claim flow (token-checked) or owner-only.
  const src = fs.readFileSync(SRC, "utf8");
  const publicFns = [...src.matchAll(/^function ([A-Za-z0-9]+)\(/gm)].map(m => m[1]);
  const flow = ["doGet", "submitClaim", "renderDecisionPage", "vpSignApprove", "presidentSign", "staffSign", "staffReturn",
    "resendToStaff", "presidentReturn", "resendToPresident", "finalApprove", "sendBackClaim", "sendDailyReminders", "validatePayload", "getClaimForEdit", "resubmitClaim",
    "checkStaffSignedPdfs"];
  publicFns.forEach(fn => assert.ok(flow.includes(fn) || admin.includes(fn), "unreviewed public function: " + fn));
}

// ================= Further fixes =================

// ---------- 11. "Claim not found": IDs that Google Sheets turned into numbers still resolve ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx;
  // What Sheets does to such IDs when they're logged as plain values:
  const cases = [["01234567", 1234567], ["12345e12", 12345e12], ["00000e00", 0]];
  cases.forEach(([id, stored]) => {
    addClaim(env, id, "expense-claim");
    const r = env.log.rows.length - 1;
    env.log.rows[r][0] = stored;                          // the sheet now holds a number
    assert.ok(c.findClaim_(id), "original link id " + id + " finds its row");
  });
  // The whole chain still works for such a claim, using the id from the original email link.
  const id = "01234567", L = c.LC;
  c.vpSignApprove(id, "tok-" + id, SIG("vp"), ["410-72104-56025100003902710000"]);
  const r1 = row(env, id);
  assert.strictEqual(r1.status, "Pending President signature");
  // later links are built from the sheet's value (1234567) and must resolve too
  assert.ok(c.findClaim_(String(r1.id)), "link built from the stored value resolves");
  c.presidentSign(String(r1.id), r1.token, SIG("pres"));
  assert.match(env.mail.at(-1).subject, /\(Claim 1234567\)$/, "subject carries the stored id");
  reply(env, id);
  assert.strictEqual(env.check(), 1, "reply matched to the claim");
  assert.strictEqual(row(env, id).status, "Approved - signed PDF from NU Staff saved");
  const led = c.getLedgerSheet_().rows.slice(1).filter(x => c.sameId_(x[L["Claim ID"]], id) || String(x[L["Claim ID"]]) === String(r1.id));
  assert.ok(led.length && led.every(x => x[L["Status"]] === "Approved"), "ledger rows follow the claim");
  // no false matches between different IDs
  assert.ok(!c.sameId_(1234567, "01234568") && !c.sameId_("abc", "") && !c.sameId_(0, ""));
  // New claims are logged with the ID kept as text, on their own row.
  c.logClaim_({ id: "0000ab12", token: "T-NEW", status: "Pending VP Finance signature" });
  const newRow = env.log.rows.find(x => x[11] === "T-NEW");
  assert.strictEqual(newRow[0], "0000ab12");
}

// ---------- 12. HEIC (and other non-PDF-friendly) photos are attached, not embedded as blanks ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx;
  const p = { lines: [
    { fileName: "IMG_1.HEIC", fileType: "image/heic", fileData: "data:image/heic;base64,AAAA" },
    { fileName: "IMG_2.heic", fileType: "", fileData: "data:application/octet-stream;base64,AAAA" },
    { fileName: "r.jpg", fileType: "image/jpeg", fileData: "data:image/jpeg;base64,AAAA" },
    { fileName: "s.png", fileType: "image/png", fileData: "data:image/png;base64,AAAA" },
    { fileName: "w.webp", fileType: "image/webp", fileData: "data:image/webp;base64,AAAA" },
    { fileName: "fake.jpg", fileType: "image/jpeg", fileData: "data:image/heic;base64,AAAA" } ] };
  const split = c.packetReceipts_(p);
  assert.deepStrictEqual(Array.from(split.embed, e => e.name), ["r.jpg", "s.png"], "only JPEG/PNG go inside the PDF");
  assert.deepStrictEqual(Array.from(split.attach, b => b.name), ["IMG_1.HEIC", "IMG_2.heic", "w.webp", "fake.jpg"], "the rest ride as attachments");
}

// ---------- 12b. PDF receipts go inside the packet as their page pictures ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx;
  const PDF = "data:application/pdf;base64,JVBERi0=", PG = "data:image/jpeg;base64,AAAA";
  const p = { lines: [
    { fileName: "a.jpg", fileType: "image/jpeg", fileData: PG },
    { fileName: "inv.pdf", fileType: "application/pdf", fileData: PDF, filePages: [PG, PG] },
    { fileName: "old.pdf", fileType: "application/pdf", fileData: PDF },                       // claim from before the fix
    { fileName: "bad.pdf", fileType: "application/pdf", fileData: PDF, filePages: [PG, "data:text/html;base64,AA"] } ] };
  const split = c.packetReceipts_(p);
  assert.deepStrictEqual(Array.from(split.embed, e => e.name), ["a.jpg", "inv.pdf (page 1 of 2)", "inv.pdf (page 2 of 2)"]);
  assert.deepStrictEqual(Array.from(split.embed, e => e.no), [1, 2, 2], "both pages are receipt 2");
  assert.deepStrictEqual(Array.from(split.attach, b => b.name), ["old.pdf", "bad.pdf"], "no pages / bad pages -> separate file, as before");
  assert.deepStrictEqual(Array.from(c.embeddedAsBlobs_(split.embed), b => b.name), ["a.jpg", "inv.pdf"], "fallback attaches the PDF once, not its pages");
  const html = c.buildPacketHtml_(Object.assign({ first: "A", last: "B", lines: p.lines.map(l => Object.assign({ date: "2026-10-01", desc: "x y z", amount: 1 }, l)) }), "id1", split.embed);
  assert.ok(html.includes("Receipt 2 - inv.pdf (page 2 of 2)"), "page label in the packet");
  // Mileage attachments carry the pages too; too big for the budget -> whole PDF attached.
  const big = "data:image/jpeg;base64," + "A".repeat(5000000);
  const m = c.packetReceipts_({ form: "mileage", attachments: [{ fileName: "map.pdf", fileType: "application/pdf", fileData: PDF, pages: [big] }] });
  assert.strictEqual(m.embed.length, 0); assert.deepStrictEqual(Array.from(m.attach, b => b.name), ["map.pdf"]);
  // The edit link gives the pages back to the form.
  const src = fs.readFileSync(SRC, "utf8");
  assert.ok(/filePages: Array\.isArray\(l\.filePages\)/.test(src) && /pages: Array\.isArray\(a\.pages\)/.test(src), "getClaimForEdit returns pages");
}

// ---------- 13. The form's script is valid and handles HEIC ----------
{
  const form = fs.readFileSync(require("./paths").file("Form.html"), "utf8");
  const js = form.slice(form.indexOf("<script>") + 8, form.lastIndexOf("</script>"));
  new Function(js); // throws on a syntax error
  const helpers = new Function(js.slice(js.indexOf("function isHeic"), js.indexOf("let heicLib")) + "; return {isHeic, isImageFile, pdfSafeImage};")();
  assert.ok(helpers.isHeic({ type: "image/heic", name: "a.HEIC" }) && helpers.isHeic({ type: "", name: "IMG_0001.HEIC" }) && helpers.isHeic({ type: "image/heif", name: "x" }));
  assert.ok(!helpers.isHeic({ type: "image/jpeg", name: "a.jpg" }));
  assert.ok(helpers.isImageFile({ type: "", name: "IMG_0001.heic" }), "HEIC with no type from Windows still treated as a photo");
  assert.ok(helpers.pdfSafeImage("image/jpeg") && !helpers.pdfSafeImage("image/heic") && !helpers.pdfSafeImage("image/webp"));
  assert.ok(form.includes("https://cdn.jsdelivr.net/npm/heic-to@1.5.2/dist/iife/heic-to.js"), "HEIC converter URL");
  assert.ok(/if\(photosBusy\(\)\)return;/.test(form) && form.split("photosBusy()").length >= 3, "Continue and Submit wait for photo conversion");
  // PDF receipts: page pictures are made, sent, and waited for.
  assert.ok(form.includes("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js") && form.includes("pdf.js/3.11.174/pdf.worker.min.js"), "pdf.js URLs");
  assert.ok(/heicPending>0\|\|pdfPending>0/.test(form), "photosBusy waits for PDFs too");
  assert.ok(form.includes("filePages:l.filePages||[]") && form.includes("pages:a.pages||[]"), "pages go in the payload");
  assert.ok(/fillMissingPdfPages\(\)\.catch/.test(form), "old PDFs get pages on resubmit");
}

// ---------- 14. One-time conversion of HEIC photos already in claims ----------
{
  const HEIC_BYTES = Buffer.concat([Buffer.from([0, 0, 0, 24]), Buffer.from("ftypheic"), Buffer.alloc(20)]);
  const heicUrl = type => "data:" + type + ";base64," + HEIC_BYTES.toString("base64");
  const JPEG = [255, 216, 255, 224, 1, 2, 3];
  let thumbReady = true;
  const fetchMock = (url, files) => {
    const m = url.match(/drive\/v3\/files\/([^?]+)\?fields=thumbnailLink/);
    if (m) {
      const f = files[m[1]];
      const ok = thumbReady && f && f.mime === "image/heic";
      return { getResponseCode: () => 200, getContentText: () => JSON.stringify(ok ? { thumbnailLink: "https://lh3.googleusercontent.com/drive-storage/abc=s220" } : {}) };
    }
    if (/lh3\.googleusercontent\.com/.test(url)) {
      assert.ok(/=s1600$/.test(url), "asks for the large preview");
      return { getResponseCode: () => 200, getBlob: () => ({ getContentType: () => "image/jpeg", getBytes: () => JPEG, setName(n) { this.name = n; return this; }, html: JPEG, type: "image/jpeg", name: "x" }) };
    }
    return { getResponseCode: () => 404, getContentText: () => "{}" };
  };
  const env = setup("claims.office@example.org", { activeUser: "vp.finance@example.org", fetch: fetchMock }), c = env.ctx;
  // A: expense claim waiting on the President — one HEIC (labelled), one normal JPG
  addClaim(env, "heic0001", "expense-claim", { lines: [
    { date: "2026-09-18", amount: 20, desc: "Pizza for nursing event", fileName: "IMG_1.HEIC", fileType: "image/heic", fileData: heicUrl("image/heic") },
    { date: "2026-09-18", amount: 5, desc: "Drinks for nursing event", fileName: "r.jpg", fileType: "image/jpeg", fileData: "data:image/jpeg;base64,/9j/AAAA" } ],
    vpSignature: SIG("vp") });
  env.log.rows[env.log.rows.length - 1][2] = "Pending President signature";
  // B: kilometre claim, approved — HEIC recognised only by its contents (Windows gave no name/type)
  addClaim(env, "heic0002", "mileage", { attachments: [{ fileName: "attachment-1", fileType: "", fileData: heicUrl("application/octet-stream") }] });
  env.log.rows[env.log.rows.length - 1][2] = "Approved - final PDF sent to VP Finance";
  // C: sent back — left alone
  addClaim(env, "heic0003", "expense-claim", { lines: [{ amount: 5, desc: "x y z", fileName: "IMG_3.heic", fileType: "image/heic", fileData: heicUrl("image/heic") }] });
  env.log.rows[env.log.rows.length - 1][2] = "Sent back for changes";
  // D: waiting on NU Staff; VP already put a converted IMG_4.jpg in the folder by hand (manual fallback)
  addClaim(env, "heic0004", "expense-claim", { lines: [{ amount: 7, desc: "Tape for posters", fileName: "IMG_4.HEIC", fileType: "image/heic", fileData: heicUrl("image/heic") }] });
  env.log.rows[env.log.rows.length - 1][2] = "Pending NU Staff signature";
  env.newFile(JPEG, "IMG_4.jpg", "image/jpeg");

  // Owner only
  const anon = setup("claims.office@example.org", { activeUser: "" });
  assert.throws(() => anon.ctx.convertHeicReceipts(), /only be run from the Apps Script editor/);

  // 1) list: read-only
  const before = JSON.stringify(env.files);
  const list = c.listHeicReceipts();
  assert.strictEqual(list.found, 3, "A, B and D (not the sent-back claim)");
  assert.strictEqual(JSON.stringify(env.files), before, "listing changes nothing");
  assert.strictEqual(env.fetches.length, 0);

  // 2) convert
  const pdfA = row(env, "heic0001").pdfId;
  const res = c.convertHeicReceipts();
  assert.strictEqual(res.fixed, 3); assert.strictEqual(res.failed, 0);
  const pA = payloadOf(env, "heic0001");
  assert.strictEqual(pA.lines[0].fileName, "IMG_1.jpg");
  assert.strictEqual(pA.lines[0].fileType, "image/jpeg");
  assert.strictEqual(pA.lines[0].fileData, "data:image/jpeg;base64," + Buffer.from(JPEG).toString("base64"));
  assert.strictEqual(pA.lines[1].fileData, "data:image/jpeg;base64,/9j/AAAA", "normal JPG untouched");
  assert.strictEqual(pA.vpSignature, SIG("vp"), "signatures kept");
  const rA = row(env, "heic0001");
  assert.notStrictEqual(rA.pdfId, pdfA, "PDF rebuilt");
  assert.strictEqual(rA.status, "Pending President signature", "status unchanged");
  assert.strictEqual(rA.token, "tok-heic0001", "email links still valid");
  assert.deepStrictEqual(env.files[rA.pdfId].viewers, ["president@example.org"], "new PDF shared with the President");
  assert.ok(env.files[rA.pdfId].content.includes("data:image/jpeg;base64," + Buffer.from(JPEG).toString("base64")), "receipt now inside the PDF");
  assert.ok(Object.values(env.files).some(f => f.name === "IMG_1.jpg" && !f.trashed), "JPG copy saved in the claim folder");
  assert.ok(Object.values(env.files).filter(f => /^convert-/.test(f.name)).every(f => f.trashed), "temporary HEIC files cleaned up");
  assert.ok(env.fetches.every(f => f.auth === "Bearer TOKEN"));
  assert.strictEqual(payloadOf(env, "heic0002").attachments[0].fileType, "image/jpeg", "HEIC found by its contents converted too");
  const rD = row(env, "heic0004");
  assert.strictEqual(payloadOf(env, "heic0004").lines[0].fileName, "IMG_4.jpg");
  assert.deepStrictEqual(env.files[rD.pdfId].viewers, [], "V5: no Drive sharing with NU Staff");
  assert.ok(env.fetches.length === 4, "D used the manual JPG (no Drive preview calls): " + env.fetches.length);
  assert.ok(payloadOf(env, "heic0003").lines[0].fileType === "image/heic", "sent-back claim untouched");
  assert.strictEqual(env.mail.length, 0, "no emails sent");
  // 3) re-run: nothing left
  assert.strictEqual(c.listHeicReceipts().found, 0);

  // 4) Drive preview not ready yet -> clear message, claim unchanged, can re-run later
  const env2 = setup("claims.office@example.org", { activeUser: "vp.finance@example.org", fetch: fetchMock });
  addClaim(env2, "heic0005", "expense-claim", { lines: [{ amount: 5, desc: "a b c", fileName: "IMG_5.HEIC", fileType: "image/heic", fileData: heicUrl("image/heic") }] });
  thumbReady = false;
  const pdf5 = row(env2, "heic0005").pdfId;
  const r5 = env2.ctx.convertHeicReceipts();
  assert.strictEqual(r5.fixed, 0); assert.strictEqual(r5.failed, 1);
  assert.match(r5.details.join("\n"), /IMG_5\.HEIC - Google Drive has not made a preview of this photo yet/);
  assert.strictEqual(row(env2, "heic0005").pdfId, pdf5, "claim left as it was");
  thumbReady = true;
  assert.strictEqual(env2.ctx.convertHeicReceipts().fixed, 1, "works on the next run");
}

// ---------- Signer folder access: VP Finance when not the owner; backfill of existing claims ----------
{
  const stages = { sh1: "Pending VP Finance signature", sh2: "Pending President signature", sh3: "Returned by President",
    sh4: "Pending NU Staff signature", sh5: "Pending VP Finance final approval", sh6: "Approved - final PDF sent to VP Finance",
    sh7: "Sent back for changes", sh8: "Pending NU Staff signed PDF" };
  const V = "vp.finance@example.org", P = "president@example.org", S = "claims.office@example.org";
  const want = { sh1: [V], sh2: [V, P], sh3: [V, P], sh4: [V, P], sh5: [V, P], sh6: [V, P], sh7: [], sh8: [V, P] };
  // script owned by another account (e.g. a personal account): VP Finance must be shared too
  const env = setup(S, { owner: "owner@example.org", activeUser: V }), c = env.ctx;
  Object.keys(stages).forEach(id => { addClaim(env, id, "expense-claim"); env.log.rows[env.log.rows.length - 1][2] = stages[id]; });
  assert.strictEqual(c.shareCurrentClaimFolders(), 7, "every claim except the sent-back one");
  Object.keys(want).forEach(id => assert.deepStrictEqual(env.files[row(env, id).pdfId].viewers, want[id], id + " shared with " + want[id]));
  assert.deepStrictEqual([...new Set(env.folder.viewers)].sort(), [V, P].sort(), "folders shared, never the owner or NU Staff");
  assert.ok(!env.folder.viewers.includes("owner@example.org"));
  // owner = VP Finance account: VP is skipped (already owns everything)
  const env2 = setup(S, { activeUser: V });
  addClaim(env2, "own00001", "expense-claim"); env2.log.rows[env2.log.rows.length - 1][2] = stages.sh4;
  env2.ctx.shareCurrentClaimFolders();
  assert.deepStrictEqual(env2.files[row(env2, "own00001").pdfId].viewers, [P]);
  // a signer's folder access is given when the claim reaches them (VP at submission is covered by submitClaim)
  const env3 = setup(S, { owner: "owner@example.org" }), c3 = env3.ctx;
  addClaim(env3, "flow0001", "expense-claim");
  c3.vpSignApprove("flow0001", row(env3, "flow0001").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  assert.deepStrictEqual(env3.folder.viewers, [P], "President gets the folder before signing");
}

// ================= Edit link for sent-back claims =================

// ---------- E1. Sent back -> the requester gets an EDIT link to their own claim, fixes it, resubmits ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx, id = "edit0001";
  const R1 = "data:image/png;base64,AAAA", R2 = "data:image/jpeg;base64,BBBB";
  addClaim(env, id, "expense-claim", { studentId: "123456", address: "1 Main St", city: "Northfield", delivery: "Pick Up",
    lines: [{ date: "2026-09-18", amount: 100, gst: 5, hst: 0, net: 96.65, desc: "Pizza for nursing event", fileName: "r1.png", fileType: "image/png", fileData: R1 }] });
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), ["410-72104-56025100003902710000"]);
  c.presidentReturn(id, row(env, id).token, "Ask for a clearer receipt");
  const LCC = c.LC["Claim ID"], led = () => c.getLedgerSheet_().rows.filter(x => x[LCC] === id);
  assert.strictEqual(led().length, 1);
  const vpTok = row(env, id).token;
  c.sendBackClaim(id, vpTok, "The receipt is blurry,\nplease attach a clearer photo");
  const m = env.mail.at(-1), tok = row(env, id).token;
  assert.strictEqual(m.to, "ana@student.university.example.edu");
  assert.notStrictEqual(tok, vpTok, "token rotated: VP Finance's old links stop working");
  assert.ok(m.htmlBody.includes("?action=edit&id=edit0001&token=" + encodeURIComponent(tok)), "the email links to THIS claim");
  assert.ok(!m.htmlBody.includes('href="' + c.CONFIG.WEB_APP_URL + '"'), "no more link to the empty form");
  assert.match(m.htmlBody, /The receipt is blurry,<br>please attach a clearer photo/);
  assert.match(m.htmlBody, /Edit your claim/);
  assert.strictEqual(row(env, id).status, "Sent back for changes");
  assert.ok(led().every(x => x[c.LC["Status"]] === "Void"));
  // the edit page loads the claim (only with the right token)
  assert.throws(() => c.getClaimForEdit(id, vpTok), /not valid any more/);
  assert.throws(() => c.getClaimForEdit(id, ""), /not valid any more/);
  assert.throws(() => c.getClaimForEdit("nope0000", tok), /not valid any more/);
  const ed = c.getClaimForEdit(id, tok);
  assert.strictEqual(ed.note, "The receipt is blurry,\nplease attach a clearer photo");
  assert.strictEqual(ed.form, "expense-claim");
  assert.strictEqual(ed.first, "Ana"); assert.strictEqual(ed.studentId, "123456"); assert.strictEqual(ed.city, "Northfield");
  assert.strictEqual(ed.lines.length, 1);
  assert.strictEqual(ed.lines[0].fileData, R1, "receipts come back too");
  assert.strictEqual(ed.lines[0].desc, "Pizza for nursing event");
  assert.strictEqual(ed.signature, SIG("req"));
  assert.ok(!("vpSignature" in ed) && !("presidentSignature" in ed) && !("presidentReturnNote" in ed), "approvers' data isn't sent to the requester");
  // resubmit: clearer receipt on line 1, a new line 2
  const fixed = { form: "expense-claim", payTo: "Student", first: "Ana", last: "Test", email: "ana@student.university.example.edu", studentId: "123456",
    address: "1 Main St", city: "Northfield", country: "Canada", usd: false, printName: "Ana Test", delivery: "Pick Up", signature: SIG("req2"),
    totals: { amt: 130, gst: 6, hst: 0, net: 126.0 },
    lines: [{ date: "2026-09-18", desc: "Pizza for nursing event", amount: "100", gst: "5", hst: "0", net: 96.65, fileName: "r1-clear.jpg", fileType: "image/jpeg", fileData: R2 },
            { date: "2026-09-19", desc: "Drinks for nursing event", amount: "30", gst: "1", hst: "0", net: 29.33, fileName: "r1.png", fileType: "image/png", fileData: R1 }] };
  const copy = () => JSON.parse(JSON.stringify(fixed));
  assert.throws(() => c.resubmitClaim(id, vpTok, copy()), /not valid any more/);
  assert.throws(() => c.resubmitClaim(id, tok, Object.assign(copy(), { lines: [] })), /missing: at least one expense/);
  const asMileage = Object.assign(copy(), { form: "mileage", lines: [{ date: "2026-09-18", km: 10, destination: "A - B" }], attachments: [{ fileName: "m.png", fileType: "image/png", fileData: R1 }] });
  assert.throws(() => c.resubmitClaim(id, tok, asMileage), /change type/);
  assert.strictEqual(row(env, id).status, "Sent back for changes", "refused attempts change nothing");
  const oldPdfId = row(env, id).pdfId, nFiles = Object.keys(env.files).length, nMail = env.mail.length;
  const res = c.resubmitClaim(id, tok, copy());
  assert.ok(res.ok && res.resubmitted);
  assert.strictEqual(res.id, id);
  const r = row(env, id);
  assert.strictEqual(r.status, "Pending VP Finance signature");
  assert.strictEqual(c.stageOf_(r.status), "vp1");
  assert.strictEqual(r.total, 130); assert.strictEqual(r.gst, 6);
  assert.notStrictEqual(r.pdfId, oldPdfId, "new PDF");
  assert.match(env.files[oldPdfId].name, /\(sent back\)\.pdf$/, "the sent-back PDF is kept, renamed");
  assert.ok(!env.files[oldPdfId].trashed);
  const added = Object.keys(env.files).slice(nFiles).map(k => env.files[k].name);
  assert.ok(added.includes("r1-clear.jpg"), "the new receipt is saved");
  assert.ok(!added.includes("r1.png"), "a receipt already in the folder isn't saved twice");
  assert.ok(added.some(n => /\(resubmitted\)\.pdf$/.test(n)));
  const p = payloadOf(env, id);
  assert.ok(!p.vpSignature && !p.presidentSignature, "approvals start over");
  assert.strictEqual(p.signature, SIG("req2"));
  assert.strictEqual(p.lines[0].gl, "410-72104-56025100003902710000", "VP's GL code pre-selected on the same line");
  assert.strictEqual(p.lines[1].gl, undefined);
  assert.strictEqual(p.resubmissions, 1);
  assert.strictEqual(p.lastSentBackNote, "The receipt is blurry,\nplease attach a clearer photo");
  assert.ok(p.firstSubmittedAt && p.submittedAt);
  assert.strictEqual(led().length, 0, "old ledger rows removed (VP Finance writes them again when signing)");
  const sent = env.mail.slice(nMail);
  assert.strictEqual(sent.length, 1);
  assert.strictEqual(sent[0].to, c.CONFIG.VP_FINANCE_EMAIL);
  assert.match(sent[0].subject, /Ana Test - \$130\.00 - resubmitted after changes - needs your signature/);
  assert.ok(sent[0].htmlBody.includes("action=decision") && sent[0].htmlBody.includes(encodeURIComponent(r.token)), "VP gets a working sign link");
  assert.match(sent[0].htmlBody, /The receipt is blurry/);
  // the edit link is used up
  assert.throws(() => c.getClaimForEdit(id, tok), /not valid any more/);
  assert.throws(() => c.resubmitClaim(id, r.token, copy()), /already resubmitted/);
  // VP Finance signs again: both lines go to the ledger
  c.vpSignApprove(id, r.token, SIG("vp2"), [p.lines[0].gl, "410-72104-56025100003902710000"]);
  assert.strictEqual(row(env, id).status, "Pending President signature");
  assert.deepStrictEqual(led().map(x => x[c.LC["Amount (incl. tax)"]]), [100, 30]);
  assert.ok(led().every(x => x[c.LC["Status"]] === "Pending"));
}

// ---------- E2. Mileage claims and locked ledger rows ----------
{
  const env = setup("claims.office@example.org"), c = env.ctx, id = "edit0002";
  const R = "data:image/png;base64,CCCC";
  addClaim(env, id, "mileage", { studentId: "1", address: "a", city: "b", attachments: [{ fileName: "map.png", fileType: "image/png", fileData: R }] });
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), []);
  c.presidentReturn(id, row(env, id).token, "check");
  const sh = c.getLedgerSheet_(), i = sh.rows.findIndex(x => x[c.LC["Claim ID"]] === id);
  sh.rows[i][c.LC["Locked"]] = "Yes"; // VP Finance corrected this row by hand
  c.sendBackClaim(id, row(env, id).token, "Add the parking receipt");
  const tok = row(env, id).token, ed = c.getClaimForEdit(id, tok);
  assert.strictEqual(ed.form, "mileage");
  assert.strictEqual(ed.lines[0].destination, "Northfield - Vancouver");
  assert.strictEqual(ed.attachments[0].fileData, R);
  assert.ok(!("fileData" in ed.lines[0]), "trips have no receipt field");
  const pay = { form: "mileage", payTo: "Student", first: "Ana", last: "Test", email: "ana@student.university.example.edu", studentId: "1", address: "a", city: "b",
    printName: "Ana Test", signature: SIG("req"), totals: { amt: 37, km: 50, parking: 7, kmReimb: 30 },
    lines: [{ date: "2026-09-18", km: "50", destination: "Northfield - Vancouver", purpose: "Conference", parking: "7", amount: 37, desc: "Northfield - Vancouver" }],
    attachments: [{ fileName: "map.png", fileType: "image/png", fileData: R }, { fileName: "parking.png", fileType: "image/png", fileData: "data:image/png;base64,DDDD" }] };
  c.resubmitClaim(id, tok, pay);
  assert.strictEqual(row(env, id).status, "Pending VP Finance signature");
  assert.strictEqual(payloadOf(env, id).lines[0].gl, vm.runInContext("MILEAGE_GL", c));
  assert.strictEqual(sh.rows.filter(x => x[c.LC["Claim ID"]] === id).length, 1, "a locked row is kept");
}

// ---------- E3. Claims sent back before edit links existed: resendEditLinks (owner only) ----------
{
  const env = setup("claims.office@example.org", { activeUser: "vp.finance@example.org" }), c = env.ctx;
  addClaim(env, "old00009", "expense-claim"); env.log.rows.at(-1)[2] = "Sent back for changes";
  addClaim(env, "old00010", "expense-claim"); env.log.rows.at(-1)[2] = "Pending VP Finance signature";
  assert.strictEqual(c.listSentBackClaims(), 1);
  assert.strictEqual(c.resendEditLinks(), 0, "nothing is sent until IDs are listed");
  vm.runInContext("RESEND_EDIT_LINK_IDS.push('old00009', 'old00010', 'missing1')", c);
  const n = env.mail.length;
  assert.strictEqual(c.resendEditLinks(), 1, "only the sent-back claim gets a link");
  const m = env.mail.at(-1), tok = row(env, "old00009").token;
  assert.strictEqual(env.mail.length, n + 1);
  assert.strictEqual(m.to, "ana@student.university.example.edu");
  assert.ok(m.htmlBody.includes("action=edit&id=old00009&token=" + encodeURIComponent(tok)));
  assert.match(m.htmlBody, /earlier email/);
  assert.strictEqual(c.getClaimForEdit("old00009", tok).note, "");
  assert.notStrictEqual(row(env, "old00009").token, "tok-old00009");
  assert.strictEqual(row(env, "old00010").token, "tok-old00010", "other claims untouched");
}

// ---------- E4. The form opens an edit link, fills the claim in, and resubmits it ----------
{
  const form = fs.readFileSync(require("./paths").file("Form.html"), "utf8");
  const js = form.slice(form.indexOf("<script>") + 8, form.lastIndexOf("</script>"));
  assert.ok(/\nboot\(\);\s*$/.test(js), "the form starts with boot()");
  assert.ok(js.includes("google.script.url.getLocation") && js.includes("getClaimForEdit(id,token)") && js.includes("run.resubmitClaim(S.edit.id,S.edit.token,buildPayload())"));
  const lib = new Function("document", "window", js.replace(/\nboot\(\);\s*$/, "") + "\nreturn {S, fillFromClaim, buildPayload, allValid, editBanner};");
  const f = lib({}, {});
  const claim = { id: "edit0001", form: "expense-claim", note: "fix it", payTo: "Student", first: "Ana", last: "Test", email: "ana@my.university.example.edu", studentId: "123",
    company: "", contact: "", address: "1 Main", city: "Northfield", country: "Canada", usd: false, printName: "Ana Test", delivery: "Mail", signature: SIG("req"),
    lines: [{ date: "2026-09-18", desc: "Pizza for nursing event", amount: 100, gst: 5, hst: "", fileName: "r1.png", fileType: "image/png", fileData: "data:image/png;base64,AAAA" }], attachments: [] };
  f.fillFromClaim(claim);
  assert.strictEqual(f.S.formType, "expense");
  assert.strictEqual(f.S.lines[0].file, "r1.png");
  assert.strictEqual(f.S.lines[0].fileThumb, "data:image/png;base64,AAAA");
  assert.strictEqual(f.S.sig, SIG("req"));
  assert.ok(f.allValid(), "a filled-in claim is ready to resubmit as is");
  const out = f.buildPayload();
  assert.strictEqual(out.form, "expense-claim");
  assert.strictEqual(out.delivery, "Mail");
  assert.strictEqual(out.lines[0].fileData, "data:image/png;base64,AAAA");
  assert.strictEqual(out.lines[0].amount, "100");
  assert.strictEqual(out.totals.amt, 100);
  f.fillFromClaim({ id: "m1", form: "mileage", first: "A", last: "B", email: "a@my.university.example.edu", studentId: "1", address: "x", city: "y", printName: "A B", signature: SIG("r"),
    lines: [{ date: "2026-09-18", km: 50, destination: "Northfield - Vancouver", purpose: "", parking: 5 }],
    attachments: [{ fileName: "map.pdf", fileType: "application/pdf", fileData: "data:application/pdf;base64,JVBE" }] });
  assert.strictEqual(f.S.formType, "mileage");
  assert.strictEqual(f.S.attachments[0].thumb, null, "PDF has no thumbnail");
  assert.ok(f.allValid());
  const mp = f.buildPayload();
  assert.strictEqual(mp.totals.amt, 35);
  assert.strictEqual(mp.attachments[0].fileData, "data:application/pdf;base64,JVBE");
}

// ================= Reminders at 5, 7 and 14 days + remindVpFinance =================
{
  const env = setup("claims.office@example.org"), c = env.ctx, VPM = vm.runInContext("CONFIG", c).VP_FINANCE_EMAIL;
  const DAY = 86400000, T0 = Date.now();
  let now = T0; // a fake clock for the script: at(d) = d days (+1 hour) after the claims were submitted
  const RealDate = Date;
  c.Date = class extends RealDate { constructor(...a) { if (a.length) super(...a); else super(now); } static now() { return now; } };
  let base = T0;
  const at = d => { now = base + d * DAY + 3600000; };
  const waitDays = (id, d) => at(d);
  const run = () => { delete env.props.LAST_REMINDER_RUN; const n = env.mail.length; const k = c.sendDailyReminders(); assert.strictEqual(env.mail.length - n, k); return env.mail.slice(n); };
  addClaim(env, "r5d00001", "expense-claim", null, new RealDate(T0));   // waiting on VP Finance (no "Decided at" yet: counts from submission)
  at(2); assert.strictEqual(run().length, 0, "nothing before 5 days");
  at(4); assert.strictEqual(run().length, 0);
  waitDays("r5d00001", 5);
  let m = run();
  assert.strictEqual(m.length, 1);
  assert.strictEqual(m[0].to, VPM);
  assert.strictEqual(m[0].subject, "[Reminder] The expense claim from Ana Test ($100.00) has not been signed for the past 5 days");
  assert.match(m[0].htmlBody, /has been waiting for 5 days/);
  assert.match(m[0].htmlBody, /sent when a claim has waited 5, 7 and 14 days/);
  assert.strictEqual(run().length, 0, "the 5-day reminder goes out once");
  waitDays("r5d00001", 6); assert.strictEqual(run().length, 0, "nothing on day 6");
  waitDays("r5d00001", 7); m = run();
  assert.strictEqual(m.length, 1); assert.match(m[0].subject, /has not been signed for the past 7 days$/);
  waitDays("r5d00001", 10); assert.strictEqual(run().length, 0);
  waitDays("r5d00001", 14); m = run();
  assert.strictEqual(m.length, 1); assert.match(m[0].subject, /for the past 14 days$/);
  waitDays("r5d00001", 20); assert.strictEqual(run().length, 0, "nothing after the 14-day reminder");
  // first run on a claim already 9 days old: only the 7-day reminder, once
  addClaim(env, "r5d00002", "mileage", null, new RealDate(now - 9 * DAY));
  m = run();
  assert.strictEqual(m.length, 1);
  assert.strictEqual(m[0].subject, "[Reminder] The kilometer reimbursement from Ana Test ($100.00) has not been signed for the past 9 days");
  assert.strictEqual(run().length, 0);
  // when the claim moves on, the count starts again for the next signer
  waitDays("r5d00001", 20);
  c.vpSignApprove("r5d00001", row(env, "r5d00001").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  base = now - 3600000;
  assert.strictEqual(run().length, 0, "the President's clock starts at 0");
  waitDays("r5d00001", 5); m = run();
  assert.strictEqual(m.length, 2, "the President's 5-day reminder + the mileage claim's 14-day one");
  assert.match(m[1].subject, /kilometer reimbursement .* for the past 14 days$/);
  m = m.filter(x => /expense claim/.test(x.subject));
  assert.strictEqual(m[0].to, "president@example.org");
  assert.match(m[0].subject, /has not been signed for the past 5 days$/);
  // returned claims: "waiting on you"
  c.presidentReturn("r5d00001", row(env, "r5d00001").token, "check");
  base = now - 3600000;
  waitDays("r5d00001", 7); m = run();
  assert.match(m[0].subject, /\(\$100\.00\) has been waiting on you for the past 7 days$/);
  // marks of finished claims are cleaned up
  c.sendBackClaim("r5d00001", row(env, "r5d00001").token, "fix");
  run();
  assert.ok(!("REM_r5d00001" in env.props) && "REM_r5d00002" in env.props);
  assert.strictEqual(c.reminderSubject_({ form: "expense-claim", name: "A", total: 1 }, "vp1", 1), "[Reminder] The expense claim from A ($1.00) has not been signed for the past 1 day");
  assert.strictEqual(c.reminderSubject_({ form: "expense-claim", name: "A", total: 1 }, "vp1", 0), "[Reminder] The expense claim from A ($1.00) has not been signed yet");

  // remindVpFinance: by hand, any time, only claims waiting on VP Finance, only to VP Finance
  const owner = setup("claims.office@example.org", { activeUser: "vp.finance@example.org" }), o = owner.ctx;
  addClaim(owner, "me000001", "expense-claim");                       // waiting on VP signature
  addClaim(owner, "me000002", "expense-claim");
  o.vpSignApprove("me000002", row(owner, "me000002").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  o.presidentReturn("me000002", row(owner, "me000002").token, "check"); // returned to VP Finance
  addClaim(owner, "me000003", "expense-claim");
  o.vpSignApprove("me000003", row(owner, "me000003").token, SIG("vp"), ["410-72104-56025100003902710000"]); // waiting on the President
  const n0 = owner.mail.length;
  assert.strictEqual(o.remindVpFinance(), 2);
  const mine = owner.mail.slice(n0);
  assert.ok(mine.every(x => x.to === VPM), "only VP Finance is emailed");
  assert.ok(mine[0].htmlBody.includes(encodeURIComponent(row(owner, "me000001").token)) || mine[0].htmlBody.includes(row(owner, "me000001").token), "working sign link");
  assert.match(mine[0].subject, /^\[Reminder\] The expense claim from Ana Test \(\$100\.00\) has not been signed/);
  assert.match(mine[1].subject, /is waiting on you|has been waiting on you/);
  assert.match(mine[0].htmlBody, /Sent by NUSA VP Finance/);
  assert.strictEqual(o.remindVpFinance(), 2, "can be re-run any time");
  const anon = setup("claims.office@example.org", { activeUser: "" });
  assert.throws(() => anon.ctx.remindVpFinance(), /only be run from the Apps Script editor/);
}

// ---------- F1. One-time fix PDF receipts of older claims go inside the PDF; restore / void ----------
{
  const env = setup("claims.office@example.org", { activeUser: "vp.finance@example.org" }), c = env.ctx;
  const PDF = "data:application/pdf;base64,JVBERi0=", PG = "data:image/jpeg;base64,AAAA";
  const pdfLine = { date: "2026-10-01", amount: 51.56, gst: 2.34, hst: 0, net: 49.99, desc: "Snacks for late night", fileName: "receipts.pdf", fileType: "application/pdf", fileData: PDF };
  // (a) pending President: pages saved, PDF rebuilt with them, status unchanged, President can open the new PDF
  addClaim(env, "fixpres1", "expense-claim", { lines: [Object.assign({}, pdfLine)] });
  c.vpSignApprove("fixpres1", row(env, "fixpres1").token, SIG("vp"), ["410-72104-56025100003902710000"]);
  const todo = c.pdfReceiptsToFix("fixpres1");
  assert.deepStrictEqual(JSON.parse(JSON.stringify(todo.items.map(i => [i.kind, i.idx, i.fileName]))), [["line", 0, "receipts.pdf"]]);
  const oldPdf = row(env, "fixpres1").pdfId, tok = row(env, "fixpres1").token;
  assert.match(c.savePdfReceiptPages("fixpres1", [{ kind: "line", idx: 0, pages: [PG, PG] }], false), /rebuilt/);
  const r1 = row(env, "fixpres1");
  assert.strictEqual(r1.status, "Pending President signature"); assert.strictEqual(r1.token, tok, "links still work");
  assert.notStrictEqual(r1.pdfId, oldPdf); assert.ok(env.files[oldPdf].trashed);
  assert.ok(pdfHtml(env, "fixpres1").includes("receipts.pdf (page 2 of 2)") && pdfHtml(env, "fixpres1").includes("SIGvp"), "receipt pages + VP signature in the PDF");
  assert.ok(env.files[r1.pdfId].viewers.includes(c.CONFIG.PRESIDENT_EMAIL));
  assert.strictEqual(c.pdfReceiptsToFix("fixpres1").items.length, 0, "nothing left to fix");
  // bad pages are ignored
  c.savePdfReceiptPages("fixpres1", [{ kind: "line", idx: 0, pages: ["data:text/html;base64,AA"] }], false);
  assert.deepStrictEqual(Array.from(payloadOf(env, "fixpres1").lines[0].filePages), [PG, PG]);
  // (b) sent back for a "missing" receipt -> restored to VP Finance with a new sign link
  addClaim(env, "fixback1", "expense-claim", { lines: [Object.assign({}, pdfLine)] });
  c.sendBackClaim("fixback1", row(env, "fixback1").token, "You are missing a receipt");
  const editTok = row(env, "fixback1").token, nMail = env.mail.length;
  assert.throws(() => c.savePdfReceiptPages("fixpres1", [], true), /not sent back/);
  assert.match(c.savePdfReceiptPages("fixback1", [{ kind: "line", idx: 0, pages: [PG] }], true), /Pending VP Finance signature/);
  const r2 = row(env, "fixback1");
  assert.strictEqual(r2.status, "Pending VP Finance signature"); assert.notStrictEqual(r2.token, editTok, "old edit link stops working");
  const m = env.mail.slice(nMail); assert.strictEqual(m.length, 1); assert.strictEqual(m[0].to, c.CONFIG.VP_FINANCE_EMAIL);
  assert.match(m[0].subject, /receipt now inside the PDF - needs your signature/);
  assert.ok(m[0].htmlBody.includes("action=decision&id=fixback1&token=" + encodeURIComponent(r2.token)));
  assert.ok(pdfHtml(env, "fixback1").includes("Receipt 1 - receipts.pdf"));
  assert.strictEqual(payloadOf(env, "fixback1").lastSentBackNote, "You are missing a receipt"); assert.ok(!("sentBackNote" in payloadOf(env, "fixback1")));
  c.vpSignApprove("fixback1", r2.token, SIG("vp"), ["410-72104-56025100003902710000"]);
  assert.strictEqual(row(env, "fixback1").status, "Pending President signature", "VP Finance can sign it as usual");
  // (c) a sent-back claim not restored just keeps the pages for its resubmission
  addClaim(env, "fixback2", "expense-claim", { lines: [Object.assign({}, pdfLine)] });
  c.sendBackClaim("fixback2", row(env, "fixback2").token, "One line per expense");
  assert.match(c.savePdfReceiptPages("fixback2", [{ kind: "line", idx: 0, pages: [PG] }], false), /when the claim is resubmitted/);
  assert.strictEqual(row(env, "fixback2").status, "Sent back for changes");
  assert.deepStrictEqual(Array.from(c.getClaimForEdit("fixback2", row(env, "fixback2").token).lines[0].filePages), [PG]);
  // (d) duplicate: the old sent-back copy is voided
  assert.throws(() => c.voidDuplicateClaim("fixpres1", "fixback2"), /not sent back/);
  addClaim(env, "fixdup01", "expense-claim"); c.sendBackClaim("fixdup01", row(env, "fixdup01").token, "dup");
  assert.match(c.voidDuplicateClaim("fixdup01", "fixback2"), /Void - replaced by claim fixback2/);
  assert.strictEqual(row(env, "fixdup01").status, "Void - replaced by claim fixback2");
  // the fix page itself is owner-only
  const HS = { createHtmlOutput: h => ({ html: h, setTitle() { return this; } }) };
  c.HtmlService = HS;
  assert.match(c.doGet({ parameter: { action: "pdfpages", ids: "a1,b2", restore: "b2" } }).html || "", /PDF receipts|CFG/);
  const anon = setup("claims.office@example.org", { activeUser: "" });
  anon.ctx.HtmlService = HS;
  assert.throws(() => anon.ctx.doGet({ parameter: { action: "pdfpages", ids: "x" } }), /only be run/);
  ["pdfReceiptsToFix", "savePdfReceiptPages", "voidDuplicateClaim"].forEach(fn => assert.throws(() => anon.ctx[fn]("x"), /only be run/));
}

// ---------- fixClaimGlCodes: wrong department picked when signing (a Yearbook claim signed with The Courier codes) ----------
{
  const env = setup(STAFF, { me: ME, owner: ME, activeUser: ME }), c = env.ctx, id = "abc12345";
  vm.runInContext("GL_FIXES.abc12345 = { 1: '410-72108-56025100003906310000', 3: '410-72108-56025100003906310000', 4: '410-72108-56025100003906310000', 5: '410-72108-55600100001000010000', 6: '410-72108-56025100003906310000' };", c);
  const MH_SE = "410-72107-56025100003906310000", MH_OS = "410-72107-55600100001000010000";
  const PI_SE = "410-72108-56025100003906310000", PI_OS = "410-72108-55600100001000010000";
  const line = (amount, desc) => ({ date: "2026-10-09", amount, gst: 0, hst: 0, net: amount, desc });
  addClaim(env, id, "expense-claim", { lines: [line(35.2, "Breakfast"), line(1.58, "Breakfast"), line(96.09, "Lunch"), line(8.48, "Lunch"), line(72.5, "Office supplies for Yearbook"), line(8.14, "Snacks")] });
  c.vpSignApprove(id, row(env, id).token, SIG("vp"), [MH_SE, PI_SE, MH_SE, MH_SE, MH_OS, MH_SE]);
  const led = () => env.ctx.getLedgerSheet_().getDataRange().getValues().filter(r => r[c.LC["Claim ID"]] === id);
  assert.deepStrictEqual(led().map(r => r[c.LC["GL Group"]]), ["The Courier", "Yearbook", "The Courier", "The Courier", "The Courier", "The Courier"]);
  // your own ledger columns survive
  const sh = c.getLedgerSheet_(), r3 = sh.rows.findIndex(r => r[0] === id + "-3");
  sh.rows[r3][c.LC["Notes"]] = "my note";
  const oldTok = row(env, id).token, oldPdf = row(env, id).pdfId, n0 = env.mail.length;
  const out = c.fixClaimGlCodes();
  assert.strictEqual(out.length, 1); assert.match(out[0], /5 GL code\(s\) changed.*sent to the President/);
  assert.deepStrictEqual(payloadOf(env, id).lines.map(l => l.gl), [PI_SE, PI_SE, PI_SE, PI_SE, PI_OS, PI_SE]);
  assert.deepStrictEqual(led().map(r => r[c.LC["GL Group"]]), Array(6).fill("Yearbook"));
  assert.deepStrictEqual(led().map(r => r[c.LC["GL Code"]]), [PI_SE, PI_SE, PI_SE, PI_SE, PI_OS, PI_SE]);
  assert.strictEqual(led()[2][c.LC["Notes"]], "my note");
  assert.ok(led().every(r => r[c.LC["Status"]] === "Pending" && r[c.LC["Claim Status"]] === "Pending President signature"));
  assert.notStrictEqual(row(env, id).pdfId, oldPdf); assert.ok(env.files[oldPdf].trashed);
  const html = pdfHtml(env, id);
  assert.ok(html.includes("72108") && !html.includes("72107"), "new PDF prints only the Yearbook codes");
  assert.ok(html.includes("SIGvp"), "VP Finance signature kept");
  assert.strictEqual(row(env, id).status, "Pending President signature");
  assert.notStrictEqual(row(env, id).token, oldTok, "old President link stops working");
  const m = env.mail.slice(n0); assert.strictEqual(m.length, 1);
  assert.strictEqual(m[0].to, "president@example.org"); assert.match(m[0].subject, /GL codes corrected - needs your signature/);
  assert.ok(m[0].htmlBody.includes("token=" + encodeURIComponent(row(env, id).token)));
  assert.throws(() => c.presidentSign(id, oldTok, SIG("pres")), /Invalid link/);
  c.presidentSign(id, row(env, id).token, SIG("pres"));
  assert.strictEqual(c.stageOf_(row(env, id).status), "staffpdf");
  // after the President signed: refused, nothing changed
  assert.throws(() => c.fixClaimGlCodes(), /only be fixed before the President signs/);
  // a mistake in the list changes nothing
  addClaim(env, "glfix002", "expense-claim");
  c.vpSignApprove("glfix002", row(env, "glfix002").token, SIG("vp"), [MH_SE]);
  vm.runInContext("delete GL_FIXES.abc12345; GL_FIXES.glfix002 = { 1: '" + PI_SE + "', 2: '" + PI_SE + "' };", c);
  assert.throws(() => c.fixClaimGlCodes(), /has no line 2/);
  vm.runInContext("GL_FIXES.glfix002 = { 1: '410-00000-00000000000000000000' };", c);
  assert.throws(() => c.fixClaimGlCodes(), /not on the GL list/);
  assert.strictEqual(payloadOf(env, "glfix002").lines[0].gl, MH_SE);
  // returned by the President: fixed, but stays with VP Finance to re-send
  c.presidentReturn("glfix002", row(env, "glfix002").token, "Wrong department");
  vm.runInContext("GL_FIXES.glfix002 = { 1: '" + PI_SE + "' };", c);
  const n1 = env.mail.length;
  assert.match(c.fixClaimGlCodes()[0], /still "Returned by President"/);
  assert.strictEqual(env.mail.length, n1, "no email"); assert.strictEqual(payloadOf(env, "glfix002").lines[0].gl, PI_SE);
  c.resendToPresident("glfix002", row(env, "glfix002").token);
  assert.strictEqual(row(env, "glfix002").status, "Pending President signature");
  const anon = setup(STAFF, { activeUser: "" });
  assert.throws(() => anon.ctx.fixClaimGlCodes(), /only be run/);
}

console.log("test_approvals: all checks passed" + (outDir ? " (samples written to " + outDir + ")" : ""));
