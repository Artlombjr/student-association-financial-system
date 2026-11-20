/**
 * NUSA Expense Claim — Google Apps Script backend
 *
 * Approval chain: requester -> VP Finance -> President -> University staff signer.
 *   The staff signer does not get a sign link: after the President signs, the packet PDF is
 *   handed to VP Finance, who sends it to the staff signer from a university address (the
 *   university mail system does not accept mail from Gmail). The signer signs the PDF itself
 *   and replies with it. checkStaffSignedPdfs() (every 10 minutes) finds the reply in the
 *   script owner's Gmail, saves the signed PDF in the claim's folder, marks the claim approved
 *   and emails it to VP Finance to forward to Accounts Payable (the system never emails AP).
 * Also keeps an Expense Ledger tab (one row per expense line) used by the budget tracker.
 * ------------------------------------------------------------------
 * Flow:
 *   1. Student opens the web app, fills the guided form, submits.
 *   2. submitClaim() saves receipts + a generated PDF packet to Drive,
 *      logs the claim to a Google Sheet, and emails the VP Finance
 *      the packet plus a one-click Approve / Send-back link.
 *   3. The President signs through a link.
 *   4. The staff signer signs the PDF and replies with it -> it is saved in the claim
 *      folder and emailed to VP Finance to send to Accounts Payable.
 *      (Or Send back -> the student gets a note and an edit link.)
 *
 * SETUP: fill in CONFIG below, then Deploy > New deployment > Web app.
 */

/* ========================= CONFIG ========================= */
const CONFIG = {
  // You — the approver. Every submission stops here first.
  VP_FINANCE_EMAIL: 'vp.finance@example.org',

  // Approval chain (in order): VP Finance signs -> President signs -> NU Staff gets
  // the PDF by email, signs it and replies with it -> the signed PDF is saved and emailed
  // to VP Finance (VP Finance sends it to Accounts Payable; the system doesn't email AP).
  // Names print on the form's "Approved by" lines; the emails route to these addresses.
  VP_FINANCE_NAME: 'Alex Rivera',
  // Gmail addresses only: university.example.edu addresses don't receive these automated emails.
  PRESIDENT_EMAIL: 'president@example.org',
  PRESIDENT_NAME: 'Jordan Blake',
  // NU Staff signer. NU email doesn't accept mail from Gmail, so the system never emails this
  // address: VP Finance gets the PDF and sends it to NU Staff from VP_FINANCE_NU_EMAIL.
  STAFF_EMAIL: 'Sam.Patel@university.example.edu',
  VP_FINANCE_NU_EMAIL: 'vpfinance@university.example.edu',
  // The signed PDF is accepted from STAFF_EMAIL, VP_FINANCE_NU_EMAIL and VP_FINANCE_EMAIL
  // (normally VP Finance forwards NU Staff's reply). Add other addresses here, comma-separated.
  STAFF_REPLY_FROM: '',
  STAFF_NAME: 'Sam Patel',
  // Positions printed under each "Approved by" line on the PDF forms.
  VP_FINANCE_TITLE: 'NUSA Vice President of Finance',
  PRESIDENT_TITLE: 'NUSA President',
  STAFF_TITLE: 'Director of Student Engagement',

  // Where claims are filed. Leave ID blank to auto-create a folder by name
  // in your Drive; or paste an existing folder's ID to use that instead.
  DRIVE_ROOT_FOLDER_NAME: 'NUSA Expense Claims',
  DRIVE_ROOT_FOLDER_ID: '',

  // The web app's STABLE /exec URL. Deploy > Manage deployments > copy the URL
  // that ends in "/exec" and paste it here. This prevents the "link expired"
  // problem: if this is blank the code falls back to ScriptApp.getService().getUrl(),
  // which — when the script is run/tested from the editor — returns the temporary
  // "/dev" URL that stops working after your editor session lapses (a couple hours).
  WEB_APP_URL: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',

  // Reminder check time, as an hour 0-23 in the script's timezone. The check runs once a day, but a reminder
  // email only goes out when a claim has waited 5, 7 and 14 days at the same step (REMINDER_DAYS).
  // Run installDailyReminderTrigger() once to schedule it.
  REMINDER_HOUR: 9,

  ORG_NAME: 'NUSA',
  BRAND_GREEN: '#0D4624',
  BRAND_CREAM: '#F1E8DE',
};
/* ========================================================== */

/* ---------- GL codes (VP Finance assigns these at approval) ----------
   Only GL codes with an allocated budget, expense lines only (revenue excluded).
   After changing this list, run refreshGlList once so the log's "GL List" tab matches. */
const GL = [
  {g:"Administration", items:[
    ["Wages - Student Work (Grants)","410-72103-50340100003802310000"],
    ["Special Events","410-72103-56025100003906310000"],
    ["Equipment Rental- IT Copier Charges","410-72103-55254100001000010000"],
    ["Office - Small Equipment","410-72103-55610100001000010000"],
    ["Hospitality / Food","410-72103-56350100001000010000"],
  ]},
  {g:"Faculty", items:[
    ["Events - Business Representative","410-72104-56025100003902210000"],
    ["Events - Education Representative","410-72104-56025100003902310000"],
    ["Events - Human Kinetics Representative","410-72104-56025100003902410000"],
    ["Events - Humanities Representative","410-72104-56025100003902510000"],
    ["Events - NATS Representative","410-72104-56025100003902610000"],
    ["Events - Nursing Representative","410-72104-56025100003902710000"],
    ["Events - SAMC Representative (MixEx/EventHosting)","410-72104-56025100001000010000"],
  ]},
  {g:"Communications", items:[
    ["Comm. - Publicity and Advertising","410-72105-55050100001000010000"],
    ["MTRL/GenSup/OfficeSuppl","410-72105-55600100001000010000"],
    ["Office - Small Equipment","410-72105-55610100001000010000"],
  ]},
  {g:"Events", items:[
    ["Events","410-72106-56025100003902110000"],
    ["SVC/ICAContract/ICA","410-72106-55275100001000010000"],
  ]},
  {g:"The Courier", items:[
    ["Wages - Student Work (Grants)","410-72107-50340100003802310000"],
    ["SVC/CommSvc/PrintPubl","410-72107-55225100001000010000"],
    ["MTRL/GenSup/OfficeSuppl","410-72107-55600100001000010000"],
    ["MTRL/GenSup/SmallEquipment","410-72107-55610100001000010000"],
    ["Food Costs","410-72107-55612100003902110000"],
    ["Special Events","410-72107-56025100003906310000"],
  ]},
  {g:"Yearbook", items:[
    ["Wages - Student Work (Grants)","410-72108-50340100003802310000"],
    ["Comm. - Printing and Publications","410-72108-55225100001000010000"],
    ["MTRL/GenSup/OfficeSuppl","410-72108-55600100001000010000"],
    ["MTRL/GenSup/SmallEquipment","410-72108-55610100001000010000"],
    ["Office - Computer Supplies","410-72108-55618100001000010000"],
    ["Special Events","410-72108-56025100003906310000"],
  ]},
  {g:"Student Lounge (The Commons)", items:[
    ["Office - Copywork","410-72109-55225100001000010000"],
    ["Office - Miscellaneous","410-72109-55600100001000010000"],
    ["Office - Eqmt & Furnishings","410-72109-55610100001000010000"],
    ["Food Costs","410-72109-55612100003902110000"],
    ["Rental Services","410-72109-56254100001000010000"],
  ]},
  {g:"Special Initiatives", items:[
    ["MixEx/EventHosting","410-72110-56025100001000010000"],
  ]},
  {g:"Proposal Committee", items:[
    ["NUSA Funding Clubs","410-72111-56025100003907010000"],
  ]},
  {g:"Capital Projects", items:[
    ["CapEquip/General (Communications)","410-73001-66000100004230010000"],
    ["CapEquip/General - Lounge Improvement","410-73001-66000100004170112510"],
  ]},
  {g:"Clubs (Fund 413)", items:[
    ["Nursing Club","413-73002-56025100003904410000"],
    ["Biology Club","413-73006-56025100003904410000"],
    ["Afro-Caribbean Club","413-73010-56025100001000010000"],
    ["The Chemistry Club","413-73016-56025100001000010000"],
    ["Games Club","413-73024-56025100003904410000"],
    ["Chinese Assoc. (CSSA) Club","413-73026-56025100001000010000"],
    ["Circus Guild Club (HKIN)","413-73041-56025100001000010000"],
    ["Tabletop Roleplaying Club (SAMC)","413-73042-56025100001000010000"],
    ["New / newly-ratified club - general (MixEx/EventHosting)","410-72111-56025100001000010000"],
  ]},
];
function glLabel_(code){for(var i=0;i<GL.length;i++){var it=GL[i].items;for(var j=0;j<it.length;j++){if(it[j][1]===code)return it[j][0];}}return '';}
function glOptionsHtml_(selected){var h='<option value="">- choose GL code -</option>';for(var i=0;i<GL.length;i++){h+='<optgroup label="'+String(GL[i].g).replace(/"/g,"&quot;")+'">';var it=GL[i].items;for(var j=0;j<it.length;j++){h+='<option value="'+it[j][1]+'"'+(selected===it[j][1]?' selected':'')+'>'+escapeHtml_(it[j][0])+'</option>';}h+='</optgroup>';}return h;}

/* ---------- Mileage (Kilometer Reimbursement) ---------- */
const MILEAGE_RATE = 0.60;
// Every km claim, whatever the department, goes to Travel-Mileage (Personal Vehicle), budgeted on the
// Student Lounge (The Commons) tab. Older claims keep the code saved on them when they were signed.
const MILEAGE_GL = '410-72103-56302100001000010000';
const MILEAGE_GL_NAME = 'Travel-Mileage (Personal Vehicle)';
const MILEAGE_GL_GROUP = 'Student Lounge (The Commons)';
function mileageGl_(p) { const l = (p && p.lines || [])[0] || {}; return l.gl || MILEAGE_GL; } // the code the claim was signed with
function isMileage_(p) { return p && p.form === 'mileage'; }



/* ---------- Web app entry point ---------- */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (e && e.parameter && e.parameter.feed === 'ledger') {
    return ledgerFeed_(e.parameter.key); // read-only CSV for the Excel tracker (V3)
  }
  if (action === 'decision') {
    return renderDecisionPage(e.parameter.id, e.parameter.token);
  }
  if (action === 'pdfpages') return pdfPagesPage_(e.parameter); // one-time maintenance page, owner only
  // Default: serve the claim form. With ?action=edit&id=..&token=.. (the link in a "sent back" email) the form
  // loads that claim itself (getClaimForEdit) so the requester can fix it and resubmit (resubmitClaim).
  return HtmlService.createHtmlOutputFromFile('Form')
    .setTitle(action === 'edit' ? 'Edit your NUSA claim' : 'NUSA Expense Claim')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* ============================================================
 *  1. SUBMISSION  (called from Form.html via google.script.run)
 * ============================================================ */
function submitClaim(payload) {
  // --- server-side re-validation (never trust the client alone) ---
  const problems = validatePayload(payload);
  if (problems.length) {
    throw new Error('This claim is missing: ' + problems.join(', '));
  }

  const id = Utilities.getUuid().slice(0, 8);
  const token = Utilities.getUuid();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  payload.submittedAt = new Date().toISOString(); // the PDF keeps this date through every re-signing
  const payeeName = ((payload.first || '') + ' ' + (payload.last || '')).trim();

  // --- Drive folder for this claim ---
  const root = getRootFolder_();
  const folder = root.createFolder(payload.last + ', ' + dateStr + ' (' + id + ')');

  // --- save every uploaded file to the claim folder ---
  receiptSources_(payload).forEach(function (r) { folder.createFile(dataUrlToBlob_(r.fileData, r.fileName)); });
  // --- split into what embeds in the packet vs. what rides as attachments (size-budgeted) ---
  const split = packetReceipts_(payload);
  const imageReceipts = split.embed;   // {name, dataUrl}
  let pdfReceipts = split.attach;      // blobs: PDFs + any images beyond the embed budget

  // --- build the packet PDF; if it is too large to merge, fall back to a form-only PDF ---
  const formLabel = formLabel_(payload);
  if (isMileage_(payload)) { (payload.lines || []).forEach(function (l) { l.gl = MILEAGE_GL; l.glName = MILEAGE_GL_NAME; }); }
  const built = htmlToPdf_(payload, id, imageReceipts, formLabel + ' - ' + payload.last + ' - ' + dateStr + '.pdf');
  const packetPdf = built.pdf;
  if (!built.merged) { pdfReceipts = embeddedAsBlobs_(imageReceipts).concat(pdfReceipts); }
  const pdfFile = folder.createFile(packetPdf);

  // --- save the full submission as JSON so the PDF can be re-stamped
  //     with the VP Finance and President signatures at each stage ---
  const jsonFile = folder.createFile(
    Utilities.newBlob(JSON.stringify(payload), 'application/json', 'claim.json'));

  // --- log to the ledger sheet ---
  const t = payload.totals || {};
  logClaim_({
    id: id, token: token, status: 'Pending VP Finance signature',
    payTo: payload.payTo, name: payeeName, email: payload.email,
    total: t.amt || 0, gst: t.gst || 0, hst: t.hst || 0, form: payload.form || 'expense-claim',
    folderUrl: folder.getUrl(), pdfId: pdfFile.getId(), jsonId: jsonFile.getId(),
  });
  sharePacketWith_(pdfFile.getId(), CONFIG.VP_FINANCE_EMAIL); // VP Finance can open the folder before signing

  // --- email VP Finance to sign & approve (stage 1 of the chain) ---
  MailApp.sendEmail({
    to: CONFIG.VP_FINANCE_EMAIL,
    subject: '[' + formLabel + '] ' + payeeName + ' - $' + money_(t.amt) + ' - needs your signature',
    htmlBody: stageEmailHtml_('vp1', { id: id, token: token, name: payeeName, total: t.amt, folderUrl: folder.getUrl(), form: payload.form }, pdfReceipts.length),
    attachments: emailAttachments_(packetPdf, pdfReceipts),
    name: CONFIG.ORG_NAME + ' Finance',
  });

  return {
    ok: true,
    id: id,
    reimbursement: money_(t.amt),
    folderName: payload.last + ', ' + dateStr,
    pdfReceiptsSeparate: pdfReceipts.length,
  };
}


/* ============================================================
 *  2. APPROVAL CHAIN  (decision pages opened from the email links)
 *     Stages (V5):  vp1 -> president -> staffpdf -> done (signed PDF saved + emailed to VP Finance)
 *     The President can return a claim to VP Finance with a note (vpreturned).
 *     Claims routed before V5 can still be at: staff (NU Staff sign link), staffreturned,
 *     vpfinal (VP final approval) — those pages keep working until they're finished.
 * ============================================================ */
const STAFF_PDF_STATUS = 'Pending NU Staff signed PDF';
function stageOf_(status) {
  status = String(status || '');
  if (status.indexOf('Pending VP Finance signature') === 0) return 'vp1';
  if (status.indexOf('Pending President') === 0) return 'president';
  if (status.indexOf(STAFF_PDF_STATUS) === 0) return 'staffpdf';
  if (status.indexOf('Pending NU Staff') === 0) return 'staff';
  if (status.indexOf('Pending VP Finance final') === 0) return 'vpfinal';
  if (status.indexOf('Returned by President') === 0) return 'vpreturned';
  if (status.indexOf('Returned by NU Staff') === 0) return 'staffreturned';
  return 'done';
}

function renderDecisionPage(id, token) {
  const row = findClaim_(id);
  if (!row) {
    return infoPage_('Claim not found', 'We could not find a claim for this link. It may have been removed. Please contact NUSA Finance.');
  }
  if (row.token !== token) {
    // The link is genuine but from an OLDER email for this claim — a later step
    // rotated the token. Tell the recipient exactly that, plus the current status,
    // so an out-of-date link is never mistaken for a broken one.
    return infoPage_('This link was replaced by a newer email',
      'This claim has moved forward since this email was sent, so this particular link is no longer the active one. '
      + 'Current status: <b>' + escapeHtml_(row.status) + '</b>.<br><br>'
      + 'Please open the most recent NUSA Finance email for this claim and use the button there.');
  }
  let payload = {};
  try { payload = loadPayload_(row.jsonId) || {}; } catch (e) { payload = {}; }
  return HtmlService.createHtmlOutput(decisionPageHtml_(row, stageOf_(row.status), payload))
    .setTitle('NUSA Expense Claim - ' + row.name)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// A simple branded message page (used for not-found / superseded-link notices).
function infoPage_(title, bodyHtml) {
  const g = CONFIG.BRAND_GREEN, cream = CONFIG.BRAND_CREAM;
  const html = '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:' + cream + ';color:#1B2A20;margin:0;padding:24px}'
    + '.card{max-width:540px;margin:40px auto;background:#FBF7F1;border:1px solid #DED2C1;border-radius:14px;padding:24px}'
    + 'h1{color:' + g + ';font-size:20px;margin:0 0 10px}p{font-size:14px;line-height:1.55;color:#3a463c}</style></head><body>'
    + '<div class="card"><h1>' + escapeHtml_(title) + '</h1><p>' + bodyHtml + '</p></div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('NUSA Finance').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Stage 1: VP Finance draws a signature and approves -> route to the President.
function vpSignApprove(id, token, signatureDataUrl, gls) {
  const row = guard_(id, token, 'vp1');
  requireSignature_(signatureDataUrl);
  const payload = loadPayload_(row.jsonId);
  if (isMileage_(payload)) {
    (payload.lines || []).forEach(function (l) { l.gl = MILEAGE_GL; l.glName = MILEAGE_GL_NAME; });
  } else {
    gls = gls || [];
    (payload.lines || []).forEach(function (l, i) {
      if (!gls[i]) throw new Error('Please assign a GL code to every expense line.');
      l.gl = gls[i];
      l.glName = glLabel_(gls[i]);
    });
  }
  payload.vpSignature = signatureDataUrl;
  payload.vpName = CONFIG.VP_FINANCE_NAME;
  savePayload_(row.jsonId, payload);
  safeLedger_(function () { writeLedgerRows_(row, payload, 'Pending President signature'); }); // Expense Ledger (V3)
  rebuildAndSavePdf_(payload, row);
  sharePacketWithPresident_(row.pdfId); // let her open the packet from her email/decision page
  const nt = rotateToken_(id);
  updateClaimStatus_(id, 'Pending President signature', CONFIG.VP_FINANCE_EMAIL);
  MailApp.sendEmail({
    to: CONFIG.PRESIDENT_EMAIL,
    subject: '[' + formLabel_(row) + '] ' + row.name + ' - $' + money_(row.total) + ' - needs your signature',
    htmlBody: stageEmailHtml_('president', { id: id, token: nt, name: row.name, total: row.total, folderUrl: row.folderUrl, form: row.form }, 0),
    attachments: capBlobs_(nonImageReceiptBlobs_(payload)),
    name: CONFIG.ORG_NAME + ' Finance',
  });
  return { ok: true };
}

// Stage 2: President draws a signature and approves -> on to the NU Staff signer.
function presidentSign(id, token, signatureDataUrl) {
  const row = guard_(id, token, 'president');
  requireSignature_(signatureDataUrl);
  requireStaffEmail_();
  const payload = loadPayload_(row.jsonId);
  payload.presidentSignature = signatureDataUrl;
  payload.presidentName = CONFIG.PRESIDENT_NAME;
  savePayload_(row.jsonId, payload);
  rebuildAndSavePdf_(payload, row);
  routeToStaff_(row, payload, CONFIG.PRESIDENT_EMAIL);
  return { ok: true };
}

// Stage 3: NU Staff draws a signature and approves -> back to VP Finance for final approval.
function staffSign(id, token, signatureDataUrl) {
  const row = guard_(id, token, 'staff');
  requireSignature_(signatureDataUrl);
  const payload = loadPayload_(row.jsonId);
  payload.staffSignature = signatureDataUrl;
  payload.staffName = CONFIG.STAFF_NAME;
  savePayload_(row.jsonId, payload);
  rebuildAndSavePdf_(payload, row);
  const nt = rotateToken_(id);
  updateClaimStatus_(id, 'Pending VP Finance final approval', CONFIG.STAFF_EMAIL);
  MailApp.sendEmail({
    to: CONFIG.VP_FINANCE_EMAIL,
    subject: '[' + formLabel_(row) + '] ' + row.name + ' - $' + money_(row.total) + ' - ready for final approval',
    htmlBody: stageEmailHtml_('vpfinal', { id: id, token: nt, name: row.name, total: row.total, folderUrl: row.folderUrl, form: row.form }, 0),
    attachments: capBlobs_(nonImageReceiptBlobs_(payload)),
    name: CONFIG.ORG_NAME + ' Finance',
  });
  return { ok: true };
}

// NU Staff returns the claim to VP Finance with a note (no signature).
function staffReturn(id, token, note) {
  const row = guard_(id, token, 'staff');
  if (!note || !note.trim()) throw new Error('Please add a note for VP Finance.');
  const payload = loadPayload_(row.jsonId);
  payload.staffReturnNote = note;
  savePayload_(row.jsonId, payload);
  const nt = rotateToken_(id);
  updateClaimStatus_(id, 'Returned by NU Staff', CONFIG.STAFF_EMAIL);
  MailApp.sendEmail({
    to: CONFIG.VP_FINANCE_EMAIL,
    subject: '[' + formLabel_(row) + '] ' + row.name + ' - returned by NU Staff',
    htmlBody: '<p>' + escapeHtml_(CONFIG.STAFF_NAME) + ' (NU Staff) returned this claim (<b>$' + money_(row.total) + '</b>, ' + escapeHtml_(row.name) + ') with a note:</p>'
      + '<blockquote style="border-left:3px solid ' + CONFIG.BRAND_GREEN + ';padding-left:12px">' + escapeHtml_(note).replace(/\n/g, '<br>') + '</blockquote>'
      + '<p><a href="' + decisionUrl_(id, nt) + '">Open the claim to re-send or send back</a></p>',
    name: CONFIG.ORG_NAME + ' Finance',
  });
  return { ok: true };
}

// VP Finance (re-)sends a claim to NU Staff: after NU Staff returned it (pre-V5 link),
// for claims that reached final approval before V4 without a NU Staff signature, and
// while waiting for NU Staff's signed PDF (e.g. they lost the email).
function resendToStaff(id, token) {
  acquireLock_();
  const row = findClaim_(id);
  if (!row || row.token !== token) throw new Error('Invalid link.');
  const payload = loadPayload_(row.jsonId);
  const st = stageOf_(row.status);
  if (!(st === 'staffreturned' || st === 'staffpdf' || (st === 'vpfinal' && !payload.staffSignature))) {
    throw new Error('This step has already been completed (' + row.status + ').');
  }
  requireStaffEmail_();
  routeToStaff_(row, payload, CONFIG.VP_FINANCE_EMAIL);
  return { ok: true };
}

// V5: NU's email doesn't accept mail from Gmail, so the system can't email NU Staff directly.
// It emails VP Finance the PDF with a ready-to-send message; VP Finance sends it to NU Staff from
// their NU email, and forwards NU Staff's signed reply back to the system's Gmail.
function routeToStaff_(row, payload, who) {
  const nt = rotateToken_(row.id);
  updateClaimStatus_(row.id, STAFF_PDF_STATUS, who);
  sendStaffPdfEmail_(row, payload, false, nt);
}

// "(Claim <id>)" in the subject (or the message) is how the signed PDF is matched to its claim.
function claimTag_(row) { return '(Claim ' + row.id + ')'; }
function staffPdfSubject_(row) {
  return formLabel_(row) + ' - ' + row.name + ' - $' + money_(row.total) + ' - please sign ' + claimTag_(row);
}
// The hand-off email to VP Finance: the PDF, plus the subject and message to send NU Staff.
// token (optional): a fresh claim link for re-send / send back; reminders go without it.
// days (optional, reminders): how long it has been waiting, for the subject.
function sendStaffPdfEmail_(row, payload, isReminder, token, days) {
  const pdf = DriveApp.getFileById(row.pdfId).getBlob();
  const extras = nonImageReceiptBlobs_(payload || {});
  const attachments = emailAttachments_(pdf, extras);
  const staffFirst = escapeHtml_(String(CONFIG.STAFF_NAME).split(' ')[0]);
  const box = 'border:1px solid #CBBCA6;border-radius:8px;background:#FBFAF7;padding:10px 12px;margin:6px 0 14px';
  const msgForStaff = 'Hi ' + staffFirst + ',<br><br>'
    + 'Attached is a NUSA ' + formLabelLc_(row) + ' (' + escapeHtml_(row.name) + ', $' + money_(row.total) + ') signed by the NUSA VP Finance and President. '
    + 'Could you please sign it on the &ldquo;Approved by &hellip; ' + escapeHtml_(CONFIG.STAFF_TITLE) + '&rdquo; line '
    + '(for example with Adobe Acrobat &ldquo;Fill &amp; Sign&rdquo;, or print, sign and scan) and reply to this email with the signed PDF attached?<br><br>'
    + 'Thank you,<br>' + escapeHtml_(CONFIG.VP_FINANCE_NAME) + '<br>' + escapeHtml_(CONFIG.VP_FINANCE_TITLE) + '<br>'
    + '<span style="color:#666">' + escapeHtml_(claimTag_(row)) + '</span>';
  MailApp.sendEmail({
    to: CONFIG.VP_FINANCE_EMAIL,
    subject: (isReminder ? '[Reminder] ' : '') + '[' + formLabel_(row) + '] ' + row.name + ' - $' + money_(row.total)
      + (isReminder && days > 0 ? ' - not signed by ' + CONFIG.STAFF_NAME + ' for the past ' + days + (days === 1 ? ' day' : ' days') : '')
      + ' - send to ' + CONFIG.STAFF_NAME + ' to sign ' + claimTag_(row),
    htmlBody: '<div style="font-family:sans-serif;max-width:600px">'
      + '<h2 style="color:' + CONFIG.BRAND_GREEN + '">' + (isReminder ? 'Still waiting for ' + staffFirst + '&#39;s signed PDF' : 'Send this PDF to ' + escapeHtml_(CONFIG.STAFF_NAME) + ' to sign') + '</h2>'
      + '<p>The VP Finance and President have signed. Send the attached PDF to <b>' + escapeHtml_(CONFIG.STAFF_NAME) + '</b> ('
      + escapeHtml_(CONFIG.STAFF_TITLE) + ') from your NU email' + (isReminder ? ', if you haven&#39;t yet' : '') + ':</p>'
      + '<ol style="line-height:1.7">'
      + '<li>Download the attached PDF' + (attachments.length > 1 ? ' (the other attachments are receipts that aren&#39;t inside the PDF; ' + staffFirst + ' doesn&#39;t need them)' : '') + '.</li>'
      + '<li>In Outlook (' + escapeHtml_(CONFIG.VP_FINANCE_NU_EMAIL) + '), write a new email to <b>' + escapeHtml_(CONFIG.STAFF_EMAIL) + '</b> and attach the PDF.</li>'
      + '<li>Subject (copy and paste):<div style="' + box + ';font-family:monospace">' + escapeHtml_(staffPdfSubject_(row)) + '</div></li>'
      + '<li>Message (copy and paste):<div style="' + box + '">' + msgForStaff + '</div></li></ol>'
      + '<p><b>When ' + staffFirst + ' replies with the signed PDF:</b> forward the reply from Outlook to <b>' + escapeHtml_(scriptMailbox_()) + '</b>. '
      + 'Keep <b>' + escapeHtml_(claimTag_(row)) + '</b> in the subject (&ldquo;RE:&rdquo; / &ldquo;FW:&rdquo; in front are fine). '
      + 'Within about 10 minutes the signed PDF is saved in the claim&#39;s folder and emailed to you here, ready for Accounts Payable.</p>'
      + (token ? '<p style="font-size:13px"><a href="' + decisionUrl_(row.id, token) + '">Open the claim</a> (send this email again, or send the claim back to the student)</p>' : '')
      + (row.folderUrl ? '<p style="font-size:12px;color:#666"><a href="' + row.folderUrl + '">Open the Drive folder (all receipts)</a></p>' : '')
      + '</div>',
    attachments: attachments,
    name: CONFIG.ORG_NAME + ' Finance',
  });
}
// The Gmail account the script runs as: NU Staff's replies land here.
function scriptMailbox_() {
  try { return String(Session.getEffectiveUser().getEmail() || ''); } catch (e) { return ''; }
}

// Stop before anything is saved if the NU Staff email was never filled in.
function requireStaffEmail_() {
  const e = String(CONFIG.STAFF_EMAIL || '').trim();
  if (!e || e.indexOf('@') < 0) throw new Error('The NU Staff email is not set up yet (CONFIG.STAFF_EMAIL). Please tell NUSA Finance.');
}

// President returns the claim to VP Finance with a note (no signature).
function presidentReturn(id, token, note) {
  const row = guard_(id, token, 'president');
  if (!note || !note.trim()) throw new Error('Please add a note for VP Finance.');
  const payload = loadPayload_(row.jsonId);
  payload.presidentReturnNote = note;
  savePayload_(row.jsonId, payload);
  const nt = rotateToken_(id);
  updateClaimStatus_(id, 'Returned by President', CONFIG.PRESIDENT_EMAIL);
  MailApp.sendEmail({
    to: CONFIG.VP_FINANCE_EMAIL,
    subject: '[' + formLabel_(row) + '] ' + row.name + ' - returned by the President',
    htmlBody: '<p>The President returned this claim (<b>$' + money_(row.total) + '</b>, ' + escapeHtml_(row.name) + ') with a note:</p>'
      + '<blockquote style="border-left:3px solid ' + CONFIG.BRAND_GREEN + ';padding-left:12px">' + escapeHtml_(note).replace(/\n/g, '<br>') + '</blockquote>'
      + '<p><a href="' + decisionUrl_(id, nt) + '">Open the claim to re-send or send back</a></p>',
    name: CONFIG.ORG_NAME + ' Finance',
  });
  return { ok: true };
}

// VP Finance re-sends a President-returned claim back to the President.
function resendToPresident(id, token) {
  const row = guard_(id, token, 'vpreturned');
  const payload = loadPayload_(row.jsonId);
  sharePacketWithPresident_(row.pdfId); // (re-)grant view access on the current packet
  const nt = rotateToken_(id);
  updateClaimStatus_(id, 'Pending President signature', CONFIG.VP_FINANCE_EMAIL);
  MailApp.sendEmail({
    to: CONFIG.PRESIDENT_EMAIL,
    subject: '[' + formLabel_(row) + '] ' + row.name + ' - $' + money_(row.total) + ' - needs your signature',
    htmlBody: stageEmailHtml_('president', { id: id, token: nt, name: row.name, total: row.total, folderUrl: row.folderUrl, form: row.form }, 0),
    attachments: capBlobs_(nonImageReceiptBlobs_(payload)),
    name: CONFIG.ORG_NAME + ' Finance',
  });
  return { ok: true };
}

// Stage 4: VP Finance gives final approval -> the final signed packet is emailed to
// VP Finance, who submits it to Accounts Payable (the system does not email AP).
const FINAL_STATUS = 'Approved - final PDF sent to VP Finance';
function finalApprove(id, token) {
  const row = guard_(id, token, 'vpfinal');
  const payload = loadPayload_(row.jsonId);
  if (!payload.staffSignature) throw new Error('This claim still needs the NU Staff signature. Use "Send to NU Staff" first.');
  const pdf = DriveApp.getFileById(row.pdfId).getBlob();
  const extras = nonImageReceiptBlobs_(payload);
  const attachments = emailAttachments_(pdf, extras);
  // 1) The final packet to VP Finance — the one email that must go out, so it goes first.
  MailApp.sendEmail({
    to: CONFIG.VP_FINANCE_EMAIL,
    subject: '[' + formLabel_(row) + ' - FINAL, ready for Accounts Payable] ' + row.name + ' - $' + money_(row.total),
    htmlBody: finalPacketEmailHtml_(row, attachments.length - 1, extras.length - (attachments.length - 1)),
    attachments: attachments,
    name: CONFIG.ORG_NAME + ' Finance',
  });
  // 2) Mark it approved right away, so a problem below can't leave it pending (and a second click can't re-send it).
  updateClaimStatus_(id, FINAL_STATUS, CONFIG.VP_FINANCE_EMAIL);
  moveClaimFolder_(row.pdfId, 'Approved');
  // 3) The requester's copy. A bad address must not undo the approval.
  if (row.email) {
    try {
      MailApp.sendEmail({
        to: row.email,
        subject: 'Your NUSA ' + formLabelLc_(row) + ' was approved',
        htmlBody: '<p>Hi ' + escapeHtml_(row.name.split(' ')[0]) + ',</p>'
          + '<p>Your ' + formLabelLc_(row) + ' for <b>$' + money_(row.total) + '</b> has been fully approved (VP Finance, President and NU Staff). '
          + 'NUSA Finance will now submit it to NU Accounts Payable for payment. Cheques are usually issued within two weeks.</p>'
          + '<p>A copy of the fully-signed ' + formLabelLc_(row) + ' is attached for your records.</p><p>- NUSA Finance</p>',
        attachments: attachments,
        name: CONFIG.ORG_NAME + ' Finance',
      });
    } catch (err) {
      console.warn('Approved, but the requester email failed for claim ' + id + ': ' + err);
      return { ok: true, requesterEmailFailed: true };
    }
  }
  return { ok: true };
}

/* ============================================================
 *  V5: NU Staff's signed PDF comes back by email
 *
 *  checkStaffSignedPdfs() looks in the script owner's Gmail for emails that carry a PDF and
 *  "(Claim <id>)" in the subject (or the message), from VP Finance (normally forwarding NU Staff's
 *  reply from VP_FINANCE_NU_EMAIL), NU Staff, or STAFF_REPLY_FROM. For a claim waiting on NU Staff it saves the PDF in the claim's folder,
 *  marks the claim approved and emails the signed PDF to VP Finance (and a copy to the
 *  requester). Each handled email thread gets the Gmail label "NUSA-signed-PDF".
 *
 *  SETUP: run installSignedPdfWatcher() ONCE from the editor (checks every 10 minutes).
 * ============================================================ */
const SIGNED_STATUS = 'Approved - signed PDF from NU Staff saved';
const SIGNED_PDF_LABEL = 'NUSA-signed-PDF';

function checkStaffSignedPdfs() {
  // The trigger must be able to call this, so it can't be owner-only; it runs at most every
  // 4 minutes so a web visitor can't use it to burn through Gmail quota.
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('LAST_SIGNED_PDF_CHECK') || 0);
  if (Date.now() - last < 240000) { console.log('Checked less than 4 minutes ago; skipped.'); return 0; }
  props.setProperty('LAST_SIGNED_PDF_CHECK', String(Date.now()));
  acquireLock_(); // one pass at a time, and no claim click mid-save

  const me = scriptMailbox_().toLowerCase();
  const allowed = emailList_([CONFIG.STAFF_EMAIL, CONFIG.VP_FINANCE_NU_EMAIL, CONFIG.STAFF_REPLY_FROM, CONFIG.VP_FINANCE_EMAIL].join(','))
    .filter(function (e) { return e !== me; }); // never our own outgoing email (it carries the unsigned PDF)
  const label = GmailApp.getUserLabelByName(SIGNED_PDF_LABEL) || GmailApp.createLabel(SIGNED_PDF_LABEL);
  if (!allowed.length) return 0;
  const threads = GmailApp.search('has:attachment -label:' + SIGNED_PDF_LABEL + ' newer_than:90d from:(' + allowed.join(' OR ') + ')', 0, 50);
  let saved = 0;
  threads.forEach(function (thread) {
    const msgs = thread.getMessages();
    for (let i = msgs.length - 1; i >= 0; i--) {         // newest reply first
      const m = msgs[i];
      const from = senderAddress_(m.getFrom());
      if (!from || from === me || allowed.indexOf(from) < 0) continue;
      // The tag is in the subject; if a subject was retyped, look for it in the message (it's in the text sent to NU Staff).
      const tag = /\(Claim ([^)\s]+)\)/;
      const idMatch = String(m.getSubject() || '').match(tag) || String(m.getPlainBody() || '').match(tag);
      if (!idMatch) continue;
      const pdfs = m.getAttachments({ includeInlineImages: false }).filter(function (a) {
        return String(a.getContentType()).toLowerCase() === 'application/pdf' || /\.pdf$/i.test(a.getName());
      });
      if (!pdfs.length) continue;
      let res;
      try { res = saveStaffSignedPdf_(idMatch[1], pdfs[0], from); }
      catch (err) { console.warn('Claim ' + idMatch[1] + ': not saved, will retry next run: ' + err); break; }
      if (res.ok) saved++;
      else notifyUnmatchedPdf_(idMatch[1], from, res.reason);
      thread.addLabel(label); // handled either way; never processed twice
      break;
    }
  });
  console.log('Signed PDFs saved: ' + saved);
  return saved;
}

// Save the signed PDF for a claim waiting on NU Staff, approve it and email VP Finance.
// (Called with the script lock held, from checkStaffSignedPdfs.)
function saveStaffSignedPdf_(id, blob, from) {
  const row = findClaim_(id);
  if (!row) return { ok: false, reason: 'no claim with this ID was found' };
  const st = stageOf_(row.status);
  if (st !== 'staffpdf' && st !== 'staff') return { ok: false, reason: 'the claim is not waiting for NU Staff (status: ' + row.status + ')' };
  const payload = loadPayload_(row.jsonId);
  const unsigned = DriveApp.getFileById(row.pdfId);
  const folder = unsigned.getParents().next();
  const name = String(unsigned.getName()).replace(/\.pdf$/i, '') + ' - signed by NU Staff.pdf';
  const file = folder.createFile(blob.copyBlob().setName(name)); // the unsigned packet stays in the folder too
  row.pdfId = file.getId();

  // 1) The signed PDF to VP Finance first: if this fails, nothing is marked and the next run retries.
  const extras = nonImageReceiptBlobs_(payload);
  const attachments = emailAttachments_(file.getBlob(), extras);
  try {
    MailApp.sendEmail({
      to: CONFIG.VP_FINANCE_EMAIL,
      subject: '[' + formLabel_(row) + ' - FINAL, ready for Accounts Payable] ' + row.name + ' - $' + money_(row.total),
      htmlBody: finalPacketEmailHtml_(row, attachments.length - 1, extras.length - (attachments.length - 1), name),
      attachments: attachments,
      name: CONFIG.ORG_NAME + ' Finance',
    });
  } catch (err) {
    try { file.setTrashed(true); } catch (e) { /* keep going */ }
    throw err;
  }
  // 2) Record it and mark the claim approved.
  payload.staffName = CONFIG.STAFF_NAME;
  payload.staffSignedPdfId = file.getId();
  payload.staffSignedFrom = from;
  payload.staffSignedAt = new Date().toISOString();
  savePayload_(row.jsonId, payload);
  updateClaimPdf_(row.id, file.getId());
  rotateToken_(row.id); // old links now only show the status
  updateClaimStatus_(row.id, SIGNED_STATUS, from);
  moveClaimFolder_(row.pdfId, 'Approved');
  // 3) The requester's copy. A bad address must not undo the approval.
  if (row.email) {
    try {
      MailApp.sendEmail({
        to: row.email,
        subject: 'Your NUSA ' + formLabelLc_(row) + ' was approved',
        htmlBody: '<p>Hi ' + escapeHtml_(String(row.name).split(' ')[0]) + ',</p>'
          + '<p>Your ' + formLabelLc_(row) + ' for <b>$' + money_(row.total) + '</b> has been fully approved (VP Finance, President and NU Staff). '
          + 'NUSA Finance will now submit it to NU Accounts Payable for payment. Cheques are usually issued within two weeks.</p>'
          + '<p>A copy of the fully-signed ' + formLabelLc_(row) + ' is attached for your records.</p><p>- NUSA Finance</p>',
        attachments: attachments,
        name: CONFIG.ORG_NAME + ' Finance',
      });
    } catch (err) {
      console.warn('Saved, but the requester email failed for claim ' + id + ': ' + err);
    }
  }
  return { ok: true };
}

// A PDF arrived that couldn't be matched: tell VP Finance once (the thread is then labelled).
function notifyUnmatchedPdf_(id, from, reason) {
  try {
    MailApp.sendEmail({
      to: CONFIG.VP_FINANCE_EMAIL,
      subject: '[NUSA Finance] Signed PDF for claim ' + id + ' was not saved',
      htmlBody: '<p>A PDF arrived from ' + escapeHtml_(from) + ' for claim <b>' + escapeHtml_(String(id)) + '</b>, but it was not saved because '
        + escapeHtml_(reason) + '.</p><p>Look for the email in ' + escapeHtml_(scriptMailbox_()) + ' (Gmail label ' + SIGNED_PDF_LABEL + ') if it needs to be filed by hand.</p>',
      name: CONFIG.ORG_NAME + ' Finance',
    });
  } catch (err) { console.warn('Could not send the unmatched-PDF notice: ' + err); }
}

function emailList_(s) {
  return String(s || '').split(',').map(function (e) { return e.trim().toLowerCase(); }).filter(function (e) { return e.indexOf('@') > 0; });
}
// "Sam Patel <sam@example.com>" -> "sam@example.com"
function senderAddress_(from) {
  const s = String(from || '');
  const m = s.match(/<([^>]+)>/);
  return (m ? m[1] : s).trim().toLowerCase();
}

// Run ONCE from the editor: checks for NU Staff's signed PDFs every 10 minutes.
function installSignedPdfWatcher() {
  ownerOnly_();
  removeSignedPdfWatcher();
  ScriptApp.newTrigger('checkStaffSignedPdfs').timeBased().everyMinutes(10).create();
  console.log('Watching ' + scriptMailbox_() + ' for signed PDFs every 10 minutes.');
}
function removeSignedPdfWatcher() {
  ownerOnly_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkStaffSignedPdfs') ScriptApp.deleteTrigger(t);
  });
}

// VP Finance sends the claim back to the student (available at vp1, vpfinal, while waiting
// for NU Staff's signed PDF, and after the President or NU Staff returned it).
function sendBackClaim(id, token, note) {
  acquireLock_();
  const row = findClaim_(id);
  if (!row || row.token !== token) throw new Error('Invalid link.');
  const st = stageOf_(row.status);
  if (['vp1', 'vpfinal', 'vpreturned', 'staffreturned', 'staffpdf'].indexOf(st) < 0) throw new Error('This step has already been completed (' + row.status + ').');
  if (!note || !note.trim()) throw new Error('Please add a note explaining what needs fixing.');
  // Keep the note with the claim (the edit page shows it) and give the requester a private edit link.
  try { const p = loadPayload_(row.jsonId); p.sentBackNote = note.trim(); p.sentBackAt = new Date().toISOString(); savePayload_(row.jsonId, p); }
  catch (e) { console.warn('Could not save the send-back note: ' + e); }
  const nt = rotateToken_(id); // the VP's old links stop working; this token is now the edit link's
  MailApp.sendEmail({
    to: row.email,
    cc: CONFIG.VP_FINANCE_EMAIL,
    subject: 'Your NUSA ' + formLabelLc_(row) + ' needs a change',
    htmlBody: sendBackEmailHtml_(row, note, editUrl_(id, nt)),
    name: CONFIG.ORG_NAME + ' Finance',
  });
  updateClaimStatus_(id, 'Sent back for changes', CONFIG.VP_FINANCE_EMAIL);
  moveClaimFolder_(row.pdfId, 'Sent back');
  return { ok: true };
}

/* ---------- Edit + resubmit a claim that was sent back ----------
 * The "needs a change" email links to editUrl_ (?action=edit&id=..&token=..). That opens the form with the
 * requester's own claim filled in (details, lines, receipts, signature) and VP Finance's note on top.
 * Resubmitting keeps the same claim ID and Drive folder: new receipts are added, a new PDF is made (the
 * sent-back one is kept, renamed), the claim goes back to "Pending VP Finance signature" and VP Finance
 * gets the sign email again. The GL codes VP Finance already picked are pre-selected on the same lines.
 * The link works until the claim is resubmitted (the token changes then). */
function editUrl_(id, token) {
  return webAppUrl_() + '?action=edit&id=' + encodeURIComponent(id) + '&token=' + encodeURIComponent(token);
}
function isSentBack_(status) { return String(status || '').indexOf('Sent back') === 0; }
function sendBackEmailHtml_(row, note, link) {
  const lbl = formLabelLc_(row);
  return '<div style="font-family:sans-serif;max-width:560px">'
    + '<p>Hi ' + escapeHtml_(String(row.name || '').split(' ')[0]) + ',</p>'
    + '<p>Thanks for your ' + lbl + ' (<b>$' + money_(row.total) + '</b>). Before it can proceed, please fix the following:</p>'
    + (note ? '<blockquote style="border-left:3px solid ' + CONFIG.BRAND_GREEN + ';padding-left:12px;color:#333">' + escapeHtml_(note).replace(/\n/g, '<br>') + '</blockquote>' : '')
    + '<p>Open your claim with the button below. Everything you entered (and your receipts) is already filled in: change what\'s needed and press <b>Resubmit claim</b>. No need to start a new claim.</p>'
    + '<p style="margin:22px 0"><a href="' + link + '" style="background:' + CONFIG.BRAND_GREEN + ';color:#fff;text-decoration:none;padding:12px 22px;border-radius:8px;font-weight:bold">Edit your claim &rarr;</a></p>'
    + '<p style="font-size:12px;color:#666">This link is only for you and only for this claim. It stops working once you resubmit.</p>'
    + '<p>- NUSA Finance</p></div>';
}
function editRow_(id, token) {
  const row = findClaim_(id);
  if (!row || !token || row.token !== token) throw new Error('This link is not valid any more. If you already resubmitted this claim, there is nothing else to do; otherwise ask VP Finance for a new link.');
  if (!isSentBack_(row.status)) throw new Error('This claim was already resubmitted (status: ' + row.status + '). There is nothing else to do.');
  return row;
}

// Called by the form when it is opened from an edit link: the claim's details for the form to fill in.
function getClaimForEdit(id, token) {
  const row = editRow_(id, token);
  const p = loadPayload_(row.jsonId);
  const out = { id: String(row.id), form: p.form || 'expense-claim', note: p.sentBackNote || '' };
  ['payTo', 'first', 'last', 'email', 'studentId', 'company', 'contact', 'address', 'city', 'country', 'printName', 'delivery']
    .forEach(function (k) { out[k] = p[k] === undefined || p[k] === null ? '' : p[k]; });
  out.usd = !!p.usd;
  out.signature = isImageDataUrl_(p.signature) ? p.signature : null;
  out.lines = (p.lines || []).map(function (l) {
    return isMileage_(p)
      ? { date: l.date || '', km: l.km === undefined ? '' : l.km, destination: l.destination || '', purpose: l.purpose || '', parking: l.parking === undefined ? '' : l.parking }
      : { date: l.date || '', desc: l.desc || '', amount: l.amount === undefined ? '' : l.amount, gst: l.gst === undefined ? '' : l.gst,
          hst: l.hst === undefined ? '' : l.hst, fileName: l.fileName || '', fileType: l.fileType || '', fileData: l.fileData || '',
          filePages: Array.isArray(l.filePages) ? l.filePages : [] };
  });
  out.attachments = (p.attachments || []).map(function (a) { return { fileName: a.fileName || '', fileType: a.fileType || '', fileData: a.fileData || '', pages: Array.isArray(a.pages) ? a.pages : [] }; });
  return out;
}

// Called by the form: saves the edited claim under the same ID and sends it to VP Finance to sign again.
function resubmitClaim(id, token, payload) {
  acquireLock_();
  const row = editRow_(id, token);
  const problems = validatePayload(payload);
  if (problems.length) throw new Error('This claim is missing: ' + problems.join(', '));
  const old = loadPayload_(row.jsonId);
  if ((payload.form || 'expense-claim') !== (old.form || 'expense-claim')) throw new Error('A claim can\'t change type when it is edited. Please start a new claim instead.');
  const now = new Date();
  payload.submittedAt = now.toISOString();
  payload.firstSubmittedAt = old.firstSubmittedAt || old.submittedAt || (row.timestamp instanceof Date ? row.timestamp.toISOString() : '');
  payload.resubmissions = (old.resubmissions || 0) + 1;
  payload.lastSentBackNote = old.sentBackNote || '';
  if (isMileage_(payload)) {
    (payload.lines || []).forEach(function (l) { l.gl = MILEAGE_GL; l.glName = MILEAGE_GL_NAME; });
  } else {
    (payload.lines || []).forEach(function (l, i) { const o = (old.lines || [])[i]; if (o && o.gl && glLabel_(o.gl)) { l.gl = o.gl; l.glName = glLabel_(o.gl); } });
  }
  // Files: only receipts that weren't in the claim before are added to its folder.
  const oldPdf = DriveApp.getFileById(row.pdfId);
  const folder = oldPdf.getParents().next();
  const had = {};
  receiptSources_(old).forEach(function (r) { had[r.fileData] = true; });
  receiptSources_(payload).forEach(function (r) { if (!had[r.fileData]) folder.createFile(dataUrlToBlob_(r.fileData, r.fileName)); });
  const split = packetReceipts_(payload);
  let pdfReceipts = split.attach;
  const formLabel = formLabel_(payload);
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const built = htmlToPdf_(payload, row.id, split.embed, formLabel + ' - ' + payload.last + ' - ' + dateStr + ' (resubmitted).pdf');
  if (!built.merged) { pdfReceipts = embeddedAsBlobs_(split.embed).concat(pdfReceipts); }
  const pdfFile = folder.createFile(built.pdf);
  try { oldPdf.setName(String(oldPdf.getName()).replace(/\.pdf$/i, '') + ' (sent back).pdf'); } catch (e) { /* keep going */ }
  savePayload_(row.jsonId, payload);

  // Claims row: new details, new PDF, new token; back to the VP Finance signature step.
  const t = payload.totals || {};
  const name = ((payload.first || '') + ' ' + (payload.last || '')).trim();
  const sh = getLogSheet_();
  sh.getRange(row.row, 4, 1, 6).setValues([[payload.payTo, name, payload.email, t.amt || 0, t.gst || 0, t.hst || 0]]);
  updateClaimPdf_(id, pdfFile.getId());
  const nt = rotateToken_(id);
  safeLedger_(function () { removeClaimLedgerRows_(id); }); // VP Finance writes them again when signing
  updateClaimStatus_(id, 'Pending VP Finance signature', payload.email);
  try { folder.moveTo(getRootFolder_()); } catch (e) { console.warn('Folder not moved back: ' + e); }
  sharePacketWith_(pdfFile.getId(), CONFIG.VP_FINANCE_EMAIL);

  MailApp.sendEmail({
    to: CONFIG.VP_FINANCE_EMAIL,
    subject: '[' + formLabel + '] ' + name + ' - $' + money_(t.amt) + ' - resubmitted after changes - needs your signature',
    htmlBody: '<p style="font-family:sans-serif"><b>' + escapeHtml_(name) + ' resubmitted this claim</b> after you sent it back'
      + (old.sentBackNote ? ' (your note: <i>' + escapeHtml_(old.sentBackNote) + '</i>)' : '') + '.</p>'
      + stageEmailHtml_('vp1', { id: row.id, token: nt, name: name, total: t.amt, folderUrl: row.folderUrl, form: payload.form }, pdfReceipts.length),
    attachments: emailAttachments_(built.pdf, pdfReceipts),
    name: CONFIG.ORG_NAME + ' Finance',
  });
  return { ok: true, id: String(row.id), resubmitted: true, reimbursement: money_(t.amt), folderName: payload.last + ', ' + dateStr, pdfReceiptsSeparate: pdfReceipts.length };
}

/* ---------- One-time fix: PDF receipts of claims filed before page pictures existed ----------
 * Those PDFs never showed inside the claim PDF (some claims were even sent back for a "missing" receipt).
 * Open the TEST deployment (Deploy > Test deployments; only the owner can open /dev) with
 *   ?action=pdfpages&ids=<id>,<id>  [&restore=<id>,<id>]  [&void=<old id>:<new id>]
 * The page draws each PDF receipt's pages (pdf.js, in the browser) and savePdfReceiptPages() stores them
 * on the claim and rebuilds its PDF (signatures kept). "restore" claims (sent back only because the receipt
 * wasn't showing) go back to "Pending VP Finance signature" and VP Finance gets the sign email again.
 * "void" marks a duplicate claim Void (replaced by the newer one) and moves its folder to the Drive trash. */
function pdfPagesPage_(p) {
  ownerOnly_();
  const list = function (s) { return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(String); };
  const cfg = { ids: list(p.ids), restore: list(p.restore), voids: list(p['void']) };
  const html = '<!doctype html><html><head><meta charset="utf-8"><title>PDF receipts fix</title></head>'
    + '<body style="font-family:sans-serif;padding:16px"><h3>Putting PDF receipts inside claim PDFs</h3><pre id="log" style="white-space:pre-wrap"></pre>'
    + '<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script><script>'
    + 'const CFG=' + JSON.stringify(cfg) + ';'
    + 'pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";'
    + 'const log=t=>{document.getElementById("log").textContent+=t+"\\n";};'
    + 'const call=(fn,...a)=>new Promise((res,rej)=>google.script.run.withSuccessHandler(res).withFailureHandler(rej)[fn](...a));'
    + 'function bytes(d){const t=String(d),b=atob(t.slice(t.indexOf(",")+1)),u=new Uint8Array(b.length);for(let i=0;i<b.length;i++)u[i]=b.charCodeAt(i);return u;}'
    + 'async function pages(d){const doc=await pdfjsLib.getDocument({data:bytes(d)}).promise;const out=[];try{for(let n=1;n<=doc.numPages;n++){'
    + 'const pg=await doc.getPage(n),v1=pg.getViewport({scale:1}),v=pg.getViewport({scale:Math.min(3,1400/Math.max(v1.width,v1.height))});'
    + 'const c=document.createElement("canvas");c.width=Math.ceil(v.width);c.height=Math.ceil(v.height);const x=c.getContext("2d");x.fillStyle="#fff";x.fillRect(0,0,c.width,c.height);'
    + 'await pg.render({canvasContext:x,viewport:v}).promise;out.push(c.toDataURL("image/jpeg",0.75));pg.cleanup();}}finally{doc.destroy();}return out;}'
    + '(async()=>{for(const id of CFG.ids){try{const c=await call("pdfReceiptsToFix",id);log(id+" "+c.name+" ("+c.status+"): "+c.items.length+" PDF receipt(s) to add");'
    + 'const out=[];for(const it of c.items){const p=await pages(it.fileData);log("  "+it.fileName+": "+p.length+" page(s)");out.push({kind:it.kind,idx:it.idx,pages:p});}'
    + 'log("  "+await call("savePdfReceiptPages",id,out,CFG.restore.indexOf(id)>=0));}catch(e){log("ERROR "+id+": "+(e&&e.message||e));}}'
    + 'for(const v of CFG.voids){try{const [o,n]=v.split(":");log(await call("voidDuplicateClaim",o,n));}catch(e){log("ERROR void "+v+": "+(e&&e.message||e));}}'
    + 'log("DONE");})();'
    + '</script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('PDF receipts fix');
}
// The claim's PDF receipts that have no page pictures yet.
function pdfReceiptsToFix(id) {
  ownerOnly_();
  const row = findClaim_(id);
  if (!row) throw new Error('Claim ' + id + ' not found.');
  const p = loadPayload_(row.jsonId);
  const items = [];
  const add = function (kind, list) {
    (list || []).forEach(function (x, i) {
      const pages = kind === 'line' ? x.filePages : x.pages;
      if (x.fileData && /^data:application\/pdf/i.test(x.fileData) && !(Array.isArray(pages) && pages.length))
        items.push({ kind: kind, idx: i, fileName: x.fileName || 'receipt.pdf', fileData: x.fileData });
    });
  };
  add('line', p.lines);
  add('att', p.attachments);
  return { id: String(row.id), name: row.name, status: row.status, items: items };
}
// Store the page pictures on the claim and rebuild its PDF; optionally put a sent-back claim back to VP Finance.
function savePdfReceiptPages(id, items, restore) {
  ownerOnly_();
  acquireLock_();
  const row = findClaim_(id);
  if (!row) throw new Error('Claim ' + id + ' not found.');
  const p = loadPayload_(row.jsonId);
  (items || []).forEach(function (it) {
    const list = it.kind === 'att' ? p.attachments : p.lines;
    const x = list && list[it.idx];
    if (!x || !Array.isArray(it.pages) || !it.pages.length || !it.pages.every(pdfEmbeddable_)) return;
    if (it.kind === 'att') x.pages = it.pages; else x.filePages = it.pages;
  });
  const st = stageOf_(row.status);
  if (restore) {
    if (!isSentBack_(row.status)) throw new Error('Claim ' + id + ' is not sent back (status: ' + row.status + ').');
    p.lastSentBackNote = p.sentBackNote || '';
    delete p.sentBackNote;
    savePayload_(row.jsonId, p);
    rebuildAndSavePdf_(p, row);
    const nt = rotateToken_(id);
    safeLedger_(function () { removeClaimLedgerRows_(id); }); // VP Finance writes them again when signing
    updateClaimStatus_(id, 'Pending VP Finance signature', CONFIG.VP_FINANCE_EMAIL);
    try { DriveApp.getFileById(row.pdfId).getParents().next().moveTo(getRootFolder_()); } catch (e) { console.warn('Folder not moved back: ' + e); }
    sharePacketWith_(row.pdfId, CONFIG.VP_FINANCE_EMAIL);
    const extras = nonImageReceiptBlobs_(p);
    MailApp.sendEmail({
      to: CONFIG.VP_FINANCE_EMAIL,
      subject: '[' + formLabel_(p) + '] ' + row.name + ' - $' + money_(row.total) + ' - receipt now inside the PDF - needs your signature',
      htmlBody: '<p style="font-family:sans-serif"><b>' + escapeHtml_(row.name) + '</b> had uploaded a PDF receipt, but it was not showing inside the claim PDF, so the claim was sent back'
        + (p.lastSentBackNote ? ' (your note: <i>' + escapeHtml_(p.lastSentBackNote) + '</i>)' : '') + '. The receipt is now inside the attached PDF and the claim is back for your signature.</p>'
        + stageEmailHtml_('vp1', { id: row.id, token: nt, name: row.name, total: row.total, folderUrl: row.folderUrl, form: p.form }, extras.length),
      attachments: emailAttachments_(DriveApp.getFileById(row.pdfId).getBlob(), extras),
      name: CONFIG.ORG_NAME + ' Finance',
    });
    return 'Back to "Pending VP Finance signature"; receipt inside the PDF; sign email sent to VP Finance.';
  }
  savePayload_(row.jsonId, p);
  if (isSentBack_(row.status) || st === 'done') return 'Pages saved (the PDF is rebuilt when the claim is resubmitted).';
  rebuildAndSavePdf_(p, row);
  if (st === 'president' || st === 'vpreturned') sharePacketWithPresident_(row.pdfId);
  return 'Pages saved and claim PDF rebuilt with the receipt inside (status unchanged: ' + row.status + ').';
}
// Mark a duplicate claim Void (replaced by a newer claim) and move its Drive folder to the trash.
function voidDuplicateClaim(oldId, newId) {
  ownerOnly_();
  acquireLock_();
  const row = findClaim_(oldId), keep = findClaim_(newId);
  if (!row || !keep) throw new Error('Claim ' + (row ? newId : oldId) + ' not found.');
  if (!isSentBack_(row.status)) throw new Error('Claim ' + oldId + ' is not sent back (status: ' + row.status + ').');
  updateClaimStatus_(oldId, 'Void - replaced by claim ' + newId, CONFIG.VP_FINANCE_EMAIL);
  safeLedger_(function () { removeClaimLedgerRows_(oldId); });
  let trashed = false;
  try {
    const folder = DriveApp.getFileById(row.pdfId).getParents().next();
    if (folder.getId() !== DriveApp.getFileById(keep.pdfId).getParents().next().getId()) { folder.setTrashed(true); trashed = true; }
  } catch (e) { console.warn('Folder not trashed: ' + e); }
  return oldId + ': Void - replaced by claim ' + newId + (trashed ? '; its Drive folder is in the trash (restorable for 30 days).' : '.');
}

// Run from the editor: change GL codes that VP Finance picked wrong when signing (e.g. a Yearbook expense
// signed with the Courier code of the same name). Only before the President signs. In GL_FIXES put
// claim ID -> { line number (1 = first expense line): new GL code }; lines left out keep their code.
// For each claim: new codes saved on the claim, PDF rebuilt (signatures kept), Expense Ledger rows
// rewritten (Locked rows are not touched) and, if it is waiting for the President, she gets the sign
// email again (her old link stops working). A "Returned by President" claim stays with you to re-send.
const GL_FIXES = {
  // 'abc12345': { 1: '410-72108-56025100003906310000', 3: '410-72108-56025100003906310000' },
};
function fixClaimGlCodes() {
  ownerOnly_();
  acquireLock_();
  // Check every claim first so a mistake in the list changes nothing.
  const jobs = Object.keys(GL_FIXES).map(function (id) {
    const row = findClaim_(id);
    if (!row) throw new Error('Claim ' + id + ' not found.');
    const st = stageOf_(row.status);
    if (st !== 'president' && st !== 'vpreturned') throw new Error('Claim ' + id + ' is "' + row.status + '": GL codes can only be fixed before the President signs.');
    const p = loadPayload_(row.jsonId);
    if (isMileage_(p)) throw new Error('Claim ' + id + ' is a kilometer claim: its GL code is fixed.');
    const fix = GL_FIXES[id];
    Object.keys(fix).forEach(function (n) {
      if (!(p.lines || [])[n - 1]) throw new Error('Claim ' + id + ' has no line ' + n + '.');
      if (!glLabel_(fix[n])) throw new Error('Claim ' + id + ' line ' + n + ': ' + fix[n] + ' is not on the GL list.');
    });
    return { id: id, row: row, st: st, p: p, fix: fix };
  });
  return jobs.map(function (j) {
    const row = j.row, p = j.p;
    Object.keys(j.fix).forEach(function (n) { p.lines[n - 1].gl = j.fix[n]; p.lines[n - 1].glName = glLabel_(j.fix[n]); });
    savePayload_(row.jsonId, p);
    safeLedger_(function () { writeLedgerRows_(row, p, row.status); });
    rebuildAndSavePdf_(p, row);
    sharePacketWithPresident_(row.pdfId);
    let msg = j.id + ' (' + row.name + '): ' + Object.keys(j.fix).length + ' GL code(s) changed, PDF rebuilt, ledger updated';
    if (j.st === 'president') {
      const nt = rotateToken_(j.id);
      updateClaimStatus_(j.id, row.status, CONFIG.VP_FINANCE_EMAIL);
      MailApp.sendEmail({
        to: CONFIG.PRESIDENT_EMAIL,
        subject: '[' + formLabel_(row) + '] ' + row.name + ' - $' + money_(row.total) + ' - GL codes corrected - needs your signature',
        htmlBody: '<p style="font-family:sans-serif">VP Finance corrected the GL codes on this claim, so its PDF was replaced. '
          + 'Please sign it with the button below (the link in the earlier email no longer works).</p>'
          + stageEmailHtml_('president', { id: j.id, token: nt, name: row.name, total: row.total, folderUrl: row.folderUrl, form: row.form }, 0),
        attachments: capBlobs_(nonImageReceiptBlobs_(p)),
        name: CONFIG.ORG_NAME + ' Finance',
      });
      msg += ', sent to the President to sign again.';
    } else {
      msg += '; still "' + row.status + '": open it from the President\'s return email to re-send it.';
    }
    console.log(msg);
    return msg;
  });
}

// The claim's (unlocked) Expense Ledger rows, removed when it is resubmitted: lines may have changed.
function removeClaimLedgerRows_(id) {
  const sh = getLedgerSheet_();
  const data = sh.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    if (sameId_(data[r][LC['Claim ID']], id) && String(data[r][LC['Locked']]).toLowerCase() !== 'yes') sh.deleteRow(r + 1);
  }
}

// ONE-TIME (run from the editor), for claims sent back BEFORE edit links existed (their email linked to
// the empty form). listSentBackClaims logs every claim at "Sent back for changes". Put the IDs you want
// to give an edit link in RESEND_EDIT_LINK_IDS below, then run resendEditLinks: each requester gets a new
// email with an "Edit your claim" button (VP Finance in copy). Leave a claim out if it was already refiled.
const RESEND_EDIT_LINK_IDS = [];
function listSentBackClaims() {
  ownerOnly_();
  const data = getLogSheet_().getDataRange().getValues();
  const out = [];
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] && isSentBack_(data[r][2])) out.push(String(data[r][0]) + '  ' + data[r][4] + '  $' + money_(data[r][6]) + '  (sent back ' + data[r][13] + ')');
  }
  console.log(out.length ? 'Claims sent back for changes:\n' + out.join('\n') : 'No claims are at "Sent back for changes".');
  return out.length;
}
function resendEditLinks() {
  ownerOnly_();
  if (!RESEND_EDIT_LINK_IDS.length) { console.log('Add claim IDs to RESEND_EDIT_LINK_IDS first (run listSentBackClaims to see them).'); return 0; }
  let sent = 0;
  RESEND_EDIT_LINK_IDS.forEach(function (id) {
    const row = findClaim_(id);
    if (!row || !isSentBack_(row.status)) { console.log('Skipped ' + id + ': ' + (row ? 'status is ' + row.status : 'not found')); return; }
    let note = '';
    try { note = loadPayload_(row.jsonId).sentBackNote || ''; } catch (e) { /* older claim */ }
    const nt = rotateToken_(row.id);
    MailApp.sendEmail({
      to: row.email, cc: CONFIG.VP_FINANCE_EMAIL,
      subject: 'Your NUSA ' + formLabelLc_(row) + ' needs a change - edit link',
      htmlBody: (note ? '' : '<p style="font-family:sans-serif">What needs fixing is in our earlier email. You can now edit your claim directly instead of filing a new one.</p>')
        + sendBackEmailHtml_(row, note, editUrl_(row.id, nt)),
      name: CONFIG.ORG_NAME + ' Finance',
    });
    console.log('Edit link sent: ' + row.id + ' (' + row.name + ')');
    sent++;
  });
  return sent;
}

/* ============================================================
 *  3. DAILY REMINDERS
 *     Once a day, email whoever's signature/approval a claim is
 *     waiting on (VP Finance, President or NU Staff) the working link.
 *
 *  SETUP: run installDailyReminderTrigger() ONCE from the editor
 *  (Run > installDailyReminderTrigger). To change the send time,
 *  edit CONFIG.REMINDER_HOUR and run it again. To turn reminders
 *  off, run removeReminderTriggers().
 * ============================================================ */

// Who to remind for each pending stage (null = no reminder for that status).
function reminderRecipient_(stage) {
  if (stage === 'president') return CONFIG.PRESIDENT_EMAIL;
  if (stage === 'staff') return CONFIG.STAFF_EMAIL; // claims routed by link before V5
  if (stage === 'staffpdf') return CONFIG.VP_FINANCE_EMAIL; // V5: VP Finance relays the PDF
  if (stage === 'vp1' || stage === 'vpfinal' || stage === 'vpreturned' || stage === 'staffreturned') return CONFIG.VP_FINANCE_EMAIL;
  return null; // 'done' / approved / sent back — nothing to chase
}

// What the recipient still needs to do, per stage.
function reminderAction_(stage) {
  if (stage === 'vp1') return 'your signature and approval';
  if (stage === 'president' || stage === 'staff' || stage === 'staffpdf') return 'your signature';
  if (stage === 'vpfinal') return 'your final approval';
  if (stage === 'vpreturned') return 'your attention (the President returned it with a note)';
  if (stage === 'staffreturned') return 'your attention (NU Staff returned it with a note)';
  return '';
}

// Whole days since a date (min 0).
function daysSince_(d) {
  if (!d) return 0;
  const ms = new Date().getTime() - new Date(d).getTime();
  return Math.max(0, Math.floor(ms / 86400000));
}

// Automatic reminders: one email when a claim has been waiting 5 days at the same step, another at
// 7 days and a last one at 14 days ("waiting since" = the last action on it, else the submission time). Each is
// sent once (remembered in Script Properties as REM_<claim id>); when the claim moves to the next step the
// count starts again for whoever it is waiting on. A claim found already past a mark (e.g. 9 days on the first
// run) gets just the latest reminder it has reached (the 7-day one).
const REMINDER_DAYS = [5, 7, 14];
function reminderDue_(props, id, since, days) {
  const reached = REMINDER_DAYS.filter(function (d) { return days >= d; });
  if (!reached.length) return 0;
  const top = reached[reached.length - 1];
  const saved = String(props.getProperty('REM_' + id) || '').split('|');
  const stamp = String(new Date(since).getTime());
  return (saved[0] === stamp && Number(saved[1]) >= top) ? 0 : top;
}
function markReminderSent_(props, id, since, top) {
  props.setProperty('REM_' + id, String(new Date(since).getTime()) + '|' + top);
}
// Forget the marks of claims that aren't waiting on anyone any more.
function cleanReminderMarks_(props, waiting) {
  try {
    props.getKeys().forEach(function (k) { if (k.indexOf('REM_') === 0 && !waiting[k.slice(4)]) props.deleteProperty(k); });
  } catch (e) { /* cosmetic */ }
}
// "[Reminder] The expense claim from Ana Test ($100.00) has not been signed for the past 5 days"
function reminderSubject_(claim, stage, days) {
  const what = 'The ' + formLabelLc_(claim) + ' from ' + claim.name + ' ($' + money_(claim.total) + ')';
  const span = 'for the past ' + days + (days === 1 ? ' day' : ' days');
  if (stage === 'vpreturned' || stage === 'staffreturned') return '[Reminder] ' + what + (days > 0 ? ' has been waiting on you ' + span : ' is waiting on you');
  return '[Reminder] ' + what + (days > 0 ? ' has not been signed ' + span : ' has not been signed yet');
}

// Scan the ledger and send the 5 / 7 / 14-day reminders that are due.
// Safe to run repeatedly; each reminder goes out once.
function sendDailyReminders() {
  // The daily trigger must be able to call this, so it can't be owner-only; instead it runs at
  // most once an hour, so nobody can use it to flood inboxes or use up the daily email quota.
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('LAST_REMINDER_RUN') || 0);
  if (Date.now() - last < 3600000) { console.log('Reminders already sent in the last hour; skipped.'); return 0; }
  props.setProperty('LAST_REMINDER_RUN', String(Date.now()));
  const sh = getLogSheet_();
  const data = sh.getDataRange().getValues();
  const waiting = {};
  let sent = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const id = row[0], status = row[2];
    if (!id) continue;
    const stage = stageOf_(status);
    const to = reminderRecipient_(stage);
    if (!to) continue; // not a pending stage

    const claim = {
      id: id, name: row[4], email: row[5], total: row[6],
      folderUrl: row[9], token: row[11], form: row[15], status: status,
    };
    // "Waiting since" = the last action (Decided at), else the submission time.
    const since = row[13] || row[1];
    const days = daysSince_(since);
    waiting[String(id)] = true;
    const due = reminderDue_(props, id, since, days);
    if (!due) continue; // not 5, 7 or 14 days yet, or that reminder was already sent
    try {
      if (stage === 'staffpdf') { // V5: VP Finance gets the hand-off email again (PDF + message for NU Staff)
        claim.payTo = row[3]; claim.pdfId = row[10];
        let payload = {};
        try { payload = loadPayload_(row[14]) || {}; } catch (e) { payload = {}; }
        sendStaffPdfEmail_(claim, payload, true, undefined, days);
        markReminderSent_(props, id, since, due);
        sent++;
        continue;
      }
      MailApp.sendEmail({
        to: to,
        subject: reminderSubject_(claim, stage, days),
        htmlBody: reminderEmailHtml_(claim, stage, days),
        name: CONFIG.ORG_NAME + ' Finance',
      });
      markReminderSent_(props, id, since, due);
      sent++;
    } catch (err) {
      console.warn('Reminder failed for claim ' + id + ': ' + err);
    }
  }
  cleanReminderMarks_(props, waiting);
  console.log('Reminders sent (5 / 7 / 14 days): ' + sent);
  return sent;
}

function reminderEmailHtml_(claim, stage, days, manual) {
  const url = decisionUrl_(claim.id, claim.token); // current, still-valid link for this stage
  const lbl = formLabelLc_(claim);
  const waited = days <= 0 ? 'since earlier today'
    : (days === 1 ? 'for 1 day' : 'for ' + days + ' days');
  return '<div style="font-family:sans-serif;max-width:520px">'
    + '<h2 style="color:' + CONFIG.BRAND_GREEN + '">Reminder: ' + (days > 0 ? 'this ' + lbl + ' has been waiting ' + waited : 'a ' + lbl + ' is waiting on you') + '</h2>'
    + '<p>This ' + lbl + ' has been waiting ' + waited + ' and still needs <b>' + reminderAction_(stage) + '</b>.</p>'
    + '<p><b>' + escapeHtml_(claim.name) + '</b> &mdash; $' + money_(claim.total) + '</p>'
    + '<p style="margin:22px 0"><a href="' + url + '" style="background:' + CONFIG.BRAND_GREEN + ';color:#fff;text-decoration:none;padding:12px 22px;border-radius:8px;font-weight:bold">Open the claim &rarr;</a></p>'
    + '<p style="font-size:12px;color:#666">' + (manual ? 'Sent by NUSA VP Finance.' : 'This is an automatic reminder, sent when a claim has waited ' + REMINDER_DAYS.join(', ').replace(/, (\d+)$/, ' and $1') + ' days. It stops as soon as the step is done.')
    + (claim.folderUrl ? ' &middot; <a href="' + claim.folderUrl + '">Drive folder</a>' : '') + '</p></div>';
}

// Create (or refresh) the once-a-day reminder trigger at CONFIG.REMINDER_HOUR.
function installDailyReminderTrigger() {
  ownerOnly_();
  removeReminderTriggers();
  ScriptApp.newTrigger('sendDailyReminders')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.REMINDER_HOUR)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  console.log('Reminder check scheduled daily at ~' + CONFIG.REMINDER_HOUR + ':00 (' + Session.getScriptTimeZone() + '); emails go out after ' + REMINDER_DAYS.join(', ') + ' days.');
}

// Remove any existing reminder triggers (so re-installing doesn't stack them up).
function removeReminderTriggers() {
  ownerOnly_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyReminders') ScriptApp.deleteTrigger(t);
  });
}

// Manual reminders, run from the editor (pick the function in the list, then Run):
//   remindPresident -> emails the President every claim waiting on her signature.
//   remindVpFinance -> emails VP Finance (you) every claim waiting on your signature, final approval,
//                      or your attention after the President / NU Staff returned it.
//   remindStaff     -> emails NU Staff every claim waiting on their signature (link claims);
//                      for claims waiting on the signed PDF (V5), re-sends VP Finance the
//                      PDF + message to forward to NU Staff.
// Only the person named is emailed. No hourly limit, so you can re-run them any time.
function remindPresident() {
  ownerOnly_();
  return remindStage_(['president'], 'the President');
}

function remindStaff() {
  ownerOnly_();
  return remindStage_(['staff', 'staffpdf'], 'NU Staff');
}

function remindVpFinance() {
  ownerOnly_();
  return remindStage_(['vp1', 'vpfinal', 'vpreturned', 'staffreturned'], 'VP Finance');
}

function remindStage_(stages, who) {
  const data = getLogSheet_().getDataRange().getValues();
  let sent = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const id = row[0], status = row[2];
    const stage = stageOf_(status);
    if (!id || stages.indexOf(stage) < 0) continue;
    const claim = {
      id: id, name: row[4], email: row[5], total: row[6],
      folderUrl: row[9], token: row[11], form: row[15], status: status,
    };
    if (stage === 'staffpdf') { // V5: NU Staff signs by email, so VP Finance gets the hand-off email again
      claim.payTo = row[3]; claim.pdfId = row[10];
      let payload = {};
      try { payload = loadPayload_(row[14]) || {}; } catch (e) { payload = {}; }
      sendStaffPdfEmail_(claim, payload, true, undefined, daysSince_(row[13] || row[1]));
      console.log('Sent to VP Finance (to forward to ' + CONFIG.STAFF_NAME + '): claim ' + id + ' (' + claim.name + ', $' + money_(claim.total) + ')');
      sent++;
      continue;
    }
    const to = reminderRecipient_(stage);
    MailApp.sendEmail({
      to: to,
      subject: reminderSubject_(claim, stage, daysSince_(row[13] || row[1])),
      htmlBody: reminderEmailHtml_(claim, stage, daysSince_(row[13] || row[1]), true),
      name: CONFIG.ORG_NAME + ' Finance',
    });
    console.log('Sent to ' + to + ': claim ' + id + ' (' + claim.name + ', $' + money_(claim.total) + ')');
    sent++;
  }
  console.log(sent ? 'Done: ' + sent + ' reminder(s) sent.' : 'No claims are waiting on ' + who + '.');
  return sent;
}

/* ---------- owner-only guard ----------
 * Every function without a trailing "_" can be called from a browser by anyone who opens the
 * public claim form (google.script.run), and it runs as the owner. The maintenance functions
 * below (setup, migration, triggers, feed key...) are meant to be run only from the Apps Script
 * editor, so they call ownerOnly_() first: in the editor the person running it is the owner;
 * from the web app the visitor is anonymous, so the call is refused.
 * Needs the userinfo.email scope in appsscript.json. */
function ownerOnly_() {
  let active = '', owner = '';
  try {
    active = Session.getActiveUser().getEmail();
    owner = Session.getEffectiveUser().getEmail();
  } catch (e) {
    throw new Error('Add "https://www.googleapis.com/auth/userinfo.email" to oauthScopes in appsscript.json, save, and run this again.');
  }
  if (!active || active !== owner) throw new Error('This can only be run from the Apps Script editor by the script owner.');
}

/* ---------- approval-chain helpers ---------- */
// Serialize claim actions so two fast clicks can't both slip through. The lock is
// released automatically when this execution ends (after the status has been written),
// so a second, queued click re-reads the advanced status below and is rejected.
function acquireLock_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('This claim is already being processed. Please wait a moment, then reopen the link.');
  }
  return lock;
}
function guard_(id, token, expectStage) {
  acquireLock_();
  const row = findClaim_(id);
  if (!row || row.token !== token) throw new Error('Invalid link.');
  if (stageOf_(row.status) !== expectStage) throw new Error('This step has already been completed (' + row.status + ').');
  return row;
}
// The base web-app URL. Prefer the stable /exec URL pinned in CONFIG; only fall
// back to getUrl() (which can be the temporary /dev URL) if that is left blank.
function webAppUrl_() {
  const pinned = CONFIG.WEB_APP_URL && String(CONFIG.WEB_APP_URL).trim();
  return pinned || ScriptApp.getService().getUrl();
}
function decisionUrl_(id, token) {
  return webAppUrl_() + '?action=decision&id=' + encodeURIComponent(id) + '&token=' + encodeURIComponent(token);
}
function loadPayload_(jsonId) {
  return JSON.parse(DriveApp.getFileById(jsonId).getBlob().getDataAsString());
}
function savePayload_(jsonId, payload) {
  DriveApp.getFileById(jsonId).setContent(JSON.stringify(payload));
}
// Every uploaded file, from per-line receipts (expense) and the shared attachments list (mileage).
// pages = the form's JPEG picture of each page of a PDF receipt (so the PDF can go inside the packet).
function receiptSources_(payload) {
  const out = [];
  (payload.lines || []).forEach(function (l, i) {
    if (l.fileData) out.push({ fileName: l.fileName || ('receipt-' + (i + 1)), fileType: l.fileType, fileData: l.fileData, pages: l.filePages });
  });
  (payload.attachments || []).forEach(function (a, i) {
    if (a.fileData) out.push({ fileName: a.fileName || ('attachment-' + (i + 1)), fileType: a.fileType, fileData: a.fileData, pages: a.pages });
  });
  return out;
}
// The merged packet PDF has a size ceiling, so only embed image receipts up to a budget;
// anything over rides along as an email attachment (and every file is saved to Drive regardless).
const PACKET_EMBED_BUDGET = 3500000; // ~3.5 MB of image data embedded into the PDF
const EMAIL_ATTACH_CAP = 20000000;   // ~20 MB total per email (Gmail cap is 25 MB); overflow stays in the Drive folder
function b64Bytes_(dataUrl){ var t=String(dataUrl||''); var i=t.indexOf(','); return i<0?0:Math.floor((t.length-i-1)*0.75); }
// Only JPEG / PNG / GIF can be drawn inside the PDF. Other images (e.g. iPhone HEIC photos, which
// the form now converts to JPEG, but older claims may still hold) go out as separate attachments
// instead of printing as a blank box. Checked on the file data itself, not the browser's label.
function pdfEmbeddable_(dataUrl){ return /^data:image\/(jpeg|jpg|png|gif);base64,/i.test(String(dataUrl||'')); }
// A PDF receipt goes inside the packet as its page pictures (all pages or none); the original PDF
// is still saved in the claim folder. Without usable pages it rides as an attachment, as before.
function pdfPages_(r){
  const pages=r.pages;
  if(!Array.isArray(pages)||!pages.length||pdfEmbeddable_(r.fileData)) return null;
  for(var i=0;i<pages.length;i++){ if(!pdfEmbeddable_(pages[i])) return null; }
  return pages;
}
function packetReceipts_(payload){
  const embed=[], attach=[]; let used=0;
  receiptSources_(payload).forEach(function(r, k){
    const pages=pdfPages_(r);
    if(pages){
      const sz=pages.reduce(function(s,p){ return s+b64Bytes_(p); },0);
      if(used+sz<=PACKET_EMBED_BUDGET){
        pages.forEach(function(p,i){ embed.push({name:r.fileName+(pages.length>1?' (page '+(i+1)+' of '+pages.length+')':''), dataUrl:p, no:k+1, source:r}); });
        used+=sz;
      } else { attach.push(dataUrlToBlob_(r.fileData, r.fileName)); }
    } else if(pdfEmbeddable_(r.fileData)){
      const sz=b64Bytes_(r.fileData);
      if(used+sz<=PACKET_EMBED_BUDGET){ embed.push({name:r.fileName, dataUrl:r.fileData, no:k+1}); used+=sz; }
      else { attach.push(dataUrlToBlob_(r.fileData, r.fileName)); }
    } else { attach.push(dataUrlToBlob_(r.fileData, r.fileName)); }
  });
  return {embed:embed, attach:attach};
}
// When the packet had to be made without receipts: the embedded files as attachments
// (a PDF receipt once, as the original PDF, not one picture per page).
function embeddedAsBlobs_(embed){
  const out=[], seen=[];
  (embed||[]).forEach(function(r){
    if(r.source){ if(seen.indexOf(r.source)<0){ seen.push(r.source); out.push(dataUrlToBlob_(r.source.fileData, r.source.fileName)); } }
    else out.push(dataUrlToBlob_(r.dataUrl, r.name));
  });
  return out;
}
// Claim IDs are 8 hex characters. Google Sheets turns some of them into numbers when they're logged
// (e.g. "01234567" -> 1234567, "12345e12" -> 1.2345E+16), and then the email link no longer matched
// its row ("Claim not found"). New IDs are stored as text (logClaim_); this also matches the old ones.
function sameId_(cell, id) {
  if (String(cell) === String(id)) return true;
  return typeof cell === 'number' && String(id).trim() !== '' && Number(id) === cell;
}
function capBlobs_(blobs){ const out=[]; let used=0; (blobs||[]).forEach(function(b){ var nn=0; try{ nn=b.getBytes().length; }catch(e){} if(used+nn<=EMAIL_ATTACH_CAP){ out.push(b); used+=nn; } }); return out; }
function emailAttachments_(first, blobs){ const out=[first]; let used=0; try{ used=first.getBytes().length; }catch(e){} (blobs||[]).forEach(function(b){ var nn=0; try{ nn=b.getBytes().length; }catch(e){} if(used+nn<=EMAIL_ATTACH_CAP){ out.push(b); used+=nn; } }); return out; }
// HTML -> PDF, with a fallback to a form-only PDF if the merged packet is too large for the converter.
function htmlToPdf_(payload, id, embed, name){
  const build = function(imgs){ return isMileage_(payload) ? buildMileageHtml_(payload, id, imgs) : buildPacketHtml_(payload, id, imgs); };
  try { return { pdf: Utilities.newBlob(build(embed),'text/html','claim.html').getAs('application/pdf').setName(name), merged:true }; }
  catch (e) { return { pdf: Utilities.newBlob(build([]),'text/html','claim.html').getAs('application/pdf').setName(name), merged:false }; }
}
function imageReceiptsFrom_(payload) { return packetReceipts_(payload).embed; }
function nonImageReceiptBlobs_(payload) { return packetReceipts_(payload).attach; }
function formLabelLc_(x) { return (x && x.form === 'mileage') ? 'kilometer reimbursement' : 'expense claim'; }
// Re-generate the PDF with whatever signatures are now on the payload, replace the old one.
function rebuildAndSavePdf_(payload, row) {
  if (!payload.submittedAt && row.timestamp instanceof Date) {
    payload.submittedAt = row.timestamp.toISOString(); // pre-V4 claim: keep its real submission date
    savePayload_(row.jsonId, payload);
  }
  const oldFile = DriveApp.getFileById(row.pdfId);
  const folder = oldFile.getParents().next();
  const name = oldFile.getName();
  const pdf = htmlToPdf_(payload, row.id, imageReceiptsFrom_(payload), name).pdf;
  const newFile = folder.createFile(pdf);
  try { oldFile.setTrashed(true); } catch (e) { /* keep going */ }
  updateClaimPdf_(row.id, newFile.getId());
  row.pdfId = newFile.getId();
  return newFile.getId();
}


/* ============================================================
 *  Validation
 * ============================================================ */
function validateMileage_(p) {
  const miss = [];
  if (!p.first || !p.last) miss.push('name');
  if (!p.email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(p.email)) miss.push('a valid email');
  if (p.payTo === 'Student' && !p.studentId) miss.push('student ID');
  if (p.payTo === 'Other' && (!p.company || !p.contact)) miss.push('company details');
  if (!p.address || !p.city) miss.push('address');
  if (!p.signature) miss.push('a signature');
  else if (!isImageDataUrl_(p.signature)) miss.push('a signature (please clear it and draw it again)');
  if (!p.printName) miss.push('a printed name');
  if (!p.lines || !p.lines.length) miss.push('at least one trip');
  (p.lines || []).forEach(function (l, i) {
    const n = ' (trip ' + (i + 1) + ')';
    if (!(parseFloat(l.km) > 0)) miss.push('kilometres' + n);
    if (!l.destination || l.destination.trim().length < 3) miss.push('a destination' + n);
  });
  if (!p.attachments || !p.attachments.length) miss.push('at least one trip screenshot or receipt');
  return miss;
}

function validatePayload(p) {
  const miss = [];
  if (!p) return ['everything'];
  if (isMileage_(p)) return validateMileage_(p);
  if (!p.first || !p.last) miss.push('name');
  if (!p.email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(p.email)) miss.push('a valid email');
  if (p.payTo === 'Student' && !p.studentId) miss.push('student ID');
  if (p.payTo === 'Other' && (!p.company || !p.contact)) miss.push('company details');
  if (!p.address || !p.city) miss.push('address');
  if (!p.signature) miss.push('a signature');
  else if (!isImageDataUrl_(p.signature)) miss.push('a signature (please clear it and draw it again)');
  if (!p.printName) miss.push('a printed name');
  if (!p.lines || !p.lines.length) miss.push('at least one expense');
  (p.lines || []).forEach(function (l, i) {
    const n = ' (receipt ' + (i + 1) + ')';
    if (!(parseFloat(l.amount) > 0)) miss.push('an amount' + n);
    if (!l.fileData) miss.push('a receipt' + n);
    if (!l.desc || l.desc.trim().split(/\s+/).length < 3) miss.push('a specific explanation' + n);
  });
  return miss;
}


/* ============================================================
 *  Drive helpers
 * ============================================================ */
function getRootFolder_() {
  if (CONFIG.DRIVE_ROOT_FOLDER_ID) return DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
  const it = DriveApp.getFoldersByName(CONFIG.DRIVE_ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.DRIVE_ROOT_FOLDER_NAME);
}

function subFolder_(name) {
  const root = getRootFolder_();
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function moveClaimFolder_(pdfId, statusFolderName) {
  try {
    const parents = DriveApp.getFileById(pdfId).getParents();
    if (!parents.hasNext()) return;
    const claimFolder = parents.next();
    claimFolder.moveTo(subFolder_(statusFolderName));
  } catch (err) {
    // Folder move is a nicety; never fail the approval because of it.
    console.warn('moveClaimFolder_ skipped: ' + err);
  }
}

function dataUrlToBlob_(dataUrl, name) {
  const m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/);
  if (!m) throw new Error('Bad file data for ' + name);
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
}


/* ============================================================
 *  Ledger (Google Sheet) helpers
 * ============================================================ */
const LOG_HEADERS = ['ID', 'Timestamp', 'Status', 'Payable to', 'Name', 'Email',
  'Total', 'GST', 'HST', 'Folder', 'PDF file id', 'Token', 'Decided by', 'Decided at', 'JSON file id', 'Form'];

function getLogSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('LOG_SS_ID');
  let ss;
  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create('NUSA Expense Claims — Log');
    props.setProperty('LOG_SS_ID', ss.getId());
    getRootFolder_().addFile(DriveApp.getFileById(ss.getId())); // keep it with the claims
  }
  let sh = ss.getSheetByName('Claims');
  if (!sh) {
    sh = ss.getSheets()[0].setName('Claims');
    sh.appendRow(LOG_HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function logClaim_(c) {
  const sh = getLogSheet_();
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); // keep claim IDs as text (see sameId_)
  sh.appendRow([
    c.id, new Date(), c.status, c.payTo, c.name, c.email,
    c.total, c.gst, c.hst, c.folderUrl, c.pdfId, c.token, '', '', c.jsonId, c.form,
  ]);
  // A row appended past the formatted range may not inherit the text format: re-write the ID as text
  // (only on this claim's own row, found by its unique token, in case two claims arrive at once).
  const last = sh.getLastRow();
  const tail = sh.getRange(Math.max(2, last - 4), 1, Math.min(5, last - 1), 12).getValues();
  for (let i = tail.length - 1; i >= 0; i--) {
    if (tail[i][11] === c.token) { sh.getRange(Math.max(2, last - 4) + i, 1).setNumberFormat('@').setValue(String(c.id)); break; }
  }
}

function findClaim_(id) {
  const sh = getLogSheet_();
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (sameId_(data[r][0], id)) {
      return {
        row: r + 1, id: data[r][0], timestamp: data[r][1], status: data[r][2], payTo: data[r][3],
        name: data[r][4], email: data[r][5], total: data[r][6], gst: data[r][7],
        hst: data[r][8], folderUrl: data[r][9], pdfId: data[r][10], token: data[r][11],
        jsonId: data[r][14], form: data[r][15],
      };
    }
  }
  return null;
}

function updateClaimPdf_(id, pdfId) {
  const sh = getLogSheet_();
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (sameId_(data[r][0], id)) { sh.getRange(r + 1, 11).setValue(pdfId); return; }
  }
}

// Rotate the claim's token so old email links can't act on later stages.
function rotateToken_(id) {
  const nt = Utilities.getUuid();
  const sh = getLogSheet_();
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (sameId_(data[r][0], id)) { sh.getRange(r + 1, 12).setValue(nt); return nt; }
  }
  return nt;
}

function updateClaimStatus_(id, status, who) {
  const sh = getLogSheet_();
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (sameId_(data[r][0], id)) {
      sh.getRange(r + 1, 3).setValue(status);
      sh.getRange(r + 1, 13).setValue(who);
      sh.getRange(r + 1, 14).setValue(new Date());
      safeLedger_(function () { syncLedgerStatus_(id, status); }); // Expense Ledger (V3)
      return;
    }
  }
}

/* ============================================================
 *  4. EXPENSE LEDGER  (feeds the Excel budget tracker)
 *     One row per expense line, in the "Expense Ledger" tab of the
 *     claim log spreadsheet. The Excel tracker reads this tab.
 *
 *  - Rows are written when VP Finance assigns GL codes and signs
 *    (vpSignApprove). Mileage rows use the fixed mileage GL.
 *  - Every status change re-syncs the claim's rows:
 *      Pending President / NU Staff / final / returned -> Pending
 *      Approved (final PDF sent to VP Finance)        -> Approved
 *      Sent back for changes                          -> Void
 *  - Your own purchases: add a row by hand (Source = Manual).
 *  - Your edits: set Locked = Yes on a row and the script will never
 *    overwrite it again (it only refreshes Claim Status / Last Updated).
 *    "Charge to Department" and "Notes" are never overwritten.
 *  - Columns after Notes (Submitted / Cheque Received from AP / Cheque Picked Up,
 *    added by the sheet's own "Claim Tracking" script) are never touched.
 *
 *  SETUP: run setupExpenseLedger() ONCE from the editor. It creates
 *  the tab, fills it with every claim already past the VP signature,
 *  and logs the private feed link to paste into the Excel tracker
 *  (Settings sheet, B3). The claim log can then be set to Restricted.
 * ============================================================ */
const LEDGER_SHEET_NAME = 'Expense Ledger';
const LEDGER_HEADERS = ['Entry ID', 'Source', 'Date', 'GL Code', 'GL Name', 'GL Group',
  'Charge to Department (optional)', 'Description', 'Payee', 'Amount (incl. tax)', 'GST', 'HST',
  'Net Expense', 'Status', 'Locked', 'Claim ID', 'Line #', 'Claim Status', 'Last Updated', 'Notes'];
const LC = {}; LEDGER_HEADERS.forEach(function (h, i) { LC[h] = i; }); // column index by header
// Department names exactly as they appear in the Excel tracker (for "Charge to Department").
const TRACKER_DEPTS = ['Administration', 'Faculty', 'Communications', 'Events', "The Courier", 'Yearbook',
  'Student Lounge (The Commons)', 'Special Initiatives', 'Proposal Committee', 'Capital Projects'];
const LEDGER_STATUSES = ['Pending', 'Approved', 'Posted', 'Void'];

function getLedgerSheet_() {
  const ss = getLogSheet_().getParent();
  let sh = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LEDGER_SHEET_NAME);
    sh.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    formatLedgerSheet_(sh);
  }
  return sh;
}

function formatLedgerSheet_(sh) {
  const n = 2000;
  sh.getRange(2, LC['Date'] + 1, n, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, LC['Amount (incl. tax)'] + 1, n, 4).setNumberFormat('#,##0.00');
  sh.getRange(2, LC['Last Updated'] + 1, n, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, LC['Entry ID'] + 1, n, 1).setNumberFormat('@');
  sh.getRange(2, LC['Claim ID'] + 1, n, 1).setNumberFormat('@');
  const list = function (values) { return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build(); };
  sh.getRange(2, LC['Status'] + 1, n, 1).setDataValidation(list(LEDGER_STATUSES));
  sh.getRange(2, LC['Locked'] + 1, n, 1).setDataValidation(list(['Yes']));
  sh.getRange(2, LC['Source'] + 1, n, 1).setDataValidation(list(['Claim', 'Manual']));
  sh.getRange(2, LC['Charge to Department (optional)'] + 1, n, 1).setDataValidation(list(TRACKER_DEPTS));
  const gl = ss_glListRange_(sh.getParent());
  sh.getRange(2, LC['GL Code'] + 1, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(gl, true).setAllowInvalid(true).build());
}

// "GL List" tab: the GL codes for the GL Code dropdown on Expense Ledger and the manual expense form.
// If the tab already holds the full list (4th column "Type" = Revenue / Expense, written by the claim log's own
// "GL Codes" script), it is kept: only claim-form codes missing from it are added. Otherwise it is rewritten with
// the claim form's codes. Returns the expense codes (the dropdown never offers revenue codes).
function ss_glListRange_(ss) {
  let sh = ss.getSheetByName('GL List');
  if (!sh) sh = ss.insertSheet('GL List');
  const rows = [];
  GL.forEach(function (g) { g.items.forEach(function (it) { rows.push([it[1], it[0], g.g, 'Expense']); }); });
  if (!rows.some(function (r) { return r[0] === MILEAGE_GL; })) rows.push([MILEAGE_GL, MILEAGE_GL_NAME, glGroup_(MILEAGE_GL) || MILEAGE_GL_GROUP, 'Expense']);
  const data = sh.getDataRange().getValues();
  if (data.length && String(data[0][3] || '').trim() === 'Type') {
    const have = {};
    data.slice(1).forEach(function (r) { have[String(r[0]).trim()] = true; });
    const add = rows.filter(function (r) { return !have[r[0]]; });
    if (add.length) sh.getRange(data.length + 1, 1, add.length, 4).setValues(add);
    return data.slice(1).filter(function (r) { return r[0] !== '' && String(r[3]).trim() !== 'Revenue'; })
      .map(function (r) { return String(r[0]).trim(); }).concat(add.map(function (r) { return r[0]; }));
  }
  sh.clearContents();
  sh.getRange(1, 1, rows.length + 1, 4).setValues([['GL Code', 'GL Name', 'Group', 'Type']].concat(rows));
  sh.getRange(1, 1, 1, 4).setFontWeight('bold');
  sh.setFrozenRows(1);
  return rows.map(function (r) { return r[0]; });
}

function glGroup_(code) {
  for (var i = 0; i < GL.length; i++) { var it = GL[i].items; for (var j = 0; j < it.length; j++) { if (it[j][1] === code) return GL[i].g; } }
  return '';
}

function ledgerStatus_(claimStatus) {
  const s = String(claimStatus || '');
  if (s.indexOf('Approved') === 0) return 'Approved'; // FINAL_STATUS, and pre-V4 "Approved & sent to AP"
  if (s.indexOf('Sent back') === 0) return 'Void';
  return 'Pending';
}

function round2_(n) { return Math.round((parseFloat(n) || 0) * 100) / 100; }

function parseDay_(v, fallback) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : (fallback || new Date());
}

// Build one ledger row per expense line of a claim.
function ledgerRowsFor_(claim, payload, claimStatus) {
  const now = new Date();
  const mileage = isMileage_(payload);
  return (payload.lines || []).map(function (l, i) {
    let amount, gst = 0, hst = 0, net, desc;
    if (mileage) {
      amount = round2_(round2_((parseFloat(l.km) || 0) * MILEAGE_RATE) + (parseFloat(l.parking) || 0));
      net = amount;
      desc = 'Mileage: ' + [l.destination, l.purpose].filter(Boolean).join(' - ') + ' (' + (parseFloat(l.km) || 0) + ' km)';
    } else {
      amount = round2_(l.amount); gst = round2_(l.gst); hst = round2_(l.hst);
      net = (typeof l.net === 'number') ? round2_(l.net) : round2_(amount - gst * 0.67 - hst * 0.7377);
      desc = l.desc || '';
    }
    const code = mileage ? (l.gl || MILEAGE_GL) : (l.gl || '');
    const row = new Array(LEDGER_HEADERS.length).fill('');
    row[LC['Entry ID']] = claim.id + '-' + (i + 1);
    row[LC['Source']] = 'Claim';
    row[LC['Date']] = parseDay_(l.date, now);
    row[LC['GL Code']] = code;
    row[LC['GL Name']] = mileage ? (l.glName || MILEAGE_GL_NAME) : (l.glName || glLabel_(code));
    row[LC['GL Group']] = glGroup_(code) || (mileage && code === MILEAGE_GL ? MILEAGE_GL_GROUP : '');
    row[LC['Description']] = desc;
    row[LC['Payee']] = claim.name || '';
    row[LC['Amount (incl. tax)']] = amount;
    row[LC['GST']] = gst;
    row[LC['HST']] = hst;
    row[LC['Net Expense']] = net;
    row[LC['Status']] = ledgerStatus_(claimStatus);
    row[LC['Claim ID']] = claim.id;
    row[LC['Line #']] = i + 1;
    row[LC['Claim Status']] = claimStatus;
    row[LC['Last Updated']] = now;
    return row;
  });
}

// Insert or refresh a claim's rows. Locked rows keep everything except Claim Status / Last Updated;
// "Charge to Department", "Locked" and "Notes" are never overwritten.
function writeLedgerRows_(claim, payload, claimStatus) {
  const sh = getLedgerSheet_();
  const data = sh.getDataRange().getValues();
  const at = {};
  for (let r = 1; r < data.length; r++) { if (data[r][LC['Entry ID']] !== '') at[String(data[r][LC['Entry ID']])] = r; }
  const keep = [LC['Charge to Department (optional)'], LC['Locked'], LC['Notes']];
  ledgerRowsFor_(claim, payload, claimStatus).forEach(function (row) {
    const r = at[String(row[LC['Entry ID']])];
    if (r === undefined) { sh.appendRow(row); return; }
    const old = data[r];
    let out;
    if (String(old[LC['Locked']]).toLowerCase() === 'yes') {
      out = old.slice(); out[LC['Claim Status']] = row[LC['Claim Status']]; out[LC['Last Updated']] = row[LC['Last Updated']];
    } else {
      out = row.slice(); keep.forEach(function (k) { out[k] = old[k]; });
    }
    sh.getRange(r + 1, 1, 1, LEDGER_HEADERS.length).setValues([out.slice(0, LEDGER_HEADERS.length)]); // only A–T: extra columns after Notes are yours
  });
}

// Called on every status change. Updates existing rows; creates them if the claim has GL codes but no rows yet.
function syncLedgerStatus_(id, claimStatus) {
  const sh = getLedgerSheet_();
  const data = sh.getDataRange().getValues();
  const now = new Date();
  let found = false;
  for (let r = 1; r < data.length; r++) {
    if (!sameId_(data[r][LC['Claim ID']], id)) continue;
    found = true;
    sh.getRange(r + 1, LC['Claim Status'] + 1).setValue(claimStatus);
    sh.getRange(r + 1, LC['Last Updated'] + 1).setValue(now);
    if (String(data[r][LC['Locked']]).toLowerCase() !== 'yes') sh.getRange(r + 1, LC['Status'] + 1).setValue(ledgerStatus_(claimStatus));
  }
  if (!found && stageOf_(claimStatus) !== 'vp1' && String(claimStatus).indexOf('Sent back') !== 0) {
    const claim = findClaim_(id);
    if (claim && claim.jsonId) writeLedgerRows_(claim, loadPayload_(claim.jsonId), claimStatus);
  }
}

// Never let a ledger problem block an approval: log it and carry on.
function safeLedger_(fn) {
  try { fn(); } catch (e) { console.error('Expense Ledger: ' + (e && e.stack || e)); }
}

// RUN ONCE: creates the tab + GL list, back-fills every claim already past the VP signature,
// and logs the private feed link for the Excel tracker. Safe to run again (it only refreshes rows
// and keeps the same link).
// Run once after changing the GL list above: rewrites the "GL List" tab (used by the manual expense form)
// and the GL Code dropdown on Expense Ledger. Doesn't touch any ledger rows.
function refreshGlList() {
  ownerOnly_();
  formatLedgerSheet_(getLedgerSheet_());
  console.log('GL List updated.');
}

function setupExpenseLedger() {
  ownerOnly_();
  const sh = getLedgerSheet_();
  formatLedgerSheet_(sh);
  const n = backfillExpenseLedger();
  const url = ledgerFeedUrl_();
  console.log('Expense Ledger ready. Claims written: ' + n + '. Paste this link into the Excel tracker (Settings sheet, B3): ' + url);
  return url;
}

/* ---------- Private feed: the Expense Ledger as CSV, for the Excel tracker ----------
 * <web app /exec URL>?feed=ledger&key=<secret>. Returns only the Expense Ledger tab (no emails,
 * no approval tokens, nothing from the Claims tab), so the claim log itself can stay Restricted.
 * The key lives in Script Properties. If the link ever leaks, run rotateLedgerFeedKey() and paste
 * the new link into the tracker; the old link stops working immediately. */
function ledgerFeedKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('LEDGER_FEED_KEY');
  if (!key) { key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); props.setProperty('LEDGER_FEED_KEY', key); }
  return key;
}
function ledgerFeedUrl_() { return webAppUrl_() + '?feed=ledger&key=' + ledgerFeedKey_(); }

function rotateLedgerFeedKey() {
  ownerOnly_();
  PropertiesService.getScriptProperties().deleteProperty('LEDGER_FEED_KEY');
  const url = ledgerFeedUrl_();
  console.log('New Expense Ledger link (the old one no longer works). Paste it into the Excel tracker (Settings sheet, B3): ' + url);
  return url;
}

function ledgerFeed_(key) {
  const expected = PropertiesService.getScriptProperties().getProperty('LEDGER_FEED_KEY');
  if (!expected || String(key || '') !== expected) {
    return ContentService.createTextOutput('Forbidden: this Expense Ledger link is not valid.').setMimeType(ContentService.MimeType.TEXT);
  }
  const tz = Session.getScriptTimeZone();
  const values = getLedgerSheet_().getDataRange().getValues();
  const csv = values.map(function (row) {
    return row.slice(0, LEDGER_HEADERS.length).map(function (v) { return csvCell_(v, tz); }).join(',');
  }).join('\r\n');
  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

function csvCell_(v, tz) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    const hasTime = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, tz, hasTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  }
  if (typeof v === 'number') return String(v);
  return '"' + String(v).replace(/"/g, '""') + '"';
}

function backfillExpenseLedger() {
  ownerOnly_();
  const data = getLogSheet_().getDataRange().getValues();
  let n = 0;
  for (let r = 1; r < data.length; r++) {
    const status = String(data[r][2] || '');
    const jsonId = data[r][14];
    if (!jsonId || stageOf_(status) === 'vp1' || status.indexOf('Sent back') === 0) continue;
    try {
      const payload = loadPayload_(jsonId);
      writeLedgerRows_({ id: data[r][0], name: data[r][4] }, payload, status);
      n++;
    } catch (e) { console.error('Could not back-fill claim ' + data[r][0] + ': ' + e); }
  }
  return n;
}


/* ============================================================
 *  ONE-TIME REPAIR — run once to bring the ledger to the current layout:
 *    ID, Timestamp, Status, Payable to, Name, Email, Total, GST, HST,
 *    Folder, PDF file id, Token, Decided by, Decided at, JSON file id, Form
 *
 *  The ledger has held three different column layouts over time:
 *    (A) V1   — an extra 'PST' column at index 8, and 'HST' near the end (15).
 *    (B) mid  — 'PST' dropped, 'HST' near the end (14), Folder at index 8.
 *    (C) new  — 'HST' moved up next to GST (index 8), Folder at index 9.
 *  This detects each old row by where its Folder URL sits and rewrites it into
 *  layout (C) in place, so every claim lines up under the correct headers.
 *
 *  SAFE to run more than once — rows already in the new layout are skipped.
 *  Run it from the editor: pick migrateLegacyRows and press Run. (No re-deploy
 *  needed.) TIP: make a copy of the ledger sheet first (File > Make a copy).
 * ============================================================ */
function migrateLegacyRows() {
  ownerOnly_();
  const sh = getLogSheet_();
  const data = sh.getDataRange().getValues();
  const isUrl = function (v) { return String(v || '').indexOf('http') === 0; };
  let fixed = 0;
  for (let r = 1; r < data.length; r++) {
    const d = data[r];
    if (!d[0]) continue;
    let n = null;
    if (isUrl(d[8])) {
      // (B) mid layout: Folder at 8, HST at 14. Move HST up to index 8.
      n = [d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7],
           d[14], d[8], d[9], d[10], d[11], d[12], d[13], d[15]];
    } else if (isUrl(d[9]) && String(d[16] || '') !== '') {
      // (A) V1 layout: PST at 8, Folder at 9, HST at 15, Form at 16.
      // Drop PST, move HST up to index 8, pull Form back to index 15.
      n = [d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[7],
           d[15], d[9], d[10], d[11], d[12], d[13], d[14], d[16]];
    } else {
      continue; // already in the new (C) layout, or unrecognized — leave it
    }
    sh.getRange(r + 1, 1, 1, n.length).setValues([n]);
    if (sh.getLastColumn() >= 17) sh.getRange(r + 1, 17).clearContent(); // drop the orphaned trailing cell
    fixed++;
  }
  // Bring the header row to the current layout.
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  if (sh.getLastColumn() >= 17) sh.getRange(1, 17).clearContent();
  SpreadsheetApp.flush();
  console.log('Claims re-aligned to the current layout: ' + fixed);
  return fixed;
}

/* ============================================================
 *  Signer access to the claim folder
 *
 *  Claim files are owned by the script owner and are otherwise private, so a
 *  signer can't open the packet or receipts unless they're shared. Every signer
 *  (VP Finance, President, NU Staff) gets view access to the claim's whole Drive
 *  folder (packet PDF, every uploaded receipt, claim data) when the claim is sent
 *  to them, so they can review everything before signing. Files added to the
 *  folder later (re-signed PDFs) are covered by the folder access; the PDF is also
 *  shared directly. The claims root folder is never shared (it holds the log).
 *  DriveApp.addViewer does not send a notification email, so it's silent.
 * ============================================================ */
function sharePacketWithPresident_(pdfId) { sharePacketWith_(pdfId, CONFIG.PRESIDENT_EMAIL); }

function sharePacketWith_(pdfId, emails) {
  try {
    const file = DriveApp.getFileById(pdfId);
    let folder = null, owner = '';
    try { folder = file.getParents().next(); } catch (err) { console.warn('No folder for ' + pdfId + ': ' + err); }
    try { owner = String(file.getOwner().getEmail() || '').toLowerCase(); } catch (err) { /* unknown owner */ }
    String(emails || '').split(',').forEach(function (e) {
      e = e.trim();
      if (!e || e.toLowerCase() === owner) return; // the owner already has full access
      try { if (folder) folder.addViewer(e); } catch (err) { console.warn('Could not share folder with ' + e + ': ' + err); }
      try { file.addViewer(e); } catch (err) { console.warn('Could not share with ' + e + ': ' + err); }
    });
  } catch (err) {
    console.warn('sharePacketWith_ failed for ' + pdfId + ': ' + err);
  }
}

// Every signer a claim has reached so far (VP Finance always; the President once VP Finance
// has signed). V5: NU Staff gets the PDF from VP Finance by NU email, so no Drive access is
// given to the NU Staff address. Sent-back claims get nobody new.
function signersSoFar_(status) {
  status = String(status || '');
  if (status.indexOf('Sent back') === 0) return [];
  const st = stageOf_(status);
  const out = [CONFIG.VP_FINANCE_EMAIL];
  if (st !== 'vp1') out.push(CONFIG.PRESIDENT_EMAIL);
  return out;
}

/* ONE-TIME: give every signer access to the folders of claims ALREADY in the system
 * (the automatic sharing above only fires on future routings). Each claim is shared
 * with the signers it has reached so far. Run it once from the editor. Safe to re-run. */
function shareCurrentClaimFolders() {
  ownerOnly_();
  const data = getLogSheet_().getDataRange().getValues();
  let shared = 0;
  for (let r = 1; r < data.length; r++) {
    const row = findClaim_(data[r][0]);
    const who = row && row.pdfId ? signersSoFar_(row.status) : [];
    if (who.length) { sharePacketWith_(row.pdfId, who.join(',')); shared++; }
  }
  console.log('Claim folders shared with their signers: ' + shared);
  return shared;
}
// Older name, kept so it still shows in the Run dropdown.
function shareCurrentPresidentPackets() { return shareCurrentClaimFolders(); }


/* ============================================================
 *  ONE-TIME: HEIC receipt photos in claims submitted before the form converted them
 *
 *  listHeicReceipts()    - read-only: logs every claim (except "Sent back") holding HEIC photos.
 *  convertHeicReceipts() - converts each one to JPEG, saves the JPEG in the claim's Drive folder
 *                          (the HEIC original is kept), updates the claim and rebuilds its PDF so the
 *                          receipt shows. Nothing is emailed; tokens/links stay the same.
 *  Conversion uses Google Drive's own preview of the photo (up to 1600 px). If Drive hasn't made
 *  one yet, run it again in a few minutes. Manual fallback: put a JPG/PNG with the same name
 *  (e.g. IMG_1562.jpg next to IMG_1562.HEIC) in the claim's folder and run it again.
 *  Safe to re-run: converted photos are no longer HEIC, so they're skipped.
 * ============================================================ */
function listHeicReceipts() { ownerOnly_(); return heicReceiptsPass_(false); }
function convertHeicReceipts() { ownerOnly_(); return heicReceiptsPass_(true); }

function heicReceiptsPass_(convert) {
  const started = Date.now();
  const data = getLogSheet_().getDataRange().getValues();
  const lines = [];
  let found = 0, fixed = 0, failed = 0, stoppedEarly = false;
  for (let r = 1; r < data.length; r++) {
    const status = String(data[r][2] || '');
    if (!data[r][0] || !data[r][14] || status.indexOf('Sent back') === 0) continue;
    let payload;
    try { payload = loadPayload_(data[r][14]); } catch (e) { continue; }
    const items = heicItems_(payload);
    if (!items.length) continue;
    found += items.length;
    const label = 'Claim ' + data[r][0] + ' (' + data[r][4] + ', ' + status + ')';
    if (!convert) { lines.push(label + ': ' + items.length + ' HEIC photo(s): ' + items.map(function (i) { return i.item.fileName; }).join(', ')); continue; }
    if (Date.now() - started > 4.5 * 60000) { stoppedEarly = true; break; } // stay under the 6-minute limit
    const row = findClaim_(data[r][0]);
    let folder;
    try { folder = DriveApp.getFileById(row.pdfId).getParents().next(); } catch (e) { failed += items.length; lines.push(label + ': claim folder not found'); continue; }
    let changed = 0;
    items.forEach(function (it) {
      try {
        const got = heicToImage_(it.item, folder), img = got.blob;
        const ext = /png/i.test(img.getContentType()) ? '.png' : '.jpg';
        const newName = String(it.item.fileName || 'receipt').replace(/\.[A-Za-z0-9]+$/, '') + ext;
        if (!got.fromFolder) folder.createFile(img.setName(newName)); // keep a copy with the claim
        it.item.fileData = 'data:' + img.getContentType() + ';base64,' + Utilities.base64Encode(img.getBytes());
        it.item.fileType = img.getContentType();
        it.item.fileName = newName;
        changed++; fixed++;
      } catch (e) {
        failed++;
        lines.push(label + ': could not convert ' + it.item.fileName + ' - ' + (e && e.message || e));
      }
    });
    if (changed) {
      savePayload_(row.jsonId, payload);
      rebuildAndSavePdf_(payload, row);
      const st = stageOf_(row.status);
      if (st === 'president') sharePacketWithPresident_(row.pdfId); // the rebuilt PDF is a new file
      lines.push(label + ': converted ' + changed + ' photo(s). PDF: https://drive.google.com/file/d/' + row.pdfId + '/view');
    }
  }
  const head = convert
    ? 'HEIC photos converted: ' + fixed + (failed ? ', not converted: ' + failed : '') + (stoppedEarly ? ' (time limit reached - run convertHeicReceipts again to continue)' : '')
    : 'HEIC photos found: ' + found + (found ? ' - run convertHeicReceipts to convert them.' : '');
  console.log([head].concat(lines).join('\n'));
  return { found: found, fixed: fixed, failed: failed, stoppedEarly: stoppedEarly, details: lines };
}

// Every uploaded file in a claim that is a HEIC photo (by name, label, or the file's own header).
function heicItems_(payload) {
  const out = [];
  (payload.lines || []).concat(payload.attachments || []).forEach(function (item) {
    if (item && item.fileData && !pdfEmbeddable_(item.fileData) && looksHeic_(item.fileName, item.fileData)) out.push({ item: item });
  });
  return out;
}
function looksHeic_(name, dataUrl) {
  const s = String(dataUrl || '');
  if (/\.(heic|heif)$/i.test(String(name || '')) || /^data:image\/hei[cf]/i.test(s)) return true;
  const i = s.indexOf(',');
  if (i < 0) return false;
  try { // HEIC/HEIF files carry "ftyp" + a brand such as heic/heix/mif1 at byte 4
    const head = Utilities.base64Decode(s.slice(i + 1, i + 17)).map(function (b) { return String.fromCharCode(b & 255); }).join('');
    return head.slice(4, 8) === 'ftyp' && /^(heic|heix|hevc|heim|heis|hevm|hevs|mif1|msf1)$/.test(head.slice(8, 12));
  } catch (e) { return false; }
}

// HEIC -> {blob: JPEG/PNG, fromFolder}. 1) a same-named JPG/PNG already in the claim folder (manual fallback);
// 2) otherwise Google Drive's preview image of the photo, fetched at up to 1600 px.
function heicToImage_(item, folder) {
  const base = String(item.fileName || '').replace(/\.[A-Za-z0-9]+$/, '').toLowerCase();
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const m = f.getName().match(/^(.*)\.(jpe?g|png)$/i);
    if (m && m[1].toLowerCase() === base && /^image\/(jpeg|png)$/.test(f.getMimeType())) {
      return { blob: f.getBlob(), fromFolder: true };
    }
  }
  const b64 = String(item.fileData).slice(String(item.fileData).indexOf(',') + 1);
  const tmp = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/heic', 'convert-' + (base || 'receipt') + '.heic'));
  try {
    const auth = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    let link = '';
    for (let i = 0; i < 9 && !link; i++) { // Drive makes the preview within seconds, sometimes longer
      if (i) Utilities.sleep(5000);
      const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + tmp.getId() + '?fields=thumbnailLink', auth);
      if (res.getResponseCode() === 200) link = JSON.parse(res.getContentText()).thumbnailLink || '';
    }
    if (!link) throw new Error('Google Drive has not made a preview of this photo yet - try again in a few minutes');
    const img = UrlFetchApp.fetch(/=s\d+$/.test(link) ? link.replace(/=s\d+$/, '=s1600') : link, auth);
    const blob = img.getBlob();
    if (img.getResponseCode() !== 200 || !/^image\/(jpeg|png)/i.test(String(blob.getContentType()))) {
      throw new Error('could not download the converted image (HTTP ' + img.getResponseCode() + ', ' + blob.getContentType() + ')');
    }
    return { blob: blob, fromFolder: false };
  } finally {
    try { tmp.setTrashed(true); } catch (e) { /* keep going */ }
  }
}


/* ============================================================
 *  HTML builders (PDF packet, emails, decision page)
 * ============================================================ */
function buildPacketHtml_(p, id, imageReceipts) {
  const t = p.totals || {};
  const esc = escapeHtml_;
  const dateStr = Utilities.formatDate(submittedDate_(p), Session.getScriptTimeZone(), 'MM/dd/yy');

  // small check-box with an X when selected
  function box(on) {
    return '<span style="display:inline-block;width:11px;height:11px;border:1px solid #000;'
      + 'text-align:center;line-height:10px;font-size:9px;vertical-align:middle">' + (on ? 'X' : '&nbsp;') + '</span>';
  }
  // a labelled fill-in field: label + underlined value
  function fld(label, val) {
    return '<tr><td style="width:120px;padding:2px 6px 2px 0;white-space:nowrap;vertical-align:bottom">' + label + '</td>'
      + '<td style="border-bottom:1px solid #000;padding:1px 4px;vertical-align:bottom">' + esc(val || '') + '</td></tr>';
  }

  // expense rows (pad to a minimum of 6 for the form look)
  const bodyRows = (p.lines || []).slice();
  const cR = 'border:1px solid #000;padding:4px;text-align:right';
  let rows = bodyRows.map(function (l) {
    return '<tr>'
      + '<td style="border:1px solid #000;padding:4px;text-align:center">' + esc(l.date || '') + '</td>'
      + '<td style="' + cR + '">$' + money_(l.amount) + '</td>'
      + '<td style="' + cR + '">$' + money_(l.gst) + '</td>'
      + '<td style="' + cR + '">$' + money_(l.hst) + '</td>'
      + '<td style="' + cR + '">$' + money_(l.net) + '</td>'
      + '<td style="border:1px solid #000;padding:4px">' + esc(l.desc || '') + '</td>'
      + '<td style="border:1px solid #000;padding:4px;font-family:monospace;font-size:8.5px;word-break:break-all">' + esc(l.gl || '') + '</td>'
      + '</tr>';
  }).join('');
  for (let i = bodyRows.length; i < 6; i++) {
    rows += '<tr><td style="border:1px solid #000;padding:4px;height:16px">&nbsp;</td>'
      + '<td style="border:1px solid #000"></td><td style="border:1px solid #000"></td>'
      + '<td style="border:1px solid #000"></td><td style="border:1px solid #000"></td><td style="border:1px solid #000"></td><td style="border:1px solid #000"></td></tr>';
  }

  const th = 'border:1px solid #000;background:#e9e9e9;padding:5px 4px;font-size:9.5px;font-weight:bold;text-align:center';

  const table = '<table style="width:100%;border-collapse:collapse;font-size:10px;table-layout:fixed;margin-top:4px">'
    + '<colgroup><col style="width:9%"><col style="width:11%"><col style="width:9%"><col style="width:9%"><col style="width:11%"><col style="width:28%"><col style="width:23%"></colgroup>'
    + '<thead><tr>'
    + '<th style="' + th + '">Date<br>(mm/dd/yy)</th><th style="' + th + '">Total Amount</th><th style="' + th + '">Total GST</th>'
    + '<th style="' + th + '">Total HST</th><th style="' + th + '">Net Expense</th><th style="' + th + '">Type of Expense</th><th style="' + th + '">GL Acct #</th>'
    + '</tr></thead><tbody>' + rows
    + '<tr>'
    + '<td style="border:1px solid #000;padding:4px;font-weight:bold">Total</td>'
    + '<td style="' + cR + ';font-weight:bold">$' + money_(t.amt) + '</td>'
    + '<td style="' + cR + ';font-weight:bold">$' + money_(t.gst) + '</td>'
    + '<td style="' + cR + ';font-weight:bold">$' + money_(t.hst) + '</td>'
    + '<td style="' + cR + ';font-weight:bold">$' + money_(t.net) + '</td>'
    + '<td style="border:1px solid #000"></td><td style="border:1px solid #000"></td>'
    + '</tr></tbody></table>';

  // advance + rebate section (mirrors the official form's formulas)
  const gstReb = (typeof t.gstReb === 'number') ? t.gstReb : Math.round((parseFloat(t.gst) || 0) * 0.67 * 100) / 100;
  const hstReb = (typeof t.hstReb === 'number') ? t.hstReb : Math.round((parseFloat(t.hst) || 0) * 0.7377 * 100) / 100;
  const uln = '<span style="display:inline-block;border-bottom:1px solid #000;min-width:110px">&nbsp;</span>';
  const rbc = 'border:1px solid #000;padding:3px 7px';
  const advance = '<table style="width:100%;font-size:10px;border-collapse:collapse;margin-top:6px"><tr>'
    + '<td style="width:52%;vertical-align:top;padding-right:16px">'
    + '<div style="padding:2px 0 6px">Less: Advance: ' + uln + '</div>'
    + '<table style="border-collapse:collapse;font-size:10px">'
    + '<tr><td style="' + rbc + ';font-weight:bold">GST Rebate</td><td style="' + rbc + ';text-align:right">$' + money_(gstReb) + '</td><td style="' + rbc + ';text-align:right">67.00%</td></tr>'
    + '<tr><td style="' + rbc + ';font-weight:bold">HST Rebate</td><td style="' + rbc + ';text-align:right">$' + money_(hstReb) + '</td><td style="' + rbc + ';text-align:right">73.77%</td></tr>'
    + '</table></td>'
    + '<td style="width:48%;vertical-align:top">'
    + '<div style="padding:2px 0">Clearing Advance: &nbsp; N ' + box(false) + ' &nbsp; Y ' + box(false) + '</div>'
    + '<div style="padding:10px 0 0">Reimbursement Refund Due Office: ' + uln + '</div>'
    + '</td></tr></table>';

  // signatures (Requested by = claimant; Approved by = VP Finance, then President, then NU Staff)
  const signatures = signaturesHtml_(p);

  const delivery = '<div style="font-size:11px;margin-top:16px">Indicate Cheque Delivery Method: &nbsp; '
    + box(/mail/i.test(p.delivery || '')) + ' Mail to the above address &nbsp;&nbsp;&nbsp; '
    + box(/pick/i.test(p.delivery || '')) + ' Pick Up</div>';

  const receiptPages = imageReceipts.map(function (r, i) {
    return '<div style="page-break-before:always;padding-top:16px">'
      + '<div style="font:12px sans-serif;color:#666;margin-bottom:8px">Receipt ' + (r.no || (i + 1)) + ' - ' + esc(r.name || '') + '</div>'
      + '<img src="' + safeImageSrc_(r.dataUrl) + '" style="max-width:100%;max-height:920px">'
      + '</div>';
  }).join('');

  return '<html><head><meta charset="utf-8"><style>'
    + 'body{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff;padding:26px;font-size:11px}'
    + '</style></head><body>'
    // header: date top-right, title band, instruction
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:4px"><tr>'
    + '<td style="vertical-align:top"><img src="' + NU_LOGO_DATAURI + '" style="width:180px;height:auto"></td>'
    + '<td style="vertical-align:top;text-align:right;font-size:11px;padding-top:8px">Date (mm/dd/yy): <b>' + esc(dateStr) + '</b></td>'
    + '</tr></table>'
    + '<div style="border:1.5px solid #000;text-align:center;font-size:20px;font-weight:bold;letter-spacing:1px;padding:8px 0">EXPENSE CLAIM FORM</div>'
    + '<div style="font-size:9px;font-style:italic;color:#333;margin:6px 0 10px">Scanned copies of receipts, bills, and invoices, containing the GST amounts should be combined with this form into a PDF file.</div>'
    // payable-to block
    + '<table style="width:100%;font-size:11px;border-collapse:collapse;margin-bottom:6px"><tr>'
    + '<td style="vertical-align:top;width:60%">Payable To: (check one) &nbsp; '
    + 'Employee ' + box(p.payTo === 'Employee') + ' &nbsp; Student ' + box(p.payTo === 'Student') + ' &nbsp; Other ' + box(p.payTo === 'Other') + '</td>'
    + '<td style="vertical-align:top;width:40%;padding-left:10px">Will be paid in CANADIAN funds unless checked below:<br>'
    + '<span style="display:inline-block;margin-top:4px">USD $ ' + box(p.usd) + (p.usd ? ' &nbsp;<b>$' + money_(t.amt) + '</b>' : '') + '</span></td>'
    + '</tr></table>'
    + '<table style="width:65%;font-size:11px;border-collapse:collapse;margin-bottom:8px">'
    + fld('Student ID#:', p.payTo === 'Student' ? p.studentId : '')
    + fld('Business Name:', p.payTo === 'Other' ? p.company : '')
    + fld('Or,  First Name:', p.first)
    + fld('Last Name:', p.last)
    + fld('Address:', p.address)
    + fld('City, Postal Code:', p.city)
    + fld('Country:', p.country)
    + '</table>'
    + table
    + advance
    + signatures
    + delivery
    + '<div style="margin-top:10px;font-size:8px;color:#999">Claim ' + esc(id) + ' &middot; Submitted ' + esc(Utilities.formatDate(submittedDate_(p), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')) + ' &middot; ' + esc(p.email) + '</div>'
    + receiptPages
    + '</body></html>';
}

// Signature block shared by both PDFs, in signing order: requester, VP Finance,
// President, NU Staff. Unsigned lines print blank; each line names its role.
function signaturesHtml_(p) {
  const esc = escapeHtml_;
  function sigCell(sig) {
    sig = safeImageSrc_(sig);
    return sig ? '<img src="' + sig + '" style="height:32px;vertical-align:middle">'
               : '<span style="display:inline-block;border-bottom:1px solid #000;width:170px">&nbsp;</span>';
  }
  function nameLine(name) {
    return '<span style="display:inline-block;border-bottom:1px solid #000;min-width:170px;padding:0 4px">' + (name ? esc(name) : '&nbsp;') + '</span>';
  }
  function sigRow(label, role, name, sig, top) {
    return '<tr><td style="width:50%;padding-top:' + top + 'px">' + label + ': ' + nameLine(name) + '</td>'
      + '<td style="width:50%;padding-top:' + top + 'px">Signature: ' + sigCell(sig) + '</td></tr>'
      + '<tr><td style="font-size:9px;color:#555">' + role + '</td><td></td></tr>';
  }
  return '<table style="width:100%;font-size:11px;border-collapse:collapse;margin-top:14px">'
    + sigRow('Requested by', 'Requester', p.printName, p.signature, 6)
    + sigRow('Approved by', esc(CONFIG.VP_FINANCE_TITLE), p.vpSignature ? (p.vpName || CONFIG.VP_FINANCE_NAME) : '', p.vpSignature, 12)
    + sigRow('Approved by', esc(CONFIG.PRESIDENT_TITLE), p.presidentSignature ? (p.presidentName || CONFIG.PRESIDENT_NAME) : '', p.presidentSignature, 12)
    + sigRow('Approved by', esc(CONFIG.STAFF_TITLE), p.staffSignature ? (p.staffName || CONFIG.STAFF_NAME) : '', p.staffSignature, 12)
    + '</table>';
}

function stageEmailHtml_(stage, ctx, pdfReceiptCount) {
  const url = decisionUrl_(ctx.id, ctx.token);
  let head, btn;
  const lbl = (ctx && ctx.form === 'mileage') ? 'kilometer reimbursement' : 'expense claim';
  if (stage === 'vp1') { head = 'A new ' + lbl + ' needs your signature and approval.'; btn = 'Review &amp; sign'; }
  else if (stage === 'president') { head = 'A ' + lbl + ' (already signed by VP Finance) needs your signature.'; btn = 'Review &amp; sign'; }
  else if (stage === 'staff') { head = 'A NUSA ' + lbl + ' (already signed by the NUSA VP Finance and President) needs your signature.'; btn = 'Review &amp; sign'; }
  else { head = 'A ' + lbl + ' has all signatures and is ready for your final approval.'; btn = 'Review &amp; approve'; }
  const extra = pdfReceiptCount
    ? '<p style="font-size:12px;color:#8a6d0b">' + pdfReceiptCount + ' receipt(s) were PDFs, attached separately.</p>' : '';
  return '<div style="font-family:sans-serif;max-width:520px">'
    + '<h2 style="color:' + CONFIG.BRAND_GREEN + '">' + head + '</h2>'
    + '<p><b>' + escapeHtml_(ctx.name) + '</b> &mdash; $' + money_(ctx.total) + '</p>' + extra
    + '<p style="margin:22px 0"><a href="' + url + '" style="background:' + CONFIG.BRAND_GREEN + ';color:#fff;text-decoration:none;padding:12px 22px;border-radius:8px;font-weight:bold">' + btn + ' &rarr;</a></p>'
    // Every signer has view access to the claim's Drive folder (receipts, PDF, claim data).
    + (ctx.folderUrl ? '<p style="font-size:12px;color:#666"><a href="' + ctx.folderUrl + '">Open the Drive folder (all receipts)</a></p>' : '')
    + '</div>';
}

function buildMileageHtml_(p, id, imageReceipts) {
  const t = p.totals || {};
  const esc = escapeHtml_;
  const dateStr = Utilities.formatDate(submittedDate_(p), Session.getScriptTimeZone(), 'MM/dd/yy');
  function box(on) { return '<span style="display:inline-block;width:11px;height:11px;border:1px solid #000;text-align:center;line-height:10px;font-size:9px;vertical-align:middle">' + (on ? 'X' : '&nbsp;') + '</span>'; }
  function fld(label, val) { return '<tr><td style="width:120px;padding:2px 6px 2px 0;white-space:nowrap;vertical-align:bottom">' + label + '</td><td style="border-bottom:1px solid #000;padding:1px 4px;vertical-align:bottom">' + esc(val || '') + '</td></tr>'; }
  const cR = 'border:1px solid #000;padding:4px;text-align:right';
  const cC = 'border:1px solid #000;padding:4px;text-align:center';
  const cL = 'border:1px solid #000;padding:4px';

  const bodyRows = (p.lines || []).slice();
  let rows = bodyRows.map(function (l) {
    return '<tr>'
      + '<td style="' + cC + '">' + esc(l.date || '') + '</td>'
      + '<td style="' + cR + '">' + esc(String(l.km || '')) + '</td>'
      + '<td style="' + cL + '">' + esc(l.destination || '') + '</td>'
      + '<td style="' + cL + '">' + esc(l.purpose || '') + '</td>'
      + '<td style="' + cR + '">$' + money_(l.parking) + '</td>'
      + '</tr>';
  }).join('');
  for (let i = bodyRows.length; i < 8; i++) {
    rows += '<tr><td style="' + cC + ';height:15px">&nbsp;</td><td style="border:1px solid #000"></td><td style="border:1px solid #000"></td><td style="border:1px solid #000"></td><td style="border:1px solid #000"></td></tr>';
  }
  const th = 'border:1px solid #000;background:#e9e9e9;padding:5px 4px;font-size:9.5px;font-weight:bold;text-align:center';
  const table = '<table style="width:100%;border-collapse:collapse;font-size:10px;table-layout:fixed;margin-top:4px">'
    + '<colgroup><col style="width:12%"><col style="width:10%"><col style="width:34%"><col style="width:26%"><col style="width:18%"></colgroup>'
    + '<thead><tr><th style="' + th + '">Date<br>(mm/dd/yy)</th><th style="' + th + '">Km\x27s</th><th style="' + th + '">Destination (from-to-to...)</th><th style="' + th + '">Purpose / Passengers</th><th style="' + th + '">Parking/Toll $<br>(attach receipts)</th></tr></thead>'
    + '<tbody>' + rows
    + '<tr><td style="' + cL + ';font-weight:bold">Total</td><td style="' + cR + ';font-weight:bold">' + esc(String(t.km || 0)) + '</td><td style="border:1px solid #000"></td><td style="border:1px solid #000"></td><td style="' + cR + ';font-weight:bold">$' + money_(t.parking) + '</td></tr>'
    + '</tbody></table>';

  const sum = '<table style="width:62%;margin-left:auto;border-collapse:collapse;font-size:10.5px;margin-top:8px">'
    + '<tr><td style="padding:3px 8px;text-align:right">Rate / km</td><td style="border:1px solid #000;padding:3px 8px;text-align:right;width:100px">$' + money_(MILEAGE_RATE) + '</td></tr>'
    + '<tr><td style="padding:3px 8px;text-align:right">Kilometer Reimbursement</td><td style="border:1px solid #000;padding:3px 8px;text-align:right">$' + money_(t.kmReimb) + '</td></tr>'
    + '<tr><td style="padding:3px 8px;text-align:right">Parking/Toll Reimbursement</td><td style="border:1px solid #000;padding:3px 8px;text-align:right">$' + money_(t.parking) + '</td></tr>'
    + '<tr><td style="padding:4px 8px;text-align:right;font-weight:bold">TOTAL REIMBURSEMENT</td><td style="border:2px solid #000;padding:4px 8px;text-align:right;font-weight:bold">$' + money_(t.amt) + '</td></tr>'
    + '</table>';
  const glLine = '<div style="font-size:10px;margin-top:6px">GL Acct #: <span style="font-family:monospace">' + esc(mileageGl_(p)) + '</span></div>';

  const signatures = signaturesHtml_(p);
  const delivery = '<div style="font-size:11px;margin-top:14px">Indicate Cheque Delivery Method: &nbsp; '
    + box(/mail/i.test(p.delivery || '')) + ' Mailed to the above address &nbsp;&nbsp;&nbsp; '
    + box(/pick/i.test(p.delivery || '')) + ' Pickup from Financial Services</div>';
  const notes = '<div style="font-size:8.5px;color:#444;margin-top:10px;line-height:1.5">FOR CASH REIMBURSEMENT (UNDER $100) - BRING TO LOWER MATTSON CASHIER<br>FOR CHEQUE REIMBURSEMENT (OVER $100) - SEND TO AP - 3RD FLOOR REIMER</div>';
  const receiptPages = imageReceipts.map(function (r, i) { return '<div style="page-break-before:always;padding-top:16px"><div style="font:12px sans-serif;color:#666;margin-bottom:8px">Receipt ' + (r.no || (i + 1)) + ' - ' + esc(r.name || '') + '</div><img src="' + safeImageSrc_(r.dataUrl) + '" style="max-width:100%;max-height:920px"></div>'; }).join('');

  return '<html><head><meta charset="utf-8"><style>body{font-family:Arial,Helvetica,sans-serif;color:#000;background:#fff;padding:26px;font-size:11px}</style></head><body>'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:4px"><tr><td style="vertical-align:top"><img src="' + NU_LOGO_DATAURI + '" style="width:180px;height:auto"></td><td style="vertical-align:top;text-align:right;font-size:11px;padding-top:8px">Date (mm/dd/yy): <b>' + esc(dateStr) + '</b></td></tr></table>'
    + '<div style="border:1.5px solid #000;text-align:center;font-size:19px;font-weight:bold;letter-spacing:1px;padding:8px 0">NUSA KILOMETER REIMBURSEMENT FORM</div>'
    + '<table style="width:100%;font-size:11px;border-collapse:collapse;margin:8px 0 6px"><tr>'
    + '<td style="vertical-align:top">Payable To: (check one) &nbsp; Employee ' + box(p.payTo === 'Employee') + ' &nbsp; Student ' + box(p.payTo === 'Student') + ' &nbsp; Other ' + box(p.payTo === 'Other') + '</td></tr></table>'
    + '<table style="width:65%;font-size:11px;border-collapse:collapse;margin-bottom:8px">'
    + fld('Student ID #:', p.payTo === 'Student' ? p.studentId : '') + fld('Business Name:', p.payTo === 'Other' ? p.company : '') + fld('or,  First Name:', p.first) + fld('Last Name:', p.last) + fld('Address:', p.address) + fld('City, Postal Code:', p.city) + fld('Country:', p.country)
    + '</table>'
    + table + sum + glLine + signatures + delivery + notes
    + '<div style="margin-top:8px;font-size:8px;color:#999">Claim ' + esc(id) + ' &middot; Submitted ' + esc(Utilities.formatDate(submittedDate_(p), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')) + ' &middot; ' + esc(p.email) + '</div>'
    + receiptPages + '</body></html>';
}

// Emailed to VP Finance on final approval, with the final signed packet attached.
// The summary is written so it can be copied straight into the email to Accounts Payable.
// signedName (V5): the name of NU Staff's signed PDF, saved in the claim's folder.
function finalPacketEmailHtml_(row, separateCount, driveOnlyCount, signedName) {
  const pdfUrl = 'https://drive.google.com/file/d/' + encodeURIComponent(String(row.pdfId)) + '/view';
  return '<div style="font-family:sans-serif;max-width:560px">'
    + '<h2 style="color:' + CONFIG.BRAND_GREEN + '">Final ' + formLabelLc_(row) + ' &mdash; ready for Accounts Payable</h2>'
    + (signedName
        ? '<p>' + escapeHtml_(CONFIG.STAFF_NAME) + ' (NU Staff) signed the PDF and sent it back. It is saved in the claim&#39;s folder as <b>'
          + escapeHtml_(signedName) + '</b> and attached here'
        : '<p>All four signatures are in (requester, VP Finance, President, NU Staff). The final signed PDF is attached')
    + (separateCount ? ', plus ' + separateCount + ' receipt file(s) that are not inside the PDF' : '') + '.</p>'
    + (driveOnlyCount ? '<p style="color:#8a6d0b"><b>' + driveOnlyCount + ' receipt file(s) were too large to attach.</b> '
      + 'Get them from the claim&#39;s Drive folder (link below) and include them for Accounts Payable.</p>' : '')
    + '<p><b>Next step:</b> forward this email (with its attachments) to NU Accounts Payable from your NU email. '
    + 'The system does not email Accounts Payable directly.</p>'
    + '<p style="font-size:13px"><a href="' + pdfUrl + '">Open the final PDF in Drive</a>'
    + (row.folderUrl ? ' &middot; <a href="' + row.folderUrl + '">Open the claim&#39;s Drive folder</a>' : '') + '</p>'
    + '<hr style="border:none;border-top:1px solid #ddd;margin:18px 0">'
    + '<p style="font-size:12px;color:#666">Summary you can paste into your email to Accounts Payable:</p>'
    + '<table style="font-size:14px;margin:8px 0">'
    + '<tr><td style="color:#666;padding:2px 12px 2px 0">Payable to</td><td><b>' + escapeHtml_(row.name) + '</b> (' + escapeHtml_(row.payTo) + ')</td></tr>'
    + '<tr><td style="color:#666;padding:2px 12px 2px 0">Amount</td><td><b>$' + money_(row.total) + '</b></td></tr>'
    + (row.form === 'mileage' ? '' : '<tr><td style="color:#666;padding:2px 12px 2px 0">GST / HST</td><td>$' + money_(row.gst) + ' / $' + money_(row.hst) + '</td></tr>')
    + '<tr><td style="color:#666;padding:2px 12px 2px 0">Approved by</td><td>' + escapeHtml_(CONFIG.VP_FINANCE_NAME) + ' (VP Finance), ' + escapeHtml_(CONFIG.PRESIDENT_NAME) + ' (President), ' + escapeHtml_(CONFIG.STAFF_NAME) + ' (NU Staff)</td></tr>'
    + '<tr><td style="color:#666;padding:2px 12px 2px 0">Claim ID</td><td>' + escapeHtml_(String(row.id)) + '</td></tr>'
    + '</table></div>';
}

function decisionPageHtml_(row, stage, payload) {
  const g = CONFIG.BRAND_GREEN, cream = CONFIG.BRAND_CREAM;
  const lines = (payload && payload.lines) || [];
  const returnNote = (payload && (stage === 'staffreturned' ? payload.staffReturnNote : payload.presidentReturnNote)) || '';
  const isMil = !!(payload && payload.form === 'mileage');
  const pdfUrl = 'https://drive.google.com/file/d/' + escapeHtml_(String(row.pdfId)) + '/view';
  const first = escapeHtml_(String(row.name).split(' ')[0]);

  const pad =
      '<div style="margin-top:16px"><div style="font-size:13px;font-weight:bold;margin-bottom:6px">Draw your signature</div>'
    + '<canvas id="pad" width="440" height="140" style="border:1px solid #CBBCA6;border-radius:8px;background:#FBFAF7;width:100%;max-width:440px;height:140px;touch-action:none;display:block"></canvas>'
    + '<button class="btn back" style="margin-top:6px;padding:6px 12px;font-size:13px" onclick="clearPad()">Clear</button></div>';
  const sendBack =
      '<div style="margin-top:22px;border-top:1px solid #E4DACB;padding-top:14px">'
    + '<div style="font-size:13px;color:#5E6B60">Or send it back to ' + first + ' to fix:</div>'
    + '<textarea id="note" placeholder="What needs fixing?"></textarea>'
    + '<button class="btn back" style="margin-top:8px" onclick="doSendBack()">Send back to ' + first + '</button></div>';
  const returnVp =
      '<div style="margin-top:22px;border-top:1px solid #E4DACB;padding-top:14px">'
    + '<div style="font-size:13px;color:#5E6B60">Or return it to VP Finance with a note:</div>'
    + '<textarea id="note" placeholder="Reason for returning"></textarea>'
    + '<button class="btn back" style="margin-top:8px" onclick="doReturn()">Return to VP Finance</button></div>';

  let heading, sub, actions = '';
  if (stage === 'vp1' && isMil) {
    heading = 'Review, sign &amp; approve &mdash; VP Finance';
    sub = 'Review the trips, draw your signature, and approve. It then goes to the President, then to NU Staff, for their signatures.';
    let trips = '<div style="margin-top:16px"><div style="font-size:13px;font-weight:bold;margin-bottom:6px">Trips</div>';
    (lines || []).forEach(function (l) {
      trips += '<div style="border:1px solid #E4DACB;border-radius:8px;padding:9px 10px;margin-bottom:6px;font-size:12.5px">'
        + '<b>' + escapeHtml_(String(l.km || '')) + ' km</b> &middot; ' + escapeHtml_(l.date || '') + '<br>' + escapeHtml_(l.destination || '')
        + (l.purpose ? ' &mdash; ' + escapeHtml_(l.purpose) : '') + (parseFloat(l.parking) ? ' &middot; parking $' + money_(l.parking) : '') + '</div>';
    });
    trips += '</div>';
    actions = trips + pad + '<button class="btn approve" style="margin-top:14px" onclick="doSign()">Sign &amp; approve</button>' + sendBack;
  } else if (stage === 'vp1') {
    heading = 'Assign GL codes, sign &amp; approve &mdash; VP Finance';
    sub = 'Assign a GL code to each expense, draw your signature, and approve. It then goes to the President, then to NU Staff, for their signatures.';
    let glAssign = '<div style="margin-top:16px"><div style="font-size:13px;font-weight:bold;margin-bottom:6px">Assign a GL code to each expense</div>';
    (lines || []).forEach(function (l, i) {
      glAssign += '<div style="border:1px solid #E4DACB;border-radius:8px;padding:10px;margin-bottom:8px">'
        + '<div style="font-size:12.5px;margin-bottom:6px"><b>$' + money_(l.amount) + '</b> &middot; ' + escapeHtml_(l.date || '') + '<br>' + escapeHtml_(l.desc || '') + '</div>'
        + '<select class="gl" data-i="' + i + '" style="width:100%;padding:8px;border:1px solid #CBBCA6;border-radius:8px;font:inherit;background:#fff">' + glOptionsHtml_(l.gl) + '</select></div>';
    });
    glAssign += '</div>';
    actions = glAssign + pad + '<button class="btn approve" style="margin-top:14px" onclick="doSign()">Sign &amp; approve</button>' + sendBack;
  } else if (stage === 'president') {
    heading = 'Sign &amp; approve &mdash; President';
    sub = 'VP Finance has signed this claim. Please review, draw your signature, and approve. It then goes to NU Staff for signature.';
    actions = pad + '<button class="btn approve" style="margin-top:14px" onclick="doSign()">Sign &amp; approve</button>' + returnVp;
  } else if (stage === 'staff') {
    heading = 'Sign &amp; approve &mdash; NU Staff';
    sub = 'The NUSA VP Finance and President have signed this claim. Please review, draw your signature, and approve.';
    actions = pad + '<button class="btn approve" style="margin-top:14px" onclick="doSign()">Sign &amp; approve</button>' + returnVp;
  } else if (stage === 'staffpdf') {
    heading = 'Waiting for NU Staff&#39;s signed PDF &mdash; VP Finance';
    sub = 'Send the PDF to ' + escapeHtml_(CONFIG.STAFF_NAME) + ' (' + escapeHtml_(CONFIG.STAFF_EMAIL) + ') from your NU email, using the subject and message in the email you received. '
      + 'When ' + escapeHtml_(String(CONFIG.STAFF_NAME).split(' ')[0]) + ' replies with the signed PDF, forward the reply to ' + escapeHtml_(scriptMailbox_())
      + ', keeping ' + escapeHtml_(claimTag_(row)) + ' in the subject. It is then saved in the claim&#39;s folder and emailed to you automatically.';
    actions = '<div style="margin-top:16px"><button class="btn approve" onclick="doResendStaff()">Email me the PDF and message again</button></div>' + sendBack;
  } else if (stage === 'vpfinal' && !(payload && payload.staffSignature)) {
    // Claims that reached this step before V4 have no NU Staff signature yet.
    heading = 'NU Staff signature needed &mdash; VP Finance';
    sub = 'This claim has the VP Finance and President signatures but not the NU Staff signature (' + escapeHtml_(CONFIG.STAFF_NAME) + '). Send it to NU Staff to sign before final approval.';
    actions = '<div style="margin-top:16px"><button class="btn approve" onclick="doResendStaff()">Send to NU Staff for signature</button></div>' + sendBack;
  } else if (stage === 'vpfinal') {
    heading = 'Final approval &mdash; VP Finance';
    sub = 'All signatures are in (VP Finance, President, NU Staff). Give final approval and the final signed PDF is emailed to you (' + escapeHtml_(CONFIG.VP_FINANCE_EMAIL) + ') to submit to Accounts Payable.';
    actions = '<button class="btn approve" style="margin-top:14px" onclick="doFinal()">Final approval &amp; email me the final PDF</button>' + sendBack;
  } else if (stage === 'staffreturned') {
    heading = 'Returned by NU Staff &mdash; VP Finance';
    sub = 'NU Staff (' + escapeHtml_(CONFIG.STAFF_NAME) + ') returned this claim. Review the note, then re-send it to NU Staff, or send it back to the student.';
    const staffNoteBox = returnNote
      ? '<div style="margin-top:14px;border-left:3px solid ' + g + ';background:#F3EEE4;padding:10px 12px;border-radius:6px;font-size:13px"><b>NU Staff note:</b><br>' + escapeHtml_(returnNote).replace(/\n/g, '<br>') + '</div>'
      : '';
    actions = staffNoteBox + '<div style="margin-top:16px"><button class="btn approve" onclick="doResendStaff()">Re-send to NU Staff</button></div>' + sendBack;
  } else if (stage === 'vpreturned') {
    heading = 'Returned by the President &mdash; VP Finance';
    sub = 'The President returned this claim. Review her note, then re-send it to her, or send it back to the student.';
    const noteBox = returnNote
      ? '<div style="margin-top:14px;border-left:3px solid ' + g + ';background:#F3EEE4;padding:10px 12px;border-radius:6px;font-size:13px"><b>President&#39;s note:</b><br>' + escapeHtml_(returnNote).replace(/\n/g, '<br>') + '</div>'
      : '';
    actions = noteBox + '<div style="margin-top:16px"><button class="btn approve" onclick="doResend()">Re-send to President</button></div>' + sendBack;
  } else {
    heading = 'Claim status';
    sub = 'This claim is now: <b>' + escapeHtml_(row.status) + '</b>.';
  }

  return '<html><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:' + cream + ';color:#1B2A20;margin:0;padding:24px}'
    + '.card{max-width:540px;margin:0 auto;background:#FBF7F1;border:1px solid #DED2C1;border-radius:14px;padding:24px}'
    + 'h1{color:' + g + ';font-size:20px;margin:0 0 4px}.muted{color:#5E6B60;font-size:13px}'
    + 'table.sum{font-size:14px;margin:14px 0;width:100%}table.sum td{padding:3px 0}table.sum td.k{color:#666;width:140px}'
    + '.btn{display:inline-block;border:none;border-radius:9px;padding:12px 20px;font-size:15px;font-weight:bold;cursor:pointer}'
    + '.approve{background:' + g + ';color:#fff}.back{background:none;border:1px solid #CBBCA6;color:#1B2A20}'
    + 'a.file{color:' + g + ';font-weight:bold}textarea{width:100%;box-sizing:border-box;min-height:70px;border:1px solid #CBBCA6;border-radius:8px;padding:10px;font:inherit;margin-top:8px}'
    + '#msg{margin-top:14px;font-weight:bold}</style></head><body>'
    + '<div class="card">'
    + '<h1>' + heading + '</h1>'
    + '<div class="muted">' + escapeHtml_(row.name) + ' &middot; Claim ' + escapeHtml_(String(row.id)) + '</div>'
    + '<p style="font-size:13.5px;color:#3a463c">' + sub + '</p>'
    + '<table class="sum">'
    + '<tr><td class="k">Payable to</td><td>' + escapeHtml_(row.name) + ' (' + escapeHtml_(row.payTo) + ')</td></tr>'
    + '<tr><td class="k">Reimbursement</td><td><b>$' + money_(row.total) + '</b></td></tr>'
    + (isMil
        ? '<tr><td class="k">Distance</td><td>' + escapeHtml_(String((payload && payload.totals && payload.totals.km) || 0)) + ' km</td></tr>'
        : '<tr><td class="k">GST / HST</td><td>$' + money_(row.gst) + ' / $' + money_(row.hst) + '</td></tr>')
    + '</table>'
    + '<p><a class="file" href="' + pdfUrl + '" target="_blank">Open the claim packet (PDF) &rarr;</a></p>'
    + (row.folderUrl ? '<p><a class="file" href="' + row.folderUrl + '" target="_blank">Open the Drive folder (all receipts) &rarr;</a></p>' : '')
    + actions
    + '<div id="msg"></div>'
    + '</div>'
    + '<script>'
    + 'var ID=' + JSON.stringify(String(row.id)) + ',TK=' + JSON.stringify(String(row.token)) + ',STAGE=' + JSON.stringify(stage) + ';'
    + 'var cv=document.getElementById("pad"),cx=cv?cv.getContext("2d"):null,drew=false,drawing=false,lx,ly;'
    + 'function pos(e){var r=cv.getBoundingClientRect();var t=e.touches?e.touches[0]:e;return {x:(t.clientX-r.left)*(cv.width/r.width),y:(t.clientY-r.top)*(cv.height/r.height)};}'
    + 'function dn(e){drawing=true;var p=pos(e);lx=p.x;ly=p.y;e.preventDefault();}'
    + 'function mv(e){if(!drawing)return;var p=pos(e);cx.beginPath();cx.moveTo(lx,ly);cx.lineTo(p.x,p.y);cx.stroke();lx=p.x;ly=p.y;drew=true;e.preventDefault();}'
    + 'function up(){drawing=false;}'
    + 'if(cx){cx.lineWidth=2.4;cx.lineCap="round";cx.lineJoin="round";cx.strokeStyle="#14263f";'
    + 'cv.addEventListener("mousedown",dn);cv.addEventListener("mousemove",mv);window.addEventListener("mouseup",up);'
    + 'cv.addEventListener("touchstart",dn,{passive:false});cv.addEventListener("touchmove",mv,{passive:false});cv.addEventListener("touchend",up);}'
    + 'function clearPad(){if(cx){cx.clearRect(0,0,cv.width,cv.height);drew=false;}}'
    + 'function msg(t){document.getElementById("msg").textContent=t;}'
    + 'var busy=false;'
    + 'function setBtns(d){var b=document.querySelectorAll("button");for(var i=0;i<b.length;i++){b[i].disabled=d;}}'
    + 'function lock(){setBtns(true);}'
    // begin() is the single gate: once an action starts, busy=true and every button is
    // disabled, so a second click (or double-click) is ignored until we hear back.
    + 'function begin(t){if(busy){return false;}busy=true;lock();msg(t);return true;}'
    + 'function ok(t){busy=true;lock();msg(t);}'
    + 'function fail(e){busy=false;setBtns(false);msg("Error: "+((e&&e.message)||e));}'
    + 'function doSign(){if(busy){return;}if(!drew){alert("Please draw your signature first.");return;}'
    + 'var gls=[],allgl=true,sels=document.querySelectorAll("select.gl");for(var i=0;i<sels.length;i++){var idx=+sels[i].getAttribute("data-i");gls[idx]=sels[i].value;if(!sels[i].value)allgl=false;}'
    + 'if(STAGE==="vp1"&&!allgl){alert("Please assign a GL code to every expense line.");return;}'
    + 'var sig=cv.toDataURL("image/png");if(!begin("Submitting your signature...")){return;}'
    + 'var done={president:"Signed. Sent to NU Staff for signature.",staff:"Signed. Sent back to VP Finance for final approval."}[STAGE]||"Signed and approved. Sent to the President for signature.";'
    + 'var run=google.script.run.withSuccessHandler(function(){ok(done);}).withFailureHandler(fail);'
    + 'if(STAGE==="president"){run.presidentSign(ID,TK,sig);}else if(STAGE==="staff"){run.staffSign(ID,TK,sig);}else{run.vpSignApprove(ID,TK,sig,gls);}}'
    + 'function doFinal(){if(busy){return;}if(!confirm("Give final approval? The final signed PDF will be emailed to you to submit to Accounts Payable.")){return;}if(!begin("Approving...")){return;}'
    + 'google.script.run.withSuccessHandler(function(r){ok("Approved. The final signed PDF was emailed to you - send it to Accounts Payable. You can also open it with the PDF link above."+(r&&r.requesterEmailFailed?" Note: the copy to the requester could not be sent - please forward it to them.":""));}).withFailureHandler(fail).finalApprove(ID,TK);}'
    + 'function doSendBack(){if(busy){return;}var n=document.getElementById("note").value;if(!n.trim()){alert("Add a note first.");return;}if(!begin("Emailing the student...")){return;}'
    + 'google.script.run.withSuccessHandler(function(){ok("Sent back to the student.");}).withFailureHandler(fail).sendBackClaim(ID,TK,n);}'
    + 'function doReturn(){if(busy){return;}var n=document.getElementById("note").value;if(!n.trim()){alert("Add a note first.");return;}if(!begin("Returning to VP Finance...")){return;}'
    + 'var rr=google.script.run.withSuccessHandler(function(){ok("Returned to VP Finance.");}).withFailureHandler(fail);'
    + 'if(STAGE==="staff"){rr.staffReturn(ID,TK,n);}else{rr.presidentReturn(ID,TK,n);}}'
    + 'function doResend(){if(busy){return;}if(!confirm("Re-send this claim to the President for signature?")){return;}if(!begin("Re-sending to the President...")){return;}'
    + 'google.script.run.withSuccessHandler(function(){ok("Re-sent to the President for signature.");}).withFailureHandler(fail).resendToPresident(ID,TK);}'
    + 'function doResendStaff(){if(busy){return;}if(!confirm("Email you the PDF and the message to send to NU Staff from your NU email?")){return;}if(!begin("Preparing the email...")){return;}'
    + 'google.script.run.withSuccessHandler(function(){ok("Done. Check your email for the PDF and the message to send to NU Staff.");}).withFailureHandler(fail).resendToStaff(ID,TK);}'
    + '</scr' + 'ipt></body></html>';
}


/* ---------- tiny utils ---------- */
// Only a genuine base64 image data URL may go into an <img src="..."> on the PDF. Anything else
// (e.g. a tampered submission carrying HTML) is dropped, so it can't add content such as a fake
// "Approved by" line to the signed form.
function isImageDataUrl_(s) { return /^data:image\/[a-z0-9.+-]+;base64,[A-Za-z0-9+\/=\s]+$/i.test(String(s || '')); }
function safeImageSrc_(s) { return isImageDataUrl_(s) ? s : ''; }
// When the claim was submitted (claims from before V4 get it from the log's Timestamp on their next re-signing).
function submittedDate_(p) {
  const d = p && p.submittedAt ? new Date(p.submittedAt) : null;
  return d && !isNaN(d.getTime()) ? d : new Date();
}
function requireSignature_(sig) {
  if (!sig) throw new Error('Please draw your signature before approving.');
  if (!isImageDataUrl_(sig)) throw new Error('The signature could not be read. Please clear it and draw it again.');
}
function money_(n) {
  n = parseFloat(n) || 0;
  return n.toLocaleString('en-CA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function formLabel_(x) { return (x && x.form === 'mileage') ? 'Kilometer Reimbursement' : 'Expense Claim'; }
function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

/* NU logo (extracted from the official template, embedded for the PDF header) */
const NU_LOGO_DATAURI = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAUAAAACcCAIAAACBeeUdAAABXklEQVR42u3TQQkAAAwDsfp3WDdVsccgkXBwKfBWJAADAwYGDAwGBgwMGBgwMBgYMDBgYDAwYGDAwICBwcCAgQEDAwYGAwMGBgwMBgYMDBgYMDAYGDAwYGDAwGBgwMCAgcHAgIEBAwMGBgMDBgYMDAYGDAwYGDAwGBgwMGBgwMBgYMDAgIHBwICBAQMDBgYDAwYGDAwYGAwMGBgwMBgYMDBgYMDAYGDAwICBAQODgQEDAwYGAwMGBgwMGBgMDBgYMDAYGDAwYGDAwGBgwMCAgQEDg4EBAwMGBgMDBgYMDBgYDAwYGDAwYGAwMGBgwMBgYMDAgIEBA4OBAQMDBgYMDAYGDAwYGAwMGBgwMGBgMDBgYMDAYGDAwICBAQODgQEDAwYGDAwGBgwMGBgMDBgYMDBgYDAwYGDAwICBwcCAgQEDg4EBAwMGBgwMBgYMDBgYDAwYGDAwYGAwMGBg4MoA2VL+AoBex0EAAAAASUVORK5CYII=";
