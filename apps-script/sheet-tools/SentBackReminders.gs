/**
 * NUSA — Remind people whose claim was sent back  (script bound to "NUSA Expense Claims — Log")
 * ------------------------------------------------------------------
 * Menu NUSA > "Remind people with sent-back claims" opens a list of every claim on the Claims tab that is
 * at "Sent back for changes" (expense claims and kilometer reimbursements): who, how much, when it was sent
 * back, your note, and when they were last reminded. Tick the people to remind, optionally add a message,
 * and press Send. Each person gets an email (VP Finance in copy) with your original note and their
 * "Edit your claim" button: the same private link as the first email, which opens their claim already filled
 * in so they can fix it and press Resubmit claim. Nothing else changes: the claim stays "Sent back for
 * changes" until they resubmit, and the link in the first email keeps working too.
 *
 * Untick anyone who already filed a new claim instead.
 *
 * SETUP (once): in the claim log, Extensions > Apps Script. Add a new script file
 * (+ > Script editor) named "Sent Back Reminders", paste this file, Save.
 * Also paste the updated ManualExpense.gs into Code.gs (it adds the menu item).
 * Reload the spreadsheet; the first time, Google asks you to authorize sending email and reading Drive.
 */

const SBR_WEB_APP_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
const SBR_VP_EMAIL = 'vp.finance@example.org'; // in copy of every reminder, and where replies go
const SBR_CLAIMS_TAB = 'Claims';
// Claims tab columns, same order as LOG_HEADERS in the claim web app's Code.gs.
const SBR_COL = { id: 0, time: 1, status: 2, payTo: 3, name: 4, email: 5, total: 6, folder: 9, token: 11, decidedAt: 13, json: 14, form: 15 };
const SBR_PROP = 'SBR_LAST_'; // + claim ID: when the last reminder went out (document properties)

// Menu item.
function remindSentBackClaims() {
  const ui = SpreadsheetApp.getUi();
  const claims = sbrSentBackClaims_(true);
  if (!claims.length) {
    ui.alert('Nothing to remind', 'No claims are at "Sent back for changes" right now.', ui.ButtonSet.OK);
    return;
  }
  const html = HtmlService.createHtmlOutput(sbrDialogHtml_(claims)).setWidth(760).setHeight(560);
  ui.showModalDialog(html, 'Remind people with sent-back claims');
}

// Called by the dialog: emails the ticked claims. Re-reads the Claims tab, so a claim resubmitted in the
// meantime is skipped.
function sbrSendReminders(ids, message) {
  if (!Array.isArray(ids) || !ids.length) throw new Error('Tick at least one person.');
  const want = {};
  ids.forEach(function (id) { want[String(id)] = true; });
  message = String(message || '').trim().slice(0, 2000);
  const props = PropertiesService.getDocumentProperties();
  const sent = [], skipped = [];
  const claims = sbrSentBackClaims_(true);
  const byId = {};
  claims.forEach(function (c) { byId[c.id] = c; });
  Object.keys(want).forEach(function (id) {
    const c = byId[id];
    if (!c) { skipped.push(id + ' (not "Sent back for changes" any more)'); return; }
    if (!c.canEdit) { skipped.push((c.name || id) + ' (' + c.why + ')'); return; }
    MailApp.sendEmail({
      to: c.email,
      cc: SBR_VP_EMAIL,
      replyTo: SBR_VP_EMAIL,
      subject: 'Reminder: your NUSA ' + (c.mileage ? 'kilometer reimbursement' : 'expense claim') + ' still needs a change',
      htmlBody: sbrEmailHtml_(c, message),
      body: sbrEmailText_(c, message),
      name: 'NUSA Finance',
    });
    props.setProperty(SBR_PROP + c.id, new Date().toISOString());
    sent.push(c.name || c.email);
  });
  return { sent: sent, skipped: skipped };
}

// Every claim at "Sent back for changes", sent back longest ago first. withNotes: also read VP Finance's
// note from the claim's saved data (Drive).
function sbrSentBackClaims_(withNotes) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBR_CLAIMS_TAB);
  if (!sh) throw new Error('No "' + SBR_CLAIMS_TAB + '" tab in this spreadsheet.');
  const data = sh.getDataRange().getValues();
  const props = PropertiesService.getDocumentProperties();
  const out = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const id = String(row[SBR_COL.id] === undefined ? '' : row[SBR_COL.id]).trim();
    if (!id || String(row[SBR_COL.status] || '').trim().indexOf('Sent back') !== 0) continue;
    const c = {
      id: id, name: String(row[SBR_COL.name] || row[SBR_COL.payTo] || '').trim(),
      email: String(row[SBR_COL.email] || '').trim(), total: Number(row[SBR_COL.total]) || 0,
      mileage: String(row[SBR_COL.form] || '') === 'mileage', submitted: row[SBR_COL.time],
      sentBack: row[SBR_COL.decidedAt] || row[SBR_COL.time], token: String(row[SBR_COL.token] || '').trim(),
      jsonId: String(row[SBR_COL.json] || '').trim(), note: '', last: props.getProperty(SBR_PROP + id) || '',
    };
    c.days = sbrDaysSince_(c.sentBack);
    c.why = !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c.email) ? 'no email address on the claim'
      : !c.token ? 'no link token on the claim'
      : !c.jsonId ? 'older claim that can\'t be edited online: ask them to file a new one' : '';
    c.canEdit = !c.why;
    if (withNotes && c.jsonId) c.note = sbrNote_(c.jsonId);
    out.push(c);
  }
  out.sort(function (a, b) { return b.days - a.days; });
  return out;
}

function sbrNote_(jsonId) {
  try {
    const p = JSON.parse(DriveApp.getFileById(jsonId).getBlob().getDataAsString());
    return String(p.sentBackNote || '').trim();
  } catch (e) {
    console.warn('Note not read for ' + jsonId + ': ' + e);
    return '';
  }
}

function sbrEditUrl_(c) {
  return SBR_WEB_APP_URL + '?action=edit&id=' + encodeURIComponent(c.id) + '&token=' + encodeURIComponent(c.token);
}

function sbrEmailHtml_(c, message) {
  const green = '#0D4624';
  const tz = Session.getScriptTimeZone();
  const quote = function (t) {
    return '<blockquote style="border-left:3px solid ' + green + ';padding-left:12px;color:#333;margin:12px 0">' + sbrEsc_(t).replace(/\n/g, '<br>') + '</blockquote>';
  };
  const lbl = c.mileage ? 'kilometer reimbursement' : 'expense claim';
  return '<div style="font-family:sans-serif;max-width:560px">'
    + '<p>Hi ' + sbrEsc_(c.name.split(' ')[0] || 'there') + ',</p>'
    + '<p>This is a reminder that your ' + lbl + ' of <b>$' + sbrMoney_(c.total) + '</b>'
    + (sbrDate_(c.submitted, tz) ? ' (submitted ' + sbrDate_(c.submitted, tz) + ')' : '')
    + ' was sent back' + (sbrDate_(c.sentBack, tz) ? ' on ' + sbrDate_(c.sentBack, tz) : '')
    + ' and is still waiting for your changes. We can\'t reimburse you until it is fixed and resubmitted.</p>'
    + (message ? quote(message) : '')
    + (c.note ? '<p>What needs fixing:</p>' + quote(c.note) : '<p>What needs fixing is in our earlier email.</p>')
    + '<p>Open your claim with the button below. Everything you entered (and your receipts) is already filled in: change what\'s needed and press <b>Resubmit claim</b>. No need to start a new claim.</p>'
    + '<p style="margin:22px 0"><a href="' + sbrEsc_(sbrEditUrl_(c)) + '" style="background:' + green + ';color:#fff;text-decoration:none;padding:12px 22px;border-radius:8px;font-weight:bold">Edit your claim &rarr;</a></p>'
    + '<p>If you no longer need this claim, just reply to this email and let us know.</p>'
    + '<p style="font-size:12px;color:#666">This link is only for you and only for this claim. It stops working once you resubmit.</p>'
    + '<p>- NUSA Finance</p></div>';
}

function sbrEmailText_(c, message) {
  return 'Hi ' + (c.name.split(' ')[0] || 'there') + ',\n\n'
    + 'This is a reminder that your ' + (c.mileage ? 'kilometer reimbursement' : 'expense claim') + ' of $' + sbrMoney_(c.total)
    + ' was sent back and is still waiting for your changes.\n\n'
    + (message ? message + '\n\n' : '')
    + (c.note ? 'What needs fixing:\n' + c.note + '\n\n' : 'What needs fixing is in our earlier email.\n\n')
    + 'Edit your claim (everything is already filled in; change what is needed and press Resubmit claim):\n' + sbrEditUrl_(c) + '\n\n'
    + 'If you no longer need this claim, just reply and let us know.\n\n- NUSA Finance';
}

function sbrDialogHtml_(claims) {
  const tz = Session.getScriptTimeZone();
  const rows = claims.map(function (c) {
    const note = c.note.length > 140 ? c.note.slice(0, 140) + '…' : c.note;
    return '<tr' + (c.canEdit ? '' : ' class="off"') + '>'
      + '<td><input type="checkbox" class="pick" value="' + sbrEsc_(c.id) + '"' + (c.canEdit ? ' checked' : ' disabled') + '></td>'
      + '<td><b>' + sbrEsc_(c.name || '(no name)') + '</b><br><span class="sub">' + (c.mileage ? 'Kilometer reimbursement' : 'Expense claim')
      + ' &middot; ' + sbrEsc_(c.email || 'no email') + '</span>' + (c.why ? '<br><span class="warn">Can\'t remind: ' + sbrEsc_(c.why) + '</span>' : '') + '</td>'
      + '<td class="n">$' + sbrMoney_(c.total) + '</td>'
      + '<td>' + sbrEsc_(sbrDate_(c.sentBack, tz)) + '<br><span class="sub">' + (c.days === 1 ? '1 day ago' : c.days + ' days ago') + '</span></td>'
      + '<td class="note" title="' + sbrEsc_(c.note) + '">' + (note ? sbrEsc_(note) : '<span class="sub">(in the first email)</span>') + '</td>'
      + '<td>' + (c.last ? sbrEsc_(sbrDate_(c.last, tz)) : '<span class="sub">Never</span>') + '</td></tr>';
  }).join('');
  return '<!doctype html><html><head><base target="_top"><style>'
    + 'body{font-family:Arial,sans-serif;font-size:13px;color:#222;margin:0;padding:4px 2px}'
    + 'p{margin:0 0 10px}table{border-collapse:collapse;width:100%}th{background:#F1E8DE;text-align:left;padding:6px;font-size:12px}'
    + 'td{padding:6px;border-bottom:1px solid #e3e3e3;vertical-align:top}td.n{white-space:nowrap}td.note{max-width:220px;color:#333}'
    + '.sub{color:#777;font-size:11.5px}.warn{color:#a42727;font-size:11.5px}tr.off td{color:#999}'
    + 'textarea{width:100%;box-sizing:border-box;height:58px;font:inherit;padding:6px;margin-top:4px}'
    + '.bar{display:flex;gap:8px;align-items:center;margin-top:12px}button{font:inherit;padding:7px 14px;border-radius:6px;border:1px solid #bbb;background:#fff;cursor:pointer}'
    + 'button.go{background:#0D4624;color:#fff;border-color:#0D4624;font-weight:bold}button:disabled{opacity:.5;cursor:default}'
    + '#msg{margin-left:auto;font-weight:bold}#msg.err{color:#a42727}a{color:#0D4624;cursor:pointer}'
    + '</style></head><body>'
    + '<p>Each person you tick gets a reminder with your note and their <b>Edit your claim</b> link (you are in copy). '
    + 'Untick anyone who already filed a new claim. <a onclick="all_(true)">Tick all</a> &middot; <a onclick="all_(false)">Untick all</a></p>'
    + '<table><tr><th></th><th>Person</th><th>Amount</th><th>Sent back</th><th>What needs fixing</th><th>Last reminder</th></tr>' + rows + '</table>'
    + '<p style="margin-top:12px">Add a message to the reminder (optional):<textarea id="note" maxlength="2000" placeholder="e.g. Please resubmit by Friday so it goes in this week\'s batch to Accounts Payable."></textarea></p>'
    + '<div class="bar"><button class="go" id="go" onclick="send_()">Send reminders</button><button onclick="google.script.host.close()">Close</button><span id="msg"></span></div>'
    + '<script>'
    + 'function picks_(){return Array.prototype.slice.call(document.querySelectorAll(".pick:checked")).map(function(x){return x.value;});}'
    + 'function all_(on){Array.prototype.forEach.call(document.querySelectorAll(".pick:not(:disabled)"),function(x){x.checked=on;});count_();}'
    + 'function count_(){var n=picks_().length;var b=document.getElementById("go");b.disabled=!n;b.textContent=n?"Send "+n+" reminder"+(n===1?"":"s"):"Send reminders";}'
    + 'Array.prototype.forEach.call(document.querySelectorAll(".pick"),function(x){x.onchange=count_;});count_();'
    + 'function say_(t,err){var m=document.getElementById("msg");m.textContent=t;m.className=err?"err":"";}'
    + 'function send_(){var ids=picks_();if(!ids.length)return;var b=document.getElementById("go");b.disabled=true;say_("Sending...");'
    + 'google.script.run.withSuccessHandler(function(r){say_("Sent "+r.sent.length+" reminder"+(r.sent.length===1?"":"s")+"."+(r.skipped.length?" Skipped: "+r.skipped.join(", "):""),r.skipped.length>0);b.textContent="Sent";})'
    + '.withFailureHandler(function(e){say_(e.message||String(e),true);count_();}).sbrSendReminders(ids,document.getElementById("note").value);}'
    + '</script></body></html>';
}

function sbrDaysSince_(d) {
  const t = d ? new Date(d).getTime() : NaN;
  return isNaN(t) ? 0 : Math.max(0, Math.floor((Date.now() - t) / 86400000));
}
function sbrDate_(d, tz) {
  if (!d) return '';
  const x = d instanceof Date ? d : new Date(d);
  return isNaN(x.getTime()) ? '' : Utilities.formatDate(x, tz, 'MMM d');
}
function sbrMoney_(n) { return (Math.round((Number(n) || 0) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function sbrEsc_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
