/**
 * NUSA — Remind the President  (script bound to "NUSA Expense Claims — Log")
 * ------------------------------------------------------------------
 * Menu NUSA > "Remind the President to sign claims" sends ONE email to the President listing every claim
 * on the Claims tab at "Pending President signature" (expense claims and kilometer reimbursements): who,
 * how much, how long it has waited, and an "Open & sign" button for each (the same link the claim system
 * emailed the President). VP Finance is in copy and replies go to VP Finance.
 *
 * Before sending it shows how many claims, when the President was last reminded, and lets you type an
 * optional note to include. It only reads the Claims tab: nothing is signed or changed.
 *
 * SETUP (once): in the claim log, Extensions > Apps Script. Add a new script file
 * (+ > Script editor) named "President Reminders", paste this file, Save.
 * Also paste the updated ManualExpense.gs into Code.gs (it adds the menu item).
 * Reload the spreadsheet; the first time, Google asks you to authorize sending email.
 */

const PR_PRESIDENT_EMAIL = 'president@example.org'; // same as PRESIDENT_EMAIL in the claim web app's Code.gs
const PR_VP_EMAIL = 'vp.finance@example.org';        // in copy, and where replies go
const PR_WEB_APP_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
const PR_CLAIMS_TAB = 'Claims';
// Claims tab columns, same order as LOG_HEADERS in the claim web app's Code.gs.
const PR_COL = { id: 0, time: 1, status: 2, payTo: 3, name: 4, total: 6, token: 11, decidedAt: 13, form: 15 };
const PR_PROP = 'PR_LAST_REMINDER'; // when the last reminder went out (document properties)

// Menu item.
function remindPresidentClaims() {
  const ui = SpreadsheetApp.getUi();
  const claims = prWaitingClaims_();
  if (!claims.length) {
    ui.alert('Nothing to remind', 'No claims are waiting for the President\'s signature right now.', ui.ButtonSet.OK);
    return;
  }
  const tz = Session.getScriptTimeZone();
  const last = Number(PropertiesService.getDocumentProperties().getProperty(PR_PROP) || 0);
  const res = ui.prompt('Remind the President',
    prCount_(claims.length, 'claim') + ' waiting for the President\'s signature ($' + prMoney_(prTotal_(claims)) + ').\n'
    + 'Last reminder from this menu: ' + (last ? Utilities.formatDate(new Date(last), tz, 'MMM d, yyyy') : 'never') + '.\n\n'
    + 'Optional: type a short note to include, then press OK to send to ' + PR_PRESIDENT_EMAIL + ' (you in copy).',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  prSendReminder_(claims, res.getResponseText());
  ui.alert('Sent', 'Reminder sent to the President with ' + prCount_(claims.length, 'claim') + '.', ui.ButtonSet.OK);
}

// Every claim at "Pending President signature", longest waiting first.
function prWaitingClaims_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(PR_CLAIMS_TAB);
  if (!sh) throw new Error('No "' + PR_CLAIMS_TAB + '" tab in this spreadsheet.');
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const id = String(row[PR_COL.id] === undefined ? '' : row[PR_COL.id]).trim();
    if (!id) continue;
    if (String(row[PR_COL.status] || '').trim().indexOf('Pending President') !== 0) continue;
    out.push({
      id: id, name: String(row[PR_COL.name] || '').trim(), payTo: String(row[PR_COL.payTo] || '').trim(),
      total: Number(row[PR_COL.total]) || 0, mileage: String(row[PR_COL.form] || '') === 'mileage',
      submitted: row[PR_COL.time], days: prDaysSince_(row[PR_COL.decidedAt] || row[PR_COL.time]),
      token: String(row[PR_COL.token] || '').trim(),
    });
  }
  out.sort(function (a, b) { return b.days - a.days; });
  return out;
}

function prSendReminder_(claims, note) {
  note = String(note || '').trim().slice(0, 2000);
  const tz = Session.getScriptTimeZone();
  MailApp.sendEmail({
    to: PR_PRESIDENT_EMAIL,
    cc: PR_VP_EMAIL,
    replyTo: PR_VP_EMAIL,
    subject: 'Reminder: ' + prCount_(claims.length, 'NUSA claim') + ' waiting for your signature',
    htmlBody: prEmailHtml_(claims, note, tz),
    body: prEmailText_(claims, note, tz),
    name: 'NUSA Finance',
  });
  PropertiesService.getDocumentProperties().setProperty(PR_PROP, String(Date.now()));
}

function prEmailHtml_(claims, note, tz) {
  const green = '#0D4624';
  return '<div style="font-family:sans-serif;max-width:680px">'
    + '<h2 style="color:' + green + '">' + prCount_(claims.length, 'claim') + ' waiting for your signature</h2>'
    + (note ? '<p style="white-space:pre-wrap;border-left:3px solid ' + green + ';padding-left:10px">' + prEsc_(note) + '</p>' : '')
    + '<p>These claims have been signed by VP Finance and need your signature before they go to NU Staff. '
    + 'Total: <b>$' + prMoney_(prTotal_(claims)) + '</b>.</p>'
    + '<table style="border-collapse:collapse;width:100%;font-size:13px">'
    + '<tr style="background:#F1E8DE;text-align:left">'
    + ['Claim', 'Amount', 'Submitted', 'Waiting', ''].map(function (h) { return '<th style="padding:6px 8px">' + h + '</th>'; }).join('')
    + '</tr>'
    + claims.map(function (c) {
      return '<tr style="border-bottom:1px solid #ddd;vertical-align:top">'
        + '<td style="padding:6px 8px"><b>' + prEsc_(c.name || c.payTo || '(no name)') + '</b><br><span style="color:#666">'
        + (c.mileage ? 'Kilometer reimbursement' : 'Expense claim') + '</span></td>'
        + '<td style="padding:6px 8px;white-space:nowrap">$' + prMoney_(c.total) + '</td>'
        + '<td style="padding:6px 8px;white-space:nowrap">' + prDate_(c.submitted, tz) + '</td>'
        + '<td style="padding:6px 8px;white-space:nowrap">' + (c.days === 1 ? '1 day' : c.days + ' days') + '</td>'
        + '<td style="padding:6px 8px;white-space:nowrap"><a href="' + prEsc_(prSignUrl_(c)) + '" style="background:' + green
        + ';color:#fff;text-decoration:none;padding:6px 12px;border-radius:6px;font-weight:bold">Open &amp; sign</a></td></tr>';
    }).join('')
    + '</table>'
    + '<p style="font-size:12px;color:#666;margin-top:20px">Sent by NUSA VP Finance. Questions? Just reply to this email.</p></div>';
}

function prEmailText_(claims, note, tz) {
  return (note ? note + '\n\n' : '') + 'Claims waiting for your signature:\n\n' + claims.map(function (c) {
    return '- ' + (c.name || c.payTo) + ', $' + prMoney_(c.total) + ', submitted ' + prDate_(c.submitted, tz) + '\n  ' + prSignUrl_(c);
  }).join('\n');
}

function prSignUrl_(c) {
  return PR_WEB_APP_URL + '?action=decision&id=' + encodeURIComponent(c.id) + '&token=' + encodeURIComponent(c.token);
}
function prTotal_(claims) { return claims.reduce(function (s, c) { return s + c.total; }, 0); }
function prDaysSince_(d) {
  const t = d ? new Date(d).getTime() : NaN;
  return isNaN(t) ? 0 : Math.max(0, Math.floor((Date.now() - t) / 86400000));
}
function prDate_(d, tz) {
  const x = d instanceof Date ? d : new Date(d);
  return isNaN(x.getTime()) ? '' : Utilities.formatDate(x, tz, 'MMM d');
}
function prMoney_(n) { return (Math.round((Number(n) || 0) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function prCount_(n, word) { return n + ' ' + word + (n === 1 ? '' : 's'); }
function prEsc_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
