/**
 * NUSA — Email me the claims waiting for my signature  (script bound to "NUSA Expense Claims — Log")
 * ------------------------------------------------------------------
 * Menu NUSA > "Email me claims waiting for my signature" sends ONE email to VP Finance listing
 * every claim on the Claims tab that still needs VP Finance: first the ones waiting for your
 * signature (and GL codes) or your final approval, then the ones waiting on you for another reason
 * (returned with a note, or a signed PDF to send to NU Staff). Each claim has its "Open & sign" link
 * (the same link the claim system emails you), its PDF and its Drive folder; the PDFs are attached
 * too (up to ~20 MB in total; the rest are linked).
 *
 * It only reads the Claims tab: nothing is signed, changed or sent to anyone else.
 * The email goes to UC_TO below (the address that receives the sign links); change it if you like.
 *
 * SETUP (once): in the claim log, Extensions > Apps Script. Add a new script file
 * (+ > Script editor) named "Unsigned Claims", paste this file, Save.
 * Also paste the updated ManualExpense.gs into Code.gs (it adds the menu item).
 * Reload the spreadsheet; the first time, Google asks you to authorize sending email and reading Drive.
 */

const UC_TO = 'vp.finance@example.org';
const UC_WEB_APP_URL = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
const UC_CLAIMS_TAB = 'Claims';
const UC_ATTACH_CAP = 20000000; // ~20 MB of PDFs per email (Gmail's limit is 25 MB)
// Claims tab columns, same order as LOG_HEADERS in the claim web app's Code.gs.
const UC_COL = { id: 0, time: 1, status: 2, payTo: 3, name: 4, total: 6, folder: 9, pdf: 10, token: 11, decidedAt: 13, form: 15 };
// Claim statuses that wait on VP Finance: [status starts with, what it needs, counts as "waiting for your signature"].
const UC_STAGES = [
  ['Pending VP Finance signature', 'Your signature and GL codes', true],
  ['Pending VP Finance final', 'Your final approval', true],
  ['Returned by President', 'Returned by the President with a note', false],
  ['Returned by NU Staff', 'Returned by NU Staff with a note', false],
  ['Pending NU Staff signed PDF', 'Send the PDF to NU Staff to sign', false],
];

// Menu item.
function emailMeUnsignedClaims() {
  const ui = SpreadsheetApp.getUi();
  const claims = ucWaitingClaims_();
  if (!claims.length) {
    ui.alert('Nothing waiting on you', 'No claims are waiting for your signature right now.', ui.ButtonSet.OK);
    return;
  }
  const sign = claims.filter(function (c) { return c.sign; }).length;
  const other = claims.length - sign;
  const what = ucCount_(sign, 'claim') + ' waiting for your signature' + (other ? ' and ' + other + ' other waiting on you' : '');
  if (ui.alert('Email me unsigned claims', 'Send the list (' + what + ') to ' + UC_TO + '?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const res = ucSendList_(claims);
  ui.alert('Sent', 'Emailed ' + what + ' to ' + UC_TO + '.'
    + (res.attached ? ' ' + ucCount_(res.attached, 'PDF') + ' attached.' : '')
    + (res.linkedOnly ? ' ' + ucCount_(res.linkedOnly, 'PDF') + ' only linked (too large to attach, or not found).' : ''), ui.ButtonSet.OK);
}

// Every claim on the Claims tab that is waiting on VP Finance; signature ones first, oldest first.
function ucWaitingClaims_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(UC_CLAIMS_TAB);
  if (!sh) throw new Error('No "' + UC_CLAIMS_TAB + '" tab in this spreadsheet.');
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const id = String(row[UC_COL.id] === undefined ? '' : row[UC_COL.id]).trim();
    if (!id) continue;
    const status = String(row[UC_COL.status] || '').trim();
    const st = UC_STAGES.filter(function (s) { return status.indexOf(s[0]) === 0; })[0];
    if (!st) continue;
    out.push({
      id: id, status: status, need: st[1], sign: st[2],
      name: String(row[UC_COL.name] || '').trim(), payTo: String(row[UC_COL.payTo] || '').trim(),
      total: Number(row[UC_COL.total]) || 0, mileage: String(row[UC_COL.form] || '') === 'mileage',
      submitted: row[UC_COL.time], days: ucDaysSince_(row[UC_COL.decidedAt] || row[UC_COL.time]),
      folder: String(row[UC_COL.folder] || '').trim(), pdfId: String(row[UC_COL.pdf] || '').trim(),
      token: String(row[UC_COL.token] || '').trim(),
    });
  }
  out.sort(function (a, b) { return (b.sign - a.sign) || (b.days - a.days); });
  return out;
}

function ucSendList_(claims) {
  const attachments = [];
  let used = 0, linkedOnly = 0;
  claims.forEach(function (c) {
    c.attached = false;
    if (!c.pdfId) { linkedOnly++; return; }
    try {
      const blob = DriveApp.getFileById(c.pdfId).getBlob();
      const n = blob.getBytes().length;
      if (used + n <= UC_ATTACH_CAP) { attachments.push(blob); used += n; c.attached = true; } else linkedOnly++;
    } catch (e) {
      console.warn('PDF not attached for claim ' + c.id + ': ' + e);
      linkedOnly++;
    }
  });
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'MMM d, yyyy');
  const sign = claims.filter(function (c) { return c.sign; });
  const other = claims.filter(function (c) { return !c.sign; });
  const total = sign.reduce(function (s, c) { return s + c.total; }, 0);
  MailApp.sendEmail({
    to: UC_TO,
    subject: 'Claims waiting for your signature: ' + sign.length + (other.length ? ' (+' + other.length + ' other)' : '') + ' - ' + today,
    htmlBody: ucEmailHtml_(sign, other, total, today, tz),
    body: ucEmailText_(claims, tz),
    attachments: attachments,
    name: 'NUSA Finance',
  });
  return { sent: claims.length, attached: attachments.length, linkedOnly: linkedOnly };
}

function ucEmailHtml_(sign, other, total, today, tz) {
  const green = '#0D4624';
  const table = function (list) {
    return '<table style="border-collapse:collapse;width:100%;font-size:13px">'
      + '<tr style="background:#F1E8DE;text-align:left">'
      + ['Claim', 'Amount', 'Submitted', 'Waiting', 'Needs', ''].map(function (h) { return '<th style="padding:6px 8px">' + h + '</th>'; }).join('')
      + '</tr>'
      + list.map(function (c) {
        const links = ['<a href="' + ucEsc_(ucSignUrl_(c)) + '" style="color:' + green + ';font-weight:bold">Open &amp; sign</a>'];
        if (c.pdfId) links.push('<a href="https://drive.google.com/file/d/' + encodeURIComponent(c.pdfId) + '/view">PDF</a>' + (c.attached ? ' (attached)' : ''));
        if (c.folder) links.push('<a href="' + ucEsc_(c.folder) + '">Folder</a>');
        return '<tr style="border-bottom:1px solid #ddd;vertical-align:top">'
          + '<td style="padding:6px 8px"><b>' + ucEsc_(c.name || c.payTo || '(no name)') + '</b><br><span style="color:#666">'
          + (c.mileage ? 'Kilometer reimbursement' : 'Expense claim') + ' &middot; ' + ucEsc_(c.id) + '</span></td>'
          + '<td style="padding:6px 8px;white-space:nowrap">$' + ucMoney_(c.total) + '</td>'
          + '<td style="padding:6px 8px;white-space:nowrap">' + ucDate_(c.submitted, tz) + '</td>'
          + '<td style="padding:6px 8px;white-space:nowrap">' + (c.days === 1 ? '1 day' : c.days + ' days') + '</td>'
          + '<td style="padding:6px 8px">' + ucEsc_(c.need) + '</td>'
          + '<td style="padding:6px 8px;white-space:nowrap">' + links.join('<br>') + '</td></tr>';
      }).join('')
      + '</table>';
  };
  return '<div style="font-family:sans-serif;max-width:760px">'
    + '<h2 style="color:' + green + '">Claims waiting for your signature</h2>'
    + '<p>As of ' + ucEsc_(today) + ': <b>' + ucCount_(sign.length, 'claim') + '</b>'
    + (sign.length ? ' totalling <b>$' + ucMoney_(total) + '</b>' : '') + '.</p>'
    + (sign.length ? table(sign) : '<p>None right now.</p>')
    + (other.length ? '<h3 style="color:' + green + ';margin-top:24px">Also waiting on you</h3>' + table(other) : '')
    + '<p style="font-size:12px;color:#666;margin-top:20px">Sent from the claim log (NUSA &gt; Email me claims waiting for my signature). '
    + 'The "Open &amp; sign" links are your private approval links: don\'t forward this email.</p></div>';
}

function ucEmailText_(claims, tz) {
  return 'Claims waiting on you:\n\n' + claims.map(function (c) {
    return '- ' + (c.name || c.payTo) + ', $' + ucMoney_(c.total) + ', submitted ' + ucDate_(c.submitted, tz) + ' - ' + c.need + '\n  ' + ucSignUrl_(c);
  }).join('\n');
}

function ucSignUrl_(c) {
  return UC_WEB_APP_URL + '?action=decision&id=' + encodeURIComponent(c.id) + '&token=' + encodeURIComponent(c.token);
}
function ucDaysSince_(d) {
  const t = d ? new Date(d).getTime() : NaN;
  return isNaN(t) ? 0 : Math.max(0, Math.floor((Date.now() - t) / 86400000));
}
function ucDate_(d, tz) {
  const x = d instanceof Date ? d : new Date(d);
  return isNaN(x.getTime()) ? '' : Utilities.formatDate(x, tz, 'MMM d');
}
function ucMoney_(n) { return (Math.round((Number(n) || 0) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function ucCount_(n, word) { return n + ' ' + word + (n === 1 ? '' : 's'); }
function ucEsc_(s) {
  return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
