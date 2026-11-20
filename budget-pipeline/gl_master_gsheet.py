# Writes "GlCodes.gs" (script bound to the claim log) from gl_rows.json (made by gl_build.py):
# every revenue + expense GL code of the 10 departments and the clubs, for the "GL List" tab of the Google Sheet.
# NOTE: gl_rows.json (the parsed master GL list) comes from a separate build step that is not part of this copy.
import os, json, datetime
S = os.environ.get("GL_WORK") or os.path.dirname(os.path.abspath(__file__))
OUT = os.environ.get("GLCODES_OUT", "GlCodes.gs")
rows = json.load(open(os.path.join(S, "gl_rows.json")))
GROUP = {"Clubs (Fund 413)": "Clubs (Fund 413)"}
KEEP = ["Administration", "Faculty", "Communications", "Events", "The Courier", "Yearbook", "Student Lounge (The Commons)",
        "Special Initiatives", "Proposal Committee", "Capital Projects", "Clubs (Fund 413)"]
out = []
for r in rows:
    if r["type"] not in ("Revenue", "Expense") or r["dept"] not in KEEP: continue
    if r["dept"] == "Clubs (Fund 413)" and not r["dnum"].startswith("73") or r["dnum"] in ("73029", "73034"): continue
    k = r["key"]
    code = f"{k[:3]}-{k[3:8]}-{k[8:]}"   # same format as the claim form and the Expense Ledger
    name = (r["claimname"].replace(" / Mileage", "").strip() if r["claimname"] else "") or r["desc"]
    if r["dept"] == "Clubs (Fund 413)":
        name = (r["claimname"] or r["club"]) + (" - Transfers In" if r["type"] == "Revenue" else " - Club expenses")
    notes = r["notes"]
    out.append([code, name, GROUP.get(r["dept"], r["dept"]), r["type"], r["cat"], "Yes" if r["claim"] else "",
                round(r["budget"], 2) if r["budget"] else "", r["jun"], notes])
exp = sum(1 for o in out if o[3] == "Expense")
js = f"""/**
 * NUSA — GL code list  (script bound to "NUSA Expense Claims — Log")
 * ------------------------------------------------------------------
 * updateGlList rewrites the "GL List" tab with every NUSA revenue and expense GL code of the ten departments and
 * the clubs ({len(out)} codes: {len(out) - exp} revenue, {exp} expense), from the master GL workbook.
 * Columns: GL Code, GL Name, Group, Type (Revenue / Expense), Category,
 * In claim form?, 2026-27 Budget, On NU list?, Notes.
 * The Expense Ledger's GL Code dropdown and NUSA > Add manual expense only offer the Expense codes.
 * Codes already on the tab that aren't in this list (e.g. added by the claim form) are kept at the bottom.
 * The claim web app's refreshGlList keeps this list too.
 *
 * SETUP: in the claim log, Extensions > Apps Script > + > Script editor, name it "GL Codes", paste this
 * file, Save. Pick updateGlList next to Run and press Run. Run it again whenever you paste a new copy.
 * Don't edit the list below by hand; it is generated from the master GL file.
 */
const GLC_TAB = 'GL List';
const GLC_HEAD = ['GL Code', 'GL Name', 'Group', 'Type', 'Category', 'In claim form?', '2026-27 Budget', 'On NU list?', 'Notes'];
const GL_CODES = {json.dumps(out, ensure_ascii=False, separators=(",", ":"))};

function updateGlList() {{
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(GLC_TAB) || ss.insertSheet(GLC_TAB);
  const known = {{}};
  GL_CODES.forEach(function (r) {{ known[r[0]] = true; }});
  const kept = sh.getDataRange().getValues().slice(1).filter(function (r) {{ return r[0] !== '' && !known[String(r[0]).trim()]; }})
    .map(function (r) {{ return [String(r[0]).trim(), r[1] || '', r[2] || '', String(r[3]).trim() === 'Revenue' ? 'Revenue' : 'Expense', r[4] || '', '', '', '', 'Kept from the previous GL List (not in the master list)']; }});
  const rows = GL_CODES.concat(kept);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  sh.getRange(1, 1, 1, GLC_HEAD.length).setValues([GLC_HEAD]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.getRange(2, 1, rows.length, GLC_HEAD.length).setValues(rows);
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('$#,##0.00;($#,##0.00);"-"');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, rows.length + 1, GLC_HEAD.length).createFilter();
  [270, 300, 170, 80, 220, 90, 110, 90, 380].forEach(function (w, i) {{ sh.setColumnWidth(i + 1, w); }});
  const expense = rows.filter(function (r) {{ return r[3] === 'Expense'; }}).map(function (r) {{ return r[0]; }});
  const led = ss.getSheetByName('Expense Ledger');
  if (led) {{
    const col = led.getRange(1, 1, 1, led.getLastColumn()).getValues()[0].indexOf('GL Code');
    if (col >= 0) led.getRange(2, col + 1, 3000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(expense, true).setAllowInvalid(true).build());
  }}
  console.log('GL List updated: ' + rows.length + ' codes (' + expense.length + ' expense, ' + (rows.length - expense.length) + ' revenue)'
    + (kept.length ? '; ' + kept.length + ' kept from before' : '') + '.');
  return rows.length;
}}
"""
open(OUT, "w", encoding="utf-8").write(js)
print(len(out), "codes,", exp, "expense ->", OUT)
