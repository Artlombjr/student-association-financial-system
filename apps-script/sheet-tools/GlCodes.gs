/**
 * NUSA - GL code list  (script bound to "NUSA Expense Claims - Log")
 * ------------------------------------------------------------------
 * updateGlList rewrites the "GL List" tab with every revenue and expense GL code in GL_CODES below (demo data:
 * the sample budget in TrackerData.gs). Columns: GL Code, GL Name, Group, Type (Revenue / Expense), Category,
 * In claim form?, 2026-27 Budget, On NU list?, Notes.
 * The Expense Ledger's GL Code dropdown and NUSA > Add manual expense only offer the Expense codes.
 * Codes already on the tab that aren't in this list (e.g. added by the claim form) are kept at the bottom.
 * The claim web app's refreshGlList keeps this list too.
 *
 * SETUP: in the claim log, Extensions > Apps Script > + > Script, name it "GL Codes", paste this
 * file, Save. Pick updateGlList in the function menu and press Run. Run it again whenever you paste a new copy.
 * In a real deployment this list is generated from the organisation's master GL file; don't edit it by hand.
 */
const GLC_TAB = 'GL List';
const GLC_HEAD = ['GL Code', 'GL Name', 'Group', 'Type', 'Category', 'In claim form?', '2026-27 Budget', 'On NU list?', 'Notes'];
const GL_CODES = [
["410-72103-41035100003300610000","Student Association Fees","Administration","Revenue","Revenue","",175000,"Yes",""],
["410-72103-41350100001000010000","Income - Miscellaneous","Administration","Revenue","Revenue","",500,"Yes",""],
["410-72103-41650100001000010000","Student Publication Fees","Administration","Revenue","Revenue","",0,"Yes",""],
["410-72103-50320100001000010000","Wages - Part Time Staff","Administration","Expense","Wages","",2300,"Yes",""],
["410-72103-50340100003802310000","Wages - Student Work (Grants)","Administration","Expense","Wages","Yes",28000,"Yes",""],
["410-72103-56727100004039510000","Serv. - Bank Service Charges","Administration","Expense","General Admin Expenses","",700,"Yes",""],
["410-72103-56025100003906310000","Special Events","Administration","Expense","General Admin Expenses","Yes",8000,"Yes",""],
["410-72103-55254100001000010000","Equipment Rental- IT Copier Charges","Administration","Expense","General Admin Expenses","Yes",1500,"Yes",""],
["410-72103-55610100001000010000","Office - Small Equipment","Administration","Expense","General Admin Expenses","Yes",1000,"Yes",""],
["410-72103-56350100001000010000","Hospitality / Food","Administration","Expense","General Admin Expenses","Yes",2500,"Yes",""],
["410-72104-42100100001000010000","Donation - Receiptable","Faculty","Revenue","Revenue","",0,"Yes",""],
["410-72104-56025100003902210000","Events - Business Representative","Faculty","Expense","Faculty Representatives","Yes",1000.0,"Yes",""],
["410-72104-56025100003902310000","Events - Education Representative","Faculty","Expense","Faculty Representatives","Yes",1000.0,"Yes",""],
["410-72104-56025100003902410000","Events - Human Kinetics Representative","Faculty","Expense","Faculty Representatives","Yes",1000.0,"Yes",""],
["410-72104-56025100003902510000","Events - Humanities Representative","Faculty","Expense","Faculty Representatives","Yes",1000.0,"Yes",""],
["410-72104-56025100003902610000","Events - NATS Representative","Faculty","Expense","Faculty Representatives","Yes",1000.0,"Yes",""],
["410-72104-56025100003902710000","Events - Nursing Representative","Faculty","Expense","Faculty Representatives","Yes",1000.0,"Yes",""],
["410-72104-56025100001000010000","Events - SAMC Representative (MixEx/EventHosting)","Faculty","Expense","Faculty Representatives","Yes",1000.0,"Yes",""],
["410-72104-50340100001000010000","Wages - Student Work (Stipends)","Faculty","Expense","Wages","",3400.0,"Yes",""],
["413-73002-56025100003904410000","Nursing Club","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["413-73006-56025100003904410000","Biology Club","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["413-73010-56025100001000010000","Afro-Caribbean Club","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["413-73016-56025100001000010000","The Chemistry Club","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["413-73024-56025100003904410000","Games Club","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["413-73026-56025100001000010000","Chinese Assoc. (CSSA) Club","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["413-73041-56025100001000010000","Circus Guild Club (HKIN)","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["413-73042-56025100001000010000","Tabletop Roleplaying Club (SAMC)","Clubs (Fund 413)","Expense","Clubs","Yes",700.0,"Yes",""],
["410-72111-56025100003907010000","NUSA Funding Clubs","Proposal Committee","Expense","(No category)","Yes",4000.0,"Yes",""],
["410-72111-56025100001000010000","MixEx/EventHosting","Proposal Committee","Expense","(No category)","Yes",2700.0,"Yes",""],
["410-72105-42200100001000010000","Advertising Revenue","Communications","Revenue","Revenue","",1500,"Yes",""],
["410-72105-41350100001000010000","Income - Miscellaneous","Communications","Revenue","Revenue","",0,"Yes",""],
["410-72105-55050100001000010000","Comm. - Publicity and Advertising","Communications","Expense","(No category)","Yes",9000,"Yes",""],
["410-72105-55600100001000010000","MTRL/GenSup/OfficeSuppl","Communications","Expense","(No category)","Yes",800,"Yes",""],
["410-72105-55610100001000010000","Office - Small Equipment","Communications","Expense","(No category)","Yes",600,"Yes",""],
["410-72106-44360100001000010000","Events Revenue","Events","Revenue","Revenue","",5500,"Yes",""],
["410-72106-42500100001000010000","Sponsorship Income","Events","Revenue","Revenue","",0,"Yes",""],
["410-72106-56025100003902110000","Events","Events","Expense","(No category)","Yes",27000,"Yes",""],
["410-72106-55275100001000010000","SVC/ICAContract/ICA","Events","Expense","(No category)","Yes",5000,"Yes",""],
["410-72107-50340100003802310000","Wages - Student Work (Grants)","The Courier","Expense","Wages","Yes",14000,"Yes",""],
["410-72107-55225100001000010000","SVC/CommSvc/PrintPubl","The Courier","Expense","(No category)","Yes",9000,"Yes",""],
["410-72107-55600100001000010000","MTRL/GenSup/OfficeSuppl","The Courier","Expense","(No category)","Yes",500,"Yes",""],
["410-72107-55610100001000010000","MTRL/GenSup/SmallEquipment","The Courier","Expense","(No category)","Yes",400,"Yes",""],
["410-72107-55612100003902110000","Food Costs","The Courier","Expense","(No category)","Yes",600,"Yes",""],
["410-72107-56025100003906310000","Special Events","The Courier","Expense","(No category)","Yes",500,"Yes",""],
["410-72108-41650100001000010000","Yearbook Sales","Yearbook","Revenue","Revenue","",1000,"Yes",""],
["410-72108-50340100003802310000","Wages - Student Work (Grants)","Yearbook","Expense","Wages","Yes",12000,"Yes",""],
["410-72108-55225100001000010000","Comm. - Printing and Publications","Yearbook","Expense","(No category)","Yes",14000,"Yes",""],
["410-72108-55600100001000010000","MTRL/GenSup/OfficeSuppl","Yearbook","Expense","(No category)","Yes",300,"Yes",""],
["410-72108-55610100001000010000","MTRL/GenSup/SmallEquipment","Yearbook","Expense","(No category)","Yes",400,"Yes",""],
["410-72108-55618100001000010000","Office - Computer Supplies","Yearbook","Expense","(No category)","Yes",900,"Yes",""],
["410-72108-56025100003906310000","Special Events","Yearbook","Expense","(No category)","Yes",400,"Yes",""],
["410-72109-55225100001000010000","Office - Copywork","Student Lounge (The Commons)","Expense","(No category)","Yes",800,"Yes",""],
["410-72109-55600100001000010000","Office - Miscellaneous","Student Lounge (The Commons)","Expense","(No category)","Yes",3000,"Yes",""],
["410-72109-55610100001000010000","Office - Eqmt & Furnishings","Student Lounge (The Commons)","Expense","(No category)","Yes",2500,"Yes",""],
["410-72109-55612100003902110000","Food Costs","Student Lounge (The Commons)","Expense","(No category)","Yes",9000,"Yes",""],
["410-72109-56254100001000010000","Rental Services","Student Lounge (The Commons)","Expense","(No category)","Yes",1200,"Yes",""],
["410-72103-56302100001000010000","Administration - Travel-Mileage (Personal Vehicle)","Student Lounge (The Commons)","Expense","(No category)","Yes",2000,"Yes",""],
["410-72110-56025100001000010000","MixEx/EventHosting","Special Initiatives","Expense","(No category)","Yes",1500,"Yes",""],
["410-73001-66000100004170112510","CapEquip/General - Lounge Improvement","Capital Projects","Expense","(No category)","Yes",12000,"Yes",""],
["410-73001-66000100004230010000","CapEquip/General (Communications)","Capital Projects","Expense","(No category)","Yes",6000,"Yes",""]
];

function updateGlList() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(GLC_TAB) || ss.insertSheet(GLC_TAB);
  const known = {};
  GL_CODES.forEach(function (r) { known[r[0]] = true; });
  const kept = sh.getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== '' && !known[String(r[0]).trim()]; })
    .map(function (r) { return [String(r[0]).trim(), r[1] || '', r[2] || '', String(r[3]).trim() === 'Revenue' ? 'Revenue' : 'Expense', r[4] || '', '', '', '', 'Kept from the previous GL List (not in the master list)']; });
  const rows = GL_CODES.concat(kept);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  sh.getRange(1, 1, 1, GLC_HEAD.length).setValues([GLC_HEAD]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.getRange(2, 1, rows.length, GLC_HEAD.length).setValues(rows);
  sh.getRange(2, 7, rows.length, 1).setNumberFormat('$#,##0.00;($#,##0.00);"-"');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, rows.length + 1, GLC_HEAD.length).createFilter();
  [270, 300, 170, 80, 220, 90, 110, 90, 380].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  const expense = rows.filter(function (r) { return r[3] === 'Expense'; }).map(function (r) { return r[0]; });
  const led = ss.getSheetByName('Expense Ledger');
  if (led) {
    const col = led.getRange(1, 1, 1, led.getLastColumn()).getValues()[0].indexOf('GL Code');
    if (col >= 0) led.getRange(2, col + 1, 3000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(expense, true).setAllowInvalid(true).build());
  }
  console.log('GL List updated: ' + rows.length + ' codes (' + expense.length + ' expense, ' + (rows.length - expense.length) + ' revenue)'
    + (kept.length ? '; ' + kept.length + ' kept from before' : '') + '.');
  return rows.length;
}
