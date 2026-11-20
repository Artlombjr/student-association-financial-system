import os, json, datetime
HERE = os.path.dirname(os.path.abspath(__file__))
exec(open(os.path.join(HERE, "parse.py"), encoding="utf-8-sig").read())
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule

F = "Arial"
NAVY = "1F3864"
H = Font(name=F, bold=True, color="FFFFFF", size=10); HF = PatternFill("solid", fgColor=NAVY)
CALC_HF = PatternFill("solid", fgColor="595959")
B = Font(name=F, size=10); BI = Font(name=F, size=10, color="0000FF"); BB = Font(name=F, size=10, bold=True)
T = Font(name=F, size=16, bold=True, color=NAVY); S = Font(name=F, size=9, italic=True, color="595959")
SEC = Font(name=F, size=12, bold=True, color=NAVY)
MONEY = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;-0.0%;"-"'
YELLOW = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="BFBFBF")
DEPTS = [t[1] for t in TABS]

def header(ws, row, cols, widths=None, start=1, calc=()):
    for i, c in enumerate(cols):
        x = ws.cell(row, start + i, c); x.font = H
        x.fill = CALC_HF if c in calc else HF
        x.alignment = Alignment(wrap_text=True, vertical="center")
        if widths: ws.column_dimensions[L(start + i)].width = widths[i]
    ws.row_dimensions[row].height = 30

# ---- GL codes the claim system offers that the AGM budget doesn't have: add them as $0 lines
FORMS_CODE = os.environ.get("FORMS_CODE", os.path.join(HERE, "..", "apps-script", "claim-web-app", "Code.gs"))   # the claim form's GL list
GROUP_DEPT = {"Administration": "Administration", "Events": "Events", "Communications": "Communications", "Faculty": "Faculty",
              "Yearbook": "Yearbook", "The Courier": "The Courier", "Student Lounge (The Commons)": "Student Lounge (The Commons)",
              "Special Initiatives": "Special Initiatives", "Proposal Committee": "Proposal Committee",
              "Capital Projects": "Capital Projects", "Clubs (Fund 413)": "Faculty"}
DNUM = {t[1]: t[2] for t in TABS}
digits = lambda s: re.sub(r"\D", "", s or "")
_src = open(FORMS_CODE, encoding="utf-8").read()
_blk = _src[_src.index("const GL = ["):]; _blk = _blk[:_blk.index("];")]
sys_codes = [(g, n, c) for g, body in re.findall(r'\{g:"([^"]+)", items:\[(.*?)\]\s*\}', _blk, re.S)
             for n, c in re.findall(r'\["([^"]*)","([^"]*)"\]', body)]
_m = re.search(r"const MILEAGE_GL = '([^']+)'", _src); _mn = re.search(r"const MILEAGE_GL_NAME = '([^']+)'", _src)
sys_codes.append(("Administration", _mn.group(1), _m.group(1)))
_legacy = os.path.join(HERE, "legacy_gl_codes.txt")
if os.path.exists(_legacy):
    for _ln in open(_legacy, encoding="utf-8"):
        if _ln.strip() and not _ln.startswith("#"):
            _g, _n, _c = _ln.rstrip("\r\n").split("\t"); sys_codes.append((_g, _n, _c))
# Every expense code of the master GL list (the Google Sheet's "GL List" tab), so any code on the ledger has a line.
_glc = os.environ.get("GLCODES_GS", os.path.join(HERE, "..", "apps-script", "sheet-tools", "GlCodes.gs"))
if os.path.exists(_glc):
    _t = open(_glc, encoding="utf-8").read()
    _rows = json.loads(re.search(r"const GL_CODES = (\[.*?\]);\n", _t, re.S).group(1))
    sys_codes += [(r[2], r[1], r[0]) for r in _rows if r[3] == "Expense"]
known ={digits(p["gl"]) for p in parents if p["section"] == "Expenditure" and digits(p["gl"])}
added = 0
for g, name, code in sys_codes:
    d = digits(code)
    if d in known: continue
    known.add(d)
    dept = GROUP_DEPT.get(g, g if g in DNUM else "Administration")
    gl_fmt = code.split("-")[0] + "  " + "".join(code.split("-")[1:])
    p = dict(tab="(Claim system)", dept=dept, dnum=DNUM[dept], section="Expenditure",
             category="Clubs (not in AGM budget)" if g.startswith("Clubs") else "(Not in AGM budget)",
             desc=name, gl=gl_fmt, actual=0, budget=0, lvl=2, src="Claim system GL list (V5)", kids=[], r=0,
             split=dept in SPLIT_DEPTS)
    parents.append(p); added += 1
    for sem in (("Fall", "Spring") if p["split"] else ("Full year",)):
        items.append(dict(p=p, desc=name, group=None, budget=0, actual=0, src=p["src"], incl="Yes", sem=sem))
print("GL codes added from the claim system:", added)

# unique tracker label per GL row
seen = collections.Counter()
for p in parents:
    m = re.match(r"4\d\d\s+(\d{5})(\d{5})", p["gl"])
    p["acct"] = m.group(2) if m else ""; p["gldept"] = m.group(1) if m else ""
    base = f"{p['dept']} | {p['desc']} | {p['acct'] or 'no GL'}"
    if p["section"] == "Revenue": base = "[Revenue] " + base
    seen[base] += 1
    p["label"] = base if seen[base] == 1 else f"{base} ({seen[base]})"

out = openpyxl.Workbook()
com = {"formula2": [], "calc_cols": []}   # applied afterwards through Excel

# ================= Lists =================
ls = out.active; ls.title = "Lists"
exp_labels = [p["label"] for p in parents if p["section"] == "Expenditure"]
cols = ["Budget Lines (expenses)", "Departments", "Status", "Payment Method", "Dashboard Department"]
header(ls, 1, cols, [70, 26, 14, 22, 26])
for i, v in enumerate(exp_labels): ls.cell(2 + i, 1, v).font = B
for i, v in enumerate(DEPTS): ls.cell(2 + i, 2, v).font = B
STAT = ["Pending", "Approved", "Posted", "Example"]
for i, v in enumerate(STAT): ls.cell(2 + i, 3, v).font = B
PAY = ["P-Card", "Reimbursement", "Cheque Requisition", "Invoice / PO", "Internal Transfer", "Other"]
for i, v in enumerate(PAY): ls.cell(2 + i, 4, v).font = B
for i, v in enumerate(["All departments"] + DEPTS): ls.cell(2 + i, 5, v).font = B
N_LAB = len(exp_labels)

# ================= GL Budget =================
gb = out.create_sheet("GL Budget", 0)
gb["A1"] = "2026-27 Budget by GL Code"; gb["A1"].font = T
gb["A2"] = ("One row per GL line on each department tab. Budget = sum of the GL's line items. "
            "Spent/Pending come from the Expense Ledger sheet (Data > Refresh All). Use the filter arrows to filter by Department, Category, GL Code or Has Budget?.")
gb["A2"].font = S
cols = ["GL Key", "Department (Tab)", "Dept #", "Section", "Category", "GL Description", "GL Code",
        "2026-27 Budget", "Spent (Approved + Posted)", "Pending", "Remaining", "% Used", "Status", "Has Budget?",
        "2025-26 Actual", "# Line Items", "GL Dept #", "Account #", "Same GL Code: Total Across All Tabs",
        "Tracker Label (used in Transactions)", "Source Cell (AGM file)",
        "Split by Semester?", "Fall Budget", "Fall Spent", "Spring Budget", "Spring Spent",
        "Tracking Period", "Period Budget", "Period Spent", "Period Pending", "Period Remaining"]
calc = {"Fall Budget", "Fall Spent", "Spring Budget", "Spring Spent", "Tracking Period", "Period Budget", "Period Spent", "Period Pending", "Period Remaining",
        "2026-27 Budget", "Spent (Approved + Posted)", "Pending", "Remaining", "% Used", "Status", "Has Budget?",
        "# Line Items", "Same GL Code: Total Across All Tabs"}
header(gb, 4, cols, [7, 22, 8, 12, 24, 34, 32, 14, 14, 12, 14, 9, 16, 9, 14, 8, 9, 9, 15, 50, 24, 9, 12, 12, 12, 12, 10, 13, 13, 13, 13], calc=calc)
gb.cell(4, 12, "% Used (this period)"); gb.cell(4, 13, "Status (this period)")
SEM = "Dashboard!$H$5"
n = len(parents); G = 4 + n; last_li = 4 + len(items)
TX = "'Expense Ledger'"
for i, p in enumerate(parents):
    r = 5 + i
    vals = [i + 1, p["dept"], p["dnum"], p["section"], p["category"], p["desc"], p["gl"],
            f"=SUMIFS('Line Items'!$K$5:$K${last_li},'Line Items'!$N$5:$N${last_li},A{r},'Line Items'!$L$5:$L${last_li},\"Yes\")",
            f'=SUMIFS({TX}!$J:$J,{TX}!$E:$E,$T{r},{TX}!$N:$N,"Approved")+SUMIFS({TX}!$J:$J,{TX}!$E:$E,$T{r},{TX}!$N:$N,"Posted")',
            f'=SUMIFS({TX}!$J:$J,{TX}!$E:$E,$T{r},{TX}!$N:$N,"Pending")',
            f"=H{r}-I{r}",
            f'=IF(AB{r}>0,AC{r}/AB{r},"")',
            f'=IF(AND(AB{r}=0,AC{r}=0),"-",IF(AB{r}=0,"Unbudgeted spend",IF(AC{r}>AB{r},"Over budget",IF(AC{r}/AB{r}>=0.9,"Almost out",IF(AC{r}/AB{r}>=0.75,"Watch","On track")))))',
            f'=IF(H{r}<>0,"Yes","No")',
            p["actual"], f"=COUNTIFS('Line Items'!$N$5:$N${last_li},A{r})", p["gldept"], p["acct"],
            f'=IF(LEFT(G{r},1)="(","",SUMIFS(H$5:H${G},G$5:G${G},G{r},D$5:D${G},D{r}))',
            p["label"], p["src"],
            "Yes" if p["split"] else "No"]
    for sem in ("Fall", "Spring"):
        vals.append(f"=SUMIFS('Line Items'!$K$5:$K${last_li},'Line Items'!$N$5:$N${last_li},A{r},'Line Items'!$L$5:$L${last_li},\"Yes\",'Line Items'!$O$5:$O${last_li},\"{sem}\")")
        vals.append(f'=SUMIFS({TX}!$J:$J,{TX}!$E:$E,$T{r},{TX}!$N:$N,"Approved",{TX}!$O:$O,"{sem}")+SUMIFS({TX}!$J:$J,{TX}!$E:$E,$T{r},{TX}!$N:$N,"Posted",{TX}!$O:$O,"{sem}")')
    vals += [f'=IF(V{r}="Yes",{SEM},"Full year")',
             f'=IF(V{r}="Yes",IF(AA{r}="Fall",W{r},Y{r}),H{r})',
             f'=IF(V{r}="Yes",IF(AA{r}="Fall",X{r},Z{r}),I{r})',
             f'=IF(V{r}="Yes",SUMIFS({TX}!$J:$J,{TX}!$E:$E,$T{r},{TX}!$N:$N,"Pending",{TX}!$O:$O,AA{r}),J{r})',
             f"=AB{r}-AC{r}"]
    for c, v in enumerate(vals, 1):
        x = gb.cell(r, c, v); x.font = BI if c == 15 else B
        if c in (8, 9, 10, 11, 15, 19) or 23 <= c <= 26 or 28 <= c <= 31: x.number_format = MONEY
        if c == 12: x.number_format = PCT
gb.add_table(Table(displayName="GLBudget", ref=f"A4:AE{G}", tableStyleInfo=TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)))
gb.freeze_panes = "C5"
tr = G + 2
gb.cell(tr, 6, "Total of visible rows (updates with filters):").font = BB
for c in (8, 9, 10, 11, 15):
    x = gb.cell(tr, c, f"=SUBTOTAL(109,{L(c)}5:{L(c)}{G})"); x.font = BB; x.number_format = MONEY
gb.conditional_formatting.add(f"A5:AE{G}", FormulaRule(formula=['$N5="No"'], font=Font(color="A6A6A6")))

# ================= Line Items =================
li = out.create_sheet("Line Items", 1)
li["A1"] = "2026-27 Budget: Line Items by GL Code"; li["A1"].font = T
li["A2"] = "Every budgeted line from the AGM budget, tied to the GL code it rolls up to. Blue = values copied from the AGM file (see Source Cell)."; li["A2"].font = S
cols = ["Line ID", "Department (Tab)", "Dept #", "Section", "Category", "GL Description", "GL Code", "Sub-group", "Line Item",
        "2025-26 Actual", "2026-27 Budget", "Counted in AGM Total?", "Source Cell (AGM file)", "GL Key", "Semester"]
header(li, 4, cols, [8, 22, 8, 12, 24, 34, 32, 14, 40, 14, 14, 12, 24, 8, 11])
for i, it in enumerate(items):
    r = 5 + i; p = it["p"]
    vals = [i + 1, p["dept"], p["dnum"], p["section"], p["category"], p["desc"], p["gl"], it["group"] or "", it["desc"],
            it["actual"], it["budget"], it["incl"], it["src"], parents.index(p) + 1, it["sem"]]
    for c, v in enumerate(vals, 1):
        x = li.cell(r, c, v); x.font = BI if c in (10, 11) else B
        if c in (10, 11): x.number_format = MONEY
li.add_table(Table(displayName="LineItems", ref=f"A4:O{last_li}", tableStyleInfo=TableStyleInfo(name="TableStyleLight9", showRowStripes=True)))
li.freeze_panes = "C5"
tr = last_li + 2
li.cell(tr, 9, "Total of visible rows:").font = BB
for c in (10, 11):
    x = li.cell(tr, c, f"=SUBTOTAL(109,{L(c)}5:{L(c)}{last_li})"); x.font = BB; x.number_format = MONEY

# ================= GL Map (GL code -> budget line) =================
gm = out.create_sheet("GL Map")
gm["A1"] = "GL Map: which budget line each GL code is charged to"; gm["A1"].font = T
gm["A2"] = ("Used by the Expense Ledger import. Key = GL code digits | department. A row with an empty department is the default "
            "when no 'Charge to Department' is set on the ledger row. To send a GL code to a different budget line, change its Budget Line here "
            "(copy the label from GL Budget, column T).")
gm["A2"].font = S
header(gm, 4, ["Key", "GL Code", "Department", "Budget Line", "Why"], [40, 32, 24, 70, 50])
by_digits = collections.defaultdict(list)
for p in parents:
    if p["section"] == "Expenditure" and digits(p["gl"]): by_digits[digits(p["gl"])].append(p)
bud_of = lambda p: sum(it["budget"] for it in items if it["p"] is p and it["incl"] == "Yes")
map_rows = []
for d, ps in by_digits.items():
    for dept in dict.fromkeys(p["dept"] for p in ps):
        cand = max((p for p in ps if p["dept"] == dept), key=bud_of)
        why = "" if sum(p["dept"] == dept for p in ps) == 1 else "Shared by several lines in this department: largest budget chosen"
        map_rows.append((d + "|" + dept, cand["gl"], dept, cand["label"], why))
    best = max(ps, key=lambda p: (bud_of(p), p["dnum"] == d[3:8]))
    why = "" if len(ps) == 1 else "On several budget lines: default is the one with the largest budget (use 'Charge to Department' to pick another)"
    map_rows.append((d + "|", best["gl"], "", best["label"], why))
for i, row in enumerate(sorted(map_rows, key=lambda x: x[0])):
    for c, v in enumerate(row, 1):
        gm.cell(5 + i, c, v).font = BI if c == 4 else B
gm.add_table(Table(displayName="GLMap", ref=f"A4:E{4 + len(map_rows)}", tableStyleInfo=TableStyleInfo(name="TableStyleLight9", showRowStripes=True)))
gm.freeze_panes = "A5"

# ================= Settings =================
st_ = out.create_sheet("Settings")
st_.column_dimensions["A"].width = 26; st_.column_dimensions["B"].width = 110
st_["A1"] = "Settings"; st_["A1"].font = T
st_["A3"] = "Expense Ledger link (private)"; st_["A3"].font = BB
st_["B3"] = ""; st_["B3"].fill = YELLOW; st_["B3"].font = BI
st_["B4"] = ("Paste the link that setupExpenseLedger() prints in Apps Script (your web app's /exec address followed by ?feed=ledger&key=...). "
             "It returns only the Expense Ledger rows, so the claim log itself can stay Restricted. Treat it like a password: anyone with it can read the ledger. "
             "If it leaks, run rotateLedgerFeedKey() in Apps Script and paste the new link here. Leave blank to load nothing.")
st_["B4"].font = S; st_["B4"].alignment = Alignment(wrap_text=True)
from openpyxl.workbook.defined_name import DefinedName
out.defined_names["LedgerURL"] = DefinedName("LedgerURL", attr_text="Settings!$B$3")

# ================= Expense Ledger (filled by Power Query in apply.ps1) =================
tx = out.create_sheet("Expense Ledger", 0)
tx["A1"] = "Expense Ledger"; tx["A1"].font = T
tx["A2"] = ("Imported from the 'Expense Ledger' tab of the NUSA Expense Claims log: claims arrive there automatically; add or edit purchases there too. "
            "Update with Data > Refresh All. Don't edit here: changes are replaced on the next refresh.")
tx["A2"].font = S
tx["A3"] = ("Counted: Approved and Posted = Spent, Pending = Pending, Void = ignored (blank Status on a manual row counts as Approved). "
            "Budget Line comes from the GL Map. Matched? = No means the row isn't counted: its GL code isn't on the map or its Status isn't Pending/Approved/Posted/Void.")
tx["A3"].font = S
LEDGER_COLS = ["Entry ID", "Source", "Date", "Department", "Budget Line", "GL Code", "GL Name", "Description", "Payee",
               "Amount (incl. tax)", "GST", "HST", "Net Expense", "Status", "Semester", "Month", "Claim ID", "Claim Status", "Notes", "Matched?"]
for i, w in enumerate([12, 9, 11, 22, 50, 32, 30, 40, 22, 13, 9, 9, 12, 10, 10, 9, 11, 26, 30, 9], 1):
    tx.column_dimensions[L(i)].width = w
tx.freeze_panes = "D5"
com["ledger"] = {"sheet": "Expense Ledger", "cell": "A4", "columns": LEDGER_COLS}

# ================= Dashboard =================
db = out.create_sheet("Dashboard", 0)
widths = {"A": 2, "B": 24, "C": 32, "D": 38, "E": 14, "F": 14, "G": 14, "H": 16, "I": 14, "J": 10, "K": 17}
for k, w in widths.items(): db.column_dimensions[k].width = w
db["B1"] = "NUSA 2026-27 Budget Tracker"; db["B1"].font = T
db["B2"] = '="As of "&TEXT(TODAY(),"mmmm d, yyyy")&"   |   Fiscal year May 1, 2026 to April 30, 2027"'; db["B2"].font = S
GB = lambda col: f"'GL Budget'!${col}$5:${col}${G}"
# KPI tiles
kpis = [("Expense budget", f'=SUMIFS({GB("H")},{GB("D")},"Expenditure")', MONEY),
        ("Spent", f'=SUMIFS({GB("I")},{GB("D")},"Expenditure")', MONEY),
        ("Pending", f'=SUMIFS({GB("J")},{GB("D")},"Expenditure")', MONEY),
        ("Remaining", "=B5-C5", MONEY),
        ("% of budget used", '=IF(B5>0,C5/B5,"")', PCT),
        ("% of year gone", '=MAX(0,MIN(1,(TODAY()-DATE(2026,5,1))/(DATE(2027,4,30)-DATE(2026,5,1)+1)))', PCT)]
tile = PatternFill("solid", fgColor="EEF2F8")
for i, (lab, f, fmt) in enumerate(kpis):
    c = 2 + i
    a = db.cell(4, c, lab); a.font = Font(name=F, size=9, color="595959"); a.fill = tile
    b = db.cell(5, c, f); b.font = Font(name=F, size=14, bold=True, color=NAVY); b.fill = tile; b.number_format = ("$#,##0" if fmt == MONEY else fmt)
    b.alignment = Alignment(horizontal="left")
a = db.cell(4, 9, "Unmatched entries"); a.font = Font(name=F, size=9, color="595959"); a.fill = tile
b = db.cell(5, 9, "=COUNTIFS('Expense Ledger'!$T:$T,\"No\",'Expense Ledger'!$N:$N,\"<>Void\")"); b.font = Font(name=F, size=14, bold=True, color=NAVY); b.fill = tile
a = db.cell(4, 8, "Current semester"); a.font = Font(name=F, size=9, color="595959"); a.fill = tile
b = db.cell(5, 8, '=IF(MONTH(TODAY())<=4,"Spring","Fall")'); b.font = Font(name=F, size=14, bold=True, color=NAVY); b.fill = tile
db.row_dimensions[5].height = 24

db["B7"] = "By department (expenses)"; db["B7"].font = SEC
header(db, 8, ["Department", "Budget", "Spent", "Pending", "Remaining", "% Used"], start=2)
for i, d in enumerate(DEPTS):
    r = 9 + i
    db.cell(r, 2, d).font = B
    for c, col in ((3, "H"), (4, "I"), (5, "J")):
        x = db.cell(r, c, f'=SUMIFS({GB(col)},{GB("B")},$B{r},{GB("D")},"Expenditure")'); x.font = B; x.number_format = MONEY
    x = db.cell(r, 6, f"=C{r}-D{r}"); x.font = B; x.number_format = MONEY
    x = db.cell(r, 7, f'=IF(C{r}>0,D{r}/C{r},"")'); x.font = B; x.number_format = PCT
rt = 9 + len(DEPTS)
db.cell(rt, 2, "Total").font = BB
for c in range(3, 7):
    x = db.cell(rt, c, f"=SUM({L(c)}9:{L(c)}{rt - 1})"); x.font = BB; x.number_format = MONEY; x.border = Border(top=thin)
x = db.cell(rt, 7, f'=IF(C{rt}>0,D{rt}/C{rt},"")'); x.font = BB; x.number_format = PCT; x.border = Border(top=thin)
db.conditional_formatting.add(f"G9:G{rt - 1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))

r0 = rt + 3
db.cell(r0, 2, "GL codes ranked: closest to budget first").font = SEC
db.cell(r0 + 1, 2, "Department:").font = BB
sel = db.cell(r0 + 1, 3, "All departments"); sel.font = BB; sel.fill = YELLOW
sel.border = Border(left=thin, right=thin, top=thin, bottom=thin)
db.cell(r0 + 1, 4, "Pick one department or All departments. Each department is ranked by % used, then by money left. Faculty lines show the current semester's half (Period column); other lines show the full year.").font = S
dvs = DataValidation(type="list", formula1=f"=Lists!$E$2:$E${2 + len(DEPTS)}", allow_blank=False)
db.add_data_validation(dvs); dvs.add(sel.coordinate)
hr = r0 + 2
header(db, hr, ["Department", "GL Code", "GL Description", "Period", "Budget", "Spent", "Pending", "Remaining", "% Used", "Status"], start=2)
first = hr + 1
spill = (f"=LET(dep,{GB('B')},gl,{GB('G')},ds,{GB('F')},bud,{GB('H')},sp,{GB('I')},pe,{GB('J')},rem,{GB('K')},"
         f"pu,{GB('L')},st,{GB('M')},sec,{GB('D')},sel,$C${r0 + 1},per,{GB('AA')},pb,{GB('AB')},psp,{GB('AC')},ppe,{GB('AD')},prem,{GB('AE')},"
         f"keep,(sec=\"Expenditure\")*(((bud<>0)+(sp<>0)+(pe<>0))>0)*(((sel=\"All departments\")+(dep=sel))>0),"
         f"ord,MATCH(dep,Lists!$B$2:$B${1 + len(DEPTS)},0),key,IF(pb>0,psp/pb,IF(psp>0,1000,-1)),"
         f"data,HSTACK(dep,gl,ds,per,pb,psp,ppe,prem,pu,st),"
         f"IFERROR(SORTBY(FILTER(data,keep),FILTER(ord,keep),1,FILTER(key,keep),-1,FILTER(prem,keep),1),\"No lines to show\"))")
com["formula2"].append({"sheet": "Dashboard", "cell": f"B{first}", "formula": spill})
last = first + 300
for r in range(first, last):
    for c in range(2, 12):
        x = db.cell(r, c); x.font = B
        if c in (6, 7, 8, 9): x.number_format = MONEY
        if c == 10: x.number_format = PCT
rng = f"B{first}:K{last}"
db.conditional_formatting.add(rng, FormulaRule(formula=[f'AND($B{first}<>"",$B{first}<>$B{first - 1})'],
                              border=Border(top=Side(style="medium", color=NAVY)), font=Font(bold=True)))
for txt, color in (("Over budget", "F8CBAD"), ("Unbudgeted spend", "F8CBAD"), ("Almost out", "FCE4D6"), ("Watch", "FFF2CC"), ("On track", "E2EFDA")):
    db.conditional_formatting.add(f"K{first}:K{last}", FormulaRule(formula=[f'$K{first}="{txt}"'], fill=PatternFill("solid", fgColor=color)))
db.conditional_formatting.add(f"J{first}:J{last}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5"))
db.freeze_panes = "A4"

# ================= Dept Summary (tie-out) =================
ds_ = out.create_sheet("Dept Summary")
ds_["A1"] = "Department Totals: Tie-out to AGM Budget"; ds_["A1"].font = T
ds_["A2"] = "Revenue/Expenditure = sum of the GL Budget sheet. AGM figures (blue) are copied from the AGM file's P&L tab, column L. Diff should be zero."; ds_["A2"].font = S
cols = ["Department (Tab)", "Dept #", "Revenue Budget", "AGM Revenue", "Diff", "Expenditure Budget", "AGM Expenditure", "Diff", "Net (Rev - Exp)", "# GL Lines", "# GL Lines with Budget"]
header(ds_, 4, cols, [26, 9, 15, 15, 10, 15, 15, 10, 15, 10, 12])
for i, (tab, dept, dnum, prr, per) in enumerate(TABS):
    r = 5 + i
    vals = [dept, dnum, f'=SUMIFS({GB("H")},{GB("B")},A{r},{GB("D")},"Revenue")', pl.cell(prr, 12).value or 0, f"=C{r}-D{r}",
            f'=SUMIFS({GB("H")},{GB("B")},A{r},{GB("D")},"Expenditure")', pl.cell(per, 12).value or 0, f"=F{r}-G{r}", f"=C{r}-F{r}",
            f'=COUNTIFS({GB("B")},A{r})', f'=COUNTIFS({GB("B")},A{r},{GB("N")},"Yes")']
    for c, v in enumerate(vals, 1):
        x = ds_.cell(r, c, v); x.font = BI if c in (4, 7) else B
        if 3 <= c <= 9: x.number_format = MONEY
r = 5 + len(TABS); ds_.cell(r, 1, "Total").font = BB
for c in range(3, 12):
    x = ds_.cell(r, c, f"=SUM({L(c)}5:{L(c)}{r - 1})"); x.font = BB; x.number_format = MONEY if c <= 9 else "0"; x.border = Border(top=thin)
ds_.cell(r + 2, 1, "Note: the AGM P&L total leaves out Capital Projects. It is listed here as its own line.").font = S

# ================= Notes =================
nt = out.create_sheet("Notes")
nt.column_dimensions["A"].width = 120
notes = ["How to use the tracker",
"• Expenses come from the 'Expense Ledger' tab of the NUSA Expense Claims log (Google Sheets). The claim system adds a row for every receipt line when VP Finance assigns GL codes and signs, and keeps its Status in sync (Pending > Approved; Sent back = Void).",
"• Your own purchases: add a row in that Google tab (Source = Manual, Date, GL Code from the dropdown, Description, Amount, Status). To correct a claim row, edit it there and set Locked = Yes so the system never overwrites it.",
"• Excel pulls the latest ledger automatically when the file opens and every 30 minutes while it is open (Data > Refresh All does it on demand). If Excel shows a yellow Security Warning bar, click Enable Content. The private link lives on the Settings sheet (keep this file private; if the link leaks, run rotateLedgerFeedKey in Apps Script). Amounts count including tax.",
"• Status: Pending = waiting on approvals. Approved and Posted count as Spent. Void is ignored. A manual row with no Status counts as Approved.",
"• Each GL code is charged to a budget line through the GL Map sheet. If a code sits on several budget lines, the one with the largest budget is the default; put a department in 'Charge to Department' on the ledger row to pick another. 'Unmatched entries' on the Dashboard counts rows that aren't counted anywhere: a GL code that isn't on the map, or a Status other than Pending/Approved/Posted/Void. Fix them in the Google tab.",
"• GL codes the claim system offers but the AGM budget doesn't have are listed on GL Budget as $0 lines (Category '(Not in AGM budget)' or 'Clubs (not in AGM budget)'), so spending on them shows as Unbudgeted spend.",
"• Faculty: every Faculty expense line is split in half, one half per semester (Fall = May-Dec, Spring = Jan-Apr). Unused Fall money does not roll over to Spring. The semester of each expense comes from its Date (the receipt date on claims).",
"• GL Budget has Fall/Spring Budget and Spent columns, plus Period columns: for Faculty the current semester, for everyone else the full year. % Used and Status use the Period numbers. The current semester is on the Dashboard (cell H5) and follows today's date.",
"• Dashboard: totals, a table by department, and every expense GL line ranked by % used (closest to budget first), grouped by department. Pick a department in the yellow cell.",
"• GL Budget Status: On track (<75%), Watch (75-90%), Almost out (90-100%), Over budget (>100%), Unbudgeted spend (money spent on a GL with $0 budget).",
"• Web dashboard: refresh and save this file, then load it on 'NUSA Budget Dashboard.html'. It reads the GL Budget, Line Items and Expense Ledger sheets, so keep those sheet names and column headers the same.",
"",
"How the budget sheets work",
"• Line Items: every budget line from the AGM tabs (Administration to Capital Projects), tied to its GL code. Blue numbers are copied from the AGM file.",
"• GL Budget: one row per GL line per tab. Budget is a formula (sum of that GL's line items). To change a budget, edit the amount on Line Items.",
"• 'Has Budget?' = Yes when the budget is not zero. For one-click filter buttons: click in the table > Table Design > Insert Slicer.",
"• 'Same GL Code: Total Across All Tabs' adds up a GL code wherever it appears in the same section (e.g. Faculty's Proposals Committee lines use 72111 codes, also on the Proposal Committee tab).",
"• Budget numbers come from the annual budget workbook (path in AGM_SRC). They are copied values, not live links.",
"• Capital Projects: a capital-equipment line can be split into the items listed under the table on its department tab.",
"",
"Things to check in the budget workbook",
"• Events: revenue lines and some event lines may have no GL code in the budget. They show as '(No GL in budget)'.",
"• Capital Projects: revenue is the 'Fund Balance at the beginning of the period', the figure the P&L uses (P&L!L18); 'Loan repayment to the university' has GL 'N/A'.",
"• Faculty: lines that share one GL code are totalled together in 'Same GL Code: Total Across All Tabs'.",
"• A tab title may show a different department number than its GL codes; this build follows the GL codes.",
"• A department tab may use another department's GL codes (e.g. the Student Lounge tab using Administration codes). See the 'GL Dept #' column.",
"• Revenue sections repeat the department's expense GL codes with $0 (same as the AGM file). The tracker dropdown only lists expense lines."]
for i, t in enumerate(notes, 1):
    x = nt.cell(i, 1, t); x.font = BB if t and not t.startswith("•") else B; x.alignment = Alignment(wrap_text=True)

out.move_sheet("Lists", offset=len(out.sheetnames))
for ws in out: ws.sheet_view.showGridLines = False
out.active = 0
out.save(OUT)
json.dump(com, open(os.path.join(HERE, "com.json"), "w", encoding="utf-8"))
print("saved", OUT, "| expense labels:", N_LAB)
