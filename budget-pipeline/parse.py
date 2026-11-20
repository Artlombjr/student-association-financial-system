import openpyxl, re, collections
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule

import os
# Input: the annual budget workbook (one tab per department plus a P&L tab). Output: the Excel tracker workbook.
SRC = os.environ.get("AGM_SRC", "budget.xlsx")  # set AGM_SRC to read a copy when the workbook is open in Excel
OUT = os.environ.get("TRACKER_OUT", "GL Budget & Expense Tracker.xlsx")  # where the Excel tracker is built

wb = openpyxl.load_workbook(SRC, data_only=True)

# tab -> (department name, dept #, P&L revenue row, P&L expense row)
TABS = [("Administration2","Administration","72103",7,23),("Faculty2","Faculty","72104",8,24),
        ("Communications2","Communications","72105",9,25),("Events2","Events","72106",10,26),
        ("Courier2","The Courier","72107",11,27),("Yearbook2","Yearbook","72108",12,28),
        ("StudentLoungeThe COMMONS","Student Lounge (The Commons)","72109",13,29),
        ("Special Initiatives","Special Initiatives","72110",14,30),
        ("Proposal Committee","Proposal Committee","72111",15,31),("Capital Projects","Capital Projects","73001",18,34)]
# Some workbooks name the department tabs with a trailing 2 (Administration2); use whichever name the file has.
TABS = [(x[0][:-1] if x[0].endswith("2") and x[0][:-1] in wb.sheetnames else x[0],) + x[1:] for x in TABS]
pl = wb["P&L"]

# GL codes are typed "410  7210341035100003300610000" or "410-72103-41035-10000-33006-10000";
# both become the first form, which the rest of the build uses as the GL key.
GLRE = re.compile(r"^\s*4\d\d[\s\xa0-]+\d")
def norm_gl(v):
    if v is None: return None
    s = str(v).replace("\xa0"," ").strip()
    if not GLRE.match(s): return s
    d = re.sub(r"\D", "", s)
    return d[:3] + "  " + d[3:] if len(d) == 28 else re.sub(r"\s+","  ",re.sub(r"^(4\d\d)[\s-]+",r"\1  ",s))

parents, items, flags = [], [], []
summary = []   # fund-balance rows inside a section ; used only when the P&L counts one
SUMMARY_RE = re.compile(r"^(Fund Balance|Excess|Retained Earning)", re.I)  # "Retained Earnings" rows too
STATUS_RE = re.compile(r"^(Ra|N/?R)$", re.I)   # Faculty2 col A marks each club Ra (ratified) or N/R (not ratified)
for tab, dept, dnum, prr, per in TABS:
    ws = wb[tab]
    hdr = next(r for r in range(1,30) if any(c.value=="Annual Budget" for c in ws[r]))
    bcol = next(c.column for c in ws[hdr] if c.value=="Annual Budget")
    acol = next(c.column for c in ws[hdr] if c.value=="Year Ended")
    cnt = collections.Counter(c.column for row in ws.iter_rows() for c in row if isinstance(c.value,str) and GLRE.match(c.value))
    gcol = cnt.most_common(1)[0][0]
    section = category = group = None; parent = None; after = False; detail = None; bcols = {}
    for r in range(hdr+1, ws.max_row+1):
        a = ws.cell(r,1).value
        if after:
            # Breakdowns listed below the table: a row naming an expense GL line, then "item | amount" rows
            # (e.g. Capital Projects: CapEquip/General (Communications) = Flash for Yearbook + hard drives + computers).
            # heading "Breakdown for <GL line>", a Description / Department / Qty header row, the
            # amount in the budget column, then Est. Subtotal / Est. Taxes / Est. Additional Fees / Est. Total rows.
            isnum = lambda v: isinstance(v,(int,float)) and not isinstance(v,bool)
            labc = next((c for c in range(1,bcol) if isinstance(ws.cell(r,c).value,str) and ws.cell(r,c).value.strip()), None)
            lab = ws.cell(r,labc).value.strip() if labc else None
            num = bcol if isnum(ws.cell(r,bcol).value) else next((c for c in range(5, ws.max_column+1) if isnum(ws.cell(r,c).value)), None)
            if lab == "Description":
                bcols = {str(ws.cell(r,c).value).strip(): c for c in range(1, ws.max_column+1) if isinstance(ws.cell(r,c).value,str)}
            elif lab and num is None:
                name = re.sub(r"^Breakdown for\s+", "", lab, flags=re.I)
                detail = next((p for p in parents if p["tab"]==tab and p["section"]=="Expenditure" and p["desc"]==name and not p["kids"]), None)
                bcols = {}
                if detail: detail.setdefault("below", [])
            elif lab and detail and re.search(r"total$", lab, re.I): pass
            elif lab and detail:
                q = ws.cell(r,bcols["Qty"]).value if "Qty" in bcols else None
                dp = ws.cell(r,bcols["Department"]).value if "Department" in bcols else None
                dp = str(dp).strip().replace("Courier", "The Courier") if dp and str(dp).strip() != "TBD" else ""
                desc = lab + (f" x{q:g}" if isnum(q) and q > 1 else "") + (f" ({dp})" if dp else "")
                detail["below"].append(dict(desc=desc,group=None,budget=ws.cell(r,num).value,actual=0,src=f"'{tab}'!{L(num)}{r}",r=r))
            elif not lab: detail = None
            continue
        if isinstance(a,str) and a.strip() in ("Revenue","Expenditures"):
            section = "Revenue" if a.strip()=="Revenue" else "Expenditure"; category=group=parent=None; continue
        if isinstance(a,str) and a.startswith("Excess"): after = True; continue
        lab, lvl = None, None
        for c in range(1,5):
            v = ws.cell(r,c).value
            if c == 1 and isinstance(v,str) and STATUS_RE.match(v.strip()): continue
            if isinstance(v,str) and v.strip(): lab, lvl = v.strip(), c; break
        if not lab or section is None: continue
        gl = norm_gl(ws.cell(r,gcol).value)
        # Capital Projects ends its table with "Retained Earning at the End of Period" in column B
        if section == "Expenditure" and not gl and re.match(r"^(Excess|Retained Earning)", lab, re.I): after = True; continue
        bud = ws.cell(r,bcol).value; act = ws.cell(r,acol).value
        src = f"'{tab}'!{L(bcol)}{r}"
        if not gl and SUMMARY_RE.match(lab):
            summary.append(dict(tab=tab,dept=dept,dnum=dnum,section=section,category="(No category)",desc=lab,
                                gl="(No GL in budget)",actual=act or 0,budget=bud or 0,lvl=lvl,src=src,kids=[],r=r))
            continue
        if gl or (bud is not None and (parent is None or lvl <= parent["lvl"])):
            parent = dict(tab=tab,dept=dept,dnum=dnum,section=section,category=category or "(No category)",
                          desc=lab,gl=gl or "(No GL in budget)",actual=act or 0,budget=bud or 0,lvl=lvl,src=src,kids=[],r=r)
            parents.append(parent); group=None
        elif bud is None:
            if lvl <= 2: category = lab; parent=None; group=None
            else: group = lab
        else:
            parent["kids"].append(dict(desc=lab,group=group,budget=bud or 0,actual=act or 0,src=src,r=r))

# use a breakdown listed below a table only when it adds up to the GL line's budget
for p in parents:
    below = p.pop("below", None)
    if not below: continue
    s = sum(k["budget"] for k in below)
    if abs(s - p["budget"]) <= 0.005: p["kids"] = below
    else: flags.append(f"{p['tab']} {p['desc']}: breakdown below the table adds to {s}, GL row says {p['budget']} (breakdown ignored)")

# Tie each department to the P&L (column L). A sub-item added under a GL row whose SUM() wasn't widened is missed by the GL row
# but counted by the tab total : when trusting the
# sub-items makes the department tie, use them.
kid_sum = lambda p: sum(k["budget"] for k in p["kids"])
for tab, dept, dnum, prr, per in TABS:
    for sec, plr in (("Revenue", prr), ("Expenditure", per)):
        ps = [p for p in parents if p["dept"] == dept and p["section"] == sec]
        target, cur = pl.cell(plr, 12).value or 0, sum(p["budget"] for p in ps)
        off = [p for p in ps if p["kids"] and abs(kid_sum(p) - p["budget"]) > 0.005]
        if abs(cur - target) > 0.005 and off and abs(cur + sum(kid_sum(p) - p["budget"] for p in off) - target) <= 0.005:
            for p in off:
                flags.append(f"{tab} {p['desc']}: GL row {p['budget']} left out sub-items; using sub-items {kid_sum(p)} (ties to P&L)")
                p["budget"] = kid_sum(p)
# A department that still doesn't tie: use the fund-balance row the P&L counts .
for tab, dept, dnum, prr, per in TABS:
    for sec, plr in (("Revenue", prr), ("Expenditure", per)):
        gap = (pl.cell(plr, 12).value or 0) - sum(p["budget"] for p in parents if p["dept"] == dept and p["section"] == sec)
        row = next((x for x in summary if x["dept"] == dept and x["section"] == sec and abs(gap) > 0.005 and abs(x["budget"] - gap) <= 0.005), None)
        if row:
            flags.append(f"{tab} {row['desc']}: {row['budget']} counted as {sec.lower()} (matches the P&L)")
            parents.append(row)
# The P&L takes Administration's student-fee revenue from 'Historical Numbers'!K25 instead of the Administration2 tab
# ; follow the P&L.
adm_rev = [p for p in parents if p["dept"] == "Administration" and p["section"] == "Revenue"]
diff = (pl.cell(7, 12).value or 0) - sum(p["budget"] for p in adm_rev)
fees = next((p for p in adm_rev if p["desc"].startswith("Student Association Fees")), None)
if abs(diff) > 0.005 and fees and not fees["kids"]:
    flags.append(f"{TABS[0][0]} {fees['desc']}: tab says {fees['budget']}, P&L says {fees['budget'] + diff}; using the P&L")
    fees["budget"] += diff; fees["src"] = "P&L!L7"

# rows the tab's own total formula leaves out
EXCLUDED = set()  # (tab, row) pairs a tab's own total leaves out;
for p in parents:
    for k in (p["kids"] or [p]):
        items.append(dict(p=p, desc=k["desc"], group=k.get("group") if k is not p else None,
                          budget=k["budget"], actual=k["actual"], src=k["src"],
                          incl="No" if (p["tab"],k["r"]) in EXCLUDED else "Yes"))
    if p["kids"]:
        s = sum(k["budget"] for k in p["kids"])
        if abs(s - p["budget"])>0.005:
            flags.append(f"{p['tab']} {p['desc']}: GL row {p['budget']} vs sub-items {s}")
            items.append(dict(p=p, desc="Adjustment to match AGM total", group=None, budget=p["budget"]-s, actual=0,
                              src=p["src"]+" minus sub-items", incl="Yes"))

# departments whose expense lines are split in half: one half per semester (Fall = May-Dec, Spring = Jan-Apr)
SPLIT_DEPTS = {"Faculty"}
split = []
for it in items:
    p = it["p"]
    if p["dept"] in SPLIT_DEPTS and p["section"] == "Expenditure":
        for sem in ("Fall", "Spring"):
            split.append({**it, "budget": it["budget"] / 2, "actual": it["actual"] / 2, "sem": sem})
    else:
        split.append({**it, "sem": "Full year"})
items = split
for p in parents:
    p["split"] = p["dept"] in SPLIT_DEPTS and p["section"] == "Expenditure"
print("mismatches:", flags)
print(len(parents),"GL rows,",len(items),"line items")
