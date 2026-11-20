import openpyxl, json, os
HERE = os.path.dirname(os.path.abspath(__file__))
P = os.environ.get("TRACKER_XLSX") or os.path.join(os.path.dirname(HERE), "GL Budget & Expense Tracker.xlsx")
wb = openpyxl.load_workbook(P)

def rows(ws, hdr=4):
    heads = [c.value for c in ws[hdr]]
    for r in ws.iter_rows(min_row=hdr + 1, values_only=True):
        yield dict(zip(heads, r))

gl = []
for d in rows(wb["GL Budget"]):
    if not isinstance(d["GL Key"], int): continue
    gl.append({"k": d["GL Key"], "dept": d["Department (Tab)"], "dnum": d["Dept #"], "sec": d["Section"], "cat": d["Category"],
               "desc": d["GL Description"], "gl": d["GL Code"], "prev": d["2025-26 Actual"] or 0,
               "label": d["Tracker Label (used in Transactions)"], "split": d["Split by Semester?"] == "Yes"})
li = []
for d in rows(wb["Line Items"]):
    if not isinstance(d["GL Key"], int): continue
    li.append({"k": d["GL Key"], "name": d["Line Item"], "grp": d["Sub-group"] or "", "b": d["2026-27 Budget"] or 0,
               "inc": d["Counted in AGM Total?"] == "Yes", "sem": d["Semester"] or "Full year"})
import datetime
data = {"asOf": datetime.date.today().isoformat(), "source": "Budget snapshot (no expenses logged yet)", "gl": gl, "li": li, "tx": []}
tpl = open(os.path.join(HERE, "dashboard_template.html"), encoding="utf-8").read()
html = tpl.replace("/*__DATA__*/null", json.dumps(data, separators=(",", ":")))
open(os.path.join(os.path.dirname(HERE), "NUSA Budget Dashboard.html"), "w", encoding="utf-8").write(html)
print(len(gl), "gl rows", len(li), "line items", len(html), "bytes")
