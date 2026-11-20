# Exports the budget data from the Excel tracker into "TrackerData.gs" (for the Google Sheets tracker builder).
# Run after build2.py/apply.ps1 whenever the AGM budget changes, then paste the new file into Apps Script and run
# rebuildBudgetTracker. Reads a copy so it works even while the Excel file is open.
import os, json, shutil, tempfile, datetime
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.environ.get("TRACKER_XLSX", os.path.join(os.path.dirname(HERE), "GL Budget & Expense Tracker.xlsx"))
OUT = os.environ.get("TRACKER_DATA_OUT", "TrackerData.gs")

tmp = os.path.join(tempfile.gettempdir(), "tracker_export_copy.xlsx")
shutil.copyfile(SRC, tmp)
wb = openpyxl.load_workbook(tmp)

def rows(name, ncols):
    ws = wb[name]
    for r in ws.iter_rows(min_row=5, max_col=ncols, values_only=True):
        yield r

clean = lambda s: (s or "").replace("'", "") if isinstance(s, str) else s   # 'Administration2'!M19 -> Administration2!M19
gl, li, mp, agm = [], [], [], []
for r in rows("GL Budget", 22):
    if not isinstance(r[0], int): continue
    # key, dept, dnum, section, category, description, GL code, 2025-26 actual, GL dept #, account #, label, source, split
    gl.append([r[0], r[1], str(r[2] or ""), r[3], r[4], r[5], r[6], r[14] or 0, str(r[16] or ""), str(r[17] or ""), r[19], clean(r[20]), 1 if r[21] == "Yes" else 0])
for r in rows("Line Items", 15):
    if not isinstance(r[0], int): continue
    # line id, GL key, sub-group, line item, 2025-26 actual, 2026-27 budget, counted in AGM total, source, semester
    li.append([r[0], r[13], r[7] or "", r[8], r[9] or 0, r[10] or 0, 1 if r[11] == "Yes" else 0, clean(r[12]), r[14]])
for r in rows("GL Map", 5):
    if not r[0]: continue
    # key, budget line, GL code, department, why
    mp.append([r[0], r[3], r[1], r[2] or "", r[4] or ""])
for r in rows("Dept Summary", 7):
    if not r[0] or r[0] == "Total" or r[1] is None: continue
    agm.append([r[0], str(r[1]), r[3] or 0, r[6] or 0])

depts = [a[0] for a in agm]
data = {"built": datetime.date.today().isoformat(), "depts": depts, "gl": gl, "li": li, "map": mp, "agm": agm}
js = ("/**\n * NUSA budget tracker: budget data (generated from the budget workbook by gsheet_export.py on "
      + data["built"] + ").\n * Don't edit by hand; change budgets on the Line Items tab instead.\n */\n"
      "const TRACKER_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n")
open(OUT, "w", encoding="utf-8").write(js)
os.remove(tmp)
print(f"{len(gl)} GL lines, {len(li)} line items, {len(mp)} map rows, {len(agm)} departments -> {OUT} ({len(js)//1024} KB)")
