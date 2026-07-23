"""Hide chart 4's unused week slots: size its category + series to the number of
weeks actually inside the selected period."""
import sys, os, re
import win32com.client as win32

P = sys.argv[1]; RENDER = "--render" in sys.argv; OUT = os.path.dirname(P)
BOOK = os.path.basename(P)
pref = f"'{BOOK}'!" if " " in BOOK else f"{BOOK}!"

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(P)
if wb.ReadOnly:
    raise SystemExit("ABORT: opened READ-ONLY - nothing would save.")
d = wb.Worksheets("Dashboard")

# locate chart 4 by its Calc_Weekly row-106 block
ch4 = None
for i in range(1, d.ChartObjects().Count + 1):
    f1 = d.ChartObjects(i).Chart.SeriesCollection(1).Formula
    if "Calc_Weekly!$C$106" in f1 or "Calc_Weekly!$C$106:" in f1:
        ch4 = d.ChartObjects(i).Chart; idx4 = i; break
print("chart4 object index:", idx4, "| series:", ch4.SeriesCollection().Count)
for i in (1, 2, ch4.SeriesCollection().Count):
    print("  s%d:" % i, ch4.SeriesCollection(i).Formula[:110])

CNT = 'MAX(1,COUNT(Calc_Weekly!$A$106:$A$125))'

def addname(n, ref):
    try: wb.Names(n).Delete()
    except Exception: pass
    wb.Names.Add(Name=n, RefersTo=ref)

addname("ch4_cats", f'=OFFSET(Calc_Weekly!$B$106,0,0,{CNT},1)')

pat = re.compile(r"Calc_Weekly!\$([A-Z]+)\$106:\$[A-Z]+\$125")
n_done = 0
for i in range(1, ch4.SeriesCollection().Count + 1):
    s = ch4.SeriesCollection(i)
    f = s.Formula
    m = pat.search(f)
    inner = f[f.index("(") + 1:f.rindex(")")]
    parts = inner.split(",")
    name_arg, order_arg = parts[0], parts[-1]
    if m:
        col = m.group(1)
        nm = f"ch4_s{i:02d}"
        addname(nm, f'=OFFSET(Calc_Weekly!${col}$106,0,0,{CNT},1)')
        val = f"{pref}{nm}"
    else:
        m2 = re.search(r"Calc_Charts!\$([A-Z]+)\$187:\$[A-Z]+\$206", f)
        if m2:
            nm = f"ch4_tot{i:02d}"
            addname(nm, f'=OFFSET(Calc_Charts!${m2.group(1)}$187,0,0,{CNT},1)')
            val = f"{pref}{nm}"
        else:
            val = parts[-2]
    s.Formula = f"=SERIES({name_arg},{pref}ch4_cats,{val},{order_arg})"
    n_done += 1
print("series repointed:", n_done)

excel.CalculateFullRebuild()
if RENDER:
    p = os.path.join(OUT, "pf_chart4.png")
    if os.path.exists(p): os.remove(p)
    co = d.ChartObjects(idx4); co.Activate(); co.Chart.Export(p)
    print("rendered", p)
wb.Save(); wb.Close(True); excel.Quit()
print("saved")
