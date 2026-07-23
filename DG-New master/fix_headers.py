"""Make the group-matrix header rows formula-driven from Settings (single source
of truth), so renames/additions flow to every chart automatically.

Adds a dedup Group List in Settings!N:O, repoints 6 Calc_Charts header rows and
Calc_Weekly!154, and gives chart 2 its missing 21st group series.
COM only (openpyxl destroys charts). Run with workbook CLOSED.
"""
import win32com.client as win32
from win32com.client import constants as C

P = r"C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
SET_LAST = 60          # Settings scan depth
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(P)
st = wb.Worksheets("Settings")
cc = wb.Worksheets("Calc_Charts")
cwk = wb.Worksheets("Calc_Weekly")

# ---- A. Dedup group list in Settings!N (marker) / O (list) ----------------
st.Cells(3, 14).Value = "(grp first-seen)"
st.Cells(3, 15).Value = "Group List (auto)"
st.Range(st.Cells(4, 14), st.Cells(SET_LAST, 14)).Formula = [
    [f'=IF($A{r}="","",IF(COUNTIF($B$4:$B{r},$B{r})=1,ROW(),""))'] for r in range(4, SET_LAST + 1)]
st.Range(st.Cells(4, 15), st.Cells(SET_LAST, 15)).Formula = [
    [f'=IFERROR(INDEX($B$4:$B${SET_LAST},SMALL($N$4:$N${SET_LAST},ROW()-3)-3),"")'] for r in range(4, SET_LAST + 1)]
excel.CalculateFullRebuild()
groups = [st.Cells(r, 15).Value for r in range(4, SET_LAST + 1)]
groups = [g for g in groups if g not in (None, "")]
print("dedup groups:", len(groups), groups[:3], "...", groups[-2:])

def idx_formula(k):
    return f'=IFERROR(INDEX(Settings!$O$4:$O${SET_LAST},{k}),"")'

# ---- B. Header rows -> formulas -----------------------------------------
# (sheet, header_row, first_col, n_slots)
BLOCKS = [
    (cc, 65,  3, 25),   # chart1  C65:AA65
    (cc, 95,  3, 25),   # chart3  C95:AA95
    (cc, 155, 2, 25),   # chart5  B155:Z155
    (cc, 386, 2, 25),   # chart7
    (cc, 541, 2, 25),   # chart8
    (cc, 696, 2, 25),   # chart9
    (cwk, 154, 3, 21),  # chart2  C154:W154 (21st slot = W, newly used)
]
for sh, hr, c0, n in BLOCKS:
    rng = sh.Range(sh.Cells(hr, c0), sh.Cells(hr, c0 + n - 1))
    # these cells were typed text and carry NumberFormat "@" -> Excel would store
    # a formula as a literal string. Reset to General before writing.
    rng.NumberFormat = "General"
    rng.Formula = [[idx_formula(k + 1) for k in range(n)]]
    print(f"{sh.Name}!row{hr}: {n} header slots -> formulas")

# ---- C. Calc_Weekly column W body (21st group) ---------------------------
cwk.Range(cwk.Cells(155, 23), cwk.Cells(174, 23)).Formula = [
    [f'=IF(W$154="","",SUMIF($C$4:$BF$4,W$154,$C{r-149}:$BF{r-149}))'] for r in range(155, 175)]

excel.CalculateFullRebuild()

# ---- D. Chart 2: add the missing 21st group series, before the total carrier
d = wb.Worksheets("Dashboard")
ch2 = None
for i in range(1, d.ChartObjects().Count + 1):
    ttl = d.ChartObjects(i).Chart.SeriesCollection(1).Formula
    if "Calc_Weekly!$C$154" in ttl:
        ch2 = d.ChartObjects(i).Chart; break
print("chart2 found:", ch2 is not None, "| series before:", ch2.SeriesCollection().Count)

names = [ch2.SeriesCollection(i).Formula for i in range(1, ch2.SeriesCollection().Count + 1)]
if not any("$W$154" in f for f in names):
    ns = ch2.SeriesCollection().NewSeries()
    ns.Formula = "=SERIES(Calc_Weekly!$W$154,,Calc_Weekly!$W$155:$W$174,%d)" % (ch2.SeriesCollection().Count)
    # colour from Settings group colour for that group
    g21 = groups[20] if len(groups) > 20 else None
    hexcol = None
    for r in range(4, SET_LAST + 1):
        if st.Cells(r, 2).Value == g21:
            hexcol = st.Cells(r, 12).Value; break
    if hexcol:
        h = str(hexcol).lstrip("#")
        ns.Format.Fill.Visible = True
        ns.Format.Fill.ForeColor.RGB = int(h[4:6] + h[2:4] + h[0:2], 16)  # BGR
    # new group segment must stack with the other groups, i.e. BEFORE the
    # week-total label carrier -> give it plot order 21 (total slides to last)
    try:
        ns.PlotOrder = 21
    except Exception as e:
        print("  PlotOrder note:", str(e)[:60])
    print("added series for", g21, "colour", hexcol, "| series now:", ch2.SeriesCollection().Count)

excel.CalculateFullRebuild()
errs = 0
for ws in wb.Worksheets:
    v = ws.UsedRange.Value
    if isinstance(v, tuple):
        for row in v:
            for cell in (row if isinstance(row, tuple) else (row,)):
                if isinstance(cell, str) and cell in {"#REF!","#DIV/0!","#VALUE!","#NAME?","#NULL!","#NUM!","#N/A"}:
                    errs += 1
print("FORMULA ERRORS:", errs)
wb.Save(); wb.Close(True); excel.Quit()
print("saved")
