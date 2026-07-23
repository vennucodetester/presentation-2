"""Option A: wire charts 2 & 4 to the Dashboard From/To filter.

Chart 2 = issues bucketed by store COMMISSION week -> keep the full commission
ladder (cohorts are independent of the call period) but only count calls whose
CALL date falls in the period (same convention every other chart uses).
Chart 4 = issues bucketed by CALL week -> also clamp the week ladder to the
selected period so it stops showing months outside the window.
COM only. Run with the workbook CLOSED.
"""
import sys, os
import win32com.client as win32

def CL(i):
    s = ""
    while i > 0:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

P = sys.argv[1]
RENDER = "--render" in sys.argv
OUT = os.path.dirname(P)

FROM = 'IF(Dashboard!$B$3="",0,Dashboard!$B$3)'
TO   = 'IF(Dashboard!$B$4="",2958465,Dashboard!$B$4+1)'
PERIOD = (f',Summary!$A$2:$A$249,">="&{FROM}'
          f',Summary!$A$2:$A$249,"<"&{TO}')

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False; excel.DisplayAlerts = False
wb = excel.Workbooks.Open(P)
if wb.ReadOnly:
    raise SystemExit("ABORT: workbook opened READ-ONLY (stale ~$ lock or Excel still running) - nothing would save.")
w = wb.Worksheets("Calc_Weekly")

LASTC = 58  # cause columns C..BF

# ---- Chart 2 source: rows 6..25, buckets on COMMISSION date (col D) -------
for r in range(6, 26):
    w.Range(w.Cells(r, 3), w.Cells(r, LASTC)).Formula = [[
        f'=IF(OR({CL(c)}$5="",$A{r}=""),"",COUNTIFS(Summary!$E$2:$E$249,{CL(c)}$5'
        f',Summary!$D$2:$D$249,">="&$A{r},Summary!$D$2:$D$249,"<"&$A{r}+7{PERIOD}))'
        for c in range(3, LASTC + 1)]]
print("chart2 source rows 6-25: period filter added")

# ---- Chart 4 ladder clamped to the period --------------------------------
start = (f'IF(Dashboard!$B$3="",MIN(Summary!$A$2:$A$249),Dashboard!$B$3)')
w.Range("A106").Formula = f'={start}-WEEKDAY({start},2)+1'
w.Range("B106").Formula = '=IF($A106="","",TEXT($A106,"mmm d"))'
end = 'IF(Dashboard!$B$4="",MAX(Summary!$A$2:$A$249),Dashboard!$B$4)'
for r in range(107, 140):
    w.Cells(r, 1).Formula = f'=IF($A{r-1}="","",IF($A{r-1}+7>{end},"",$A{r-1}+7))'
    w.Cells(r, 2).Formula = f'=IF($A{r}="","",TEXT($A{r},"mmm d"))'
print("chart4 week ladder: clamped to period")

# ---- Chart 4 source: rows 106..125, buckets on CALL date (col A) ----------
for r in range(106, 126):
    w.Range(w.Cells(r, 3), w.Cells(r, LASTC)).Formula = [[
        f'=IF(OR({CL(c)}$105="",$A{r}=""),"",COUNTIFS(Summary!$E$2:$E$249,{CL(c)}$105'
        f',Summary!$A$2:$A$249,">="&$A{r},Summary!$A$2:$A$249,"<"&$A{r}+7{PERIOD}))'
        for c in range(3, LASTC + 1)]]
print("chart4 source rows 106-125: period filter added")

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
print("ch2 week1 label/total:", w.Range("B155").Value, w.Range("C155").Value)
print("ch4 ladder:", [w.Cells(r, 2).Value for r in range(106, 126)])

if RENDER:
    d = wb.Worksheets("Dashboard")
    for i in (2, 4):
        p = os.path.join(OUT, f"pf_chart{i}.png")
        if os.path.exists(p): os.remove(p)
        co = d.ChartObjects(i); co.Activate(); co.Chart.Export(p)
        print("rendered", p)

wb.Save(); wb.Close(True); excel.Quit()
print("saved")
