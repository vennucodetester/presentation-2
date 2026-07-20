$ErrorActionPreference = "Stop"
$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\outputs\019f6dfb-339c-7212-a7a5-679970584164\DG-template-reorganized-phase5.xlsx"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
try {
  $wb = $xl.Workbooks.Open($WorkbookPath)
  try { $xl.Calculation = -4135 } catch {}
  $wsCharts = $wb.Worksheets.Item("Calc_Charts")
  $wsChecks = $wb.Worksheets.Item("Data Checks")

  for ($r = 213; $r -le 217; $r++) {
    $idx = $r - 212
    $wsCharts.Cells.Item($r, 1).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
    $wsCharts.Cells.Item($r, 2).Formula = "=IF(A$r="""","""",SUMIF(Summary!`$L`$5:`$L`$241,A$r,Summary!`$A`$5:`$A`$241))"
    $wsCharts.Cells.Item($r, 3).Formula = "=IF(A$r="""","""",SUMPRODUCT(SUBTOTAL(103,OFFSET(Summary!`$E`$5,ROW(Summary!`$E`$5:`$E`$241)-ROW(Summary!`$E`$5),0,1)),--(Summary!`$L`$5:`$L`$241=A$r)))"
    $wsCharts.Cells.Item($r, 4).Formula = "=IF(A$r="""","""",COUNTIFS(Summary!`$L`$5:`$L`$241,A$r,Summary!`$G`$5:`$G`$241,""Yes""))"
  }

  $wsCase = $wb.Worksheets.Item("Case Nomenclature Graph")
  for ($r = 2; $r -le 9; $r++) {
    $wsCase.Cells.Item($r, 2).Formula = "=SUMPRODUCT(--ISNUMBER(SEARCH(A$r,Summary!`$H`$5:`$H`$241)),--ISNUMBER(Summary!`$A`$5:`$A`$241),--(Summary!`$A`$5:`$A`$241>=IF(set_from="""",0,set_from)),--(Summary!`$A`$5:`$A`$241<IF(set_to="""",2958465,set_to+1)))"
  }
  $wsCase.Range("B10").Formula = "=set_from"
  $wsCase.Range("B11").Formula = "=set_to"

  $wsChecks.Range("B16").Formula = "=SUMPRODUCT(--(Calc_Charts!`$C`$65:`$AA`$65<>Calc_Charts!`$C`$95:`$AA`$95))+SUMPRODUCT(--(Calc_Charts!`$B`$155:`$Z`$155<>Calc_Charts!`$B`$386:`$Z`$386))+SUMPRODUCT(--(Calc_Charts!`$B`$155:`$Z`$155<>Calc_Charts!`$B`$541:`$Z`$541))+SUMPRODUCT(--(Calc_Charts!`$B`$155:`$Z`$155<>Calc_Charts!`$B`$696:`$Z`$696))+SUMPRODUCT(--(Calc_Charts!`$C`$65:`$AA`$65<>Calc_Charts!`$B`$155:`$Z`$155))"

  try { $xl.Calculation = -4105 } catch {}
  try { $xl.CalculateFullRebuild() } catch { $xl.CalculateFull() }
  $wb.Save()
  $wb.Close($true)
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  $xl.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
