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

  for ($r = 3; $r -le 58; $r++) {
    $settingsRow = $r + 1
    $wsCharts.Cells.Item($r, 34).Formula = "=IF(Settings!A$settingsRow="""",999,IF(COUNTIF(Settings!`$B`$4:Settings!B$settingsRow,Settings!B$settingsRow)=1,ROW(Settings!B$settingsRow),999))"
  }
  for ($i = 1; $i -le 25; $i++) {
    $r = 2 + $i
    $wsCharts.Cells.Item($r, 33).Formula = "=IFERROR(IF(SMALL(`$AH`$3:`$AH`$58,$i)>900,"""",INDEX(Settings!`$B`$4:`$B`$59,SMALL(`$AH`$3:`$AH`$58,$i)-3)),"""")"
  }
  foreach ($n in @($wb.Names)) {
    if ($n.Name -eq "set_groupList") { $n.Delete() }
  }
  $wb.Names.Add("set_groupList", "=OFFSET(Calc_Charts!`$AG`$3,0,0,MAX(1,COUNTIF(Calc_Charts!`$AG`$3:`$AG`$27,""?*"")),1)") | Out-Null

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

