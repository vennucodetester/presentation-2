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
  $wsWeekly = $wb.Worksheets.Item("Calc_Weekly")

  foreach ($row in @(65,95)) {
    $rng = $wsCharts.Range($wsCharts.Cells.Item($row,3), $wsCharts.Cells.Item($row,27))
    $rng.NumberFormat = "General"
    for ($col = 3; $col -le 27; $col++) {
      $idx = $col - 2
      $wsCharts.Cells.Item($row, $col).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
    }
  }

  foreach ($row in @(155,386,541,696)) {
    $rng = $wsCharts.Range($wsCharts.Cells.Item($row,2), $wsCharts.Cells.Item($row,26))
    $rng.NumberFormat = "General"
    for ($col = 2; $col -le 26; $col++) {
      $idx = $col - 1
      $wsCharts.Cells.Item($row, $col).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
    }
  }

  $wsWeekly.Range($wsWeekly.Cells.Item(154,3), $wsWeekly.Cells.Item(154,27)).NumberFormat = "General"
  for ($col = 3; $col -le 27; $col++) {
    $idx = $col - 2
    $wsWeekly.Cells.Item(154, $col).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
  }

  try { $xl.Calculation = -4105 } catch {}
  $xl.Calculate()
  $wb.Save()
  $wb.Close($true)
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  $xl.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

