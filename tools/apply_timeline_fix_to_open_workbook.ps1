$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$FixedLocalPath = Join-Path $env:TEMP "DG-New-master-timeline-group-basis-fix.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\timeline-group-basis-fix-verify"

if (!(Test-Path $FixedLocalPath)) {
  throw "Fixed local workbook does not exist: $FixedLocalPath"
}

function Export-Chart-Png($chartObj, [string]$pngPath) {
  Remove-Item $pngPath -ErrorAction SilentlyContinue
  $chartObj.Activate()
  Start-Sleep -Milliseconds 500
  $chartObj.Chart.Export($pngPath, "PNG") | Out-Null
  if (!(Test-Path $pngPath) -or ((Get-Item $pngPath).Length -eq 0)) {
    throw "Chart export failed or produced a zero-byte PNG: $pngPath"
  }
}

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false

$target = $null
foreach ($candidate in @($xl.Workbooks)) {
  if ([string]::Equals($candidate.FullName, $WorkbookPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    $target = $candidate
    break
  }
}
if (!$target) {
  $target = $xl.Workbooks.Open($WorkbookPath)
}

$source = $null
try {
  $source = $xl.Workbooks.Open($FixedLocalPath, $null, $true)

  $srcTimeline = $source.Worksheets.Item("Timeline Calc")
  $srcChecks = $source.Worksheets.Item("Data Checks")
  $srcDashboard = $source.Worksheets.Item("Dashboard")
  $dstTimeline = $target.Worksheets.Item("Timeline Calc")
  $dstChecks = $target.Worksheets.Item("Data Checks")
  $dstDashboard = $target.Worksheets.Item("Dashboard")

  $dstTimeline.Range("A8:E191").Formula = $srcTimeline.Range("A8:E191").Formula
  $dstChecks.Range("C4").Formula = $srcChecks.Range("C4").Formula
  $dstChecks.Range("D4").Value2 = $srcChecks.Range("D4").Value2

  $dstChart = $dstDashboard.ChartObjects("Chart 10 - Clean Point Lines")
  $srcChart = $srcDashboard.ChartObjects("Chart 10 - Clean Point Lines")
  for ($i = 1; $i -le $srcChart.Chart.SeriesCollection().Count; $i++) {
    $dstChart.Chart.SeriesCollection($i).Formula = $srcChart.Chart.SeriesCollection($i).Formula
  }

  $dstDashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dstDashboard.Range("B4").Value2 = [datetime]"2026-07-15"
  $dstDashboard.Range("B125").Value2 = "Doors"
  $dstDashboard.Range("B126").Value2 = "Call date"
  $xl.CalculateFull()
  $doorsCall = [int]$xl.WorksheetFunction.Sum($dstTimeline.Range("B8:B191"))
  $expected = [int]$xl.Evaluate("SUMPRODUCT(--(Summary!K2:K240=""Doors""),--(Summary!A2:A240>=DATE(2026,6,1)),--(Summary!A2:A240<DATE(2026,7,16)))")
  Export-Chart-Png $dstChart (Join-Path $WorkDir "doors-call-date-open-workbook.png")

  $dstDashboard.Range("B126").Value2 = "Commission date"
  $xl.CalculateFull()
  $doorsCommission = [int]$xl.WorksheetFunction.Sum($dstTimeline.Range("B8:B191"))
  Export-Chart-Png $dstChart (Join-Path $WorkDir "doors-commission-date-open-workbook.png")

  if ($doorsCall -ne $expected -or $doorsCommission -ne $expected) {
    throw "Open workbook Doors totals failed: call=$doorsCall commission=$doorsCommission expected=$expected"
  }
  if ([string]$dstChecks.Range("C4").Text -ne "OK") {
    throw "Timeline Data Check is not OK in open workbook: $($dstChecks.Range("C4").Text)"
  }

  $dstDashboard.Range("B125").Value2 = "False Alarm"
  $dstDashboard.Range("B126").Value2 = "Call date"
  $dstDashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dstDashboard.Range("B4").Value2 = [datetime]"2026-07-15"
  $xl.CalculateFull()

  $target.Save()
  Write-Host "Updated open workbook $WorkbookPath"
  Write-Host "Doors expected=$expected call=$doorsCall commission=$doorsCommission"
  Write-Host "Verification previews: $WorkDir"
}
finally {
  if ($source) { try { $source.Close($false) } catch {} }
  try { $xl.EnableEvents = $true } catch {}
}
