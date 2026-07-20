$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$BackupPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.before-impl3-cleanpoint.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\impl3-cleanpoint-verify"
$LocalPath = Join-Path $env:TEMP "DG-New-master-impl3-working.xlsx"
$LogPath = Join-Path $WorkDir "impl3-script.log"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
Set-Content -Path $LogPath -Value "Impl3 run $(Get-Date -Format o)"
if (!(Test-Path $BackupPath)) {
  Copy-Item -LiteralPath $WorkbookPath -Destination $BackupPath -Force
}
Copy-Item -LiteralPath $WorkbookPath -Destination $LocalPath -Force

function Release-Com($obj) {
  if ($null -ne $obj) {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
  }
}

function Rgb([int]$r, [int]$g, [int]$b) {
  return $r + ($g * 256) + ($b * 65536)
}

function Set-FormulaChecked($cell, [string]$formula) {
  try {
    $cell.Formula = $formula
  } catch {
    throw "Formula write failed at $($cell.Worksheet.Name)!$($cell.Address($false,$false)): $formula :: $($_.Exception.Message)"
  }
  $readBack = [string]$cell.Formula
  if ($readBack -match "\[\d+\]") {
    throw "External workbook reference appeared in $($cell.Address($false,$false)): $readBack"
  }
}

function Ensure-Sheet($wb, [string]$name, $afterSheet, [bool]$visible = $true) {
  for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
    $ws = $wb.Worksheets.Item($i)
    if ($ws.Name -eq $name) {
      $ws.Visible = $(if ($visible) { -1 } else { 0 })
      return $ws
    }
  }
  $newWs = $wb.Worksheets.Add([System.Type]::Missing, $afterSheet)
  $newWs.Name = $name
  $newWs.Visible = $(if ($visible) { -1 } else { 0 })
  return $newWs
}

function Export-Chart-Png($chartObj, [string]$pngPath) {
  $chartObj.Activate()
  $chartObj.Chart.Export($pngPath, "PNG") | Out-Null
  if (!(Test-Path $pngPath) -or ((Get-Item $pngPath).Length -eq 0)) {
    throw "Chart export failed or produced a zero-byte PNG: $pngPath"
  }
}

function Remove-Dashboard-Chart($dashboard, [string]$chartName) {
  for ($i = $dashboard.ChartObjects().Count; $i -ge 1; $i--) {
    $co = $dashboard.ChartObjects($i)
    if ($co.Name -eq $chartName) { $co.Delete() }
  }
}

function Log-Step([string]$message) {
  Add-Content -Path $LogPath -Value "$(Get-Date -Format o) $message"
}

function Clean-Blank-Marker-Legend($chart, $timeline) {
  if (-not $chart.HasLegend) { return }
  for ($k = 6; $k -ge 1; $k--) {
    $labelText = [string]$timeline.Cells.Item(8, 20 + $k).Text
    if ($labelText.Trim().Length -eq 0) {
      try { $chart.Legend.LegendEntries($k + 1).Delete() } catch {}
    }
  }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $null
$verify = $null

try {
  $wb = $xl.Workbooks.Open($LocalPath)
  try { $xl.Calculation = -4135 } catch {}

  $dashboard = $wb.Worksheets.Item("Dashboard")
  $summary = $wb.Worksheets.Item("Summary")
  $settings = $wb.Worksheets.Item("Settings")
  $heatmap = $wb.Worksheets.Item("Heatmap Calc")
  $checks = $wb.Worksheets.Item("Data Checks")
  $weekly = $wb.Worksheets.Item("Calc_Weekly")

  Log-Step "Normalizing Summary labels and validation"
  # Carry-over cleanup so Summary root-cause validation/checks pass against this master Settings sheet.
  for ($r = 2; $r -le 240; $r++) {
    $v = [string]$summary.Cells.Item($r, 5).Value2
    if ($v -eq "DG Installation issue") { $summary.Cells.Item($r, 5).Value2 = "Installation issue" }
    if ($v -eq "Incorrect Program") { $summary.Cells.Item($r, 5).Value2 = "Wrong program" }
    if ($v -eq "Case End Seal") { $summary.Cells.Item($r, 5).Value2 = "Double defrost" }
  }

  $summary.Range("E2:E240").Validation.Delete()
  $summary.Range("E2:E240").Validation.Add(3, 1, 1, "='Settings'!`$A`$4:`$A`$80")
  $summary.Range("E2:E240").Validation.IgnoreBlank = $true
  $summary.Range("E2:E240").Validation.InCellDropdown = $true
  $summary.Range("E2:E240").Validation.ErrorTitle = "Invalid root cause"
  $summary.Range("E2:E240").Validation.ErrorMessage = "Pick a valid Root Cause from Settings."
  $summary.Range("E2:E240").Validation.ShowError = $true

  Log-Step "Building Interventions"
  $interventions = Ensure-Sheet $wb "Interventions" $heatmap $true
  $interventions.Cells.Clear()
  $interventions.Range("A1").Value2 = "Intervention log - one row per clean point / fix action. Drives the clean-point lines chart. Settings' Clean-Point Date column still drives the after/before charts."
  $interventions.Range("A1:C1").Merge() | Out-Null
  $interventions.Range("A1").Font.Italic = $true
  $interventions.Range("A1").Font.Size = 9
  $interventions.Range("A3:C3").Value2 = @("Root Cause", "Date", "Label")
  $interventions.Range("A3:C3").Font.Bold = $true
  $interventions.Range("A3:C3").Interior.Color = Rgb 217 225 242

  $outRow = 4
  for ($r = 4; $r -le 80; $r++) {
    $cause = [string]$settings.Cells.Item($r, 1).Value2
    $date = $settings.Cells.Item($r, 5).Value2
    if ($cause -and $date) {
      $interventions.Cells.Item($outRow, 1).Value2 = $cause
      $interventions.Cells.Item($outRow, 2).Value2 = [double]$date
      $interventions.Cells.Item($outRow, 3).Value2 = "Clean point"
      $outRow++
    }
  }
  $interventions.Range("A4:A204").Validation.Delete()
  $interventions.Range("A4:A204").Validation.Add(3, 1, 1, "='Settings'!`$A`$4:`$A`$80")
  $interventions.Range("A4:A204").Validation.IgnoreBlank = $true
  $interventions.Range("A4:A204").Validation.InCellDropdown = $true
  $interventions.Range("B4:B204").NumberFormat = "yyyy-mm-dd"
  $interventions.Range("A3:C204").Borders.LineStyle = 1
  $interventions.Range("A3:C204").Borders.Color = Rgb 217 217 217
  $interventions.Columns("A:C").AutoFit() | Out-Null

  Log-Step "Building Timeline Calc"
  $timeline = Ensure-Sheet $wb "Timeline Calc" $interventions $false
  $timeline.Cells.Clear()
  $timeline.Range("A1").Value2 = "Clean-point chart source - formulas feed Dashboard chart. Edit picker in B2; edit clean-point rows on Interventions."
  $timeline.Range("A1:N1").Merge() | Out-Null
  $timeline.Range("A2").Value2 = "Root cause:"
  $timeline.Range("B2").Value2 = "False Alarm"
  $timeline.Range("B2").Validation.Delete()
  $timeline.Range("B2").Validation.Add(3, 1, 1, "='Settings'!`$A`$4:`$A`$80")
  $timeline.Range("B2").Validation.IgnoreBlank = $false
  try { $wb.Names.Item("cp_pick").Delete() } catch {}
  $wb.Names.Add("cp_pick", "='Timeline Calc'!`$B`$2") | Out-Null
  Set-FormulaChecked $timeline.Range("B3") "=B2&"" - weekly calls with clean points (""&'Dashboard'!B5&"")"""

  $timeline.Range("A7:H7").Value2 = @("Week Label", "Calls", "Marker 1", "Marker 2", "Marker 3", "Marker 4", "Marker 5", "Marker 6")
  $timeline.Range("I7:N7").Value2 = @("Date 1", "Date 2", "Date 3", "Date 4", "Date 5", "Date 6")
  $timeline.Range("O7:T7").Value2 = @("Label 1", "Label 2", "Label 3", "Label 4", "Label 5", "Label 6")
  $timeline.Range("U7:Z7").Value2 = @("Chart Label 1", "Chart Label 2", "Chart Label 3", "Chart Label 4", "Chart Label 5", "Chart Label 6")
  $timeline.Range("AA7").Value2 = "Ceiling"
  $timeline.Range("AB7:AG7").Value2 = @("X 1", "X 2", "X 3", "X 4", "X 5", "X 6")
  $timeline.Range("A7:AG7").Font.Bold = $true

  Log-Step "Writing weekly formulas"
  for ($i = 1; $i -le 19; $i++) {
    $r = 7 + $i
    $weekRow = 105 + $i
    Set-FormulaChecked $timeline.Cells.Item($r, 1) "=TEXT(Calc_Weekly!A$weekRow,""mm/dd"")"
    Set-FormulaChecked $timeline.Cells.Item($r, 2) "=COUNTIFS('Summary'!`$E`$2:`$E`$240,cp_pick,'Summary'!`$A`$2:`$A`$240,"">=""&'Calc_Weekly'!`$A`$$weekRow,'Summary'!`$A`$2:`$A`$240,""<""&'Calc_Weekly'!`$A`$$weekRow+7,'Summary'!`$A`$2:`$A`$240,"">=""&IF('Dashboard'!`$B`$3="""",0,'Dashboard'!`$B`$3),'Summary'!`$A`$2:`$A`$240,""<""&IF('Dashboard'!`$B`$4="""",2958465,'Dashboard'!`$B`$4+1))"
    for ($k = 1; $k -le 6; $k++) {
      $dateCol = 8 + $k
      $markerCol = 2 + $k
      Set-FormulaChecked $timeline.Cells.Item($r, $markerCol) "=IF(AND($($timeline.Cells.Item(8,$dateCol).Address($true,$true))<>"""",$($timeline.Cells.Item(8,$dateCol).Address($true,$true))>='Calc_Weekly'!`$A`$$weekRow,$($timeline.Cells.Item(8,$dateCol).Address($true,$true))<'Calc_Weekly'!`$A`$$weekRow+7),`$AA`$8,NA())"
    }
  }
  Set-FormulaChecked $timeline.Range("AA8") '=MAX(1,MAX(B8:B26)*1.15)'
  Log-Step "Writing intervention marker formulas"
  for ($k = 1; $k -le 6; $k++) {
    $dateCell = $timeline.Cells.Item(8, 8 + $k)
    $labelCell = $timeline.Cells.Item(8, 14 + $k)
    $chartLabelCell = $timeline.Cells.Item(8, 20 + $k)
    $xCell = $timeline.Cells.Item(8, 27 + $k)
    $dateExpr = "SMALL(INDEX((Interventions!`$A`$4:`$A`$204=cp_pick)*Interventions!`$B`$4:`$B`$204+(Interventions!`$A`$4:`$A`$204<>cp_pick)*9E+99,0),$k)"
    Set-FormulaChecked $dateCell "=IFERROR(IF($dateExpr>9E+98,"""",$dateExpr),"""")"
    Set-FormulaChecked $labelCell "=IF($($dateCell.Address($true,$true))="""","""",INDEX(Interventions!`$C`$4:`$C`$204,SUMPRODUCT((Interventions!`$A`$4:`$A`$204=cp_pick)*(Interventions!`$B`$4:`$B`$204=$($dateCell.Address($true,$true)))*(ROW(Interventions!`$A`$4:`$A`$204)-ROW(Interventions!`$A`$4)+1))))"
    Set-FormulaChecked $chartLabelCell "=IF($($dateCell.Address($true,$true))="""","""",$($labelCell.Address($true,$true))&CHAR(10)&TEXT($($dateCell.Address($true,$true)),""mmm d""))"
    Set-FormulaChecked $xCell "=IF($($dateCell.Address($true,$true))="""",NA(),MATCH(TEXT($($dateCell.Address($true,$true))-WEEKDAY($($dateCell.Address($true,$true)),2)+1,""mm/dd""),`$A`$8:`$A`$26,0))"
  }
  $timeline.Range("I8:N8").NumberFormat = "yyyy-mm-dd"
  $timeline.Columns("A:AG").AutoFit() | Out-Null

  Log-Step "Writing Data Checks"
  # Data Checks: preserve existing heatmap row and add/update checks for Impl 3.
  $checks.Range("A1:E1").Value2 = @("Check", "Count", "Status", "What to fix", "Formula scope")
  $checks.Range("A3").Value2 = "Root causes not in Settings"
  Set-FormulaChecked $checks.Range("B3") '=SUMPRODUCT(--(''Summary''!$E$2:$E$240<>""),--ISNA(MATCH(''Summary''!$E$2:$E$240,''Settings''!$A$4:$A$80,0)))'
  Set-FormulaChecked $checks.Range("C3") '=IF(B3=0,"OK","Review")'
  $checks.Range("D3").Value2 = "Pick a valid Root Cause from the dropdown."
  $checks.Range("E3").Value2 = "Summary vs Settings"
  $checks.Range("A4").Value2 = "Clean-point chart weekly reconciliation"
  Set-FormulaChecked $checks.Range("B4") '=SUM(''Timeline Calc''!B8:B26)'
  Set-FormulaChecked $checks.Range("C4") '=IF(B4=SUMPRODUCT(--(''Summary''!$E$2:$E$240=cp_pick),--(''Summary''!$A$2:$A$240>=IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3)),--(''Summary''!$A$2:$A$240<IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1))),"OK","Review")'
  $checks.Range("D4").Value2 = "Picked-cause weekly sum must equal Summary lines in the active Dashboard date window."
  $checks.Range("E4").Value2 = "Timeline Calc vs Summary"
  Set-FormulaChecked $dashboard.Range("B6") '=COUNTIF(''Data Checks''!C:C,"Review")'
  $checks.Columns("A:E").AutoFit() | Out-Null

  Log-Step "Creating Dashboard chart"
  Remove-Dashboard-Chart $dashboard "Chart 10 - Clean Point Lines"
  $chartObj = $dashboard.ChartObjects().Add(609.3, 1680, 560, 300)
  $chartObj.Name = "Chart 10 - Clean Point Lines"
  $chart = $chartObj.Chart
  $chart.ChartType = 51
  $chart.HasTitle = $true
  try { $chart.ChartTitle.Formula = "='Timeline Calc'!`$B`$3" } catch { $chart.ChartTitle.Text = $timeline.Range("B3").Text }
  $chart.ChartTitle.Font.Bold = $true
  $chart.ChartTitle.Font.Size = 12
  $chart.HasLegend = $true
  try { $chart.Legend.Position = -4160 } catch {}

  while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() }
  $series = $chart.SeriesCollection().NewSeries()
  $series.Name = "Calls"
  $series.Values = "='Timeline Calc'!`$B`$8:`$B`$26"
  $series.XValues = "='Timeline Calc'!`$A`$8:`$A`$26"
  $series.Format.Fill.ForeColor.RGB = Rgb 68 114 196
  $series.Format.Line.Visible = 0
  $series.HasDataLabels = $true
  $series.DataLabels().NumberFormat = "0;;;"
  $series.DataLabels().Font.Size = 9
  try { $series.DataLabels().Position = 0 } catch {}

  for ($k = 1; $k -le 6; $k++) {
    $marker = $chart.SeriesCollection().NewSeries()
    $marker.Name = "='Timeline Calc'!`$$([char](84+$k))`$8"
    $colLetter = [char](66 + $k)
    $marker.Values = "='Timeline Calc'!`$$colLetter`$8:`$$colLetter`$26"
    $marker.XValues = "='Timeline Calc'!`$A`$8:`$A`$26"
    $marker.Format.Fill.ForeColor.RGB = Rgb 192 0 0
    try { $marker.Format.Fill.Transparency = 0.35 } catch {}
    $marker.Format.Line.Visible = 0
    $marker.HasDataLabels = $false
  }
  $chart.ChartGroups(1).Overlap = 100
  $chart.ChartGroups(1).GapWidth = 45
  $chart.Axes(1).TickLabels.Orientation = 45
  $chart.Axes(1).TickLabels.Font.Size = 9
  $chart.Axes(2).MinimumScale = 0
  $chart.Axes(2).HasTitle = $true
  $chart.Axes(2).AxisTitle.Text = "Calls"
  $chart.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = Rgb 217 217 217

  Log-Step "Built sheets and chart; exporting all-data state"
  $dashboard.Range("B3").Value2 = ""
  $dashboard.Range("B4").Value2 = ""
  $timeline.Range("B2").Value2 = "False Alarm"
  $xl.CalculateFull()
  try { $chart.ChartTitle.Formula = "='Timeline Calc'!`$B`$3" } catch {}
  Export-Chart-Png $chartObj (Join-Path $WorkDir "cleanpoint-false-alarm-all-data.png")

  $interventions.Range("A205").Value2 = "False Alarm"
  $interventions.Range("B205").Value2 = [datetime]"2026-06-26"
  $interventions.Range("C205").Value2 = "Recheck"
  Log-Step "Exporting two-line state"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "cleanpoint-false-alarm-two-lines.png")
  $interventions.Range("A205:C205").ClearContents()
  $xl.CalculateFull()

  $timeline.Range("B2").Value2 = "Door Torque"
  Log-Step "Exporting Door Torque state"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "cleanpoint-door-torque.png")
  $timeline.Range("B2").Value2 = "False Alarm"

  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-06-30"
  Log-Step "Exporting narrowed-window state"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "cleanpoint-to-jun30.png")

  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-07-15"
  Log-Step "Exporting final-window state"
  $xl.CalculateFull()
  Clean-Blank-Marker-Legend $chart $timeline
  Export-Chart-Png $chartObj (Join-Path $WorkDir "cleanpoint-final-window.png")

  $wb.SaveAs($LocalPath, 51)
  $wb.Close($true)
  $wb = $null
  Copy-Item -LiteralPath $LocalPath -Destination $WorkbookPath -Force

  $verify = $xl.Workbooks.Open($WorkbookPath)
  Log-Step "Reopening copied workbook for verification"
  $xl.CalculateFull()
  $vd = $verify.Worksheets.Item("Dashboard")
  $vc = $verify.Worksheets.Item("Data Checks")
  $vt = $verify.Worksheets.Item("Timeline Calc")
  $vt.Range("B2").Value2 = "False Alarm"
  $vd.Range("B3").Value2 = [datetime]"2026-06-01"
  $vd.Range("B4").Value2 = [datetime]"2026-07-15"
  $xl.CalculateFull()
  if ([string]$vc.Range("C2").Text -ne "OK" -or [string]$vc.Range("C3").Text -ne "OK" -or [string]$vc.Range("C4").Text -ne "OK") {
    throw "Data Checks not OK after reopen: C2=$($vc.Range("C2").Text), C3=$($vc.Range("C3").Text), C4=$($vc.Range("C4").Text)"
  }
  $errors = @()
  foreach ($ws in @($verify.Worksheets)) {
    $ur = $ws.UsedRange
    foreach ($err in @("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")) {
      if ($ws.Name -eq "Timeline Calc" -and $err -eq "#N/A") { continue }
      $found = $ur.Find($err)
      if ($found) { $errors += "$($ws.Name)!$($found.Address($false,$false))=$($found.Text)" }
    }
  }
  if ($errors.Count -gt 0) {
    throw "Formula errors found: $($errors -join '; ')"
  }
  $verify.Save()
  $verify.Close($true)
  $verify = $null

  Write-Host "Updated $WorkbookPath"
  Write-Host "Backup $BackupPath"
  Write-Host "Verification previews: $WorkDir"
  Write-Host "Data Checks C2:C4 OK"
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($verify) { try { $verify.Close($false) } catch {} }
  $xl.Quit()
  Release-Com $xl
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
