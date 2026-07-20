$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$BackupPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.before-impl3-daily-cleanpoint.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\impl3-daily-cleanpoint-verify"
$LocalPath = Join-Path $env:TEMP "DG-New-master-impl3-daily-working.xlsx"
$LogPath = Join-Path $WorkDir "impl3-daily-script.log"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
Set-Content -Path $LogPath -Value "Impl3 daily run $(Get-Date -Format o)"
if (!(Test-Path $BackupPath)) {
  Copy-Item -LiteralPath $WorkbookPath -Destination $BackupPath -Force
}
Copy-Item -LiteralPath $WorkbookPath -Destination $LocalPath -Force

function Rgb([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }
function Log-Step([string]$message) { Add-Content -Path $LogPath -Value "$(Get-Date -Format o) $message" }
function Set-FormulaChecked($cell, [string]$formula) {
  try { $cell.Formula = $formula } catch {
    throw "Formula write failed at $($cell.Worksheet.Name)!$($cell.Address($false,$false)): $formula :: $($_.Exception.Message)"
  }
  $readBack = [string]$cell.Formula
  if ($readBack -match "\[\d+\]") {
    throw "External workbook reference appeared in $($cell.Address($false,$false)): $readBack"
  }
}
function Ensure-Sheet($wb, [string]$name, $afterSheet, [bool]$visible) {
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
function Remove-Dashboard-Chart($dashboard, [string]$chartName) {
  for ($i = $dashboard.ChartObjects().Count; $i -ge 1; $i--) {
    $co = $dashboard.ChartObjects($i)
    if ($co.Name -eq $chartName) { $co.Delete() }
  }
}
function Export-Chart-Png($chartObj, [string]$pngPath) {
  Remove-Item $pngPath -ErrorAction SilentlyContinue
  $chartObj.Activate()
  Start-Sleep -Milliseconds 400
  $chartObj.Chart.Export($pngPath, "PNG") | Out-Null
  if (!(Test-Path $pngPath) -or ((Get-Item $pngPath).Length -eq 0)) {
    throw "Chart export failed or produced a zero-byte PNG: $pngPath"
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

  Log-Step "Normalize Summary labels and validation"
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

  Log-Step "Add clean-point columns"
  $settings.Range("K3").Value2 = "Clean Point 2"
  $settings.Range("L3").Value2 = "Clean Point 3"
  $settings.Range("K2:L2").Merge() | Out-Null
  $settings.Range("K2").Value2 = "Up to 3 clean points per cause; all drawn as lines on the timeline chart."
  $settings.Range("K2").Font.Italic = $true
  $settings.Range("K3:L3").Font.Bold = $true
  $settings.Range("K3:L80").NumberFormat = "yyyy-mm-dd"
  $settings.Range("K4:L80").Validation.Delete()
  $settings.Range("K4:L80").Validation.Add(4, 1, 1, "1", "2958465")
  $settings.Columns("K:L").ColumnWidth = 14

  Log-Step "Dashboard picker cells"
  $dashboard.Range("A83").Value2 = "Timeline cause"
  $dashboard.Range("B83").Value2 = "False Alarm"
  $dashboard.Range("A84").Value2 = "Date basis"
  $dashboard.Range("B84").Value2 = "Call date"
  $dashboard.Range("A83:B84").Interior.Color = Rgb 242 242 242
  $dashboard.Range("A83:A84").Font.Bold = $true
  $dashboard.Range("B83").Validation.Delete()
  $dashboard.Range("B83").Validation.Add(3, 1, 1, "='Settings'!`$A`$4:`$A`$80")
  $dashboard.Range("B84").Validation.Delete()
  $dashboard.Range("B84").Validation.Add(3, 1, 1, "Call date,Commission date")
  try { $wb.Names.Item("tl_cause").Delete() } catch {}
  try { $wb.Names.Item("tl_basis").Delete() } catch {}
  $wb.Names.Add("tl_cause", "='Dashboard'!`$B`$83") | Out-Null
  $wb.Names.Add("tl_basis", "='Dashboard'!`$B`$84") | Out-Null

  Log-Step "Build Timeline Calc"
  $timeline = Ensure-Sheet $wb "Timeline Calc" $heatmap $false
  $timeline.Cells.Clear()
  $timeline.Range("A1").Value2 = "Daily clean-point timeline source. Pickers live on Dashboard!B83:B84."
  $timeline.Range("A1:N1").Merge() | Out-Null
  $timeline.Range("A2").Value2 = "Cause"
  Set-FormulaChecked $timeline.Range("B2") "=tl_cause"
  $timeline.Range("A3").Value2 = "Basis"
  Set-FormulaChecked $timeline.Range("B3") "=tl_basis"
  Set-FormulaChecked $timeline.Range("B4") "='Dashboard'!B83&"" - issues by ""&'Dashboard'!B84&"", with clean points (""&'Dashboard'!B5&"")"""
  $timeline.Range("A7:H7").Value2 = @("Date", "Issues", "Clean Point 1", "Clean Point 2", "Clean Point 3", "CP Date 1", "CP Date 2", "CP Date 3")
  $timeline.Range("I7:K7").Value2 = @("CP Label 1", "CP Label 2", "CP Label 3")
  $timeline.Range("L7").Value2 = "Ceiling"
  $timeline.Range("A7:L7").Font.Bold = $true
  $timeline.Range("A8").Value2 = [datetime]"2026-03-01"
  for ($r = 9; $r -le 191; $r++) {
    Set-FormulaChecked $timeline.Cells.Item($r, 1) "=A$($r-1)+1"
  }
  $timeline.Range("A8:A191").NumberFormat = "mm/dd"
  for ($r = 8; $r -le 191; $r++) {
    Set-FormulaChecked $timeline.Cells.Item($r, 2) "=IF(tl_basis=""Call date"",COUNTIFS('Summary'!`$E`$2:`$E`$240,tl_cause,'Summary'!`$A`$2:`$A`$240,"">=""&`$A$r,'Summary'!`$A`$2:`$A`$240,""<""&`$A$r+1,'Summary'!`$A`$2:`$A`$240,"">=""&IF('Dashboard'!`$B`$3="""",0,'Dashboard'!`$B`$3),'Summary'!`$A`$2:`$A`$240,""<""&IF('Dashboard'!`$B`$4="""",2958465,'Dashboard'!`$B`$4+1)),COUNTIFS('Summary'!`$E`$2:`$E`$240,tl_cause,'Summary'!`$D`$2:`$D`$240,"">=""&`$A$r,'Summary'!`$D`$2:`$D`$240,""<""&`$A$r+1,'Summary'!`$D`$2:`$D`$240,"">=""&IF('Dashboard'!`$B`$3="""",0,'Dashboard'!`$B`$3),'Summary'!`$D`$2:`$D`$240,""<""&IF('Dashboard'!`$B`$4="""",2958465,'Dashboard'!`$B`$4+1)))"
  }
  Set-FormulaChecked $timeline.Range("F8") '=IFERROR(INDEX(''Settings''!$E$4:$E$80,MATCH(tl_cause,''Settings''!$A$4:$A$80,0)),"")'
  Set-FormulaChecked $timeline.Range("G8") '=IFERROR(INDEX(''Settings''!$K$4:$K$80,MATCH(tl_cause,''Settings''!$A$4:$A$80,0)),"")'
  Set-FormulaChecked $timeline.Range("H8") '=IFERROR(INDEX(''Settings''!$L$4:$L$80,MATCH(tl_cause,''Settings''!$A$4:$A$80,0)),"")'
  Set-FormulaChecked $timeline.Range("I8") '=IF(OR(F8="",F8=0),"","Clean point 1"&CHAR(10)&TEXT(F8,"mmm d"))'
  Set-FormulaChecked $timeline.Range("J8") '=IF(OR(G8="",G8=0),"","Clean point 2"&CHAR(10)&TEXT(G8,"mmm d"))'
  Set-FormulaChecked $timeline.Range("K8") '=IF(OR(H8="",H8=0),"","Clean point 3"&CHAR(10)&TEXT(H8,"mmm d"))'
  Set-FormulaChecked $timeline.Range("L8") '=MAX(1,MAX(B8:B191)+1)'
  $timeline.Range("F8:H8").NumberFormat = "yyyy-mm-dd"
  for ($r = 8; $r -le 191; $r++) {
    Set-FormulaChecked $timeline.Cells.Item($r, 3) "=IF(AND(`$F`$8<>"""",A$r=`$F`$8),`$L`$8,NA())"
    Set-FormulaChecked $timeline.Cells.Item($r, 4) "=IF(AND(`$G`$8<>"""",A$r=`$G`$8),`$L`$8,NA())"
    Set-FormulaChecked $timeline.Cells.Item($r, 5) "=IF(AND(`$H`$8<>"""",A$r=`$H`$8),`$L`$8,NA())"
  }
  $timeline.Columns("A:L").AutoFit() | Out-Null

  Log-Step "Data Checks"
  $checks.Range("A1:E1").Value2 = @("Check", "Count", "Status", "What to fix", "Formula scope")
  $checks.Range("A3").Value2 = "Root causes not in Settings"
  Set-FormulaChecked $checks.Range("B3") '=SUMPRODUCT(--(''Summary''!$E$2:$E$240<>""),--ISNA(MATCH(''Summary''!$E$2:$E$240,''Settings''!$A$4:$A$80,0)))'
  Set-FormulaChecked $checks.Range("C3") '=IF(B3=0,"OK","Review")'
  $checks.Range("A4").Value2 = "Clean-point daily chart reconciliation"
  Set-FormulaChecked $checks.Range("B4") '=SUM(''Timeline Calc''!B8:B191)'
  Set-FormulaChecked $checks.Range("C4") '=IF(B4=IF(tl_basis="Call date",COUNTIFS(''Summary''!$E$2:$E$240,tl_cause,''Summary''!$A$2:$A$240,">="&IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3),''Summary''!$A$2:$A$240,"<"&IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1)),COUNTIFS(''Summary''!$E$2:$E$240,tl_cause,''Summary''!$D$2:$D$240,">="&IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3),''Summary''!$D$2:$D$240,"<"&IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1))),"OK","Review")'
  Set-FormulaChecked $dashboard.Range("B6") '=COUNTIF(''Data Checks''!C:C,"Review")'
  $checks.Columns("A:E").AutoFit() | Out-Null

  Log-Step "Create chart"
  Remove-Dashboard-Chart $dashboard "Chart 10 - Clean Point Lines"
  $chartObj = $dashboard.ChartObjects().Add(609.3, 1680, 560, 300)
  $chartObj.Name = "Chart 10 - Clean Point Lines"
  $chart = $chartObj.Chart
  $chart.ChartType = 51
  while ($chart.SeriesCollection().Count -gt 0) { $chart.SeriesCollection(1).Delete() }
  $chart.HasTitle = $true
  try { $chart.ChartTitle.Formula = "='Timeline Calc'!`$B`$4" } catch { $chart.ChartTitle.Text = $timeline.Range("B4").Text }
  $chart.ChartTitle.Font.Bold = $true
  $chart.ChartTitle.Font.Size = 12
  $chart.HasLegend = $false

  $calls = $chart.SeriesCollection().NewSeries()
  $calls.Name = "Issues"
  $calls.Values = "='Timeline Calc'!`$B`$8:`$B`$191"
  $calls.XValues = "='Timeline Calc'!`$A`$8:`$A`$191"
  $calls.Format.Fill.ForeColor.RGB = Rgb 68 114 196
  $calls.Format.Line.Visible = 0
  $calls.HasDataLabels = $true
  $calls.DataLabels().NumberFormat = "0;;;"
  $calls.DataLabels().Font.Size = 8

  for ($k = 1; $k -le 3; $k++) {
    $colLetter = [char](66 + $k)
    $labelCol = [char](72 + $k)
    $m = $chart.SeriesCollection().NewSeries()
    $m.Name = "='Timeline Calc'!`$$labelCol`$8"
    $m.Values = "='Timeline Calc'!`$$colLetter`$8:`$$colLetter`$191"
    $m.XValues = "='Timeline Calc'!`$A`$8:`$A`$191"
    $m.Format.Fill.ForeColor.RGB = Rgb 192 0 0
    $m.Format.Line.Visible = 0
    $m.HasDataLabels = $false
  }
  $chart.ChartGroups(1).Overlap = 100
  $chart.ChartGroups(1).GapWidth = 30
  try { $chart.Axes(1).CategoryType = 3 } catch {}
  try { $chart.Axes(1).BaseUnit = 0 } catch {}
  try { $chart.Axes(1).MajorUnit = 7 } catch {}
  $chart.Axes(1).TickLabels.NumberFormat = "mm/dd"
  $chart.Axes(1).TickLabels.Orientation = 45
  $chart.Axes(1).TickLabels.Font.Size = 8
  $chart.Axes(2).MinimumScale = 0
  $chart.Axes(2).MajorUnit = 1
  $chart.Axes(2).HasTitle = $true
  $chart.Axes(2).AxisTitle.Text = "Issues that day"
  $chart.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = Rgb 217 217 217

  for ($i = $chart.Shapes.Count; $i -ge 1; $i--) {
    $shape = $chart.Shapes.Item($i)
    if ($shape.Name -like "CleanPointLabel*") { $shape.Delete() }
  }
  $labelPositions = @(
    @{ Cell = "I8"; Left = 240; Top = 70 },
    @{ Cell = "J8"; Left = 330; Top = 70 },
    @{ Cell = "K8"; Left = 420; Top = 70 }
  )
  foreach ($cfg in $labelPositions) {
    $shape = $chart.Shapes.AddTextbox(1, [double]$cfg.Left, [double]$cfg.Top, 70, 34)
    $shape.Name = "CleanPointLabel_$($cfg.Cell)"
    $shape.TextFrame.Characters().Text = ""
    $shape.DrawingObject.Formula = "='Timeline Calc'!`$$($cfg.Cell)"
    $shape.TextFrame.HorizontalAlignment = -4108
    $shape.TextFrame.VerticalAlignment = -4108
    $shape.TextFrame.Characters().Font.Color = Rgb 192 0 0
    $shape.TextFrame.Characters().Font.Bold = $true
    $shape.TextFrame.Characters().Font.Size = 8
    $shape.Line.Visible = 0
    $shape.Fill.Visible = 0
  }

  Log-Step "Acceptance exports"
  $dashboard.Range("B3").Value2 = ""
  $dashboard.Range("B4").Value2 = ""
  $dashboard.Range("B83").Value2 = "False Alarm"
  $dashboard.Range("B84").Value2 = "Call date"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "daily-false-alarm-call-all-data.png")

  $settings.Range("K18").Value2 = [datetime]"2026-06-26"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "daily-false-alarm-two-lines.png")
  $settings.Range("K18").ClearContents()

  $dashboard.Range("B83").Value2 = "Unknown"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "daily-unknown-zero-extra-lines.png")

  $dashboard.Range("B83").Value2 = "False Alarm"
  $dashboard.Range("B84").Value2 = "Commission date"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "daily-false-alarm-commission-date.png")

  $dashboard.Range("B83").Value2 = "Door Torque"
  $dashboard.Range("B84").Value2 = "Call date"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "daily-door-torque-call-date.png")

  $dashboard.Range("B83").Value2 = "False Alarm"
  $dashboard.Range("B84").Value2 = "Call date"
  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-06-30"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "daily-false-alarm-to-jun30.png")

  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-07-15"
  $dashboard.Range("B83").Value2 = "False Alarm"
  $dashboard.Range("B84").Value2 = "Call date"
  $xl.CalculateFull()
  Export-Chart-Png $chartObj (Join-Path $WorkDir "daily-final-window.png")

  $wb.SaveAs($LocalPath, 51)
  $wb.Close($true)
  $wb = $null
  Copy-Item -LiteralPath $LocalPath -Destination $WorkbookPath -Force

  Log-Step "Reopen verify"
  $verify = $xl.Workbooks.Open($WorkbookPath)
  $xl.CalculateFull()
  $vd = $verify.Worksheets.Item("Dashboard")
  $vc = $verify.Worksheets.Item("Data Checks")
  if ([string]$vc.Range("C2").Text -ne "OK" -or [string]$vc.Range("C3").Text -ne "OK" -or [string]$vc.Range("C4").Text -ne "OK") {
    throw "Data Checks not OK after reopen: C2=$($vc.Range("C2").Text), C3=$($vc.Range("C3").Text), C4=$($vc.Range("C4").Text)"
  }
  $errors = @()
  foreach ($ws in @($verify.Worksheets)) {
    $ur = $ws.UsedRange
    foreach ($err in @("#REF!", "#DIV/0!", "#VALUE!", "#NAME?")) {
      $found = $ur.Find($err)
      if ($found) { $errors += "$($ws.Name)!$($found.Address($false,$false))=$($found.Text)" }
    }
  }
  if ($errors.Count -gt 0) { throw "Formula errors found: $($errors -join '; ')" }
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
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
