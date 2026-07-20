$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$BackupPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.before-impl5-gap-histogram.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\impl5-gap-histogram-verify"
$LocalPath = Join-Path $env:TEMP "DG-New-master-impl5-gap-histogram.xlsx"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
if (!(Test-Path $BackupPath)) {
  Copy-Item -LiteralPath $WorkbookPath -Destination $BackupPath -Force
}
Copy-Item -LiteralPath $WorkbookPath -Destination $LocalPath -Force

function Get-ColumnLetter([int]$columnNumber) {
  $n = $columnNumber
  $name = ""
  while ($n -gt 0) {
    $mod = ($n - 1) % 26
    $name = [char](65 + $mod) + $name
    $n = [math]::Floor(($n - $mod) / 26)
  }
  return $name
}

function Set-FormulaChecked($cell, [string]$formula) {
  try { $cell.Formula = $formula } catch {
    throw "Formula write failed at $($cell.Worksheet.Name)!$($cell.Address($false,$false)): $formula :: $($_.Exception.Message)"
  }
  $readBack = [string]$cell.Formula
  if ($readBack -match "\[\d+\]") {
    throw "External workbook reference appeared in $($cell.Worksheet.Name)!$($cell.Address($false,$false)): $readBack"
  }
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

function Rects-Overlap($a, $b) {
  return !(
    ($a.Left + $a.Width) -le $b.Left -or
    ($b.Left + $b.Width) -le $a.Left -or
    ($a.Top + $a.Height) -le $b.Top -or
    ($b.Top + $b.Height) -le $a.Top
  )
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $null
$verify = $null

try {
  $wb = $xl.Workbooks.Open($LocalPath)
  try { $xl.Calculation = -4105 } catch {}

  $summary = $wb.Worksheets.Item("Summary")
  $timeline = $wb.Worksheets.Item("Timeline Calc")
  $dashboard = $wb.Worksheets.Item("Dashboard")
  $checks = $wb.Worksheets.Item("Data Checks")

  $gapCol = $null
  for ($c = 1; $c -le 30; $c++) {
    if ([string]$summary.Cells.Item(1, $c).Text -eq "Gap (days)") {
      $gapCol = $c
      break
    }
  }
  if (!$gapCol) {
    $lastHeader = 1
    for ($c = 1; $c -le 30; $c++) {
      if ([string]$summary.Cells.Item(1, $c).Text -ne "") { $lastHeader = $c }
    }
    $gapCol = $lastHeader + 1
  }
  $gapLetter = Get-ColumnLetter $gapCol

  $summary.Cells.Item(1, $gapCol).Value2 = "Gap (days)"
  $summary.Cells.Item(1, $gapCol).Font.Bold = $true
  $summary.Cells.Item(1, $gapCol).Interior.Color = 14277081
  for ($r = 2; $r -le 240; $r++) {
    Set-FormulaChecked $summary.Cells.Item($r, $gapCol) "=IF(OR(`$A$r="""",`$D$r="""",NOT(ISNUMBER(`$A$r)),NOT(ISNUMBER(`$D$r))),"""",`$A$r-`$D$r)"
  }
  $summary.Range("${gapLetter}2:${gapLetter}240").NumberFormat = "0"
  $summary.Columns.Item($gapCol).ColumnWidth = 11

  $timeline.Range("X1:AD247").Clear()
  $timeline.Range("X1:AD1").Merge() | Out-Null
  $timeline.Range("X1").Value2 = "Gap histogram source - shares tl_cause"
  $timeline.Range("X1").Font.Bold = $true
  $timeline.Range("X1").Interior.Color = 14277081

  $timeline.Range("X4:AA4").Value2 = @("Bucket", "Lo", "Hi", "Count")
  $timeline.Range("X5:X10").Value2 = @(
    @("0-7 days"),
    @("8-14 days"),
    @("15-30 days"),
    @("31-60 days"),
    @("61-90 days"),
    @("90+ days")
  )
  $timeline.Range("Y5:Y10").Value2 = @(@(0), @(8), @(15), @(31), @(61), @(91))
  $timeline.Range("Z5:Z10").Value2 = @(@(7), @(14), @(30), @(60), @(90), @(100000))
  $timeline.Range("X4:AA10").Borders.LineStyle = 1
  $timeline.Range("X4:AA4").Font.Bold = $true
  $timeline.Range("X4:AA4").Interior.Color = 14277081

  for ($r = 5; $r -le 10; $r++) {
    $formula = "=IF(ISNUMBER(MATCH(tl_cause,Settings!`$B`$4:`$B`$83,0)),COUNTIFS(Summary!`$K`$2:`$K`$240,tl_cause,Summary!`$$gapLetter`$2:`$$gapLetter`$240,"">=""&Y$r,Summary!`$$gapLetter`$2:`$$gapLetter`$240,""<=""&Z$r,Summary!`$A`$2:`$A`$240,"">=""&IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A`$2:`$A`$240,""<""&IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),COUNTIFS(Summary!`$E`$2:`$E`$240,tl_cause,Summary!`$$gapLetter`$2:`$$gapLetter`$240,"">=""&Y$r,Summary!`$$gapLetter`$2:`$$gapLetter`$240,""<=""&Z$r,Summary!`$A`$2:`$A`$240,"">=""&IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A`$2:`$A`$240,""<""&IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)))"
    Set-FormulaChecked $timeline.Cells.Item($r, 27) $formula
  }

  $timeline.Range("AB7").Value2 = "In-scope gap"
  $timeline.Range("AB7").Font.Bold = $true
  $timeline.Range("AB7").Interior.Color = 14277081
  for ($r = 8; $r -le 247; $r++) {
    $sr = $r - 6
    $formula = "=IF(ISNUMBER(MATCH(tl_cause,Settings!`$B`$4:`$B`$83,0)),IF(AND(Summary!`$K$sr=tl_cause,Summary!`$$gapLetter$sr<>"""",Summary!`$A$sr>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A$sr<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),Summary!`$$gapLetter$sr,""""),IF(AND(Summary!`$E$sr=tl_cause,Summary!`$$gapLetter$sr<>"""",Summary!`$A$sr>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A$sr<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),Summary!`$$gapLetter$sr,""""))"
    Set-FormulaChecked $timeline.Cells.Item($r, 28) $formula
  }

  $timeline.Range("AC4:AD4").Value2 = @("Metric", "Value")
  $timeline.Range("AC5").Value2 = "n"
  Set-FormulaChecked $timeline.Range("AD5") "=SUM(AA5:AA10)"
  $timeline.Range("AC6").Value2 = "median"
  Set-FormulaChecked $timeline.Range("AD6") "=IFERROR(MEDIAN(AB8:AB247),"""")"
  $timeline.Range("AC7").Value2 = "title"
  Set-FormulaChecked $timeline.Range("AD7") "=tl_cause&"" - days from commission to service call (n=""&AD5&"", median ""&AD6&""d)"""
  $timeline.Range("AC4:AD7").Borders.LineStyle = 1
  $timeline.Range("AC4:AD4").Font.Bold = $true
  $timeline.Range("AC4:AD4").Interior.Color = 14277081
  $timeline.Range("AA5:AA10,AD5:AD6").NumberFormat = "0"
  $timeline.Range("X:AD").Columns.AutoFit() | Out-Null

  $checks.Range("A6").Value2 = "Gap histogram reconciliation"
  Set-FormulaChecked $checks.Range("B6") "='Timeline Calc'!`$AD`$5"
  $directFormula = "=IF(B6=IF(ISNUMBER(MATCH(tl_cause,Settings!`$B`$4:`$B`$83,0)),SUMPRODUCT(--(Summary!`$$gapLetter`$2:`$$gapLetter`$240<>""""),--(Summary!`$A`$2:`$A`$240>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3)),--(Summary!`$A`$2:`$A`$240<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),--(Summary!`$K`$2:`$K`$240=tl_cause)),SUMPRODUCT(--(Summary!`$$gapLetter`$2:`$$gapLetter`$240<>""""),--(Summary!`$A`$2:`$A`$240>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3)),--(Summary!`$A`$2:`$A`$240<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),--(Summary!`$E`$2:`$E`$240=tl_cause))),""OK"",""Review"")"
  Set-FormulaChecked $checks.Range("C6") $directFormula
  $checks.Range("D6").Value2 = "Gap histogram n must equal direct in-scope gap-bearing Summary rows."
  $checks.Range("E6").Value2 = "Timeline Calc vs Summary"
  $checks.Range("A1:E6").Columns.AutoFit() | Out-Null

  foreach ($co in @($dashboard.ChartObjects())) {
    if ($co.Name -eq "Chart 11 - Gap Histogram") { $co.Delete() }
  }

  $left = 293.0
  $top = 2160.0
  $width = 560.0
  $height = 300.0
  $newChart = $dashboard.ChartObjects().Add($left, $top, $width, $height)
  $newChart.Name = "Chart 11 - Gap Histogram"
  $chart = $newChart.Chart
  $chart.ChartType = 51
  $chart.HasLegend = $false
  $chart.SetSourceData($timeline.Range("X4:AA10"))
  while ($chart.SeriesCollection().Count -gt 1) { $chart.SeriesCollection(2).Delete() }
  $series = $chart.SeriesCollection(1)
  $series.Name = "=""Issues"""
  $series.XValues = "='Timeline Calc'!`$X`$5:`$X`$10"
  $series.Values = "='Timeline Calc'!`$AA`$5:`$AA`$10"
  $series.Format.Fill.ForeColor.RGB = 12874308
  $series.Format.Line.Visible = 0
  $series.HasDataLabels = $true
  $series.DataLabels().NumberFormat = "0;;;"
  $series.DataLabels().Font.Size = 9
  $chart.ChartGroups(1).GapWidth = 40
  $chart.HasTitle = $true
  try {
    $chart.ChartTitle.Formula = "='Timeline Calc'!`$AD`$7"
  } catch {
    $chart.ChartTitle.Text = [string]$timeline.Range("AD7").Text
  }
  $chart.ChartTitle.Font.Bold = $true
  $chart.ChartTitle.Font.Size = 12
  $chart.Axes(1).HasTitle = $true
  $chart.Axes(1).AxisTitle.Text = "Days from commission to call"
  $chart.Axes(2).HasTitle = $true
  $chart.Axes(2).AxisTitle.Text = "Issue count"
  $chart.Axes(2).MajorUnit = 1
  $chart.Axes(2).MinimumScale = 0
  $chart.Axes(2).TickLabels.NumberFormat = "0"
  $chart.ChartArea.Format.Line.Visible = 0
  $chart.PlotArea.Format.Line.Visible = 0

  $newRect = @{ Left = $newChart.Left; Top = $newChart.Top; Width = $newChart.Width; Height = $newChart.Height }
  $overlaps = @()
  foreach ($co in @($dashboard.ChartObjects())) {
    if ($co.Name -ne $newChart.Name) {
      $rect = @{ Left = $co.Left; Top = $co.Top; Width = $co.Width; Height = $co.Height }
      if (Rects-Overlap $newRect $rect) { $overlaps += $co.Name }
    }
  }
  foreach ($shape in @($dashboard.Shapes)) {
    if ($shape.Name -ne $newChart.Name -and $shape.Type -ne 3) {
      $rect = @{ Left = $shape.Left; Top = $shape.Top; Width = $shape.Width; Height = $shape.Height }
      if (Rects-Overlap $newRect $rect) { $overlaps += $shape.Name }
    }
  }
  if ($overlaps.Count -gt 0) {
    throw "Chart 11 overlaps: $($overlaps -join ', ')"
  }

  $originalCause = [string]$wb.Names.Item("tl_cause").RefersToRange.Value2
  $originalBasis = [string]$wb.Names.Item("tl_basis").RefersToRange.Value2
  $originalFrom = $dashboard.Range("B3").Value2
  $originalTo = $dashboard.Range("B4").Value2

  $dashboard.Range("B3").ClearContents()
  $dashboard.Range("B4").ClearContents()
  $wb.Names.Item("tl_cause").RefersToRange.Value2 = "Doors"
  $xl.CalculateFull()
  $doorsCounts = @()
  for ($r = 5; $r -le 10; $r++) { $doorsCounts += [int]$timeline.Cells.Item($r, 27).Value2 }
  $doorsN = [int]$timeline.Range("AD5").Value2
  $doorsMedian = [int]$timeline.Range("AD6").Value2
  Export-Chart-Png $newChart (Join-Path $WorkDir "doors-all-data-gap-histogram.png")

  $knownCounts = @(6, 1, 8, 10, 11, 5)
  $knownMatch = ($doorsN -eq 41 -and $doorsMedian -eq 42)
  for ($i = 0; $i -lt 6; $i++) {
    if ($doorsCounts[$i] -ne $knownCounts[$i]) { $knownMatch = $false }
  }
  if (!$knownMatch) {
    Write-Host "WARNING: Current workbook Doors all-data differs from spec target. Current=$($doorsCounts -join '/') n=$doorsN median=$($doorsMedian)d; target=6/1/8/10/11/5 n=41 median=42d."
  }

  $installCause = "DG Installation issue"
  $hasInstallCause = $false
  for ($r = 8; $r -le 167; $r++) {
    if ([string]$timeline.Cells.Item($r, 22).Text -eq $installCause) { $hasInstallCause = $true; break }
  }
  if ($hasInstallCause) {
    $wb.Names.Item("tl_cause").RefersToRange.Value2 = $installCause
    $xl.CalculateFull()
    Export-Chart-Png $newChart (Join-Path $WorkDir "dg-installation-gap-histogram.png")
    Export-Chart-Png $dashboard.ChartObjects("Chart 10 - Clean Point Lines") (Join-Path $WorkDir "dg-installation-cleanpoint-lines.png")
  } else {
    Write-Host "WARNING: Picker value '$installCause' not found in current Timeline Calc pick list; skipped picker export."
  }

  $wb.Names.Item("tl_cause").RefersToRange.Value2 = "Doors"
  $dashboard.Range("B3").ClearContents()
  $dashboard.Range("B4").Value2 = [datetime]"2026-06-30"
  $xl.CalculateFull()
  $windowN = [int]$timeline.Range("AD5").Value2
  if ($windowN -ge $doorsN) { throw "Window test failed: Jun 30 n=$windowN did not drop below all-data n=$doorsN" }

  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-07-15"
  if ($originalCause -ne "") { $wb.Names.Item("tl_cause").RefersToRange.Value2 = $originalCause }
  if ($originalBasis -ne "") { $wb.Names.Item("tl_basis").RefersToRange.Value2 = $originalBasis }
  $xl.CalculateFull()

  $errors = @()
  foreach ($ws in @($wb.Worksheets)) {
    $ur = $ws.UsedRange
    foreach ($err in @("#REF!", "#DIV/0!", "#VALUE!", "#NAME?")) {
      $found = $ur.Find($err)
      if ($found) { $errors += "$($ws.Name)!$($found.Address($false,$false))=$($found.Text)" }
    }
  }
  if ($errors.Count -gt 0) { throw "Formula errors found: $($errors -join '; ')" }

  try { $xl.Calculation = -4105 } catch {}
  $wb.SaveAs($LocalPath, 51)
  $wb.Close($true)
  $wb = $null
  Copy-Item -LiteralPath $LocalPath -Destination $WorkbookPath -Force

  $verify = $xl.Workbooks.Open($WorkbookPath)
  try { $xl.Calculation = -4105 } catch {}
  $vt = $verify.Worksheets.Item("Timeline Calc")
  $vd = $verify.Worksheets.Item("Dashboard")
  $vc = $verify.Worksheets.Item("Data Checks")
  $vd.Range("B3").ClearContents()
  $vd.Range("B4").ClearContents()
  $verify.Names.Item("tl_cause").RefersToRange.Value2 = "Doors"
  $xl.CalculateFull()
  $reopenCounts = @()
  for ($r = 5; $r -le 10; $r++) { $reopenCounts += [int]$vt.Cells.Item($r, 27).Value2 }
  $reopenN = [int]$vt.Range("AD5").Value2
  $reopenMedian = [int]$vt.Range("AD6").Value2
  if ([string]$vc.Range("C6").Text -ne "OK") {
    throw "Gap histogram Data Check is not OK after reopen: $($vc.Range("C6").Text)"
  }
  if ($reopenCounts -join "/" -ne $doorsCounts -join "/" -or $reopenN -ne $doorsN -or $reopenMedian -ne $doorsMedian) {
    throw "Reopen values changed: before=$($doorsCounts -join '/') n=$doorsN median=$doorsMedian; after=$($reopenCounts -join '/') n=$reopenN median=$reopenMedian"
  }
  $vd.Range("B3").Value2 = [datetime]"2026-06-01"
  $vd.Range("B4").Value2 = [datetime]"2026-07-15"
  if ($originalCause -ne "") { $verify.Names.Item("tl_cause").RefersToRange.Value2 = $originalCause }
  if ($originalBasis -ne "") { $verify.Names.Item("tl_basis").RefersToRange.Value2 = $originalBasis }
  $xl.CalculateFull()
  $verify.Save()
  $verify.Close($true)
  $verify = $null

  Write-Host "Updated $WorkbookPath"
  Write-Host "Backup $BackupPath"
  Write-Host "Gap column Summary!$gapLetter"
  Write-Host "Doors all-data buckets=$($doorsCounts -join '/') n=$doorsN median=$($doorsMedian)d"
  Write-Host "Window test Jun30 n=$windowN"
  Write-Host "Verification previews: $WorkDir"
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($verify) { try { $verify.Close($false) } catch {} }
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
