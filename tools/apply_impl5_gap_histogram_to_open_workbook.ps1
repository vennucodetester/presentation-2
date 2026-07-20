$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$BackupPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.before-impl5-gap-histogram.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\impl5-gap-histogram-verify"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
if (!(Test-Path $BackupPath)) {
  Copy-Item -LiteralPath $WorkbookPath -Destination $BackupPath -Force
}

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
  $cell.Formula = $formula
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

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$oldCalc = $xl.Calculation
$wb = $null

try {
  foreach ($candidate in @($xl.Workbooks)) {
    if ([string]::Equals($candidate.FullName, $WorkbookPath, [System.StringComparison]::OrdinalIgnoreCase)) {
      $wb = $candidate
      break
    }
  }
  if (!$wb) { $wb = $xl.Workbooks.Open($WorkbookPath) }
  $xl.Calculation = -4135

  $summary = $wb.Worksheets.Item("Summary")
  $timeline = $wb.Worksheets.Item("Timeline Calc")
  $dashboard = $wb.Worksheets.Item("Dashboard")
  $checks = $wb.Worksheets.Item("Data Checks")

  $gapCol = $null
  for ($c = 1; $c -le 30; $c++) {
    if ([string]$summary.Cells.Item(1, $c).Text -eq "Gap (days)") { $gapCol = $c; break }
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
  $gapFormula = "=IF(OR(`$A2="""",`$D2="""",NOT(ISNUMBER(`$A2)),NOT(ISNUMBER(`$D2))),"""",INT(`$A2)-INT(`$D2))"
  Set-FormulaChecked $summary.Cells.Item(2, $gapCol) $gapFormula
  $summary.Range("${gapLetter}2:${gapLetter}240").FillDown() | Out-Null
  $summary.Range("${gapLetter}2:${gapLetter}240").NumberFormat = "0"
  $summary.Columns.Item($gapCol).ColumnWidth = 11

  $timeline.Range("X1:AD247").Clear()
  $timeline.Range("X1:AD1").Merge() | Out-Null
  $timeline.Range("X1").Value2 = "Gap histogram source - shares tl_cause"
  $timeline.Range("X1").Font.Bold = $true
  $timeline.Range("X1").Interior.Color = 14277081
  $timeline.Range("X4").Value2 = "Bucket"
  $timeline.Range("Y4").Value2 = "Lo"
  $timeline.Range("Z4").Value2 = "Hi"
  $timeline.Range("AA4").Value2 = "Count"
  $bucketLabels = @("0-7 days", "8-14 days", "15-30 days", "31-60 days", "61-90 days", "90+ days")
  $bucketLo = @(0, 8, 15, 31, 61, 91)
  $bucketHi = @(7, 14, 30, 60, 90, 100000)
  for ($i = 0; $i -lt 6; $i++) {
    $row = 5 + $i
    $timeline.Cells.Item($row, 24).Value2 = $bucketLabels[$i]
    $timeline.Cells.Item($row, 25).Value2 = [double]$bucketLo[$i]
    $timeline.Cells.Item($row, 26).Value2 = [double]$bucketHi[$i]
  }
  $timeline.Range("X4:AA10").Borders.LineStyle = 1
  $timeline.Range("X4:AA4").Font.Bold = $true
  $timeline.Range("X4:AA4").Interior.Color = 14277081

  Set-FormulaChecked $timeline.Range("AA5") "=IF(ISNUMBER(MATCH(tl_cause,Settings!`$B`$4:`$B`$83,0)),COUNTIFS(Summary!`$K`$2:`$K`$240,tl_cause,Summary!`$$gapLetter`$2:`$$gapLetter`$240,"">=""&Y5,Summary!`$$gapLetter`$2:`$$gapLetter`$240,""<=""&Z5,Summary!`$A`$2:`$A`$240,"">=""&IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A`$2:`$A`$240,""<""&IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),COUNTIFS(Summary!`$E`$2:`$E`$240,tl_cause,Summary!`$$gapLetter`$2:`$$gapLetter`$240,"">=""&Y5,Summary!`$$gapLetter`$2:`$$gapLetter`$240,""<=""&Z5,Summary!`$A`$2:`$A`$240,"">=""&IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A`$2:`$A`$240,""<""&IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)))"
  $timeline.Range("AA5:AA10").FillDown() | Out-Null

  $timeline.Range("AB7").Value2 = "In-scope gap"
  $timeline.Range("AB7").Font.Bold = $true
  $timeline.Range("AB7").Interior.Color = 14277081
  Set-FormulaChecked $timeline.Range("AB8") "=IF(ISNUMBER(MATCH(tl_cause,Settings!`$B`$4:`$B`$83,0)),IF(AND(Summary!`$K2=tl_cause,Summary!`$$gapLetter`2<>"""",Summary!`$A2>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A2<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),Summary!`$$gapLetter`2,""""),IF(AND(Summary!`$E2=tl_cause,Summary!`$$gapLetter`2<>"""",Summary!`$A2>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A2<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),Summary!`$$gapLetter`2,""""))"
  $timeline.Range("AB8:AB247").FillDown() | Out-Null

  $timeline.Range("AC4").Value2 = "Metric"
  $timeline.Range("AD4").Value2 = "Value"
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
  Set-FormulaChecked $checks.Range("C6") "=IF(B6=IF(ISNUMBER(MATCH(tl_cause,Settings!`$B`$4:`$B`$83,0)),SUMPRODUCT(--(Summary!`$$gapLetter`$2:`$$gapLetter`$240<>""""),--(Summary!`$A`$2:`$A`$240>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3)),--(Summary!`$A`$2:`$A`$240<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),--(Summary!`$K`$2:`$K`$240=tl_cause)),SUMPRODUCT(--(Summary!`$$gapLetter`$2:`$$gapLetter`$240<>""""),--(Summary!`$A`$2:`$A`$240>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3)),--(Summary!`$A`$2:`$A`$240<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)),--(Summary!`$E`$2:`$E`$240=tl_cause))),""OK"",""Review"")"
  $checks.Range("D6").Value2 = "Gap histogram n must equal direct in-scope gap-bearing Summary rows."
  $checks.Range("E6").Value2 = "Timeline Calc vs Summary"
  $checks.Range("A1:E6").Columns.AutoFit() | Out-Null

  foreach ($co in @($dashboard.ChartObjects())) {
    if ($co.Name -eq "Chart 11 - Gap Histogram") { $co.Delete() }
  }
  $newChart = $dashboard.ChartObjects().Add(293, 2160, 560, 300)
  $newChart.Name = "Chart 11 - Gap Histogram"
  $chart = $newChart.Chart
  $chart.ChartType = 51
  $chart.HasLegend = $false
  $series = $chart.SeriesCollection().NewSeries()
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
  try { $chart.ChartTitle.Formula = "='Timeline Calc'!`$AD`$7" } catch { $chart.ChartTitle.Text = [string]$timeline.Range("AD7").Text }
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
  if ($overlaps.Count -gt 0) { throw "Chart 11 overlaps: $($overlaps -join ', ')" }

  $originalCause = [string]$wb.Names.Item("tl_cause").RefersToRange.Value2
  $originalBasis = [string]$wb.Names.Item("tl_basis").RefersToRange.Value2
  $dashboard.Range("B3").ClearContents()
  $dashboard.Range("B4").ClearContents()
  $wb.Names.Item("tl_cause").RefersToRange.Value2 = "Doors"
  $xl.Calculation = -4105
  $xl.CalculateFull()
  $doorsCounts = @()
  for ($r = 5; $r -le 10; $r++) { $doorsCounts += [int]$timeline.Cells.Item($r, 27).Value2 }
  $doorsN = [int]$timeline.Range("AD5").Value2
  $doorsMedian = [int]$timeline.Range("AD6").Value2
  Export-Chart-Png $newChart (Join-Path $WorkDir "doors-all-data-gap-histogram.png")
  if (($doorsCounts -join "/") -ne "6/1/8/10/11/5" -or $doorsN -ne 41 -or $doorsMedian -ne 42) {
    Write-Host "WARNING: Current workbook Doors all-data differs from spec target. Current=$($doorsCounts -join '/') n=$doorsN median=$($doorsMedian)d; target=6/1/8/10/11/5 n=41 median=42d."
  }

  $hasInstallCause = $false
  for ($r = 8; $r -le 167; $r++) {
    if ([string]$timeline.Cells.Item($r, 22).Text -eq "DG Installation issue") { $hasInstallCause = $true; break }
  }
  if ($hasInstallCause) {
    $wb.Names.Item("tl_cause").RefersToRange.Value2 = "DG Installation issue"
    $xl.CalculateFull()
    Export-Chart-Png $newChart (Join-Path $WorkDir "dg-installation-gap-histogram.png")
    Export-Chart-Png $dashboard.ChartObjects("Chart 10 - Clean Point Lines") (Join-Path $WorkDir "dg-installation-cleanpoint-lines.png")
  } else {
    Write-Host "WARNING: Picker value 'DG Installation issue' not found; skipped picker export."
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

  if ([string]$checks.Range("C6").Text -ne "OK") {
    throw "Gap histogram Data Check is not OK: $($checks.Range("C6").Text)"
  }
  $errors = @()
  foreach ($ws in @($wb.Worksheets)) {
    $ur = $ws.UsedRange
    foreach ($err in @("#REF!", "#DIV/0!", "#VALUE!", "#NAME?")) {
      $found = $ur.Find($err)
      if ($found) { $errors += "$($ws.Name)!$($found.Address($false,$false))=$($found.Text)" }
    }
  }
  if ($errors.Count -gt 0) { throw "Formula errors found: $($errors -join '; ')" }

  $xl.Calculation = -4105
  $wb.Save()
  Write-Host "Updated open workbook $WorkbookPath"
  Write-Host "Backup $BackupPath"
  Write-Host "Gap column Summary!$gapLetter"
  Write-Host "Doors all-data buckets=$($doorsCounts -join '/') n=$doorsN median=$($doorsMedian)d"
  Write-Host "Window test Jun30 n=$windowN"
  Write-Host "Verification previews: $WorkDir"
}
finally {
  try { $xl.Calculation = -4105 } catch {}
  try { $xl.EnableEvents = $true } catch {}
}
