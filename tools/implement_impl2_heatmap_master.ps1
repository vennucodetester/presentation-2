$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$BackupPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.before-impl2-heatmap.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\impl2-heatmap-verify"
$LocalPath = Join-Path $env:TEMP "DG-New-master-impl2-working.xlsx"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
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
  $cell.Formula = $formula
  $readBack = [string]$cell.Formula
  if ($readBack -match "\[\d+\]") {
    throw "External workbook reference appeared in $($cell.Address($false,$false)): $readBack"
  }
}

function Normalize-Picture-Formula([string]$formula) {
  $trimmed = $formula.Trim()
  if ($trimmed.StartsWith("=")) { return $trimmed.Substring(1) }
  return $trimmed
}

function Ensure-Sheet($wb, [string]$name, $afterSheet) {
  for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
    $ws = $wb.Worksheets.Item($i)
    if ($ws.Name -eq $name) { return $ws }
  }
  $newWs = $wb.Worksheets.Add([System.Type]::Missing, $afterSheet)
  $newWs.Name = $name
  return $newWs
}

function Export-Range-Png($ws, [string]$address, [string]$pngPath) {
  $range = $ws.Range($address)
  $range.CopyPicture(1, 2)
  Start-Sleep -Milliseconds 400
  $chartObj = $ws.ChartObjects().Add([double]$range.Left, [double]$range.Top, [double]$range.Width, [double]$range.Height)
  $chartObj.Activate()
  $chartObj.Chart.Paste()
  $chartObj.Chart.Export($pngPath, "PNG") | Out-Null
  $chartObj.Delete()
  if (!(Test-Path $pngPath) -or ((Get-Item $pngPath).Length -eq 0)) {
    throw "Range export failed or produced a zero-byte PNG: $pngPath"
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
  try { $xl.Calculation = -4105 } catch {}

  $dashboard = $wb.Worksheets.Item("Dashboard")
  $summary = $wb.Worksheets.Item("Summary")
  $settings = $wb.Worksheets.Item("Settings")
  $weekly = $wb.Worksheets.Item("Calc_Weekly")
  $checks = Ensure-Sheet $wb "Data Checks" $dashboard
  $heat = Ensure-Sheet $wb "Heatmap Calc" $checks

  $checks.Visible = -1
  $checks.Cells.Clear()
  $checks.Range("A1:E1").Value2 = @("Check", "Count", "Status", "What to fix", "Formula scope")
  $checks.Range("A1:E1").Font.Bold = $true
  $checks.Range("A1:E1").Interior.Color = Rgb 217 225 242
  $checks.Range("A2").Value2 = "Heatmap classified line reconciliation"
  Set-FormulaChecked $checks.Range("B2") "=SUM('Heatmap Calc'!`$B`$5:`$T`$30)"
  Set-FormulaChecked $checks.Range("C2") "=IF(B2=SUMPRODUCT(--(Summary!`$E`$2:`$E`$240<>""""),--ISNUMBER(MATCH(Summary!`$E`$2:`$E`$240,Settings!`$A`$4:`$A`$29,0)),--(Summary!`$A`$2:`$A`$240>=IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3)),--(Summary!`$A`$2:`$A`$240<IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1))),""OK"",""Review"")"
  $checks.Range("D2").Value2 = "Heatmap sum must equal classified Summary lines in the active Dashboard date window."
  $checks.Range("E2").Value2 = "Heatmap Calc vs Summary"
  $checks.Columns("A:E").AutoFit() | Out-Null
  Set-FormulaChecked $dashboard.Range("B6") '=COUNTIF(''Data Checks''!C:C,"Review")'

  $heat.Visible = -1
  $heat.Tab.Color = Rgb 128 128 128
  $heat.Cells.Clear()
  $heat.Cells.FormatConditions.Delete()
  $heat.Activate()
  $heat.Application.ActiveWindow.DisplayGridlines = $false

  $heat.Range("A1").Value2 = "Cause x week heatmap - source range for the linked picture on Dashboard. Format HERE, not on the picture."
  $heat.Range("A1:T1").Merge() | Out-Null
  $heat.Range("A1").Font.Size = 9
  $heat.Range("A1").Font.Italic = $true
  $heat.Range("A1").Font.Color = Rgb 90 90 90

  Set-FormulaChecked $heat.Range("A3") '="Issues by root cause x call week  ("&Dashboard!$B$5&")"'
  $heat.Range("A3:T3").Merge() | Out-Null
  $heat.Range("A3").HorizontalAlignment = -4108
  $heat.Range("A3").Font.Bold = $true
  $heat.Range("A3").Font.Size = 12
  $heat.Range("A3").Font.Name = "Calibri"

  for ($c = 2; $c -le 20; $c++) {
    $sourceRow = 104 + $c
    Set-FormulaChecked $heat.Cells.Item(4, $c) "=TEXT(Calc_Weekly!A$sourceRow,""mm/dd"")"
  }
  $heat.Range("B4:T4").Font.Size = 9
  $heat.Range("B4:T4").HorizontalAlignment = -4108
  $heat.Range("B4:T4").Orientation = 45

  for ($r = 5; $r -le 30; $r++) {
    $settingsRow = $r - 1
    Set-FormulaChecked $heat.Cells.Item($r, 1) "=IF(Settings!A$settingsRow="""","""",Settings!A$settingsRow)"
  }
  $heat.Range("A5:A30").Font.Size = 9.5
  $heat.Range("A5:A30").HorizontalAlignment = -4152

  for ($r = 5; $r -le 30; $r++) {
    for ($c = 2; $c -le 20; $c++) {
      $weekRow = 104 + $c
      $formula = "=IF(`$A$r="""","""",COUNTIFS(Summary!`$E`$2:`$E`$240,`$A$r,Summary!`$A`$2:`$A`$240,"">=""&Calc_Weekly!`$A`$$weekRow,Summary!`$A`$2:`$A`$240,""<""&Calc_Weekly!`$A`$$weekRow+7,Summary!`$A`$2:`$A`$240,"">=""&IF(Dashboard!`$B`$3="""",0,Dashboard!`$B`$3),Summary!`$A`$2:`$A`$240,""<""&IF(Dashboard!`$B`$4="""",2958465,Dashboard!`$B`$4+1)))"
      Set-FormulaChecked $heat.Cells.Item($r, $c) $formula
    }
  }

  $grid = $heat.Range("B5:T30")
  $grid.NumberFormat = "[>=3]0;[<3]"""";"""";@"
  $grid.HorizontalAlignment = -4108
  $grid.VerticalAlignment = -4108
  $grid.Font.Size = 8
  $grid.Font.Color = Rgb 255 255 255
  $grid.Borders.LineStyle = 1
  $grid.Borders.Weight = 2
  $grid.Borders.Color = Rgb 255 255 255
  $grid.FormatConditions.Delete()
  $scale = $grid.FormatConditions.AddColorScale(3)
  $scale.ColorScaleCriteria.Item(1).Type = 1
  $scale.ColorScaleCriteria.Item(1).FormatColor.Color = Rgb 255 255 255
  $scale.ColorScaleCriteria.Item(2).Type = 5
  $scale.ColorScaleCriteria.Item(2).Value = 50
  $scale.ColorScaleCriteria.Item(2).FormatColor.Color = Rgb 158 197 244
  $scale.ColorScaleCriteria.Item(3).Type = 2
  $scale.ColorScaleCriteria.Item(3).FormatColor.Color = Rgb 31 78 121

  $heat.Columns.Item(1).ColumnWidth = 26
  for ($c = 2; $c -le 20; $c++) { $heat.Columns.Item($c).ColumnWidth = 5.5 }
  $heat.Columns.Item(20).ColumnWidth = 7.5
  $heat.Rows.Item(3).RowHeight = 22
  $heat.Rows.Item(4).RowHeight = 40
  for ($r = 5; $r -le 30; $r++) { $heat.Rows.Item($r).RowHeight = 15 }

  $dashboard.Range("B3").Value2 = ""
  $dashboard.Range("B4").Value2 = ""
  $xl.CalculateFullRebuild()
  Export-Range-Png $heat "A3:T30" (Join-Path $WorkDir "heatmap-all-data.png")

  for ($i = $dashboard.Pictures().Count; $i -ge 1; $i--) {
    $pic = $dashboard.Pictures($i)
    if ($pic.Name -eq "Heatmap - Cause x Week") { $pic.Delete() }
  }
  $heat.Range("A3:T30").Copy()
  $dashboard.Activate()
  $dashboard.Pictures().Paste($true) | Out-Null
  Start-Sleep -Milliseconds 300
  $pic = $dashboard.Pictures($dashboard.Pictures().Count)
  $pic.Name = "Heatmap - Cause x Week"
  $pic.Formula = "='Heatmap Calc'!`$A`$3:`$T`$30"
  if ((Normalize-Picture-Formula ([string]$pic.Formula)) -ne "'Heatmap Calc'!`$A`$3:`$T`$30") {
    throw "Linked picture formula was not set correctly: $($pic.Formula)"
  }
  $pic.Left = 610
  $pic.Top = 1390
  $pic.Width = 560

  $dashboard.Range("B4").Value2 = [datetime]"2026-06-30"
  $xl.CalculateFullRebuild()
  Export-Range-Png $heat "A3:T30" (Join-Path $WorkDir "heatmap-to-jun30.png")

  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-07-15"
  $xl.CalculateFullRebuild()
  Export-Range-Png $heat "A3:T30" (Join-Path $WorkDir "heatmap-final-window.png")

  $wb.SaveAs($LocalPath, 51)
  $wb.Close($true)
  $wb = $null
  Copy-Item -LiteralPath $LocalPath -Destination $WorkbookPath -Force

  $verify = $xl.Workbooks.Open($WorkbookPath)
  $xl.CalculateFullRebuild()
  $vd = $verify.Worksheets.Item("Dashboard")
  $vc = $verify.Worksheets.Item("Data Checks")
  $picFormula = ""
  for ($i = 1; $i -le $vd.Pictures().Count; $i++) {
    if ($vd.Pictures($i).Name -eq "Heatmap - Cause x Week") {
      $picFormula = [string]$vd.Pictures($i).Formula
    }
  }
  if ((Normalize-Picture-Formula $picFormula) -ne "'Heatmap Calc'!`$A`$3:`$T`$30") {
    throw "Reopened linked picture formula failed: $picFormula"
  }
  if ([string]$vc.Range("C2").Text -ne "OK") {
    throw "Heatmap Data Check is not OK after reopen: $($vc.Range("C2").Text)"
  }

  $vd.Range("B4").Value2 = [datetime]"2026-06-30"
  $xl.CalculateFullRebuild()
  $narrowStatus = [string]$vc.Range("C2").Text
  $vd.Range("B3").Value2 = [datetime]"2026-06-01"
  $vd.Range("B4").Value2 = [datetime]"2026-07-15"
  $xl.CalculateFullRebuild()

  $errors = @()
  foreach ($ws in @($verify.Worksheets)) {
    $ur = $ws.UsedRange
    foreach ($err in @("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A")) {
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
  Write-Host "Linked picture formula: $picFormula"
  Write-Host "Final heatmap check: OK"
  Write-Host "Narrow heatmap check: $narrowStatus"
  Write-Host "Verification previews: $WorkDir"
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($verify) { try { $verify.Close($false) } catch {} }
  $xl.Quit()
  Release-Com $xl
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
