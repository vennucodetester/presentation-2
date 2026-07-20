$ErrorActionPreference = "Stop"

$ProjectRoot = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2"
$OriginalPath = Join-Path $ProjectRoot "DG-template.xlsx"
$Fix1Path = Join-Path $ProjectRoot "DG-template-fix1.xlsx"
$Impl2Path = Join-Path $ProjectRoot "DG-template-impl2.xlsx"
$WorkDir = Join-Path $ProjectRoot "outputs\impl2-heatmap-verify"
$LocalPath = Join-Path $env:TEMP "DG-template-impl2-working.xlsx"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

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

function Last-Used-Row($ws) {
  return $ws.Cells.Find("*", $ws.Cells.Item(1,1), -4163, 2, 1, 2, $false).Row
}

function Export-Range-Png($ws, [string]$address, [string]$pngPath) {
  $range = $ws.Range($address)
  $range.CopyPicture(1, 2)
  Start-Sleep -Milliseconds 300
  $chartObj = $ws.ChartObjects().Add([double]$range.Left, [double]$range.Top, [double]$range.Width, [double]$range.Height)
  $chartObj.Activate()
  $chartObj.Chart.Paste()
  $chartObj.Chart.Export($pngPath, "PNG") | Out-Null
  $chartObj.Delete()
  if (!(Test-Path $pngPath) -or ((Get-Item $pngPath).Length -eq 0)) {
    throw "Range export failed or produced a zero-byte PNG: $pngPath"
  }
}

function Export-Chart-Png($chartObj, [string]$pngPath) {
  $chartObj.Activate()
  $chartObj.Chart.Export($pngPath, "PNG") | Out-Null
  if (!(Test-Path $pngPath) -or ((Get-Item $pngPath).Length -eq 0)) {
    throw "Chart export failed or produced a zero-byte PNG: $pngPath"
  }
}

function Apply-Fix1($wb) {
  $graphs = $wb.Worksheets.Item("Graphs")
  $summary = $wb.Worksheets.Item("Summary")
  $settings = $wb.Worksheets.Item("Root Cause Settings")
  $week = $wb.Worksheets.Item("By Commission Week")
  $checks = $wb.Worksheets.Item("Data Checks")

  for ($i = 1; $i -le 20; $i++) {
    Set-FormulaChecked $graphs.Cells.Item($i, 40) "=TEXT('By Commission Week'!A$($i + 154),""mm/dd"")"
    Set-FormulaChecked $graphs.Cells.Item($i, 41) "=TEXT('By Commission Week'!A$($i + 105),""mm/dd"")"
  }

  Set-FormulaChecked $graphs.Range("H6") '="Issues by root cause — after vs before clean point  ("&$B$5&")"'

  $summary.Range("E4:E240").Validation.Delete()
  $summary.Range("E4:E240").Validation.Add(3, 1, 1, "='Root Cause Settings'!`$A`$4:`$A`$80")
  $summary.Range("E4:E240").Validation.IgnoreBlank = $true
  $summary.Range("E4:E240").Validation.InCellDropdown = $true
  $summary.Range("E4:E240").Validation.ErrorTitle = "Invalid root cause"
  $summary.Range("E4:E240").Validation.ErrorMessage = "Pick a valid Root Cause from Root Cause Settings."
  $summary.Range("E4:E240").Validation.ShowError = $true

  Set-FormulaChecked $checks.Range("C4") '=IF(B4>=138,"OK","Review")'
  $checks.Range("D4").Value2 = "Count can only grow from the migrated 138."
  Set-FormulaChecked $checks.Range("B5") '=COUNTIF(''Root Cause Settings''!$A$4:$A$80,"<>")'
  Set-FormulaChecked $checks.Range("C5") '=IF(B5=COUNTIF(''Root Cause Settings''!$A$4:$A$80,"<>"),"OK","Review")'
  $checks.Range("D5").Value2 = "Live count of non-blank root causes in Settings."

  # The implementation target and existing Summary data use these labels. Align the copy so
  # validation, chart labels, and the heatmap all reconcile without touching the original file.
  for ($r = 4; $r -le 80; $r++) {
    if ([string]$settings.Cells.Item($r, 1).Value2 -eq "Installation issue") {
      $settings.Cells.Item($r, 1).Value2 = "DG Installation issue"
    }
  }
  for ($r = 4; $r -le 240; $r++) {
    if ([string]$summary.Cells.Item($r, 5).Value2 -eq "Incorrect Program") {
      $summary.Cells.Item($r, 5).Value2 = "Wrong program"
    }
    if ([string]$summary.Cells.Item($r, 5).Value2 -eq "Installation issue") {
      $summary.Cells.Item($r, 5).Value2 = "DG Installation issue"
    }
  }
  for ($c = 2; $c -le 30; $c++) {
    if ([string]$week.Cells.Item(154, $c).Value2 -eq "Incorrect Program") {
      $week.Cells.Item(154, $c).Value2 = "Wrong program"
    }
    if ([string]$week.Cells.Item(154, $c).Value2 -eq "Installation issue") {
      $week.Cells.Item(154, $c).Value2 = "DG Installation issue"
    }
  }
}

function Build-Heatmap($wb) {
  $graphs = $wb.Worksheets.Item("Graphs")
  $summary = $wb.Worksheets.Item("Summary")
  $settings = $wb.Worksheets.Item("Root Cause Settings")
  $week = $wb.Worksheets.Item("By Commission Week")
  $checks = $wb.Worksheets.Item("Data Checks")

  $heat = Ensure-Sheet $wb "Heatmap Calc" $checks
  $heat.Tab.Color = Rgb 128 128 128
  $heat.Cells.Clear()
  $heat.Cells.FormatConditions.Delete()
  $heat.Visible = -1
  $heat.Activate()
  $heat.Application.ActiveWindow.DisplayGridlines = $false

  $heat.Range("A1").Value2 = "Cause x week heatmap - source range for the linked picture on Graphs. Format HERE, not on the picture."
  $heat.Range("A1:U1").Merge() | Out-Null
  $heat.Range("A1").Font.Size = 9
  $heat.Range("A1").Font.Italic = $true
  $heat.Range("A1").Font.Color = Rgb 90 90 90

  Set-FormulaChecked $heat.Range("A3") '="Issues by root cause x call week  ("&Graphs!$B$5&")"'
  $heat.Range("A3:U3").Merge() | Out-Null
  $heat.Range("A3").HorizontalAlignment = -4108
  $heat.Range("A3").Font.Bold = $true
  $heat.Range("A3").Font.Size = 12
  $heat.Range("A3").Font.Name = "Calibri"

  for ($c = 2; $c -le 20; $c++) {
    $sourceRow = 104 + $c
    Set-FormulaChecked $heat.Cells.Item(4, $c) "=TEXT('By Commission Week'!A$sourceRow,""mm/dd"")"
  }
  $heat.Range("B4:U4").Font.Size = 9
  $heat.Range("B4:U4").HorizontalAlignment = -4108
  $heat.Range("B4:U4").Orientation = 45

  for ($r = 5; $r -le 30; $r++) {
    $settingsRow = $r - 1
    Set-FormulaChecked $heat.Cells.Item($r, 1) "=IF('Root Cause Settings'!A$settingsRow="""","""",'Root Cause Settings'!A$settingsRow)"
  }
  $heat.Range("A5:A30").Font.Size = 9.5
  $heat.Range("A5:A30").HorizontalAlignment = -4152

  for ($r = 5; $r -le 30; $r++) {
    for ($c = 2; $c -le 20; $c++) {
      $weekRow = 104 + $c
      $formula = "=IF(`$A$r="""","""",COUNTIFS(Summary!`$E`$4:`$E`$240,`$A$r,Summary!`$A`$4:`$A`$240,"">=""&'By Commission Week'!`$A`$$weekRow,Summary!`$A`$4:`$A`$240,""<""&'By Commission Week'!`$A`$$weekRow+7,Summary!`$A`$4:`$A`$240,"">=""&IF(Graphs!`$B`$3="""",0,Graphs!`$B`$3),Summary!`$A`$4:`$A`$240,""<""&IF(Graphs!`$B`$4="""",2958465,Graphs!`$B`$4+1)))"
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

  $graphs.Activate()
  $graphs.Range("B3").Value2 = ""
  $graphs.Range("B4").Value2 = ""
  $wb.Application.CalculateFullRebuild()
  Export-Range-Png $heat "A3:T30" (Join-Path $WorkDir "heatmap-source-all-data.png")

  for ($i = $graphs.Pictures().Count; $i -ge 1; $i--) {
    $pic = $graphs.Pictures($i)
    if ($pic.Name -eq "Heatmap - Cause x Week") { $pic.Delete() }
  }

  $heat.Range("A3:T30").Copy()
  $graphs.Activate()
  $pasted = $graphs.Pictures().Paste($true)
  Start-Sleep -Milliseconds 300
  $pic = $graphs.Pictures($graphs.Pictures().Count)
  $pic.Name = "Heatmap - Cause x Week"
  $pic.Formula = "='Heatmap Calc'!`$A`$3:`$T`$30"
  if ((Normalize-Picture-Formula ([string]$pic.Formula)) -ne "'Heatmap Calc'!`$A`$3:`$T`$30") {
    throw "Linked picture formula was not set correctly: $($pic.Formula)"
  }
  $pic.Left = 32
  $pic.Top = 132
  $pic.Width = 620

  if ([string]$graphs.Range("A1").Value2 -notmatch "Heatmap Calc") {
    $existing = [string]$graphs.Range("A1").Value2
    if ($existing.Trim().Length -gt 0) {
      $graphs.Range("A1").Value2 = $existing + "`nHeatmap is a linked picture; format it on Heatmap Calc."
    } else {
      $graphs.Range("A1").Value2 = "Heatmap is a linked picture; format it on Heatmap Calc."
    }
    $graphs.Range("A1").WrapText = $true
  }

  $row = [Math]::Max((Last-Used-Row $checks) + 1, 10)
  $checks.Cells.Item($row, 1).Value2 = "Heatmap classified line reconciliation"
  Set-FormulaChecked $checks.Cells.Item($row, 2) "=SUM('Heatmap Calc'!`$B`$5:`$T`$30)"
  Set-FormulaChecked $checks.Cells.Item($row, 3) "=IF(B$row=SUMPRODUCT(--(Summary!`$E`$4:`$E`$240<>""""),--ISNUMBER(MATCH(Summary!`$E`$4:`$E`$240,'Root Cause Settings'!`$A`$4:`$A`$29,0)),--(Summary!`$A`$4:`$A`$240>=IF(Graphs!`$B`$3="""",0,Graphs!`$B`$3)),--(Summary!`$A`$4:`$A`$240<IF(Graphs!`$B`$4="""",2958465,Graphs!`$B`$4+1))),""OK"",""Review"")"
  $checks.Cells.Item($row, 4).Value2 = "Heatmap sum must equal classified Summary lines in the active Graphs date window."
  $checks.Cells.Item($row, 5).Value2 = "Heatmap Calc vs Summary"
  $checks.Range("A$($row):E$($row)").Interior.Color = Rgb 242 242 242

  $graphs.Range("B4").Value2 = [datetime]"2026-06-30"
  $wb.Application.CalculateFullRebuild()
  Export-Range-Png $heat "A3:T30" (Join-Path $WorkDir "heatmap-source-to-jun30.png")

  $originalCause = [string]$summary.Range("E4").Value2
  $summary.Range("E4").Value2 = "High Store RH"
  $wb.Application.CalculateFullRebuild()
  Export-Range-Png $heat "A3:T30" (Join-Path $WorkDir "heatmap-source-summary-change-test.png")
  $summary.Range("E4").Value2 = $originalCause

  $graphs.Range("B3").Value2 = [datetime]"2026-06-01"
  $graphs.Range("B4").Value2 = [datetime]"2026-07-15"
  $wb.Application.CalculateFullRebuild()
  Export-Range-Png $heat "A3:T30" (Join-Path $WorkDir "heatmap-source-final-window.png")

  Export-Chart-Png $graphs.ChartObjects(1) (Join-Path $WorkDir "chart1-spotcheck.png")
  Export-Chart-Png $graphs.ChartObjects(5) (Join-Path $WorkDir "chart5-spotcheck.png")

  return @{
    DataCheckRow = $row
    PictureFormula = [string]$pic.Formula
  }
}

foreach ($p in @($Fix1Path, $Impl2Path, $LocalPath)) {
  if (Test-Path $p) { Remove-Item -LiteralPath $p -Force }
}
Copy-Item -LiteralPath $OriginalPath -Destination $Fix1Path -Force
Copy-Item -LiteralPath $Fix1Path -Destination $LocalPath -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $null
try {
  $wb = $xl.Workbooks.Open($LocalPath)
  try { $xl.Calculation = -4105 } catch {}

  Apply-Fix1 $wb
  $xl.CalculateFullRebuild()
  $wb.SaveAs($Fix1Path, 51)

  $result = Build-Heatmap $wb
  $wb.SaveAs($LocalPath, 51)
  $wb.Close($true)
  $wb = $null

  Copy-Item -LiteralPath $LocalPath -Destination $Impl2Path -Force

  $verify = $xl.Workbooks.Open($Impl2Path)
  $xl.CalculateFullRebuild()
  $vg = $verify.Worksheets.Item("Graphs")
  $vh = $verify.Worksheets.Item("Heatmap Calc")
  $vc = $verify.Worksheets.Item("Data Checks")

  $vg.Range("B3").Value2 = ""
  $vg.Range("B4").Value2 = ""
  $xl.CalculateFullRebuild()
  Export-Range-Png $vh "A3:T30" (Join-Path $WorkDir "reopened-all-data.png")
  $vg.Range("B4").Value2 = [datetime]"2026-06-30"
  $xl.CalculateFullRebuild()
  Export-Range-Png $vh "A3:T30" (Join-Path $WorkDir "reopened-to-jun30.png")

  $vg.Range("B3").Value2 = [datetime]"2026-06-01"
  $vg.Range("B4").Value2 = [datetime]"2026-07-15"
  $xl.CalculateFullRebuild()

  $picFormula = ""
  for ($i = 1; $i -le $vg.Pictures().Count; $i++) {
    if ($vg.Pictures($i).Name -eq "Heatmap - Cause x Week") {
      $picFormula = [string]$vg.Pictures($i).Formula
    }
  }
  if ((Normalize-Picture-Formula $picFormula) -ne "'Heatmap Calc'!`$A`$3:`$T`$30") {
    throw "Reopened linked picture formula failed: $picFormula"
  }
  if ([string]$vc.Cells.Item($result.DataCheckRow, 3).Text -ne "OK") {
    throw "Heatmap Data Checks row is not OK after reopen: $($vc.Cells.Item($result.DataCheckRow, 3).Text)"
  }

  $verify.Save()
  $verify.Close($true)
  $verify = $null

  Write-Host "Created $Fix1Path"
  Write-Host "Created $Impl2Path"
  Write-Host "Heatmap check row: $($result.DataCheckRow)"
  Write-Host "Linked picture formula: $picFormula"
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
