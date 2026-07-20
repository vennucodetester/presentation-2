$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$BackupPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.before-timeline-group-basis-fix.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\timeline-group-basis-fix-verify"
$LocalPath = Join-Path $env:TEMP "DG-New-master-timeline-group-basis-fix.xlsx"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
if (!(Test-Path $BackupPath)) {
  Copy-Item -LiteralPath $WorkbookPath -Destination $BackupPath -Force
}
Copy-Item -LiteralPath $WorkbookPath -Destination $LocalPath -Force

function Set-FormulaChecked($cell, [string]$formula) {
  try { $cell.Formula = $formula } catch {
    throw "Formula write failed at $($cell.Worksheet.Name)!$($cell.Address($false,$false)): $formula :: $($_.Exception.Message)"
  }
  $readBack = [string]$cell.Formula
  if ($readBack -match "\[\d+\]") {
    throw "External workbook reference appeared in $($cell.Address($false,$false)): $readBack"
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
  $timeline = $wb.Worksheets.Item("Timeline Calc")
  $checks = $wb.Worksheets.Item("Data Checks")
  $chartObj = $dashboard.ChartObjects("Chart 10 - Clean Point Lines")

  Set-FormulaChecked $timeline.Range("A8") '=MIN(IF(''Dashboard''!$B$3="",2958465,''Dashboard''!$B$3),IFERROR(IF(ISNUMBER(MATCH(tl_cause,''Settings''!$B$4:$B$83,0)),IF(tl_basis="Call date",AGGREGATE(15,6,''Summary''!$A$2:$A$240/((''Summary''!$K$2:$K$240=tl_cause)*(''Summary''!$A$2:$A$240>=IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3))*(''Summary''!$A$2:$A$240<IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1))*(''Summary''!$A$2:$A$240>0)),1),AGGREGATE(15,6,''Summary''!$D$2:$D$240/((''Summary''!$K$2:$K$240=tl_cause)*(''Summary''!$A$2:$A$240>=IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3))*(''Summary''!$A$2:$A$240<IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1))*(''Summary''!$D$2:$D$240>0)),1)),IF(tl_basis="Call date",AGGREGATE(15,6,''Summary''!$A$2:$A$240/((''Summary''!$E$2:$E$240=tl_cause)*(''Summary''!$A$2:$A$240>=IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3))*(''Summary''!$A$2:$A$240<IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1))*(''Summary''!$A$2:$A$240>0)),1),AGGREGATE(15,6,''Summary''!$D$2:$D$240/((''Summary''!$E$2:$E$240=tl_cause)*(''Summary''!$A$2:$A$240>=IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3))*(''Summary''!$A$2:$A$240<IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1))*(''Summary''!$D$2:$D$240>0)),1))),IF(''Dashboard''!$B$3="",2958465,''Dashboard''!$B$3)),IF($F$8>0,$F$8,IF(''Dashboard''!$B$3="",2958465,''Dashboard''!$B$3)),IF($G$8>0,$G$8,IF(''Dashboard''!$B$3="",2958465,''Dashboard''!$B$3)),IF($H$8>0,$H$8,IF(''Dashboard''!$B$3="",2958465,''Dashboard''!$B$3)))'
  for ($r = 9; $r -le 191; $r++) {
    Set-FormulaChecked $timeline.Cells.Item($r, 1) "=A$($r - 1)+1"
  }

  for ($r = 8; $r -le 191; $r++) {
    $formula = "=IF(ISNUMBER(MATCH(tl_cause,'Settings'!`$B`$4:`$B`$83,0)),IF(tl_basis=""Call date"",COUNTIFS('Summary'!`$K`$2:`$K`$240,tl_cause,'Summary'!`$A`$2:`$A`$240,"">=""&`$A$r,'Summary'!`$A`$2:`$A`$240,""<""&`$A$r+1,'Summary'!`$A`$2:`$A`$240,"">=""&IF('Dashboard'!`$B`$3="""",0,'Dashboard'!`$B`$3),'Summary'!`$A`$2:`$A`$240,""<""&IF('Dashboard'!`$B`$4="""",2958465,'Dashboard'!`$B`$4+1)),COUNTIFS('Summary'!`$K`$2:`$K`$240,tl_cause,'Summary'!`$D`$2:`$D`$240,"">=""&`$A$r,'Summary'!`$D`$2:`$D`$240,""<""&`$A$r+1,'Summary'!`$A`$2:`$A`$240,"">=""&IF('Dashboard'!`$B`$3="""",0,'Dashboard'!`$B`$3),'Summary'!`$A`$2:`$A`$240,""<""&IF('Dashboard'!`$B`$4="""",2958465,'Dashboard'!`$B`$4+1))),IF(tl_basis=""Call date"",COUNTIFS('Summary'!`$E`$2:`$E`$240,tl_cause,'Summary'!`$A`$2:`$A`$240,"">=""&`$A$r,'Summary'!`$A`$2:`$A`$240,""<""&`$A$r+1,'Summary'!`$A`$2:`$A`$240,"">=""&IF('Dashboard'!`$B`$3="""",0,'Dashboard'!`$B`$3),'Summary'!`$A`$2:`$A`$240,""<""&IF('Dashboard'!`$B`$4="""",2958465,'Dashboard'!`$B`$4+1)),COUNTIFS('Summary'!`$E`$2:`$E`$240,tl_cause,'Summary'!`$D`$2:`$D`$240,"">=""&`$A$r,'Summary'!`$D`$2:`$D`$240,""<""&`$A$r+1,'Summary'!`$A`$2:`$A`$240,"">=""&IF('Dashboard'!`$B`$3="""",0,'Dashboard'!`$B`$3),'Summary'!`$A`$2:`$A`$240,""<""&IF('Dashboard'!`$B`$4="""",2958465,'Dashboard'!`$B`$4+1))))"
    Set-FormulaChecked $timeline.Cells.Item($r, 2) $formula
  }
  for ($r = 8; $r -le 191; $r++) {
    Set-FormulaChecked $timeline.Cells.Item($r, 3) "=IF(AND(`$F`$8<>"""",`$A$r=`$F`$8),1,NA())"
    Set-FormulaChecked $timeline.Cells.Item($r, 4) "=IF(AND(`$G`$8<>"""",`$G`$8>0,`$A$r=`$G`$8),1,NA())"
    Set-FormulaChecked $timeline.Cells.Item($r, 5) "=IF(AND(`$H`$8<>"""",`$H`$8>0,`$A$r=`$H`$8),1,NA())"
  }

  $chartObj.Chart.SeriesCollection(1).Formula = '=SERIES("Calls",''Timeline Calc''!$A$8:$A$191,''Timeline Calc''!$B$8:$B$191,1)'
  $chartObj.Chart.SeriesCollection(2).Formula = '=SERIES(''Timeline Calc''!$I$8,''Timeline Calc''!$A$8:$A$191,''Timeline Calc''!$C$8:$C$191,2)'
  $chartObj.Chart.SeriesCollection(3).Formula = '=SERIES(''Timeline Calc''!$J$8,''Timeline Calc''!$A$8:$A$191,''Timeline Calc''!$D$8:$D$191,3)'
  $chartObj.Chart.SeriesCollection(4).Formula = '=SERIES(''Timeline Calc''!$K$8,''Timeline Calc''!$A$8:$A$191,''Timeline Calc''!$E$8:$E$191,4)'

  Set-FormulaChecked $checks.Range("C4") '=IF(B4=IF(ISNUMBER(MATCH(tl_cause,''Settings''!$B$4:$B$83,0)),COUNTIFS(''Summary''!$K$2:$K$240,tl_cause,''Summary''!$A$2:$A$240,">="&IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3),''Summary''!$A$2:$A$240,"<"&IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1)),COUNTIFS(''Summary''!$E$2:$E$240,tl_cause,''Summary''!$A$2:$A$240,">="&IF(''Dashboard''!$B$3="",0,''Dashboard''!$B$3),''Summary''!$A$2:$A$240,"<"&IF(''Dashboard''!$B$4="",2958465,''Dashboard''!$B$4+1))),"OK","Review")'
  $checks.Range("D4").Value2 = "Timeline sum must equal selected cause/group rows in the Dashboard call-date window; basis only changes plotted dates."

  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-07-15"

  $dashboard.Range("B125").Value2 = "Doors"
  $dashboard.Range("B126").Value2 = "Call date"
  $xl.CalculateFull()
  $doorsCallTimeline = [int]$xl.WorksheetFunction.Sum($timeline.Range("B8:B191"))
  $doorsCallExpected = [int]$xl.Evaluate("SUMPRODUCT(--(Summary!K2:K240=""Doors""),--(Summary!A2:A240>=DATE(2026,6,1)),--(Summary!A2:A240<DATE(2026,7,16)))")
  Export-Chart-Png $chartObj (Join-Path $WorkDir "doors-call-date.png")

  $dashboard.Range("B126").Value2 = "Commission date"
  $xl.CalculateFull()
  $doorsCommissionTimeline = [int]$xl.WorksheetFunction.Sum($timeline.Range("B8:B191"))
  Export-Chart-Png $chartObj (Join-Path $WorkDir "doors-commission-date.png")

  if ($doorsCallTimeline -ne $doorsCallExpected -or $doorsCommissionTimeline -ne $doorsCallExpected) {
    throw "Doors totals failed: call timeline=$doorsCallTimeline commission timeline=$doorsCommissionTimeline expected=$doorsCallExpected"
  }
  if ([string]$checks.Range("C4").Text -ne "OK") {
    throw "Timeline Data Check is not OK after Doors Commission test: $($checks.Range("C4").Text)"
  }

  $dashboard.Range("B125").Value2 = "False Alarm"
  $dashboard.Range("B126").Value2 = "Call date"
  $xl.CalculateFull()

  try { $xl.Calculation = -4105 } catch {}
  $wb.SaveAs($LocalPath, 51)
  $wb.Close($true)
  $wb = $null
  Copy-Item -LiteralPath $LocalPath -Destination $WorkbookPath -Force

  $verify = $xl.Workbooks.Open($WorkbookPath)
  $xl.CalculateFull()
  $vd = $verify.Worksheets.Item("Dashboard")
  $vt = $verify.Worksheets.Item("Timeline Calc")
  $vc = $verify.Worksheets.Item("Data Checks")

  $vd.Range("B125").Value2 = "Doors"
  $vd.Range("B126").Value2 = "Call date"
  $xl.CalculateFull()
  $reopenDoorsCall = [int]$xl.WorksheetFunction.Sum($vt.Range("B8:B191"))
  $vd.Range("B126").Value2 = "Commission date"
  $xl.CalculateFull()
  $reopenDoorsCommission = [int]$xl.WorksheetFunction.Sum($vt.Range("B8:B191"))
  $reopenExpected = [int]$xl.Evaluate("SUMPRODUCT(--(Summary!K2:K240=""Doors""),--(Summary!A2:A240>=DATE(2026,6,1)),--(Summary!A2:A240<DATE(2026,7,16)))")
  if ($reopenDoorsCall -ne $reopenExpected -or $reopenDoorsCommission -ne $reopenExpected) {
    throw "Reopen Doors totals failed: call=$reopenDoorsCall commission=$reopenDoorsCommission expected=$reopenExpected"
  }

  $vd.Range("B125").Value2 = "False Alarm"
  $vd.Range("B126").Value2 = "Call date"
  $vd.Range("B3").Value2 = [datetime]"2026-06-01"
  $vd.Range("B4").Value2 = [datetime]"2026-07-15"
  $xl.CalculateFull()

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
  Write-Host "Doors expected=$doorsCallExpected call=$doorsCallTimeline commission=$doorsCommissionTimeline"
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
