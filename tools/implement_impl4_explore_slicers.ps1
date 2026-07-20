$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.xlsx"
$BackupPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\DG-New master.before-impl4-explore.xlsx"
$WorkDir = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\DG-New master\impl4-explore-verify"
$LocalPath = Join-Path $env:TEMP "DG-New-master-impl4-working.xlsx"
$LogPath = Join-Path $WorkDir "impl4-script.log"

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
Set-Content -Path $LogPath -Value "Impl4 run $(Get-Date -Format o)"
if (!(Test-Path $BackupPath)) {
  Copy-Item -LiteralPath $WorkbookPath -Destination $BackupPath -Force
}
Copy-Item -LiteralPath $WorkbookPath -Destination $LocalPath -Force

function Rgb([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }
function Log-Step([string]$message) { Add-Content -Path $LogPath -Value "$(Get-Date -Format o) $message" }
function Set-FormulaChecked($cell, [string]$formula, [bool]$array = $false) {
  try {
    if ($array) { $cell.FormulaArray = $formula } else { $cell.Formula = $formula }
  } catch {
    throw "Formula write failed at $($cell.Worksheet.Name)!$($cell.Address($false,$false)): $formula :: $($_.Exception.Message)"
  }
  $readBack = [string]$cell.Formula
  if ($readBack -match "\[\d+\]") {
    throw "External workbook reference appeared in $($cell.Address($false,$false)): $readBack"
  }
}
function Delete-Sheet-If-Exists($wb, [string]$name) {
  for ($i = $wb.Worksheets.Count; $i -ge 1; $i--) {
    $ws = $wb.Worksheets.Item($i)
    if ($ws.Name -eq $name) {
      $ws.Delete()
      return
    }
  }
}
function Export-Range-Png($ws, [string]$address, [string]$pngPath) {
  Remove-Item $pngPath -ErrorAction SilentlyContinue
  $range = $ws.Range($address)
  $range.CopyPicture(1, 2)
  Start-Sleep -Milliseconds 500
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
  Remove-Item $pngPath -ErrorAction SilentlyContinue
  $chartObj.Activate()
  Start-Sleep -Milliseconds 500
  $chartObj.Chart.Export($pngPath, "PNG") | Out-Null
  if (!(Test-Path $pngPath) -or ((Get-Item $pngPath).Length -eq 0)) {
    throw "Chart export failed or produced a zero-byte PNG: $pngPath"
  }
}
function Rect-Overlap($a, $b) {
  return (($a.Left -lt ($b.Left + $b.Width)) -and (($a.Left + $a.Width) -gt $b.Left) -and ($a.Top -lt ($b.Top + $b.Height)) -and (($a.Top + $a.Height) -gt $b.Top))
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
  $checks = $wb.Worksheets.Item("Data Checks")

  Log-Step "Write Summary Nomenclature helper"
  $summary.Range("L1").Value2 = "Nomenclature"
  $summary.Range("L1").Font.Bold = $true
  try { $summary.Range("A1:K1").Copy(); $summary.Range("L1").PasteSpecial(-4122) | Out-Null } catch {}
  Set-FormulaChecked $summary.Range("L2") '=IF($H2="","",IFERROR(INDEX(''Case Nomenclature Graph''!$A$2:$A$7,MATCH(TRUE,ISNUMBER(SEARCH(''Case Nomenclature Graph''!$A$2:$A$7,$H2)),0)),"other"))' $true
  $summary.Range("L2").Copy()
  $summary.Range("L3:L240").PasteSpecial(-4123) | Out-Null
  $summary.Columns("L").ColumnWidth = 16

  Log-Step "Create Explore sheet"
  Delete-Sheet-If-Exists $wb "Explore"
  $explore = $wb.Worksheets.Add([System.Type]::Missing, $dashboard)
  $explore.Name = "Explore"
  $explore.Activate()
  $explore.Application.ActiveWindow.DisplayGridlines = $false
  $explore.Range("A1").Value2 = "Slice-and-dice. Slicers filter the chart and table live. Right-click Refresh after adding Summary rows."
  $explore.Range("A1:N1").Merge() | Out-Null
  $explore.Range("A1").Font.Bold = $true
  $explore.Range("A1").Font.Size = 11
  $explore.Range("A1").Interior.Color = Rgb 242 242 242

  Log-Step "Create PivotTable"
  $sourceRange = "'Summary'!R1C1:R240C12"
  $cache = $wb.PivotCaches().Create(1, $sourceRange)
  $pt = $cache.CreatePivotTable($explore.Range("A58"), "ExplorePivot")
  $pt.ManualUpdate = $true
  $pt.PivotFields("Group").Orientation = 1
  $pt.PivotFields("Group").Position = 1
  $dataField = $pt.AddDataField($pt.PivotFields("Work Order"), "Count of Work Order", -4112)
  $pt.PivotFields("Group").AutoSort(2, $dataField.Name)
  $pt.ManualUpdate = $false
  $pt.RefreshTable() | Out-Null

  Log-Step "Create PivotChart"
  $chartObj = $explore.ChartObjects().Add(365.0, 42.0, 805.0, 385.0)
  $chartObj.Name = "Explore PivotChart"
  $chart = $chartObj.Chart
  $chart.SetSourceData($pt.TableRange1)
  $chart.ChartType = 51
  $chart.HasTitle = $true
  $chart.ChartTitle.Text = "Filtered issues - use the slicers"
  $chart.ChartTitle.Font.Bold = $true
  $chart.ChartTitle.Font.Size = 12
  $chart.HasLegend = $false
  try { $chart.ShowAllFieldButtons = $false } catch {}
  try { $chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = Rgb 68 114 196 } catch {}
  try { $chart.SeriesCollection(1).Format.Line.Visible = 0 } catch {}
  try { $chart.SeriesCollection(1).HasDataLabels = $true; $chart.SeriesCollection(1).DataLabels().NumberFormat = "0;;;" } catch {}
  try { $chart.ChartGroups(1).GapWidth = 60 } catch {}
  try { $chart.Axes(1).TickLabels.Orientation = 45 } catch {}

  Log-Step "Create slicers"
  $slicerSpecs = @(
    @{ Field = "Group"; Name = "Slicer_Group"; Caption = "Group"; Left = 15.0; Top = 38.0; Width = 320.0; Height = 118.0; Columns = 1 },
    @{ Field = "Root Cause"; Name = "Slicer_RootCause"; Caption = "Root Cause"; Left = 15.0; Top = 164.0; Width = 320.0; Height = 130.0; Columns = 1 },
    @{ Field = "Nomenclature"; Name = "Slicer_Nomenclature"; Caption = "Nomenclature"; Left = 15.0; Top = 302.0; Width = 320.0; Height = 118.0; Columns = 1 },
    @{ Field = "After Clean Point"; Name = "Slicer_AfterClean"; Caption = "After Clean Point"; Left = 15.0; Top = 428.0; Width = 320.0; Height = 90.0; Columns = 1 },
    @{ Field = "Store"; Name = "Slicer_Store"; Caption = "Store"; Left = 15.0; Top = 526.0; Width = 320.0; Height = 175.0; Columns = 1 }
  )
  $createdObjects = @()
  $slicerCachesByField = @{}
  foreach ($spec in $slicerSpecs) {
    $sc = $wb.SlicerCaches().Add($pt, $spec.Field)
    $slicerCachesByField[$spec.Field] = $sc
    $slicer = $sc.Slicers().Add($explore, [System.Type]::Missing, $spec.Name, $spec.Caption, [double]$spec.Top, [double]$spec.Left, [double]$spec.Width, [double]$spec.Height)
    try { $slicer.Style = "SlicerStyleLight1" } catch {}
    try { $slicer.NumberOfColumns = $spec.Columns } catch {}
    $createdObjects += @{ Name = $spec.Name; Left = $spec.Left; Top = $spec.Top; Width = $spec.Width; Height = $spec.Height }
  }
  $createdObjects += @{ Name = "Explore PivotChart"; Left = 365.0; Top = 42.0; Width = 805.0; Height = 385.0 }
  $pivotRect = @{ Name = "PivotTable"; Left = [double]$pt.TableRange1.Left; Top = [double]$pt.TableRange1.Top; Width = [double]$pt.TableRange1.Width; Height = [double]$pt.TableRange1.Height }
  $createdObjects += $pivotRect
  for ($i = 0; $i -lt $createdObjects.Count; $i++) {
    for ($j = $i + 1; $j -lt $createdObjects.Count; $j++) {
      if (Rect-Overlap $createdObjects[$i] $createdObjects[$j]) {
        throw "Explore object overlap: $($createdObjects[$i].Name) overlaps $($createdObjects[$j].Name)"
      }
    }
  }

  Log-Step "Dashboard/Data Checks integration"
  $note = [string]$dashboard.Range("A1").Value2
  if ($note -notmatch "Explore") {
    $dashboard.Range("A1").Value2 = ($note.Trim() + "`nExplore sheet: use slicers there for slice-and-dice; refresh pivot after new data.").Trim()
    $dashboard.Range("A1").WrapText = $true
  }
  $checks.Range("A5").Value2 = "Explore pivot grand total reconciliation"
  Set-FormulaChecked $checks.Range("B5") '=GETPIVOTDATA("Work Order",Explore!$A$58)'
  Set-FormulaChecked $checks.Range("C5") '=IF(B5=SUMPRODUCT(--(Summary!$E$2:$E$240<>"")),"OK","Review")'
  $checks.Range("D5").Value2 = "Clear Explore slicers and refresh pivot if this shows Review."
  $checks.Range("E5").Value2 = "Explore Pivot vs Summary"
  $checks.Columns("A:E").AutoFit() | Out-Null

  Log-Step "Acceptance exports/click tests"
  $xl.Calculation = -4105
  $xl.CalculateFull()
  $pt.RefreshTable() | Out-Null
  Export-Range-Png $explore "A1:T70" (Join-Path $WorkDir "explore-full-clear.png")
  Export-Chart-Png $dashboard.ChartObjects("Chart 1 - Root Cause") (Join-Path $WorkDir "dashboard-chart1-spotcheck.png")
  Export-Chart-Png $dashboard.ChartObjects("Chart 10 - Clean Point Lines") (Join-Path $WorkDir "dashboard-timeline-spotcheck.png")

  # Click test 2: Nomenclature RLN3MA + After Clean Point Yes.
  $slicerCachesByField["Nomenclature"].ClearManualFilter()
  $slicerCachesByField["After Clean Point"].ClearManualFilter()
  $slicerCachesByField["Nomenclature"].SlicerItems("RLN3MA").Selected = $true
  for ($i = 1; $i -le $slicerCachesByField["Nomenclature"].SlicerItems.Count; $i++) {
    $item = $slicerCachesByField["Nomenclature"].SlicerItems($i)
    if ($item.Name -ne "RLN3MA") { try { $item.Selected = $false } catch {} }
  }
  $slicerCachesByField["After Clean Point"].SlicerItems("Yes").Selected = $true
  for ($i = 1; $i -le $slicerCachesByField["After Clean Point"].SlicerItems.Count; $i++) {
    $item = $slicerCachesByField["After Clean Point"].SlicerItems($i)
    if ($item.Name -ne "Yes") { try { $item.Selected = $false } catch {} }
  }
  $pt.RefreshTable() | Out-Null
  $pivotFiltered = $pt.DataBodyRange.Cells.Item($pt.DataBodyRange.Rows.Count, 1).Value2
  $liveFiltered = $xl.Evaluate("SUMPRODUCT(--(Summary!L2:L240=""RLN3MA""),--(Summary!G2:G240=""Yes""))")
  Log-Step "Click test 2 RLN3MA+Yes pivot=$pivotFiltered live=$liveFiltered"
  Export-Range-Png $explore "A1:T70" (Join-Path $WorkDir "explore-filter-rln3ma-yes.png")

  # Click test 3: Group Doors.
  foreach ($fieldName in @("Nomenclature", "After Clean Point", "Group", "Root Cause", "Store")) {
    try { $slicerCachesByField[$fieldName].ClearManualFilter() } catch {}
  }
  $slicerCachesByField["Group"].SlicerItems("Doors").Selected = $true
  for ($i = 1; $i -le $slicerCachesByField["Group"].SlicerItems.Count; $i++) {
    $item = $slicerCachesByField["Group"].SlicerItems($i)
    if ($item.Name -ne "Doors") { try { $item.Selected = $false } catch {} }
  }
  $pt.RefreshTable() | Out-Null
  Export-Range-Png $explore "A1:T70" (Join-Path $WorkDir "explore-filter-doors.png")
  foreach ($fieldName in @("Group", "Nomenclature", "After Clean Point", "Root Cause", "Store")) {
    try { $slicerCachesByField[$fieldName].ClearManualFilter() } catch {}
  }
  $pt.RefreshTable() | Out-Null
  $dashboard.Range("B3").Value2 = [datetime]"2026-06-01"
  $dashboard.Range("B4").Value2 = [datetime]"2026-07-15"
  $xl.CalculateFull()
  $wb.SaveAs($LocalPath, 51)
  $wb.Close($true)
  $wb = $null
  Copy-Item -LiteralPath $LocalPath -Destination $WorkbookPath -Force

  Log-Step "Reopen verification"
  $verify = $xl.Workbooks.Open($WorkbookPath)
  $xl.CalculateFull()
  if ($xl.Calculation -ne -4105) { throw "Calculation is not automatic after reopen: $($xl.Calculation)" }
  $ve = $verify.Worksheets.Item("Explore")
  $vc = $verify.Worksheets.Item("Data Checks")
  $vd = $verify.Worksheets.Item("Dashboard")
  $vpt = $ve.PivotTables("ExplorePivot")
  $vpt.RefreshTable() | Out-Null
  if ([string]$vc.Range("C5").Text -ne "OK") { throw "Explore Data Check C5 is not OK: $($vc.Range("C5").Text)" }
  if ([string]$vd.Range("B3").Text -ne "6/1/2026" -or [string]$vd.Range("B4").Text -ne "7/15/2026") {
    throw "Dashboard dates were not restored"
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
  Write-Host "Click test 2 pivot=$pivotFiltered live=$liveFiltered"
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($verify) { try { $verify.Close($false) } catch {} }
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
