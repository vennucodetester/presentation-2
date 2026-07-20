$ErrorActionPreference = "Stop"

$WorkbookPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\outputs\019f6dfb-339c-7212-a7a5-679970584164\DG-template-reorganized-phase5.xlsx"
$LogPath = "C:\Users\silam\OneDrive\Documents\Presentation\Presentation-2\outputs\019f6dfb-339c-7212-a7a5-679970584164\phase5-implementation.log"
Set-Content -LiteralPath $LogPath -Value "Phase 5 implementation started $(Get-Date -Format o)"

function Log-Step([string]$message) {
  Add-Content -LiteralPath $LogPath -Value "$(Get-Date -Format o) $message"
}

function Set-Fill($range, [int]$color) {
  $range.Interior.Color = $color
}

function Set-Border($range) {
  $range.Borders.LineStyle = 1
  $range.Borders.Color = 14277081
}

function Add-Or-Get-Sheet($wb, [string]$name, [bool]$visible = $true) {
  foreach ($ws in $wb.Worksheets) {
    if ($ws.Name -eq $name) {
      $ws.Visible = $(if ($visible) { -1 } else { 0 })
      return $ws
    }
  }
  $newWs = $wb.Worksheets.Add([System.Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
  $newWs.Name = $name
  $newWs.Visible = $(if ($visible) { -1 } else { 0 })
  return $newWs
}

function Delete-Name-If-Exists($wb, [string]$name) {
  foreach ($n in @($wb.Names)) {
    if ($n.Name -eq $name) {
      $n.Delete()
      return
    }
  }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.ScreenUpdating = $false
try { $xl.CalculateBeforeSave = $false } catch { Log-Step "Could not disable CalculateBeforeSave before open: $($_.Exception.Message)" }
try { $xl.Calculation = -4135; Log-Step "Set calculation manual before open" } catch { Log-Step "Could not set calculation before open: $($_.Exception.Message)" }

try {
  Log-Step "Opening workbook"
  $wb = $xl.Workbooks.Open($WorkbookPath)
  try { $xl.CalculateBeforeSave = $false } catch { Log-Step "Could not disable CalculateBeforeSave after open: $($_.Exception.Message)" }
  try { $xl.Calculation = -4135; Log-Step "Set calculation manual after open" } catch { Log-Step "Could not set calculation after open: $($_.Exception.Message)" }
  $wsReadme = $wb.Worksheets.Item("README")
  $wsDash = $wb.Worksheets.Item("Dashboard")
  $wsSummary = $wb.Worksheets.Item("Summary")
  $wsRaw = $wb.Worksheets.Item("Raw Data")
  $wsSettings = $wb.Worksheets.Item("Settings")
  $wsStore = $wb.Worksheets.Item("Store List")
  $wsWeekly = $wb.Worksheets.Item("Calc_Weekly")
  $wsCharts = $wb.Worksheets.Item("Calc_Charts")
  $wsChecks = $wb.Worksheets.Item("Data Checks")

  # Part A1: group registry in unused right-side columns.
  Log-Step "Part A1/A5 formula engine"
  $wsCharts.Range("AG1:AH30").Clear()
  $wsCharts.Range("AG1").Value2 = "T0. Group Registry"
  $wsCharts.Range("AG2").Value2 = "Group"
  $wsCharts.Range("AH2").Value2 = "First row"
  for ($i = 1; $i -le 25; $i++) {
    $r = 2 + $i
    $wsCharts.Cells.Item($r, 33).Formula = "=IFERROR(INDEX(Settings!`$B`$4:`$B`$59,SMALL(`$AH`$3:`$AH`$58,$i)-3),"""")"
  }
  for ($r = 3; $r -le 58; $r++) {
    $settingsRow = $r + 1
    $wsCharts.Cells.Item($r, 34).Formula = "=IF(Settings!A$settingsRow="""","""",IF(COUNTIF(Settings!`$B`$4:Settings!B$settingsRow,Settings!B$settingsRow)=1,ROW(Settings!B$settingsRow),""""))"
  }
  Delete-Name-If-Exists $wb "set_groupList"
  $wb.Names.Add("set_groupList", "=OFFSET(Calc_Charts!`$AG`$3,0,0,MAX(1,COUNTA(Calc_Charts!`$AG`$3:`$AG`$27)),1)") | Out-Null
  Set-Fill $wsCharts.Range("AG1:AH2") 12632256
  $wsCharts.Range("AG:AH").ColumnWidth = 24

  # Part A5: replace old Graphs/Counts mirror engine in place.
  for ($r = 6; $r -le 61; $r++) {
    $idx = $r - 5
    $wsCharts.Cells.Item($r, 1).Formula = "=IFERROR(INDEX(Settings!`$A`$4:`$A`$59,$idx),"""")"
    $wsCharts.Cells.Item($r, 2).Formula = "=IF(`$A$r="""","""",SUMPRODUCT((Summary!`$E`$5:`$E`$241=`$A$r)*ISNUMBER(Summary!`$A`$5:`$A`$241)*(Summary!`$A`$5:`$A`$241>=IF(set_from="""",0,set_from))*(Summary!`$A`$5:`$A`$241<IF(set_to="""",2958465,set_to+1))))"
    $wsCharts.Cells.Item($r, 3).Formula = "=IF(`$A$r="""","""",SUMPRODUCT((Summary!`$E`$5:`$E`$241=`$A$r)*(Summary!`$G`$5:`$G`$241=""Yes"")*ISNUMBER(Summary!`$A`$5:`$A`$241)*(Summary!`$A`$5:`$A`$241>=IF(set_from="""",0,set_from))*(Summary!`$A`$5:`$A`$241<IF(set_to="""",2958465,set_to+1))))"
    $tie = 62 - $r
    $wsCharts.Cells.Item($r, 4).Formula = "=IF(`$A$r="""","""",B$r+($tie/set_tiebreakDivisor))"
  }
  for ($r = 4; $r -le 59; $r++) {
    $chartRow = $r + 2
    $wsCharts.Cells.Item($r, 28).Formula = "=Settings!B$r"
    $wsCharts.Cells.Item($r, 29).Formula = "=Calc_Charts!B$chartRow"
    $wsCharts.Cells.Item($r, 30).Formula = "=Calc_Charts!C$chartRow"
    $wsCharts.Cells.Item($r, 31).Formula = "=IF(Settings!A$r="""","""",IF(COUNTIF(Settings!`$B`$4:Settings!B$r,Settings!B$r)=1,ROW(Settings!B$r),""""))"
  }
  for ($r = 36; $r -le 60; $r++) {
    $idx = $r - 35
    $tie = 61 - $r
    $wsCharts.Cells.Item($r, 1).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
    $wsCharts.Cells.Item($r, 2).Formula = "=IF(OR(A$r="""",A$r=0),"""",MATCH(A$r,Settings!`$B`$4:`$B`$59,0)+3)"
    $wsCharts.Cells.Item($r, 3).Formula = "=IF(A$r="""","""",SUMIF(`$AB`$4:`$AB`$59,A$r,`$AC`$4:`$AC`$59))"
    $wsCharts.Cells.Item($r, 4).Formula = "=IF(A$r="""","""",SUMIF(`$AB`$4:`$AB`$59,A$r,`$AD`$4:`$AD`$59))"
    $wsCharts.Cells.Item($r, 5).Formula = "=IF(A$r="""","""",C$r+($tie/set_tiebreakDivisor))"
    $wsCharts.Cells.Item($r, 6).Formula = "=IF(A$r="""","""",D$r+($tie/set_tiebreakDivisor))"
  }
  for ($r = 66; $r -le 90; $r++) {
    $idx = $r - 65
    $wsCharts.Cells.Item($r, 1).Formula = "=IFERROR(INDEX(`$A`$36:`$A`$60,MATCH(LARGE(`$E`$36:`$E`$60,$idx),`$E`$36:`$E`$60,0)),"""")"
    $wsCharts.Cells.Item($r, 2).Formula = "=IF(A$r="""","""",INDEX(`$C`$36:`$C`$60,MATCH(A$r,`$A`$36:`$A`$60,0)))"
  }
  for ($r = 96; $r -le 120; $r++) {
    $idx = $r - 95
    $wsCharts.Cells.Item($r, 1).Formula = "=IFERROR(INDEX(`$A`$36:`$A`$60,MATCH(LARGE(`$F`$36:`$F`$60,$idx),`$F`$36:`$F`$60,0)),"""")"
    $wsCharts.Cells.Item($r, 2).Formula = "=IF(A$r="""","""",INDEX(`$D`$36:`$D`$60,MATCH(A$r,`$A`$36:`$A`$60,0)))"
  }
  for ($r = 126; $r -le 150; $r++) {
    $idx = $r - 125
    $wsCharts.Cells.Item($r, 1).Formula = "=IFERROR(INDEX(`$A`$36:`$A`$60,MATCH(LARGE(`$F`$36:`$F`$60,$idx),`$F`$36:`$F`$60,0)),"""")"
    $wsCharts.Cells.Item($r, 2).Formula = "=IF(A$r="""","""",INDEX(`$D`$36:`$D`$60,MATCH(A$r,`$A`$36:`$A`$60,0)))"
    $wsCharts.Cells.Item($r, 3).Formula = "=IF(A$r="""","""",INDEX(`$C`$36:`$C`$60,MATCH(A$r,`$A`$36:`$A`$60,0))-B$r)"
    $wsCharts.Cells.Item($r, 4).Formula = "=IF(A$r="""","""",B$r+C$r)"
  }
  for ($r = 187; $r -le 206; $r++) {
    $wk1 = $r - 32
    $wk2 = $r - 81
    $wsCharts.Cells.Item($r, 2).Formula = "=SUM(Calc_Weekly!C$wk1:AA$wk1)"
    $wsCharts.Cells.Item($r, 4).Formula = "=SUM(Calc_Weekly!C$wk2:AA$wk2)"
  }
  for ($r = 218; $r -le 243; $r++) {
    $idx = $r - 212
    $wsCharts.Cells.Item($r, 1).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
    $wsCharts.Cells.Item($r, 2).Formula = "=IF(A$r="""","""",SUMIF(Summary!`$L`$5:`$L`$241,A$r,Summary!`$A`$5:`$A`$241))"
    $wsCharts.Cells.Item($r, 3).Formula = "=IF(A$r="""","""",SUMPRODUCT(SUBTOTAL(103,OFFSET(Summary!`$E`$5,ROW(Summary!`$E`$5:`$E`$241)-ROW(Summary!`$E`$5),0,1)),--(Summary!`$L`$5:`$L`$241=A$r)))"
    $wsCharts.Cells.Item($r, 4).Formula = "=IF(A$r="""","""",COUNTIFS(Summary!`$L`$5:`$L`$241,A$r,Summary!`$G`$5:`$G`$241,""Yes""))"
  }

  # Part A1/A2: formula-driven group headers.
  foreach ($row in @(65,95)) {
    for ($col = 3; $col -le 27; $col++) {
      $idx = $col - 2
      $wsCharts.Cells.Item($row, $col).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
    }
  }
  foreach ($row in @(155,386,541,696)) {
    for ($col = 2; $col -le 26; $col++) {
      $idx = $col - 1
      $wsCharts.Cells.Item($row, $col).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
    }
  }
  for ($col = 3; $col -le 27; $col++) {
    $idx = $col - 2
    $wsWeekly.Cells.Item(154, $col).Formula = "=IFERROR(INDEX(set_groupList,$idx),"""")"
  }

  # Part A3: tighten validation.
  Log-Step "Part A3 validation"
  $wsSummary.Range("E5:E241").Validation.Delete()
  [void]$wsSummary.Range("E5:E241").Validation.Add(3, 1, 1, "=set_causes")
  $wsSummary.Range("E5:E241").Validation.IgnoreBlank = $true
  $wsSummary.Range("E5:E241").Validation.InCellDropdown = $true
  $wsSummary.Range("E5:E241").Validation.ErrorTitle = "Invalid root cause"
  $wsSummary.Range("E5:E241").Validation.ErrorMessage = "Pick a root cause from Settings - add new causes there first"
  foreach ($addr in @("B143:B146","B225:B241")) {
    try { $wsSummary.Range($addr).Validation.Delete() } catch {}
  }

  # Part A6: central chart label.
  Log-Step "Part A6 chart labels"
  $wsCharts.Range("AG33").Value2 = "Before / at clean point"
  foreach ($chartIndex in @(5,7,8,9)) {
    try {
      $co = $wsDash.ChartObjects($chartIndex)
      foreach ($ser in $co.Chart.SeriesCollection()) {
        if ([string]$ser.Name -eq "Before / at clean point") {
          $ser.Name = "=Calc_Charts!`$AG`$33"
        }
      }
    } catch {}
  }

  # Part B0: enrich Summary columns I:N.
  Log-Step "Part B0 Summary enrichment"
  $headers = @("Branch","City","State","Group","Nomenclature","Age at failure (days)")
  for ($i = 0; $i -lt $headers.Count; $i++) {
    $wsSummary.Cells.Item(4, 9 + $i).Value2 = $headers[$i]
  }
  for ($r = 5; $r -le 241; $r++) {
    $wsSummary.Cells.Item($r, 9).Formula = "=IFERROR(INDEX('Raw Data'!`$I`$5:`$I`$400,MATCH(`$B$r,'Raw Data'!`$H`$5:`$H`$400,0)),"""")"
    $wsSummary.Cells.Item($r, 10).Formula = "=IFERROR(INDEX('Raw Data'!`$K`$5:`$K`$400,MATCH(`$B$r,'Raw Data'!`$H`$5:`$H`$400,0)),"""")"
    $wsSummary.Cells.Item($r, 11).Formula = "=IFERROR(INDEX('Raw Data'!`$L`$5:`$L`$400,MATCH(`$B$r,'Raw Data'!`$H`$5:`$H`$400,0)),"""")"
    $wsSummary.Cells.Item($r, 12).Formula = "=IFERROR(INDEX(Settings!`$B`$4:`$B`$59,MATCH(`$E$r,set_causes,0)),"""")"
    $wsSummary.Cells.Item($r, 13).Formula = "=IF(`$H$r="""","""",IF(ISNUMBER(SEARCH(Settings!`$A`$75,`$H$r)),Settings!`$A`$75,IF(ISNUMBER(SEARCH(Settings!`$A`$76,`$H$r)),Settings!`$A`$76,IF(ISNUMBER(SEARCH(Settings!`$A`$77,`$H$r)),Settings!`$A`$77,IF(ISNUMBER(SEARCH(Settings!`$A`$78,`$H$r)),Settings!`$A`$78,IF(ISNUMBER(SEARCH(Settings!`$A`$79,`$H$r)),Settings!`$A`$79,IF(ISNUMBER(SEARCH(Settings!`$A`$80,`$H$r)),Settings!`$A`$80,""other"")))))))"
    $wsSummary.Cells.Item($r, 14).Formula = "=IF(OR(`$A$r="""",`$D$r=""""),"""",`$A$r-`$D$r)"
  }
  $wsSummary.Range("I4:N4").Font.Bold = $true
  Set-Fill $wsSummary.Range("I4:N4") 12632256
  $wsSummary.Range("I:N").ColumnWidth = 18
  $wsSummary.Range("N5:N241").NumberFormat = "0"

  # Part A7: new data checks.
  Log-Step "Part A7 data checks"
  $checks = @(
    @("Matrix headers match group registry","=SUMPRODUCT(--(Calc_Charts!`$C`$65:`$AA`$65<>TRANSPOSE(TRANSPOSE(Calc_Charts!`$AG`$3:`$AG`$27))))+SUMPRODUCT(--(Calc_Charts!`$C`$65:`$AA`$65<>Calc_Charts!`$C`$95:`$AA`$95))+SUMPRODUCT(--(Calc_Charts!`$B`$155:`$Z`$155<>Calc_Charts!`$B`$386:`$Z`$386))+SUMPRODUCT(--(Calc_Charts!`$B`$155:`$Z`$155<>Calc_Charts!`$B`$541:`$Z`$541))+SUMPRODUCT(--(Calc_Charts!`$B`$155:`$Z`$155<>Calc_Charts!`$B`$696:`$Z`$696))","=IF(B16=0,""OK"",""Review"")","Matrix headers should all follow set_groupList.","Calc_Charts"),
    @("Group count within matrix capacity","=MAX(0,COUNTA(set_groupList)-25)","=IF(B17=0,""OK"",""Review"")","Add matrix/chart capacity before adding more than 25 groups.","Settings vs Calc_Charts"),
    @("Summary rows near cap","=SUMPRODUCT(--(Summary!`$B`$5:`$B`$241<>""""))","=IF(B18>=220,""Review"",""OK"")","Review when classified/input rows approach the 237-row Summary capacity.","Summary"),
    @("Duplicate cause names in Settings","=SUMPRODUCT(--(Settings!`$A`$4:`$A`$59<>""""),--(COUNTIF(Settings!`$A`$4:`$A`$59,Settings!`$A`$4:`$A`$59)>1))","=IF(B19=0,""OK"",""Review"")","Never create two causes with the same name.","Settings")
  )
  for ($i = 0; $i -lt $checks.Count; $i++) {
    $r = 16 + $i
    for ($j = 0; $j -lt 5; $j++) {
      if ($j -eq 1 -or $j -eq 2) { $wsChecks.Cells.Item($r, $j + 1).Formula = $checks[$i][$j] }
      else { $wsChecks.Cells.Item($r, $j + 1).Value2 = $checks[$i][$j] }
    }
  }

  # README additions.
  Log-Step "README notes"
  $start = [Math]::Max(35, $wsReadme.UsedRange.Rows.Count + 2)
  $wsReadme.Cells.Item($start, 1).Value2 = "Phase 5 notes"
  $wsReadme.Cells.Item($start, 1).Font.Bold = $true
  $wsReadme.Cells.Item($start + 1, 1).Value2 = "To rename a root cause or group: 1) Edit the name in Settings (col A or B). 2) If a cause was renamed: Find & Replace the old name in Summary!E:E (Match entire cell contents). 3) Press F9, open Data Checks - Root causes not in Settings must be 0. Colors: the chart fill for a group/cause is static XML; after reordering Settings rows or changing hexes, the fills must be re-synced (Phase-4 note). Never create two causes with the same name."
  $wsReadme.Cells.Item($start + 2, 1).Value2 = "Legacy Graphs/Counts sheets remain hidden for package stability. Phase 5 ports Dashboard-facing formulas away from them."
  $wsReadme.Cells.Item($start + 3, 1).Value2 = "Pivots do not auto-refresh - right-click Refresh or use Data > Refresh All after adding rows."
  $wsReadme.Range("A" + ($start + 1) + ":C" + ($start + 3)).WrapText = $true

  # Part B1: Explore sheet with pivot, fallback formulas, and timeline block.
  Log-Step "Part B1 Explore"
  $wsExplore = Add-Or-Get-Sheet $wb "Explore" $true
  $wsExplore.Cells.Clear()
  $wsExplore.Range("A1").Value2 = "Explore"
  $wsExplore.Range("A2").Value2 = "Pivot and slicers are exploratory. Refresh after data changes."
  $wsExplore.Range("A1").Font.Bold = $true
  Set-Fill $wsExplore.Range("A1:N2") 15395562
  try {
    $pc = $wb.PivotCaches().Create(1, "Summary!R4C1:R241C14")
    $pt = $pc.CreatePivotTable($wsExplore.Range("A4"), "ptSummaryExplore")
    $pt.PivotFields("Group").Orientation = 1
    $pt.PivotFields("Group").Position = 1
    $pt.AddDataField($pt.PivotFields("Work Order"), "Count of WO", -4112) | Out-Null
    foreach ($fieldName in @("Group","Root Cause","Nomenclature","After Clean Point","State","Branch")) {
      try { $wb.SlicerCaches.Add2($pt, $fieldName).Slicers.Add($wsExplore, [System.Type]::Missing, "sl_$fieldName", $fieldName, 10 + (115 * [Array]::IndexOf(@("Group","Root Cause","Nomenclature","After Clean Point","State","Branch"), $fieldName)), 40, 110, 90) | Out-Null } catch {}
    }
  } catch {
    $wsExplore.Range("A4").Value2 = "Pivot creation failed on this Excel session; source range is Summary!A4:N241."
  }
  $wsExplore.Range("J3").Value2 = "Cause picker"
  $wsExplore.Range("K3").Formula = "=INDEX(set_causes,1)"
  $wsExplore.Range("K3").Validation.Delete()
  [void]$wsExplore.Range("K3").Validation.Add(3,1,1,"=set_causes")
  $wsExplore.Range("J5:L5").Value2 = @("Week Start","Calls","Intervention")
  for ($i = 0; $i -lt 25; $i++) {
    $r = 6 + $i
    $weekRow = 155 + $i
    $wsExplore.Cells.Item($r, 10).Formula = "=Calc_Weekly!A$weekRow"
    $wsExplore.Cells.Item($r, 11).Formula = "=COUNTIFS(Summary!`$E`$5:`$E`$241,`$K`$3,Summary!`$A`$5:`$A`$241,"">=""&J$r,Summary!`$A`$5:`$A`$241,""<""&J$r+7)"
    $wsExplore.Cells.Item($r, 12).Formula = "=IFERROR(INDEX(Interventions!`$C`$5:`$C`$200,MATCH(1,(Interventions!`$A`$5:`$A`$200=`$K`$3)*(Interventions!`$B`$5:`$B`$200>=J$r)*(Interventions!`$B`$5:`$B`$200<J$r+7),0)),"""")"
  }
  $co = $wsExplore.ChartObjects().Add(470, 120, 420, 250)
  $co.Chart.ChartType = 51
  $co.Chart.SetSourceData($wsExplore.Range("J5:K30"))
  $co.Chart.HasTitle = $true
  $co.Chart.ChartTitle.Text = "Weekly calls with interventions"

  # Part B2: Interventions table seeded from Settings clean points.
  Log-Step "Part B2 Interventions"
  $wsInterventions = Add-Or-Get-Sheet $wb "Interventions" $true
  $wsInterventions.Cells.Clear()
  $wsInterventions.Range("A1").Value2 = "Interventions"
  $wsInterventions.Range("A2").Value2 = "Additive analysis layer. Settings clean-point dates still drive before/after charts."
  $wsInterventions.Range("A4:C4").Value2 = @("Root Cause","Date","Label")
  $seedRow = 5
  for ($r = 4; $r -le 59; $r++) {
    $cause = $wsSettings.Cells.Item($r,1).Text
    $dt = $wsSettings.Cells.Item($r,5).Value2
    if ($cause -ne "" -and $dt -ne $null -and $dt -ne "") {
      $wsInterventions.Cells.Item($seedRow,1).Value2 = $cause
      $wsInterventions.Cells.Item($seedRow,2).Value2 = [double]$dt
      $wsInterventions.Cells.Item($seedRow,3).Value2 = "Clean point"
      $seedRow++
    }
  }
  $wsInterventions.Range("A5:A200").Validation.Delete()
  [void]$wsInterventions.Range("A5:A200").Validation.Add(3,1,1,"=set_causes")
  $wsInterventions.Range("B5:B200").NumberFormat = "yyyy-mm-dd"
  Set-Fill $wsInterventions.Range("A1:C4") 15395562
  Set-Border $wsInterventions.Range("A4:C200")

  # Part B3: Heatmaps.
  Log-Step "Part B3 Heatmaps"
  $wsHeat = Add-Or-Get-Sheet $wb "Heatmaps" $true
  $wsHeat.Cells.Clear()
  $wsHeat.Range("A1").Value2 = "Heatmaps"
  $wsHeat.Range("A2").Value2 = "Conditional-format heatmaps for cause/week and store/group patterns."
  $wsHeat.Range("A4").Value2 = "Cause x call-week"
  for ($i = 1; $i -le 25; $i++) { $wsHeat.Cells.Item(5 + $i, 1).Formula = "=IFERROR(INDEX(set_causes,$i),"""")" }
  for ($i = 0; $i -lt 20; $i++) {
    $c = 2 + $i
    $weekRow = 155 + $i
    $wsHeat.Cells.Item(5, $c).Formula = "=Calc_Weekly!A$weekRow"
    for ($r = 6; $r -le 30; $r++) {
      $wsHeat.Cells.Item($r, $c).Formula = "=IF(`$A$r="""","""",COUNTIFS(Summary!`$E`$5:`$E`$241,`$A$r,Summary!`$A`$5:`$A`$241,"">=""&B`$5,Summary!`$A`$5:`$A`$241,""<""&B`$5+7))"
    }
  }
  $rngHeat1 = $wsHeat.Range("B6:U30")
  $rngHeat1.NumberFormat = "0;;;"
  $rngHeat1.FormatConditions.Delete()
  $cf = $rngHeat1.FormatConditions.AddColorScale(3)
  $cf.ColorScaleCriteria.Item(1).FormatColor.Color = 16777215
  $cf.ColorScaleCriteria.Item(2).FormatColor.Color = 16041374
  $cf.ColorScaleCriteria.Item(3).FormatColor.Color = 9785112

  $wsHeat.Range("A34").Value2 = "Store x top groups"
  $wsHeat.Range("A35").Value2 = "Store"
  $wsHeat.Range("B35").Value2 = "Calls"
  for ($i = 1; $i -le 10; $i++) { $wsHeat.Cells.Item(35, 2 + $i).Formula = "=IFERROR(INDEX(set_groupList,$i),"""")" }
  $wsHeat.Cells.Item(35,13).Value2 = "Other"
  for ($i = 0; $i -lt 170; $i++) {
    $r = 36 + $i
    $storeRow = 4 + $i
    $wsHeat.Cells.Item($r,1).Formula = "='Store List'!A$storeRow"
    $wsHeat.Cells.Item($r,2).Formula = "=IF(A$r="""","""",COUNTIFS(Summary!`$C`$5:`$C`$241,A$r))"
    for ($j = 1; $j -le 10; $j++) {
      $c = 2 + $j
      $hdr = $wsHeat.Cells.Item(35,$c).Address($false,$true)
      $wsHeat.Cells.Item($r,$c).Formula = "=IF(OR(`$A$r="""",$hdr=""""),"""",COUNTIFS(Summary!`$C`$5:`$C`$241,`$A$r,Summary!`$L`$5:`$L`$241,$hdr))"
    }
    $wsHeat.Cells.Item($r,13).Formula = "=IF(`$A$r="""","""",MAX(0,`$B$r-SUM(C$r:L$r)))"
  }
  $rngHeat2 = $wsHeat.Range("C36:M205")
  $rngHeat2.NumberFormat = "0;;;"
  $rngHeat2.FormatConditions.Delete()
  $cf2 = $rngHeat2.FormatConditions.AddColorScale(3)
  $cf2.ColorScaleCriteria.Item(1).FormatColor.Color = 16777215
  $cf2.ColorScaleCriteria.Item(2).FormatColor.Color = 16041374
  $cf2.ColorScaleCriteria.Item(3).FormatColor.Color = 9785112
  $wsHeat.Activate()
  $wsHeat.Range("B6").Select()
  $xl.ActiveWindow.FreezePanes = $true

  # Part B4: cohort failure rate block and Dashboard chart.
  Log-Step "Part B4 Cohorts"
  $wsCharts.Range("AG40:AJ70").Clear()
  $wsCharts.Range("AG40:AJ40").Value2 = @("Commission week","Stores commissioned","Issues","Issues per store")
  for ($i = 0; $i -lt 25; $i++) {
    $r = 41 + $i
    $weekRow = 155 + $i
    $wsCharts.Cells.Item($r,33).Formula = "=Calc_Weekly!A$weekRow"
    $wsCharts.Cells.Item($r,34).Formula = "=COUNTIFS('Store List'!`$B`$4:`$B`$300,"">=""&AG$r,'Store List'!`$B`$4:`$B`$300,""<""&AG$r+7)"
    $wsCharts.Cells.Item($r,35).Formula = "=COUNTIFS(Summary!`$D`$5:`$D`$241,"">=""&AG$r,Summary!`$D`$5:`$D`$241,""<""&AG$r+7)"
    $wsCharts.Cells.Item($r,36).Formula = "=IF(AH$r=0,"""",AI$r/AH$r)"
  }
  $coCohort = $wsDash.ChartObjects().Add(25, 720, 520, 260)
  $coCohort.Chart.ChartType = 51
  $coCohort.Chart.SetSourceData($wsCharts.Range("AG40:AG65,AJ40:AJ65"))
  $coCohort.Chart.HasTitle = $true
  $coCohort.Chart.ChartTitle.Text = "Issues per store by commission cohort"
  $coCohort.Chart.Shapes.AddTextbox(1, 300, 10, 190, 36).TextFrame.Characters().Text = "older cohorts have had more exposure time"

  # Part B5: repeat offenders block and Dashboard table/chart.
  Log-Step "Part B5 Repeat offenders"
  $wsCharts.Range("AG75:AN140").Clear()
  $wsCharts.Range("AG75:AJ75").Value2 = @("Serial","Calls","Rank","Causes")
  $serialCounts = @{}
  $serialCauses = @{}
  for ($sr = 5; $sr -le 241; $sr++) {
    $serial = ([string]$wsSummary.Cells.Item($sr,8).Text).Trim()
    $cause = [string]$wsSummary.Cells.Item($sr,5).Text
    if ($serial -ne "") {
      if (-not $serialCounts.ContainsKey($serial)) {
        $serialCounts[$serial] = 0
        $serialCauses[$serial] = New-Object System.Collections.Generic.HashSet[string]
      }
      $serialCounts[$serial] = [int]$serialCounts[$serial] + 1
      if ($cause -ne "") { [void]$serialCauses[$serial].Add($cause) }
    }
  }
  $topSerials = $serialCounts.GetEnumerator() | Sort-Object @{Expression="Value";Descending=$true}, @{Expression="Key";Descending=$false} | Select-Object -First 60
  $outRow = 76
  foreach ($entry in $topSerials) {
    $wsCharts.Cells.Item($outRow,33).Value2 = $entry.Key
    $wsCharts.Cells.Item($outRow,34).Formula = "=IF(AG$outRow="""","""",COUNTIF(Summary!`$H`$5:`$H`$241,AG$outRow))"
    $wsCharts.Cells.Item($outRow,35).Formula = "=IF(AG$outRow="""","""",AH$outRow+(100-ROW(A1))/set_tiebreakDivisor)"
    $wsCharts.Cells.Item($outRow,36).Value2 = [string]::Join(", ", [string[]]$serialCauses[$entry.Key])
    $outRow++
  }
  $wsCharts.Range("AL75:AN79").Value2 = @("Metric","Units","Calls")
  $wsCharts.Range("AL76").Value2 = "1 call"
  $wsCharts.Range("AL77").Value2 = "2 calls"
  $wsCharts.Range("AL78").Value2 = "3+ calls"
  $wsCharts.Range("AM76").Formula = "=SUMPRODUCT(--(AH76:AH135=1))"
  $wsCharts.Range("AM77").Formula = "=SUMPRODUCT(--(AH76:AH135=2))"
  $wsCharts.Range("AM78").Formula = "=SUMPRODUCT(--(AH76:AH135>=3))"
  $wsCharts.Range("AN76").Formula = "=AM76"
  $wsCharts.Range("AN77").Formula = "=AM77*2"
  $wsCharts.Range("AN78").Formula = "=SUMPRODUCT((AH76:AH135>=3)*AH76:AH135)"
  $coRep = $wsDash.ChartObjects().Add(570, 720, 360, 260)
  $coRep.Chart.ChartType = 51
  $coRep.Chart.SetSourceData($wsCharts.Range("AL75:AM78"))
  $coRep.Chart.HasTitle = $true
  $coRep.Chart.ChartTitle.Text = "Repeat-offender serial distribution"

  # Final workbook settings.
  Log-Step "Final save"
  $wsDash.Range("B3").Value2 = [datetime]"2026-06-01"
  $wsDash.Range("B4").Value2 = [datetime]"2026-07-15"
  try { $xl.Calculation = -4105 } catch { Log-Step "Could not restore auto calc before save: $($_.Exception.Message)" }
  $wb.Save()
  Log-Step "Saved workbook"
  $wb.Close($true)
}
finally {
  if ($wb) { try { $wb.Close($false) } catch {} }
  $xl.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
