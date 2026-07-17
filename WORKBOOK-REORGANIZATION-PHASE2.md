# DG-template-reorganized.xlsx — Phase 2: Fix & Finish

Follow-up to `WORKBOOK-REORGANIZATION.md`. Phase 1 was implemented in
`outputs/019f6dfb-339c-7212-a7a5-679970584164/DG-template-reorganized.xlsx` and got the
layout right, but left two classes of problems:

- **BUG (user-visible): changing From/To does not update the charts.**
- **Incomplete: the charts still run on the old hidden engine; the new `Calc_Charts`
  layer is an unused mirror.**

This file tells the LLM exactly what to fix. Work on a copy. Same tooling warning as
phase 1: **openpyxl deletes charts on save** — use Excel/COM (xlwings) or direct XML
editing of the unzipped package only.

---

## Part A — Fix the frozen-charts bug (do this first)

**Root cause 1: calculation mode was left on manual.**
`xl/workbook.xml` now contains `<calcPr calcId="144525" calcMode="manual" forceFullCalc="1"/>`.
The original file had automatic calc. Nothing recalculates when the user edits a date,
so every chart shows stale cached values.

Fix: set `calcMode="auto"` (or simply remove the `calcMode` attribute) and keep
`fullCalcOnLoad="1"` so the first open recalculates everything:
`<calcPr calcId="144525" fullCalcOnLoad="1"/>`.
If working via COM: `app.calculation = "automatic"` before the final save, then save.

**Root cause 2: the From/To cells are formulas, not values.**
`Dashboard!B3` is stored as `<f>46174</f>` (a formula that returns the date serial) and
`B4` as `=46218`. They must be plain literal date values (46174 = 2026-06-01,
46218 = 2026-07-15) with a date number format, so a user can type a date or clear the
cell. Keep the rule: **blank = unbounded**. Add date data-validation that allows blanks.

**Verify Part A before moving on:** open in Excel, change To to Jun 30 → chart 1/3/5/6
bars and titles change immediately; clear both dates → titles read "All data".

## Part B — Finish the engine swap

### B1. Current state (measured, don't re-derive)

- Charts live on `Dashboard` (drawing1) but reference the OLD engine:

| Chart | Title ref | Cats ref | Vals refs | Extra series |
|---|---|---|---|---|
| 1 | `Graphs!$H$1` | `rc_cat1` | `rc_val1` (name cell `Graphs!$B$7`) | — |
| 2 | `Graphs!$H$3` | `Graphs!$AN$1:$AN$20` (hardcoded text) | `Calc_Weekly!$C$155:$V$174` per group | total labels `Graphs!$AP$1:$AP$20` |
| 3 | `Graphs!$H$2` | `rc_cat2` | `rc_val2` (name cell `Graphs!$F$7`) | — |
| 4 | `Graphs!$H$4` | `Graphs!$AO$1:$AO$20` (hardcoded text) | `Calc_Weekly!$C$106:$AA$125` per cause | total labels `Graphs!$AQ$1:$AQ$20` |
| 5 | `Graphs!$H$6` | `rc_cat5` | `rc_d5_1`…`rc_d5_25` (name cells `Graphs!$DI$40`…`$EG$40`), `rc_before5`, `rc_total5`, `rc_after5` | — |
| 6 | `'Case Nomenclature Graph'!$B$1` | `'Case Nomenclature Graph'!$A$2:$A$7` | `...!$B$2:$B$7` | — |

- New names `ch1_cats/ch1_vals`, `ch3_cats/ch3_vals`, `ch5_cats/ch5_after/ch5_before/
  ch5_total/ch5_seg01..25`, `ch6_cats/ch6_vals` already exist and point at `Calc_Charts`
  (T3 at rows 65+, T4 at 95+, T5 at 125+, T6 matrix anchored row 156, T9 at 244+).
  **They are currently used by nothing.**
- `Calc_Charts` cells are mirrors (`=Graphs!V8` style), NOT real logic. T7 (rows 184–206+)
  already has the formula-driven week labels (`=TEXT(Calc_Weekly!A155,"mm/dd")`).
- Old sheets `Graphs`, `Counts`, `Case Nomenclature Graph` are hidden but still the live
  engine; all `rc_*` names still exist.
- Summary data starts at **row 5** now (banner row added); all `$5:$241` references are
  already consistent — do not "fix" them back.

### B2. Port the real logic into Calc_Charts

Replace every mirror formula (`=Graphs!...`, `='Case Nomenclature Graph'!...`) on
`Calc_Charts` with the actual computation, referencing only: `Summary`, `Settings`,
`Calc_Weekly`, `Dashboard!B3/B4` (or `set_from`/`set_to`), and other `Calc_Charts` tables.
The logic to port, per table (source of truth is the old `Graphs` sheet):

- **T1 Cause totals** ← `Graphs!K8:N…`: cause list from `Settings!A`, two SUMPRODUCTs
  (window-filtered total & after-clean), epsilon column.
- **T2 Group totals** ← `Graphs!P8:U32` + dedup col `Graphs!AA4:AA80` (COUNTIF
  first-occurrence + SMALL); group SUMIFs over T1; two epsilon columns.
- **T3 / T4 Ranked lists** ← `Graphs!V:W` / `X:Y`: `INDEX(T2, MATCH(LARGE(epsilon,k)))`.
- **T5 Chart-5 stack** ← `Graphs!DD41:DG65`: ranked group, After (INDEX/MATCH into T2),
  Before = Total − After, Total.
- **T6 Diagonal matrix** ← `Graphs!DI40:EG65`: header row = dynamic names
  (`=IF(<T5 rank n cat>="","",<cat>)`), body = After value on the diagonal, 0 elsewhere.
- **T7** already correct (labels from Calc_Weekly); add the two total columns
  (`=SUM(Calc_Weekly!C155:V155)` per commission week, `=SUM(Calc_Weekly!C106:AA106)` per
  call week) if not already real formulas.
- **T9 Case nomenclature** ← the SEARCH-based SUMPRODUCT from
  `'Case Nomenclature Graph'!B2:B7`, with prefixes from `Settings` and dates from
  `set_from`/`set_to` (not `Graphs!B3/B4`).
- **T10 Chart titles** — create if missing: four cells building the dynamic titles from a
  local Period label (`=IF(AND(set_from="",set_to=""),"All data",...)` — port from
  `Graphs!B5`, but reference `set_from`/`set_to` directly). One row per chart, labeled.

After porting, `Calc_Charts` must have **zero references** to `Graphs`, `Counts`, or
`Case Nomenclature Graph` (grep the sheet XML for those names to prove it).

### B3. Re-point the six charts (XML edit of `xl/charts/chart*.xml`)

Swap references per the table in B1, changing **nothing else** (keep every `srgbClr`,
order, gap/overlap, data-label settings):

- Chart 1: title → T10 cell; cats `[0]!rc_cat1` → `[0]!ch1_cats`; vals `rc_val1` →
  `ch1_vals`; series-name cell → a label cell on Calc_Charts.
- Chart 3: same with `ch3_cats`/`ch3_vals` and its T10 cell.
- Chart 5: `rc_cat5→ch5_cats`, `rc_d5_N→ch5_segNN` (N 1–25, zero-padded), `rc_before5→
  ch5_before`, `rc_total5→ch5_total`, `rc_after5→ch5_after`; 25 series-name cells
  `Graphs!$DI$40..$EG$40` → T6 header row cells; title → T10.
- Charts 2 & 4: category refs `Graphs!$AN$1:$AN$20` / `$AO$1:$AO$20` → the T7 label
  columns (this is the week-label bug fix); total-label series `Graphs!$AP/AQ$1:$20` →
  T7 total columns; titles → T10 cells. Value refs into `Calc_Weekly` stay as they are.
- Chart 6: title/cats/vals → T9 cells (`ch6_cats`/`ch6_vals` exist; consider making them
  OFFSET-dynamic over the Settings prefix list, else keep fixed 6 rows).

Also update the cached `<c:strCache>`/`<c:numCache>` blocks or simply leave them — Excel
refreshes caches on first recalc (guaranteed by `fullCalcOnLoad`).

### B4. Delete the old engine (only after B5 passes once)

1. Delete defined names: all 33 `rc_*` names.
2. Delete sheets `Graphs`, `Counts`, `Case Nomenclature Graph` (this also removes the
   `Dashboard!B5` period mirror — make sure Dashboard's Period cell was re-pointed to the
   T10/local period formula first, and that `Graphs!B3/B4` pass-through cells are no
   longer referenced by anything).
3. Re-grep the entire package for `Graphs!`, `rc_`, `Counts!`, `Nomenclature` — zero hits
   allowed outside README prose.

### B5. Cleanup + verification checklist

- [ ] Part A works: date edits recalc instantly (automatic calc confirmed in
      Excel Options → Formulas).
- [ ] All six charts render pixel-identical to `previews/dashboard_contact_sheet.png`
      with From=Jun 1, To=Jul 15.
- [ ] Change To → Jun 30: charts 1/3/5/6 change, titles update. Clear both: "All data".
- [ ] Zero-window test: pick a window where a group has 0 → bar disappears (charts 1/3/5),
      chart-5 colors stay stable per rank.
- [ ] Add a Summary row in a brand-new week → charts 2/4 gain a column AND its axis label
      matches the computed week start (the original bug — must pass now).
- [ ] Name Manager: only `ch*` and `set_*` names, none broken.
- [ ] No `#REF!`/`#NAME?` after full recalc; Data Checks all OK/Info.
- [ ] Optional leftovers from phase 1, do if cheap: normalize duplicate category names
      ("Installation issue" vs "DG Installation issue", "Incorrect Program" vs
      "Wrong program") across Settings, Summary, and chart series names; make Data Checks
      expected counts formula-driven instead of hardcoded 138/25.
