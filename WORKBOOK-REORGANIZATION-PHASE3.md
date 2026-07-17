# Phase 3 — Per-nomenclature "after vs before clean point" charts

Follow-up to `WORKBOOK-REORGANIZATION.md` and `WORKBOOK-REORGANIZATION-PHASE2.md`.
Work on the phase-2 output workbook (`DG-template-reorganized-phase2.xlsx`).

Scope decision: Phase 3 is split. **Phase 3A (this file, do now)** adds Charts 7–9 in the
existing Chart-5 style. **Phase 3B (color identity normalization) is deferred to Phase 4**
— see the last section; do NOT restyle existing charts in this phase. This keeps
Charts 1–6 bit-untouched and the verification surface small.

Tooling rules (hard-learned):
- Never round-trip through openpyxl (deletes charts).
- **Prefer Excel/COM (xlwings): copy the Chart 5 object three times on the Dashboard,
  then re-point each copy's series formulas via the COM API.** Fall back to direct XML
  edits only for what COM cannot express, and open-test the workbook in Excel **after
  each chart is added**, not after all three. (Phase 2 corruption came from package-level
  edits done in bulk.)

## Goal

Three new charts — **Chart 7 (RLN3MA), Chart 8 (RLN2MA), Chart 9 (RMN4MA)** — each a
clone of Chart 5 *as it currently exists* (stacked column, outlined total bar with the
colored "after" segment inside, rank-keyed colors, sorted desc, zero bars hidden,
timeline-filtered, dynamic title), restricted to issues whose **case serial
(Summary col H) contains that nomenclature prefix**.

Blocks must be structurally identical so a fourth prefix can be added by copying one
block + one chart. Each block reads its prefix from a labeled cell — never hardcode the
prefix inside formulas.

## Step 0 — Fix the Settings prefix list (data bug, do first)

`Settings` T-S3 currently lists: RLN2MA, RLN3MA, RLN4MA, **RMN2MA**, RMN3MA, RMN5MA.
The original workbook's list was: RLN2MA, RLN3MA, RLN4MA, **RMN3MA, RMN4MA**, RMN5MA —
there was never an RMN2MA. Phase 1 almost certainly transcribed RMN4MA as RMN2MA.

Fix: **replace RMN2MA with RMN4MA** (do not just append — verify first that
`SUMPRODUCT(--ISNUMBER(SEARCH("RMN2MA",Summary!$H$5:$H$241)))` = 0 and the same for
RMN4MA is ~14; if RMN2MA truly has matches, add RMN4MA instead and flag it). Then confirm
Chart 6 / T9 picks up the corrected list and shows six nonzero-capable bars matching the
original chart (RMN4MA ≈ 14 for Jun 1–Jul 15).

## Phase-2 state you inherit (verified facts)

- Charts 1–6 live on `Dashboard`, referencing only `Calc_Charts` / `Calc_Weekly` /
  `ch*` names; all `rc_*` names gone.
- `Graphs`, `Counts`, `Case Nomenclature Graph` remain **hidden** as legacy fallbacks —
  deleting them corrupted the package in phase 2. Keep them; reference nothing in them.
- `Calc_Charts`: T1 cause totals, T2 group totals, T3/T4 ranked, T5 chart-5 stack,
  T6 diagonal matrix, T7 week labels, T9 nomenclature counts, T10 titles.
  `ch5_cats/after/before/total/seg01..25` show the name pattern to copy.
- Summary data rows `$5:$241`; A = call date, E = root cause, G = "Yes"/"No" after-clean,
  H = case serial text. Timeline: `set_from`/`set_to`, blank = unbounded.
- Chart 5 formatting to inherit via cloning: stacked barChart, overlap 100, gapWidth 30,
  25 rank-colored "after" series, "Before / at clean point" outline series, gray (404040)
  total-label series, after-label series.

## Step 1 — Calc_Charts: three new blocks (T11, T12, T13)

Append below the last table, stacked vertically, standard 2-row banner each, e.g.:

> **T11. RLN3MA — after vs before by group**
> Prefix from cell below · filters Summary by serial prefix + timeline · feeds Chart 7 via ch7_*.

Each block, top to bottom (write once, copy 3×):

1. **Prefix cell** — labeled `Nomenclature:`, value `=Settings!<T-S3 cell>` (link, not
   typed text). Name it `ch7_prefix` (ch8_/ch9_ likewise).
2. **Cause totals (prefix-filtered)** — one row per root cause (from `Settings!A4:A59`):
   - `Total` = `SUMPRODUCT((Summary!$E$5:$E$241=<cause>) * ISNUMBER(SEARCH(<prefix>,Summary!$H$5:$H$241)) * ISNUMBER(Summary!$A$5:$A$241) * (Summary!$A$5:$A$241>=IF(set_from="",0,set_from)) * (Summary!$A$5:$A$241<IF(set_to="",2958465,set_to+1)))`
   - `After` = same × `(Summary!$G$5:$G$241="Yes")`.
   (Blank serials never match — correct behavior.)
3. **Group totals** — unique group list (reuse T2's technique), Total/After SUMIFs over
   the cause rows, plus epsilon column `Total + (25-k)/set_tiebreakDivisor`.
4. **Ranked stack feed** (25 rows, mirrors T5): ranked group via
   `INDEX/MATCH(LARGE(epsilon,k))`; After; Before = Total − After; Total.
5. **Diagonal matrix** (mirrors T6, current rank-keyed style): header row of dynamic
   ranked names, 25×25 body with the After value on the diagonal, 0 elsewhere.
   (Phase 4 will convert T6 and T11–T13 matrices to group-keyed together.)
6. **Title cell** (extend T10): `="RLN3MA — issues by root cause, after vs before clean point ("&<period text>&")"`.
   Keep titles short; Chart 5's title box currently overflows its plot — size the new
   charts' title font the same or smaller, and optionally shrink Chart 5's while here.

## Step 2 — Defined names

Per chart n ∈ {7,8,9}: `chN_cats`, `chN_after`, `chN_before`, `chN_total`,
`chN_seg01`…`chN_seg25`, same OFFSET body as the ch5 family:
`=OFFSET(<anchor>,0,0,MAX(1,COUNTIF(<total 25-row range>,">0")),1)` — zero-bar hiding
comes free (a nomenclature with few affected groups shows few bars).

## Step 3 — The charts (COM-first workflow)

1. Via COM: `Dashboard.ChartObjects("<chart5>").Duplicate`/Copy-Paste three times; place
   below the existing six, 2-per-row grid, Chart-5 size; order 7 RLN3MA, 8 RLN2MA,
   9 RMN4MA.
2. For each copy, re-point via COM (`FullSeriesCollection(i).Formula` swap): `ch5_*` →
   `chN_*`; 25 series-name cells → the new block's matrix header row; title link → the
   new T10 cell. Keep every fill/format as inherited from the copy.
3. **Save and open-test in Excel after each chart** before starting the next.
4. Only if COM cannot set something (e.g. a name-based ref it mangles): unzip and edit
   that one chartN.xml surgically, then open-test again.

## Step 4 — Documentation & checks

- README: add Charts 7–9 to the chart list; document the copy-a-block recipe for adding
  another nomenclature; note the RMN2MA→RMN4MA correction.
- Data Checks: one row per prefix — `SUM(<block total column>)` must equal the T9 count
  for that prefix (both window-filtered); status Review otherwise.

## Verification checklist

- [ ] Step 0: Settings T-S3 = the six original prefixes incl. RMN4MA; Chart 6 matches the
      original (RLN2MA 21, RLN3MA 32, RMN4MA 14 for Jun 1–Jul 15).
- [ ] With From=Jun 1 / To=Jul 15: Chart 7 totals sum to 32, Chart 8 to 21, Chart 9 to 14
      (equal to their Chart 6 bars). Re-check after changing the window.
- [ ] Each new chart: bars sorted desc, zero groups absent, "after" segment colored with
      count label inside, total outline + total label above — same styling as Chart 5.
- [ ] Change To → Jun 30: charts 7–9 retitle and re-rank; clear both dates → "All data".
- [ ] Titles fit inside the chart area.
- [ ] Legacy-sheet rule (softened): hidden sheets `Graphs`/`Counts`/`Case Nomenclature
      Graph` may exist, but **no chart series, defined name, Dashboard formula, or any
      new T11–T13 formula references them** (check Name Manager + the new blocks + the
      nine charts' series formulas).
- [ ] Charts 1–6 byte-untouched except Chart 6's corrected prefix data (and Chart 5's
      title box only if you opted to fix it); workbook opens cleanly at every checkpoint.

## Phase 4 (deferred) — Color identity normalization

Agreed principle, deliberately not done now: one group = one color everywhere, driven by
`Settings!set_colors` (`H4:I59`). When scheduled, in one pass: convert T6 + T11–T13
matrices from rank-keyed (dynamic ranked headers, color follows bar slot) to group-keyed
(static group headers in Settings order; cell r,g = After if ranked-group-at-r = g else
0), re-fill the 25 series of Charts 5/7/8/9 with canonical hexes; apply the same matrix
technique to Charts 1/3 (new `ch1_seg*`/`ch3_seg*` names); re-sync Charts 2/4 fills to
the map.

Charts 2 & 4 are NOT exempt ("time-series" is a misread): their stacked segments ARE the
root-cause groups (Chart 2) and root causes (Chart 4) — Chart 2 vs Chart 1 is where the
color drift is most visible, so Chart 2 must sync to the group colors (`set_colors` col I).
Chart 4 is cause-granularity: use **cause-level colors** (`set_colors` col H) so causes
within a group stay distinguishable; document that chart 4 is the one cause-level chart.
Consistency rule: every chart matches the map at its own granularity (group charts ↔
col I, the cause chart ↔ col H).

Legend note: after conversion, Charts 5/7/8/9 carry 25 static group series (including
zero-data groups) — either hide those legends (axis labels already name the bars) or
accept the longer legend; decide once, apply to all four.

Expect all charts' colors to change — screenshot before/after and verify values, order,
and shapes are identical. Also note in README **next to the color map** that chart fills
are static XML: editing a hex in Settings requires re-running the sync; it will never
repaint automatically.
