# Phase 5 — Single source of truth + exploratory analytics

Applies to `outputs/019f6dfb-339c-7212-a7a5-679970584164/DG-template-reorganized-phase4.xlsx`
(work on a copy named `...-phase5.xlsx`). Read the audit in Part 0 before changing anything —
every location and count below was measured on the actual file, not assumed.

Tooling rules (unchanged from phases 2–4):
- Never round-trip through openpyxl (it deletes charts). Use Excel/COM (xlwings) first,
  surgical XML edits only as fallback.
- Open-test the workbook in Excel after each major step, not at the end.
- OneDrive has silently reverted saves in this folder before: after each save, close,
  reopen, and re-verify one changed cell before proceeding.
- Restore Dashboard From/To to 2026-06-01 / 2026-07-15 and calc mode = automatic before
  the final save.

---

## Part 0 — Audit: where the single-source-of-truth breaks today (measured)

The Settings sheet is *supposed* to be the source of truth for root causes and groups.
These are all the places that currently bypass it:

| # | Location | What's wrong | Blast radius when a name changes |
|---|---|---|---|
| 0.1 | `Calc_Charts` rows **65, 95, 155, 386, 541, 696** — six matrix header rows, ~21 columns each = **126 cells** | Group names are **typed text**, not formulas (user found `E65` = "Drain tube blocked"). The matrix body matches by name: `IF($A66=C$65, value, "")` | Rename a group in Settings → that group's bars **silently disappear** from charts 1, 3, 5, 7, 8, 9 (no error, just missing data) |
| 0.2 | `Calc_Weekly!C154:V154` — **19 cells** | Chart 2's series-name row AND its SUMIF criteria are typed group names | Rename a group → its weekly stack segment goes to zero in chart 2, silently |
| 0.3 | `Calc_Charts` rows **6–206**: **390 formulas still reference hidden legacy sheet `Graphs`, 100 reference hidden `Counts`** | The phase-2 "port the real logic" step was never actually completed for T1–T7. Charts 1/2/3/4/5 totals flow **through the hidden legacy engine** (Settings → Graphs/Counts → Calc_Charts mirrors → charts). Only T9 and the phase-3 blocks (rows ≥330, charts 7–9) are genuinely self-contained | Legacy sheets can never be deleted; any edit inside hidden Graphs breaks live charts; two parallel row-bound systems must stay in sync |
| 0.4 | `Summary` column E has **no dropdown validation at all** (the only DV in the sheet is a stray remnant on `B143:B146,B225:B241` — the Work Order column) | Root causes are free-typed | A typo creates a cause that matches nothing → the row **vanishes from every chart**. Data Check #6 ("Root causes not in Settings") is the only safety net |
| 0.5 | Charts 5, 7, 8, 9 XML: series name **"Before / at clean point" is a hardcoded literal** in each chart | Cosmetic label not driven by any cell | Can't retitle centrally; inconsistent if edited in one chart only |
| 0.6 | Inherent: the **cause name itself is the join key** everywhere (Summary E ↔ Settings A ↔ COUNTIFS). Renaming a cause in Settings orphans all existing Summary rows that carry the old name | By design (no stable IDs) | Must be handled by procedure, not formulas — see A4 |
| 0.7 | Bounds hardcoded by position: causes capped at `Settings!A4:A59`, matrices at 25 columns (21 groups used), weeks at 20 rows, Summary at row 241, colors keyed to Settings **row position** (insert/reorder rows in Settings = colors and mappings shift) | Structural caps | Acceptable if documented; add Data Checks (A7) so breaches surface loudly |

Answer to "how many things like this exist": **145 hardcoded name cells (0.1 + 0.2), 490
legacy mirror formulas (0.3), one missing validation (0.4), four hardcoded chart labels
(0.5)** — plus the two systemic issues 0.6/0.7 that need procedure + checks rather than formulas.

---

## Part A — Make Settings the real single source of truth

Do A1–A3 first (they fix the user-visible rename bug); A5 is the big one, do it once
charts are stable; A4/A6/A7 close it out.

### A1. Group Registry + formula-driven matrix headers (fixes 0.1)
1. Create **T0 "Group Registry"** at the top of `Calc_Charts` (insert above T1 or place in
   free columns; do NOT shift existing rows — charts and Data Checks reference absolute
   addresses. Safest: put T0 in unused columns to the right, e.g. `AA1:AB30`):
   - Column 1: unique group list in first-appearance order from `Settings!B4:B59`
     (the proven dedup pattern: `COUNTIF` first-occurrence flag + `SMALL` + `INDEX` —
     copy it from the legacy Graphs AA-column technique or rebuild per phase-2 B2).
   - Name the range `set_groupList` (OFFSET-dynamic over non-blank rows).
2. Replace all **126 typed header cells** with formulas indexing T0:
   `C65 = IFERROR(INDEX(set_groupList,1),"")`, `D65 = IFERROR(INDEX(set_groupList,2),"")`, …
   across all six header rows (65, 95, 155, 386, 541, 696). Column k = group k, so column
   order (and therefore each series' static color) still follows Settings order.
3. **Do not touch the matrix body formulas** — they already match dynamically by name.
4. Chart series *names* reference these header cells already (verify per chart; fix any
   that point elsewhere).

### A2. Calc_Weekly chart-2 header row (fixes 0.2)
Replace `Calc_Weekly!C154:V154` typed names with `=IFERROR(INDEX(set_groupList,k),"")`.
The SUMIF criteria in rows 155–174 reference row 154, so they follow automatically.
Chart 2 series fills stay keyed to columns — same color rule as A1.

### A3. Summary dropdown (fixes 0.4)
1. Add list data validation on `Summary!E5:E241`: source `=set_causes`
   (allow blank; show dropdown; error style Stop with message "Pick a root cause from
   Settings — add new causes there first").
2. Delete the stray remnant DV on `B143:B146,B225:B241`.
3. Existing typed values that already match Settings are unaffected; run Data Check #6
   after — it must stay 0.

### A4. Rename procedure (handles 0.6 — document, don't engineer)
Add to README, verbatim intent:
> **To rename a root cause or group:** 1) Edit the name in `Settings` (col A or B).
> 2) If a *cause* was renamed: Find & Replace the old name in `Summary!E:E`
> (Match entire cell contents). 3) Press F9, open Data Checks — "Root causes not in
> Settings" must be 0. Colors: the chart fill for a group/cause is static XML; after
> reordering Settings rows or changing hexes, the fills must be re-synced (Phase-4 note).
> Never create two causes with the same name.
Optionally add a Data Check: `COUNTIF` duplicate names in `Settings!A4:A59` and in the
group registry → Review if any.

### A5. Finish the engine port (fixes 0.3) — the big item
Goal: **zero formulas in Calc_Charts/Calc_Weekly referencing `Graphs` or `Counts`.**
The phase-3 blocks (T11–T13, rows ≥330) are the template — they already compute
cause totals → group rollup → ranked stack → matrix entirely from Summary + Settings.
Port the same shapes into:
- **T1** (cause totals, window-filtered) — replace `=Graphs!K/L/M` mirrors with the
  SUMPRODUCT pattern from T11's cause-totals section (minus the prefix factor).
- **T2** (group totals + epsilons) — SUMIFs over T1 keyed by `set_groupList`.
- **T3/T4** (ranked lists, rows 65+/95+) — `INDEX/MATCH(LARGE(epsilon,k))` over T2
  (T3 by total, T4 by after-clean).
- **T5** (chart-5 stack, rows 125+) — ranked group / after / before / total from T2.
- **T7 total columns** (`B187`/`C187` down = currently `=Graphs!AP1`/`AQ1` mirrors) —
  replace with `=SUM(Calc_Weekly!C155:V155)` and `=SUM(Calc_Weekly!C106:AA106)` row-wise.
- Also sweep `Counts!` references (100 formulas) the same way — their logic is simple
  SUMIFS over Summary/Settings.
Keep every cell address identical (in-place formula replacement only) so charts, names,
and Data Checks are untouched. After the sweep, verify with a full-sheet scan that
`Graphs!`/`Counts!` appear in **zero** formulas outside the legacy sheets themselves.
Keep the legacy sheets hidden (deleting corrupted the package in phase 2) — but after
A5 they are genuinely dead, and a README line should say so.

### A6. "Before / at clean point" label (fixes 0.5)
Add one label cell on Calc_Charts (e.g. next to T10 titles), text
`Before / at clean point`; point the series name of that series in charts 5/7/8/9 at it
(COM: `FullSeriesCollection(n).Name = "=Calc_Charts!$X$Y"`).

### A7. New Data Checks (guards 0.7)
- "Matrix headers match group registry": `SUMPRODUCT(--(C65:W65<>C95:W95))` style
  cross-row comparison + first header row vs registry → 0 or Review.
- "Group count within matrix capacity": `COUNTA(set_groupList) <= 25` → Review if over.
- "Summary rows near cap": count of used rows vs 237 capacity → Review at ≥ 220.
- "Duplicate cause names in Settings" (from A4).

### Part A verification
- [ ] Baseline export of all 9 charts (Jun 1–Jul 15) before starting; pixel-compare after.
- [ ] **The rename test (the user's original complaint):** change group "Drain tube
      blocked" → "Drain tube blocked X" in Settings col B. All matrix headers, chart 2
      series, and legends update; no bars disappear; Data Checks stay OK. Revert.
- [ ] Rename a cause per the A4 procedure end-to-end; verify charts and checks. Revert.
- [ ] Summary E shows the dropdown; typing an invalid cause is rejected.
- [ ] Zero `Graphs!`/`Counts!` refs outside legacy sheets (scripted scan, not eyeball).
- [ ] All four date-window scenarios from the phase-4 QA still match independent counts.

---

## Part B — New analytics (approved ideas: 8, 7, 6, 4-extended, 3, 2)

Build order matters: B0 unlocks everything else.

### B0. Enrich Summary (new formula columns, right of existing ones)
Add to `Summary` (headers row 4, formulas rows 5–241, same IFERROR-lookup style as col A):
- **I: Branch**, **J: City**, **K: State** — INDEX/MATCH from `Raw Data` by Work Order
  (Raw Data headers row 4: BRANCH / CITY / STATE columns; match on BOS WORK ORDER col H
  exactly as Summary!A does).
- **L: Group** — `=IFERROR(INDEX(Settings!$B$4:$B$59,MATCH($E5,set_causes,0)),"")`.
- **M: Nomenclature** — first matching prefix from Settings T-S3 over col H
  (nested SEARCH or a small LOOKUP over the prefix list; "other" if no match).
- **N: Age at failure (days)** — `=IF(OR($A5="",$D5=""),"",$A5-$D5)`.
- Column I currently holds a leftover header "Column1" — remove/overwrite it.
These columns are pure additions; nothing existing references I:N, so risk is minimal.
Verify one known WO manually against Raw Data.

### B1. "Explore" sheet — PivotTable + slicers (idea 8)
1. Define `Summary!A4:N241` as the pivot source (or convert to a real Table
   `tSummary` **only if** a full recalc + chart QA passes afterward — the workbook's
   SUMPRODUCTs use plain ranges and must keep working).
2. New visible sheet `Explore`, one PivotTable + one PivotChart (stacked column,
   default: rows = Group, values = count of WO).
3. Slicers: **Group, Root Cause, Nomenclature, After Clean Point, State, Branch** +
   a **Timeline** on Call Date. Arrange slicers in one row above the chart.
4. README: "Pivots don't auto-refresh — right-click → Refresh (or Data → Refresh All)
  after adding rows." (.xlsx cannot carry auto-refresh code; do not add macros.)
5. This sheet is for exploration; the Dashboard remains the presentation layer.

### B2. Interventions timeline (idea 4, extended to multiple clean points)
1. New visible sheet `Interventions` with table: **Root Cause | Date | Label**
   (DV on cause col = `set_causes`; date-validated col 2). Seed it with the existing
   single clean points from `Settings!E` (one row each) so day one isn't empty.
   Settings keeps its Clean-Point column — it still drives the after/before charts;
   this table is an *additive analysis layer* (state that in the sheet banner).
2. On `Explore` (or a small `Timeline` sheet): a **cause picker cell** (DV =
   `set_causes`) and a computed block: 25 week rows
   (`COUNTIFS(Summary E = picked cause, call date in week)`) reusing Calc_Weekly's
   week starts.
3. Chart: column chart of that block + intervention markers as an XY-scatter overlay
   (X = intervention date, Y = 0, vertical **error bar** up to the plot max, data label
   = the Label text, series color = status red `C00000`/`D03B3B`). Up to 6 markers;
   blank rows collapse via `#N/A`.
4. Title: `=<cause> & " — weekly calls with interventions (" & period & ")"`.
5. Verify: pick False Alarm → one marker at 2026-05-22 and visibly more calls after it
   than before (known result from QA data: 12 of 15 after).

### B3. Heatmaps (ideas 2 & 3) — conditional formatting, no chart objects
New visible sheet `Heatmaps`, two blocks, each with the standard banner:
1. **Cause × call-week**: rows = causes (formula-linked to `set_causes`), columns =
   the ~20 week starts (formula-linked to Calc_Weekly col A), body = COUNTIFS.
   3-color scale white→blue (use theme-consistent blues, e.g. min `FFFFFF` /
   mid `9EC5F4` / max `184F95`); zeros shown as blank via number format `0;;;`.
2. **Store × cause**: rows = stores that have ≥1 call. Practical approach: full store
   list from `Store List` with a call-count column, sorted/filtered by count desc
   (a helper rank + INDEX block, or an AutoFilter the user can sort — pick the simpler
   and document it), columns = top groups + "Other". Same color scale.
3. Freeze panes on both blocks; these sheets print/screenshot well as-is.

### B4. Cohort failure-rate chart (idea 6)
On `Calc_Charts`, a small block: per commission week (reuse Calc_Weekly week starts):
stores commissioned that week (COUNTIFS over `Store List!B`), issues attributed to
those stores' cases (COUNTIFS Summary commission-date in week), and
**rate = issues / stores commissioned** (blank if denominator 0).
Chart on Dashboard: column chart of rate with the n= denominator as a second-row
category label (like the sample), title "Issues per store by commission cohort".
Include the caveat as an on-chart text note: "older cohorts have had more exposure time."

### B5. Repeat-offender block (idea 7)
1. Calc_Charts block: normalized serial (`TRIM`), calls-per-serial (COUNTIF), then:
   distribution row (units with 1 / 2 / 3+ calls) and a **top-repeats table**
   (serial, call count, cause list) for count ≥ 2, sorted desc — INDEX/MATCH ranking
   like T3.
2. Dashboard: small column chart of the distribution + the top-repeats table pasted as
   a formatted range next to it (a table is more actionable than a chart here).
3. Data note: serials are free text from Raw Data (one row contains two serials
   comma-joined). Normalize with TRIM only; do not attempt to split multi-serial cells —
   count them as their own unit and note this in the block banner.

### Part B verification
- [ ] B0: spot-check 3 WOs' branch/city/state/group/nomenclature/age against Raw Data.
- [ ] B1: slicer combination "RLN3MA + After = Yes" shows 9 work orders (known value
      from phase-4 QA data).
- [ ] B2: False Alarm picker → marker at May 22, counts match Calc_Weekly's False Alarm
      row; picking a cause with no interventions shows no markers, no errors.
- [ ] B3: heatmap grand totals equal 157 (all-data) and match Chart 1's totals per cause.
- [ ] B4: rates recompute when Store List gains a row; blank (not #DIV/0!) where a week
      commissioned 0 stores.
- [ ] B5: distribution sums to distinct serial count; total calls in the block ≈ 157
      minus blank-serial rows.
- [ ] Full regression: the 9 existing charts unchanged across the 4 QA date windows;
      all Data Checks OK/Info; workbook opens cleanly after every step.

## Suggested execution order
A1 → A2 → A3 (rename bug fixed, low risk) → checkpoint → A5 (engine port, biggest risk,
verify hard) → A4/A6/A7 → checkpoint → B0 → B1 → B2 → B3 → B4 → B5 → full regression.
