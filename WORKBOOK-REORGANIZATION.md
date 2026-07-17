# DG-template.xlsx — Reorganization Instructions

Instructions for an LLM (or a careful human) to reorganize `DG-template.xlsx` so it is
maintainable, self-documenting, and easy to extend — **without changing any numbers or
chart behavior**. Read the whole file before touching the workbook.

---

## 1. What this workbook does (do not break any of this)

Six bar charts driven by a service-call log:

| Chart | Title cell | What it shows | Dynamic behaviors |
|---|---|---|---|
| 1 | `Graphs!H1` | Issues by root-cause **group**, sorted desc | timeline filter, zero bars hidden, dynamic title |
| 2 | `Graphs!H3` | Stacked issues per **store commission week**, by group | fixed colors per group, week-total label series |
| 3 | `Graphs!H2` | After-clean-point issues by group, sorted desc | timeline filter, zero bars hidden, dynamic title |
| 4 | `Graphs!H4` | Stacked issues per **call week**, by root cause | fixed colors per cause, week-total label series |
| 5 | `Graphs!H6` | After vs before clean point per group, stacked, sorted desc | timeline filter, zero bars hidden, fixed color per bar rank (diagonal-matrix trick), total-label series |
| 6 | `'Case Nomenclature Graph'` | Issues by case serial prefix (RLN2MA…RMN5MA) | timeline filter |

The four mechanisms that make them dynamic (all must survive the reorganization):

1. **Timeline filter** — `Graphs!B3` (From) and `Graphs!B4` (To); blank = unbounded.
   Every filtered count uses the pattern
   `(dates >= IF(From="",0,From)) * (dates < IF(To="",2958465,To+1))`.
   Chart titles are formula cells concatenating the period text, referenced by charts as `StrRef`.
2. **Sort-desc with tiebreak** — helper column adds a descending epsilon
   (`value + (N-k)/100000`) so `LARGE`/`MATCH` never tie; `INDEX(list, MATCH(LARGE(...)))`
   produces the ranked list.
3. **Zero-bar hiding** — workbook-level defined names like
   `=OFFSET(anchor,0,0,MAX(1,COUNTIF(totals,">0")),1)`. Because lists are sorted desc,
   zeros sink to the bottom and OFFSET truncates them. Charts reference these names, not ranges.
4. **Stable colors** — every root cause / group is its own chart series with a hardcoded
   `srgbClr` fill in the chart XML. Chart 5 additionally uses a **diagonal matrix**: series
   *n* holds the "after" value only at category row *n* (0 elsewhere), so each bar position
   keeps a fixed color while series names update dynamically to the group ranked there.

## 2. What is wrong with the current layout

- Helper math is scattered in unlabeled grid islands: sort machinery at `Graphs!K8:Y32`,
  dedup helper at `Graphs!AA4:AA80`, chart-5 matrix at `Graphs!DD40:EG65`, label/total
  columns at `Graphs!AN1:AQ20`. Nothing says what feeds which chart.
- `By Commission Week` holds **three** stacked pivot blocks (rows 4–100, 104–125, 153–174)
  with no visual separation or explanation.
- Chart 2 & 4 x-axis labels (`Graphs!AN1:AN20`, `AO1:AO20`) are **hardcoded text** while the
  underlying week rows are computed from `MIN(Summary dates)` — they will silently desync.
- Series colors exist only inside chart XML; there is no visible color legend/source of truth.
- Hard caps are invisible: Summary row 240, Raw Data row 400, Settings row 80, 20 weeks,
  25 causes, 25 rank rows.
- `Graphs!H5` is an orphan duplicate of `H6`; `H6` contains a stray pasted string "RLN3MA".
- Two names for the same concept in different charts ("Installation issue" vs
  "DG Installation issue", "Incorrect Program" vs "Wrong program"); chart 5 has four series
  all named "Doors" (`ED40:EG40` fall back to blank ranks).

## 3. Target sheet structure

Rebuild into these sheets, in this tab order. Keep original sheet names where users
already know them; new sheets are marked. Color the tabs: inputs = blue, outputs = green,
engine = gray (and hide the engine sheets once verified).

| # | Sheet | Role | Visibility |
|---|---|---|---|
| 1 | `README` *(new)* | How the workbook works, update workflow, limits table | visible |
| 2 | `Dashboard` *(new, replaces top of Graphs)* | From/To controls, period label, data-check status lights, all 6 charts | visible |
| 3 | `Summary` | The one user-edited log (unchanged columns) | visible |
| 4 | `Raw Data` | CRM paste target (unchanged) | visible |
| 5 | `Settings` *(rename of Root Cause Settings)* | Root causes, groups, clean points, **plus the color map** | visible |
| 6 | `Store List` | unchanged | visible |
| 7 | `Calc_Weekly` *(rename of By Commission Week)* | the three week pivots, restructured | hidden |
| 8 | `Calc_Charts` *(new, replaces helper areas of Graphs + Counts)* | all chart-feed tables | hidden |
| 9 | `Data Checks` | checks, expanded | visible |

Delete the old `Graphs`, `Counts`, and `Case Nomenclature Graph` sheets **only after**
charts are re-pointed and values verified (Counts' group rollup moves to `Calc_Charts`;
the case-nomenclature table moves there too; its chart moves to `Dashboard`).

## 4. Layout rules (apply to every sheet)

These rules are the point of the exercise. Follow them everywhere:

1. **One table = one block.** Every table starts with a 2-row header banner:
   - Row A: `T<n>. <Table name>` in bold with a fill color, e.g. `T3. Group totals (filtered, sorted)`.
   - Row B: one sentence: source → transformation → consumer, e.g.
     `From T2 · sorts groups by total desc · feeds Chart 1 categories/values via names c1_cats/c1_vals`.
2. **Stack tables vertically in column A**, separated by exactly 2 blank rows. Never place
   tables side-by-side in far-right columns. (Exception: `Calc_Weekly` pivots are wide by
   nature; still stack them vertically with banners.)
3. **Every column has a header cell.** No formula column without a name.
4. **Cell coloring convention** (add a legend to README): blue text = user input;
   black = formula (never edit); gray fill = helper/epsilon columns; yellow fill = the two
   timeline cells on Dashboard.
5. **All magic numbers become labeled cells** on `Settings` (see §6): max rows, week count,
   epsilon divisor. Formulas reference the labeled cell or a named constant, not `240`/`400`.
6. **Named ranges use a scheme**: `c<chart#>_<role>` for chart feeds (`c1_cats`, `c1_vals`,
   `c5_before`, `c5_total`, `c5_seg01`…`c5_seg25`), `set_*` for settings
   (`set_from`, `set_to`, `set_causes`, `set_groups`, `set_colors`). Delete the old
   `rc_*` names after re-pointing.

## 5. Sheet-by-sheet specification

### 5.1 README (new)
Plain text: purpose, the weekly update workflow (paste CRM rows → classify in Summary →
check Data Checks → charts update), the color/format legend, the limits table
(max WOs, max causes, max weeks — with the cell on Settings that controls each), and a
one-line description of every calc table (T1…Tn) with what chart it feeds.

### 5.2 Dashboard (new)
- Top-left control block: `From` / `To` date cells (yellow, data-validated as dates,
  named `set_from`/`set_to`), computed `Period` label cell, and a mirror of the Data
  Checks status column (e.g. `=COUNTIF('Data Checks'!C:C,"Review")` with an OK/Review light).
- All 6 charts arranged in a grid below, each preceded by nothing — titles live in the
  charts themselves (linked to title cells on `Calc_Charts`).
- No helper formulas on this sheet other than the control block.

### 5.3 Summary — keep as is
Keep columns A–H exactly (formulas in A, C, D, G, H; user enters B, E, F). Only changes:
- Convert to an Excel Table named `tSummary` **only if** you verify the OFFSET/SUMPRODUCT
  formulas still work against it; otherwise keep plain ranges but bound them with a single
  named constant `set_maxSummaryRow` instead of the literal 240.
- Add a thin banner row explaining which columns to edit.

### 5.4 Settings (rename of Root Cause Settings)
Three stacked tables:
- **T-S1 Root causes** — existing columns (Root Cause, Group, Owner, Status, Clean-Point
  Date, Clean-Point Serial, Display Order) **plus two new columns**: `Cause Color (hex)`
  and `Group Color (hex)`. Populate from the current chart XML fills so recoloring is a
  visible, documented act (current map, group-level: Ref. leak/Low charge `70AD47`,
  Doors `FFC000`, Drain tube blocked `4472C4`, Wrong/Loose wiring `ED7D31`, Communication
  problem `5B9BD5`, High Store RH `A5A5A5`, Enclosure Seal `264478`, Component Failure
  `9E480E`, Defrost Sync `636363`, False Alarm `997300`, Incorrect Program `255E91`,
  Overstocking `43682B`, Case End Seal `698ED0`, Return Air blocked `F1975A`,
  Unidentified air leak `B7B7B7`, Freedom `1F4E79`, Heavy case load `C00000`,
  DG Installation issue `7030A0`, Unknown `2E75B6`, Program change - Hussmann `BF9000`;
  cause-level fills for chart 4 are in `chart4.xml`). Also fill each cell with its own
  color so the legend is visual.
- **T-S2 Constants** — labeled cells: `Max Summary rows` (240), `Max Raw rows` (400),
  `Max causes` (80), `Weeks shown` (20), `Tiebreak divisor` (100000). Name each.
- **T-S3 Serial prefixes** — the RLN2MA…RMN5MA list (moved from Case Nomenclature Graph
  col A), so chart 6's categories are settings-driven.

While migrating, **normalize the duplicate labels**: pick one canonical name for
"Installation issue"/"DG Installation issue" and "Incorrect Program"/"Wrong program",
update Settings, Summary values, and chart series accordingly.

### 5.5 Calc_Charts (new, hidden) — the engine
All tables stacked vertically, each with its banner. Suggested order:

- **T1 Cause totals (filtered)** — from old `Graphs!K8:N…`: per root cause, Total-in-window
  and After-clean-in-window (the two SUMPRODUCTs), + gray epsilon column.
- **T2 Group totals (filtered)** — from old `Graphs!P8:U32` + dedup helper `AA`: unique
  group list (keep the COUNTIF-first-occurrence + SMALL technique, but put the helper
  column *inside* this table with a header like `First-occurrence row (dedup helper)`),
  group Total, group After, two epsilon columns.
- **T3 Ranked by total** — old `V/W`: `INDEX/MATCH(LARGE(...))` over T2. Feeds `c1_cats`/`c1_vals`.
- **T4 Ranked by after-clean** — old `X/Y`. Feeds `c3_cats`/`c3_vals`.
- **T5 Chart-5 stack feed** — old `DD41:DG65`: ranked group, After, Before(=Total−After),
  Total. Feeds `c5_cats`, `c5_after`, `c5_before`, `c5_total`.
- **T6 Chart-5 diagonal matrix** — old `DI40:EG65`, kept verbatim (dynamic name row on top,
  25×25 diagonal below). Banner must explain the trick: *"Series n = After value only at
  rank n; gives each bar a fixed color while names follow the ranking. Feeds c5_seg01…c5_seg25."*
- **T7 Week labels + totals for charts 2 & 4** — replaces `Graphs!AN1:AQ20`.
  **Fix the hardcoded labels**: derive them from the week-start dates on `Calc_Weekly`
  (`=TEXT('Calc_Weekly'!A<row>,"mm/dd")`), never as typed text. Total columns = SUM of the
  matching `Calc_Weekly` row (these feed the dark-gray label series).
- **T8 Group rollup (from old Counts)** — Total/Visible/After per cause and per group.
- **T9 Case nomenclature counts** — prefix list from Settings T-S3 + the SEARCH-based
  SUMPRODUCT (keep as normal formula; it does not need Ctrl+Shift+Enter in modern Excel).
- **T10 Chart titles** — the four dynamic title formulas (old `H1,H2,H4,H6`), one per row
  with a label column. Rebuild the chart-5 title cleanly (drop the stray "RLN3MA", drop
  the orphan duplicate).

### 5.6 Calc_Weekly (rename of By Commission Week, hidden)
Same three pivots, restructured with banners:
- **T-W1 Commission week × root cause** (wide COUNTIFS matrix; headers pulled from Settings).
- **T-W2 Call week × root cause** (same shape, keyed on call date).
- **T-W3 Commission week × group** (SUMIF rollup of T-W1). Feeds chart 2 series.
Chart 4 series feed from T-W2. Keep the `MIN(date)-WEEKDAY(...)+1` week-start logic; drive
the number of week rows from the `Weeks shown` constant if practical, else document the cap
in the banner.

### 5.7 Data Checks
Keep existing checks but make expected counts formula-driven (e.g. compare Summary causes
against Settings dynamically) instead of hardcoded `138`/`25` from the original migration.
Add two new checks: *week-label desync* (T7 label = TEXT of Calc_Weekly week start) and
*causes missing a color* (blank hex in Settings T-S1).

## 6. Chart re-pointing rules

Charts keep their exact visual design; only their references move.

1. Create all new defined names first (§4.6 scheme), pointing at the new table locations,
   with the same `OFFSET(...,MAX(1,COUNTIF(...,">0")),1)` bodies.
2. Edit each series (Select Data or XML) to swap old refs → new names / new sheet ranges.
   Keep every `srgbClr` fill exactly as is (or re-derive from the Settings color map).
3. Title `StrRef`s → the T10 cells on `Calc_Charts`.
4. Chart 2/4 category refs → the T7 formula-driven label column.
5. Only after all 6 charts render identically: delete old names (`rc_*`), then old sheets.

Tooling note: openpyxl **destroys existing charts on save**. Do this reorganization either
(a) inside Excel itself (scripted via `xlwings`/COM or by hand following this spec), or
(b) by unzipping the `.xlsx` and editing `xl/charts/chart*.xml`, sheet XML, and
`workbook.xml` (defined names) directly, then re-zipping. Do **not** round-trip the file
through openpyxl.

## 7. Verification checklist (must pass before deleting anything old)

- [ ] With From/To = Jun 1 / Jul 15, every chart shows identical bars, order, colors,
      labels, and titles as a screenshot/copy of the original file.
- [ ] Blank both dates → titles say "All data", counts grow accordingly.
- [ ] Set a window where some group = 0 → its bar disappears from charts 1, 3, 5 (no gap,
      no zero-height bar), and colors of remaining chart-5 bars stay stable per rank.
- [ ] Add a fake Summary row in a new week → charts 2/4 pick it up **and the axis label
      matches the computed week start** (this is the bug fix — the old file fails this).
- [ ] Data Checks all OK; no `#REF!`/`#NAME?` anywhere (recalculate the whole book).
- [ ] All `rc_*` names deleted; every chart series resolves (Formulas → Name Manager shows
      no broken refs).
