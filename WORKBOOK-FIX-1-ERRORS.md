# Fix 1 — Repair the original workbook's defects (no reorganization)

**Start file: `DG-template.xlsx` in the project root — the ORIGINAL workbook.**
Ignore every `DG-template-reorganized-*.xlsx` under `outputs/` and every
`WORKBOOK-REORGANIZATION-*.md`; that lineage is abandoned. Never modify the original
in place: copy it to `DG-template-fix1.xlsx` (project root) and work on the copy.
The follow-up `WORKBOOK-IMPL-2-HEATMAP-DASHBOARD.md` continues from `DG-template-fix1.xlsx`.

Scope: repair the defects below. Do NOT restructure sheets, rename tabs, move helper
blocks, or re-point charts beyond what each fix requires. The workbook works today —
six charts on the `Graphs` sheet with timeline filtering (`Graphs!B3/B4`), sorted
bars, zero-bar hiding via `rc_*` OFFSET names, and fixed series colors. Every fix must
leave all of that pixel-identical except for the specific defect it repairs.

## Part A — Working rules (each exists because a prior attempt failed without it)

1. **Local copy workflow.** OneDrive silently reverted mid-session saves in this
   project at least three times. Copy the workbook to `%TEMP%`, do all work there,
   copy back once at the end (kill stray `EXCEL` processes first; retry the copy in a
   loop if locked). After copy-back, reopen from the real path and re-verify one item
   per fix.
2. **Render and inspect.** After each fix, export the affected chart(s) via
   `Chart.Export` as PNG and LOOK at the image. A previous build shipped an unreadable
   deliverable because no one ever rendered it. Excel must be `Visible = $true` when
   exporting or the PNGs come out 0-byte.
3. **Tooling:** use Excel/COM (never openpyxl — it deletes charts on save). This
   machine's COM surface: no `Range.Formula2` (use `.Formula`/`.FormulaArray`), no
   `SlicerCaches.Add2`, `ChartObjects().Add` args must be `[double]`, PowerShell
   `foreach` over COM parameterized collections yields the property not items (use
   indexed access `X(1)`). Wrap independent fixes in try/catch so one abort doesn't
   kill the rest — and check the transcript afterwards: a catch that fired is a fix
   that didn't happen.
4. **After writing any cross-sheet formula via COM, read it back** and assert it does
   not contain `[1]` (COM sometimes rewrites references as external-workbook links).
5. Original workbook conventions (verified): Summary data rows `$4:$240` (headers
   row 3; A=call date, B=WO, D=commission date, E=root cause, G=after-clean flag,
   H=case serial). Settings sheet is named `Root Cause Settings` (causes `A4:A80`,
   groups B, clean-point dates E). Weekly pivots live on `By Commission Week`:
   commission weeks `A6:A25`, call weeks `A106:A125`, commission-by-group block rows
   `154:174`. Charts and their helper blocks live on `Graphs`. Calculation is
   automatic. Before final save: `Graphs!B3` = 2026-06-01, `B4` = 2026-07-15.

## Part B — The defects to fix

### B1. Hardcoded week labels on charts 2 and 4 (silent-desync bug — the important one)
`Graphs!AN1:AN20` and `AO1:AO20` are **typed text** ("03/02"…"07/13") used as the
category axes of the two weekly stacked charts, while the actual week rows on
`By Commission Week` are computed from `MIN()` of the data. New data in a new week
shifts the computed weeks but not these labels — the axis silently lies.

Fix: replace each with a formula deriving from the real week starts:
- `AN1` = `=TEXT('By Commission Week'!A155,"mm/dd")` … `AN20` → `A174`
  (the commission-by-group block the chart actually plots).
- `AO1` = `=TEXT('By Commission Week'!A106,"mm/dd")` … `AO20` → `A125`.
Verify each cell displays the same label it showed before (the current data hasn't
shifted yet, so text must be unchanged). Then the regression test: add a temporary
Summary row dated in a brand-new week → charts 2/4 must show the new week AND its
axis label must match; delete the temp row.

### B2. Chart 5 title garbage
`Graphs!H6` (chart 5's linked title) contains a stray pasted "RLN3MA" mid-sentence:
`="Issues by root cause — after vs before clean point RLN3MA("&$B$5&")"`.
Fix: set `H6` to the clean version (H5 already holds it):
`="Issues by root cause — after vs before clean point  ("&$B$5&")"`.
Leave `H5` alone. Export chart 5, confirm the title reads correctly.

### B3. Root-cause dropdown on Summary
Verify whether `Summary!E4:E240` carries list data-validation. If missing or partial,
add it: list source = the cause list (`='Root Cause Settings'!$A$4:$A$80` or an
equivalent named range), blanks allowed, error style Stop. Existing values must all
still pass Data Check "Root causes not in Settings" (= 0). This prevents the
typo-creates-invisible-data failure mode.

### B4. Data Checks frozen at migration-day values
`Data Checks!C4` compares to a hardcoded `138` and `C5` to `25` — every legitimate
data addition flips them to "Review", training users to ignore the sheet.
Fix: make them informational or self-consistent:
- Row 4 (migrated list lines): change check to `=IF(B4>=138,"OK","Review")` and the
  note to "count can only grow from the migrated 138".
- Row 5 (root causes): compare against the live count of non-blank
  `'Root Cause Settings'!A4:A80` instead of literal 25.
All other checks stay as-is.

### B5. Label inconsistencies between charts (display-text only — tread carefully)
The same concept appears under two names in different charts' series-name cells:
"Installation issue" (chart 4) vs "DG Installation issue" (chart 2), and
"Incorrect Program" (chart 2) vs "Wrong program" (chart 5 block). These are
display/series-name cells, some typed, some derived.
Fix ONLY what is safe: where the inconsistent name is **typed text in a series-name
cell** (e.g. `'By Commission Week'!` row 154 headers), align it to the name used in
`Root Cause Settings`. Do NOT rename anything in `Root Cause Settings` itself or in
Summary data (that would orphan rows — out of scope here). If a label turns out to
be the SUMIF criterion as well as the display name (row 154 cells are both), the
Settings spelling wins and you must verify the chart's totals are unchanged after the
edit; if totals change, revert that cell and log it as skipped.

## Part C — Acceptance (on the reopened `DG-template-fix1.xlsx` from the real path)

- [ ] All six charts render pixel-identical to the original file at Jun 1–Jul 15
      (export before/after and compare) — except chart 5's title, which is now clean.
- [ ] New-week regression test (B1) passes: fresh week appears with a correct axis
      label on charts 2 and 4.
- [ ] Summary E has the dropdown; an invalid cause is rejected; Data Check row 6 = 0.
- [ ] Data Checks shows no "Review" on rows 4/5 with the current, untouched data.
- [ ] Timeline still works: blank both dates → titles say "All data"; set Jun 1–Jun 30
      → charts 1/3/5/6 re-rank and re-title. Restore Jun 1 / Jul 15 before final save.
- [ ] Original `DG-template.xlsx` is byte-untouched; the fixed copy lives at
      `DG-template-fix1.xlsx` and survives close/reopen from the OneDrive path.
