# YourOrg Analytics — Engineering Guide

## Project Context

RCM (Revenue Cycle Management) Wave training analytics for YourOrg Health System. The
source of truth is the **RCM Wave Data Master File** (DATA sheet, one row per
person), which the daily pipeline updates from the HR, MVP, Epic and Cornerstone
exports and loads into SQL Server (`.\SQLEXPRESS`, `AnalyticsDB`).
Reports are generated from SQL, not maintained by hand.

**Data lives under two roots — never hand-build a path, resolve everything
through `scripts/onedrive_paths.py`:**

| Root | Holds |
|---|---|
| **Local** `yourorg_analytics\data\` (since 2026-08-11) | Master Wave File + its daily backups, DATA snapshots and review CSVs. Not synced, not shared. |
| **OneDrive** `shared_workspace\data\` | All raw exports, and every shared deliverable — RCM Wave Team File (the only roster that gets shared), Training Tracker + Daily Log, Morning Review, dashboards, boss gap reports. |

The daily refresh (`scripts/daily_refresh.py`, runbook in
`scripts/DAILY_REFRESH.md`) runs prep → apply on a schedule; the Master is
written ONLY by `update_master_from_mvp.py` (surgical XML),
`update_master_from_hr_safe.py` and `update_master_training_status.py`
(both xlwings/COM). Never write the Master with openpyxl.

Older local folders (`EPIC EXPORTS/`, `HR EXPORTS/`, `MASTER WAVE FILE/`,
`REPORTS/`) are pre-migration leftovers — `REPORTS/` keeps only Job Role Change
Report inputs and templates; the rest is archived.

---

## Python / Library Conventions

### Library roles — use the right tool

| Need | Library |
|---|---|
| Read values from xlsx | `openpyxl` with `data_only=True` |
| Read formulas from xlsx | `openpyxl` with `data_only=False` |
| Write new xlsx with charts, styles, formatting | `xlsxwriter` |
| Modify existing xlsx formulas/structure | `openpyxl` OR zipfile+XML (see below) |
| Data manipulation | `pandas` |
| Read old `.xls` files | `xlrd` |

### Always back up before modifying

```python
import shutil
shutil.copy(str(SRC), str(SRC.with_suffix('.bak.xlsx')))
```

### Loading pattern

```python
wb_vals = openpyxl.load_workbook(path, data_only=True)   # cached values only
wb_form = openpyxl.load_workbook(path, data_only=False)  # formula strings
```

Never load a large workbook both ways simultaneously unless required — double memory hit.

### XML patching (when openpyxl can't round-trip)

openpyxl silently drops: dynamic/spill arrays, x14 data-validation extensions, some chart XML.
For those, patch the xlsx XML directly via `zipfile`:

```python
import zipfile, shutil, re
SRC = Path('REPORTS/file.xlsx')
TMP = SRC.with_suffix('.patched.xlsx')
with zipfile.ZipFile(SRC, 'r') as zin, zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'xl/worksheets/sheet2.xml':
            text = data.decode('utf-8')
            text = text.replace(OLD, NEW)
            data = text.encode('utf-8')
        zout.writestr(item, data)
shutil.move(str(TMP), str(SRC))
```

Order of edits inside the XML loop matters — apply broadest replacements last.

---

## Formula Best Practices

### Preferred functions (fast, non-volatile)

| Use | Instead of | Why |
|---|---|---|
| `XLOOKUP` | `VLOOKUP` / `INDEX+MATCH` | Exact match both ways, handles missing, no column index |
| `COUNTIFS` / `SUMIFS` | `SUMPRODUCT` on large ranges | Native C loop, much faster |
| `IFS` | nested `IF` | Readable, same speed |
| `UNIQUE` / `SORT` / `FILTER` | manual list maintenance | Dynamic spill — updates automatically |
| Structured table refs `Table1[Column]` | `$A$2:$A$9999` | Auto-expands, self-documenting |

### Never use on large ranges (volatile / slow)

- `INDIRECT()` — recalculates on every sheet change, blocks caching
- `OFFSET()` — volatile, use `INDEX` instead
- `NOW()` / `TODAY()` in cells that feed calculations — triggers full recalc constantly
- Entire-column references (`A:A`) in SUMIFS/COUNTIFS on sheets with >50k rows
- Nested `VLOOKUP` inside `SUMPRODUCT` — O(n²) effective cost

### Avoid array formula sprawl

Traditional `{Ctrl+Shift+Enter}` array formulas entered over hundreds of cells are a
common crash cause. Replace with:
- A single spill formula in one cell (`=FILTER(...)`, `=UNIQUE(...)`)
- `SUMIFS` / `COUNTIFS` (they are inherently array-aware)
- A helper column on the Data sheet computed once

### Range sizing discipline

Tight ranges always outperform open-ended ones:
- Audit with `ws.max_row` / `ws.max_column` before writing formulas
- When writing XLOOKUP lookup arrays, size to actual data + 20% buffer, not entire columns
- Document the intended range in a comment cell or named range

---

## Workbook Performance — Preventing Crashes

### Size thresholds (rough guidelines)

| Metric | Safe | Caution | Danger |
|---|---|---|---|
| Total cells with formulas | < 50k | 50k–200k | > 200k |
| Unique volatile calls | 0 | < 10 | Any INDIRECT/OFFSET in loops |
| External links | 0 | 1–2 | > 2 |
| Charts on one sheet | < 8 | 8–15 | > 15 |
| Conditional formatting rules | < 20 | 20–50 | > 50 |

### Audit before delivering a workbook

Always run `scripts/validate_tracker.py` pattern before handing off any workbook:

```python
error_codes = {'#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#N/A', '#NULL!', '#NUM!', '#SPILL!', '#CALC!'}
for sname in wb_v.sheetnames:
    for row in wb_v[sname].iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value in error_codes:
                # log sname, cell.coordinate, cell.value
```

Check for:
1. Cached error values in every sheet
2. `#REF!` embedded in formula strings (formula is broken, not just the value)
3. External link references `[filename.xlsx]` — break when file moves
4. Formulas that are just `=` (empty formula — openpyxl write artifact)

### Calculation mode

When building a workbook programmatically, set manual calculation to prevent
Excel from recalculating on every cell write:

```python
wb.calculation.calcMode = 'manual'  # user hits F9 to refresh
# — or —
wb.calculation.calcMode = 'auto'    # default; fine for small workbooks
```

For xlsxwriter:
```python
workbook = xlsxwriter.Workbook('output.xlsx')
workbook.set_calc_mode('manual')
```

---

## Dashboards

### Sheet structure pattern

```
_Data (hidden)     ← lookup arrays, named ranges, raw helper cols
Data               ← source rows, one per record
Summary            ← KPI cards, aggregate formulas referencing Data
Exec Dashboard     ← charts, percent bars, leadership view
[Dimension]Tracking ← breakdowns (Manager, Class, Leader, etc.)
```

### Named ranges — always use them for dashboard formulas

```python
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName
wb = load_workbook('file.xlsx')
# Define a named range pointing to a sheet range
ref = "'Summary'!$B$4:$N$4"
wb.defined_names['KPI_Row'] = DefinedName('KPI_Row', attr_text=ref)
```

Then formulas reference `KPI_Row` instead of hardcoded coordinates — survives row inserts.

### Freeze panes and print setup

```python
ws.freeze_panes = 'B5'          # freeze rows 1-4 and column A
ws.sheet_view.showGridLines = False
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
```

### Conditional formatting (performance-safe)

Apply rules to exact data ranges, not whole columns:

```python
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

green = PatternFill(bgColor='C6EFCE')
red   = PatternFill(bgColor='FFC7CE')
# Apply only to the actual data range
ws.conditional_formatting.add(
    f'D2:D{last_row}',
    CellIsRule(operator='greaterThanOrEqual', formula=['0.8'], fill=green)
)
```

---

## Charts

### xlsxwriter (preferred for new workbooks with charts)

```python
import xlsxwriter
wb = xlsxwriter.Workbook('dashboard.xlsx')
ws = wb.add_worksheet('Exec Dashboard')

chart = wb.add_chart({'type': 'bar', 'subtype': 'stacked'})
chart.add_series({
    'name':       ['Summary', 0, 1],       # row 0, col 1 = header cell
    'categories': ['Summary', 1, 0, 10, 0],
    'values':     ['Summary', 1, 1, 10, 1],
    'fill':       {'color': '#4472C4'},
})
chart.set_title({'name': 'Training Progress by Manager'})
chart.set_size({'width': 480, 'height': 288})
ws.insert_chart('B2', chart)
wb.close()
```

### openpyxl charts (modifying existing workbooks)

openpyxl can add charts but cannot reliably preserve existing ones on round-trip.
Use XML patching to preserve existing charts when only changing data ranges.

### Chart data — always from a named table or explicit range

Never hardcode row counts in chart series references. Use:
```python
last_row = ws.max_row  # compute at write time
chart.add_series({'values': f"='Data'!$C$2:$C${last_row}"})
```

### Pivot charts

openpyxl/xlsxwriter cannot create native Excel pivot tables.
Pattern: compute aggregation in pandas → write results to a summary sheet → add a regular chart over that data.

```python
pivot = df.pivot_table(index='Epic Job Category', values='Fully Trained', aggfunc='sum')
pivot.to_excel(writer, sheet_name='PT_JobRole', startrow=1)
# Then add xlsxwriter chart pointing at PT_JobRole sheet
```

---

## Large Dataset Patterns

### Read performance

```python
# Stream rows instead of loading entire sheet into memory
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
    if all(v is None for v in row):
        continue
    # process row
```

### Write performance

```python
# Use write_only mode for large output — 10x faster, lower memory
wb = openpyxl.Workbook(write_only=True)
ws = wb.create_sheet('Data')
ws.append(headers)
for record in records:
    ws.append(record)
wb.save('output.xlsx')
```

### pandas ↔ Excel

```python
# Reading
df = pd.read_excel('file.xlsx', sheet_name='Data', engine='openpyxl')

# Writing multiple sheets
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    df_summary.to_excel(writer, sheet_name='Summary', index=False)
    df_detail.to_excel(writer, sheet_name='Data',    index=False)
```

Use `engine='openpyxl'` when you need to preserve existing sheets.
Use `engine='xlsxwriter'` when writing new files with charts/formatting.

---

## Validation Checklist (run before delivering any workbook)

- [ ] No error values (`#REF!`, `#VALUE!`, `#NAME?`, etc.) in any sheet
- [ ] No `#REF!` embedded in formula strings
- [ ] No external links pointing to other xlsx files
- [ ] KPI cells cross-validated against re-computation from raw Data sheet
- [ ] Percentage formulas use consistent denominator (check `validate_tracker.py` pattern)
- [ ] Named ranges exist for all lookup arrays used by formulas
- [ ] Conditional formatting applied to bounded ranges, not whole columns
- [ ] File size < 10 MB before distributing (images and charts bloat fast)
- [ ] Calculation mode set intentionally (auto vs manual)
- [ ] Backup copy made before any destructive script run

---

## Agent team & routing (local only, since 2026-09-14)

The main session is the **Chief of Staff**: it routes, holds every approval,
and is the only place that asks the analyst questions. Six specialists live in
`agents\roles\` and preload the operations handbook. Playbooks
`/start-day` and `/end-day` build the morning brief and the wrap-up.

| She says | Delegate to | Then the main session |
|---|---|---|
| update my Master / run the refresh / what failed / did it run | `morning-review` (prep + diagnosis) | runs apply after her OK |
| change X on the Master / who changed / add these people / Team File | `wave-file-steward` (preview + change list) | runs the apply command after her OK |
| why is this person…, what does Fully Trained mean, Epic vs Cornerstone | `training-registration-expert` | relays the finding |
| today's numbers / how many / by leader / pull me a list | `data-analyst` | relays the tables |
| build / rebuild / fix the tracker, LAVA, dashboard, a workbook for the boss | `report-builder` | confirms before any NEW file lands on OneDrive |
| add a column / new view / script errors / change the pipeline | `pipeline-engineer` | reviews the diff, runs nothing against the live Master |
| anything from an email | main session builds the date timeline first (EMAIL_LOG rule), then routes | |

**Stop-and-return contract.** Specialists cannot ask the human anything. When an agent
reaches an action on its "Never" list, or would overwrite a hand-entered cell,
it stops and returns a `DECISION NEEDED` block (what / evidence /
recommended). The orchestrator asks the human and acts. Apply steps, Master
writes, overwrites of manual edits and first-time OneDrive publishes are
always main-session actions taken after her explicit OK.

The pre-write guard applies to agent tool calls too.
Agent and skill files are local; never copy them to any `shared_workspace` path.
