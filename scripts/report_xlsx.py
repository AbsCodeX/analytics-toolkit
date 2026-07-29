# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
report_xlsx.py

Shared helpers for writing formatted xlsx reports (bold header, autofilter,
frozen header row, sane column widths) — used by the small single-purpose
report scripts (missing_from_wave_hr.py, missing_from_wave_epic.py,
wave_missing_from_hr_epic_mvp.py) so the formatting isn't duplicated in each.
"""

from pathlib import Path

import pandas as pd
import xlsxwriter


def _cell_str(v) -> str:
    """Stringify a cell value, treating None/NaN (incl. pandas.NA) as blank.
    (.astype(str) alone doesn't reliably turn float NaN into '' across dtypes,
    e.g. rows with unmatched leader-lookup columns left as real NaN floats.)"""
    if v is None:
        return ""
    if isinstance(v, float) and v != v:  # NaN
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def _write_sheet(workbook, worksheet, df: pd.DataFrame, link_col: str | None = None,
                 stamp: str | None = None) -> None:
    """Write df into an already-created worksheet: bold header, autofilter,
    frozen header row, and column widths sized to content (capped at 50).

    If link_col names a column, its non-blank values are treated as local
    file/folder paths and written as clickable external hyperlinks (the path
    string stays visible as the cell text).

    If stamp is given, it is written as a muted one-line banner in row 1
    (e.g. "Last updated: 2026-07-28 09:48") and the header/data shift down
    one row — header row stays frozen and filtered either way."""
    columns = list(df.columns)
    bold = workbook.add_format({"bold": True})
    link_fmt = workbook.add_format({"font_color": "#4472C4", "underline": 1})
    stamp_fmt = workbook.add_format({"italic": True, "font_color": "#808080"})
    link_c = columns.index(link_col) if link_col in columns else None

    head_r = 0
    if stamp:
        worksheet.write_string(0, 0, stamp, stamp_fmt)
        head_r = 1

    for c, name in enumerate(columns):
        worksheet.write(head_r, c, name, bold)
    for r, row in enumerate(df.itertuples(index=False), start=head_r + 1):
        for c, val in enumerate(row):
            s = _cell_str(val)
            if c == link_c and s:
                worksheet.write_url(r, c, f"external:{s}", link_fmt, string=s)
            else:
                worksheet.write_string(r, c, s)

    last_row = head_r + max(len(df), 1)
    if columns:
        worksheet.autofilter(head_r, 0, last_row, len(columns) - 1)
    worksheet.freeze_panes(head_r + 1, 0)

    sample = df.head(500)
    for c, name in enumerate(columns):
        widths = [len(_cell_str(v)) for v in sample[name]] if len(sample) else []
        sample_width = max(widths, default=0)
        width = max(len(name), sample_width) + 2
        worksheet.set_column(c, c, min(width, 50))


def write_formatted_xlsx(df: pd.DataFrame, path: Path, sheet_name: str) -> None:
    """Write df as a single-sheet formatted xlsx."""
    write_formatted_workbook({sheet_name: df}, path)


def write_formatted_workbook(sheets: dict[str, pd.DataFrame], path: Path,
                             link_cols: dict[str, str] | None = None,
                             stamps: dict[str, str] | None = None) -> None:
    """Write one formatted xlsx with one worksheet per (sheet_name, df) pair,
    in the given dict order. link_cols optionally maps sheet name -> column
    name whose path values become clickable hyperlinks; stamps optionally maps
    sheet name -> banner text written above the header (see _write_sheet)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(path))
    for sheet_name, df in sheets.items():
        worksheet = workbook.add_worksheet(sheet_name)
        _write_sheet(workbook, worksheet, df, (link_cols or {}).get(sheet_name),
                     (stamps or {}).get(sheet_name))
    workbook.close()
