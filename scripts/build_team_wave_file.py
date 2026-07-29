# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_team_wave_file.py — publish the shareable TEAM copy of the Master.

The Master Wave File is the internal working file (change logs, validation
columns, staging tabs, run logs). Team members just need to OPEN AN EXCEL FILE —
no SQL — so this script publishes a slimmed copy containing only:

    DASHBOARD          (with its hidden 'dashboard data' source, so it keeps working)
    Member Lookup      (team member lookup)
    Template Lookup    (job template lookup, with its hidden helper sheets)
    DATA               (the wave roster, with internal columns hidden)

Everything else (wave_distribution, JobCategoriesData, No Longer Rev Cycle,
Run Log, staging sheets) is deleted from the copy. The live Master is NEVER
touched — the script works on a temp copy and saves to a separate file:

    output:      data\\reports\\Main Reports\\RCM Wave Team File.xlsx        (overwritten each publish)
    dated copy:  data\\processed\\wave\\wave_data_copies\\wave_distribution_copies\\<month>\\

Uses Excel itself (COM) so the dashboard's pivots/slicers/charts survive intact
(openpyxl would strip them). Run AFTER the morning Master update:

    python scripts\\build_team_wave_file.py
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parent))
import onedrive_paths as op  # noqa: E402

# Sheets the team file keeps. Hidden helpers stay hidden but must survive so
# the visible sheets' formulas/pivots keep working.
KEEP_SHEETS = {
    "DASHBOARD",
    "dashboard data",       # pivot/formula source for DASHBOARD
    "Member Lookup",
    "Template Lookup",
    "template names",       # helper for Template Lookup (if present)
    "Wave_JobTemplate",     # data behind Template Lookup
    "DATA",
    "REF",                  # lists some formulas point at
}

# Kept for formulas but hidden from the team regardless of how the Master
# shows them.
HELPER_SHEETS = ("dashboard data", "template names", "Wave_JobTemplate", "REF")

# DATA columns hidden in the team copy (matched by header prefix, case-insensitive):
# internal change logs and bookkeeping fields. (The validation / recommendation /
# Epic-mirror columns that used to be listed here were deleted from the Master
# on 2026-07-06 — Epic data lives in SQL now.)
HIDE_COLUMN_PREFIXES = [
    "hr info",
    "Users Wave Change Log",
    "Users HR Change Log",
    "Recommended User Type",
]

XL_TO_LEFT = -4159
XL_XLSX = 51


def main() -> None:
    if not op.MASTER_WAVE_PATH.exists():
        sys.exit(f"ERROR: Master not found: {op.MASTER_WAVE_PATH}")

    # Work on a temp copy so the live Master (possibly open in Excel) is untouched.
    tmp_dir = Path(tempfile.mkdtemp(prefix="team_wave_"))
    work = tmp_dir / "work.xlsx"
    shutil.copy(op.MASTER_WAVE_PATH, work)

    # DispatchEx = a NEW hidden Excel instance (never the user's open session).
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(work))

        deleted = []
        for name in [ws.Name for ws in wb.Worksheets]:
            if name not in KEEP_SHEETS:
                wb.Worksheets(name).Delete()
                deleted.append(name)

        ws = wb.Worksheets("DATA")
        last_col = ws.Cells(1, ws.Columns.Count).End(XL_TO_LEFT).Column
        hidden = []
        for c in range(1, last_col + 1):
            header = str(ws.Cells(1, c).Value or "").strip().lower()
            if any(header.startswith(p.lower()) for p in HIDE_COLUMN_PREFIXES):
                ws.Columns(c).Hidden = True
                hidden.append(str(ws.Cells(1, c).Value))

        # Helper sheets stay hidden in the team copy no matter how they're
        # shown in the Master (xlSheetHidden = 0).
        for name in HELPER_SHEETS:
            try:
                wb.Worksheets(name).Visible = 0
            except Exception:
                pass  # sheet may not exist in this Master version

        wb.SaveAs(str(op.TEAM_WAVE_FILE_PATH), FileFormat=XL_XLSX)
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
        shutil.rmtree(tmp_dir, ignore_errors=True)

    dated = op.write_dated_copy(
        op.TEAM_WAVE_FILE_PATH,
        op.month_subdir(op.WAVE_DISTRIBUTION_DIR),
        "RCM_Wave_Team_File",
    )

    print(f"Published: {op.TEAM_WAVE_FILE_PATH}")
    print(f"Dated copy: {dated}")
    print(f"Sheets removed ({len(deleted)}): {', '.join(deleted)}")
    print(f"DATA columns hidden ({len(hidden)}): {', '.join(hidden)}")


if __name__ == "__main__":
    main()
