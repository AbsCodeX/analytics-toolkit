# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
set_master_auto_calc.py — switch the Master Wave File to AUTOMATIC calculation.

One-time fix requested by the analyst 2026-07-28: the workbook was left on manual
calculation by the 2026-05-13 performance fix, which made post-refresh values
look stale/"broken" until F9. The underlying freeze cause (re-expanded shared
formula ranges + unanchored CF) was fixed separately, so auto calc is safe now.

Goes through desktop Excel (xlwings/COM) like every Master writer — never
openpyxl. Performs a full rebuild recalc before saving so all cached values
are fresh. Refuses if the Master is open in Excel.

Usage:
    python scripts\\set_master_auto_calc.py
"""

from __future__ import annotations

import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

import activity_log                  # noqa: E402
import master_backup                 # noqa: E402
import onedrive_paths as op          # noqa: E402

XL_CALCULATION_AUTOMATIC = -4105  # Excel enum xlCalculationAutomatic


def master_locked() -> bool:
    lock = op.MASTER_WAVE_PATH.parent / ("~$" + op.MASTER_WAVE_PATH.name)
    return lock.exists()


def main() -> None:
    if master_locked():
        sys.exit("ERROR: the Master is open in Excel. Close it and rerun.")

    try:
        backup = master_backup.ensure_daily_full_backup(op.MASTER_WAVE_PATH)
    except PermissionError:
        sys.exit(f"ERROR: Could not create backup (permission denied): {op.MASTER_WAVE_PATH}")
    print(f"Backup: {backup.name}")

    import xlwings as xw

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(op.MASTER_WAVE_PATH))
        before = app.api.Calculation
        app.api.Calculation = XL_CALCULATION_AUTOMATIC
        app.api.CalculateFullRebuild()
        wb.save()
        wb.close()
        print(f"Calculation mode: {before} -> automatic; full recalc done; saved.")
    finally:
        if app is not None:
            app.quit()

    snap = master_backup.save_data_snapshot(op.MASTER_WAVE_PATH)
    print(f"DATA snapshot: {snap}")


if __name__ == "__main__":
    with activity_log.track_run("set_master_auto_calc.py"):
        main()
