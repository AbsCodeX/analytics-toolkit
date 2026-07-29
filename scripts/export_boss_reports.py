# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
export_boss_reports.py — publish the recurring manager gap reports.

Read-only (never touches the Master). Two reports land in the shared OneDrive
folder data\\reports\\gap_reports\\ — a stable "(Latest)" copy at the top level
plus a dated copy in a monthly subfolder:

  DAILY   "Epic Not In HR (Latest).xlsx"
          Epic Team Member Lookup people with no row in the HR file, from SQL
          view report.epic_not_in_hr. FirstSeenDate/NewToday come from
          history.epic_not_in_hr_daily (stamped by sql\\refresh.py's snapshot
          step, which must run first — daily_refresh.py enforces the order).
          NewToday = first appearance ever, flagged only once at least two
          snapshot days exist (the baseline day flags nobody).
          Default scope: on the Master OR RCM-Centralized curriculum;
          --all exports every Epic row missing from HR.

  WEEKLY  "HR Missing From Wave (Latest).xlsx"
          A copy of the newest hr_missing_from_wave_*.xlsx produced by
          missing_from_wave_hr.py. Gate: today is Monday OR a fresh HR.xlsx
          landed today (--weekly forces it).

Usage:
    python scripts\\export_boss_reports.py            # daily report (+ weekly if due)
    python scripts\\export_boss_reports.py --all      # daily report without scope filter
    python scripts\\export_boss_reports.py --weekly   # force the weekly HR copy too
"""

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

import activity_log                      # noqa: E402
import onedrive_paths as op              # noqa: E402
from refresh import CONN                 # noqa: E402  (shared SQL connection string)
from report_xlsx import write_formatted_xlsx  # noqa: E402
from sqlalchemy import create_engine     # noqa: E402

EPIC_LATEST = op.GAP_REPORTS_DIR / "Epic Not In HR (Latest).xlsx"
HR_LATEST = op.GAP_REPORTS_DIR / "HR Missing From Wave (Latest).xlsx"

EPIC_NOT_IN_HR_SQL = """
WITH first_seen AS (
    SELECT UniversalID, MIN(SnapshotDate) AS FirstSeenDate
    FROM history.epic_not_in_hr_daily
    GROUP BY UniversalID
),
days AS (
    SELECT MAX(SnapshotDate) AS LatestDay,
           COUNT(DISTINCT SnapshotDate) AS DayCount
    FROM history.epic_not_in_hr_daily
)
SELECT e.UniversalID, e.FullName, e.JobTitle, e.CurriculumType, e.Wave,
       e.WorkerType, e.BusinessUnit, e.DirectManager, e.HireDate,
       e.EpicTrainingNeeded, e.EpicEligible, e.HrStatusPDM, e.OffshoreVendorPDM,
       e.OnMaster, e.IsRCMCurriculum,
       fs.FirstSeenDate,
       CASE WHEN d.DayCount > 1 AND fs.FirstSeenDate = d.LatestDay
            THEN 'Yes' ELSE '' END AS NewToday
FROM report.epic_not_in_hr e
LEFT JOIN first_seen fs ON fs.UniversalID = e.UniversalID
CROSS JOIN days d
{where}
ORDER BY CASE WHEN d.DayCount > 1 AND fs.FirstSeenDate = d.LatestDay THEN 0 ELSE 1 END,
         e.FullName
"""


def epic_not_in_hr_frame(engine, scope_all: bool = False) -> pd.DataFrame:
    """The daily Epic-not-in-HR data, NewToday rows first. Shared with
    build_morning_review.py so the boss report and the review sheet match."""
    where = "" if scope_all else "WHERE e.OnMaster = 1 OR e.IsRCMCurriculum = 1"
    return pd.read_sql(EPIC_NOT_IN_HR_SQL.format(where=where), engine)


def export_epic_not_in_hr(engine, scope_all: bool) -> None:
    df = epic_not_in_hr_frame(engine, scope_all)
    write_formatted_xlsx(df, EPIC_LATEST, "Epic Not In HR")
    dated = op.write_dated_copy(EPIC_LATEST, op.month_subdir(op.GAP_REPORTS_DIR),
                                "epic_not_in_hr")
    new_today = int((df["NewToday"] == "Yes").sum()) if len(df) else 0
    print(f"Epic Not In HR: {len(df)} people ({new_today} new today) -> {EPIC_LATEST.name}"
          f"  (+ {dated.name})")


def weekly_hr_copy_due() -> tuple[bool, str]:
    if date.today().weekday() == 0:
        return True, "Monday"
    raw_mtime = datetime.fromtimestamp(op.RAW_HR_PATH.stat().st_mtime).date() \
        if op.RAW_HR_PATH.exists() else None
    if raw_mtime == date.today():
        return True, "fresh HR file today"
    return False, "not Monday and no fresh HR file"


def export_hr_missing_weekly(force: bool) -> None:
    due, reason = weekly_hr_copy_due()
    if not (due or force):
        print(f"HR Missing From Wave: skipped ({reason})")
        return
    src = op.latest_dated_file(op.MISSING_FROM_WAVE_DIR, "hr_missing_from_wave", ".xlsx")
    if src is None:
        print("HR Missing From Wave: no hr_missing_from_wave_*.xlsx found "
              "(run missing_from_wave_hr.py first) — skipped")
        return
    op.GAP_REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy(src, HR_LATEST)
    dated_dir = op.month_subdir(op.GAP_REPORTS_DIR)
    shutil.copy(src, dated_dir / src.name)
    print(f"HR Missing From Wave: copied {src.name} -> {HR_LATEST.name} ({reason})")


def main() -> None:
    ap = argparse.ArgumentParser(description="Publish the manager gap reports.")
    ap.add_argument("--all", action="store_true",
                    help="Epic report without the OnMaster/RCM-curriculum scope filter")
    ap.add_argument("--weekly", action="store_true",
                    help="force the weekly HR-missing-from-wave copy")
    args = ap.parse_args()

    engine = create_engine(CONN)
    try:
        export_epic_not_in_hr(engine, args.all)
    finally:
        engine.dispose()
    export_hr_missing_weekly(args.weekly)
    print(f"Boss reports published to {op.GAP_REPORTS_DIR}")


if __name__ == "__main__":
    with activity_log.track_run("export_boss_reports.py"):
        main()
