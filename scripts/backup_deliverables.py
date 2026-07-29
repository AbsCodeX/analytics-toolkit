# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
backup_deliverables.py — once-a-day pre-rebuild safety copies of the
deliverables that are OVERWRITTEN in place by the apply stage and had no
dated archive of their own (the Master, Team File, boss reports, and Morning
Review already keep dated copies via their own scripts).

Runs as the first apply step after the guard, BEFORE anything is rebuilt, so
each .bak.<date> file holds the previous day's closing version — same
semantics as the Master's daily .bak. At most one backup per calendar day
per file (re-runs are no-ops). Also calls
master_backup.ensure_daily_full_backup() so the Master gets its daily .bak
even on days when no updater happens to write it first.

Restore = copy the .bak file back over the live file (Excel closed).

Usage:
    python scripts\\backup_deliverables.py
"""

from __future__ import annotations

import shutil
import sys
from datetime import datetime
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPTS))

import activity_log                  # noqa: E402
import master_backup                 # noqa: E402
import onedrive_paths as op          # noqa: E402

MASTER_REPORTS_DIR = op.MAIN_REPORTS_DIR
ARCHIVE_DIR = op.MAIN_REPORTS_BACKUPS_DIR

# Files rebuilt in place by apply that need a pre-rebuild daily copy.
TARGETS = [
    MASTER_REPORTS_DIR / "RCM Training Tracker.xlsx",
    MASTER_REPORTS_DIR / "RCM Training Daily Log.xlsx",
]


def backup_once_daily(src: Path, archive_base: Path) -> str:
    if not src.exists():
        return f"missing (skipped): {src.name}"
    month_dir = op.month_subdir(archive_base)
    stamp = datetime.now().strftime("%Y-%m-%d")
    dest = month_dir / f"{src.stem}.bak.{stamp}{src.suffix}"
    if dest.exists():
        return f"already backed up today: {dest.name}"
    shutil.copy(src, dest)
    return f"backed up: {dest.name}"


def main() -> None:
    results = []
    for target in TARGETS:
        msg = backup_once_daily(target, ARCHIVE_DIR)
        print(f"  {msg}")
        results.append(msg)
    master_bak = master_backup.ensure_daily_full_backup(op.MASTER_WAVE_PATH)
    done = sum(1 for m in results if m.startswith("backed up"))
    print(f"Daily safety copies done: {done} new report backup(s); "
          f"Master backup: {master_bak.name}")


if __name__ == "__main__":
    with activity_log.track_run("backup_deliverables.py"):
        main()
