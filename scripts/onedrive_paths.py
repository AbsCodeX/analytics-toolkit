# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
onedrive_paths.py

Single source of truth for ALL OneDrive raw/processed/report paths. Every
active script in scripts/ and the SQL loaders in sql/ resolve their data
paths through this module — no live code builds its own data path
(verified 2026-07-13). Anything still on the old local EXPORTS/REPORTS
layout lives in scripts/_archive/ and is retired.
"""

import shutil
from datetime import date
from pathlib import Path

ONEDRIVE_ROOT = Path(
    r"C:\path\to\shared\workspace"
)

# ---------------------------------------------------------------------------
# Raw inputs
# ---------------------------------------------------------------------------
RAW_HR_PATH = ONEDRIVE_ROOT / "data" / "raw" / "hr" / "HR.xlsx"

RAW_MVP_DIR = ONEDRIVE_ROOT / "data" / "raw" / "mvp"
RAW_MVP_USER_MAPPINGS = RAW_MVP_DIR / "User MappingsData.csv"
RAW_MVP_ROLE_MAPPINGS = RAW_MVP_DIR / "RoleMappingsData.csv"
RAW_MVP_JOB_CATEGORIES = RAW_MVP_DIR / "JobCategoriesData.csv"

RAW_EPIC_TM_DIR = ONEDRIVE_ROOT / "data" / "raw" / "epic" / "epic_tm_lookup"
RAW_EPIC_STATUS_DIR = ONEDRIVE_ROOT / "data" / "raw" / "epic" / "epic_status_details"
RAW_EPIC_CLASS_SCHEDULES_DIR = ONEDRIVE_ROOT / "data" / "raw" / "epic" / "epic_class_schedules"

RAW_CORNERSTONE_DIR = ONEDRIVE_ROOT / "data" / "raw" / "cornerstone" / "enterprise_training_reports"

RAW_WAVE_CHANGE_REQUESTS_DIR = ONEDRIVE_ROOT / "data" / "raw" / "wave" / "wave_change_requests"

# ---------------------------------------------------------------------------
# Main Reports — the ONE folder for all primary deliverables (2026-07-29).
# Master, Team File, Training Tracker + Daily Log, Morning Review workbook,
# Wave/Job Role Change Report, CSI Tracker live here; their backups go to
# Backups\; HTML dashboards to Dashboards\.
# ---------------------------------------------------------------------------
MAIN_REPORTS_DIR = ONEDRIVE_ROOT / "data" / "reports" / "Main Reports"
MAIN_REPORTS_BACKUPS_DIR = MAIN_REPORTS_DIR / "Backups"
DASHBOARDS_DIR = MAIN_REPORTS_DIR / "Dashboards"

# Master Wave File — sole source of truth going forward
MASTER_WAVE_PATH = MAIN_REPORTS_DIR / "RCM Wave Data Master File.xlsx"

# Shareable team copy of the Master (dashboard + lookups + clean roster only),
# published by scripts/build_team_wave_file.py. Dated copies of each publish
# land in WAVE_DISTRIBUTION_DIR/<month>/.
TEAM_WAVE_FILE_PATH = MAIN_REPORTS_DIR / "RCM Wave Team File.xlsx"
WAVE_DISTRIBUTION_DIR = (
    ONEDRIVE_ROOT / "data" / "processed" / "wave" / "wave_data_copies" / "wave_distribution_copies"
)

# Ad-hoc SQL pulls (query.py exports) — shared with the team via OneDrive.
SQL_PULLS_DIR = ONEDRIVE_ROOT / "data" / "reports" / "sql_pulls"

# Boss-facing gap reports (weekly HR-not-on-wave, daily Epic-not-in-HR):
# stable "(Latest)" copies at the top level + dated copies in monthly folders.
GAP_REPORTS_DIR = ONEDRIVE_ROOT / "data" / "reports" / "gap_reports"

# The consolidated morning review workbook (built by build_morning_review.py):
# the ONE file the analyst opens before running the apply stage.
# Workbook lives in Main Reports; the DIR keeps the supporting preview CSVs
# and the morning-review archive.
MORNING_REVIEW_DIR = ONEDRIVE_ROOT / "data" / "reports" / "morning_review"
MORNING_REVIEW_PATH = MAIN_REPORTS_DIR / "Morning Review.xlsx"

# Standard SQL-built reports (export_sql_reports.py): no-shows + unregistered.
NO_SHOW_REPORTS_DIR = ONEDRIVE_ROOT / "data" / "reports" / "no_show"
UNREGISTERED_REPORTS_DIR = ONEDRIVE_ROOT / "data" / "reports" / "unregistered"

# Wave 3 daily registration package (export_w3_registration_reports.py):
# one multi-tab workbook (unregistered / no-shows / no-show resolution /
# daily leader summary / leader metrics) + the rolling look-back CSV.
W3_REGISTRATION_REPORTS_DIR = ONEDRIVE_ROOT / "data" / "reports" / "registration_daily"

# CSI interactions tracker (scripts/csi_tracker.py): drop the daily
# "CSI interactions by day" PDFs into CSI_INBOX_DIR, run the script, and it
# merges them into the rolling tracker workbook (upsert by IMS number) and
# archives each ingested PDF into CSI_ARCHIVE_DIR/<YYYY-MM>/.
CSI_INBOX_DIR = ONEDRIVE_ROOT / "data" / "raw" / "csi_interactions"
CSI_ARCHIVE_DIR = CSI_INBOX_DIR / "archive"
CSI_TRACKER_PATH = MAIN_REPORTS_DIR / "CSI Interactions Tracker.xlsx"

# Reference lists (lookup logic kept out of the Master file): one workbook,
# one list per sheet. Loaded into SQL as raw.ref_<sheetname> every refresh.
REFERENCES_DIR = ONEDRIVE_ROOT / "data" / "references"
REF_LISTS_PATH = REFERENCES_DIR / "wave_reference_lists.xlsx"

# Source registry: one row per export type. The daily refresh loads every
# Enabled row into raw.<source>_<table> and creates the <source>.<table> view.
# Adding a new export = drop the file + add one row here. (sql\SOURCES.md)
SOURCE_REGISTRY_PATH = REFERENCES_DIR / "source_registry.xlsx"

# ---------------------------------------------------------------------------
# Processed outputs
# ---------------------------------------------------------------------------
# data/processed/ is grouped into per-source-area folders (hr/, epic/, mvp/,
# wave/, cornerstone/, ad_hoc/, power_bi_ready/) — updated here to match that
# grouping. Only the folders scripts actually write to are given constants
# below; the empty placeholder folders (clean_mvp_user_mappings,
# clean_epic_unregistered, power_bi_ready, etc.) belong to not-yet-migrated
# scripts and aren't referenced yet.
PROCESSED_DIR = ONEDRIVE_ROOT / "data" / "processed"
CLEAN_HR_DIR = PROCESSED_DIR / "hr" / "clean_hr_copies"
CLEAN_EPIC_TM_DIR = PROCESSED_DIR / "epic" / "clean_epic_team_member_copies"
MISSING_FROM_WAVE_DIR = PROCESSED_DIR / "wave" / "missing_from_wave"

# Raw (pre-filter) combined Epic Team Member Lookup, for manual review only
# (single pre-combined workbook today; per-wave W1-W4 files also supported).
# Single stable file, OVERWRITTEN every run — not dated, not read by any other
# script. The dated, filtered deliverable lives in CLEAN_EPIC_TM_DIR instead.
EPIC_TM_RAW_REVIEW_DIR = PROCESSED_DIR / "epic" / "epic_team_member_raw_review"
EPIC_TM_RAW_REVIEW_PATH = EPIC_TM_RAW_REVIEW_DIR / "combined.xlsx"

# Daily history snapshots (sql/refresh.py `snapshot` step): a dated CSV per day
# per history table, in monthly subfolders. The durable copy of the SQL
# history.* tables — `refresh.py backload_history` rebuilds them from here.
HISTORY_SNAPSHOTS_DIR = PROCESSED_DIR / "history_snapshots"

# ---------------------------------------------------------------------------
# Runlogs + Master backups/versioning
# ---------------------------------------------------------------------------
RUNLOGS_DIR = ONEDRIVE_ROOT / "data" / "runlogs"

# Full-workbook Master backups (at most one per calendar day) and DATA-sheet-only
# versioned snapshots (one per actual Master edit) — both organized into monthly
# subfolders via month_subdir() so the flat folders don't grow unbounded.
MASTER_ARCHIVE_DIR = ONEDRIVE_ROOT / "data" / "reports" / "archive"
WAVE_REPOSITORY_DIR = PROCESSED_DIR / "wave" / "wave_data_copies" / "wave_repository_copies"


def month_subdir(base_dir: Path, when=None) -> Path:
    """base_dir/<YYYY-MM>/, created if missing."""
    d = base_dir / f"{(when or date.today()):%Y-%m}"
    d.mkdir(parents=True, exist_ok=True)
    return d


def dated_stem(stem: str) -> str:
    """'{stem}_{today as YYYY.MM.DD}' — sortable, human-readable dated filename stem."""
    return f"{stem}_{date.today():%Y.%m.%d}"


def write_dated_copy(saved_path: Path, dest_dir: Path, stem: str) -> Path:
    """Copy an already-saved stable-name output to a dated companion filename
    in dest_dir, for lookback/archive. Returns the dated copy's path."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f"{dated_stem(stem)}{saved_path.suffix}"
    shutil.copy(saved_path, dest)
    return dest


def latest_dated_file(directory: Path, stem: str, ext: str) -> Path | None:
    """Return the most recently dated '{stem}_YYYY.MM.DD{ext}' file in directory,
    or None if none exist. Dated filenames sort chronologically as plain strings
    (fixed-width zero-padded YYYY.MM.DD), so a plain sort finds the latest."""
    if not directory.is_dir():
        return None
    matches = sorted(directory.glob(f"{stem}_*{ext}"))
    return matches[-1] if matches else None


def latest_hr_cleaned() -> Path | None:
    """Most recent dated 'HR_cleaned_YYYY.MM.DD.xlsx' in CLEAN_HR_DIR, or None."""
    return latest_dated_file(CLEAN_HR_DIR, "HR_cleaned", ".xlsx")
