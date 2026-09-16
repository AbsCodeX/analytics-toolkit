# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
export_audit.py — are the export files on disk fresh, loaded into SQL, and sane?

Since 2026-08-04 the daily exports OVERWRITE fixed filenames (no dated copies),
so filename comparison alone can't say whether a feed is current. This audit
checks every feed end-to-end and produces one row per feed with a flag:

  1. File exists on disk (FAIL if missing)
  2. File updated on schedule — daily feeds: modified today on weekdays;
     HR: within 9 days (WARN if stale)
  3. SQL holds the current file — raw table _loaded_at vs file modified time
     (WARN if a newer file is on disk but not loaded yet)
  4. Row count sanity — >20% drop vs the previous audit run (WARN)
  5. Enterprise cap — loaded rows near Cornerstone's 1,000,000-record export
     cap means the export was truncated (FAIL: counts can't be trusted)

History appends to data\\runlogs\\export_audit\\export_audit_history.csv
(drives the row-drop check and gives a durable per-day record of every feed).

Run standalone:  python scripts\\export_audit.py   (exit code 1 if any FAIL)
Used by:         build_morning_review.py — "Export Audit" sheet + Start Here flags
"""

from __future__ import annotations

import csv
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

import onedrive_paths as op                  # noqa: E402

HISTORY_DIR = op.RUNLOGS_DIR / "export_audit"
HISTORY_CSV = HISTORY_DIR / "export_audit_history.csv"

ROW_DROP_PCT = 0.20        # WARN when a feed loses more than this vs last audit
ROW_DROP_MIN_PREV = 1000   # ...but only for feeds that had at least this many
CAP_ROWS = 995_000         # near the 1,000,000 Cornerstone export cap = truncated


def _newest(folder: Path, pattern: str) -> Path | None:
    files = [p for p in folder.glob(pattern) if not p.name.startswith("~$")]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


# Feed -> (file resolver, raw table, cadence). Cadence: "daily" = expect a
# weekday-morning overwrite; "weekly" = HR's bi-weekly-ish drop (9-day grace);
# None = no schedule expectation (checked for presence + load only).
FEEDS = [
    ("Master Wave File",
     lambda: op.MASTER_WAVE_PATH, "master", "daily"),
    ("HR.xlsx (raw)",
     lambda: op.RAW_HR_PATH, "hr", "weekly"),
    ("MVP User Mappings",
     lambda: op.RAW_MVP_USER_MAPPINGS, "mvp", "daily"),
    ("Epic Curriculum Status",
     lambda: _newest(op.RAW_EPIC_STATUS_DIR, "*.xlsx"), "epic_status", "daily"),
    ("Epic Team Member Lookup",
     lambda: op.RAW_EPIC_TM_DIR / "Epic Team Member Lookup.xlsx",
     "epic_lookup", "daily"),
    ("Cornerstone Enterprise Training Report",
     lambda: _newest(op.RAW_CORNERSTONE_DIR, "Enterprise_Training_Report*.xlsx"),
     "cornerstone", "daily"),
    ("Epic Class Schedule",
     lambda: _newest(op.RAW_EPIC_CLASS_SCHEDULES_DIR, "epic_class_schedule*.xlsx"),
     "epic_class_schedule", "daily"),
    ("Epic Unregistered Sessions",
     lambda: _newest(op.ONEDRIVE_ROOT / "data" / "raw" / "epic" / "unregistered",
                     "Unregistered Sessions by User*.xlsx"),
     "epic_unregistered", "daily"),
    ("Reference Lists",
     lambda: op.REF_LISTS_PATH, "ref_leaders", None),
]


def _sql_state(engine, table: str) -> tuple[datetime | None, str, int | None]:
    """(_loaded_at, _source_file, row count) for raw.<table>, or (None, '', None)."""
    try:
        df = pd.read_sql(
            f"SELECT MAX(_loaded_at) AS la, MAX(_source_file) AS sf, "
            f"COUNT(*) AS n FROM raw.{table}", engine)
        la = pd.to_datetime(df.at[0, "la"]) if df.at[0, "la"] is not None else None
        return (la.to_pydatetime() if la is not None else None,
                str(df.at[0, "sf"] or ""), int(df.at[0, "n"]))
    except Exception:
        return None, "", None


def _prev_rows() -> dict[str, int]:
    """Feed -> SQL row count at the most recent prior audit run."""
    if not HISTORY_CSV.exists():
        return {}
    prev: dict[str, int] = {}
    with HISTORY_CSV.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):   # chronological; last write per feed wins
            try:
                prev[row["Feed"]] = int(row["SqlRows"])
            except (KeyError, ValueError):
                continue
    return prev


def _append_history(rows: list[dict]) -> None:
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    fields = ["RunAt", "Feed", "File", "FileModified", "SqlLoadedAt",
              "SqlRows", "Status", "Detail"]
    new_file = not HISTORY_CSV.exists()
    with HISTORY_CSV.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        if new_file:
            w.writeheader()
        run_at = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
        for r in rows:
            w.writerow({**r, "RunAt": run_at})


def audit(engine) -> pd.DataFrame:
    """Run every check; returns one row per feed (worst finding wins the Status).
    Appends the results to the history CSV."""
    prev_rows = _prev_rows()
    today = date.today()
    is_weekday = today.weekday() < 5
    out: list[dict] = []

    for feed, resolver, table, cadence in FEEDS:
        path = resolver()
        loaded_at, source_file, sql_rows = _sql_state(engine, table)
        findings: list[tuple[str, str]] = []   # (severity, message)

        if path is None or not path.exists():
            findings.append(("FAIL", "export file MISSING on disk"))
            mtime = None
        else:
            mtime = datetime.fromtimestamp(path.stat().st_mtime)
            if cadence == "daily" and is_weekday and mtime.date() < today:
                findings.append(("WARN",
                                 f"not updated today (last drop {mtime:%Y-%m-%d %H:%M})"))
            elif cadence == "weekly" and (today - mtime.date()).days > 9:
                findings.append(("WARN",
                                 f"stale — last drop {mtime:%Y-%m-%d} "
                                 f"({(today - mtime.date()).days}d ago)"))

        if sql_rows is None:
            findings.append(("FAIL", f"raw.{table} missing/unreadable in SQL"))
        elif sql_rows == 0:
            findings.append(("FAIL", f"raw.{table} is EMPTY in SQL"))
        else:
            if (mtime is not None and loaded_at is not None
                    and mtime > loaded_at):
                findings.append(("WARN",
                                 "newer file on disk than SQL load "
                                 f"(file {mtime:%H:%M}, loaded {loaded_at:%Y-%m-%d %H:%M}) "
                                 "— next auto-refresh should pick it up"))
            prev = prev_rows.get(feed)
            if (prev is not None and prev >= ROW_DROP_MIN_PREV
                    and sql_rows < prev * (1 - ROW_DROP_PCT)):
                findings.append(("WARN",
                                 f"row count dropped {prev:,} -> {sql_rows:,} "
                                 "since last audit — verify the export"))
            if table == "cornerstone" and sql_rows >= CAP_ROWS:
                findings.append(("FAIL",
                                 f"{sql_rows:,} rows loaded — at/near the 1,000,000 "
                                 "Cornerstone export cap: export is TRUNCATED, "
                                 "re-pull before trusting counts"))

        status = ("FAIL" if any(s == "FAIL" for s, _ in findings)
                  else "WARN" if findings else "OK")
        out.append({
            "Feed": feed,
            "Status": status,
            "File": path.name if path is not None and path.exists() else "",
            "FileModified": f"{mtime:%Y-%m-%d %H:%M}" if mtime else "",
            "SqlLoadedAt": f"{loaded_at:%Y-%m-%d %H:%M}" if loaded_at else "",
            "SqlRows": sql_rows if sql_rows is not None else "",
            "Detail": "; ".join(msg for _, msg in findings) or "current and loaded",
        })

    _append_history(out)
    return pd.DataFrame(out)


def main() -> int:
    from refresh import CONN
    from sqlalchemy import create_engine
    engine = create_engine(CONN)
    try:
        df = audit(engine)
    finally:
        engine.dispose()
    print(df.to_string(index=False))
    fails = (df["Status"] == "FAIL").sum()
    warns = (df["Status"] == "WARN").sum()
    print(f"\n{len(df)} feeds: {fails} FAIL, {warns} WARN "
          f"(history: {HISTORY_CSV})")
    return 1 if fails else 0


if __name__ == "__main__":
    import activity_log
    with activity_log.track_run("export_audit.py"):
        rc = main()
    sys.exit(rc)
