# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_metrics_summary.py

READ-ONLY. Appends ONE row to data/runlogs/metrics_summary.xlsx every time
it's run — a trend table (not a per-day snapshot file) so you can see how
the numbers move over time in one place.

Deliberately basic for a first pass, per the user: pulls only from the
Master DATA sheet and whatever the already-migrated scripts have already
computed and written (never recomputes anything those scripts already did).
Structured as small, independent "extract_*" functions returning a flat
{metric: value} dict — bolt on more later without touching the existing ones.

Not yet available (flagged, not faked): Cornerstone/unregistered/no-show
counts, and a clean single "total job role changes" number (only a mixed
change-log entry count exists today) — add once those scripts migrate.

Run:
    python scripts/build_metrics_summary.py
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "sql"))
import onedrive_paths
import runlog
import activity_log
from leader_names import CANONICAL_NAMES, NOT_REV_CYCLE, normalize_leader
from refresh import CONN
from sqlalchemy import create_engine, text

METRICS_PATH = onedrive_paths.RUNLOGS_DIR / "metrics_summary.xlsx"
METRICS_TAB = "Metrics"


def _is_yes(v) -> bool:
    """True for boolean True (IsDepartedInactive? is a real Excel bool) or any
    string containing 'yes' (Vendor Yes/No, Fully Registered/Trained are text)."""
    if isinstance(v, bool):
        return v
    return "yes" in str(v).strip().lower()


# ---------------------------------------------------------------------------
# Extractors — each returns a flat {metric: value} dict, or {} if its source
# isn't available yet (e.g. no Master-update run has happened today).
# ---------------------------------------------------------------------------

def extract_master_metrics() -> dict:
    """Read the Master DATA sheet directly: headcounts, by-leader, by-wave,
    training needed, departed, vendor, no-longer-rev-cycle."""
    path = onedrive_paths.MASTER_WAVE_PATH
    if not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["DATA"]
    it = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h is not None else "" for h in next(it)]
    col = {h: i for i, h in enumerate(headers) if h}

    def get(row, name):
        i = col.get(name)
        return row[i] if (i is not None and i < len(row)) else None

    total = 0
    leaders = Counter()
    waves = Counter()
    training_yes = training_no = departed = vendor_yes = no_longer_rev_cycle = 0
    for row in it:
        uid = get(row, "UniversalID")
        if not uid:
            continue
        total += 1
        leader = normalize_leader(get(row, "Leaders"))
        if leader:
            leaders[leader] += 1
            if leader == NOT_REV_CYCLE:
                no_longer_rev_cycle += 1
        wave = str(get(row, "GoLiveWave") or "").strip()
        if wave:
            waves[wave] += 1
        tn = str(get(row, "Training Needed (Yes/No)") or "").strip().lower()
        if tn == "yes":
            training_yes += 1
        elif tn == "no":
            training_no += 1
        if _is_yes(get(row, "IsDepartedInactive?")):
            departed += 1
        if _is_yes(get(row, "Vendor Yes/No")):
            vendor_yes += 1
    wb.close()

    metrics = {
        "Total Team Members": total,
        "Total Vendors": vendor_yes,
        "Total Training Needed": training_yes,
        "Total Departed/Inactive": departed,
        "Total No Longer Rev Cycle": no_longer_rev_cycle,
    }
    for name in CANONICAL_NAMES:
        metrics[f"By Leader: {name}"] = leaders.get(name, 0)
    for wave_name in sorted(waves):
        metrics[f"By Wave: {wave_name}"] = waves[wave_name]
    return metrics


def extract_mvp_update_metrics() -> dict:
    """Latest 'MVP Update' row already computed by update_master_from_mvp.py —
    reused, not recomputed."""
    path = onedrive_paths.RUNLOGS_DIR / "master_runlog.xlsx"
    if not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "MVP Update" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["MVP Update"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if len(rows) < 2:
        return {}
    headers, latest = rows[0], rows[-1]
    r = dict(zip(headers, latest))
    return {
        "MVP: Rows Updated": r.get("Master Rows Updated"),
        "MVP: Rows Not in MVP": r.get("Rows NOT in MVP"),
        "MVP: Wave 1→2": r.get("Wave 1→2"),
        "MVP: Wave 2→3": r.get("Wave 2→3"),
        "MVP: Wave 3→4": r.get("Wave 3→4"),
        "MVP: Wave Changes (Other)": r.get("Wave Changes (Other)"),
    }


def extract_epic_summary_metrics() -> dict:
    """RCM population counts straight from SQL (raw.epic_lookup, the Epic TM
    Lookup load). Metric names unchanged. Replaced the retired dated
    epic_team_member_lookup workbook 2026-07-27 (build step trimmed)."""
    eng = create_engine(CONN)
    try:
        with eng.connect() as c:
            row = c.execute(text(
                "SELECT COUNT(*) AS t, "
                "SUM(CASE WHEN Fully_Registered LIKE '%Yes%' THEN 1 ELSE 0 END) AS r, "
                "SUM(CASE WHEN Fully_Trained LIKE '%Yes%' THEN 1 ELSE 0 END) AS tr "
                "FROM raw.epic_lookup "
                "WHERE Curriculum_Type = 'Revenue Cycle - Centralized'")).one()
    except Exception:
        return {}
    finally:
        eng.dispose()
    return {
        "Epic: Total Team Members": int(row[0] or 0),
        "Epic: Fully Registered": int(row[1] or 0),
        "Epic: Fully Trained": int(row[2] or 0),
    }


def extract_missing_from_wave_metrics() -> dict:
    """Gap counts straight from the SQL views. Metric names unchanged.
    Replaced the retired dated gap workbooks 2026-07-27 (gaps_epic /
    gaps_wave steps trimmed; gaps_hr workbook still published weekly but the
    metric reads the same view it exports)."""
    eng = create_engine(CONN)
    try:
        with eng.connect() as c:
            hr_n = c.execute(text(
                "SELECT COUNT(*) FROM report.hr_missing_from_wave")).scalar()
            epic_n = c.execute(text(
                "SELECT COUNT(*) FROM report.epic_missing_from_wave")).scalar()
            w = c.execute(text(
                "SELECT SUM(CASE WHEN NotInHR = 1 THEN 1 ELSE 0 END) AS h, "
                "SUM(CASE WHEN NotInEpic = 1 THEN 1 ELSE 0 END) AS e, "
                "SUM(CASE WHEN NotInMVP = 1 THEN 1 ELSE 0 END) AS m "
                "FROM report.wave_missing_from_sources")).one()
    except Exception:
        return {}
    finally:
        eng.dispose()
    return {
        "Total New Members (Missing From HR)": int(hr_n or 0),
        "Total New Members (Missing From Epic)": int(epic_n or 0),
        "Wave Users Not in HR": int(w[0] or 0),
        "Wave Users Not in Epic": int(w[1] or 0),
        "Wave Users Not in MVP": int(w[2] or 0),
    }


EXTRACTORS = [
    extract_master_metrics,
    extract_mvp_update_metrics,
    extract_epic_summary_metrics,
    extract_missing_from_wave_metrics,
]

NOT_YET_AVAILABLE = [
    "Cornerstone (unregistered/no-show) counts — add once combine_unregistered_sessions.py "
    "/ build_noshow_from_cornerstone_ent.py migrate",
    "Total Job Role Changes as a single clean number — only a mixed change-log "
    "entry count exists today",
]


def main() -> int:
    metrics: dict = {}
    for fn in EXTRACTORS:
        result = fn()
        if not result:
            print(f"  (skipped {fn.__name__} — source not available)")
        metrics.update(result)

    if not metrics:
        print("ERROR: no metric sources were available. Nothing to write.")
        return 1

    from datetime import datetime
    headers = ["Run Timestamp"] + list(metrics.keys())
    row = [datetime.now().strftime("%Y-%m-%d %H:%M:%S")] + list(metrics.values())
    runlog.append(METRICS_TAB, headers, row, path=METRICS_PATH)

    print(f"\n{len(metrics)} metric(s) written to {METRICS_PATH}")
    for k, v in metrics.items():
        print(f"  {k}: {v}")
    if NOT_YET_AVAILABLE:
        print("\nNot yet available (add later):")
        for note in NOT_YET_AVAILABLE:
            print(f"  - {note}")
    return 0


if __name__ == "__main__":
    with activity_log.track_run("build_metrics_summary.py"):
        sys.exit(main())
