# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_morning_review.py — assemble the ONE workbook the analyst opens each morning.

Read-only against the Master (never touches it). Since 2026-07-28 this is the
post-run report AND personal dashboard for the automated morning refresh
(daily_refresh.py auto): open it and know everything — what ran, what failed,
what changed, what needs review, and where every file lives.

Sheets (stable name Morning Review.xlsx, overwritten each run; dated copy in
morning_review\\archive\\<YYYY-MM>\\):

  Start Here            run status + errors + export flags + rows to review
  Today's Run           every pipeline step today: OK / FAILED / SKIPPED + duration
  Files Updated         each deliverable, updated-today flag, timestamp, link
  Export Audit          every export feed end-to-end: on disk, fresh, loaded in
                        SQL, row counts sane (export_audit.py; OK / WARN / FAIL)
  HR Leader Changes     preview CSV from review_hr_leader_changes.py
  Master vs Sources     SQL report.master_vs_sources (only rows with a Differs flag)
  Exceptions            SQL report.exceptions (in-scope, missing from Epic/HR)
  Missing Wave - HR     SQL report.hr_missing_from_wave
  Missing Wave - Epic   SQL report.epic_missing_from_wave
  Wave vs Sources       SQL report.wave_missing_from_sources (flags flattened)
  Mapping Gaps          SQL report.mapping_gaps (values not in the reference lists)
  Wave Change Requests  raw.wave_change_requests (ALL active form files, combined)
  Yesterday Changes     SQL report.daily_changes at the latest snapshot
  Run History           cumulative Daily Refresh log (from master_runlog.xlsx)
  Activity Log          last 7 days of per-script runs (from activity_log.xlsx)
  Directory             persistent map of every folder/file, clickable links
  My ...                any sheet whose name starts with "My " is YOURS: its
                        values are carried forward on every rebuild (add notes,
                        to-dos, reference lists — they survive; heavy formatting
                        does not, values only)

Runs as the LAST apply step (review_refresh) so it reports the finished run;
also still runs in a manual prep for the pre-apply review flow.

Usage:
    python scripts\\build_morning_review.py
"""

from __future__ import annotations

import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

import activity_log                          # noqa: E402
import export_audit                          # noqa: E402
import onedrive_paths as op                  # noqa: E402
from refresh import CONN                     # noqa: E402
from report_xlsx import write_formatted_workbook      # noqa: E402
from sqlalchemy import create_engine         # noqa: E402

# review_hr_leader_changes.py writes its preview CSV next to Morning Review
# in Main Reports (moved from the retired morning_review folder 2026-07-30).
HR_PREVIEW_CSV = op.HR_PREVIEW_CSV_PATH

DAILY_REFRESH_LOG_DIR = op.RUNLOGS_DIR / "daily_refresh"
MASTER_REPORTS_DIR = op.MAIN_REPORTS_DIR
TRACKER_PATH = MASTER_REPORTS_DIR / "RCM Training Tracker.xlsx"
DAILY_LOG_PATH = MASTER_REPORTS_DIR / "RCM Training Daily Log.xlsx"
NEW_MEMBERS_DIR = op.NEW_MEMBERS_DIR      # local since 2026-08-11 (follows the Master)
EXEC_BRIEF_PATH = op.DASHBOARDS_DIR / "RCM W3 Executive Brief.html"

# User sheets carried forward across rebuilds (values only).
PERSONAL_PREFIX = "my "


# Shown as the "What It Means" column on Start Here (kept in sync with the
# review table in DAILY_REFRESH.md).
SHEET_GUIDE = {
    "Today's Run":
        "Every pipeline step today with OK/FAILED/SKIPPED and duration. A "
        "FAILED row = fix, then rerun that stage with --from <step>.",
    "Files Updated":
        "Each deliverable with an Updated Today? flag and a clickable link. "
        "A 'No' on a weekday morning means its step didn't run/failed.",
    "HR Leader Changes":
        "Leader changes the HR apply step made/will make. Scan; if wrong, fix "
        "HR/aliases and rerun.",
    "Master vs Sources":
        "Rows where the Master disagrees with HR (AVP/VP) or MVP (job role). "
        "Spot-check; apply aligns most.",
    "Exceptions":
        "2 checks: Centralized on Epic TM Lookup but NOT on the wave file "
        "(hard problem — should be 0); in-scope but missing from HR (bad UID).",
    "Missing Wave - HR":
        "People in HR's RCM population with no wave assignment. Decide who to add.",
    "Missing Wave - Epic":
        "Epic team members with no wave assignment. Decide who to add / follow up.",
    "Wave vs Sources":
        "Wave (Master) people NOT found in HR / Epic / MVP (see Missing From "
        "col). Verify terms/transfers.",
    "Mapping Gaps":
        "Master values not in the reference lists (BU/vendor/leader). Fix "
        "wave_reference_lists.xlsx or the Master value.",
    "Wave Change Requests":
        "Every open request form in the folder (see _source_file). Disposition "
        "each; archive a finished form and it drops off tomorrow.",
    "Yesterday Changes":
        "What actually changed at the last snapshot. Sanity-check yesterday's apply.",
    "Run History":
        "Cumulative daily-refresh runs (newest first). The durable log lives "
        "in master_runlog.xlsx — this is a view of it.",
    "Activity Log":
        "Last 7 days of individual script runs (newest first), from "
        "activity_log.xlsx.",
    "Export Audit":
        "Every export feed checked end-to-end: file on disk, dropped on "
        "schedule, loaded into SQL, row counts sane, Enterprise under the 1M "
        "cap. WARN/FAIL rows need a look; OK means current and loaded.",
    "Directory":
        "Persistent map of everything — reports, raw drops, backups, logs. "
        "Links are clickable. Backup rows explain how to revert.",
}


def _note(text: str) -> pd.DataFrame:
    return pd.DataFrame({"Note": [text]})


def _is_note(df: pd.DataFrame) -> bool:
    return list(df.columns) == ["Note"]


def _sql(engine, sql: str, empty_note: str) -> pd.DataFrame:
    try:
        df = pd.read_sql(sql, engine)
    except Exception as e:  # surface the problem on the sheet, don't die
        return _note(f"Query failed: {e}")
    return df if len(df) else _note(empty_note)


def _audit_start_rows(audit_df: pd.DataFrame) -> list[dict]:
    """Start Here lines from the export audit: a one-line summary, plus one
    line per flagged feed so problems are visible without changing sheets."""
    flagged = audit_df[audit_df["Status"] != "OK"]
    if flagged.empty:
        summary = f"All {len(audit_df)} feeds OK — on disk, fresh, and loaded"
    else:
        summary = (f"{len(flagged)} of {len(audit_df)} feeds flagged "
                   "— see Export Audit sheet")
    rows = [{"Item": "Export audit", "Detail": summary}]
    for r in flagged.itertuples():
        rows.append({"Item": f"Export {r.Status}",
                     "Detail": f"{r.Feed}: {r.Detail}"})
    return rows


# ---------------------------------------------------------------------------
# Today's Run — parse today's daily_refresh stage logs
# ---------------------------------------------------------------------------
_RE_RUN = re.compile(r"^===== (prep|apply) run (\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) =====$")
_RE_STEP = re.compile(r"^===== (\w+): (.+) =====$")
_RE_RESULT = re.compile(r"^(OK |X {2})(\w+) \((\d+(?:\.\d+)?)s\)$")
_RE_SKIP = re.compile(r"^(\w+): skipped \((.+)\)$")


def _parse_stage_log(stage: str) -> tuple[list[dict], str | None]:
    """Return (step rows, last run start ts) for today's <stage> log. Each
    step keeps its LAST occurrence, so re-runs/resumes show the final state."""
    log = (DAILY_REFRESH_LOG_DIR / f"{date.today():%Y-%m}"
           / f"daily_refresh_{date.today():%Y.%m.%d}_{stage}.log")
    if not log.exists():
        return [], None
    steps: dict[str, dict] = {}
    order: list[str] = []
    last_run_ts = None
    known: set[str] = set()
    for raw in log.read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.rstrip()
        m = _RE_RUN.match(line)
        if m and m.group(1) == stage:
            last_run_ts = m.group(2)
            continue
        m = _RE_STEP.match(line)
        if m:
            name, label = m.group(1), m.group(2)
            known.add(name)
            if name not in steps:
                order.append(name)
            steps[name] = {"Stage": stage, "Step": name, "Result": "in progress",
                           "Duration (s)": "", "Detail": label}
            continue
        m = _RE_RESULT.match(line)
        if m and m.group(2) in known:
            name = m.group(2)
            steps[name]["Result"] = "OK" if m.group(1).startswith("OK") else "FAILED"
            steps[name]["Duration (s)"] = m.group(3)
            continue
        m = _RE_SKIP.match(line)
        if m:
            name = m.group(1)
            if name not in steps:
                order.append(name)
            steps[name] = {"Stage": stage, "Step": name, "Result": "SKIPPED",
                           "Duration (s)": "", "Detail": m.group(2)}
    rows = [steps[n] for n in order]
    # The build we're inside right now can't have its OK line yet.
    for r in rows:
        if r["Step"] == "review_refresh" and r["Result"] == "in progress":
            r["Result"] = "OK"
            r["Detail"] = "this Morning Review rebuild"
    return rows, last_run_ts


def _todays_run() -> tuple[pd.DataFrame, list[str]]:
    """(Today's Run sheet, status lines for Start Here)."""
    all_rows, status = [], []
    for stage in ("prep", "apply"):
        rows, run_ts = _parse_stage_log(stage)
        all_rows.extend(rows)
        if not rows:
            status.append(f"{stage.upper()}: did not run today")
            continue
        failed = [r["Step"] for r in rows if r["Result"] == "FAILED"]
        hanging = [r["Step"] for r in rows if r["Result"] == "in progress"]
        ok = sum(1 for r in rows if r["Result"] == "OK")
        skipped = sum(1 for r in rows if r["Result"] == "SKIPPED")
        when = f" (started {run_ts[11:16]})" if run_ts else ""
        if failed:
            status.append(f"{stage.upper()}: FAILED at {', '.join(failed)}{when} — "
                          f"fix, then: python scripts\\daily_refresh.py {stage} "
                          f"--from {failed[0]}")
        elif hanging:
            status.append(f"{stage.upper()}: incomplete — stopped during "
                          f"{', '.join(hanging)}{when}")
        else:
            status.append(f"{stage.upper()}: Success{when} — {ok} steps OK, "
                          f"{skipped} skipped")
    df = (pd.DataFrame(all_rows) if all_rows
          else _note("No daily_refresh run logged today."))
    return df, status


# ---------------------------------------------------------------------------
# Files Updated
# ---------------------------------------------------------------------------
def _file_row(label: str, path: Path) -> dict:
    if path.exists():
        m = datetime.fromtimestamp(path.stat().st_mtime)
        return {"Deliverable": label,
                "Updated Today?": "Yes" if m.date() == date.today() else "No",
                "Last Updated": f"{m:%Y-%m-%d %H:%M}",
                "Location": str(path)}
    return {"Deliverable": label, "Updated Today?": "—",
            "Last Updated": "MISSING", "Location": str(path)}


def _files_updated() -> pd.DataFrame:
    today = f"{date.today():%Y-%m-%d}"
    archive_month = op.MASTER_ARCHIVE_DIR / f"{date.today():%Y-%m}"
    reports_archive_month = op.MAIN_REPORTS_BACKUPS_DIR / f"{date.today():%Y-%m}"
    rows = [
        _file_row("Master Wave File", op.MASTER_WAVE_PATH),
        _file_row("RCM Wave Team File (share this)", op.TEAM_WAVE_FILE_PATH),
        _file_row("RCM Training Tracker", TRACKER_PATH),
        _file_row("RCM Training Daily Log", DAILY_LOG_PATH),
        _file_row("W3 Executive Brief (html)", EXEC_BRIEF_PATH),
        _file_row("Metrics summary (trend rows)", op.RUNLOGS_DIR / "metrics_summary.xlsx"),
        _file_row("Master backup (today's revert point)",
                  archive_month / f"{op.MASTER_WAVE_PATH.stem}.bak.{today}.xlsx"),
        _file_row("Tracker backup (today's revert point)",
                  reports_archive_month / f"RCM Training Tracker.bak.{today}.xlsx"),
        _file_row("Daily Log backup (today's revert point)",
                  reports_archive_month / f"RCM Training Daily Log.bak.{today}.xlsx"),
        _file_row("New members added (today's CSV)",
                  NEW_MEMBERS_DIR / f"New_Members_Added_{today}.csv"),
    ]
    if op.GAP_REPORTS_DIR.is_dir():
        for p in sorted(op.GAP_REPORTS_DIR.glob("*.xlsx")):
            row = _file_row(f"Boss gap report: {p.stem}", p)
            # The HR-missing-from-wave report is WEEKLY (Mondays) — a "No" on
            # other days is expected, so only flag it once it's genuinely stale.
            if "HR Missing From Wave" in p.stem and row["Updated Today?"] == "No":
                m = datetime.fromtimestamp(p.stat().st_mtime)
                row["Updated Today?"] = ("No (weekly — OK)"
                                         if (datetime.now() - m).days <= 8
                                         else "No (weekly — STALE)")
            rows.append(row)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Run History + Activity Log (views of the durable runlog workbooks)
# ---------------------------------------------------------------------------
def _run_history() -> pd.DataFrame:
    path = op.RUNLOGS_DIR / "master_runlog.xlsx"
    try:
        df = pd.read_excel(path, sheet_name="Daily Refresh", engine="openpyxl")
        return df.iloc[::-1].reset_index(drop=True) if len(df) else \
            _note("No daily-refresh runs logged yet.")
    except Exception as e:
        return _note(f"Could not read {path.name}: {e}")


def _activity_log_week() -> pd.DataFrame:
    path = op.RUNLOGS_DIR / "activity_log.xlsx"
    try:
        df = pd.read_excel(path, sheet_name="Activity", engine="openpyxl")
        if not len(df):
            return _note("Activity log is empty.")
        ts = pd.to_datetime(df["Timestamp"], errors="coerce")
        recent = df[ts >= (datetime.now() - timedelta(days=7))]
        if not len(recent):
            return _note("No script runs in the last 7 days.")
        return recent.iloc[::-1].reset_index(drop=True)
    except Exception as e:
        return _note(f"Could not read {path.name}: {e}")


# ---------------------------------------------------------------------------
# Directory — the persistent map (regenerated, content is stable)
# ---------------------------------------------------------------------------
def _directory() -> pd.DataFrame:
    d = [
        ("Daily outputs", "Master Wave File", "SOURCE OF TRUTH roster. Never share directly.",
         op.MASTER_WAVE_PATH),
        ("Daily outputs", "RCM Wave Team File", "Shareable copy of the Master — share THIS.",
         op.TEAM_WAVE_FILE_PATH),
        ("Daily outputs", "RCM Training Tracker", "Leader dropdown tracker + Registration Audit.",
         TRACKER_PATH),
        ("Daily outputs", "RCM Training Daily Log", "Daily/weekly trend history since 07/06.",
         DAILY_LOG_PATH),
        ("Daily outputs", "Morning Review (this file)", "Rebuilt at the end of every morning run.",
         op.MORNING_REVIEW_PATH),
        ("Daily outputs", "Boss gap reports", "Daily Epic-not-in-HR + weekly HR-missing-from-wave.",
         op.GAP_REPORTS_DIR),
        ("Daily outputs", "W3 Executive Brief", "Flagship HTML brief (10 AM task rebuilds it).",
         EXEC_BRIEF_PATH),
        ("Daily outputs", "Dashboards folder", "All HTML dashboards (exec brief + variants).",
         op.DASHBOARDS_DIR),
        ("Daily outputs", "New members added", "One CSV per day the apply appended people.",
         NEW_MEMBERS_DIR),
        ("Daily outputs", "SQL pulls", "Ad-hoc query exports (sql\\query.py --excel).",
         op.SQL_PULLS_DIR),
        ("Daily outputs", "CSI Interactions Tracker", "Rolling tracker (csi_tracker.py).",
         op.CSI_TRACKER_PATH),
        ("Raw drops (inbound)", "HR export", "HR.xlsx lands here (bi-weekly-ish).",
         op.RAW_HR_PATH.parent),
        ("Raw drops (inbound)", "MVP exports", "User MappingsData.csv + role/job files (daily ~9:00-9:30).",
         op.RAW_MVP_DIR),
        ("Raw drops (inbound)", "Epic exports", "TM lookup / status details / class schedules.",
         op.RAW_EPIC_TM_DIR.parent),
        ("Raw drops (inbound)", "Cornerstone Enterprise export", "The only loading Cornerstone feed.",
         op.RAW_CORNERSTONE_DIR),
        ("Raw drops (inbound)", "Wave change request forms", "Open forms load to SQL; archive when done.",
         op.RAW_WAVE_CHANGE_REQUESTS_DIR),
        ("Raw drops (inbound)", "CSI inbox", "Drop daily CSI PDFs here, run csi_tracker.py.",
         op.CSI_INBOX_DIR),
        ("References", "Wave reference lists", "Leader allowlist + lookup lists (raw.ref_*).",
         op.REF_LISTS_PATH),
        ("References", "Source registry", "One row per export type the SQL refresh loads.",
         op.SOURCE_REGISTRY_PATH),
        ("Backups / revert", "Master + report backups", "Daily .bak.<date> copies, monthly folders. "
         "REVERT = close Excel, copy the .bak over the live file.",
         op.MASTER_ARCHIVE_DIR),
        ("Backups / revert", "Tracker + Daily Log + Morning Review backups",
         "Daily pre-rebuild .bak.<date> copies + dated Morning Review archive. "
         "REVERT = copy the .bak over the live file.",
         op.MAIN_REPORTS_BACKUPS_DIR),
        ("Backups / revert", "Master DATA snapshots", "Values-only DATA sheet after every edit "
         "(git-like history).",
         op.WAVE_REPOSITORY_DIR),
        ("Backups / revert", "Team File dated copies", "One per publish.",
         op.WAVE_DISTRIBUTION_DIR),
        ("Backups / revert", "SQL history snapshots", "Dated CSV per history table per day "
         "(backload_history rebuilds SQL from these).",
         op.HISTORY_SNAPSHOTS_DIR),
        ("Logs", "Daily refresh logs", "Full console output per stage per day.",
         DAILY_REFRESH_LOG_DIR),
        ("Logs", "Master runlog", "Detailed per-script metric tabs incl. Daily Refresh summary.",
         op.RUNLOGS_DIR / "master_runlog.xlsx"),
        ("Logs", "Activity log", "One row per script run (did it work, one-liner).",
         op.RUNLOGS_DIR / "activity_log.xlsx"),
        ("Tools", "Scripts folder (local)", "All pipeline scripts; daily_refresh.py is the runbook.",
         ROOT / "scripts"),
        ("Tools", "SQL folder (local)", "refresh.py, query.py, view definitions, DATA_DICTIONARY.md.",
         ROOT / "sql"),
    ]
    return pd.DataFrame(
        [{"Section": s, "Item": i, "What It Is": w, "Location": str(p)}
         for s, i, w, p in d])


# ---------------------------------------------------------------------------
# Personal sheets ("My ...") carried forward from the existing workbook
# ---------------------------------------------------------------------------
def _preserved_personal_sheets(taken_names: set[str]) -> dict[str, pd.DataFrame]:
    if not op.MORNING_REVIEW_PATH.exists():
        return {}
    out: dict[str, pd.DataFrame] = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(op.MORNING_REVIEW_PATH, read_only=True,
                                    data_only=True)
        try:
            for name in wb.sheetnames:
                if not name.lower().startswith(PERSONAL_PREFIX) or name in taken_names:
                    continue
                rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
                rows = [r for r in rows if any(v is not None for v in r)]
                if not rows:
                    continue
                header, seen = [], set()
                for i, v in enumerate(rows[0]):
                    h = str(v) if v is not None else f"Col{i + 1}"
                    while h in seen:
                        h += "_"
                    seen.add(h)
                    header.append(h)
                out[name] = pd.DataFrame(rows[1:], columns=header)
        finally:
            wb.close()
    except Exception as e:
        print(f"  WARNING: could not carry forward personal sheets ({e})")
    return out


def _save_with_retry(sheets: dict, link_cols: dict, stamps: dict) -> Path:
    """The workbook may be open in Excel during the unattended morning run —
    retry briefly, then fall back to a side-by-side file rather than failing
    the whole pipeline."""
    for attempt in range(3):
        try:
            write_formatted_workbook(sheets, op.MORNING_REVIEW_PATH, link_cols, stamps)
            return op.MORNING_REVIEW_PATH
        except PermissionError:
            if attempt < 2:
                print("  Morning Review.xlsx is open in Excel — retrying in 15s...")
                time.sleep(15)
    fallback = op.MAIN_REPORTS_DIR / "Morning Review (LATEST - close the other copy).xlsx"
    write_formatted_workbook(sheets, fallback, link_cols, stamps)
    print(f"WARNING: Morning Review.xlsx was open in Excel; fresh copy saved as "
          f"{fallback.name}. Close the open copy — tomorrow's run will overwrite "
          f"the stable file again.")
    return fallback


def main() -> None:
    engine = create_engine(CONN)
    try:
        run_df, run_status = _todays_run()

        sheets: dict[str, pd.DataFrame] = {}

        if HR_PREVIEW_CSV.exists():
            hr_prev = pd.read_csv(HR_PREVIEW_CSV, dtype=str)
            sheets["HR Leader Changes"] = hr_prev if len(hr_prev) else \
                _note("No HR leader changes pending.")
        else:
            sheets["HR Leader Changes"] = _note(
                "hr_leader_changes_preview.csv not found — run review_hr_leader_changes.py.")

        sheets["Master vs Sources"] = _sql(
            engine,
            "SELECT * FROM report.master_vs_sources "
            "WHERE AVP_Differs = 1 OR VP_Differs = 1 OR JobRole_Differs = 1",
            "No Master-vs-source differences today.")
        sheets["Exceptions"] = _sql(
            engine, "SELECT * FROM report.exceptions ORDER BY Issue, Leader",
            "No data-quality exceptions today.")
        # since 2026-07-27 these three read the SQL views directly (fresher
        # than the retired dated gap workbooks, which lagged a day behind)
        sheets["Missing Wave - HR"] = _sql(
            engine, "SELECT * FROM report.hr_missing_from_wave",
            "Nobody in HR under our leaders is missing from the wave file.")
        sheets["Missing Wave - Epic"] = _sql(
            engine, "SELECT * FROM report.epic_missing_from_wave",
            "Nobody Epic-centralized is missing from the wave file.")
        sheets["Wave vs Sources"] = _sql(
            engine,
            "SELECT 'HR' AS [Missing From], UniversalID, FullName, Wave, "
            "Leader, TrainingNeeded, Departed "
            "FROM report.wave_missing_from_sources WHERE NotInHR = 1 "
            "UNION ALL SELECT 'Epic', UniversalID, FullName, Wave, Leader, "
            "TrainingNeeded, Departed "
            "FROM report.wave_missing_from_sources WHERE NotInEpic = 1 "
            "UNION ALL SELECT 'MVP', UniversalID, FullName, Wave, Leader, "
            "TrainingNeeded, Departed "
            "FROM report.wave_missing_from_sources WHERE NotInMVP = 1 "
            "ORDER BY [Missing From], Leader, FullName",
            "Every wave user is findable in HR, Epic, and MVP.")

        sheets["Mapping Gaps"] = _sql(
            engine,
            "SELECT * FROM report.mapping_gaps ORDER BY Issue, People DESC",
            "All Master values map cleanly to the reference lists.")
        sheets["Wave Change Requests"] = _sql(
            engine,
            "SELECT * FROM raw.wave_change_requests",
            "No wave change requests loaded.")
        sheets["Yesterday Changes"] = _sql(
            engine,
            "SELECT * FROM report.daily_changes WHERE SnapshotDate = "
            "(SELECT MAX(SnapshotDate) FROM history.roster_daily) "
            "ORDER BY BecameRegistered DESC, Leader, FullName",
            "No day-over-day changes at the latest snapshot (or only one snapshot so far).")

        files_df = _files_updated()
        directory_df = _directory()
        try:
            audit_df = export_audit.audit(engine)
        except Exception as e:   # audit must never sink the review build
            audit_df = _note(f"Export audit failed to run: {e}")

        # ----- Start Here: the summary hub -----
        total = sum(0 if _is_note(df) else len(df) for df in sheets.values())
        failed_steps = [] if _is_note(run_df) else \
            [f"{r.Stage}/{r.Step}" for r in run_df.itertuples() if r.Result == "FAILED"]
        not_updated = [r["Deliverable"] for _, r in files_df.iterrows()
                       if r["Updated Today?"] in ("No", "No (weekly — STALE)")
                       and "backup" not in r["Deliverable"].lower()
                       and "New members" not in r["Deliverable"]]
        today_bak = (op.MASTER_ARCHIVE_DIR / f"{date.today():%Y-%m}"
                     / f"{op.MASTER_WAVE_PATH.stem}.bak.{date.today():%Y-%m-%d}.xlsx")

        start = [{"Item": "Built", "Detail": f"{datetime.now():%Y-%m-%d %H:%M:%S}"}]
        for line in run_status:
            start.append({"Item": "Run status", "Detail": line})
        start.append({"Item": "Errors today",
                      "Detail": ("NONE" if not failed_steps
                                 else "FAILED: " + ", ".join(failed_steps)
                                 + " — see Today's Run")})
        start.append({"Item": "Rows to review (all sheets)", "Detail": str(total)})
        if not_updated:
            start.append({"Item": "Not updated today",
                          "Detail": ", ".join(not_updated) + " — see Files Updated"})
        start.append({"Item": "Revert point",
                      "Detail": (f"Master backup: {today_bak.name} (in "
                                 f"{op.MASTER_ARCHIVE_DIR}\\{date.today():%Y-%m}) — "
                                 "copy over the live file to roll back"
                                 if today_bak.exists()
                                 else "No Master backup taken today (no Master "
                                      "writes yet)")})
        start.append({"Item": "Your pages",
                      "Detail": "Sheets named 'My ...' are carried forward on "
                                "every rebuild — use them for notes/references."})
        start += (_audit_start_rows(audit_df) if not _is_note(audit_df)
                  else [{"Item": "Export audit",
                         "Detail": str(audit_df.at[0, "Note"])}])
        for name in ["Today's Run", "Files Updated", "Export Audit",
                     *sheets.keys(), "Run History", "Activity Log", "Directory"]:
            if name in sheets:
                df = sheets[name]
                n = 0 if _is_note(df) else len(df)
                start.append({"Item": f"Sheet: {name}", "Detail": f"{n} rows to review",
                              "What It Means": SHEET_GUIDE.get(name, "")})
            else:
                start.append({"Item": f"Sheet: {name}", "Detail": "",
                              "What It Means": SHEET_GUIDE.get(name, "")})

        ordered = {"Start Here": pd.DataFrame(start).fillna(""),
                   "Today's Run": run_df,
                   "Files Updated": files_df,
                   "Export Audit": audit_df,
                   **sheets,
                   "Run History": _run_history(),
                   "Activity Log": _activity_log_week(),
                   "Directory": directory_df}
        # Every generated sheet carries a "Last updated" banner in row 1.
        # Personal "My ..." sheets are excluded — they're carried forward
        # verbatim, so a stamp row would turn into data on the next rebuild.
        stamp = f"Last updated: {datetime.now():%Y-%m-%d %H:%M}"
        stamps = {name: stamp for name in ordered}
        ordered.update(_preserved_personal_sheets(set(ordered.keys())))

        link_cols = {"Files Updated": "Location", "Directory": "Location"}

        op.MAIN_REPORTS_BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
        saved_to = _save_with_retry(ordered, link_cols, stamps)
        op.write_dated_copy(saved_to,
                            op.month_subdir(op.MAIN_REPORTS_BACKUPS_DIR),
                            "morning_review")

        print(f"Morning Review built: {total} total rows to review, "
              f"{len(failed_steps)} failed step(s) today -> {saved_to}")
    finally:
        engine.dispose()


if __name__ == "__main__":
    with activity_log.track_run("build_morning_review.py"):
        main()
