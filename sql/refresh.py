# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
refresh.py  —  RUN DAILY  (the whole job is: python sql\\refresh.py)

What it does:
  Reads the latest export from each source and (re)loads it into the
  matching raw.* table in SQL Server. Tables are created automatically,
  so there is no table DDL to maintain by hand.

  Sources (newest file of each type, read from OneDrive) -> raw tables:
    Master Wave File   -> raw.master        (DATA sheet)   = who's in / source of truth
    HR.xlsx  (raw)     -> raw.hr             = full 67-col HR detail
    HR_cleaned (proc.) -> raw.hr_cleaned     = resolved leader chain (AVP/VP/SVP/Leaders)
    MVP User Mappings  -> raw.mvp            = job roles
    MVP Role Mappings  -> raw.mvp_roles      = dept/job-code -> Epic job-category mapping
    MVP Job Categories -> raw.mvp_job_categories = Epic job-category reference
    Cornerstone        -> raw.cornerstone    = enterprise training transcripts
    Epic Curriculum    -> raw.epic_status    = per-user Fully Trained/Registered (per wave)
    Epic TM Lookup     -> raw.epic_lookup    = roster: Wave, training-needed, eligibility
    Epic Class Sched.  -> raw.epic_class_schedule = sessions: date/location/instructor/seats
    Wave Change Reqs   -> raw.wave_change_requests = add/edit/remove requests (form tab)

  Everything lands as text (NVARCHAR). The report views in
  3_report_views.sql do the typing, cleaning, and joining.
  File paths come from scripts/onedrive_paths.py (single source of truth).

  A full run ends with a `snapshot` step: today's report.roster and
  report.epic_not_in_hr are appended to history.* tables (one snapshot per
  day) and mirrored to dated CSVs on OneDrive, so day-over-day trend and
  change queries are possible and the history survives a lost/rebuilt DB.

Usage:
    python sql\\refresh.py                    # load everything + today's history snapshot
    python sql\\refresh.py mvp epic_status     # load only these
    python sql\\refresh.py snapshot            # just (re)stamp today's history
    python sql\\refresh.py backload_history    # rebuild history.* from the OneDrive CSVs
"""

from __future__ import annotations
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.types import NVARCHAR

# ----------------------------------------------------------------------
# CONFIG  — edit these if the server name or file locations change
# ----------------------------------------------------------------------
SERVER   = r".\SQLEXPRESS"
DATABASE = "AnalyticsDB"
ROOT     = Path(__file__).resolve().parent.parent   # yourorg_analytics/

# Reuse the OneDrive path constants + latest_*() helpers the other scripts use.
sys.path.insert(0, str(ROOT / "scripts"))
import onedrive_paths as op   # noqa: E402

CONN = (
    "mssql+pyodbc://@" + SERVER + "/" + DATABASE +
    "?driver=ODBC+Driver+18+for+SQL+Server"
    "&Trusted_Connection=yes&TrustServerCertificate=yes"
)

NVARLEN = 4000  # raw columns are NVARCHAR(4000); plenty for these files


# ----------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------
def clean_col(name: str) -> str:
    """Turn a messy header into a safe SQL column name."""
    s = re.sub(r"[^0-9a-zA-Z]+", "_", str(name).strip()).strip("_")
    if not s:
        s = "col"
    if s[0].isdigit():
        s = "c_" + s
    return s


def dedupe(cols):
    """Append _2, _3 ... to duplicate column names."""
    seen, out = {}, []
    for c in cols:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 1
            out.append(c)
    return out


def prepare(df: pd.DataFrame, source_file: str, extra: dict | None = None) -> pd.DataFrame:
    """Clean columns, add metadata, make everything text/None."""
    df = df.copy()
    df.columns = dedupe([clean_col(c) for c in df.columns])
    # drop completely empty rows
    df = df.dropna(how="all")
    # add metadata
    df["_source_file"] = Path(source_file).name
    df["_loaded_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if extra:
        for k, v in extra.items():
            df[k] = v
    # everything to clean strings (NaN -> None so SQL gets NULL)
    df = df.astype(object).where(pd.notna(df), None)
    df = df.map(lambda v: str(v).strip() if v is not None else None)
    return df


def write_table(engine, table: str, df: pd.DataFrame):
    # Stage-and-swap: load into raw._loading_<table>, then swap it in. An
    # interrupted load can never leave raw.<table> half-full — the old rows
    # survive until the (instant) swap.
    dtype = {c: NVARCHAR(NVARLEN) for c in df.columns}
    tmp = f"_loading_{table}"
    df.to_sql(
        tmp, engine, schema="raw", if_exists="replace", index=False,
        dtype=dtype, chunksize=1000, method=None,
    )
    with engine.begin() as con:
        con.execute(text(f"DROP TABLE IF EXISTS raw.[{table}]"))
        con.execute(text(f"EXEC sp_rename 'raw.[{tmp}]', '{table}'"))
    print(f"   -> raw.{table}: {len(df):,} rows x {len(df.columns)} cols")


# ----------------------------------------------------------------------
# per-source loaders
# ----------------------------------------------------------------------
def build_uids(engine, table: str, col: str):
    """Distinct clean UniversalIDs of raw.<table> -> raw.uids_<table>, with a
    unique clustered index. report.* anti-joins (e.g. report.exceptions) seek
    these tiny key tables instead of scanning the NVARCHAR(4000) heaps — the
    optimizer picks catastrophic nested-loop plans against the raw tables."""
    with engine.begin() as con:
        con.execute(text(f"DROP TABLE IF EXISTS raw.uids_{table}"))
        con.execute(text(
            f"SELECT DISTINCT CAST(UPPER(LTRIM(RTRIM([{col}]))) AS NVARCHAR(450)) AS UniversalID "
            f"INTO raw.uids_{table} FROM raw.[{table}] "
            f"WHERE LTRIM(RTRIM(ISNULL([{col}],''))) <> ''"))
        con.execute(text(
            f"CREATE UNIQUE CLUSTERED INDEX IX_uids_{table} ON raw.uids_{table}(UniversalID)"))


def _newest(paths):
    """Newest path by name (dated/timestamped names sort chronologically), or None."""
    paths = sorted(paths)
    return paths[-1] if paths else None


def _newest_by_mtime(paths):
    """Newest path by modified time, skipping Excel lock files (~$...), or None.
    Use this for exports whose filenames don't sort chronologically (e.g.
    Cornerstone's 12-hour AM/PM timestamps)."""
    paths = [p for p in paths if not p.name.startswith("~$")]
    return max(paths, key=lambda p: p.stat().st_mtime, default=None)


def load_master(engine):
    """Master DATA sheet. If the live file is open in Excel (locked), fall back to
    the newest DATA_*.xlsx snapshot so the refresh still runs."""
    path, note = op.MASTER_WAVE_PATH, ""
    try:
        df = pd.read_excel(path, sheet_name="DATA", dtype=str, engine="calamine")
    except Exception:  # locked / corrupt / mid-sync — calamine + os errors vary
        snap = _newest(op.WAVE_REPOSITORY_DIR.glob("*/DATA_*.xlsx"))
        if snap is None:
            raise
        path, note = snap, "  (live master locked; used snapshot)"
        df = pd.read_excel(path, sheet_name="DATA", dtype=str, engine="calamine")
    print("MASTER  ", path.name, note)
    write_table(engine, "master", prepare(df, path.name))


def load_hr(engine):
    print("HR      ", op.RAW_HR_PATH.name)
    df = pd.read_excel(op.RAW_HR_PATH, sheet_name=0, dtype=str, engine="calamine")
    write_table(engine, "hr", prepare(df, op.RAW_HR_PATH.name))
    build_uids(engine, "hr", "UniversalID")


def load_hr_cleaned(engine):
    """Cleaned HR (resolved leader chain) — the 'All Staff' tab of the newest
    HR_cleaned_YYYY.MM.DD.xlsx produced by scripts/clean_hr.py."""
    path = op.latest_hr_cleaned()
    if path is None:
        print("HR_CLEAN (no HR_cleaned_*.xlsx found, skipped)")
        return
    print("HR_CLEAN", path.name)
    df = pd.read_excel(path, sheet_name="All Staff", dtype=str, engine="calamine")
    write_table(engine, "hr_cleaned", prepare(df, path.name))


def load_mvp(engine):
    print("MVP     ", op.RAW_MVP_USER_MAPPINGS.name)
    df = pd.read_csv(op.RAW_MVP_USER_MAPPINGS, dtype=str, encoding="utf-8-sig", low_memory=False)
    write_table(engine, "mvp", prepare(df, op.RAW_MVP_USER_MAPPINGS.name))


def load_mvp_roles(engine):
    print("MVP_ROLE", op.RAW_MVP_ROLE_MAPPINGS.name)
    df = pd.read_csv(op.RAW_MVP_ROLE_MAPPINGS, dtype=str, encoding="utf-8-sig", low_memory=False)
    write_table(engine, "mvp_roles", prepare(df, op.RAW_MVP_ROLE_MAPPINGS.name))


def load_mvp_job_categories(engine):
    print("MVP_JOBC", op.RAW_MVP_JOB_CATEGORIES.name)
    df = pd.read_csv(op.RAW_MVP_JOB_CATEGORIES, dtype=str, encoding="utf-8-sig", low_memory=False)
    write_table(engine, "mvp_job_categories", prepare(df, op.RAW_MVP_JOB_CATEGORIES.name))


def load_cornerstone(engine):
    path = _newest_by_mtime(op.RAW_CORNERSTONE_DIR.glob("Enterprise_Training_Report_*.xlsx"))
    if path is None:
        print("CORNER   (no Enterprise_Training_Report_*.xlsx found, skipped)")
        return
    print("CORNER  ", path.name)
    # Banner + filter block on top, and its height varies between exports —
    # locate the real header row ("User Full Name" in column A) instead of
    # hardcoding it.
    probe = pd.read_excel(path, sheet_name="Enterprise Training Report", header=None,
                          nrows=30, dtype=str, engine="calamine")
    hits = probe.index[probe[0].astype(str).str.strip() == "User Full Name"]
    if len(hits) == 0:
        raise ValueError(f"no 'User Full Name' header row in the first 30 rows of {path.name}")
    df = pd.read_excel(path, sheet_name="Enterprise Training Report", header=int(hits[0]),
                       dtype=str, engine="calamine")
    write_table(engine, "cornerstone", prepare(df, path.name))


def load_epic_status(engine):
    """Curriculum Status Detail by User - W*.xlsx (one file per wave, no date in the
    name). Keep the newest file per wave by modified time; combine into raw.epic_status."""
    files = list(op.RAW_EPIC_STATUS_DIR.glob("Curriculum Status Detail by User - W*.xlsx"))
    if not files:
        print("EPIC_ST  (no Curriculum Status Detail files found, skipped)")
        return
    pat = re.compile(r"-\s*(W\d)", re.IGNORECASE)
    best: dict[str, Path] = {}
    for f in files:
        m = pat.search(f.name)
        if not m:
            continue
        wave = m.group(1).upper()
        if wave not in best or f.stat().st_mtime > best[wave].stat().st_mtime:
            best[wave] = f
    frames = []
    for wave, f in sorted(best.items()):
        snap = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d")
        print(f"EPIC_ST  {wave}  {f.name}")
        df = pd.read_excel(f, sheet_name="Export", dtype=str, engine="calamine")
        frames.append(prepare(df, f.name, extra={"_wave": wave, "_snapshot_date": snap}))
    write_table(engine, "epic_status", pd.concat(frames, ignore_index=True, sort=False))
    build_uids(engine, "epic_status", "Universal_Id")


def load_epic_lookup(engine):
    path = op.RAW_EPIC_TM_DIR / "Epic Team Member Lookup.xlsx"
    print("EPIC_LK ", path.name)
    df = pd.read_excel(path, sheet_name="Export", dtype=str, engine="calamine")
    write_table(engine, "epic_lookup", prepare(df, path.name))


def load_epic_class_schedule(engine):
    """Newest epic_class_schedule_*.xlsx (dated filenames accumulate; older copies
    live in archive/, which the non-recursive glob ignores)."""
    path = _newest_by_mtime(op.RAW_EPIC_CLASS_SCHEDULES_DIR.glob("epic_class_schedule_*.xlsx"))
    if path is None:
        print("EPIC_CS  (no epic_class_schedule_*.xlsx found, skipped)")
        return
    print("EPIC_CS ", path.name)
    df = pd.read_excel(path, sheet_name="Export", dtype=str, engine="calamine")
    write_table(engine, "epic_class_schedule", prepare(df, path.name))


def load_refs(engine):
    """Reference lists workbook (data/references/wave_reference_lists.xlsx) —
    every sheet loads as raw.ref_<sheetname>. Edit the workbook, rerun this."""
    path = op.REF_LISTS_PATH
    if not path.exists():
        print("REFS     (wave_reference_lists.xlsx not found, skipped)")
        return
    xl = pd.ExcelFile(path, engine="calamine")
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, dtype=str)
        table = "ref_" + clean_col(sheet).lower()
        write_table(engine, table, prepare(df, path.name))


# ----------------------------------------------------------------------
# per-source schemas (cornerstone / epic / mvp / hr / wave)
# ----------------------------------------------------------------------
# Every source system gets its own schema so its tables can be browsed —
# and joined — without hunting through raw.*. The alias views below are
# re-created on every run (so they always match the raw tables' columns);
# the analyst's own views in these schemas are left alone and are backed
# up daily alongside sandbox (see _backup_user_views).
SOURCE_SCHEMAS = ("cornerstone", "epic", "mvp", "hr", "wave")

# View names carry the data stage so raw exports and processed files can't be
# confused: raw_* = the export exactly as received (data\raw\...);
# processed_* = cleaned by the pipeline (data\processed\...); wave.master_data
# is the Master Wave File itself (data\reports\). Each alias also exposes a
# _stage column with the same value.
BUILTIN_SOURCE_ALIASES = [
    # (schema, view name, raw table, stage)
    ("cornerstone", "raw_enterprise_training",   "cornerstone",          "raw"),
    ("epic",        "raw_team_member_lookup",    "epic_lookup",          "raw"),
    ("epic",        "raw_curriculum_status",     "epic_status",          "raw"),
    ("epic",        "raw_class_schedule",        "epic_class_schedule",  "raw"),
    ("mvp",         "raw_user_mappings",         "mvp",                  "raw"),
    ("mvp",         "raw_role_mappings",         "mvp_roles",            "raw"),
    ("mvp",         "raw_job_categories",        "mvp_job_categories",   "raw"),
    ("hr",          "raw_workday_export",        "hr",                   "raw"),
    ("hr",          "processed_cleaned",         "hr_cleaned",           "processed"),
    ("wave",        "master_data",               "master",               "reports"),
    ("wave",        "raw_change_requests",       "wave_change_requests", "raw"),
]


def _ensure_schema(con, schema: str):
    con.execute(text(f"IF SCHEMA_ID('{schema}') IS NULL EXEC('CREATE SCHEMA [{schema}]')"))


def _probe_header_row(path, sheet: str) -> int:
    """0-based header row for exports whose header position moves. Cornerstone
    files carry a variable-length title/criteria block above the header (it
    grows or shrinks when the report's saved filters change), so a fixed
    HeaderRow breaks on the next filter edit — registry rows can say 'auto'
    instead. The header is the first row with 5+ non-empty cells; the
    metadata/criteria rows above it never have more than 2."""
    probe = pd.read_excel(path, sheet_name=sheet if sheet else 0, header=None,
                          nrows=30, engine="calamine")
    for i, row in probe.iterrows():
        if row.notna().sum() >= 5:
            return int(i)
    raise ValueError(f"no header row found in the first 30 rows of {path.name}")


def load_sources(engine):
    """Per-source schemas: (re)create the builtin alias views, then process the
    source registry (data/references/source_registry.xlsx) — each Enabled row
    loads the newest matching file into raw.<source>_<table> and creates the
    <source>.<table> view. Adding a new export needs NO code: drop the file in
    its folder and add one registry row."""
    with engine.begin() as con:
        for schema in SOURCE_SCHEMAS:
            _ensure_schema(con, schema)
        for schema, view, raw_table, stage in BUILTIN_SOURCE_ALIASES:
            if con.execute(text(f"SELECT OBJECT_ID('raw.[{raw_table}]')")).scalar():
                con.execute(text(
                    f"CREATE OR ALTER VIEW [{schema}].[{view}] AS "
                    f"/* {schema}.{view} — alias over raw.{raw_table} "
                    f"(stage: {stage}); auto-created by refresh.py, do not edit */ "
                    f"SELECT *, CONVERT(varchar(9), '{stage}') AS _stage "
                    f"FROM raw.[{raw_table}]"))
    print(f"SOURCES  {len(BUILTIN_SOURCE_ALIASES)} builtin alias views refreshed "
          f"({', '.join(SOURCE_SCHEMAS)})")

    path = op.SOURCE_REGISTRY_PATH
    if not path.exists():
        print("SOURCES  (source_registry.xlsx not found, registry skipped)")
        return
    reg = pd.read_excel(path, sheet_name=0, dtype=str).fillna("")
    for _, row in reg.iterrows():
        if str(row.get("Enabled", "")).strip().lower() != "yes":
            continue
        source = clean_col(str(row["Source"]).strip()).lower()
        table = clean_col(str(row["Table"]).strip()).lower()
        folder = op.ONEDRIVE_ROOT / "data" / str(row["Folder"]).strip().strip("\\/")
        pattern = str(row.get("FilePattern", "")).strip() or "*.xlsx"
        files = [p for p in folder.glob(pattern) if not p.name.startswith("~$")]
        if not files:
            print(f"SOURCES  {source}.{table}: nothing matches '{pattern}' in {folder} — skipped")
            continue
        newest = max(files, key=lambda p: p.stat().st_mtime)
        sheet = str(row.get("Sheet", "")).strip()
        hdr_txt = str(row.get("HeaderRow", "")).strip()
        if hdr_txt.lower() == "auto":
            hdr = _probe_header_row(newest, sheet)
        else:
            hdr = (int(float(hdr_txt)) - 1) if hdr_txt else 0
        if newest.suffix.lower() == ".csv":
            df = pd.read_csv(newest, dtype=str, encoding="utf-8-sig",
                             low_memory=False, header=hdr)
        else:
            df = pd.read_excel(newest, sheet_name=sheet if sheet else 0,
                               dtype=str, engine="calamine", header=hdr)
        # Stage comes from the first folder segment (raw\... or processed\...),
        # and prefixes the view name so the stage is visible at a glance.
        stage = str(row["Folder"]).strip().replace("/", "\\").split("\\")[0].lower() or "raw"
        view = f"{stage}_{table}"
        raw_table = f"{source}_{table}"
        print(f"SOURCES  {newest.name}")
        write_table(engine, raw_table, prepare(df, newest.name))
        with engine.begin() as con:
            _ensure_schema(con, source)
            con.execute(text(
                f"CREATE OR ALTER VIEW [{source}].[{view}] AS "
                f"/* {source}.{view} — registry export over raw.{raw_table} "
                f"(stage: {stage}); auto-created by refresh.py from "
                f"source_registry.xlsx, do not edit */ "
                f"SELECT *, CONVERT(varchar(9), '{stage}') AS _stage "
                f"FROM raw.[{raw_table}]"))
        print(f"         -> raw.{raw_table} + view {source}.{view}")


def load_runlogs(engine):
    """Run-history workbooks -> raw.activity_log / raw.metrics_summary, so run
    status and metric trends are queryable (and visible in Power BI)."""
    for table, path in [
        ("activity_log", op.RUNLOGS_DIR / "activity_log.xlsx"),
        ("metrics_summary", op.RUNLOGS_DIR / "metrics_summary.xlsx"),
    ]:
        if not path.exists():
            print(f"RUNLOGS  ({path.name} not found, skipped)")
            continue
        print("RUNLOGS ", path.name)
        df = pd.read_excel(path, sheet_name=0, dtype=str, engine="calamine")
        write_table(engine, table, prepare(df, path.name))


def load_wave_change_requests(engine):
    """ALL 'Wave Change Request Form*.xlsx' workbooks in the folder, combined —
    only their 'Wave Change Request' data tab (the Instructions / quick-guide /
    ref tabs are for the requesters). One form per requester accumulates here;
    archive a form (move to archive\\<YYYY-MM>\\) once its requests are handled
    and it drops out on the next load. The strict filename pattern keeps stray
    workbooks dropped in the folder from being read as request forms."""
    files = sorted(p for p in op.RAW_WAVE_CHANGE_REQUESTS_DIR
                   .glob("Wave Change Request Form*.xlsx")
                   if not p.name.startswith("~$"))
    if not files:
        print("WAVE_CR  (no Wave Change Request Form*.xlsx found, skipped)")
        return
    frames = []
    for f in files:
        print("WAVE_CR ", f.name)
        df = pd.read_excel(f, sheet_name="Wave Change Request", dtype=str, engine="calamine")
        frames.append(prepare(df, f.name))
    write_table(engine, "wave_change_requests",
                pd.concat(frames, ignore_index=True, sort=False))


# ----------------------------------------------------------------------
# daily history snapshot (for trend / day-over-day comparison queries)
# ----------------------------------------------------------------------
# history table -> the query whose result is today's state worth keeping.
# roster_daily keeps the PERSON-LEVEL state (not aggregates) so history can
# be sliced by leader, vendor, wave, or any future category after the fact.
HISTORY_TABLES = {
    "roster_daily": "SELECT * FROM report.roster",
    # Daily record of Epic team members with no HR row — lastname13 the
    # FirstSeen/NewToday flags on the boss's Epic-not-in-HR gap report.
    "epic_not_in_hr_daily": "SELECT * FROM report.epic_not_in_hr",
    # Wave 3 per-leader daily totals (registered/trained, unregistered,
    # no-shows standing & to date) — lastname13 report.w3_registration_summary
    # and the W3 registration look-back CSVs. Added 2026-07-15.
    "w3_leader_daily": "SELECT * FROM report.w3_leader_snapshot",
}


def _ensure_history_schema(engine):
    with engine.begin() as con:
        con.execute(text("IF SCHEMA_ID('history') IS NULL EXEC('CREATE SCHEMA history')"))


def _sync_history_columns(engine, table: str, df: pd.DataFrame):
    """Add any columns df has that history.<table> lacks (as nullable NVARCHAR),
    so the roster can grow new columns without breaking the daily append."""
    with engine.begin() as con:
        existing = {r[0] for r in con.execute(text(
            f"SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('history.{table}')"))}
        if not existing:
            return  # table doesn't exist yet; to_sql will create it
        for col in df.columns:
            if col not in existing:
                con.execute(text(f"ALTER TABLE history.{table} ADD [{col}] NVARCHAR(4000) NULL"))
                print(f"HISTORY  history.{table}: added new column [{col}]")


def _backup_user_views(engine):
    """Script every hand-built view in the analyst's schemas (sandbox + the
    per-source schemas) to ONE .sql file on OneDrive (overwritten each run;
    OneDrive keeps versions). If the local DB is ever lost/rebuilt, running
    that file restores them all."""
    schemas = "', '".join(("sandbox",) + SOURCE_SCHEMAS)
    with engine.connect() as con:
        rows = con.execute(text(f"""
            SELECT SCHEMA_NAME(v.schema_id) AS sch, v.name, m.definition
            FROM sys.views v
            JOIN sys.sql_modules m ON m.object_id = v.object_id
            WHERE SCHEMA_NAME(v.schema_id) IN ('{schemas}')
            ORDER BY SCHEMA_NAME(v.schema_id), v.name""")).fetchall()
    if not rows:
        return
    defs = []
    for _, _, d in rows:
        # normalize to CREATE OR ALTER so the restore file is rerunnable
        defs.append(re.sub(r"(?i)CREATE\s+(OR\s+ALTER\s+)?VIEW",
                           "CREATE OR ALTER VIEW", d.strip(), count=1))
    path = op.REFERENCES_DIR / "user_views_backup.sql"
    path.parent.mkdir(parents=True, exist_ok=True)
    schema_ddl = "\n".join(
        f"IF SCHEMA_ID('{s}') IS NULL EXEC('CREATE SCHEMA [{s}]');"
        for s in ("sandbox",) + SOURCE_SCHEMAS)
    header = (
        "-- Hand-built view definitions (sandbox + source schemas) — auto-exported\n"
        "-- by sql\\refresh.py (snapshot step). Restore them all on a fresh DB with:\n"
        "--   sqlcmd -S \".\\SQLEXPRESS\" -E -C -d AnalyticsDB -i \"data\\references\\user_views_backup.sql\"\n"
        + schema_ddl + "\nGO\n\n"
    )
    path.write_text(header + "\nGO\n\n".join(defs) + "\nGO\n", encoding="utf-8")
    print(f"HISTORY  user views backed up: {len(rows)} -> {path.name}")


def snapshot(engine):
    """Append today's state to the history.* tables — one snapshot per calendar
    day; rerunning on the same day replaces that day's rows. Each day's rows are
    also mirrored to a dated CSV on OneDrive (history_snapshots/), which is the
    durable copy: the local DB can always be rebuilt from those via
    `refresh.py backload_history`. Also backs up hand-built sandbox.* views."""
    _ensure_history_schema(engine)
    today = f"{date.today():%Y-%m-%d}"
    for table, src in HISTORY_TABLES.items():
        df = pd.read_sql(src, engine)
        if df.empty:
            print(f"HISTORY  history.{table}: source empty, skipped")
            continue
        df.insert(0, "SnapshotDate", today)
        _sync_history_columns(engine, table, df)
        with engine.begin() as con:
            if con.execute(text(f"SELECT OBJECT_ID('history.{table}')")).scalar():
                con.execute(text(f"DELETE FROM history.{table} WHERE SnapshotDate = :d"),
                            {"d": today})
        df.to_sql(table, engine, schema="history", if_exists="append", index=False)
        csv_dir = op.month_subdir(op.HISTORY_SNAPSHOTS_DIR / table)
        csv_path = csv_dir / f"{table}_{date.today():%Y.%m.%d}.csv"
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        print(f"HISTORY  history.{table}: {len(df):,} rows stamped {today}  (CSV: {csv_path.name})")
    _backup_user_views(engine)


def backload_history(engine):
    """Rebuild the history.* tables from the dated CSVs on OneDrive — for a new
    machine or a lost/rebuilt local DB. Replaces the SQL tables entirely."""
    _ensure_history_schema(engine)
    for table in HISTORY_TABLES:
        files = sorted((op.HISTORY_SNAPSHOTS_DIR / table).glob(f"*/{table}_*.csv"))
        if not files:
            print(f"BACKLOAD history.{table}: no CSVs found, skipped")
            continue
        df = pd.concat([pd.read_csv(f, dtype=str, encoding="utf-8-sig") for f in files],
                       ignore_index=True, sort=False)
        df.to_sql(table, engine, schema="history", if_exists="replace", index=False)
        print(f"BACKLOAD history.{table}: {len(df):,} rows from {len(files)} daily CSVs")


LOADERS = {
    "master": load_master,
    "hr": load_hr,
    "hr_cleaned": load_hr_cleaned,
    "mvp": load_mvp,
    "mvp_roles": load_mvp_roles,
    "mvp_job_categories": load_mvp_job_categories,
    "cornerstone": load_cornerstone,
    "epic_status": load_epic_status,
    "epic_lookup": load_epic_lookup,
    "epic_class_schedule": load_epic_class_schedule,
    "wave_change_requests": load_wave_change_requests,
    "refs": load_refs,
    "runlogs": load_runlogs,
    "sources": load_sources,  # per-source schemas + the Excel source registry
    "snapshot": snapshot,   # keep last: a full run stamps history after all loads
}

# backload_history is deliberately NOT part of a full run — only when named.
COMMANDS = {**LOADERS, "backload_history": backload_history}


# ----------------------------------------------------------------------
# main
# ----------------------------------------------------------------------
def main():
    requested = [a.lower() for a in sys.argv[1:]] or list(LOADERS)
    unknown = [r for r in requested if r not in COMMANDS]
    if unknown:
        print("Unknown source(s):", unknown, "\nValid:", list(COMMANDS))
        sys.exit(1)

    engine = create_engine(CONN, fast_executemany=True)
    start = datetime.now()
    print(f"Loading {requested} into {DATABASE} at {start:%H:%M:%S}\n")
    failed = []
    for name in requested:
        try:
            COMMANDS[name](engine)
        except Exception as e:
            failed.append(name)
            print(f"   !! {name} FAILED: {e}")
    engine.dispose()
    print(f"\nDone in {(datetime.now() - start).total_seconds():.0f}s.")
    print("Next: run 3_report_views.sql once (or after editing views) to refresh report.*")
    if failed:
        print(f"FAILED sources: {', '.join(failed)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
