# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
update_master_training_status.py — mirror the tracker's per-person training
status onto the Master's DATA table, so the wave file carries the same four
fields the RCM Training Tracker's "Completion Status" sheet shows:

    Training Status                        <- TrainingStatus
    Anticipated Training Completion Date   <- FinalScheduledDate
    Epic Status Fully Registered           <- FullyRegisteredYN
    Epic Status Fully Trained              <- FullyTrainedYN

Source is report.tracker_training_status (built from raw.epic_status — Epic's
Curriculum Status Detail export — via report.tracker_detail), so the Master and
the tracker always agree. "Anticipated" = the LAST session date the person is
booked into; blank when nothing is scheduled (e.g. no Epic curriculum yet).
People outside the training-needed population (IsInScope = 0) stay blank in all
four columns — mirror-exactly, same rule as the EpicTMLookup_ block.

The first run also renames the two empty stub columns "Cornerstone Training
Status" / "Cornerstone Latest Registered Class Date" to the names above; the
data is Epic-sourced, not Cornerstone, and mislabeling the source would be
misleading. Renames are one-shot and idempotent.

Rules honored: xlwings/COM only (never openpyxl on the Master), Master must be
closed, columns located by header name and appended as the Table's last columns
so every other writer is unaffected. Idempotent — re-running recomputes and
overwrites the same four columns.

Ordering in the daily apply (2026-08-11): SQL is reloaded BEFORE this step
(`sql_stage`) and again after it (`sql_post`). The first reload means
report.users — and report.tracker_training_status on top of it — reflect TODAY'S
roster, so people add_members appended minutes earlier get their status in the
same run rather than a day later. The second reload puts these four columns into
raw.master and stamps the day's history from the finished workbook. Running this
script standalone skips both: reload first with `python sql\\refresh.py master`
if the Master changed since the last SQL load.

Usage:
    python scripts\\update_master_training_status.py            # preview only
    python scripts\\update_master_training_status.py --apply    # write
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

import activity_log                   # noqa: E402
import master_backup                  # noqa: E402
import onedrive_paths as op           # noqa: E402
from refresh import CONN              # noqa: E402
from sqlalchemy import create_engine   # noqa: E402

TABLE_NAME = "Table1"
DATA_SHEET = "DATA"

# Empty stub columns the user added earlier -> their corrected names.
RENAMES = {
    "Cornerstone Training Status":              "Training Status",
    "Cornerstone Latest Registered Class Date":  "Anticipated Training Completion Date",
}

# (Master header, source column, kind) in the order they should sit on DATA.
COLUMNS = [
    ("Training Status",                      "TrainingStatus",    "text"),
    ("Anticipated Training Completion Date", "AnticipatedDate",   "date"),
    ("Last Training Completed Date",         "CompletedDate",     "date"),
    ("Epic Status Fully Registered",         "FullyRegisteredYN", "text"),
    ("Epic Status Fully Trained",            "FullyTrainedYN",    "text"),
]

# Column whose fill-rate drives the shrink guard, and the two date columns.
GUARD_COL   = COLUMNS[0][0]
ANTICIPATED = COLUMNS[1][0]
COMPLETED   = COLUMNS[2][0]

DATE_FORMAT = "m/d/yyyy"
PREVIEW_CSV = op.LOCAL_MAIN_REPORTS_DIR / "training_status_preview.csv"

# OnLeaderList = 1 matters (QA 2026-08-11): report.tracker_training_status is
# IsInScope-only, but its underlying tracker_detail applies the canonical leader
# filter — so an off-list person with real Epic classes comes back as "No Epic
# Curriculum" with no date. Blank is the honest value for anyone outside our
# reporting population; without this filter 71 rows carry filter-driven labels.
# Two date columns, split 2026-08-19 after an audit of the single old one.
#
# "Anticipated Training Completion Date" used to be MAX(session date) over every
# class the person was registered for OR had completed, which conflated two
# different questions and went stale three ways:
#   * 112 of 121 Fully Trained W3 people carried their LAST ATTENDED date — an
#     actual completion, sitting in a column named "anticipated".
#   * 61 In Progress people carried a date already in the past (median 12 days,
#     max 47): every classroom session done, but an online class or test still
#     outstanding. Those items carry no session date, so the column could never
#     move forward to describe the work that was actually left.
#   * 15 cells held the literal text "Completed (Equivalent)" — mixed types in a
#     date column, which breaks Excel sort/filter and any downstream date math.
# It is now strictly forward-looking, and since 2026-08-19 it is sourced from
# CORNERSTONE too (per the analyst): the latest class the person is REGISTERED for.
# Cornerstone's "Registered" status excludes anything completed, so a fully
# trained person falls out to blank on its own — no separate rule needed. Epic's
# curriculum export is loaded for W3 only, so the old Epic-sourced version dated
# nobody in W1/W2; Cornerstone covers every wave (W1 67, W2 87, W3 564) and also
# catches registrations Epic lags on — USER04UID read 2026-08-10 from Epic while
# actually booked into 2026-10-01, a 52-day understatement.
#
# "Last Training Completed Date" is the actual completion, and comes from
# CORNERSTONE (raw.cornerstone — the Enterprise Training Report), not Epic.
# Cornerstone is the system of record for what was completed and when; Epic
# feeds from it. It also stamps a real date on equivalency credit, which Epic
# never does — that is what retires the "Completed (Equivalent)" label rather
# than just blanking it. Coverage is the other reason: Epic's curriculum export
# is loaded for W3 only, so it dates 489 W3 people and nobody in W1/W2, while
# Cornerstone dates 2,426 W2 and 1,107 W1 people as well.
#
# It is deliberately named "Last ..." — for a fully trained person it IS their
# completion date, but for someone mid-training it is their most recent finished
# item. That is true in every wave and needs no curriculum data to be correct;
# a plain "Training Completed Date" would be a claim we cannot support for W1/W2,
# where no required-curriculum list is loaded to say whether they are done.
QUERY = """
SELECT t.UniversalID, t.TrainingStatus,
       t.FullyRegisteredYN, t.FullyTrainedYN
FROM report.tracker_training_status t
WHERE t.OnLeaderList = 1
"""

# The completion date is rolled up in pandas, NOT in SQL. raw.cornerstone is
# ~793k rows of unindexed nvarchar(4000), so GROUP BY UPPER(LTRIM(RTRIM(...)))
# runs for minutes on SQLEXPRESS — measured >5 min and still going, against 5.5s
# to stream the same rows out and aggregate them here. This step runs inside the
# daily refresh, so the SQL version was a non-starter.
CORNERSTONE_QUERY = """
SELECT User_ID, Transcript_Status, Training_Start_Date, Transcript_Completed_Date
FROM raw.cornerstone
WHERE Training_Provider = 'EPIC'
  AND (Transcript_Status LIKE 'Completed%' OR Transcript_Status = 'Registered')
"""


def load_cornerstone_dates(eng) -> tuple[pd.Series, pd.Series]:
    """
    (last completed item, latest registered class) per UniversalID, from the
    Cornerstone Enterprise Training Report.

    'Registered' in Cornerstone means booked but NOT yet completed, so the
    second series is empty for anyone who has finished — which is exactly the
    estimated-completion semantics we want, with no extra filtering.
    """
    cs = pd.read_sql(CORNERSTONE_QUERY, eng)
    cs["UniversalID"] = cs["User_ID"].astype(str).str.strip().str.upper()
    cs = cs[(cs["UniversalID"] != "") & (cs["UniversalID"] != "NAN")]

    done = cs[cs["Transcript_Status"].str.startswith("Completed", na=False)].copy()
    done["d"] = pd.to_datetime(done["Transcript_Completed_Date"], errors="coerce")
    completed = done[done["d"].notna()].groupby("UniversalID")["d"].max()

    reg = cs[cs["Transcript_Status"] == "Registered"].copy()
    reg["d"] = pd.to_datetime(reg["Training_Start_Date"], errors="coerce")
    registered = reg[reg["d"].notna()].groupby("UniversalID")["d"].max()

    return completed, registered

# Empty-load guard: this script blanks any Master row the query doesn't return,
# so a stale/failed SQL side (views not applied, Epic status not loaded) would
# quietly wipe thousands of cells. Refuse to write if the population collapses.
MIN_ROWS = 1000                # hard floor — matches the MVP updater's convention
MAX_DROP_PCT = 0.20            # or a >20% fall vs what the Master already carries
MIN_COMPLETED = 2000           # Cornerstone completion dates (≈4,000 today)


def load_status() -> pd.DataFrame:
    eng = create_engine(CONN)
    try:
        df = pd.read_sql(QUERY, eng)
        completed, registered = load_cornerstone_dates(eng)
    finally:
        eng.dispose()
    if len(df) < MIN_ROWS:
        sys.exit(f"ABORT: report.tracker_training_status returned only {len(df):,} "
                 f"row(s) (floor {MIN_ROWS:,}). Refusing to blank the Master — "
                 f"check that the SQL refresh and report views ran.")
    df["UniversalID"] = df["UniversalID"].astype(str).str.strip().str.upper()
    df["CompletedDate"] = df["UniversalID"].map(completed)
    df["AnticipatedDate"] = df["UniversalID"].map(registered)

    # Same empty-load logic as above, for the Cornerstone side specifically:
    # raw.cornerstone is a separate feed, so it can go stale or fail to load
    # while report.tracker_training_status is perfectly healthy. Without this
    # the completion-date column would silently blank across every wave.
    n_completed = int(df["CompletedDate"].notna().sum())
    if n_completed < MIN_COMPLETED:
        sys.exit(f"ABORT: only {n_completed:,} Cornerstone completion date(s) "
                 f"(floor {MIN_COMPLETED:,}). Refusing to blank "
                 f"'{COMPLETED}' — check that raw.cornerstone loaded.")
    print(f"Cornerstone completion dates: {n_completed:,}")
    return df.drop_duplicates(subset="UniversalID").set_index("UniversalID")


def office_lock_present() -> bool:
    lock = op.MASTER_WAVE_PATH.parent / ("~$" + op.MASTER_WAVE_PATH.name)
    return lock.exists()


def cell_value(row, src: str, kind: str):
    """One Master cell: None (truly blank) when the person has no value."""
    if row is None:
        return None
    v = row[src]
    if v is None or (isinstance(v, float) and pd.isna(v)) or pd.isna(v):
        # A date column stays a date column: no value means blank, never a
        # text label. Equivalency credit now gets a real date from Cornerstone.
        return None
    if kind == "date":
        return v.normalize().to_pydatetime()   # session dates carry a start time
    s = str(v).strip()
    return s or None


def summarize(values: dict[str, list], uids: list[str]) -> pd.DataFrame:
    rows = []
    for header, _src, _kind in COLUMNS:
        col = values[header]
        rows.append({
            "Column": header,
            "Rows filled": sum(1 for v in col if v is not None),
            "Rows blank": sum(1 for v in col if v is None),
            "Total rows": len(uids),
        })
    return pd.DataFrame(rows)


def main(apply: bool) -> None:
    if office_lock_present():
        sys.exit("ERROR: Master Wave File is open in Excel. Close it and re-run.")

    status = load_status()
    print(f"report.tracker_training_status rows: {len(status):,}")

    import xlwings as xw

    app = wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = app.screen_updating = False
        wb = app.books.open(str(op.MASTER_WAVE_PATH), update_links=False)
        ws = wb.sheets[DATA_SHEET]
        lo = ws.tables[TABLE_NAME].api

        def headers() -> list[str]:
            return [str(h).strip() if h is not None else ""
                    for h in lo.HeaderRowRange.Value[0]]

        # -- one-shot rename of the empty Cornerstone stubs ------------------
        for old, new in RENAMES.items():
            hs = headers()
            if new in hs:
                continue
            if old in hs:
                if not apply:
                    print(f"  would rename '{old}' -> '{new}'")
                    continue
                lo.ListColumns(hs.index(old) + 1).Name = new
                print(f"  renamed '{old}' -> '{new}'")

        # -- ensure every target column exists -------------------------------
        col_pos: dict[str, int] = {}
        for header, _src, kind in COLUMNS:
            hs = headers()
            if header in hs:
                col_pos[header] = hs.index(header) + 1
            elif apply:
                lo.ListColumns.Add()
                n = lo.ListColumns.Count
                lo.ListColumns(n).Name = header
                col_pos[header] = n
                print(f"  added column '{header}' (table col {n})")
            else:
                print(f"  would add column '{header}'")

        first_row = lo.Range.Row + 1                   # first data row on sheet
        n_rows = lo.Range.Rows.Count - 1
        uid_col = lo.Range.Column + headers().index("UniversalID")
        raw_uids = ws.range((first_row, uid_col),
                            (first_row + n_rows - 1, uid_col)).value
        uids = [str(u).strip().upper() if u is not None else "" for u in raw_uids]

        by_uid = {u: status.loc[u] for u in set(uids) if u in status.index}
        values = {
            header: [cell_value(by_uid.get(u), src, kind) for u in uids]
            for header, src, kind in COLUMNS
        }

        summary = summarize(values, uids)
        print(f"\nMaster DATA rows: {n_rows:,}   matched to tracker: "
              f"{sum(1 for u in uids if u in by_uid):,}\n")
        print(summary.to_string(index=False))
        n_ant  = sum(1 for v in values[ANTICIPATED] if v is not None)
        n_comp = sum(1 for v in values[COMPLETED]   if v is not None)
        bad = [h for h in (ANTICIPATED, COMPLETED)
               for v in values[h] if v is not None and not hasattr(v, "year")]
        print(f"\n  {ANTICIPATED}: {n_ant:,} future dates"
              f"\n  {COMPLETED}: {n_comp:,} dates")
        if bad:
            sys.exit(f"ABORT: non-date value(s) built for {sorted(set(bad))} — "
                     f"date columns must never carry text.")

        preview = pd.DataFrame({"UniversalID": uids,
                                **{h: values[h] for h, _s, _k in COLUMNS}})
        preview = preview[preview[COLUMNS[0][0]].notna()]
        PREVIEW_CSV.parent.mkdir(parents=True, exist_ok=True)
        preview.to_csv(PREVIEW_CSV, index=False)
        print(f"\nPreview CSV ({len(preview):,} rows): {PREVIEW_CSV}")

        # Second half of the empty-load guard: compare against what the Master
        # already carries. A big fall means the source shrank, not that people
        # genuinely left the population — don't overwrite good data with blanks.
        status_header = COLUMNS[0][0]
        prev_col = ws.range(
            (first_row, lo.Range.Column + col_pos[status_header] - 1),
            (first_row + n_rows - 1, lo.Range.Column + col_pos[status_header] - 1)).value
        prev_filled = sum(1 for v in (prev_col or []) if v not in (None, ""))
        new_filled = sum(1 for v in values[status_header] if v is not None)
        if prev_filled and new_filled < prev_filled * (1 - MAX_DROP_PCT):
            sys.exit(f"ABORT: '{status_header}' would fall from {prev_filled:,} to "
                     f"{new_filled:,} filled rows (>{MAX_DROP_PCT:.0%} drop). "
                     f"Nothing written — investigate the source before re-running.")
        if prev_filled:
            print(f"  guard: {status_header} {prev_filled:,} -> {new_filled:,} filled")

        if not apply:
            print("\nDRY RUN — nothing written. Re-run with --apply to write.")
            return

        master_backup.ensure_daily_full_backup(op.MASTER_WAVE_PATH)

        for header, _src, kind in COLUMNS:
            sheet_col = lo.Range.Column + col_pos[header] - 1
            rng = ws.range((first_row, sheet_col), (first_row + n_rows - 1, sheet_col))
            rng.value = [[v] for v in values[header]]
            if kind == "date":
                rng.number_format = DATE_FORMAT
        wb.save()
        print(f"\nWrote {len(COLUMNS)} column(s) across {n_rows:,} rows; saved.")
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            if app is not None:
                app.quit()

    if apply:
        snap = master_backup.save_data_snapshot(op.MASTER_WAVE_PATH)
        print(f"DATA snapshot: {snap}")


if __name__ == "__main__":
    run_apply = "--apply" in sys.argv[1:]
    with activity_log.track_run("update_master_training_status.py"):
        main(run_apply)
