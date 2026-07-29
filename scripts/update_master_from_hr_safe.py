# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
update_master_from_hr_safe.py

Updates leadership columns in the Master Wave File DATA sheet with the latest
values from the HR export, matched on UniversalID (case-insensitive).

WHY THIS VERSION EXISTS — workbook-safety
-----------------------------------------
The previous versions used openpyxl to OPEN and SAVE the Master Wave File. When
openpyxl saves an .xlsx it rebuilds the entire file package from its own in-memory
model. It silently drops or mangles anything it does not fully model — Power Query
connections, external data connections, slicers, dashboard drawing objects, some
chart XML, pivot caches, and assorted workbook metadata. For a workbook that is the
team's source of truth, that is an unacceptable risk.

This version NEVER uses openpyxl (or pandas/xlsxwriter) to write the Master. Instead
it drives the real Excel desktop application through COM via xlwings:

    * Excel itself opens the workbook, so every internal object it understands
      (queries, connections, slicers, charts, tables, formatting, formulas, named
      ranges, metadata) is preserved exactly as Excel maintains it.
    * We write ONLY cell VALUES into the specific target columns of the DATA sheet
      (and the single dashboard date cell). We never touch structure, formatting,
      tables, sheets, calculation mode (permanently), or connections.
    * Excel performs the save in its native format, so the package is never rebuilt
      by a third-party library.

openpyxl/pandas ARE still used to READ the CLEANED HR file's "All Staff" tab
(HR_cleaned.xlsx, produced by clean_hr.py), in which leader-chain Universal IDs
have already been resolved to full names and leader spellings standardized. That
file is a disposable derived export, so rebuilding it is irrelevant (we never
save it).

Business logic (unchanged from update_master_from_hrv2.py)
----------------------------------------------------------
  - Read the cleaned HR file's "All Staff" tab, which is unfiltered (all rows) —
    do NOT filter by IsRCM
  - Deduplicate HR on UniversalID (keep first occurrence)
  - Match on UniversalID (normalized to upper) — never by name
  - If UID found in HR with a real senior chain: overwrite SeniorManager,
    Director, Sr. Director, AVP, VP, SVP; recompute Leaders; clear stale notes
  - If UID found in HR but VP/AVP/SVP are ALL blank: PRESERVE existing leadership
    cells, flag "blank leadership" note for manual review, skip change-log
  - If UID NOT in HR: leave leadership unchanged; add "User Not in HR Export." note
  - Leaders = VP > AVP > SVP fallback; non-canonical leader -> "No Longer Rev Cycle"
  - LEADER_UID_OVERRIDES win over computed Leaders
  - Change log ('Users HR Change Log' column, located by header name)
    accumulates "MM/DD/YY - Old to New per MM/DD/YY HR file"
  - Dashboard date cell written via Excel (DASHBOARD!C3)

How to run — this script ALWAYS applies. There is no mode flag to remember:

    python scripts/update_master_from_hr_safe.py
        Backs up the Master, writes only the changed cells through Excel,
        recomputes Leaders, migrates legacy "Not Rev Cycle" -> "No Longer Rev
        Cycle", updates DASHBOARD!C3, and logs.

To PREVIEW what would change WITHOUT writing anything, run the dedicated,
read-only preview script (it never opens the Master for writing, never backs
up, never saves):

    python scripts/review_hr_leader_changes.py

That split is intentional: previewing and applying are two different scripts, so
there is no "did I add/remove --dry-run?" footgun. This script still refuses to
run if the Master is open/locked in Excel, and requires the cleaned HR file
(run clean_hr.py first).
"""

import argparse
import datetime
import os
import sys
from pathlib import Path

import openpyxl  # HR export READ ONLY — never used to write the Master
import xlwings as xw  # drives the real Excel app so the Master package is preserved

sys.path.insert(0, str(Path(__file__).parent))
import runlog  # writes to a SEPARATE master_runlog.xlsx — never the Master Wave File
import onedrive_paths
import master_backup
import activity_log
from leader_names import (
    CANONICAL_SET,
    LEGACY_NOT_REV_CYCLE,
    NOT_REV_CYCLE,
    NOTE_HR_BLANK_LEADERSHIP,
    NOTE_NOT_IN_HR,
    add_note,
    append_change_log,
    classify_leader_change,
    format_leader_change,
    normalize_leader,
    parse_hr_export_date,
    remove_note,
)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
# Derive the project root from this script's own location (scripts/ -> parent),
# so the script works no matter which directory it's launched from. An optional
# YOURORG_ANALYTICS_BASE env var still overrides it if ever needed.
DEFAULT_BASE = Path(__file__).resolve().parent.parent
BASE = Path(os.getenv("YOURORG_ANALYTICS_BASE", DEFAULT_BASE))

HR_PATH = onedrive_paths.latest_hr_cleaned()  # most recent dated HR_cleaned_*.xlsx, or None
HR_SHEET = "All Staff"  # the unfiltered tab of HR_cleaned.xlsx (clean_hr.py output)
MASTER_PATH = onedrive_paths.MASTER_WAVE_PATH
LOG_PATH = BASE / "scripts/update_master_from_hr_log.txt"

RUNLOG_TAB = "HR Update"
RUNLOG_HEADERS = [
    "Run Timestamp", "HR Export Date", "HR Rows Loaded", "Master Rows Updated",
    "Rows NOT in HR", "Rows: HR Blank Leadership (Preserved)",
    "Rows Skipped (Blank UID)", "Total Data Rows",
    "Leaders: VP Used", "Leaders: AVP Used", "Leaders: SVP Fallback",
    "Leaders Changed (Total)", "Leaders Changed: Reassigned within RCM",
    "Leaders Changed: Fell out (Not Rev Cycle)", "Leaders Changed: Re-entered RCM",
    "Leaders Changed: Newly assigned", "Leaders Changed: Cleared",
    "UIDs Not in HR", "HR File", "Master File", "Backup File",
    "Data Snapshot File",
]

# HR column name -> Master DATA column name. Leaders (computed) is handled separately.
LEAD_MAP = {
    "SeniorManager": "SeniorManager",
    "Director": "Director",
    "SeniorDirector": "Sr. Director",  # header name differs between files
    "AVP": "AVP",
    "VP": "VP",
    "SVP": "SVP",
}

# Only VP/AVP/SVP are canonicalized against the leaders list. SeniorManager/
# Director/Sr. Director can hold any employee name and must NOT be run through
# last-name lookup (it would match the wrong people).
_CANONICALIZE = {"VP", "AVP", "SVP"}

# UID-keyed Leaders overrides. For Rev Cycle people whose HR senior chain routes
# them through a non-canonical VP/AVP (so compute_leaders would otherwise drop them
# to "No Longer Rev Cycle") even though they belong under a canonical leader. Applied
# AFTER compute_leaders so it always wins regardless of HR data.
LEADER_UID_OVERRIDES = {
    "BKREBS":     "Lastname03, Firstname03",   # Lastname09, Firstname09
    "LEADER02UID": "Lastname03, Firstname03",   # Lastname02, Firstname02
}

# The exact set of DATA columns this script is permitted to write. Nothing else
# in the workbook is ever touched.
TARGET_COLS = [
    "SeniorManager", "Director", "Sr. Director", "AVP", "VP", "SVP",
    "Leaders", "Notes", "Users HR Change Log",
]

# Dashboard date cell — mirrors leader_names.update_dashboard_date, but written
# through Excel/COM so openpyxl never rebuilds the Master package.
DASHBOARD_SHEET = "DASHBOARD"
DASHBOARD_DATE_CELL = "C3"


# ===========================================================================
# Helpers — pure logic (shared, no I/O side effects)
# ===========================================================================
def clean_uid(raw_val) -> str | None:
    """Convert an Excel cell value to a standardized upper-case string UID."""
    if not raw_val:
        return None
    val_str = str(raw_val).strip()
    # Excel/COM sometimes reads numeric IDs as floats (e.g. 12345.0)
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str.upper() or None


def clean_leadership(raw, canonicalize: bool = False) -> str | None:
    """Clean a raw HR leadership value; scrub null variations and 'nan, nan' text."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan, nan", "nan"):
        return None
    return normalize_leader(s) if canonicalize else s


def compute_leaders(vp: str | None, avp: str | None, svp: str | None) -> str | None:
    """VP > AVP > SVP fallback; non-canonical chosen leader becomes 'No Longer Rev Cycle'."""
    raw = vp or avp or svp
    if raw is None:
        return None
    return raw if raw in CANONICAL_SET else NOT_REV_CYCLE


def find_duplicate_headers(headers) -> list:
    """Return a list of header names that appear more than once (blanks ignored)."""
    seen, dups = set(), []
    for h in headers:
        if h is None:
            continue
        key = str(h).strip()
        if not key:
            continue
        if key in seen and key not in dups:
            dups.append(key)
        seen.add(key)
    return dups


# ===========================================================================
# HR export — READ ONLY via openpyxl (safe: we never save this file)
# ===========================================================================
def load_hr(path: Path) -> tuple[dict, str]:
    """Load HR file, dedupe on UniversalID (keep first). Returns (dict, sheet_name)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet_name = HR_SHEET if HR_SHEET in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    rows = ws.iter_rows(values_only=True)
    try:
        headers = [str(h).strip() if h is not None else "" for h in next(rows)]
    except StopIteration:
        wb.close()
        raise ValueError(f"The HR export at {path} contains no rows.")

    # Safety: duplicate HR headers would make column lookup ambiguous.
    hr_dups = find_duplicate_headers(headers)
    if hr_dups:
        wb.close()
        sys.exit(f"ERROR: HR export has duplicate header(s): {hr_dups}. Aborting.")

    col = {h: i for i, h in enumerate(headers) if h}

    # Safety: confirm every HR column we read actually exists before processing.
    hr_required = ["UniversalID"] + list(LEAD_MAP.keys())
    hr_missing = [c for c in hr_required if c not in col]
    if hr_missing:
        wb.close()
        sys.exit(f"ERROR: HR export is missing required column(s): {hr_missing}. Aborting.")

    uid_i = col["UniversalID"]

    hr = {}
    for row in rows:
        if uid_i >= len(row):
            continue
        uid = clean_uid(row[uid_i])
        if not uid:
            continue
        if uid not in hr:  # keep first occurrence
            hr[uid] = {headers[i]: (row[i] if i < len(row) else None)
                       for i in range(len(headers))}

    wb.close()
    return hr, sheet_name


# ===========================================================================
# Excel/COM lock + open-state checks
# ===========================================================================
def office_lock_present(path: Path) -> bool:
    """True if the Office owner-lock file (~$name) exists next to the workbook."""
    return path.with_name("~$" + path.name).exists()


def already_open_in_excel(path: Path) -> bool:
    """True if the workbook is already open in any running Excel instance."""
    target = path.resolve()
    try:
        for app in xw.apps:
            for bk in app.books:
                try:
                    if Path(bk.fullname).resolve() == target:
                        return True
                except Exception:
                    continue
    except Exception:
        # No Excel running, or COM not reachable — treat as "not open".
        pass
    return False


# ===========================================================================
# Excel column read/write helpers (xlwings) — values only, single column
# ===========================================================================
def read_col(ws, col_idx: int, first_row: int, last_row: int) -> list:
    """Read one column (first_row..last_row) into a flat Python list, one COM call."""
    if last_row < first_row:
        return []
    if last_row == first_row:
        return [ws.range((first_row, col_idx)).value]
    return list(ws.range((first_row, col_idx), (last_row, col_idx)).value)


def write_col(ws, col_idx: int, first_row: int, values: list) -> None:
    """Write a flat Python list down one column starting at first_row. Values only."""
    if not values:
        return
    if len(values) == 1:
        ws.range((first_row, col_idx)).value = values[0]
    else:
        # transpose=True writes a 1-D list vertically; only .value is set, so
        # number formats, fonts, borders, and everything else are untouched.
        ws.range((first_row, col_idx)).options(transpose=True).value = values


def last_header_col(ws) -> int:
    """
    Last non-empty header cell in row 1, via End-left from the sheet's right edge.
    More reliable than used_range, which can under-report its extent.
    """
    last = ws.range((1, ws.cells.last_cell.column)).end("left").column
    return max(last, 1)


def last_data_row(ws, key_col: int) -> int:
    """
    Last data row, via End-up from the bottom of the key (UID) column.

    Excel's used_range.last_cell is unreliable — in testing it under-reported by
    one and silently dropped the final data row (which would leave that row never
    updated). End-up on the key column is the canonical, stable way to find the
    last populated row. Rows past the last UID have no key to match/update anyway.
    """
    bottom = ws.range((ws.cells.last_cell.row, key_col)).end("up").row
    return max(bottom, 1)


def diff_segments(orig_vals: list, new_vals: list) -> list:
    """
    Find the cells that actually changed and group them into contiguous runs.

    Returns a list of (start_index, sublist) where start_index is the 0-based
    offset into the column and sublist is the run of new values to write. This
    lets us write ONLY changed cells while still using one bulk range write per
    run (instead of thousands of single-cell COM calls).
    """
    segments = []
    i, n = 0, len(new_vals)
    while i < n:
        if new_vals[i] != orig_vals[i]:
            j = i
            while j < n and new_vals[j] != orig_vals[j]:
                j += 1
            segments.append((i, new_vals[i:j]))
            i = j
        else:
            i += 1
    return segments


# ===========================================================================
# Run log workbook (separate file — never the Master)
# ===========================================================================
def append_runlog(backup_label, hr_export_date, hr_rcm_count, updated, not_in_hr,
                  hr_blank, skipped_blank, total_rows, leaders_vp, leaders_avp,
                  leaders_svp, changes, not_in_hr_uids, data_snapshot_path=None):
    """Append execution metrics to the persistent historical run log tracker."""
    uid_list = ", ".join(sorted(not_in_hr_uids)) if not_in_hr_uids else ""
    runlog.append(RUNLOG_TAB, RUNLOG_HEADERS, [
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        hr_export_date, hr_rcm_count, updated, not_in_hr, hr_blank,
        skipped_blank, total_rows,
        leaders_vp, leaders_avp, leaders_svp,
        changes["total"], changes["reassigned"], changes["fell_out"],
        changes["reentered"], changes["newly_assigned"], changes["cleared"],
        uid_list, str(HR_PATH), str(MASTER_PATH), str(backup_label),
        str(data_snapshot_path),
    ])


# ===========================================================================
# Command-line mode — this script ALWAYS applies; there is no mode flag.
# To preview without writing, run review_hr_leader_changes.py instead.
# --apply / --yes are accepted as harmless no-ops for backward compatibility
# (e.g. hr_apply.bat). --dry-run is intentionally rejected with a redirect so
# that stale muscle memory can never cause a silent, unexpected write.
# ===========================================================================
def parse_args():
    ap = argparse.ArgumentParser(
        description="Apply Master Wave File leadership updates from the cleaned HR file. "
                    "This script always applies; to preview without writing, run "
                    "review_hr_leader_changes.py instead.",
        epilog="No mode flag needed — running this always applies.",
    )
    ap.add_argument("--apply", action="store_true",
                    help="(No-op) kept for compatibility; this script always applies.")
    ap.add_argument("--yes", action="store_true",
                    help="(No-op) kept for compatibility; live runs no longer prompt.")
    ap.add_argument("--dry-run", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()
    if args.dry_run:
        sys.exit(
            "This script no longer has a dry-run mode (apply is the only behavior).\n"
            "For a read-only preview that writes NOTHING, run:\n"
            "    python scripts/review_hr_leader_changes.py"
        )
    return args


# ===========================================================================
# Main
# ===========================================================================
def main(dry_run: bool):
    mode_banner = "DRY RUN — NOTHING WILL BE WRITTEN" if dry_run else "LIVE RUN — WILL WRITE"
    print("=" * 70)
    print(f"  update_master_from_hr_safe.py   [{mode_banner}]")
    print("=" * 70 + "\n")

    # -- 1. Verify source assets exist ----------------------------------------
    if HR_PATH is None:
        sys.exit(f"ERROR: No cleaned HR file found in {onedrive_paths.CLEAN_HR_DIR}\n       Run scripts/clean_hr.py first to generate it.")
    if not MASTER_PATH.exists():
        sys.exit(f"ERROR: Master Wave File not found at: {MASTER_PATH}")

    # -- 2. Refuse to run if the Master is already open / locked --------------
    # Editing a workbook that the user has open in Excel risks fighting over the
    # file and losing their unsaved edits. Stop cleanly in BOTH modes.
    if office_lock_present(MASTER_PATH) or already_open_in_excel(MASTER_PATH):
        sys.exit(
            "ERROR: The Master Wave File appears to be open in Excel (or locked).\n"
            "       Close it completely in every Excel window, then run again."
        )

    # -- 3. Load HR (read-only; validates HR headers + required columns) ------
    print("Loading HR export (read-only)...")
    hr, sheet_name = load_hr(HR_PATH)
    hr_export_date = sheet_name
    hr_date_short = parse_hr_export_date(sheet_name)
    run_date_short = datetime.date.today().strftime("%m/%d/%y")
    print(f"  Sheet: {sheet_name}")
    print(f"  HR file date: {hr_date_short}  |  Run date: {run_date_short}")
    print(f"  {len(hr):,} unique UIDs loaded\n")

    # -- 4. Backup (once per calendar day, real runs only) --------------------
    # Full-workbook copy, skipped if today's already exists (see master_backup.py).
    # Does not touch workbook internals.
    backup_label = "(dry-run: no backup)"
    data_snapshot_path = None
    if not dry_run:
        try:
            backup_path = master_backup.ensure_daily_full_backup(MASTER_PATH)
        except PermissionError:
            sys.exit(f"ERROR: Could not create backup (permission denied): {MASTER_PATH}")
        backup_label = str(backup_path)
        print()

    # -- 5. Open the Master through Excel itself ------------------------------
    # add_book=False so no blank workbook is created. We track whether WE started
    # Excel so we only quit the instance we own.
    app = None
    created_app = False
    wb = None
    orig_calc = None
    orig_alerts = None
    orig_screen = None
    try:
        app = xw.App(visible=False, add_book=False)
        created_app = True
        # App-instance conveniences; all restored in finally before we quit.
        orig_alerts = app.display_alerts
        orig_screen = app.screen_updating
        app.display_alerts = False
        app.screen_updating = False

        print("Opening Master Wave File through Excel...")
        # update_links=False: do not refresh external links on open.
        # read_only in dry run so we physically cannot modify the file.
        wb = app.books.open(
            str(MASTER_PATH),
            update_links=False,
            read_only=dry_run,
        )

        # Capture the workbook's calculation mode so we can put it back exactly.
        # We set manual only for the duration of writing, then RESTORE it — the
        # saved workbook keeps whatever mode it already had (never changed
        # permanently).
        orig_calc = app.calculation
        app.calculation = "manual"

        if DASHBOARD_SHEET not in [s.name for s in wb.sheets]:
            pass  # dashboard optional; date write is skipped later if absent
        if "DATA" not in [s.name for s in wb.sheets]:
            sys.exit("ERROR: A sheet named 'DATA' was not found in the Master Wave File.")
        ws = wb.sheets["DATA"]

        # -- 6. Read + validate the DATA header row ---------------------------
        # End-based extent detection — used_range can under-report and silently
        # drop the final data row(s) (observed in testing); that would leave the
        # last rows never updated.
        last_col = last_header_col(ws)
        header_row = ws.range((1, 1), (1, last_col)).value
        if not isinstance(header_row, list):
            header_row = [header_row]

        data_dups = find_duplicate_headers(header_row)
        if data_dups:
            sys.exit(f"ERROR: Master DATA sheet has duplicate header(s): {data_dups}. Aborting.")

        col_idx = {
            str(h).strip(): i + 1
            for i, h in enumerate(header_row)
            if h is not None and str(h).strip() != ""
        }

        required = TARGET_COLS + ["UniversalID"]
        missing = [c for c in required if c not in col_idx]
        if missing:
            sys.exit(f"ERROR: Master DATA sheet missing required column(s): {missing}. Aborting.")

        uid_col = col_idx["UniversalID"]
        last_row = last_data_row(ws, uid_col)
        total_rows = max(0, last_row - 1)
        print(f"  DATA rows to process: {total_rows:,}\n")

        # -- 7. Bulk-read only the columns we need ----------------------------
        first = 2
        uid_vals = read_col(ws, uid_col, first, last_row)
        orig = {c: read_col(ws, col_idx[c], first, last_row) for c in TARGET_COLS}
        # Working copies we mutate; written back only where they differ from orig.
        new = {c: list(orig[c]) for c in TARGET_COLS}
        n = len(uid_vals)

        # -- 8. Process rows in memory (identical business logic) -------------
        updated = not_in_hr = hr_blank_leadership = skipped_blank = 0
        leaders_vp = leaders_avp = leaders_svp = 0
        not_in_hr_uids = []
        change_preview = []  # (uid, old, new) for dry-run display
        changes = {"total": 0, "reassigned": 0, "fell_out": 0,
                   "reentered": 0, "newly_assigned": 0, "cleared": 0}

        for r in range(n):
            uid = clean_uid(uid_vals[r])
            if not uid:
                skipped_blank += 1
                continue

            hr_row = hr.get(uid)

            if not hr_row:
                # Not in HR: leave leadership untouched; add the note.
                new["Notes"][r] = add_note(new["Notes"][r], NOTE_NOT_IN_HR)
                not_in_hr += 1
                not_in_hr_uids.append(uid)
                continue

            # Clean each mapped leadership field.
            clean_fields = {
                LEAD_MAP[f]: clean_leadership(hr_row.get(f), canonicalize=(f in _CANONICALIZE))
                for f in LEAD_MAP
            }
            vp = clean_fields["VP"]
            avp = clean_fields["AVP"]
            svp = clean_fields["SVP"]

            # Blank-senior-chain guard: HR has the user but no VP/AVP/SVP.
            # Preserve all master leadership cells, flag for manual review,
            # skip the change-log. (Biweekly HR drops make a missing chain more
            # likely a stale/partial export than a real reporting-structure change.)
            if vp is None and avp is None and svp is None:
                na = remove_note(new["Notes"][r], NOTE_NOT_IN_HR)
                na = add_note(na, NOTE_HR_BLANK_LEADERSHIP)
                new["Notes"][r] = na
                hr_blank_leadership += 1
                continue

            # Compute Leaders, then apply hardcoded UID override (wins outright).
            new_leaders = compute_leaders(vp, avp, svp)
            if uid in LEADER_UID_OVERRIDES:
                new_leaders = LEADER_UID_OVERRIDES[uid]

            old_leaders = orig["Leaders"][r]

            # Write each leadership column value into the working arrays.
            for master_col, val in clean_fields.items():
                new[master_col][r] = val
            new["Leaders"][r] = new_leaders

            # Tally which tier drove Leaders.
            if vp:
                leaders_vp += 1
            elif avp:
                leaders_avp += 1
            else:
                leaders_svp += 1

            # Diff old vs new Leaders and append a change-log entry.
            bucket = classify_leader_change(old_leaders, new_leaders)
            if bucket:
                changes["total"] += 1
                changes[bucket] += 1
                entry = format_leader_change(run_date_short, hr_date_short,
                                             old_leaders, new_leaders)
                if entry:
                    new["Users HR Change Log"][r] = append_change_log(
                        orig["Users HR Change Log"][r], entry)
                    if len(change_preview) < 30:
                        change_preview.append((uid, old_leaders, new_leaders))

            # Clear stale "not in HR" / "blank leadership" notes now resolved.
            na = remove_note(new["Notes"][r], NOTE_NOT_IN_HR)
            na = remove_note(na, NOTE_HR_BLANK_LEADERSHIP)
            new["Notes"][r] = na
            updated += 1

        # -- 8b. One-time wording migration -----------------------------------
        # Flip any remaining legacy "Not Rev Cycle" Leaders cell to the current
        # "No Longer Rev Cycle" wording. Covers rows that were NOT recomputed this
        # run (users not in HR / with a blank HR chain, whose Leaders is preserved).
        # Comparison treats the two spellings as equal, so this is not logged as a
        # leader change — it only updates the displayed text. Self-cleaning: a no-op
        # once every cell carries the new wording.
        migrated_wording = 0
        for r in range(n):
            val = new["Leaders"][r]
            if val is not None and str(val).strip() == LEGACY_NOT_REV_CYCLE:
                new["Leaders"][r] = NOT_REV_CYCLE
                migrated_wording += 1

        # -- 9. Commit (or report, in dry run) --------------------------------
        # Compute per-column changed segments once; reuse for both modes. We write
        # ONLY cells whose value actually changed (grouped into contiguous runs).
        col_segments = {c: diff_segments(orig[c], new[c]) for c in TARGET_COLS}
        changed_cols = [c for c in TARGET_COLS if col_segments[c]]
        changed_cells = sum(len(sub) for c in TARGET_COLS for _, sub in col_segments[c])

        if dry_run:
            print("=== DRY RUN — NOTHING WAS WRITTEN ===")
            print(f"  Rows that WOULD update from HR:        {updated:,}")
            print(f"  Rows NOT in HR (would be noted):       {not_in_hr:,}")
            print(f"  Rows w/ HR blank leadership (preserve): {hr_blank_leadership:,}")
            print(f"  Rows skipped (blank UID):              {skipped_blank:,}")
            print(f"  Leader changes that WOULD be logged:   {changes['total']:,}")
            print(f"  Legacy 'Not Rev Cycle' -> '{NOT_REV_CYCLE}': {migrated_wording:,}  (wording only)")
            print(f"  Cells that WOULD be written:           {changed_cells:,}")
            print(f"  Columns that WOULD be written:         "
                  f"{', '.join(changed_cols) if changed_cols else '(none)'}")
            if change_preview:
                print("\n  Sample of leader changes (first 30):")
                for uid, old_l, new_l in change_preview:
                    print(f"    {uid:<14} {old_l!r}  ->  {new_l!r}")
            print("\n  Dashboard date cell would be set to today.")
            print("  Re-run with --apply to write these changes.")
        else:
            # Write back ONLY the cells that changed, grouped into contiguous runs
            # per column. Each run is one bulk value write — no formatting/structure
            # touched, and untouched rows are never written.
            if migrated_wording:
                print(f"Migrating {migrated_wording:,} legacy 'Not Rev Cycle' "
                      f"cell(s) to '{NOT_REV_CYCLE}' (wording only, not logged).")
            print(f"Writing {changed_cells:,} changed cell(s) across "
                  f"{len(changed_cols)} column(s): "
                  f"{', '.join(changed_cols) if changed_cols else '(none)'}")
            for c in changed_cols:
                for start, sub in col_segments[c]:
                    write_col(ws, col_idx[c], first + start, sub)

            # Dashboard run date — single cell only (mirrors update_dashboard_date).
            if DASHBOARD_SHEET in [s.name for s in wb.sheets]:
                wb.sheets[DASHBOARD_SHEET].range(DASHBOARD_DATE_CELL).value = (
                    datetime.date.today().strftime("%m/%d/%Y")
                )

            # Restore original calculation mode BEFORE saving so the workbook
            # keeps whatever mode it already had — never changed permanently.
            app.calculation = orig_calc
            orig_calc = None  # already restored; don't double-restore in finally

            wb.save()  # native Excel save — package preserved
            print("Saved through Excel (workbook package preserved).")

    finally:
        # Always restore app state and clean up, even on error.
        try:
            if app is not None and orig_calc is not None:
                app.calculation = orig_calc
        except Exception:
            pass
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        try:
            if app is not None:
                app.display_alerts = orig_alerts if orig_alerts is not None else True
                app.screen_updating = orig_screen if orig_screen is not None else True
        except Exception:
            pass
        # Only quit Excel if WE created the instance.
        try:
            if created_app and app is not None:
                app.quit()
        except Exception:
            pass

    # -- 10. Console summary (always) -----------------------------------------
    print("\n=== RESULTS ===")
    print(f"  HR export date:               {hr_export_date}  ({hr_date_short})")
    print(f"  Rows updated from HR:         {updated:,}")
    print(f"  Rows NOT in HR (noted):       {not_in_hr:,}")
    print(f"  Rows w/ HR blank leadership:  {hr_blank_leadership:,}  (preserved + flagged)")
    print(f"  Rows skipped (blank UID):     {skipped_blank:,}")
    print("\n  Leaders breakdown:")
    print(f"    VP used:          {leaders_vp:,}")
    print(f"    AVP stepped in:   {leaders_avp:,}")
    print(f"    SVP fallback:     {leaders_svp:,}  (both VP+AVP missing)")
    print("\n  Leader changes (col AO 'Users HR Change Log'):")
    print(f"    Total changes:           {changes['total']:,}")
    print(f"    Reassigned within RCM:   {changes['reassigned']:,}")
    print(f"    Fell out (No Longer Rev Cycle):{changes['fell_out']:,}")
    print(f"    Re-entered RCM:          {changes['reentered']:,}")
    print(f"    Newly assigned:          {changes['newly_assigned']:,}")
    print(f"    Cleared:                 {changes['cleared']:,}")

    # -- 11. Side-effect outputs — real runs only -----------------------------
    if dry_run:
        print("\n(DRY RUN: text log and run-log workbook were NOT written.)")
        return

    # Versioned DATA-only snapshot (lightweight, one per successful run). Read
    # fresh from disk now that Excel has saved and the workbook/app are closed.
    # Best-effort: never masks the successful save above (returns None on failure).
    data_snapshot_path = master_backup.save_data_snapshot(MASTER_PATH)
    if data_snapshot_path:
        print(f"\nData snapshot: {data_snapshot_path}")

    with open(LOG_PATH, "w", encoding="utf-8") as log:
        log.write("update_master_from_hr_safe.py — run log\n")
        log.write("=" * 60 + "\n\n")
        log.write(f"HR path:     {HR_PATH}\n")
        log.write(f"HR sheet:    {sheet_name}\n")
        log.write(f"Master path: {MASTER_PATH}\n")
        log.write(f"Backup:      {backup_label}\n")
        log.write(f"Data snapshot: {data_snapshot_path}\n\n")
        log.write(f"HR UIDs loaded:               {len(hr):,}\n")
        log.write(f"Master rows updated:          {updated:,}\n")
        log.write(f"Master rows NOT in HR:        {not_in_hr:,}\n")
        log.write(f"Rows w/ HR blank leadership:  {hr_blank_leadership:,}  (preserved + flagged)\n")
        log.write(f"Rows skipped (blank UID):     {skipped_blank:,}\n\n")
        log.write(f"Leaders — VP used:            {leaders_vp:,}\n")
        log.write(f"Leaders — AVP stepped in:     {leaders_avp:,}\n")
        log.write(f"Leaders — SVP fallback:       {leaders_svp:,}\n\n")
        log.write("Leader changes (col AO):\n")
        log.write(f"  Total:                  {changes['total']:,}\n")
        log.write(f"  Reassigned within RCM:  {changes['reassigned']:,}\n")
        log.write(f"  Fell out:               {changes['fell_out']:,}\n")
        log.write(f"  Re-entered RCM:         {changes['reentered']:,}\n")
        log.write(f"  Newly assigned:         {changes['newly_assigned']:,}\n")
        log.write(f"  Cleared:                {changes['cleared']:,}\n\n")
        if not_in_hr_uids:
            log.write(f"UIDs not found in HR ({len(not_in_hr_uids)}):\n")
            for missing_uid in sorted(not_in_hr_uids):
                log.write(f"  {missing_uid}\n")

    append_runlog(
        backup_label, hr_export_date, len(hr), updated, not_in_hr,
        hr_blank_leadership, skipped_blank, total_rows,
        leaders_vp, leaders_avp, leaders_svp, changes, not_in_hr_uids,
        data_snapshot_path,
    )

    print(f"\n  Log written: {LOG_PATH.name}")
    print(f"  Run log appended: master_runlog.xlsx (tab: {RUNLOG_TAB})")
    print(f"  Saved: {MASTER_PATH.name}")


if __name__ == "__main__":
    parse_args()  # accepts/validates legacy flags; always applies
    with activity_log.track_run("update_master_from_hr_safe.py"):
        main(dry_run=False)
