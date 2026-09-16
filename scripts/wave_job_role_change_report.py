# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
wave_job_role_change_report.py

Auto-detects all dated *_UserMappingsData.csv snapshots in
REPORTS/Job Role Change Reports/User Mappings Monthly, sorted by date, and
produces a wave & job role change report for Dr. Lastname03.

Monthly routine: drop the new first-week-of-month User Mappings export into
that folder (filename must start with YYYY-MM-DD) and rerun — no code changes.
Each file is a point-in-time snapshot of MVP as of its filename date.

Snapshots vs periods — these are named on different clocks, deliberately:
  * A SNAPSHOT is named for its own month ("August" = the roster as of Aug 2).
  * A PERIOD (the delta between two snapshots) is named for the month the
    window mostly falls in — the STARTING snapshot's month. Jul 5 → Aug 2 is
    ~26 days of July against 2 of August, so it reports as July. Through
    2026-08 periods were named for the destination snapshot, which ran every
    column one month ahead of the activity it described.
Every period carries its exact window; the current month gets no period until
the following month's export closes it.

Leaders column joined from the Master Wave File DATA sheet on UniversalID.
Scope: users whose Leaders value is on the canonical leader allowlist
(leader_names.CANONICAL_NAMES — same list as SQL/clean_hr).

Scope is FROZEN per snapshot (2026-09-10, her rule: "I don't want last month's
data to change; when a new file is added, add that to it"). The first time a
snapshot is processed, the Master's Leaders mapping of that day is saved to
User Mappings Monthly/_scope_ledger/leaders_<snapshot date>.csv and reused on
every later run. A person counts in a period when the END snapshot's ledger
has them on the allowlist, so a later move to "No Longer Rev Cycle" only
affects periods after the move; earlier periods keep them, and they stay in
the report under their current Leaders value ("No Longer Rev Cycle") rather
than disappearing. Delete a ledger file only to deliberately re-scope that
snapshot with today's Master.

Output: OneDrive "Main Reports/Wave_JobRole_Change_Report.xlsx"
(previous run auto-backed up to "Main Reports/Backups/")
  Dashboard        — KPI cards, charts, wave destinations table
  Summary          — overall counts, role transition tables, definitions
  By Leader        — wave movements, role change counts, role transition detail
  Wave Changes     — per-user wave history + Departed / Wave Exception flags
  Job Role Changes — per-user role history + What Changed column + flags
  New Additions    — users present in later snapshots but not the first
"""

import datetime
import io
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

# Windows console defaults to cp1252; force UTF-8 so arrow characters print cleanly
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd
import xlsxwriter

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE        = Path(r"C:\path\to\analytics")
sys.path.insert(0, str(BASE / "scripts"))
sys.path.insert(0, str(BASE / "sql"))
from onedrive_paths import (MASTER_WAVE_PATH,      # noqa: E402
                            MAIN_REPORTS_DIR, MAIN_REPORTS_BACKUPS_DIR)
from leader_names import CANONICAL_NAMES           # noqa: E402

# Monthly User Mappings CSV snapshots stay LOCAL (heavy, inputs not reports);
# the report itself + its backups live in the OneDrive Main Reports folder.
MVP_DIR     = BASE / "REPORTS/Job Role Change Reports/User Mappings Monthly"
MASTER_PATH = MASTER_WAVE_PATH
OUT_PATH    = MAIN_REPORTS_DIR / "Wave_JobRole_Change_Report.xlsx"
BACKUP_DIR  = MAIN_REPORTS_BACKUPS_DIR
# Per-snapshot frozen Leaders mappings (local, next to the CSVs they scope).
LEDGER_DIR  = MVP_DIR / "_scope_ledger"

_DATE_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})")


def _detect_snapshots(mvp_dir: Path) -> list[tuple[str, Path, datetime.datetime]]:
    """
    Find all *_UserMappingsData.csv files whose names start with YYYY-MM-DD,
    sort chronologically, and assign a month-name label.
    If two files fall in the same calendar month, both get a 'Mon DD' label.

    Returns (label, path, snapshot_date). The label names the *snapshot* (a
    point-in-time roster, e.g. "August" = as of Aug 2); change periods between
    snapshots are labelled separately — see _period_month().
    """
    found = []
    for p in mvp_dir.glob("*_UserMappingsData.csv"):
        m = _DATE_RE.match(p.name)
        if m:
            dt = datetime.datetime.strptime(m.group(1), "%Y-%m-%d")
            found.append((dt, p))
    if not found:
        raise FileNotFoundError(
            f"No *_UserMappingsData.csv files found in {mvp_dir}\n"
            "Expected filenames starting with YYYY-MM-DD."
        )
    found.sort(key=lambda x: x[0])
    # Labels must stay unique: same month+year twice -> add the day;
    # same month across different years -> add the year.
    my_counts    = Counter((dt.year, dt.month) for dt, _ in found)
    month_years  = {}
    for dt, _ in found:
        month_years.setdefault(dt.month, set()).add(dt.year)
    result = []
    for dt, p in found:
        if my_counts[(dt.year, dt.month)] > 1:
            label = f"{dt.strftime('%b')} {dt.day}"
            if len(month_years[dt.month]) > 1:
                label += f" {dt.year}"
        elif len(month_years[dt.month]) > 1:
            label = dt.strftime("%b %Y")
        else:
            label = dt.strftime("%B")
        result.append((label, p, dt))
    return result


SNAPSHOTS = _detect_snapshots(MVP_DIR)
if len(SNAPSHOTS) < 2:
    raise ValueError("At least two snapshots are required to produce a change report.")

SNAP_DATES = {lbl: dt for lbl, _, dt in SNAPSHOTS}


def _fmt_day(dt: datetime.datetime) -> str:
    """'Jul 5' — no zero-padded day (%-d is not portable to Windows)."""
    return f"{dt.strftime('%b')} {dt.day}"


def _period_month(a: str) -> str:
    """
    Calendar month a change period belongs to = the month of the START snapshot.

    Exports are pulled in the first week of each month, so the a→b delta covers
    (say) Jul 5 → Aug 2: ~26 days of July against 2 of August. That is July's
    activity, not August's. Labelling by the destination snapshot — as this
    report did through 2026-08 — put every period one month later than the
    activity it described.

    MVP snapshots carry no change timestamps, so a true calendar Jul 1–31 split
    is not derivable; the window is always reported alongside the month name.
    """
    return SNAP_DATES[a].strftime("%B")


def _period_window(a: str, b: str) -> str:
    """'Jul 5 – Aug 2' — the dates the change period actually spans."""
    return f"{_fmt_day(SNAP_DATES[a])} – {_fmt_day(SNAP_DATES[b])}"


def _period_window_short(a: str, b: str) -> str:
    """'7/5-8/2' — fits the 12-char Dashboard columns without overflowing."""
    da, db = SNAP_DATES[a], SNAP_DATES[b]
    return f"{da.month}/{da.day}-{db.month}/{db.day}"

MVP_LOAD_COLS = [
    "UniversalID", "FirstName", "LastName",
    "GoLiveWave",
    "IndividualCategoryUpdate1Name",
    "IndividualCategoryUpdate2Name",
    "IndividualCategoryUpdate3Name",
    "IndividualCategoryUpdate4Name",
    "IsDepartedInactive",
    "HasWaveException",
]

JOB_ROLE_COLS   = [f"IndividualCategoryUpdate{i}Name" for i in range(1, 5)]
JOB_ROLE_LABELS = [f"Job Role {i}" for i in range(1, 5)]

# Canonical leader allowlist — single source of truth shared with SQL/clean_hr
LEADER_ALLOWLIST = frozenset(CANONICAL_NAMES)

_WAVE_PREFERRED = ["Wave 1", "Wave 2", "Wave 3", "Wave 4", "Wave 5"]


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_snapshot(path: Path, label: str) -> pd.DataFrame:
    """Load one MVP CSV. Returns DataFrame indexed by UniversalID (upper)."""
    # Read only columns that exist in this file (older exports may lack new cols)
    probe    = pd.read_csv(path, nrows=0, encoding="utf-8-sig").columns.tolist()
    use_cols = [c for c in MVP_LOAD_COLS if c in probe]

    df = pd.read_csv(path, encoding="utf-8-sig", usecols=use_cols,
                     dtype=str, keep_default_na=False)

    # Ensure missing optional columns exist as empty strings
    for col in MVP_LOAD_COLS:
        if col not in df.columns:
            df[col] = ""

    df["UniversalID"] = df["UniversalID"].str.strip().str.upper()
    df = df[df["UniversalID"] != ""].drop_duplicates("UniversalID", keep="last")
    df["GoLiveWave"] = df["GoLiveWave"].str.strip().str.title()
    for col in JOB_ROLE_COLS:
        df[col] = df[col].str.strip()

    # Normalize boolean flags to "Yes" / "No" strings so outer-join fillna("") works
    for bool_col in ("IsDepartedInactive", "HasWaveException"):
        df[bool_col] = (
            df[bool_col].str.strip().str.upper()
            .isin({"TRUE", "1", "YES"})
            .map({True: "Yes", False: "No"})
        )

    first = df["FirstName"].str.strip()
    last  = df["LastName"].str.strip()
    df[f"FullName_{label}"] = (last + ", " + first).str.strip(", ")

    rename = {
        "GoLiveWave":         f"Wave_{label}",
        "IsDepartedInactive": f"IsDepart_{label}",
        "HasWaveException":   f"HasException_{label}",
        **{col: f"{lbl}_{label}" for col, lbl in zip(JOB_ROLE_COLS, JOB_ROLE_LABELS)},
    }
    keep = (
        [f"FullName_{label}", f"Wave_{label}", f"IsDepart_{label}", f"HasException_{label}"]
        + [f"{lbl}_{label}" for lbl in JOB_ROLE_LABELS]
    )
    return df.rename(columns=rename).set_index("UniversalID")[keep]


def load_leaders() -> pd.DataFrame:
    """Load UniversalID → Full Name + Leaders from Master DATA sheet.
    Falls back to SQL raw.master (same-day copy) if the file is locked."""
    try:
        df = pd.read_excel(MASTER_PATH, sheet_name="DATA", engine="calamine", dtype=str)
    except PermissionError:
        print("  Master file is locked (open in Excel / OneDrive sync) — "
              "using SQL raw.master instead (same-day copy from the daily refresh).")
        from sqlalchemy import create_engine
        from refresh import CONN
        engine = create_engine(CONN)
        df = pd.read_sql(
            'SELECT UniversalID, Full_Name AS "Full Name", Leaders FROM raw.master',
            engine, dtype=str)
    df.columns = df.columns.str.strip()
    df["UniversalID"] = df["UniversalID"].str.strip().str.upper()
    df = df.dropna(subset=["UniversalID"])
    df = df[df["UniversalID"] != ""].drop_duplicates("UniversalID", keep="last")
    df["Leaders"]   = df["Leaders"].fillna("").str.strip()
    df["Full Name"] = df["Full Name"].fillna("").str.strip()
    return df.set_index("UniversalID")[["Full Name", "Leaders"]]


def load_scope_ledgers(leaders_now: pd.DataFrame) -> dict[str, pd.Series]:
    """Leaders mapping to use for each snapshot: the saved ledger if one exists,
    otherwise today's Master (saved now so it is frozen from here on).
    Returns {snapshot label: Series(Leaders, index=UniversalID)}."""
    LEDGER_DIR.mkdir(parents=True, exist_ok=True)
    ledgers, frozen, new = {}, [], []
    for label, _path, dt in SNAPSHOTS:
        f = LEDGER_DIR / f"leaders_{dt:%Y-%m-%d}.csv"
        if f.exists():
            df = pd.read_csv(f, dtype=str, keep_default_na=False)
            df["UniversalID"] = df["UniversalID"].str.strip().str.upper()
            ledgers[label] = df.drop_duplicates("UniversalID").set_index("UniversalID")["Leaders"].str.strip()
            frozen.append(label)
        else:
            ser = leaders_now["Leaders"].copy()
            ser.rename_axis("UniversalID").reset_index().to_csv(f, index=False, encoding="utf-8-sig")
            ledgers[label] = ser
            new.append(label)
    print(f"  Scope ledgers: {len(frozen)} frozen ({', '.join(frozen) or '-'}) | "
          f"{len(new)} saved from today's Master ({', '.join(new) or '-'})")
    return ledgers


def apply_frozen_scope(combined: pd.DataFrame, ledgers: dict, labels: list) -> pd.DataFrame:
    """Blank out a user's changes in any period whose END snapshot ledger does
    not have them on the allowlist, so later re-scoping never rewrites history."""
    for a, b in zip(labels, labels[1:]):
        in_b = combined.index.isin(ledgers[b][ledgers[b].isin(LEADER_ALLOWLIST)].index)
        combined.loc[~in_b, f"WaveChange_{a}_{b}"] = ""
        combined.loc[~in_b, f"RoleChanged_{a}_{b}"] = False
    return combined


# ---------------------------------------------------------------------------
# Dataset assembly
# ---------------------------------------------------------------------------

def build_combined(snapshots_dict: dict, leaders: pd.DataFrame) -> pd.DataFrame:
    """
    Outer-join all snapshots, inner-join to Master (scope to tracked users).
    Adds In_{label} boolean presence flags for each snapshot.
    """
    labels   = list(snapshots_dict.keys())
    combined = None
    for df in snapshots_dict.values():
        combined = df if combined is None else combined.join(df, how="outer")

    for label, df_snap in snapshots_dict.items():
        combined[f"In_{label}"] = combined.index.isin(df_snap.index)

    for col in combined.columns:
        if not pd.api.types.is_bool_dtype(combined[col]):
            combined[col] = combined[col].fillna("")

    combined = combined.join(leaders, how="inner")
    combined["Leaders"]   = combined["Leaders"].fillna("")
    combined["Full Name"] = combined["Full Name"].fillna("")

    combined["Team Member"] = combined["Full Name"]
    for label in reversed(labels):
        mask = combined["Team Member"] == ""
        combined.loc[mask, "Team Member"] = combined.loc[mask, f"FullName_{label}"]

    fn_cols = [f"FullName_{l}" for l in labels]
    return combined.drop(columns=fn_cols + ["Full Name"])


# ---------------------------------------------------------------------------
# Change detection
# ---------------------------------------------------------------------------

def wave_change_col(df: pd.DataFrame, a: str, b: str) -> pd.Series:
    """
    Return 'OldWave → NewWave' when user is present in both snapshots and wave differs.
    Blank wave shown as '(None)' — captures assignments and removals.
    Vectorized; no row-by-row apply.
    """
    wa, wb_ = f"Wave_{a}", f"Wave_{b}"
    present_in_both = df[f"In_{a}"] & df[f"In_{b}"]
    changed  = present_in_both & (df[wa] != df[wb_])
    wa_disp  = df[wa].replace("", "(None)")
    wb_disp  = df[wb_].replace("", "(None)")
    result   = wa_disp + " → " + wb_disp
    return result.where(changed, "")


def role_changed_col(df: pd.DataFrame, a: str, b: str) -> pd.Series:
    """True when any Job Role 1-4 differs between snapshots a and b (both present)."""
    present_in_both = df[f"In_{a}"] & df[f"In_{b}"]
    changed = pd.Series(False, index=df.index)
    for lbl in JOB_ROLE_LABELS:
        ca, cb = f"{lbl}_{a}", f"{lbl}_{b}"
        either_present = (df[ca] != "") | (df[cb] != "")
        changed = changed | (present_in_both & either_present & (df[ca] != df[cb]))
    return changed


def add_change_cols(df: pd.DataFrame, labels: list) -> pd.DataFrame:
    df = df.copy()
    pairs = [(labels[i], labels[i + 1]) for i in range(len(labels) - 1)]
    for a, b in pairs:
        df[f"WaveChange_{a}_{b}"]  = wave_change_col(df, a, b)
        df[f"RoleChanged_{a}_{b}"] = role_changed_col(df, a, b)
    return df


def wave_transition_counts(df: pd.DataFrame, change_col: str) -> pd.DataFrame:
    counts = (
        df.loc[df[change_col] != "", change_col]
        .value_counts()
        .reset_index()
    )
    counts.columns = ["Wave Transition", "Count"]
    return counts.sort_values("Wave Transition").reset_index(drop=True)


def role_transition_counts(df: pd.DataFrame, a: str, b: str) -> pd.DataFrame:
    """(From Role, To Role, Count) for every changed role slot in the a→b period."""
    mask    = df[f"RoleChanged_{a}_{b}"] & df[f"In_{a}"] & df[f"In_{b}"]
    changed = df[mask]
    parts   = []
    for lbl in JOB_ROLE_LABELS:
        ca, cb = f"{lbl}_{a}", f"{lbl}_{b}"
        moved  = changed.loc[changed[ca] != changed[cb], [ca, cb]].copy()
        moved.columns = ["From", "To"]
        moved["From"] = moved["From"].replace("", "(None)")
        moved["To"]   = moved["To"].replace("", "(None)")
        parts.append(moved)
    if not parts:
        return pd.DataFrame(columns=["From", "To", "Count"])
    all_moves = pd.concat(parts, ignore_index=True)
    tbl = all_moves.groupby(["From", "To"]).size().reset_index(name="Count")
    return tbl.sort_values("Count", ascending=False).reset_index(drop=True)


def describe_role_changes(role_users: pd.DataFrame, pairs: list,
                          period_month: dict) -> pd.Series:
    """Build a human-readable 'What Changed' string for each row in role_users.

    Each change is prefixed with the period's month — the month the a→b window
    falls in, not the destination snapshot's month.
    """
    results = []
    for _, r in role_users.iterrows():
        parts = []
        for a, b in pairs:
            if not r.get(f"RoleChanged_{a}_{b}", False):
                continue
            if not (r.get(f"In_{a}", False) and r.get(f"In_{b}", False)):
                continue
            slot_changes = []
            for slot in JOB_ROLE_LABELS:
                va, vb = r[f"{slot}_{a}"], r[f"{slot}_{b}"]
                if va != vb:
                    slot_changes.append(f"{va or '(None)'} → {vb or '(None)'}")
            if slot_changes:
                parts.append(f"{period_month[f'{a}_{b}']}: " + "; ".join(slot_changes))
        results.append(" | ".join(parts))
    return pd.Series(results, index=role_users.index)


# ---------------------------------------------------------------------------
# Excel formatting helpers
# ---------------------------------------------------------------------------

def make_formats(wb: xlsxwriter.Workbook) -> dict:
    return {
        "title":        wb.add_format({"bold": True, "font_size": 13, "font_color": "#1F3864"}),
        "dash_section": wb.add_format({"bold": True, "bg_color": "#F7F8FA",
                                        "font_color": "#374151", "border": 1,
                                        "border_color": "#D1D5DB", "font_size": 10}),
        "dash_kpi_lbl": wb.add_format({"bg_color": "#FFFFFF", "font_color": "#6B7280",
                                        "border": 1, "border_color": "#E5E7EB",
                                        "align": "center", "valign": "vcenter",
                                        "text_wrap": True, "font_size": 9}),
        "dash_kpi_val": wb.add_format({"bold": True, "font_size": 20,
                                        "font_color": "#1F2937", "bg_color": "#FAFAFA",
                                        "border": 1, "border_color": "#E5E7EB",
                                        "num_format": "#,##0", "align": "center",
                                        "valign": "vcenter"}),
        "dash_kpi_dep": wb.add_format({"bold": True, "font_size": 16,
                                        "font_color": "#9CA3AF", "bg_color": "#F9FAFB",
                                        "border": 1, "border_color": "#E5E7EB",
                                        "num_format": "#,##0", "align": "center",
                                        "valign": "vcenter"}),
        "dash_tbl_hdr": wb.add_format({"bold": True, "bg_color": "#F3F4F6",
                                        "font_color": "#374151", "border": 1,
                                        "border_color": "#E5E7EB", "font_size": 9}),
        "tbl_dat": wb.add_format({"border": 1, "border_color": "#E5E7EB", "font_size": 9}),
        "tbl_num": wb.add_format({"border": 1, "border_color": "#E5E7EB", "font_size": 9,
                                   "num_format": "#,##0"}),
        "gray":    wb.add_format({"italic": True, "font_color": "#888888"}),
        "section": wb.add_format({"bold": True, "font_size": 11,
                                   "bg_color": "#1F3864", "font_color": "#FFFFFF",
                                   "border": 1}),
        "header":  wb.add_format({"bold": True, "bg_color": "#D6E4F0",
                                   "border": 1, "text_wrap": True, "valign": "vcenter"}),
        "subhdr":  wb.add_format({"bold": True, "bg_color": "#E9ECEF", "border": 1}),
        "data":    wb.add_format({"border": 1}),
        "num":     wb.add_format({"border": 1, "num_format": "#,##0"}),
        "arrow":   wb.add_format({"border": 1, "bg_color": "#FFF2CC"}),
        "new":     wb.add_format({"border": 1, "bg_color": "#E2EFDA"}),
        "bold":    wb.add_format({"bold": True}),
        "yes":     wb.add_format({"border": 1, "bg_color": "#C6EFCE"}),
        "no":      wb.add_format({"border": 1}),
        "depart":  wb.add_format({"border": 1, "bg_color": "#F2F2F2",
                                   "font_color": "#9CA3AF", "italic": True}),
        "exc":     wb.add_format({"border": 1, "bg_color": "#FCE4D6"}),
        "what_chg": wb.add_format({"border": 1, "bg_color": "#EBF3FB",
                                    "text_wrap": True, "font_size": 8}),
        "role_from": wb.add_format({"border": 1, "bg_color": "#FFF2CC", "font_size": 9}),
        "role_to":   wb.add_format({"border": 1, "bg_color": "#E2EFDA", "font_size": 9}),
        # Dashboard refactor additions
        "divider":       wb.add_format({"bottom": 2, "border_color": "#D1D5DB"}),
        "dash_subhdr":   wb.add_format({"bold": True, "font_size": 9,
                                         "font_color": "#6B7280", "align": "left"}),
        "dash_kpi_lbl_up": wb.add_format({"bg_color": "#F0F9F4", "font_color": "#1F7A3F",
                                           "border": 1, "border_color": "#D1FAE5",
                                           "align": "center", "valign": "vcenter",
                                           "text_wrap": True, "font_size": 9, "bold": True}),
        "dash_kpi_val_up": wb.add_format({"bold": True, "font_size": 18,
                                           "font_color": "#1F7A3F", "bg_color": "#F0F9F4",
                                           "border": 1, "border_color": "#D1FAE5",
                                           "num_format": "#,##0", "align": "center",
                                           "valign": "vcenter"}),
        "dash_kpi_lbl_dn": wb.add_format({"bg_color": "#FEF2F2", "font_color": "#9B1C1C",
                                           "border": 1, "border_color": "#FECACA",
                                           "align": "center", "valign": "vcenter",
                                           "text_wrap": True, "font_size": 9, "bold": True}),
        "dash_kpi_val_dn": wb.add_format({"bold": True, "font_size": 18,
                                           "font_color": "#9B1C1C", "bg_color": "#FEF2F2",
                                           "border": 1, "border_color": "#FECACA",
                                           "num_format": "#,##0", "align": "center",
                                           "valign": "vcenter"}),
        "alt_dat":  wb.add_format({"border": 1, "border_color": "#E5E7EB",
                                    "bg_color": "#F9FAFB", "font_size": 9}),
        "alt_num":  wb.add_format({"border": 1, "border_color": "#E5E7EB",
                                    "bg_color": "#F9FAFB", "font_size": 9,
                                    "num_format": "#,##0"}),
        "chart_title_font": {"size": 10, "bold": True, "color": "#374151"},
    }


def section_header(ws, row: int, text: str, fmt, ncols: int = 4) -> int:
    ws.write(row, 0, text, fmt)
    for c in range(1, ncols):
        ws.write(row, c, "", fmt)
    return row + 1


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    run_date = datetime.date.today().strftime("%m/%d/%Y")

    # -- Abort early if the output workbook is open in Excel -------------------
    if OUT_PATH.exists():
        lock = OUT_PATH.with_name("~$" + OUT_PATH.name)
        try:
            with open(OUT_PATH, "r+b"):
                pass
        except PermissionError:
            sys.exit(f"ABORT: {OUT_PATH.name} is open in Excel — close it and rerun.")
        if lock.exists():
            print(f"  (note: stale Excel lock file {lock.name} present but file is writable)")

    # -- Backup existing output ------------------------------------------------
    if OUT_PATH.exists():
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        bak = BACKUP_DIR / f"{OUT_PATH.stem}_{timestamp}.bak.xlsx"
        shutil.copy(str(OUT_PATH), str(bak))
        print(f"Backed up previous report → Backups\\{bak.name}")

    # -- Load snapshots --------------------------------------------------------
    print("Loading MVP snapshots...")
    labels         = [lbl for lbl, _, _ in SNAPSHOTS]
    snapshots_dict = {}
    for label, path, _dt in SNAPSHOTS:
        snapshots_dict[label] = load_snapshot(path, label)
        print(f"  {label:8s}: {len(snapshots_dict[label]):,} unique UIDs  ({path.name})")

    # -- Load leaders ----------------------------------------------------------
    print("Loading Leaders from Master Wave File...")
    leaders = load_leaders()
    print(f"  {len(leaders):,} users in Master")
    ledgers = load_scope_ledgers(leaders)
    # Ever in scope = on the allowlist in ANY snapshot's ledger. People who
    # later left Rev Cycle stay in the report under their current Leaders value.
    ever_scope = set()
    for ser in ledgers.values():
        ever_scope |= set(ser[ser.isin(LEADER_ALLOWLIST)].index)
    # Users dropped from today's Master but present in an older ledger keep
    # their last ledger Leaders value so the inner join does not lose them.
    missing = [u for u in ever_scope if u not in leaders.index]
    if missing:
        last = {}
        for ser in ledgers.values():
            for u in missing:
                if u in ser.index:
                    last[u] = ser[u]
        extra = pd.DataFrame({"Full Name": "", "Leaders": pd.Series(last)})
        extra.index.name = "UniversalID"
        leaders = pd.concat([leaders, extra])
        print(f"  {len(missing):,} user(s) no longer on the Master kept from older ledgers")

    # -- Build combined dataset ------------------------------------------------
    print("Building combined dataset...")
    combined = build_combined(snapshots_dict, leaders)
    combined = add_change_cols(combined, labels)
    combined = apply_frozen_scope(combined, ledgers, labels)
    in_scope = combined[combined.index.isin(ever_scope)].copy()
    left_rc_n = int((~in_scope["Leaders"].isin(LEADER_ALLOWLIST)).sum())
    print(f"  {len(combined):,} in Master  |  {len(in_scope):,} in scope "
          f"(of which {left_rc_n:,} since left Rev Cycle, kept)")

    pairs         = [(labels[i], labels[i + 1]) for i in range(len(labels) - 1)]
    period_keys   = [f"{a}_{b}" for a, b in pairs]

    # Period label = the month the change window actually falls in (the START
    # snapshot's month) plus the exact window. Snapshots land in the first week
    # of a month, so Jul 5 → Aug 2 is July's activity — labelling it by the
    # destination snapshot ("August") ran every period a month ahead of itself.
    period_month  = {f"{a}_{b}": _period_month(a)     for a, b in pairs}
    period_window = {f"{a}_{b}": _period_window(a, b) for a, b in pairs}
    period_win_sh = {f"{a}_{b}": _period_window_short(a, b) for a, b in pairs}
    period_labels = {k: f"{period_month[k]}  ({period_window[k]})" for k in period_keys}
    # Compact form for chart category axes — month alone unless it repeats
    # (two snapshots inside one calendar month), in which case add the window.
    _month_counts = Counter(period_month.values())
    period_chart  = {
        k: (v if _month_counts[v] == 1 else f"{v} ({period_window[k]})")
        for k, v in period_month.items()
    }

    # Departed / wave-exception: use most recent snapshot the user appears in
    def _latest_flag(r, col_prefix):
        for lbl in reversed(labels):
            if r.get(f"In_{lbl}", False):
                return r.get(f"{col_prefix}{lbl}", "")
        return ""

    in_scope = in_scope.copy()
    in_scope["Departed"]   = in_scope.apply(lambda r: _latest_flag(r, "IsDepart_"),    axis=1)
    in_scope["WaveExcept"] = in_scope.apply(lambda r: _latest_flag(r, "HasException_"), axis=1)
    departed_n = int((in_scope["Departed"] == "Yes").sum())

    # -- Wave names from actual data (not hardcoded) ---------------------------
    all_wave_vals = {
        v for lbl in labels for v in in_scope[f"Wave_{lbl}"].unique() if v
    }
    wave_names = sorted(
        all_wave_vals,
        key=lambda w: _WAVE_PREFERRED.index(w) if w in _WAVE_PREFERRED else 99,
    )

    # -- Derive sub-populations ------------------------------------------------
    wave_users = in_scope[
        in_scope[[f"WaveChange_{a}_{b}" for a, b in pairs]].ne("").any(axis=1)
    ].copy().sort_values(["Leaders", "Team Member"])

    role_users = in_scope[
        in_scope[[f"RoleChanged_{a}_{b}" for a, b in pairs]].any(axis=1)
    ].copy().sort_values(["Leaders", "Team Member"])

    new_users = in_scope[~in_scope[f"In_{labels[0]}"]].copy()
    new_users = new_users.sort_values(["Leaders", "Team Member"])

    print(f"  Wave changes: {len(wave_users):,}  |  Role changes: {len(role_users):,}  |  "
          f"New: {len(new_users):,}  |  Departed: {departed_n:,}")

    # Precompute "What Changed" for Job Role Changes sheet
    role_what_changed = describe_role_changes(role_users, pairs, period_month)

    # -- Chart / dashboard data ------------------------------------------------
    wave_ch_data = {key: int((in_scope[f"WaveChange_{key}"] != "").sum()) for key in period_keys}
    role_ch_data = {key: int(in_scope[f"RoleChanged_{key}"].sum()) for key in period_keys}

    new_by_snap: dict[str, int] = {}
    for i, lbl in enumerate(labels[1:], start=1):
        prior = labels[:i]
        new_by_snap[lbl] = len(in_scope[
            ~in_scope[[f"In_{p}" for p in prior]].any(axis=1) & in_scope[f"In_{lbl}"]
        ])

    changed_uids = set(wave_users.index) | set(role_users.index) | set(new_users.index)
    leader_total = (
        in_scope[in_scope.index.isin(changed_uids)]
        .groupby("Leaders").size()
        .sort_values(ascending=False).head(10)
        .reset_index(name="Count")
    )
    leader_total.columns = ["Leader", "Count"]

    leader_dash_rows = [
        (leader,
         int((wave_users["Leaders"] == leader).sum()),
         int((role_users["Leaders"] == leader).sum()))
        for leader in sorted(LEADER_ALLOWLIST)
    ]

    _all_wc = [f"WaveChange_{key}" for key in period_keys]
    w1_to_w2 = int((in_scope[_all_wc] == "Wave 1 → Wave 2").any(axis=1).sum())
    w2_to_w3 = int((in_scope[_all_wc] == "Wave 2 → Wave 3").any(axis=1).sum())
    w3_to_w2 = int((in_scope[_all_wc] == "Wave 3 → Wave 2").any(axis=1).sum())
    w2_to_w1 = int((in_scope[_all_wc] == "Wave 2 → Wave 1").any(axis=1).sum())

    wave_dest_data = {}
    for key in period_keys:
        wc_col  = f"WaveChange_{key}"
        changed = in_scope[in_scope[wc_col] != ""]
        wave_dest_data[key] = {w: int(changed[wc_col].str.endswith(w).sum()) for w in wave_names}

    latest_lbl       = labels[-1]
    wave_dist_latest = {w: int((in_scope[f"Wave_{latest_lbl}"] == w).sum()) for w in wave_names}

    # FIXED donut categories — mutually exclusive, sum to len(in_scope)
    # "New" takes priority (entry event); wave+role both excludes new users
    wave_uid_set = set(wave_users.index)
    role_uid_set = set(role_users.index)
    new_uid_set  = set(new_users.index)
    wave_only_n  = len(wave_uid_set - role_uid_set - new_uid_set)
    role_only_n  = len(role_uid_set - wave_uid_set - new_uid_set)
    both_n       = len((wave_uid_set & role_uid_set) - new_uid_set)
    new_n        = len(new_uid_set)
    stable_n     = len(in_scope) - len(wave_uid_set | role_uid_set | new_uid_set)

    # Aggregated role transition table (all periods combined) for Summary sheet
    all_rt_parts = [role_transition_counts(in_scope, a, b) for a, b in pairs]
    all_rt_parts = [df for df in all_rt_parts if not df.empty]
    if all_rt_parts:
        all_rt = (
            pd.concat(all_rt_parts, ignore_index=True)
            .groupby(["From", "To"])["Count"].sum()
            .reset_index()
            .sort_values("Count", ascending=False)
            .reset_index(drop=True)
        )
    else:
        all_rt = pd.DataFrame(columns=["From", "To", "Count"])

    # =========================================================================
    print(f"\nWriting {OUT_PATH.name}...")
    wb  = xlsxwriter.Workbook(str(OUT_PATH))
    wb.set_calc_mode("manual")
    fmt = make_formats(wb)

    # =========================================================================
    # Dashboard
    # =========================================================================
    nperiods = len(period_keys)
    n_waves  = len(wave_names)
    n_snaps  = len(new_by_snap)

    # ---- Hidden _ChartData sheet --------------------------------------------
    # All chart source data lives here so the Dashboard sheet stays presentation-only.
    ws_cd = wb.add_worksheet("_ChartData")
    ws_cd.hide()
    HC = 0  # _ChartData uses cols A (label) + B (value)

    # T1: wave changes by period
    ws_cd.write(0, HC,     "Period")
    ws_cd.write(0, HC + 1, "Wave Changes")
    for i, key in enumerate(period_keys):
        ws_cd.write(1 + i, HC,     period_chart[key])
        ws_cd.write(1 + i, HC + 1, wave_ch_data[key])

    # T2: role changes by period
    t2_base = nperiods + 2
    ws_cd.write(t2_base, HC,     "Period")
    ws_cd.write(t2_base, HC + 1, "Role Changes")
    for i, key in enumerate(period_keys):
        ws_cd.write(t2_base + 1 + i, HC,     period_chart[key])
        ws_cd.write(t2_base + 1 + i, HC + 1, role_ch_data[key])

    # T3: new additions by first snapshot
    t3_base = t2_base + nperiods + 2
    ws_cd.write(t3_base, HC,     "First Snapshot")
    ws_cd.write(t3_base, HC + 1, "New Additions")
    for i, (snap_lbl, snap_val) in enumerate(new_by_snap.items()):
        ws_cd.write(t3_base + 1 + i, HC,     snap_lbl)
        ws_cd.write(t3_base + 1 + i, HC + 1, snap_val)

    # T4: top 10 leaders (for the horizontal bar chart)
    t4_base   = t3_base + n_snaps + 2
    n_leaders = len(leader_total)
    ws_cd.write(t4_base, HC,     "Leader")
    ws_cd.write(t4_base, HC + 1, "Total Changes")
    for i, lt_row in enumerate(leader_total.itertuples(index=False)):
        ws_cd.write(t4_base + 1 + i, HC,     lt_row.Leader)
        ws_cd.write(t4_base + 1 + i, HC + 1, int(lt_row.Count))

    # T5: wave destinations by period (Period | →W1 | →W2 | ...)
    t5_base = t4_base + n_leaders + 2
    ws_cd.write(t5_base, HC, "Period")
    for wi, wn in enumerate(wave_names):
        ws_cd.write(t5_base, HC + 1 + wi, f"→ {wn}")
    for i, key in enumerate(period_keys):
        ws_cd.write(t5_base + 1 + i, HC, period_chart[key])
        for wi, wn in enumerate(wave_names):
            ws_cd.write(t5_base + 1 + i, HC + 1 + wi, wave_dest_data[key][wn])

    # ---- Dashboard sheet ----------------------------------------------------
    ws_d = wb.add_worksheet("Dashboard")
    ws_d.activate()
    ws_d.set_first_sheet()
    ws_d.hide_gridlines(2)
    ws_d.set_zoom(90)

    # Column widths: A=margin, B-K=content (10 cols), L=margin
    ws_d.set_column(0, 0, 2)            # A: left margin
    ws_d.set_column(1, 10, 12)          # B-K: main content (12 chars ≈ 90px)
    ws_d.set_column(11, 11, 2)          # L: right margin

    # Row positions (0-indexed)
    R_TITLE       = 0
    R_SUB         = 1
    R_DIV1        = 2
    R_SEC_KPI     = 3
    R_KPI_LBL     = 4
    R_KPI_VAL     = 5
    R_SNAP_LBL    = 6
    R_SNAP_VAL    = 7
    R_WAVE_SUB    = 8
    R_WAVE_R1_LBL = 9
    R_WAVE_R1_VAL = 10
    R_WAVE_R2_LBL = 11
    R_WAVE_R2_VAL = 12
    R_DIV2        = 13
    R_SEC_DEST    = 14
    R_DEST_HDR    = 15
    R_DEST_FIRST  = 16
    R_DEST_LAST   = R_DEST_FIRST + nperiods - 1
    R_DIV3        = R_DEST_LAST + 1
    R_CHART_A     = R_DIV3 + 1
    R_CHART_B     = R_CHART_A + 15
    R_DIV4        = R_CHART_B + 15
    R_SEC_LDR     = R_DIV4 + 1
    R_LDR         = R_SEC_LDR + 1   # header row of leader table / top of leader chart

    # Row heights
    ws_d.set_row(R_TITLE, 24)
    ws_d.set_row(R_SUB,   14)
    ws_d.set_row(R_DIV1,   6)
    ws_d.set_row(R_SEC_KPI, 18)
    ws_d.set_row(R_KPI_LBL,  20)
    ws_d.set_row(R_KPI_VAL,  38)
    ws_d.set_row(R_SNAP_LBL, 20)
    ws_d.set_row(R_SNAP_VAL, 38)
    ws_d.set_row(R_WAVE_SUB, 16)
    ws_d.set_row(R_WAVE_R1_LBL, 18)
    ws_d.set_row(R_WAVE_R1_VAL, 34)
    ws_d.set_row(R_WAVE_R2_LBL, 18)
    ws_d.set_row(R_WAVE_R2_VAL, 34)
    ws_d.set_row(R_DIV2,   6)
    ws_d.set_row(R_SEC_DEST, 18)
    ws_d.set_row(R_DIV3,   6)
    ws_d.set_row(R_DIV4,   6)
    ws_d.set_row(R_SEC_LDR, 18)

    # Title + subtitle
    ws_d.merge_range(R_TITLE, 1, R_TITLE, 10,
                     f"Wave & Job Role Change Report  —  {run_date}", fmt["title"])
    ws_d.merge_range(R_SUB, 1, R_SUB, 10,
        f"Snapshots: {' | '.join(labels)}  |  "
        f"In-scope: {len(in_scope):,} users  |  16-leader allowlist",
        fmt["gray"])

    # Divider 1
    for c in range(1, 11):
        ws_d.write(R_DIV1, c, "", fmt["divider"])

    # Section: POPULATION KPIs
    ws_d.merge_range(R_SEC_KPI, 1, R_SEC_KPI, 10, "POPULATION KPIs", fmt["dash_section"])

    # Primary row — 5 KPI cards across B-K (each 2 cols)
    primary_kpis = [
        ("Total In-Scope Users",            len(in_scope)),
        ("Users w/ Wave Changes",           len(wave_users)),
        ("Users w/ Role Changes",           len(role_users)),
        (f"New Additions (vs {labels[0]})", len(new_users)),
        ("Departed (latest snapshot)",      departed_n),
    ]
    for idx, (lbl, val) in enumerate(primary_kpis):
        c = 1 + idx * 2
        lbl_fmt = fmt["dash_kpi_lbl"]
        val_fmt = fmt["dash_kpi_dep"] if lbl.startswith("Departed") else fmt["dash_kpi_val"]
        ws_d.merge_range(R_KPI_LBL, c, R_KPI_LBL, c + 1, lbl, lbl_fmt)
        ws_d.merge_range(R_KPI_VAL, c, R_KPI_VAL, c + 1, val, val_fmt)

    # Snapshot row — 4 cards (most recent snapshots), centered visually
    snap_labels = labels[-4:]
    for idx, lbl in enumerate(snap_labels):
        # Cards at B-C, D-E, G-H, I-J (gap at F, tail at K) for visual balance
        c_starts = [1, 3, 6, 8]
        c = c_starts[idx]
        ws_d.merge_range(R_SNAP_LBL, c, R_SNAP_LBL, c + 1, f"{lbl} Snapshot", fmt["dash_kpi_lbl"])
        ws_d.merge_range(R_SNAP_VAL, c, R_SNAP_VAL, c + 1,
                         int(in_scope[f"In_{lbl}"].sum()), fmt["dash_kpi_val"])

    # Wave movements subheader + 2x2 sub-grid (ups left B-E, downs right H-K)
    ws_d.merge_range(R_WAVE_SUB, 1, R_WAVE_SUB, 10, "Wave Movements", fmt["dash_subhdr"])
    # Top-left up, top-right down
    ws_d.merge_range(R_WAVE_R1_LBL, 1, R_WAVE_R1_LBL, 4, "Wave 1 → 2  ↑", fmt["dash_kpi_lbl_up"])
    ws_d.merge_range(R_WAVE_R1_VAL, 1, R_WAVE_R1_VAL, 4, w1_to_w2,        fmt["dash_kpi_val_up"])
    ws_d.merge_range(R_WAVE_R1_LBL, 7, R_WAVE_R1_LBL, 10, "Wave 3 → 2  ↓", fmt["dash_kpi_lbl_dn"])
    ws_d.merge_range(R_WAVE_R1_VAL, 7, R_WAVE_R1_VAL, 10, w3_to_w2,        fmt["dash_kpi_val_dn"])
    # Bottom-left up, bottom-right down
    ws_d.merge_range(R_WAVE_R2_LBL, 1, R_WAVE_R2_LBL, 4, "Wave 2 → 3  ↑", fmt["dash_kpi_lbl_up"])
    ws_d.merge_range(R_WAVE_R2_VAL, 1, R_WAVE_R2_VAL, 4, w2_to_w3,        fmt["dash_kpi_val_up"])
    ws_d.merge_range(R_WAVE_R2_LBL, 7, R_WAVE_R2_LBL, 10, "Wave 2 → 1  ↓", fmt["dash_kpi_lbl_dn"])
    ws_d.merge_range(R_WAVE_R2_VAL, 7, R_WAVE_R2_VAL, 10, w2_to_w1,        fmt["dash_kpi_val_dn"])

    # Divider 2
    for c in range(1, 11):
        ws_d.write(R_DIV2, c, "", fmt["divider"])

    # Section: WAVE DESTINATIONS BY PERIOD
    ws_d.merge_range(R_SEC_DEST, 1, R_SEC_DEST, 10,
                     "WAVE DESTINATIONS BY PERIOD   "
                     "(period = the month the change window falls in)",
                     fmt["dash_section"])
    ws_d.write(R_DEST_HDR, 1, "Period",  fmt["dash_tbl_hdr"])
    ws_d.write(R_DEST_HDR, 2, "Window",  fmt["dash_tbl_hdr"])
    for wi, wn in enumerate(wave_names):
        ws_d.write(R_DEST_HDR, 3 + wi, f"→ {wn}", fmt["dash_tbl_hdr"])
    for i, key in enumerate(period_keys):
        row = R_DEST_FIRST + i
        ws_d.write(row, 1, period_month[key],  fmt["tbl_dat"])
        ws_d.write(row, 2, period_win_sh[key], fmt["tbl_dat"])
        for wi, wn in enumerate(wave_names):
            ws_d.write(row, 3 + wi, wave_dest_data[key][wn], fmt["tbl_num"])

    # Divider 3
    for c in range(1, 11):
        ws_d.write(R_DIV3, c, "", fmt["divider"])

    # ---- Charts -------------------------------------------------------------
    C_WAVE = (["#A8C5DA", "#7EB5A6", "#F4A261", "#C9B1D9", "#B5B5B5"] + ["#888888"] * 10)[:n_waves]
    TITLE_FONT = fmt["chart_title_font"]

    def _style_chart(c, *, legend=False, x_grid=False, y_grid=True, y_reverse=False):
        if legend:
            c.set_legend({"position": "bottom", "font": {"size": 9}})
        else:
            c.set_legend({"none": True})
        c.set_x_axis({"major_gridlines": {"visible": x_grid, "line": {"color": "#E5E7EB"}},
                      "line": {"none": True},
                      "num_font": {"size": 9, "color": "#374151"}})
        y_ax = {"major_gridlines": {"visible": y_grid, "line": {"color": "#E5E7EB"}},
                "line": {"none": True},
                "num_font": {"size": 9, "color": "#374151"},
                "min": 0}
        if y_reverse:
            y_ax["reverse"] = True
        c.set_y_axis(y_ax)
        c.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
        c.set_plotarea({"border": {"none": True}})

    def _col_chart(title, cats, vals, color):
        c = wb.add_chart({"type": "column"})
        c.add_series({"name": title, "categories": cats, "values": vals,
                      "fill": {"color": color},
                      "data_labels": {"value": True, "position": "outside_end",
                                      "font": {"size": 8, "color": "#374151"}}})
        c.set_title({"name": title, "name_font": TITLE_FONT})
        _style_chart(c)
        return c

    # Top chart row: Wave Changes (left) | Job Role Changes (right)
    chart_wc = _col_chart(
        "Wave Changes by Period",
        ["_ChartData", 1, 0, nperiods, 0],
        ["_ChartData", 1, 1, nperiods, 1],
        "#5B8DB8",
    )
    chart_wc.set_size({"width": 420, "height": 280})
    ws_d.insert_chart(R_CHART_A, 1, chart_wc)

    chart_rc = _col_chart(
        "Job Role Changes by Period",
        ["_ChartData", t2_base + 1, 0, t2_base + nperiods, 0],
        ["_ChartData", t2_base + 1, 1, t2_base + nperiods, 1],
        "#88ADC8",
    )
    chart_rc.set_size({"width": 420, "height": 280})
    ws_d.insert_chart(R_CHART_A, 6, chart_rc)

    # Middle chart row: Wave Destinations (wide, 60%) | New Additions (narrow, 40%)
    chart_dest = wb.add_chart({"type": "column"})
    for wi, (wn, wc) in enumerate(zip(wave_names, C_WAVE)):
        chart_dest.add_series({
            "name":        wn,
            "categories":  ["_ChartData", t5_base + 1, 0,          t5_base + nperiods, 0],
            "values":      ["_ChartData", t5_base + 1, 1 + wi,     t5_base + nperiods, 1 + wi],
            "fill":        {"color": wc},
            "data_labels": {"value": True, "position": "outside_end",
                            "font": {"size": 8, "color": "#374151"}},
        })
    chart_dest.set_title({"name": "Users Moving to Each Wave by Period", "name_font": TITLE_FONT})
    _style_chart(chart_dest, legend=True)
    chart_dest.set_size({"width": 520, "height": 280})
    ws_d.insert_chart(R_CHART_B, 1, chart_dest)

    chart_new = _col_chart(
        "New Additions by First Snapshot",
        ["_ChartData", t3_base + 1, 0, t3_base + n_snaps, 0],
        ["_ChartData", t3_base + 1, 1, t3_base + n_snaps, 1],
        "#85B369",
    )
    chart_new.set_size({"width": 320, "height": 280})
    ws_d.insert_chart(R_CHART_B, 7, chart_new)

    # Divider 4
    for c in range(1, 11):
        ws_d.write(R_DIV4, c, "", fmt["divider"])

    # ---- Leader section: side-by-side chart + table ------------------------
    ws_d.merge_range(R_SEC_LDR, 1, R_SEC_LDR, 10, "CHANGES BY LEADER", fmt["dash_section"])

    # Leader chart (left, cols B-F)
    chart_leaders = wb.add_chart({"type": "bar"})
    chart_leaders.add_series({
        "name":        "Total Changes",
        "categories":  ["_ChartData", t4_base + 1, 0, t4_base + n_leaders, 0],
        "values":      ["_ChartData", t4_base + 1, 1, t4_base + n_leaders, 1],
        "fill":        {"color": "#7A9BB5"},
        "data_labels": {"value": True, "position": "outside_end",
                        "font": {"size": 8, "color": "#374151"}},
    })
    chart_leaders.set_title({"name": "Top Leaders by Total Changes", "name_font": TITLE_FONT})
    _style_chart(chart_leaders, x_grid=True, y_grid=False, y_reverse=True)
    chart_leaders.set_size({"width": 440, "height": 360})
    ws_d.insert_chart(R_LDR, 1, chart_leaders)

    # Leader table (right, cols H-K): Leader | Wave | Role | Total — sorted desc by Total
    leader_table_rows = sorted(
        [(leader, wc, rc, wc + rc) for (leader, wc, rc) in leader_dash_rows],
        key=lambda x: x[3], reverse=True,
    )
    ws_d.write(R_LDR, 7, "Leader",     fmt["header"])
    ws_d.write(R_LDR, 8, "Wave Chg",   fmt["header"])
    ws_d.write(R_LDR, 9, "Role Chg",   fmt["header"])
    ws_d.write(R_LDR, 10, "Total",     fmt["header"])
    for i, (leader, wc, rc, total) in enumerate(leader_table_rows):
        r = R_LDR + 1 + i
        # Alt row shading
        row_dat = fmt["alt_dat"] if (i % 2 == 1) else fmt["tbl_dat"]
        row_num = fmt["alt_num"] if (i % 2 == 1) else fmt["tbl_num"]
        ws_d.write(r, 7, leader, row_dat)
        ws_d.write(r, 8, wc,     row_num)
        ws_d.write(r, 9, rc,     row_num)
        ws_d.write(r, 10, total, row_num)

    # Data bars on Total column
    if leader_table_rows:
        last_row = R_LDR + len(leader_table_rows)
        ws_d.conditional_format(R_LDR + 1, 10, last_row, 10, {
            "type": "data_bar",
            "bar_color": "#7A9BB5",
            "bar_solid": True,
            "bar_only": False,
        })

    # ---- Freeze + print area ------------------------------------------------
    ws_d.freeze_panes(3, 0)  # freeze rows 0-2 (title, subtitle, divider)
    last_data_row = R_LDR + max(1, len(leader_table_rows))
    ws_d.print_area(0, 0, last_data_row, 11)
    ws_d.set_landscape()
    ws_d.fit_to_pages(1, 1)

    # =========================================================================
    # Summary
    # =========================================================================
    ws1 = wb.add_worksheet("Summary")
    ws1.set_column(0, 0, 36)
    ws1.set_column(1, 2, 20)
    ws1.set_column(3, 3, 60)
    ws1.hide_gridlines(2)

    ws1.write("A1", f"Wave & Job Role Change Report  —  {run_date}", fmt["title"])
    ws1.write("A2", f"Snapshots: {' | '.join(labels)}  |  Leaders from Master Wave File",
              fmt["gray"])

    row = 4
    row = section_header(ws1, row, "SNAPSHOT SIZES", fmt["section"], 3)
    ws1.write(row, 0, "Snapshot",  fmt["subhdr"])
    ws1.write(row, 1, "Users",     fmt["subhdr"])
    ws1.write(row, 2, "Pulled",    fmt["subhdr"]); row += 1
    for lbl in labels:
        ws1.write(row, 0, lbl, fmt["bold"])
        ws1.write(row, 1, int(in_scope[f"In_{lbl}"].sum()), fmt["num"])
        ws1.write(row, 2, SNAP_DATES[lbl].strftime("%m/%d/%Y"), fmt["data"])
        row += 1
    ws1.write(row, 0, "In Master Wave File",              fmt["bold"])
    ws1.write(row, 1, len(combined),                      fmt["num"]); row += 1
    ws1.write(row, 0, "In Scope (16-leader allowlist)",   fmt["bold"])
    ws1.write(row, 1, len(in_scope),                      fmt["num"]); row += 1
    ws1.write(row, 0, "  of which since left Rev Cycle (kept)", fmt["data"])
    ws1.write(row, 1, left_rc_n,                          fmt["num"]); row += 1
    ws1.write(row, 0, "Departed Users in Scope (latest)", fmt["bold"])
    ws1.write(row, 1, departed_n,                         fmt["num"]); row += 3

    row = section_header(ws1, row,
        "WAVE MOVEMENTS  (in-scope users only — period = month the window falls in)",
        fmt["section"], 3)
    for key in period_keys:
        ws1.write(row, 0, period_labels[key], fmt["subhdr"])
        ws1.write(row, 1, "Count",            fmt["subhdr"])
        row += 1
        tbl = wave_transition_counts(in_scope, f"WaveChange_{key}")
        if tbl.empty:
            ws1.write(row, 0, "No wave changes detected", fmt["gray"]); row += 1
        else:
            for _, r in tbl.iterrows():
                ws1.write(row, 0, r["Wave Transition"], fmt["arrow"])
                ws1.write(row, 1, int(r["Count"]),      fmt["num"]); row += 1
            ws1.write(row, 0, "Total users with wave change", fmt["bold"])
            ws1.write(row, 1, int((in_scope[f"WaveChange_{key}"] != "").sum()),
                      fmt["num"]); row += 1
        row += 1

    row += 1
    row = section_header(ws1, row, "JOB ROLE CHANGES  (in-scope users only)", fmt["section"], 3)
    ws1.write(row, 0, "Period",                      fmt["subhdr"])
    ws1.write(row, 1, "Users w/ >= 1 Role Change",   fmt["subhdr"]); row += 1
    for key in period_keys:
        ws1.write(row, 0, period_labels[key],                        fmt["data"])
        ws1.write(row, 1, int(in_scope[f"RoleChanged_{key}"].sum()), fmt["num"]); row += 1

    row += 2
    row = section_header(ws1, row, "ROLE TRANSITIONS — FROM → TO  (all periods combined)",
                         fmt["section"], 4)
    if all_rt.empty:
        ws1.write(row, 0, "No role transitions detected", fmt["gray"]); row += 1
    else:
        ws1.write(row, 0, "From Role", fmt["subhdr"])
        ws1.write(row, 1, "To Role",   fmt["subhdr"])
        ws1.write(row, 2, "Count",     fmt["subhdr"]); row += 1
        for _, rr in all_rt.iterrows():
            ws1.write(row, 0, rr["From"],      fmt["role_from"])
            ws1.write(row, 1, rr["To"],        fmt["role_to"])
            ws1.write(row, 2, int(rr["Count"]), fmt["tbl_num"]); row += 1

    row += 2
    row = section_header(ws1, row,
        f"NEW ADDITIONS  (in-scope, not in {labels[0]} snapshot)", fmt["section"], 3)
    ws1.write(row, 0, "First appears in", fmt["subhdr"])
    ws1.write(row, 1, "Users",            fmt["subhdr"]); row += 1
    for i, lbl in enumerate(labels[1:], start=1):
        prior  = labels[:i]
        subset = in_scope[
            ~in_scope[[f"In_{p}" for p in prior]].any(axis=1) & in_scope[f"In_{lbl}"]
        ]
        ws1.write(row, 0, lbl,        fmt["data"])
        ws1.write(row, 1, len(subset), fmt["num"]); row += 1

    # Definitions legend
    row += 3
    row = section_header(ws1, row, "DEFINITIONS", fmt["section"], 4)
    defs = [
        ("In Scope",       "Users in Master Wave File whose Leaders column is in the 16-leader allowlist."),
        ("Snapshot",       "A point-in-time MVP export — the roster as of its Pulled date above."),
        ("Period",         "Change window between two consecutive snapshots, named for the month it mostly "
                           "falls in (the starting snapshot's month). Exact window always shown alongside."),
        ("Why the window",  "Exports are pulled in the first week, so Jul 5 – Aug 2 is ~26 days of July vs "
                           "2 of August — reported as July, not August."),
        ("Current month",  "Has no period until next month's export lands and closes the window. "
                           "Snapshots carry no timestamps, so windows cannot be split by calendar day."),
        ("New Addition",   f"User present in a later snapshot but absent from {labels[0]}."),
        ("Wave Change",    "GoLiveWave differs between two consecutive monthly snapshots."),
        ("Role Change",    "Any Job Role 1-4 (IndividualCategoryUpdate1-4Name) differs between consecutive snapshots."),
        ("Leaders",        "From Master Wave File col AD. VP if VP set, else AVP."),
        ("Departed",       "IsDepartedInactive = True in user's most recent snapshot. Flagged but included in counts."),
        ("Wave Exception", "HasWaveException = True in user's most recent snapshot. Surfaced on Wave Changes sheet."),
        ("(None)",         "GoLiveWave or Job Role was blank in that snapshot."),
    ]
    for term, defn in defs:
        ws1.write(row, 0, term, fmt["bold"])
        ws1.write(row, 3, defn, fmt["data"]); row += 1

    # =========================================================================
    # By Leader
    # =========================================================================
    ws2 = wb.add_worksheet("By Leader")
    ws2.set_column(0, 0, 28)
    ws2.set_column(1, 1, 30)
    ws2.set_column(2, max(5, len(period_keys) + 1), 16)
    ws2.hide_gridlines(2)

    ws2.write("A1", "Wave Movements & Job Role Changes by Leader", fmt["title"])
    row = 3

    row = section_header(ws2, row, "WAVE MOVEMENTS BY LEADER", fmt["section"], 4)
    for key in period_keys:
        change_col = f"WaveChange_{key}"
        ws2.write(row, 0, period_labels[key], fmt["subhdr"])
        ws2.write(row, 1, "Wave Transition",  fmt["subhdr"])
        ws2.write(row, 2, "Users",            fmt["subhdr"]); row += 1

        changed = in_scope[in_scope[change_col] != ""]
        if changed.empty:
            ws2.write(row, 0, "No wave changes", fmt["gray"]); row += 2; continue

        grp = (
            changed.groupby(["Leaders", change_col], sort=True)
            .size().reset_index(name="Count")
        )
        prev_leader = None
        for _, r in grp.iterrows():
            leader_cell = r["Leaders"] if r["Leaders"] != prev_leader else ""
            prev_leader = r["Leaders"]
            ws2.write(row, 0, leader_cell,     fmt["data"])
            ws2.write(row, 1, r[change_col],   fmt["arrow"])
            ws2.write(row, 2, int(r["Count"]), fmt["num"]); row += 1
        row += 2

    row += 1
    row = section_header(ws2, row, "JOB ROLE CHANGES BY LEADER",
                         fmt["section"], len(period_keys) + 1)
    ws2.write(row, 0, "Leader", fmt["header"])
    for i, key in enumerate(period_keys):
        ws2.write(row, i + 1, f"{period_month[key]} ({period_win_sh[key]})", fmt["header"])
    row += 1
    for leader in sorted(LEADER_ALLOWLIST):
        sub = in_scope[in_scope["Leaders"] == leader]
        ws2.write(row, 0, leader, fmt["data"])
        for i, key in enumerate(period_keys):
            ws2.write(row, i + 1, int(sub[f"RoleChanged_{key}"].sum()), fmt["num"])
        row += 1

    # Role transition detail by leader
    row += 2
    row = section_header(ws2, row,
        "ROLE TRANSITIONS BY LEADER — FROM → TO  (all periods combined)", fmt["section"], 4)
    for leader in sorted(LEADER_ALLOWLIST):
        sub  = in_scope[in_scope["Leaders"] == leader]
        parts = [role_transition_counts(sub, a, b) for a, b in pairs]
        leader_rt = (
            pd.concat(parts, ignore_index=True)
            .groupby(["From", "To"])["Count"].sum()
            .reset_index()
            .sort_values("Count", ascending=False)
        ) if parts else pd.DataFrame(columns=["From", "To", "Count"])

        if leader_rt.empty:
            continue

        ws2.write(row, 0, leader,   fmt["bold"])
        ws2.write(row, 1, "From",   fmt["subhdr"])
        ws2.write(row, 2, "To",     fmt["subhdr"])
        ws2.write(row, 3, "Count",  fmt["subhdr"]); row += 1
        for _, rr in leader_rt.iterrows():
            ws2.write(row, 1, rr["From"],       fmt["role_from"])
            ws2.write(row, 2, rr["To"],         fmt["role_to"])
            ws2.write(row, 3, int(rr["Count"]), fmt["num"]); row += 1
        row += 1

    # =========================================================================
    # Wave Changes by User
    # =========================================================================
    ws3 = wb.add_worksheet("Wave Changes by User")
    ws3.set_column(0, 0, 18)
    ws3.set_column(1, 1, 32)
    ws3.set_column(2, 2, 28)
    ws3.set_column(3, 3 + len(labels) - 1, 16)
    ws3.set_column(3 + len(labels), 3 + len(labels) + len(pairs) - 1, 28)
    ws3.set_column(3 + len(labels) + len(pairs), 3 + len(labels) + len(pairs) + 1, 14)
    ws3.hide_gridlines(2)

    ws3.write(0, 0, f"Wave Changes by User  ({len(wave_users):,} users with >= 1 wave change)",
              fmt["title"])

    headers3 = (
        ["Universal ID", "Team Member", "Leaders"]
        + [f"Wave ({l})" for l in labels]
        + [f"Wave Chg — {period_labels[f'{a}_{b}']}" for a, b in pairs]
        + ["Departed", "Wave Exception"]
    )
    row = 2
    for i, h in enumerate(headers3):
        ws3.write(row, i, h, fmt["header"])
    row += 1

    for uid, r in wave_users.iterrows():
        is_dep  = r.get("Departed", "") == "Yes"
        base_f  = fmt["depart"] if is_dep else fmt["data"]
        col = 0
        ws3.write(row, col, uid,              base_f); col += 1
        ws3.write(row, col, r["Team Member"], base_f); col += 1
        ws3.write(row, col, r["Leaders"],     base_f); col += 1
        for lbl in labels:
            ws3.write(row, col, r[f"Wave_{lbl}"], base_f); col += 1
        for a, b in pairs:
            val = r[f"WaveChange_{a}_{b}"]
            ws3.write(row, col, val or "—",
                      (fmt["arrow"] if (val and not is_dep) else base_f)); col += 1
        ws3.write(row, col, r.get("Departed",   ""),
                  fmt["depart"] if is_dep else fmt["data"]); col += 1
        ws3.write(row, col, r.get("WaveExcept", ""),
                  fmt["exc"] if r.get("WaveExcept", "") == "Yes" else fmt["data"]); col += 1
        row += 1

    ws3.autofilter(2, 0, row - 1, len(headers3) - 1)
    ws3.freeze_panes(3, 3)

    # =========================================================================
    # Job Role Changes
    # =========================================================================
    ws4 = wb.add_worksheet("Job Role Changes")
    n_role_cols = len(JOB_ROLE_LABELS) * len(labels)
    chg_col_start = 3 + n_role_cols
    what_col      = chg_col_start + len(pairs)
    dep_col       = what_col + 1
    ws4.set_column(0, 0, 18)
    ws4.set_column(1, 1, 32)
    ws4.set_column(2, 2, 28)
    ws4.set_column(3, 3 + n_role_cols - 1, 26)
    ws4.set_column(chg_col_start, what_col - 1, 12)
    ws4.set_column(what_col, what_col, 60)
    ws4.set_column(dep_col,  dep_col + 1, 14)
    ws4.hide_gridlines(2)

    ws4.write(0, 0,
        f"Job Role Changes  ({len(role_users):,} users with >= 1 role change, any period)",
        fmt["title"])

    headers4 = (
        ["Universal ID", "Team Member", "Leaders"]
        + [f"{lbl} ({snap})" for lbl in JOB_ROLE_LABELS for snap in labels]
        + [f"Role Chg — {period_month[f'{a}_{b}']} ({period_win_sh[f'{a}_{b}']})"
           for a, b in pairs]
        + ["What Changed", "Departed", "Wave Exception"]
    )
    row = 2
    for i, h in enumerate(headers4):
        ws4.write(row, i, h, fmt["header"])
    row += 1

    for uid, r in role_users.iterrows():
        is_dep  = r.get("Departed", "") == "Yes"
        base_f  = fmt["depart"] if is_dep else fmt["data"]
        col = 0
        ws4.write(row, col, uid,              base_f); col += 1
        ws4.write(row, col, r["Team Member"], base_f); col += 1
        ws4.write(row, col, r["Leaders"],     base_f); col += 1
        for lbl in JOB_ROLE_LABELS:
            for snap in labels:
                ws4.write(row, col, r[f"{lbl}_{snap}"], base_f); col += 1
        for a, b in pairs:
            changed = bool(r[f"RoleChanged_{a}_{b}"])
            ws4.write(row, col, "Y" if changed else "",
                      fmt["yes"] if (changed and not is_dep) else base_f); col += 1
        ws4.write(row, col, role_what_changed.get(uid, ""), fmt["what_chg"]); col += 1
        ws4.write(row, col, r.get("Departed",   ""),
                  fmt["depart"] if is_dep else fmt["data"]); col += 1
        ws4.write(row, col, r.get("WaveExcept", ""),
                  fmt["exc"] if r.get("WaveExcept", "") == "Yes" else fmt["data"]); col += 1
        row += 1

    ws4.autofilter(2, 0, row - 1, len(headers4) - 1)
    ws4.freeze_panes(3, 3)

    # =========================================================================
    # New Additions
    # =========================================================================
    ws5 = wb.add_worksheet("New Additions")
    ws5.set_column(0, 0, 18)
    ws5.set_column(1, 1, 32)
    ws5.set_column(2, 2, 28)
    ws5.set_column(3, 3, 18)
    ws5.set_column(4, 4 + len(labels) - 2, 16)
    ws5.set_column(4 + len(labels) - 1, 4 + len(labels) + 3, 26)
    ws5.set_column(4 + len(labels) + 4, 4 + len(labels) + 5, 14)
    ws5.hide_gridlines(2)

    ws5.write(0, 0, f"New Additions  ({len(new_users):,} users not in {labels[0]} snapshot)",
              fmt["title"])
    ws5.write(1, 0,
        f"Users present in a later snapshot but absent from the {labels[0]} snapshot.",
        fmt["gray"])

    headers5 = (
        ["Universal ID", "Team Member", "Leaders", "First Snapshot"]
        + [f"Wave ({l})" for l in labels[1:]]
        + [f"Job Role {i} ({labels[-1]})" for i in range(1, 5)]
        + ["Departed", "Wave Exception"]
    )
    row = 3
    for i, h in enumerate(headers5):
        ws5.write(row, i, h, fmt["header"])
    row += 1

    for uid, r in new_users.iterrows():
        first_snap = next((lbl for lbl in labels[1:] if r[f"In_{lbl}"]), labels[-1])
        is_dep  = r.get("Departed", "") == "Yes"
        base_f  = fmt["depart"] if is_dep else fmt["new"]
        col = 0
        ws5.write(row, col, uid,              base_f); col += 1
        ws5.write(row, col, r["Team Member"], base_f); col += 1
        ws5.write(row, col, r["Leaders"],     base_f); col += 1
        ws5.write(row, col, first_snap,       base_f); col += 1
        for lbl in labels[1:]:
            ws5.write(row, col, r[f"Wave_{lbl}"], base_f); col += 1
        for i in range(1, 5):
            ws5.write(row, col, r[f"Job Role {i}_{labels[-1]}"], base_f); col += 1
        ws5.write(row, col, r.get("Departed",   ""),
                  fmt["depart"] if is_dep else fmt["data"]); col += 1
        ws5.write(row, col, r.get("WaveExcept", ""),
                  fmt["exc"] if r.get("WaveExcept", "") == "Yes" else fmt["data"]); col += 1
        row += 1

    ws5.autofilter(3, 0, row - 1, len(headers5) - 1)
    ws5.freeze_panes(4, 3)

    wb.close()

    # =========================================================================
    # Console summary
    # =========================================================================
    print("\n=== REPORT COMPLETE ===")
    print(f"  Output: {OUT_PATH}")

    print("\n  Wave movements (in-scope users):")
    for key in period_keys:
        n   = int((in_scope[f"WaveChange_{key}"] != "").sum())
        tbl = wave_transition_counts(in_scope, f"WaveChange_{key}")
        print(f"    {period_labels[key].replace(chr(8594), '->')}: {n:,} users")
        for _, r in tbl.iterrows():
            print(f"      {r['Wave Transition'].replace(chr(8594), '->')}: {int(r['Count']):,}")

    print("\n  Job role changes (in-scope users):")
    for key in period_keys:
        n = int(in_scope[f"RoleChanged_{key}"].sum())
        print(f"    {period_labels[key].replace(chr(8594), '->')}: {n:,} users")

    if not all_rt.empty:
        print("\n  Top role transitions (all periods):")
        for _, rr in all_rt.head(10).iterrows():
            print(f"    {rr['From']} → {rr['To']}: {int(rr['Count']):,}")

    print(f"\n  New additions (not in {labels[0]}):")
    for i, lbl in enumerate(labels[1:], start=1):
        prior  = labels[:i]
        subset = in_scope[
            ~in_scope[[f"In_{p}" for p in prior]].any(axis=1) & in_scope[f"In_{lbl}"]
        ]
        print(f"    First in {lbl}: {len(subset):,} users")

    print(f"\n  Departed users in scope: {departed_n:,}")
    print("\n  Sheets written:")
    print("    Dashboard        — KPIs + wave dest table + 5 charts (helper data on hidden _ChartData)")
    print("    Summary          — counts + role transition table + definitions")
    print("    By Leader        — wave movements + role counts + role transition detail")
    print(f"    Wave Changes     — {len(wave_users):,} rows  (+ Departed, Wave Exception)")
    print(f"    Job Role Changes — {len(role_users):,} rows  (+ What Changed, Departed, Wave Exception)")
    print(f"    New Additions    — {len(new_users):,} rows  (+ Departed, Wave Exception)")


if __name__ == "__main__":
    main()
