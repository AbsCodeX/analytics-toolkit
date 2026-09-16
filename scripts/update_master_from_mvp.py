# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
update_master_from_mvp.py

Updates columns in the Master Wave File DATA sheet with latest values
from the MVP User Mappings CSV, matched on UniversalID (case-insensitive).

SAVE MECHANISM (changed 2026-05-18)
-----------------------------------
This script used to load/save the Master through openpyxl. openpyxl cannot
round-trip this workbook: it mangled the DATA sheet's ~7,599 shared formulas
and ~15,198 array formulas, dropped Table1's 8 calculated columns, and blew
out the table ref over styled empty rows — which is why Excel "repaired" the
file on open, the table changed, and empty rows rendered white.

It now edits the workbook surgically via scripts/surgical_xlsx.py: only the
specific value cells this logic changes are rewritten inside the DATA
worksheet XML; every other part (formulas, styles, conditional formatting,
data validation, other sheets) is preserved byte-for-byte. calcChain.xml is
dropped cleanly so Excel rebuilds it silently instead of prompting a repair.
No openpyxl, fix_table_style, or restore_power_query pass is needed anymore.

BEHAVIOR NOTES vs. the old version
----------------------------------
* Blank/gap row deletion: the maintained Master has zero blank rows. Safely
  renumber-deleting mid-sheet rows in raw XML is unsafe here (fragmented CF
  sqref ranges + shared-formula masters — the exact edit that corrupted this
  file on 2026-05-13). If blank rows are ever found, the script now WARNS and
  asks you to delete them in Excel rather than doing it itself.
* Run Log: the auto-archive-on-schema-drift behavior is dropped; the script
  appends to the existing 'Run Log' sheet and adds any new metric columns.

Rules:
  - Match is UniversalID (normalized to upper) — never by name
  - If UID found in MVP: overwrite mapped columns with MVP values;
    clear any "not in MVP" note from Notes column
  - If UID NOT in MVP: leave all data unchanged; write note to Notes column:
    "User Not in MVP User Mappings."
  - Badge Buddies updated from IsEpicSuperUser (string -> bool)
  - Full Name computed as "LastName, FirstName MI" for every row
  - Training Preference auto-filled when blank (Remote/Onsite rules)
  - Training Needed: set to "No" ONLY when departed/inactive OR EVERY populated
    Job Role 1-4 is on the no-training list (wave_reference_lists.xlsx,
    'no_training_job_roles' sheet — the shared editable source; REF!I is
    display-only fallback). One qualifying role anywhere in Job Role 1-4 keeps
    the person trainable (her rule 2026-09-10: HIM scanners with an ROI /
    Deficiency Analyst second role were being forced to No by Job Role 1 alone).
    Blank cells are filled with "Yes".
    Existing Yes/No values are never overwritten otherwise. The old "Why
    reason → No" override (LOA / Terminated / View Only) is no longer applied.
  - Why Training Not Needed: was never written by this script (per user
    request 2026-06-09); the column itself was DELETED from the Master on
    2026-07-07, so it is no longer required or referenced at all.
  - Notes: kept short — "Not in MVP" / "Dup UID" (old long forms auto-migrate).
  - Latest MVP Record Update Date?: the column was DELETED from the Master on
    2026-07-23 (per user) — the run-date stamp is no longer written; the Users
    Wave Change Log alone records per-row changes.
  - EpicTMLookup_* block (14 columns, added by user 2026-07-23): mirrored
    EXACTLY from the raw Epic Team Member Lookup export (EPIC_COL_MAP) in the
    same pass — overwrite every run; a person absent from today's export gets
    the whole block cleared to blank ("not known to Epic today", never stale).
    Fully Registered / Fully Trained normalized to clean Yes/No (emoji
    stripped); hire date as MM/DD/YYYY; Total Curriculums as a plain integer.
  - Validation Checks / Training Needed Validation Checks: the user's own
    in-cell LET/array formulas over Table1 — PROTECTED, never written.
  - Promotion from 'New Users to Add' sheet is DISABLED — only DATA, Run Log,
    and DASHBOARD!C3 are written this run (Power Query / other sheets safe).
  - HR backfill (fill-blanks-only): after the MVP pass, the CLEANED HR file's
    "All Staff" tab (HR_cleaned.xlsx, produced by clean_hr.py — where leader-
    chain Universal IDs are already resolved to full names) is loaded and used
    to fill BLANK leader cells (SeniorManager, Director,
    Sr. Director, AVP, VP, SVP, Leaders) for any matched UniversalID. Never
    overwrites a populated leader cell. Leaders is recomputed only when its
    cell is blank (VP > AVP > SVP, with No Longer Rev Cycle fallback and the
    Lastname09/Lastname02 UID overrides).
  - Identity vs MVP stub rows (2026-09-04): MVP carries stub rows for some
    Locked/ServiceNow users (no names/EmployeeID/title/department/manager).
    A BLANK MVP value never overwrites a populated identity cell
    (IDENTITY_KEEP_IF_MVP_BLANK); afterwards any row still missing
    FirstName/LastName is backfilled — blanks only — from the RAW HR export
    (HR_IDENTITY_MAP + MI) and Full Name recomputed. Raw HR is read only
    when a row needs it.
  - Excel Table1 ref + sheet dimension updated to actual row count
  - DASHBOARD C3 updated with today's run date (MM/DD/YYYY)
  - Run Log row appended inside the Master, plus external master_runlog.xlsx

Column mapping (Master -> MVP): unchanged from prior version (see COL_MAP).
"""

import csv
import datetime
import re
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).parent))
import runlog
import onedrive_paths
import master_backup
import activity_log
from surgical_xlsx import SurgicalWorkbook, num_to_col
from leader_names import (
    CANONICAL_SET,
    NOT_REV_CYCLE,
    NOTE_DUPLICATE_UID,
    NOTE_NOT_IN_MVP,
    add_note,
    normalize_leader,
    remove_note,
)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE        = Path(__file__).resolve().parent.parent
MVP_PATH    = onedrive_paths.RAW_MVP_USER_MAPPINGS
HR_PATH     = onedrive_paths.latest_hr_cleaned()  # most recent dated HR_cleaned_*.xlsx, or None
HR_SHEET    = "All Staff"  # unfiltered tab of HR_cleaned.xlsx (clean_hr.py output)
MASTER_PATH = onedrive_paths.MASTER_WAVE_PATH
LOG_PATH    = BASE / "scripts/update_master_from_mvp_log.txt"

RUNLOG_TAB  = "MVP Update"
RUNLOG_HEADERS = [
    "Run Timestamp", "MVP UIDs Loaded", "Master Rows Updated",
    "Rows NOT in MVP", "Blank Rows Deleted", "Training Needed Set to No",
    "Total Data Rows", "MVP File", "Master File", "Backup File",
    "Data Snapshot File",
    "Wave 1→2", "Wave 2→3", "Wave 3→4", "Wave Changes (Other)",
    "Total Departed/Inactive", "Total TN=Yes", "Total TN=No",
]

RUN_LOG_SHEET = "Run Log"
RUN_LOG_HEADERS = [
    "Run Timestamp",
    "Total Rows (Start)",
    "Total Rows (End)",
    "Blank Rows Deleted",
    "Rows Updated (MVP Match)",
    "Rows NOT in MVP",
    "Newly Departed/Inactive",
    "Training Needed → No (Total)",
    "  → Departed/Inactive",
    "  → Why Training Reason",
    "  → No-Training Job Role",
    "Training Pref Auto-Filled",
    "  → Remote",
    "  → Onsite",
    "Wave 1 → 2",
    "Wave 2 → 3",
    "Wave 3 → 4",
    "Wave Changes (Other)",
    "Total Departed/Inactive",
    "Total Training Needed = Yes",
    "Total Training Needed = No",
    "MVP File",
    "Duplicate Rows (Review)",
    "New Users Promoted",
    "Staging Skipped (Already on DATA)",
    "Change Log Entries Written",
    "HR Backfill: HR Sheet",
    "HR Backfill: HR UIDs Loaded",
    "HR Backfill: Rows Touched",
    "HR Backfill: SeniorManager",
    "HR Backfill: Director",
    "HR Backfill: Sr. Director",
    "HR Backfill: AVP",
    "HR Backfill: VP",
    "HR Backfill: SVP",
    "HR Backfill: Leaders",
]

WAVE_LOG_COL_NAME = "Users Wave Change Log"
TRACKED_JOB_ROLES = ("Job Role 1", "Job Role 2", "Job Role 3", "Job Role 4")
MAX_LOG_CELL_CHARS = 30000

NEW_USERS_SHEET = "New Users to Add"
NEW_USERS_HEADER_ALIASES = {
    "vendor (offshore/onshore)": "User Type (Offshore/Onshore/YourOrg)",
    "seniordirector":            "Sr. Director",
}

_OLD_JR1_FLAG = "USER NOT LISTED IN MVP"
NO_TRAINING_REASONS = {"loa", "terminated", "view only / no access"}

# HR backfill (fill-blanks-only). HR field name -> Master column name.
HR_LEAD_MAP = {
    "SeniorManager":  "SeniorManager",
    "Director":       "Director",
    "SeniorDirector": "Sr. Director",   # name differs between HR and Master
    "AVP":            "AVP",
    "VP":             "VP",
    "SVP":            "SVP",
}
# Only VP/AVP/SVP are canonicalized against the leaders list. SeniorManager,
# Director, Sr. Director can hold any employee name and must NOT run through
# last-name lookup (would resolve to the wrong person).
_HR_CANONICALIZE = {"VP", "AVP", "SVP"}
# UID-keyed Leaders overrides for VPs whose HR record routes their SVP slot
# through Lastname03's own manager (Skipmgrlast, Skipmgrfirst) instead of Lastname03.
HR_LEADER_UID_OVERRIDES = {
    "BKREBS":     "Lastname03, Firstname03",   # Lastname09, Firstname09
    "LEADER02UID": "Lastname03, Firstname03",   # Lastname02, Firstname02
}

# Master column name -> MVP CSV field name
COL_MAP = {
    "EmployeeID":                   "EmployeeID",
    "FirstName":                    "FirstName",
    "MI":                           "MI",
    "LastName":                     "LastName",
    "DepartmentLocation":           "DepartmentLocation",
    "JobTitle":                     "JobTitle",
    "BusinessUnit":                 "BusinessUnit",
    "ValidatingManagerUniversalID": "ValidatingManagerUniversalID",
    "ValidatingManagerFirstName":   "ValidatingManagerFirstName",
    "ValidatingManagerLastName":    "ValidatingManagerLastName",
    "HRManagerUniversalID":         "HRManagerUniversalID",
    "OwningApplication":            "OwningApplication",
    "Job Role 1":                   "IndividualCategoryUpdate1Name",
    "Job Role 2":                   "IndividualCategoryUpdate2Name",
    "Job Role 3":                   "IndividualCategoryUpdate3Name",
    "Job Role 4":                   "IndividualCategoryUpdate4Name",
    "Status":                       "Status",
    "GoLiveWave":                   "GoLiveWave",
    "IsDepartedInactive?":          "IsDepartedInactive",
    "Badge Buddies":                "IsEpicSuperUser",
}

# 2026-09-04: MVP User Mappings carries STUB rows for some Locked/ServiceNow
# users — lowercase UID, no names, EmployeeID, department, title or manager
# (215 such rows in the 9/4 export). Mirroring those blanks wiped the
# HR-sourced identity of 6 people add_missing_to_master.py had appended the
# same day (and VEN-JKLEIN before that). For these identity columns a BLANK
# MVP value keeps whatever the Master already holds; a populated MVP value
# still overwrites exactly as before. Roles / Status / Wave / flags remain
# mirror-exactly — MVP is their authority and a blank there is meaningful.
IDENTITY_KEEP_IF_MVP_BLANK = frozenset({
    "EmployeeID", "FirstName", "MI", "LastName", "DepartmentLocation", "JobTitle",
    "BusinessUnit", "ValidatingManagerUniversalID", "ValidatingManagerFirstName",
    "ValidatingManagerLastName", "HRManagerUniversalID",
})

# Identity backfill (fill-blanks-only) from the RAW HR export, for rows whose
# FirstName / LastName are blank after the MVP pass (MVP stub or not in MVP).
# Raw HR field -> Master column. Same fields add_missing_to_master.py seeds.
HR_IDENTITY_MAP = {
    "FirstName":          "FirstName",
    "LastName":           "LastName",
    "EmployeeID":         "EmployeeID",
    "JobTitle":           "JobTitle",
    "Department Name":    "DepartmentLocation",
    "ManagerUniversalID": "ValidatingManagerUniversalID",
}

REQUIRED_COLS = (
    "UniversalID", "EmployeeID", "FirstName", "MI", "LastName", "Full Name",
    "Notes", "Training Needed (Yes/No)",
    "IsDepartedInactive?", "Job Role 1",
    "Vendor Yes/No", "User Type (Offshore/Onshore/YourOrg)", "Training Preference",
)

# Epic TM Lookup mirror: Master column -> raw export column ("Export" sheet of
# Epic Team Member Lookup.xlsx). Export headers are resolved case- and
# whitespace-insensitively (some carry double internal spaces). A missing
# column on either side warns and is skipped — never fatal.
EPIC_UID_HEADER = "Universal ID"      # export join key (with space)
EPIC_COL_MAP = {
    "EpicTMLookup_Curriculum Type":                            "Curriculum Type",
    "EpicTMLookup_WorkerType":                                 "Worker Type",
    "EpicTMLookup_EpicTrainingNeeded(MVP)":                    "Epic Training Needed (MVP)",
    "EpicTMLookup_ActiveWorker(PDM)":                          "Active Worker (PDM)",
    "EpicTMLookup_HRStatus(PDM)":                              "Hr Status (PDM)",
    "EpicTMLookup_TeamMemberHasActiveCurriculums(Cornerstone)": "Team Member Has Active Curriculums (Cornerstone)",
    "EpicTMLookup_OffshoreVendor(PDM)":                        "Offshore Vender (PDM)",
    "EpicTMLookup_EpicTrainingEligible(Appears on Dashboard)": "Epic Training Eligible (Appears on Dashboard)",
    "EpicTMLookup_FullyRegistered":                            "Fully Registered",
    "EpicTMLookup_FullyTrained":                               "Fully Trained",
    "EpicTMLookup_TeamMemberHireDate":                         "Team Member Hire Date",
    "EpicTMLookup_TotalCurriculums(MVP)":                      "Total Curriculums (MVP)",
    "EpicTMLookup_EpicCurriculums(Cornerstone)":               "Epic Curriculums (Cornerstone)",
    "EpicTMLookup_EpicAccess":                                 "Epic Access",
}
# Epic exports these flags emoji-prefixed ("✅ Yes" / "⛔ No") — normalize to
# clean Yes/No (user decision 2026-07-23).
_EPIC_YESNO_COLS = {"EpicTMLookup_FullyRegistered", "EpicTMLookup_FullyTrained"}
_EPIC_DATE_COLS  = {"EpicTMLookup_TeamMemberHireDate"}
_EPIC_INT_COLS   = {"EpicTMLookup_TotalCurriculums(MVP)"}
_YES_NO_RE = re.compile(r"\b(yes|no)\b", re.I)

_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

def load_mvp(path: Path) -> dict:
    """Load all MVP rows (unfiltered). dict keyed by UniversalID.upper()."""
    mvp = {}
    dup_uids: list[str] = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            uid = row.get("UniversalID", "").strip().upper()
            if not uid:
                continue
            if uid in mvp:
                dup_uids.append(uid)
            mvp[uid] = row
    if dup_uids:
        sample = ", ".join(dup_uids[:5])
        more = " ..." if len(dup_uids) > 5 else ""
        print(
            f"  WARNING: MVP CSV contains {len(dup_uids)} duplicate UniversalID "
            f"row(s) (last wins). Sample: {sample}{more}"
        )
    return mvp


def clean_val(raw) -> str | None:
    val = _ILLEGAL_XML_RE.sub("", str(raw)).strip()
    if len(val) > 32000:
        val = val[:32000] + "... [TRUNCATED BY SCRIPT]"
    return val if val else None


def convert_bool(raw: str):
    s = raw.strip().lower()
    if s in {"true", "1", "yes", "y", "t"}:
        return True
    if s in {"false", "0", "no", "n", "f"}:
        return False
    return None


def _norm_cell(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def append_change_log(existing, new_entries: list[str], ts_str: str) -> str:
    base = (str(existing) if existing else "").rstrip()
    new_block = "\n".join(f"{ts_str} - {e}" for e in new_entries)
    combined = f"{base}\n{new_block}" if base else new_block
    if len(combined) > MAX_LOG_CELL_CHARS:
        tail = combined[-MAX_LOG_CELL_CHARS:]
        nl = tail.find("\n")
        if nl >= 0:
            tail = tail[nl + 1:]
        combined = "[... older entries truncated ...]\n" + tail
    return combined


def _norm_epic_val(mcol: str, raw) -> str | None:
    """Normalize one raw Epic lookup value for its Master column (or None)."""
    if raw is None:
        return None
    if isinstance(raw, float) and raw != raw:      # NaN without needing pandas
        return None
    if mcol in _EPIC_DATE_COLS:
        if isinstance(raw, (datetime.datetime, datetime.date)):
            return raw.strftime("%m/%d/%Y")
        return clean_val(raw)
    if mcol in _EPIC_INT_COLS:
        try:
            return str(int(float(raw)))
        except (TypeError, ValueError):
            return clean_val(raw)
    if mcol in _EPIC_YESNO_COLS:
        m = _YES_NO_RE.search(str(raw))
        return m.group(1).title() if m else clean_val(raw)
    return clean_val(raw)


def load_epic_lookup() -> tuple[dict | None, str]:
    """
    Raw Epic Team Member Lookup export -> ({UID: {master_col: value}}, filename).

    Values arrive already normalized per _norm_epic_val, keyed by MASTER column
    name so the row loop just writes them. Returns (None, reason) when the file
    or its UID column is unavailable — the mirror pass is then SKIPPED entirely
    (fail-soft: an unchanged block beats a wrongly blanked one).
    """
    path = onedrive_paths.RAW_EPIC_TM_DIR / "Epic Team Member Lookup.xlsx"
    if not path.exists():
        cands = sorted(onedrive_paths.RAW_EPIC_TM_DIR.glob("*.xlsx"),
                       key=lambda f: f.stat().st_mtime, reverse=True)
        if not cands:
            return None, f"no xlsx found in {onedrive_paths.RAW_EPIC_TM_DIR}"
        path = cands[0]

    import pandas as pd
    df = pd.read_excel(path, sheet_name="Export", dtype=object, engine="calamine")

    def _hnorm(h) -> str:
        return re.sub(r"\s+", " ", str(h)).strip().lower()

    by_norm = {_hnorm(c): c for c in df.columns}
    uid_actual = by_norm.get(_hnorm(EPIC_UID_HEADER))
    if uid_actual is None:
        return None, f"'{EPIC_UID_HEADER}' column missing from {path.name}"

    resolved: dict[str, str] = {}          # master col -> actual export header
    for mcol, ecol in EPIC_COL_MAP.items():
        actual = by_norm.get(_hnorm(ecol))
        if actual is None:
            print(f"  WARNING: Epic lookup export missing column '{ecol}' — "
                  f"'{mcol}' will not be updated this run.")
        else:
            resolved[mcol] = actual

    data: dict[str, dict] = {}
    for rec in df[[uid_actual, *resolved.values()]].itertuples(index=False):
        uid = str(rec[0] or "").strip().upper()
        if not uid or uid == "NAN" or uid in data:   # keep first per UID
            continue
        data[uid] = {mcol: _norm_epic_val(mcol, rec[i + 1])
                     for i, mcol in enumerate(resolved)}
    return data, path.name


def load_hr(path: Path) -> tuple[dict, str]:
    """
    Load HR file, dedupe on UniversalID (keep first). Returns (dict, sheet_name).
    Dict is keyed by UID.upper() -> {header: value, ...}.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet_name = HR_SHEET if HR_SHEET in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    rows = ws.iter_rows(values_only=True)
    try:
        headers = [str(h).strip() if h is not None else "" for h in next(rows)]
    except StopIteration:
        wb.close()
        raise ValueError(f"HR export at {path} contains no rows.")
    col = {h: i for i, h in enumerate(headers) if h}
    if "UniversalID" not in col:
        wb.close()
        raise KeyError("HR sheet is missing required 'UniversalID' column.")
    uid_i = col["UniversalID"]
    hr: dict[str, dict] = {}
    for row in rows:
        if uid_i >= len(row):
            continue
        raw = row[uid_i]
        if not raw:
            continue
        s = str(raw).strip()
        if s.endswith(".0"):
            s = s[:-2]
        uid = s.upper()
        if not uid or uid in hr:
            continue
        hr[uid] = {headers[i]: (row[i] if i < len(row) else None)
                   for i in range(len(headers))}
    wb.close()
    return hr, sheet_name


def clean_hr_leadership(raw, canonicalize: bool = False) -> str | None:
    """Clean an HR leadership value. 'nan'/'nan, nan' treated as null."""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan, nan", "nan"):
        return None
    return normalize_leader(s) if canonicalize else s


def compute_hr_leaders(vp: str | None, avp: str | None, svp: str | None) -> str | None:
    """VP > AVP > SVP fallback; non-canonical leader -> 'No Longer Rev Cycle'."""
    raw = vp or avp or svp
    if raw is None:
        return None
    return raw if raw in CANONICAL_SET else NOT_REV_CYCLE


def _is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def backfill_blanks_from_hr(wb: SurgicalWorkbook, hr: dict) -> dict:
    """
    Fill BLANK leader cells in Master with values from HR (matched on UID).
    Never overwrites a populated cell. Leaders is recomputed only when blank.
    Returns counters dict.
    """
    counts = {
        "rows_touched":  0,
        "SeniorManager": 0,
        "Director":      0,
        "Sr. Director":  0,
        "AVP":           0,
        "VP":            0,
        "SVP":           0,
        "Leaders":       0,
    }

    uid_col = wb.col_of["UniversalID"]
    leaders_col = wb.col_of.get("Leaders")

    # Resolve column letters once
    col_letters: dict[str, tuple[str, str]] = {}  # hr_field -> (master_col, letter)
    for hr_field, master_col in HR_LEAD_MAP.items():
        c = wb.col_of.get(master_col)
        if c:
            col_letters[hr_field] = (master_col, c)
    vp_col  = wb.col_of.get("VP")
    avp_col = wb.col_of.get("AVP")
    svp_col = wb.col_of.get("SVP")

    for r in wb.rows:
        rn = r[0]
        if rn == 1:
            continue
        uid_raw = wb.get_value(rn, uid_col)
        if not uid_raw:
            continue
        s = str(uid_raw).strip()
        if s.endswith(".0"):
            s = s[:-2]
        uid = s.upper()
        hr_row = hr.get(uid)
        if not hr_row:
            continue

        row_changed = False
        for hr_field, (master_col, letter) in col_letters.items():
            if not _is_blank(wb.get_value(rn, letter)):
                continue
            new_val = clean_hr_leadership(
                hr_row.get(hr_field),
                canonicalize=(hr_field in _HR_CANONICALIZE),
            )
            if new_val is None:
                continue
            wb.set_value(rn, letter, new_val)
            counts[master_col] += 1
            row_changed = True

        # Recompute Leaders only if its own cell is blank
        if leaders_col and _is_blank(wb.get_value(rn, leaders_col)):
            def _read(letter):
                v = wb.get_value(rn, letter) if letter else None
                return clean_hr_leadership(v, canonicalize=True) if v is not None else None
            vp_val  = _read(vp_col)
            avp_val = _read(avp_col)
            svp_val = _read(svp_col)
            new_leaders = compute_hr_leaders(vp_val, avp_val, svp_val)
            if uid in HR_LEADER_UID_OVERRIDES:
                new_leaders = HR_LEADER_UID_OVERRIDES[uid]
            if new_leaders:
                wb.set_value(rn, leaders_col, new_leaders)
                counts["Leaders"] += 1
                row_changed = True

        if row_changed:
            counts["rows_touched"] += 1

    return counts


def _hr_text(v) -> str | None:
    """Raw-HR cell -> stripped text, or None for blank / NaN / 'nan'."""
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "nan, nan"):
        return None
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def rows_missing_identity(wb: SurgicalWorkbook) -> list[tuple[int, str]]:
    """(row, UID) for every data row with a UID but a blank FirstName or LastName."""
    uid_col, f_col, l_col = wb.col_of["UniversalID"], wb.col_of["FirstName"], wb.col_of["LastName"]
    out = []
    for r in wb.rows:
        rn = r[0]
        if rn == 1:
            continue
        uid = _hr_text(wb.get_value(rn, uid_col))
        if not uid:
            continue
        if _is_blank(wb.get_value(rn, f_col)) or _is_blank(wb.get_value(rn, l_col)):
            out.append((rn, uid.upper()))
    return out


def load_hr_raw_identity(path: Path, uids: set[str]) -> dict:
    """Raw HR.xlsx rows for the given UIDs only, keyed by UID.upper().

    Loaded lazily (pandas + calamine, ~15 s) and only when some row needs it —
    the file is 146k rows x 67 cols, so it is not read on a clean day.
    """
    import pandas as pd  # noqa: PLC0415 — deliberate lazy import (see above)
    with pd.ExcelFile(path, engine="calamine") as xl:
        df = pd.read_excel(xl, sheet_name=xl.sheet_names[0], dtype=str)
    if "UniversalID" not in df.columns:
        raise KeyError("raw HR sheet is missing 'UniversalID'")
    df["_UID"] = df["UniversalID"].map(_hr_text).str.upper()
    df = df[df["_UID"].isin(uids)].drop_duplicates(subset=["_UID"])
    return {row["_UID"]: row for _, row in df.iterrows()}


def backfill_identity_from_raw_hr(wb: SurgicalWorkbook, needing: list[tuple[int, str]],
                                  hr_raw: dict) -> dict:
    """
    Fill BLANK identity cells (HR_IDENTITY_MAP + MI from MiddleName) from raw
    HR for the rows in `needing`, then recompute Full Name where it is blank.
    Never overwrites a populated cell. Returns counters dict.
    """
    counts = {"rows_needing": len(needing), "rows_in_hr": 0, "rows_touched": 0, "cells": 0}
    fn_col, mi_col = wb.col_of["Full Name"], wb.col_of["MI"]
    f_col, l_col = wb.col_of["FirstName"], wb.col_of["LastName"]
    for rn, uid in needing:
        h = hr_raw.get(uid)
        if h is None:
            continue
        counts["rows_in_hr"] += 1
        changed = False
        for hr_field, master_col in HR_IDENTITY_MAP.items():
            col = wb.col_of.get(master_col)
            if col is None or not _is_blank(wb.get_value(rn, col)):
                continue
            v = _hr_text(h.get(hr_field))
            if v is None:
                continue
            if master_col == "EmployeeID" and v.isdigit():
                v = int(v)
            wb.set_value(rn, col, v)
            counts["cells"] += 1
            changed = True
        mid = _hr_text(h.get("MiddleName"))
        if mid and _is_blank(wb.get_value(rn, mi_col)):
            wb.set_value(rn, mi_col, mid[0] + ".")
            counts["cells"] += 1
            changed = True
        if _is_blank(wb.get_value(rn, fn_col)):
            first = _hr_text(wb.get_value(rn, f_col)) or ""
            last = _hr_text(wb.get_value(rn, l_col)) or ""
            mi = _hr_text(wb.get_value(rn, mi_col)) or ""
            if first or last:
                given = f"{first} {mi}".strip() if mi else first
                wb.set_value(rn, fn_col, f"{last}, {given}" if last else given)
                counts["cells"] += 1
                changed = True
        if changed:
            counts["rows_touched"] += 1
    return counts


def load_no_training_roles(wb: SurgicalWorkbook) -> set:
    """No-training Job Roles (lowercased) — read from the reference workbook
    (data/references/wave_reference_lists.xlsx, 'no_training_job_roles' sheet),
    the single editable source of truth shared with the SQL warehouse
    (changed 2026-07-09; previously read the Master's REF sheet column I,
    which was position-sensitive and had drifted). Falls back to REF!I
    (rows 3+, label rows filtered) if the workbook can't be read."""
    try:
        import openpyxl
        ref_wb = openpyxl.load_workbook(
            onedrive_paths.REF_LISTS_PATH, read_only=True, data_only=True)
        ws = ref_wb["no_training_job_roles"]
        roles = {
            str(row[0]).strip().lower()
            for row in ws.iter_rows(min_row=2, max_col=1, values_only=True)
            if row[0] is not None and str(row[0]).strip()
        }
        ref_wb.close()
        if roles:
            return roles
        print(f"  WARNING: 'no_training_job_roles' sheet in "
              f"{onedrive_paths.REF_LISTS_PATH.name} is empty — falling back to REF!I.")
    except Exception as e:  # missing file/sheet — Master's REF copy still works
        print(f"  WARNING: couldn't read {onedrive_paths.REF_LISTS_PATH.name} "
              f"({e}) — falling back to the Master's REF sheet column I.")
    label_junk = {"job roles_no_training_need", "no training needed job roles"}
    return {
        str(v).strip().lower()
        for v in wb.read_column("REF", "I", min_row=3)
        if v is not None and str(v).strip()
        and str(v).strip().lower() not in label_junk
    }


def promote_new_users(wb: SurgicalWorkbook) -> dict:
    """
    Promote rows from 'New Users to Add' into DATA, then remove the processed
    staging rows. UIDs already on DATA are skipped (and cleared from staging).
    """
    counters = {"promoted": 0, "skipped_existing": 0, "blank_skipped": 0}

    if not wb.has_sheet(NEW_USERS_SHEET):
        print(f"  '{NEW_USERS_SHEET}' sheet not found — skipping promotion.")
        return counters

    stage_hdr, stage_rows = wb.read_simple_sheet(NEW_USERS_SHEET)
    if not stage_rows:
        print(f"  '{NEW_USERS_SHEET}' is empty — nothing to promote.")
        return counters
    if "UniversalID" not in stage_hdr:
        print(
            f"  WARNING: '{NEW_USERS_SHEET}' has no UniversalID column "
            "— skipping promotion."
        )
        return counters

    uid_col = wb.col_of["UniversalID"]
    existing_uids = {
        str(wb.get_value(r[0], uid_col) or "").strip().upper()
        for r in wb.rows
        if r[0] != 1 and wb.get_value(r[0], uid_col)
    }

    data_ci = {name.lower(): col for name, col in wb.col_of.items()}

    def resolve(stage_header: str):
        key = stage_header.strip().lower()
        if key in NEW_USERS_HEADER_ALIASES:
            return wb.col_of.get(NEW_USERS_HEADER_ALIASES[key])
        return data_ci.get(key)

    stage_uid_col = stage_hdr["UniversalID"]
    to_delete: set[int] = set()

    for rn, cells in stage_rows:
        uid_norm = str(cells.get(stage_uid_col) or "").strip().upper()

        if not uid_norm:
            if all(
                not str(v or "").strip() for v in cells.values()
            ):
                to_delete.add(rn)
                counters["blank_skipped"] += 1
            continue

        if uid_norm in existing_uids:
            counters["skipped_existing"] += 1
            to_delete.add(rn)
            continue

        new_vals = {uid_col: uid_norm}
        for sname, scol in stage_hdr.items():
            if sname.lower() == "universalid":
                continue
            dcol = resolve(sname)
            if not dcol:
                continue
            val = cells.get(scol)
            if val is None or (isinstance(val, str) and not val.strip()):
                continue
            new_vals[dcol] = val
        wb.append_row(new_vals)
        existing_uids.add(uid_norm)
        counters["promoted"] += 1
        to_delete.add(rn)

    if to_delete:
        wb.delete_simple_rows(NEW_USERS_SHEET, to_delete)
    return counters


def fill_training_preference(wb: SurgicalWorkbook) -> tuple[int, int]:
    uid_col      = wb.col_of["UniversalID"]
    vendor_col   = wb.col_of.get("Vendor Yes/No")
    usertype_col = wb.col_of.get("User Type (Offshore/Onshore/YourOrg)")
    pref_col     = wb.col_of.get("Training Preference")

    if not all([vendor_col, usertype_col, pref_col]):
        print("  WARNING: Training Preference columns not found — skipping fill.")
        return 0, 0

    remote_count = onsite_count = 0
    for r in wb.rows:
        rn = r[0]
        if rn == 1 or not wb.get_value(rn, uid_col):
            continue
        if wb.get_value(rn, pref_col):
            continue

        vendor    = str(wb.get_value(rn, vendor_col) or "").strip()
        user_type = str(wb.get_value(rn, usertype_col) or "").strip()

        vendor_yes    = vendor.lower() in {"yes", "y", "true", "1"}
        user_type_key = user_type.lower()
        is_remote     = user_type_key in {"offshore", "onshore", "on shore", "off shore"}
        is_onsite     = user_type_key == "yourorg"

        if vendor_yes or is_remote:
            wb.set_value(rn, pref_col, "Remote")
            remote_count += 1
        elif is_onsite:
            wb.set_value(rn, pref_col, "Onsite")
            onsite_count += 1

    return remote_count, onsite_count


def append_external_runlog(mvp_uid_count, updated, not_in_mvp, blank_deleted,
                           tn_set_no, total_rows,
                           w1_to_w2, w2_to_w3, w3_to_w4, w_other_moves,
                           total_departed_end, tn_yes_total, tn_no_total,
                           backup_path, data_snapshot_path):
    runlog.append(RUNLOG_TAB, RUNLOG_HEADERS, [
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        mvp_uid_count, updated, not_in_mvp, blank_deleted, tn_set_no,
        total_rows, str(MVP_PATH), str(MASTER_PATH), str(backup_path),
        str(data_snapshot_path),
        w1_to_w2, w2_to_w3, w3_to_w4, w_other_moves,
        total_departed_end, tn_yes_total, tn_no_total,
    ])


def _hkey(s: str) -> str:
    """Whitespace-insensitive header key. The legacy Run Log sheet stored the
    nested metrics without the cosmetic 2-space indent in RUN_LOG_HEADERS
    ('→ Remote' vs '  → Remote'); matching on this key prevents spawning
    duplicate columns and leaving the originals blank."""
    return " ".join(str(s).split()).lower()


def write_run_log_row(wb: SurgicalWorkbook, values: list) -> None:
    """
    Append the run's metrics to the in-workbook 'Run Log' sheet.

    Self-healing: the sheet is rewritten into the canonical RUN_LOG_HEADERS
    layout every run, remapping any existing rows by whitespace-insensitive
    header name. This both fixes the historical column drift (duplicate
    AA–AE columns / blank I,J,K,M,N from the indentation mismatch) and keeps
    new rows correctly aligned. Run Log has no formulas/CF/validation, so a
    full rewrite of it is safe.
    """
    if not wb.has_sheet(RUN_LOG_SHEET):
        print(f"  WARNING: '{RUN_LOG_SHEET}' sheet not in workbook — "
              "in-workbook run log skipped (external runlog still written).")
        return

    header, existing = wb.read_simple_sheet(RUN_LOG_SHEET)
    # canonical column letter per header (A, B, C, …)
    canon_col = {h: num_to_col(i + 1) for i, h in enumerate(RUN_LOG_HEADERS)}
    key_to_canon = {_hkey(h): canon_col[h] for h in RUN_LOG_HEADERS}
    # existing sheet column-letter -> its header key
    col_to_key = {col: _hkey(name) for name, col in header.items()}

    data_rows: list[dict[str, object]] = []
    for _rn, cells in existing:
        remapped: dict[str, object] = {}
        for col, val in cells.items():
            tgt = key_to_canon.get(col_to_key.get(col, ""))
            if tgt is not None:
                remapped[tgt] = val
        if remapped:
            data_rows.append(remapped)

    new_row = {canon_col[h]: v for h, v in zip(RUN_LOG_HEADERS, values)}
    data_rows.append(new_row)

    wb.rewrite_simple_sheet(RUN_LOG_SHEET, RUN_LOG_HEADERS, data_rows)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Console summary uses → / arrows; Windows consoles default to cp1252 and
    # would crash on them. Force UTF-8 (replace on the off chance it can't).
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    ts = datetime.datetime.now()

    try:
        with open(MASTER_PATH, "r+b"):
            pass
    except (PermissionError, OSError):
        print(f"ERROR: {MASTER_PATH.name} is locked (likely open in Excel).")
        print("Close the file in Excel and re-run this script.")
        sys.exit(1)

    backup_path = master_backup.ensure_daily_full_backup(MASTER_PATH)

    print("Loading MVP User Mappings...")
    mvp = load_mvp(MVP_PATH)
    print(f"  {len(mvp):,} unique UIDs in MVP")

    print("Loading Epic Team Member Lookup (EpicTMLookup_ mirror)...")
    epic_rows, epic_src = load_epic_lookup()
    if epic_rows is None:
        print(f"  WARNING: Epic mirror skipped — {epic_src}")
    elif len(epic_rows) < 1000:
        # Mirror semantics blank absent people — a truncated/partial export
        # would wrongly clear most of the block. Refuse implausibly small files.
        print(f"  WARNING: Epic mirror skipped — {epic_src} has only "
              f"{len(epic_rows):,} UIDs (expected ~100k); export looks partial.")
        epic_rows, epic_src = None, epic_src
    else:
        print(f"  {len(epic_rows):,} unique UIDs in {epic_src}")

    print("Loading Master Wave File (surgical, no openpyxl round-trip)...")
    wb = SurgicalWorkbook(MASTER_PATH)

    # Header integrity: reject duplicate / empty headers before any write.
    seen: dict[str, str] = {}
    for col, val in wb.iter_header_cells():
        if not isinstance(val, str):
            continue
        key = val.strip()
        if not key:
            raise ValueError(f"DATA sheet has an empty header in column {col}.")
        if key in seen:
            raise ValueError(
                f"DATA sheet has duplicate header '{key}' "
                f"(columns {seen[key]} and {col})."
            )
        seen[key] = col

    missing = [c for c in REQUIRED_COLS if c not in wb.col_of]
    if missing:
        raise KeyError(
            f"Master DATA sheet is missing required column(s): {missing}\n"
            "Add the header(s) to row 1 before running this script."
        )

    # -- Protected columns — never write validation columns -----------------
    # 'User Type Validation' / 'Training Needed Validation': per user request
    # 2026-06-26; both DELETED from the Master 2026-07-07 (guard self-
    # reactivates if they return). 'Validation Checks' / 'Training Needed
    # Validation Checks' (added 2026-07-23): the user's own in-cell LET/array
    # formulas over Table1 — scripts must never touch them.
    PROTECTED_COLS = (
        "User Type Validation", "Training Needed Validation",
        "Validation Checks", "Training Needed Validation Checks",
    )
    _protected_letters = {wb.col_of[n] for n in PROTECTED_COLS if n in wb.col_of}
    _orig_set_value = wb.set_value
    _protected_warned = [False]

    def _guarded_set_value(row, col, value, shared=False):
        if col in _protected_letters:
            if not _protected_warned[0]:
                print(f"  NOTE: writes to protected validation column(s) "
                      f"{sorted(_protected_letters)} are blocked — left untouched.")
                _protected_warned[0] = True
            return False
        return _orig_set_value(row, col, value, shared=shared)

    wb.set_value = _guarded_set_value

    uid_col   = wb.col_of["UniversalID"]
    notes_col = wb.col_of["Notes"]
    tn_col    = wb.col_of["Training Needed (Yes/No)"]
    fn_col    = wb.col_of["Full Name"]
    jr1_col   = wb.col_of["Job Role 1"]
    jr_cols   = [wb.col_of[h] for h in TRACKED_JOB_ROLES if h in wb.col_of]
    dep_col   = wb.col_of["IsDepartedInactive?"]
    fname_col = wb.col_of["FirstName"]
    mi_col    = wb.col_of["MI"]
    lname_col = wb.col_of["LastName"]
    wave_col  = wb.col_of.get("GoLiveWave")
    wave_log_col = wb.col_of.get(WAVE_LOG_COL_NAME)

    # Epic mirror targets: only columns present on BOTH sides get written.
    epic_cols: dict[str, str] = {}
    if epic_rows is not None:
        sample = next(iter(epic_rows.values()), {})
        for mcol in EPIC_COL_MAP:
            letter = wb.col_of.get(mcol)
            if letter is None:
                print(f"  WARNING: Master column '{mcol}' not found — "
                      "skipped by Epic mirror.")
            elif mcol in sample:
                epic_cols[mcol] = letter

    total_rows_start = sum(1 for r in wb.rows if r[0] != 1)

    # -- STEP 1: Blank/gap rows — detect only (safe-by-design) ---------------
    print("Checking for blank/gap rows...")
    blanks = wb.count_blank_rows(uid_col)
    blank_deleted = 0
    if blanks:
        print(
            f"  WARNING: {len(blanks)} blank row(s) found (rows "
            f"{blanks[:10]}{'...' if len(blanks) > 10 else ''}). "
            "Not deleted automatically — mid-sheet XML row deletion is unsafe "
            "for this file's conditional formatting / shared formulas. "
            "Please delete these rows in Excel and re-run."
        )
    else:
        print("  No blank rows found.")

    # -- STEP 1b: Promote rows from 'New Users to Add' → DATA ---------------
    # DISABLED per user request: only DATA, Run Log, and DASHBOARD!C3 should be
    # modified this run. To re-enable, uncomment the call to promote_new_users.
    promote_counts = {"promoted": 0, "skipped_existing": 0, "blank_skipped": 0}
    print(f"Skipping '{NEW_USERS_SHEET}' promotion (DATA + Run Log only mode).")

    # -- Detect duplicate UIDs ----------------------------------------------
    print("Checking for duplicate UIDs...")
    uid_to_rows: dict[str, list[int]] = {}
    for r in wb.rows:
        rn = r[0]
        if rn == 1:
            continue
        v = str(wb.get_value(rn, uid_col) or "").strip().upper()
        if v:
            uid_to_rows.setdefault(v, []).append(rn)
    duplicate_uid_to_rows = {u: rs for u, rs in uid_to_rows.items() if len(rs) > 1}
    duplicate_row_set     = {r for rs in duplicate_uid_to_rows.values() for r in rs}
    duplicate_rows_count  = len(duplicate_row_set)
    duplicate_uids_count  = len(duplicate_uid_to_rows)
    if duplicate_uids_count:
        sample = sorted(duplicate_uid_to_rows.keys())[:5]
        more = " ..." if duplicate_uids_count > 5 else ""
        print(f"  Found {duplicate_rows_count} duplicate row(s) across "
              f"{duplicate_uids_count} UID(s). Sample: {', '.join(sample)}{more}")
    else:
        print("  No duplicate UIDs.")

    no_training_roles = load_no_training_roles(wb)
    print(f"  {len(no_training_roles)} no-training job roles loaded from REF sheet")

    # -- STEP 2: MVP update --------------------------------------------------
    updated = not_in_mvp = tn_set_no = tn_departed = tn_why_reason = 0
    tn_job_role = newly_departed = 0
    w1_to_w2 = w2_to_w3 = w3_to_w4 = w_other_moves = 0
    tn_yes_total = tn_no_total = total_departed_end = 0
    change_log_entries_written = 0
    epic_matched = epic_not_in_lookup = epic_cells_changed = 0

    if not wave_log_col:
        print(f"  WARNING: '{WAVE_LOG_COL_NAME}' column not found — "
              "per-row change log skipped this run.")

    ts_short = ts.strftime("%m/%d/%y")
    row_nums = [r[0] for r in wb.rows if r[0] != 1]
    print(f"  {len(row_nums):,} data rows to process...")

    for rn in row_nums:
        uid_raw = wb.get_value(rn, uid_col)
        if not uid_raw:
            continue
        uid = str(uid_raw).strip().upper()
        mvp_row = mvp.get(uid)

        old_wave = (str(wb.get_value(rn, wave_col) or "").strip().title()
                    if wave_col else "")
        old_dep = wb.get_value(rn, dep_col)
        was_departed = (old_dep is True) or (
            convert_bool(str(old_dep)) is True if old_dep is not None else False
        )

        old_jr_vals = {}
        if mvp_row and wave_log_col:
            for jr_name in TRACKED_JOB_ROLES:
                c = wb.col_of.get(jr_name)
                if c:
                    old_jr_vals[jr_name] = _norm_cell(wb.get_value(rn, c))

        if mvp_row:
            for master_col, mvp_field in COL_MAP.items():
                col = wb.col_of.get(master_col)
                if col is None:
                    continue
                raw = mvp_row.get(mvp_field, "")
                if master_col in ("IsDepartedInactive?", "Badge Buddies"):
                    val = False if not raw.strip() else convert_bool(raw)
                elif master_col == "GoLiveWave":
                    cleaned = clean_val(raw)
                    val = cleaned.title() if cleaned else None
                elif master_col == "EmployeeID":
                    s = _ILLEGAL_XML_RE.sub("", raw).strip()
                    if not s:
                        val = None
                    else:
                        try:
                            val = int(float(s))
                        except ValueError:
                            val = clean_val(s)
                else:
                    val = clean_val(raw)
                if (val is None and master_col in IDENTITY_KEEP_IF_MVP_BLANK
                        and not _is_blank(wb.get_value(rn, col))):
                    continue  # MVP stub row: keep the populated identity cell
                wb.set_value(rn, col, val)

            existing_note = wb.get_value(rn, notes_col)
            cleared = remove_note(existing_note, NOTE_NOT_IN_MVP)
            if cleared != existing_note:
                wb.set_value(rn, notes_col, cleared)
            updated += 1

            if wave_col and old_wave:
                new_wave = str(wb.get_value(rn, wave_col) or "").strip()
                if new_wave and old_wave != new_wave:
                    if   (old_wave, new_wave) == ("Wave 1", "Wave 2"): w1_to_w2 += 1
                    elif (old_wave, new_wave) == ("Wave 2", "Wave 3"): w2_to_w3 += 1
                    elif (old_wave, new_wave) == ("Wave 3", "Wave 4"): w3_to_w4 += 1
                    else:                                               w_other_moves += 1
        else:
            jr1_val = wb.get_value(rn, jr1_col)
            if str(jr1_val or "").strip() == _OLD_JR1_FLAG:
                wb.set_value(rn, jr1_col, None)
            existing_note = wb.get_value(rn, notes_col)
            new_note = add_note(existing_note, NOTE_NOT_IN_MVP)
            if new_note != existing_note:
                wb.set_value(rn, notes_col, new_note)
            not_in_mvp += 1

        # Full Name
        first = str(wb.get_value(rn, fname_col) or "").strip()
        mi    = str(wb.get_value(rn, mi_col) or "").strip()
        last  = str(wb.get_value(rn, lname_col) or "").strip()
        if first or last:
            given = f"{first} {mi}".strip() if mi else first
            new_fn = f"{last}, {given}" if last else given
        else:
            new_fn = None
        wb.set_value(rn, fn_col, new_fn)

        # Training Needed — set "No" only for departed/inactive, or when EVERY
        # populated Job Role 1-4 is on the no-training list (a single trainable
        # role anywhere keeps the person trainable — 2026-09-10). Otherwise
        # fill blanks with "Yes"; never overwrite an existing Yes/No value.
        departed_raw = wb.get_value(rn, dep_col)
        is_departed  = (departed_raw is True) or (
            convert_bool(str(departed_raw)) is True
            if departed_raw is not None else False
        )
        jr_vals = [str(wb.get_value(rn, c) or "").strip().lower() for c in jr_cols]
        jr_vals = [v for v in jr_vals if v]
        all_roles_no_training = bool(jr_vals) and all(v in no_training_roles for v in jr_vals)
        current_tn = str(wb.get_value(rn, tn_col) or "").strip()

        if is_departed and not was_departed:
            newly_departed += 1

        # "Why Training Not Needed" is intentionally never written by this
        # script (per user request 2026-06-09) — leave it exactly as-is.
        if is_departed:
            wb.set_value(rn, tn_col, "No")
            tn_set_no += 1
            tn_departed += 1
        elif all_roles_no_training:
            wb.set_value(rn, tn_col, "No")
            tn_set_no += 1
            tn_job_role += 1
        elif not current_tn:
            wb.set_value(rn, tn_col, "Yes")
        # else: existing Yes/No left untouched

        if is_departed:
            total_departed_end += 1
        tn_final = wb.get_value(rn, tn_col)
        if tn_final == "Yes":
            tn_yes_total += 1
        elif tn_final == "No":
            tn_no_total += 1

        # Per-row change log — fires on Job Role 1-4 changes, wave moves, and
        # newly-marked Departed/Inactive.
        if mvp_row and wave_log_col:
            changes: list[str] = []
            for jr_name in TRACKED_JOB_ROLES:
                c = wb.col_of.get(jr_name)
                if not c:
                    continue
                new_val = _norm_cell(wb.get_value(rn, c))
                old_val = old_jr_vals.get(jr_name, "")
                if old_val == new_val:
                    continue
                if not old_val:
                    changes.append(f"{jr_name} added")
                elif not new_val:
                    changes.append(f"{jr_name} removed")
                else:
                    changes.append(f"{jr_name}: {old_val} → {new_val}")

            if wave_col and old_wave:
                new_wave_now = _norm_cell(wb.get_value(rn, wave_col))
                if new_wave_now and old_wave != new_wave_now:
                    new_tail = (new_wave_now[5:].strip()
                                if new_wave_now.lower().startswith("wave ")
                                else new_wave_now)
                    changes.append(f"moved from {old_wave} - {new_tail}")

            if is_departed and not was_departed:
                changes.append("marked Departed/Inactive")

            if changes:
                existing_log = wb.get_value(rn, wave_log_col)
                wb.set_value(rn, wave_log_col,
                             append_change_log(existing_log, changes, ts_short))
                change_log_entries_written += len(changes)

        # -- Epic TM Lookup mirror (EpicTMLookup_* block) -------------------
        # Mirror exactly: overwrite every run; not in today's export -> the
        # whole block cleared to blank. set_value no-ops unchanged cells.
        if epic_cols:
            evals = epic_rows.get(uid)
            if evals is not None:
                epic_matched += 1
            else:
                epic_not_in_lookup += 1
            for mcol, letter in epic_cols.items():
                target = evals.get(mcol) if evals is not None else None
                if wb.set_value(rn, letter, target, shared=True):
                    epic_cells_changed += 1

    # -- STEP 3: Fill blank Training Preference -----------------------------
    print("Filling blank Training Preference cells...")
    pref_remote, pref_onsite = fill_training_preference(wb)
    pref_filled = pref_remote + pref_onsite
    print(f"  Filled {pref_filled} cell(s): {pref_remote} Remote, "
          f"{pref_onsite} Onsite.")

    # -- STEP 3b: Sync duplicate-UID note -----------------------------------
    for r in wb.rows:
        rn = r[0]
        if rn == 1:
            continue
        existing = wb.get_value(rn, notes_col)
        if rn in duplicate_row_set:
            new_val = add_note(existing, NOTE_DUPLICATE_UID)
        else:
            new_val = remove_note(existing, NOTE_DUPLICATE_UID)
        if new_val != existing:
            wb.set_value(rn, notes_col, new_val)

    # -- STEP 3c: HR backfill — fill-blanks-only ---------------------------
    # Fill BLANK leader cells (SeniorManager, Director, Sr. Director, AVP, VP,
    # SVP, Leaders) from the latest HR file. Never overwrites populated values.
    hr_counts = {
        "rows_touched": 0, "SeniorManager": 0, "Director": 0,
        "Sr. Director": 0, "AVP": 0, "VP": 0, "SVP": 0, "Leaders": 0,
    }
    hr_sheet_name = ""
    hr_uids_loaded = 0
    if HR_PATH is None:
        print(f"  WARNING: No cleaned HR file found in {onedrive_paths.CLEAN_HR_DIR} "
              "— run clean_hr.py first; backfill skipped.")
    else:
        print(f"Loading HR file for blank-leader backfill: {HR_PATH.name}")
        try:
            hr_data, hr_sheet_name = load_hr(HR_PATH)
        except Exception as e:
            print(f"  WARNING: Could not load HR file ({e}) — backfill skipped.")
            hr_data = {}
        hr_uids_loaded = len(hr_data)
        if hr_uids_loaded:
            print(f"  HR sheet: {hr_sheet_name}  |  {hr_uids_loaded:,} unique UIDs")
            print("Backfilling blank leader cells from HR (no overwrites)...")
            hr_counts = backfill_blanks_from_hr(wb, hr_data)
            print(f"  Rows touched: {hr_counts['rows_touched']:,}")
            for name in ("SeniorManager", "Director", "Sr. Director",
                         "AVP", "VP", "SVP", "Leaders"):
                if hr_counts[name]:
                    print(f"    → {name}: {hr_counts[name]:,} filled")

    # -- STEP 3d: identity backfill from RAW HR — fill-blanks-only ----------
    # 2026-09-04: rows whose FirstName/LastName are blank after the MVP pass
    # (MVP stub rows, or people not in MVP at all) get names / EmployeeID /
    # JobTitle / Department / manager UID from the raw HR export, then Full
    # Name is recomputed. Never overwrites a populated cell. Raw HR is only
    # read when at least one row needs it.
    id_counts = {"rows_needing": 0, "rows_in_hr": 0, "rows_touched": 0, "cells": 0}
    needing = rows_missing_identity(wb)
    if needing:
        raw_hr_path = onedrive_paths.RAW_HR_PATH
        print(f"Identity backfill: {len(needing)} row(s) with blank FirstName/LastName "
              f"— loading raw HR ({raw_hr_path.name})...")
        try:
            hr_raw = load_hr_raw_identity(raw_hr_path, {u for _, u in needing})
            id_counts = backfill_identity_from_raw_hr(wb, needing, hr_raw)
            print(f"  Rows in raw HR: {id_counts['rows_in_hr']} | rows touched: "
                  f"{id_counts['rows_touched']} | cells filled: {id_counts['cells']}")
            still = [u for _, u in needing if u not in hr_raw]
            if still:
                print(f"  Still blank (not in raw HR either): {', '.join(still[:20])}"
                      + (" ..." if len(still) > 20 else ""))
        except Exception as e:
            print(f"  WARNING: identity backfill skipped ({e}).")

    # -- STEP 4: Table1 ref + dimension + DASHBOARD date --------------------
    wb.sync_table_and_dimension()
    wb.set_other_cell("DASHBOARD", "C3",
                      datetime.date.today().strftime("%m/%d/%Y"))

    # -- STEP 5: Run Log row (in-workbook) ----------------------------------
    total_rows_end = sum(1 for r in wb.rows if r[0] != 1)
    write_run_log_row(wb, [
        ts.strftime("%Y-%m-%d %H:%M:%S"), total_rows_start, total_rows_end,
        blank_deleted, updated, not_in_mvp, newly_departed, tn_set_no,
        tn_departed, tn_why_reason, tn_job_role, pref_filled, pref_remote,
        pref_onsite, w1_to_w2, w2_to_w3, w3_to_w4, w_other_moves,
        total_departed_end, tn_yes_total, tn_no_total, MVP_PATH.name,
        duplicate_rows_count, promote_counts["promoted"],
        promote_counts["skipped_existing"], change_log_entries_written,
        hr_sheet_name, hr_uids_loaded, hr_counts["rows_touched"],
        hr_counts["SeniorManager"], hr_counts["Director"],
        hr_counts["Sr. Director"], hr_counts["AVP"],
        hr_counts["VP"], hr_counts["SVP"], hr_counts["Leaders"],
    ])

    # -- Save (surgical) + drop calcChain so Excel rebuilds it silently -----
    wb.drop_calc_chain()
    try:
        wb.save()
    except PermissionError:
        print(f"\nERROR: Cannot save {MASTER_PATH.name} — file became locked.")
        print(f"Your work was NOT saved. Backup intact: {backup_path.name}")
        sys.exit(1)

    # -- Versioned DATA-only snapshot (lightweight, one per successful run) --
    # Best-effort: never masks the successful save above (returns None on failure).
    data_snapshot_path = master_backup.save_data_snapshot(MASTER_PATH)
    if data_snapshot_path:
        print(f"Data snapshot: {data_snapshot_path}")

    # -- Console summary ----------------------------------------------------
    print("\n=== RESULTS ===")
    print(f"  Rows updated from MVP:          {updated:,}")
    print(f"  Rows NOT in MVP (noted):        {not_in_mvp:,}")
    print(f"  Blank rows (manual review):     {len(blanks):,}")
    print(f"  New users promoted to DATA:     {promote_counts['promoted']:,}"
          f"  (skipped {promote_counts['skipped_existing']} already on DATA)")
    print(f"  Change log entries written:     {change_log_entries_written:,}")
    print(f"  Duplicate rows (review):        {duplicate_rows_count:,}"
          f" ({duplicate_uids_count} UID{'s' if duplicate_uids_count != 1 else ''})")
    print(f"  Newly departed/inactive:        {newly_departed:,}")
    print(f"  Training Needed set to No:      {tn_set_no:,}")
    print(f"    → Departed/Inactive:          {tn_departed:,}")
    print(f"    → Why Training reason:        {tn_why_reason:,}")
    print(f"    → No-training job role:       {tn_job_role:,}")
    print(f"  Training Pref auto-filled:      {pref_filled:,}  "
          f"({pref_remote} Remote / {pref_onsite} Onsite)")
    total_wave_moves = w1_to_w2 + w2_to_w3 + w3_to_w4 + w_other_moves
    if total_wave_moves:
        print(f"  Wave movements this run:        {total_wave_moves:,}")
        print(f"    Wave 1 → 2:                   {w1_to_w2:,}")
        print(f"    Wave 2 → 3:                   {w2_to_w3:,}")
        print(f"    Wave 3 → 4:                   {w3_to_w4:,}")
        if w_other_moves:
            print(f"    Other wave changes:           {w_other_moves:,}")
    else:
        print("  Wave movements this run:        none")
    print("\n  Epic TM Lookup mirror (EpicTMLookup_ block):")
    if epic_cols:
        print(f"    Source: {epic_src}  ({len(epic_cols)} column(s) mapped)")
        print(f"    Rows matched:           {epic_matched:,}")
        print(f"    Not in lookup (blank):  {epic_not_in_lookup:,}")
        print(f"    Cells changed:          {epic_cells_changed:,}")
    else:
        print(f"    Skipped ({epic_src if epic_rows is None else 'no mapped columns found'}).")

    print("\n  HR backfill (blank leader cells):")
    if hr_uids_loaded:
        print(f"    HR sheet: {hr_sheet_name}  ({hr_uids_loaded:,} UIDs)")
        print(f"    Rows touched:           {hr_counts['rows_touched']:,}")
        for _name in ("SeniorManager", "Director", "Sr. Director",
                      "AVP", "VP", "SVP", "Leaders"):
            print(f"    {_name+':':<24}{hr_counts[_name]:,}")
    else:
        print("    Skipped (HR file unavailable or empty).")

    print("\n  End-of-run totals:")
    print(f"    Total users:                  {total_rows_end:,}")
    print(f"    Training Needed = Yes:        {tn_yes_total:,}")
    print(f"    Training Needed = No:         {tn_no_total:,}")
    print(f"    Departed/Inactive:            {total_departed_end:,}")
    print(f"\n  Run Log row appended: '{RUN_LOG_SHEET}' sheet in Master Wave File.")

    # -- Text log file ------------------------------------------------------
    with open(LOG_PATH, "w", encoding="utf-8") as log:
        log.write("update_master_from_mvp.py — run log\n")
        log.write("=" * 60 + "\n\n")
        log.write(f"Run:         {ts.strftime('%Y-%m-%d %H:%M:%S')}\n")
        log.write(f"MVP path:    {MVP_PATH}\n")
        log.write(f"Master path: {MASTER_PATH}\n")
        log.write(f"Backup:      {backup_path}\n")
        log.write(f"Data snapshot: {data_snapshot_path}\n\n")
        log.write(f"MVP unique UIDs:                {len(mvp):,}\n")
        log.write(f"Total rows (start):             {total_rows_start:,}\n")
        log.write(f"Total rows (end):               {total_rows_end:,}\n")
        log.write(f"Blank rows (manual review):     {len(blanks):,}\n")
        log.write(f"New users promoted to DATA:     {promote_counts['promoted']:,}"
                  f" (skipped {promote_counts['skipped_existing']} already on DATA,"
                  f" {promote_counts['blank_skipped']} blank staging rows)\n")
        log.write(f"Change log entries written:     {change_log_entries_written:,}\n")
        log.write(f"Duplicate rows (review):        {duplicate_rows_count:,}"
                  f" ({duplicate_uids_count} UIDs)\n")
        log.write(f"Rows updated from MVP:          {updated:,}\n")
        log.write(f"Rows NOT in MVP (noted):        {not_in_mvp:,}\n")
        log.write(f"Newly departed/inactive:        {newly_departed:,}\n")
        log.write(f"Training Needed set to No:      {tn_set_no:,}\n")
        log.write(f"  → Departed/Inactive:          {tn_departed:,}\n")
        log.write(f"  → Why Training reason:        {tn_why_reason:,}\n")
        log.write(f"  → No-training job role:       {tn_job_role:,}\n")
        log.write(f"Training Pref auto-filled:      {pref_filled:,} "
                  f"({pref_remote} Remote / {pref_onsite} Onsite)\n")
        log.write(f"Wave 1 → 2:                     {w1_to_w2:,}\n")
        log.write(f"Wave 2 → 3:                     {w2_to_w3:,}\n")
        log.write(f"Wave 3 → 4:                     {w3_to_w4:,}\n")
        log.write(f"Wave changes (other):           {w_other_moves:,}\n")
        log.write(f"Total departed/inactive (end):  {total_departed_end:,}\n")
        log.write(f"Total Training Needed = Yes:    {tn_yes_total:,}\n")
        log.write(f"Total Training Needed = No:     {tn_no_total:,}\n")
        log.write(f"Total users (end):              {total_rows_end:,}\n\n")
        log.write("Epic TM Lookup mirror (EpicTMLookup_ block):\n")
        if epic_cols:
            log.write(f"  Source:                  {epic_src}\n")
            log.write(f"  Columns mapped:          {len(epic_cols)}\n")
            log.write(f"  Rows matched:            {epic_matched:,}\n")
            log.write(f"  Not in lookup (blank):   {epic_not_in_lookup:,}\n")
            log.write(f"  Cells changed:           {epic_cells_changed:,}\n\n")
        else:
            log.write("  Skipped "
                      f"({epic_src if epic_rows is None else 'no mapped columns found'}).\n\n")
        log.write("HR backfill — fill-blanks-only:\n")
        if hr_uids_loaded:
            log.write(f"  HR sheet:                {hr_sheet_name}\n")
            log.write(f"  HR UIDs loaded:          {hr_uids_loaded:,}\n")
            log.write(f"  Rows touched:            {hr_counts['rows_touched']:,}\n")
            for _name in ("SeniorManager", "Director", "Sr. Director",
                          "AVP", "VP", "SVP", "Leaders"):
                log.write(f"  {_name+':':<24} {hr_counts[_name]:,}\n")
        else:
            log.write("  Skipped (HR file unavailable or empty).\n")
        if duplicate_uid_to_rows:
            log.write("\nDuplicate UIDs found (manual review needed):\n")
            log.write("-" * 60 + "\n")
            for uid_val, rs in sorted(duplicate_uid_to_rows.items()):
                log.write(f"  {uid_val}: rows "
                          f"{', '.join(str(r) for r in sorted(rs))}\n")

    append_external_runlog(
        len(mvp), updated, not_in_mvp, blank_deleted,
        tn_set_no, total_rows_end,
        w1_to_w2, w2_to_w3, w3_to_w4, w_other_moves,
        total_departed_end, tn_yes_total, tn_no_total,
        backup_path, data_snapshot_path,
    )

    print(f"  Log written: {LOG_PATH.name}")
    print(f"  External run log appended: master_runlog.xlsx (tab: {RUNLOG_TAB})")
    print(f"  Saved: {MASTER_PATH.name}")


if __name__ == "__main__":
    with activity_log.track_run("update_master_from_mvp.py"):
        main()
