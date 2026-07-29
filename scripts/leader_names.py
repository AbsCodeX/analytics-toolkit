# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
Shared utilities for leader name normalization, Notes column management,
no-training job role lookup, and Dashboard date writing.
"""

import datetime
import re

# ---------------------------------------------------------------------------
# Canonical leader names (exact spelling for VP / Leaders / SVP columns)
# ---------------------------------------------------------------------------
# Removed 2026-07-09 (confirmed with user): "FormerLeaderA, Firstname" (retired, no
# longer with the company) and "FormerLeaderB, Firstname" (AVP under Lastname16, not a
# Leader). List now matches the Master REF sheet's leader list exactly.
CANONICAL_NAMES = [
    "Lastname01, Firstname01",
    "Lastname02, Firstname02",
    "Lastname03, Firstname03",
    "Lastname04, Firstname04",
    "Lastname05, Firstname05",
    "Lastname06, Firstname06",
    "Lastname07, Firstname07",
    "Lastname08, Firstname08",
    "Lastname09, Firstname09",
    "Lastname10, Firstname10",
    "Lastname11, Firstname11",
    "Lastname12, Firstname12",
    "Lastname13, Firstname13",
    "Lastname14, Firstname14",
    "Lastname15, Firstname15",
    "Lastname16, Firstname16",
]

# Fast membership check for "is this leader on our list?"
CANONICAL_SET = frozenset(CANONICAL_NAMES)

# Leaders-cell value used when a computed leader is outside CANONICAL_SET.
# Renamed 2026-06-05 from "Not Rev Cycle". The legacy spelling is still
# recognized everywhere (comparison, classification, migration) so old cells
# flip to the new wording cleanly without being re-logged as changes.
NOT_REV_CYCLE = "No Longer Rev Cycle"
LEGACY_NOT_REV_CYCLE = "Not Rev Cycle"
_NOT_REV_CYCLE_VARIANTS = {NOT_REV_CYCLE.lower(), LEGACY_NOT_REV_CYCLE.lower()}


def is_not_rev_cycle(value) -> bool:
    """True if value is the current OR legacy 'no longer rev cycle' wording."""
    if value is None:
        return False
    return str(value).strip().lower() in _NOT_REV_CYCLE_VARIANTS

# Explicit aliases: HR values that don't resolve via last-name lookup
_ALIASES = {
    "aliasfirst aliaslast":  "Lastname12, Firstname12",
    "aliaslast, aliasfirst": "Lastname12, Firstname12",
    "aliaslast aliasfirst":  "Lastname12, Firstname12",
}

# Build last-name → canonical lookup (lowercase key for case-insensitive matching)
_BY_LAST = {name.split(",")[0].strip().lower(): name for name in CANONICAL_NAMES}


def normalize_leader(raw) -> str | None:
    """
    Map a raw leader name string to its canonical form.

    Resolution order:
      1. None / empty / 'nan, nan'  → None
      2. References Lastname03          → 'Lastname03, Firstname03'
      3. Known alias                → canonical form
      4. Last-name match            → canonical form
      5. No match                   → cleaned string as-is
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan, nan", "nan"):
        return None

    # Lastname03 special case — only when the name actually references Lastname03
    # ("Lastname03 Jr, Firstname03", "Firstname03, Lastname03") or is the bare "Jr, Firstname03"
    # HR artifact. A plain substring test on 'firstname03' wrongly rewrote other
    # real leaders (Otherlast1, Otherfirst1; Otherlast2, Firstname03; Otherlast3, Firstname03;
    # Somelast, Somefirst) to Lastname03 — fixed 2026-07-13.
    if "lastname03" in s.lower() or s.lower() in ("jr, firstname03", "firstname03 jr"):
        return "Lastname03, Firstname03"

    # Explicit alias table
    if s.lower() in _ALIASES:
        return _ALIASES[s.lower()]

    # Last-name lookup
    if "," in s:
        last = s.split(",")[0].strip().lower()
    else:
        parts = s.strip().split()
        last = parts[-1].lower() if parts else ""

    return _BY_LAST.get(last, s)


# ---------------------------------------------------------------------------
# Notes column helpers
# ---------------------------------------------------------------------------
NOTE_NOT_IN_MVP    = "Not in MVP"
NOTE_NOT_IN_HR     = "User Not in HR Export."
NOTE_HR_BLANK_LEADERSHIP = "User Has Blank Leadership in HR Export — manual review."
NOTE_DUPLICATE_UID = "Dup UID"
_SEP = " | "

# Legacy long-form notes written by prior runs -> current short form. Existing
# Master cells are upgraded in place whenever add_note/remove_note touch them,
# so the short forms never collide with (or leave fragments of) the old strings
# — important because the short "Not in MVP" is a substring of the old text.
_LEGACY_NOTE_REWRITES = {
    "User Not in MVP User Mappings.": NOTE_NOT_IN_MVP,
    "DUPLICATE UID — manual review needed.": NOTE_DUPLICATE_UID,
    "DUPLICATE UID - manual review needed.": NOTE_DUPLICATE_UID,
}


def _migrate_legacy(s: str) -> str:
    for old, new in _LEGACY_NOTE_REWRITES.items():
        if old in s:
            s = s.replace(old, new)
    return s


def add_note(existing, note: str) -> str:
    """Append note to existing Notes value, deduplicated."""
    if not existing:
        return note
    s = _migrate_legacy(str(existing))
    if note in s:
        return s
    return s + _SEP + note


def remove_note(existing, note: str):
    """Remove a specific note from Notes value. Returns None if result is empty."""
    if not existing:
        return None
    s = _migrate_legacy(str(existing))
    s = s.replace(_SEP + note, "").replace(note + _SEP, "").replace(note, "")
    s = s.strip().strip("|").strip()
    return s if s else None


# ---------------------------------------------------------------------------
# No-training job roles (REF sheet, column I, rows 4+)
# ---------------------------------------------------------------------------
def load_no_training_roles(wb) -> set:
    """
    Load the no-training job role list from column I of the REF sheet.
    Returns a lowercase set for case-insensitive matching.
    First 3 rows of col I are headers — skipped.
    """
    roles = set()
    if "REF" not in wb.sheetnames:
        return roles
    ws = wb["REF"]
    for row in ws.iter_rows(min_row=4, min_col=9, max_col=9, values_only=True):
        val = row[0]
        if val is not None:
            roles.add(str(val).strip().lower())
    return roles


# ---------------------------------------------------------------------------
# Dashboard date
# ---------------------------------------------------------------------------
def update_dashboard_date(wb) -> None:
    """Write today's run date (MM/DD/YYYY) to DASHBOARD cell C3."""
    if "DASHBOARD" not in wb.sheetnames:
        return
    wb["DASHBOARD"]["C3"].value = datetime.date.today().strftime("%m/%d/%Y")


# ---------------------------------------------------------------------------
# Leader change-log helpers (col AO "Users HR Change Log")
# ---------------------------------------------------------------------------
CHANGE_LOG_SEP = " | "


def parse_hr_export_date(sheet_name: str) -> str:
    """
    Extract MM/DD/YY from an HR sheet name like 'OCM_HRDatafile_04012026'.
    Falls back to today's date if no MMDDYYYY substring is found.
    """
    m = re.search(r"(\d{2})(\d{2})(\d{4})", sheet_name or "")
    if not m:
        return datetime.date.today().strftime("%m/%d/%y")
    mm, dd, yyyy = m.groups()
    return f"{mm}/{dd}/{yyyy[-2:]}"


def _display(v):
    """Stripped string for display; None for empty/missing."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s else None


def _compare_key(v):
    """Case-folded key for equality comparison. Treats 'Lastname05' == 'Lastname05',
    and the legacy 'Not Rev Cycle' == the current 'No Longer Rev Cycle' (so the
    rename never registers as a leader change)."""
    d = _display(v)
    if d is None:
        return None
    if is_not_rev_cycle(d):
        return NOT_REV_CYCLE.casefold()
    return d.casefold()


def format_leader_change(run_date: str, hr_date: str, old, new) -> str | None:
    """
    Build a change-log entry. Returns None when old == new (no change).
    Equality is case-insensitive so spelling normalizations (Lastname05 -> Lastname05)
    don't get logged as changes. Display uses the actual strings as they appear.

    Wording:
      newly assigned : '05/20/26 - Newly assigned to Lastname08 per 05/20/26 HR file'
      cleared        : '05/20/26 - Lastname07 removed per 05/20/26 HR file'
      transition     : '05/20/26 - Lastname07 to Lastname08 per 05/20/26 HR file'
    """
    if _compare_key(old) == _compare_key(new):
        return None

    old_d = _display(old)
    new_d = _display(new)

    if old_d is None:
        body = f"Newly assigned to {new_d}"
    elif new_d is None:
        body = f"{old_d} removed"
    else:
        body = f"{old_d} to {new_d}"

    return f"{run_date} - {body} per {hr_date} HR file"


def append_change_log(existing, entry: str) -> str:
    """Append a change-log entry to existing AO value, deduplicated."""
    if not existing:
        return entry
    s = str(existing).strip()
    if not s:
        return entry
    if entry in s:
        return s
    return s + CHANGE_LOG_SEP + entry


def classify_leader_change(old, new) -> str | None:
    """
    Classify a leader transition into one bucket for aggregate counting.
    Returns one of: 'reassigned', 'fell_out', 'reentered', 'newly_assigned',
    'cleared', or None when no change. Equality is case-insensitive.
    """
    if _compare_key(old) == _compare_key(new):
        return None
    old_d = _display(old)
    new_d = _display(new)
    if old_d is None:
        return "newly_assigned"
    if new_d is None:
        return "cleared"
    if is_not_rev_cycle(new_d):
        return "fell_out"
    if is_not_rev_cycle(old_d):
        return "reentered"
    return "reassigned"
