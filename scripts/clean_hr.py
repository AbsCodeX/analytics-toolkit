# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
Clean the HR.xlsx file and produce a two-tab output workbook for the RCM Wave program.

Tab "HR Cleaned" (filtered) columns:
    Universal ID, Full Name, SeniorManager, Director, SeniorDirector, AVP, VP, SVP,
    EVP, Leaders

Tab "All Staff" (all rows, unfiltered) columns:
    UniversalID, Full Name, GoLiveWave, Email, SeniorManager, Director,
    SeniorDirector, AVP, VP, SVP, EVP, Leaders, IsDeparted/Inactive?, IsRCM

ID -> name resolution:
    Leader-chain cells that contain a Universal ID instead of a name (e.g. "JDOE1")
    are resolved to "Last, First" using a UniversalID -> Employee Name map built from
    the full HR file. This runs BEFORE name standardization, so VP/AVP/SVP and the
    derived Leaders column never show a raw ID. Unresolved IDs are left as-is and
    reported. Resolution and standardization are applied to the full dataset, so both
    tabs are clean.

Row filter (HR Cleaned tab only):
    Keep rows where IsRCM=True OR Firstname09 Lastname09 or Firstname02 Lastname02 appear anywhere
    in the leader chain. Lastname09 and Lastname02 themselves are always kept.
    Exception: a row is dropped even if it matched above when AVP, VP, and SVP
    are ALL blank (no way to roll it up to Lastname03's org), unless the employee
    is one of the CANONICAL_LEADERS themselves (e.g. Lastname03, whose own row has
    nothing populated in AVP/VP/SVP because he sits at the top of the chain).

Leaders column logic (exactly as specified, no extra cases):
    Leaders is only ever an AVP or VP name, with SVP as a fallback — never a
    SeniorDirector name (SeniorDirector is an intermediate rank, not a
    recognized Leader tier).
    1. If the user IS the VP (VP cell == own name) -> SVP
    2. Else if VP is set -> VP
    3. Else if AVP is set -> AVP
    4. Else if SVP is set -> SVP (covers users who report directly to an SVP
       with no VP/AVP layer between them — e.g. VPs themselves, or anyone
       whose closest populated rollup is only a SeniorDirector)
    5. Else -> the user's own name (covers users at the top of the chain,
       e.g. Lastname03 himself, whose own row has nothing populated even up
       through SVP). Leaders is never left blank.

Name standardization:
    Fixed alias map first (Lastname07 Nickname07 -> Lastname07 Firstname07, etc.), then a strict
    fuzzy backup: only matches when the last name (part before the comma) is an
    exact case-insensitive match AND the first name fuzzy-matches at cutoff 0.85.
    This blocks false positives like Lastname12, Melissa -> Lastname12, Firstname12 or
    SimilarLastname, Firstname10 -> Lastname10, Firstname10. Anything genuine the fuzzy
    matcher misses can be added explicitly to NAME_ALIASES.
"""

from __future__ import annotations

import argparse
import re
import sys
from difflib import get_close_matches
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).parent.parent

sys.path.insert(0, str(Path(__file__).parent))
import onedrive_paths  # noqa: E402


# ============================================================
# CONSTANTS
# ============================================================

NBSP = "\xa0"
INVISIBLE_PATTERN = "[\u200b-\u200d\ufeff]"
NULL_STRINGS = {"nan", "nan, nan", "nat", "none", "null"}

FUZZY_CUTOFF = 0.85

HR_LEADER_COLUMNS = [
    "SeniorManager",
    "Director",
    "SeniorDirector",
    "AVP",
    "VP",
    "SVP",
    "EVP",
]

CANONICAL_LEADERS = [
    "Lastname14, Firstname14",
    "Lastname15, Firstname15",
    "Lastname04, Firstname04",
    "Lastname10, Firstname10",
    "Lastname11, Firstname11",
    "Lastname16, Firstname16",
    "Lastname05, Firstname05",
    "Lastname08, Firstname08",
    "Lastname07, Firstname07",
    "Lastname01, Firstname01",
    "Lastname03, Firstname03",
    "Lastname09, Firstname09",
    "Lastname02, Firstname02",
    "Lastname12, Firstname12",
    "Lastname06, Firstname06",
    "Lastname13, Firstname13",
    "FormerLeaderB, Firstname",
]

# Explicit aliases — left side compared case-insensitively after trimming.
# Use this for known misspellings the fuzzy matcher would miss (nicknames, swapped names).
NAME_ALIASES = {
    "lastname07, nickname07": "Lastname07, Firstname07",
    "lastname07, nickname07": "Lastname07, Firstname07",
    "jr, firstname03": "Lastname03, Firstname03",
    "lastname03 jr, firstname03": "Lastname03, Firstname03",
    "aliasfirst aliaslast": "Lastname12, Firstname12",
    "aliaslast aliasfirst": "Lastname12, Firstname12",
    "aliaslast, aliasfirst": "Lastname12, Firstname12",
    "lastname06, firstname06 a.": "Lastname06, Firstname06",
}

ROW_FILTER_LEADERS = {"lastname09, firstname09", "lastname02, firstname02"}

TRUE_STRINGS = {"true", "1", "yes", "y", "t"}

# Source columns in the raw HR file used to build the output tabs.
SRC_UID = "UniversalID"
SRC_FULL_NAME = "Employee Name"
SRC_WAVE = "Wave"
SRC_EMAIL = "Email Address"
SRC_DEPARTED = "IsDepartedInactive"
SRC_IS_RCM = "IsRCM"

# Secondary ID -> name source: MVP User Mappings. Used only to resolve IDs that are
# not present as employee rows in the HR file (HR's own Employee Name wins on conflict).
MVP_PATH = onedrive_paths.RAW_MVP_USER_MAPPINGS
MVP_UID = "UniversalID"
MVP_FIRST = "FirstName"
MVP_LAST = "LastName"

# Output column order for the "All Staff" tab (all rows, unfiltered).
ALL_STAFF_COLUMNS = [
    "UniversalID",
    "Full Name",
    "GoLiveWave",
    "Email",
    *HR_LEADER_COLUMNS,
    "Leaders",
    "IsDeparted/Inactive?",
    "IsRCM",
]

_ID_RE = re.compile(r"[A-Za-z]")


def looks_like_id(value: str) -> bool:
    """True if a cell looks like a Universal ID rather than a 'Last, First' name."""
    value = value.strip()
    return (
        value != ""
        and "," not in value
        and " " not in value
        and bool(_ID_RE.search(value))
    )


# ============================================================
# ARGUMENTS
# ============================================================

def parse_args() -> argparse.Namespace:
    """Parse command-line arguments (input file and optional output path)."""
    parser = argparse.ArgumentParser(description=__doc__.strip().splitlines()[0])
    parser.add_argument(
        "input_file",
        nargs="?",
        default=str(onedrive_paths.RAW_HR_PATH),
        help=f"HR .xlsx file. Defaults to {onedrive_paths.RAW_HR_PATH}.",
    )
    parser.add_argument(
        "-o", "--output",
        help=f"Output .xlsx. Defaults to a dated HR_cleaned_YYYY.MM.DD.xlsx in {onedrive_paths.CLEAN_HR_DIR}.",
    )
    return parser.parse_args()


# ============================================================
# PATH HELPERS
# ============================================================

def resolve_input_path(input_file: str) -> Path:
    """Resolve and validate the HR input path; must be an existing .xlsx file."""
    path = Path(input_file).expanduser()
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
        if not path.exists():
            path = (ROOT / input_file).resolve()
    else:
        path = path.resolve()

    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"Unsupported file type: {path.suffix}. HR cleaner only supports .xlsx.")

    if not path.exists():
        raise FileNotFoundError(f"HR file not found: {path}")

    return path


def resolve_output_path(input_path: Path, output_file: str | None) -> Path:
    """Resolve the output path. Defaults to a dated file in the OneDrive
    clean_hr_copies folder (e.g. HR_cleaned_2026.07.01.xlsx) — one file per run,
    never overwritten by a later day's run. Consumers look up the latest one via
    onedrive_paths.latest_hr_cleaned() rather than a fixed filename."""
    if output_file:
        path = Path(output_file).expanduser()
        if not path.is_absolute():
            path = (Path.cwd() / path).resolve()
        else:
            path = path.resolve()
    else:
        path = onedrive_paths.CLEAN_HR_DIR / f"{onedrive_paths.dated_stem('HR_cleaned')}.xlsx"

    if path.suffix.lower() != ".xlsx":
        raise ValueError("Output must end with .xlsx")
    if path == input_path:
        raise ValueError("Output cannot be the same as the input.")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_text_series(series: pd.Series) -> pd.Series:
    """Normalize text: replace NBSP/invisible chars, collapse whitespace, null out blanks."""
    series = series.astype("string")
    series = series.str.replace(NBSP, " ", regex=False)
    series = series.str.replace(INVISIBLE_PATTERN, "", regex=True)
    series = series.str.replace(r"\s+", " ", regex=True)
    series = series.str.strip()
    series = series.where(series != "", pd.NA)
    null_mask = series.str.lower().isin(NULL_STRINGS).fillna(False)
    series = series.mask(null_mask, pd.NA)
    return series


# ============================================================
# NAME STANDARDIZATION
# ============================================================

def split_last_first(name: str) -> tuple[str, str] | None:
    """Parse 'Last, First' into (last_lower, first_lower). Returns None if no comma."""
    if "," not in name:
        return None
    last, _, first = name.partition(",")
    last = last.strip().lower()
    first = first.strip().lower()
    if not last or not first:
        return None
    return last, first


def build_name_map(observed_values: list[str]) -> dict[str, str]:
    """
    Map each unique observed leader name to its canonical spelling.
    Computed once over the unique set, then applied to each leader column.

    Match order:
      1. Explicit NAME_ALIASES (case-insensitive).
      2. Exact canonical match (case-insensitive).
      3. Strict fuzzy: last name must match exactly (case-insensitive); first name
         must fuzzy-match at cutoff FUZZY_CUTOFF.
    """
    canonical_by_key = {name.lower(): name for name in CANONICAL_LEADERS}

    # Group canonical names by last-name for the strict-fuzzy step
    canonical_by_last: dict[str, list[tuple[str, str]]] = {}
    for canonical_name in CANONICAL_LEADERS:
        parsed = split_last_first(canonical_name)
        if parsed is None:
            continue
        last, first = parsed
        canonical_by_last.setdefault(last, []).append((first, canonical_name))

    mapping: dict[str, str] = {}
    for value in observed_values:
        key = value.strip().lower()

        if key in NAME_ALIASES:
            mapping[value] = NAME_ALIASES[key]
            continue

        if key in canonical_by_key:
            mapping[value] = canonical_by_key[key]
            continue

        parsed = split_last_first(value)
        if parsed is not None:
            obs_last, obs_first = parsed
            candidates = canonical_by_last.get(obs_last, [])
            if candidates:
                candidate_firsts = [first for first, _ in candidates]
                fuzzy_match = get_close_matches(
                    obs_first, candidate_firsts, n=1, cutoff=FUZZY_CUTOFF
                )
                if fuzzy_match:
                    matched_first = fuzzy_match[0]
                    canonical_name = next(
                        canonical for first, canonical in candidates if first == matched_first
                    )
                    mapping[value] = canonical_name
                    continue

        mapping[value] = value

    return mapping


def apply_name_map(series: pd.Series, name_map: dict[str, str]) -> pd.Series:
    """Map values through name_map, leaving NaN and unmapped values unchanged."""
    return series.map(lambda value: name_map.get(value, value) if pd.notna(value) else value)


# ============================================================
# ID -> NAME RESOLUTION
# ============================================================

def build_id_to_name_map(df: pd.DataFrame) -> dict[str, str]:
    """Map UniversalID (upper-cased) -> 'Last, First' full name from the HR file itself."""
    id_map: dict[str, str] = {}
    uids = df[SRC_UID]
    names = df[SRC_FULL_NAME]
    for uid, name in zip(uids, names):
        if pd.notna(uid) and pd.notna(name):
            key = str(uid).strip().upper()
            value = str(name).strip()
            if key and value:
                id_map[key] = value
    return id_map


def _fix_name_case(name: str) -> str:
    """Title-case only fully-upper or fully-lower names (e.g. 'ACOSTA, ILEANA');
    leave already mixed-case names (Lastname05, O'Kane, McGrinder) untouched."""
    if name == name.upper() or name == name.lower():
        return name.title()
    return name


def build_mvp_id_to_name_map() -> dict[str, str]:
    """Fallback map UniversalID (upper) -> 'Last, First' from MVP User Mappings.csv.

    Returns an empty map (with a warning) if the file is missing — the HR cleaner
    still runs, it just resolves fewer IDs.
    """
    if not MVP_PATH.exists():
        print(f"  NOTE: MVP fallback file not found, skipping: {MVP_PATH}", file=sys.stderr)
        return {}
    mvp = pd.read_csv(MVP_PATH, dtype=str, usecols=[MVP_UID, MVP_FIRST, MVP_LAST])
    id_map: dict[str, str] = {}
    for uid, first, last in zip(mvp[MVP_UID], mvp[MVP_FIRST], mvp[MVP_LAST]):
        if pd.notna(uid) and pd.notna(first) and pd.notna(last):
            key = str(uid).strip().upper()
            first_s, last_s = str(first).strip(), str(last).strip()
            if key and first_s and last_s and key not in id_map:
                id_map[key] = _fix_name_case(f"{last_s}, {first_s}")
    return id_map


def resolve_ids_in_series(
    series: pd.Series, id_map: dict[str, str], unresolved: set[str]
) -> tuple[pd.Series, int]:
    """Replace ID-looking cells with their resolved full name; collect unresolved IDs."""
    resolved = 0

    def _resolve(value):
        nonlocal resolved
        if pd.isna(value):
            return value
        text = str(value).strip()
        if looks_like_id(text):
            hit = id_map.get(text.upper())
            if hit:
                resolved += 1
                return hit
            unresolved.add(text)
        return value

    return series.map(_resolve), resolved


# ============================================================
# LEADERS COLUMN
# ============================================================

def compute_leaders_column(df: pd.DataFrame) -> pd.Series:
    """Vectorized implementation of the Leaders chain rule."""
    employee_name = df["Employee Name"].astype("string")
    vp = df["VP"].astype("string")
    avp = df["AVP"].astype("string")
    svp = df["SVP"].astype("string")

    name_norm = employee_name.str.strip().str.lower()
    vp_norm = vp.str.strip().str.lower()

    is_self_vp = vp.notna() & employee_name.notna() & (vp_norm == name_norm)

    vp_set = vp.notna() & (vp.str.strip() != "")
    avp_set = avp.notna() & (avp.str.strip() != "")
    svp_set = svp.notna() & (svp.str.strip() != "")

    conditions = [
        is_self_vp,
        ~is_self_vp & vp_set,
        ~is_self_vp & ~vp_set & avp_set,
        ~is_self_vp & ~vp_set & ~avp_set & svp_set,
    ]
    choices = [svp, vp, avp, svp]

    leaders = np.select(conditions, choices, default=employee_name)
    return pd.Series(leaders, index=df.index, dtype="object")


# ============================================================
# ROW FILTER
# ============================================================

def build_keep_mask(df: pd.DataFrame) -> tuple[pd.Series, dict[str, int]]:
    """Build the keep mask (IsRCM, target person, or reports to a target leader) plus stats.

    A row that matches on those grounds is still dropped if AVP, VP, and SVP are
    ALL blank — there is no chain data to roll it up to Lastname03's org — unless the
    employee is one of CANONICAL_LEADERS (e.g. Lastname03 himself sits at the top with
    nothing populated above him).
    """
    is_rcm = (
        df["IsRCM"].astype("string").str.strip().str.lower().isin(TRUE_STRINGS).fillna(False)
    )

    name_lower = df["Employee Name"].astype("string").str.strip().str.lower()
    is_target_person = name_lower.isin(ROW_FILTER_LEADERS).fillna(False)

    has_target_leader = pd.Series(False, index=df.index)
    for column in HR_LEADER_COLUMNS:
        column_lower = df[column].astype("string").str.strip().str.lower()
        has_target_leader = has_target_leader | column_lower.isin(ROW_FILTER_LEADERS).fillna(False)

    matched = is_rcm | is_target_person | has_target_leader

    def _blank(column: str) -> pd.Series:
        series = df[column].astype("string")
        return series.isna() | (series.str.strip() == "")

    no_rollup = _blank("AVP") & _blank("VP") & _blank("SVP")
    is_canonical_leader = name_lower.isin({leader.lower() for leader in CANONICAL_LEADERS}).fillna(False)
    excluded_no_rollup = matched & no_rollup & ~is_canonical_leader

    keep_mask = matched & ~excluded_no_rollup

    stats = {
        "is_rcm": int(is_rcm.sum()),
        "is_target_person": int(is_target_person.sum()),
        "has_target_leader": int(has_target_leader.sum()),
        "excluded_no_rollup": int(excluded_no_rollup.sum()),
        "total_kept": int(keep_mask.sum()),
    }
    return keep_mask, stats


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    """Read the HR export, clean/resolve it, and write the two-tab cleaned workbook."""
    args = parse_args()
    input_path = resolve_input_path(args.input_file)
    output_path = resolve_output_path(input_path, args.output)

    print(f"Reading: {input_path}")
    try:
        all_sheets = pd.read_excel(input_path, sheet_name=None, dtype=str, engine="openpyxl")
    except PermissionError as error:
        raise PermissionError(f"Close the file in Excel and re-run: {input_path}") from error

    sheet_name = next(iter(all_sheets))
    df = all_sheets[sheet_name]
    print(f"Sheet: {sheet_name!r} ({len(df):,} rows, {len(df.columns)} columns)")

    required = [
        SRC_UID, SRC_FULL_NAME, SRC_IS_RCM, SRC_WAVE, SRC_EMAIL, SRC_DEPARTED,
        *HR_LEADER_COLUMNS,
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"HR file is missing expected columns: {missing}.\n"
            f"Found columns: {list(df.columns)}"
        )

    print("Cleaning text in working columns...")
    working_columns = [SRC_UID, SRC_FULL_NAME, SRC_EMAIL, *HR_LEADER_COLUMNS]
    for column in working_columns:
        df[column] = clean_text_series(df[column])

    df[SRC_UID] = df[SRC_UID].str.upper()

    # --- Resolve ID-looking leader cells to full names ---
    # Primary source: HR's own Employee Name. Fallback: MVP User Mappings (for IDs
    # not present as HR employee rows). HR wins on any key collision.
    print("Resolving Universal IDs to full names in leader columns...")
    hr_map = build_id_to_name_map(df)
    mvp_map = build_mvp_id_to_name_map()
    id_map = {**mvp_map, **hr_map}
    print(
        f"  ID->name map: {len(hr_map):,} from HR, "
        f"{len(mvp_map):,} from MVP fallback, {len(id_map):,} combined."
    )
    unresolved_ids: set[str] = set()
    total_resolved = 0
    for column in HR_LEADER_COLUMNS:
        df[column], resolved = resolve_ids_in_series(df[column], id_map, unresolved_ids)
        total_resolved += resolved
    print(f"  Resolved {total_resolved:,} ID cell(s) to names.")
    if unresolved_ids:
        preview = sorted(unresolved_ids)[:15]
        print(
            f"  WARNING: {len(unresolved_ids)} ID value(s) could not be resolved "
            f"(left as-is): {preview}{' ...' if len(unresolved_ids) > 15 else ''}",
            file=sys.stderr,
        )

    print("Standardizing leader names...")
    observed_names: set[str] = set()
    for column in HR_LEADER_COLUMNS:
        observed_names.update(df[column].dropna().unique().tolist())

    name_map = build_name_map(sorted(observed_names))
    fixes = {orig: new for orig, new in name_map.items() if orig != new}
    print(f"  Mapped {len(fixes)} of {len(name_map)} unique leader name(s) to canonical spelling.")
    if fixes:
        sample = list(fixes.items())[:10]
        for original, canonical in sample:
            print(f"    {original!r} -> {canonical!r}")
        if len(fixes) > 10:
            print(f"    ... and {len(fixes) - 10} more")

    for column in HR_LEADER_COLUMNS:
        df[column] = apply_name_map(df[column], name_map)

    # --- Lastname03 tier rule: he is the SVP, never an AVP/VP ---
    # If a (standardized) AVP or VP cell says Lastname03, re-home him to SVP and
    # clear the cell, so the Leaders fallback (VP > AVP > SVP) still lands on
    # him without misstating his tier.
    lastname03 = "Lastname03, Firstname03"
    for column in ("AVP", "VP"):
        mask = df[column].astype("string").str.strip().eq(lastname03).fillna(False)
        if mask.any():
            print(f"  Lastname03 tier rule: moved {int(mask.sum())} {column} cell(s) to SVP.")
            df.loc[mask, "SVP"] = lastname03
            df.loc[mask, column] = pd.NA
    df["Full Name"] = df[SRC_FULL_NAME]
    df["GoLiveWave"] = clean_text_series(df[SRC_WAVE]).str.title()
    df["Email"] = df[SRC_EMAIL]
    df["IsDeparted/Inactive?"] = clean_text_series(df[SRC_DEPARTED])
    df["IsRCM"] = clean_text_series(df[SRC_IS_RCM])

    print("Computing Leaders column...")
    df["Leaders"] = compute_leaders_column(df)
    populated = int(df["Leaders"].notna().sum())
    print(f"  Leaders populated for {populated:,} of {len(df):,} rows.")

    # --- Tab 2: All Staff (all rows, unfiltered) ---
    all_staff_df = df[ALL_STAFF_COLUMNS].copy()
    print(f"All Staff tab: {len(all_staff_df):,} rows.")

    # --- Tab 1: HR Cleaned (filtered: IsRCM OR Lastname09/Lastname02 in chain) ---
    print("Filtering rows (IsRCM=True OR Lastname09/Lastname02 in chain)...")
    keep_mask, stats = build_keep_mask(df)
    print(
        f"  Kept {stats['total_kept']:,} of {len(df):,} rows."
        f" IsRCM={stats['is_rcm']:,},"
        f" Lastname09/Lastname02 person={stats['is_target_person']:,},"
        f" Lastname09/Lastname02 in chain={stats['has_target_leader']:,},"
        f" excluded (no AVP/VP/SVP rollup)={stats['excluded_no_rollup']:,}."
    )
    filtered_columns = [SRC_UID, "Full Name", *HR_LEADER_COLUMNS, "Leaders"]
    out_df = df.loc[keep_mask, filtered_columns].rename(columns={SRC_UID: "Universal ID"})

    blank_ids = int(out_df["Universal ID"].isna().sum())
    if blank_ids:
        print(
            f"  WARNING: {blank_ids:,} filtered row(s) have a blank Universal ID.",
            file=sys.stderr,
        )

    leader_counts = out_df["Leaders"].value_counts(dropna=True)
    print("Leaders distribution (filtered tab):")
    for leader, count in leader_counts.items():
        print(f"  {leader}: {count:,}")

    print(f"Writing: {output_path}")
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="HR Cleaned")
            all_staff_df.to_excel(writer, index=False, sheet_name="All Staff")
    except PermissionError as error:
        raise PermissionError(f"Close the output in Excel and re-run: {output_path}") from error

    print(
        f"Done. HR Cleaned: {len(out_df):,} rows x {len(out_df.columns)} cols; "
        f"All Staff: {len(all_staff_df):,} rows x {len(all_staff_df.columns)} cols -> {output_path}"
    )


if __name__ == "__main__":
    import activity_log
    with activity_log.track_run("clean_hr.py"):
        try:
            main()
        except Exception as error:
            print(f"ERROR: {error}", file=sys.stderr)
            raise SystemExit(1) from error
