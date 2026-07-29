# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_reference_lists.py — regenerate wave_reference_lists.xlsx clean.

The reference workbook's sheets auto-load into SQL as raw.ref_<sheet> on every
refresh (sql\\refresh.py refs) and back the dim.* views + report.mapping_gaps.
They were originally seeded from wave-file contents and picked up pivot debris
(header leaks, note rows, spilled analysis cells). This generator rebuilds
each sheet from its proper source while carrying every hand-made mapping
forward:

  business_units             row set = MVP (IsRCMUser=True) BUs  UNION  Master
                             BUs  UNION  existing mapped rows; IsVendor/UserType
                             carried; Status = Mapped/NeedsMapping; Source =
                             MVP/Legacy. Nothing mapped is dropped silently.
  leaders                    generated from leader_names.CANONICAL_NAMES —
                             the sheet is a build artifact of the Python
                             source of truth (single copy problem solved).
  leader_training_preference row set = CANONICAL_NAMES; preferences carried
                             (names canonicalized via normalize_leader, so
                             'AliasLast, AliasFirst' folds into 'Lastname12, Firstname12' etc.).
  no_training_job_roles      real roles kept, junk rows dropped, deduped.
  vendor_aliases             junk rows dropped (Grand Total / rule text),
                             deduped case-insensitively on Alias.
  job_titles_leader_flag     deduped only (content preserved).

USAGE
-----
    python scripts\\build_reference_lists.py            # DRY RUN — console diff + changes CSV
    python scripts\\build_reference_lists.py --apply    # backup workbook -> write rebuilt sheets

After --apply, reload SQL:  python sql\\refresh.py refs
"""

import argparse
import datetime
import shutil
import sys
import warnings
from pathlib import Path

import openpyxl
import pandas as pd

warnings.filterwarnings("ignore")

sys.path.insert(0, str(Path(__file__).parent))
import onedrive_paths as op
import runlog
from activity_log import track_run
from leader_names import CANONICAL_NAMES, normalize_leader

REF_PATH = op.REF_LISTS_PATH
ARCHIVE_DIR = op.REFERENCES_DIR / "archive"
CHANGES_CSV = Path(__file__).parent / "build_reference_lists_changes.csv"

EXPECTED_SHEETS = [
    "business_units",
    "job_titles_leader_flag",
    "no_training_job_roles",
    "leaders",
    "leader_training_preference",
    "vendor_aliases",
]

# Exact junk strings observed in the sheets (deterministic removal — anything
# else unexpected is kept and flagged rather than silently dropped).
NO_TRAINING_JUNK = {"no training needed job roles", "mapping in progress"}
ALIAS_JUNK_CANONICALS = {"grand total"}

RUNLOG_TAB = "Reference Lists Rebuild"
RUNLOG_HEADERS = ["Run Timestamp", "Mode", "Sheet", "Before", "After",
                  "Added", "Dropped", "Flagged", "Detail"]


def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    return s


changes = []  # (sheet, action, key, detail)


def log_change(sheet, action, key, detail=""):
    changes.append({"Sheet": sheet, "Action": action, "Key": key, "Detail": detail})


# ---------------------------------------------------------------------------
# Sources
# ---------------------------------------------------------------------------
def load_current_sheets():
    if not REF_PATH.exists():
        sys.exit(f"ERROR: reference workbook not found: {REF_PATH}")
    xl = pd.ExcelFile(REF_PATH, engine="openpyxl")
    missing = [s for s in EXPECTED_SHEETS if s not in xl.sheet_names]
    if missing:
        sys.exit(f"ERROR: workbook is missing expected sheet(s) {missing} — "
                 f"found {xl.sheet_names}. Aborting (sheet names drive raw.ref_* tables).")
    extra = [s for s in xl.sheet_names if s not in EXPECTED_SHEETS]
    if extra:
        print(f"  NOTE: extra sheet(s) preserved untouched: {extra}")
    return {s: pd.read_excel(xl, sheet_name=s, dtype=str) for s in xl.sheet_names}


def load_mvp_rcm_bus():
    path = op.RAW_MVP_USER_MAPPINGS
    if not path.exists():
        sys.exit(f"ERROR: MVP file not found: {path}")
    mvp = pd.read_csv(path, dtype=str, usecols=["BusinessUnit", "IsRCMUser"])
    rcm = mvp[mvp["IsRCMUser"].astype(str).str.strip().str.lower() == "true"]
    bus = sorted({b for b in (clean(x) for x in rcm["BusinessUnit"]) if b})
    if not bus:
        sys.exit("ERROR: 0 RCM business units loaded from MVP — refusing to continue.")
    return bus


def load_master_bus():
    wb = openpyxl.load_workbook(op.MASTER_WAVE_PATH, read_only=True, data_only=True)
    ws = wb["DATA"]
    headers = [str(c.value).strip() if c.value is not None else ""
               for c in next(ws.iter_rows(max_row=1))]
    bu_i = headers.index("BusinessUnit")
    bus = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        b = clean(row[bu_i]) if bu_i < len(row) else None
        if b:
            bus.add(b)
    wb.close()
    return sorted(bus)


# ---------------------------------------------------------------------------
# Sheet builders — each returns the new DataFrame
# ---------------------------------------------------------------------------
def build_business_units(cur: pd.DataFrame) -> pd.DataFrame:
    existing = {}
    for _, r in cur.iterrows():
        bu = clean(r.get("BusinessUnit"))
        if not bu:
            log_change("business_units", "dropped", str(r.get("BusinessUnit")),
                       "blank/junk BU key")
            continue
        existing[bu.upper()] = {"BusinessUnit": bu,
                                "IsVendor": clean(r.get("IsVendor")),
                                "UserType": clean(r.get("UserType"))}

    mvp_bus = load_mvp_rcm_bus()
    master_bus = load_master_bus()
    mvp_upper = {b.upper() for b in mvp_bus}

    all_bus = {}
    for b in mvp_bus + master_bus + [v["BusinessUnit"] for v in existing.values()]:
        all_bus.setdefault(b.upper(), b)

    rows = []
    for key in sorted(all_bus):
        bu = existing.get(key, {}).get("BusinessUnit") or all_bus[key]
        is_vendor = existing.get(key, {}).get("IsVendor")
        user_type = existing.get(key, {}).get("UserType")
        # A YourOrg BU (IsVendor=No) with blank UserType defaults cleanly:
        if is_vendor and str(is_vendor).strip().lower() == "no" and not user_type:
            user_type = "YourOrg"
            log_change("business_units", "edited", bu,
                       "UserType blank + IsVendor=No -> defaulted to YourOrg")
        status = "Mapped" if (is_vendor and user_type) else "NeedsMapping"
        source = "MVP" if key in mvp_upper else "Legacy"
        if key not in existing:
            log_change("business_units", "added", bu, f"new from {source}; needs mapping")
        if status == "NeedsMapping":
            log_change("business_units", "flagged", bu, "NeedsMapping (IsVendor/UserType incomplete)")
        rows.append({"BusinessUnit": bu, "IsVendor": is_vendor, "UserType": user_type,
                     "Status": status, "Source": source})
    return pd.DataFrame(rows)


def build_leaders(cur: pd.DataFrame) -> pd.DataFrame:
    cur_names = {clean(v) for v in cur.get("Leader", pd.Series(dtype=str))}
    cur_names.discard(None)
    for junk in sorted(cur_names - set(CANONICAL_NAMES)):
        log_change("leaders", "dropped", junk, "not in leader_names.CANONICAL_NAMES")
    for new in sorted(set(CANONICAL_NAMES) - cur_names):
        log_change("leaders", "added", new, "in CANONICAL_NAMES but missing from sheet")
    return pd.DataFrame({"Leader": list(CANONICAL_NAMES)})


def build_leader_training_preference(cur: pd.DataFrame) -> pd.DataFrame:
    prefs = {}
    for _, r in cur.iterrows():
        raw_name, pref = clean(r.get("Leader")), clean(r.get("TrainingPreference"))
        if not raw_name:
            continue
        canon = normalize_leader(raw_name)
        if canon in set(CANONICAL_NAMES):
            if pref and pref != "—":
                if canon in prefs and prefs[canon] != pref:
                    log_change("leader_training_preference", "flagged", canon,
                               f"conflicting carried prefs: kept '{prefs[canon]}', ignored '{pref}' (from '{raw_name}')")
                    continue
                prefs[canon] = pref
                if canon != raw_name:
                    log_change("leader_training_preference", "edited", raw_name,
                               f"folded into canonical '{canon}'")
        else:
            log_change("leader_training_preference", "dropped", raw_name,
                       "not a canonical leader (junk/retired row)")
    rows = []
    for name in CANONICAL_NAMES:
        pref = prefs.get(name)
        if not pref:
            log_change("leader_training_preference", "flagged", name, "no preference on file")
        rows.append({"Leader": name, "TrainingPreference": pref})
    return pd.DataFrame(rows)


def build_no_training_job_roles(cur: pd.DataFrame) -> pd.DataFrame:
    kept, seen = [], set()
    for v in cur.get("JobRole", pd.Series(dtype=str)):
        role = clean(v)
        if not role:
            log_change("no_training_job_roles", "dropped", str(v), "blank row")
            continue
        if role.lower() in NO_TRAINING_JUNK:
            log_change("no_training_job_roles", "dropped", role, "junk/header row")
            continue
        if role.upper() in seen:
            log_change("no_training_job_roles", "dropped", role, "duplicate")
            continue
        seen.add(role.upper())
        kept.append(role)
    return pd.DataFrame({"JobRole": sorted(kept)})


def build_vendor_aliases(cur: pd.DataFrame) -> pd.DataFrame:
    rows, seen = [], set()
    for _, r in cur.iterrows():
        alias, canon = clean(r.get("Alias")), clean(r.get("CanonicalVendor"))
        if not alias or alias.lower() in ALIAS_JUNK_CANONICALS:
            log_change("vendor_aliases", "dropped", str(alias), "blank/junk alias")
            continue
        if canon and (canon.lower() in ALIAS_JUNK_CANONICALS or canon.lower().startswith("rule:")):
            log_change("vendor_aliases", "dropped", alias, f"junk canonical: '{canon[:60]}'")
            continue
        if alias.upper() in seen:
            log_change("vendor_aliases", "dropped", alias, "duplicate alias")
            continue
        seen.add(alias.upper())
        if not canon:
            log_change("vendor_aliases", "flagged", alias, "blank CanonicalVendor")
        rows.append({"Alias": alias, "CanonicalVendor": canon})
    return pd.DataFrame(rows).sort_values("Alias").reset_index(drop=True)


def build_job_titles_leader_flag(cur: pd.DataFrame) -> pd.DataFrame:
    rows, seen = [], set()
    for _, r in cur.iterrows():
        title, flag = clean(r.get("JobTitle")), clean(r.get("IsLeader"))
        if not title:
            log_change("job_titles_leader_flag", "dropped", str(r.get("JobTitle")), "blank row")
            continue
        if title.upper() in seen:
            log_change("job_titles_leader_flag", "dropped", title, "duplicate")
            continue
        seen.add(title.upper())
        rows.append({"JobTitle": title, "IsLeader": flag or "Yes"})
    return pd.DataFrame(rows).sort_values("JobTitle").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main(apply: bool):
    banner = "LIVE RUN — WILL REWRITE THE WORKBOOK" if apply else "DRY RUN — NOTHING WILL BE WRITTEN"
    print("=" * 72)
    print(f"  build_reference_lists.py   [{banner}]")
    print(f"  Workbook: {REF_PATH}")
    print("=" * 72 + "\n")

    current = load_current_sheets()
    builders = {
        "business_units": build_business_units,
        "job_titles_leader_flag": build_job_titles_leader_flag,
        "no_training_job_roles": build_no_training_job_roles,
        "leaders": build_leaders,
        "leader_training_preference": build_leader_training_preference,
        "vendor_aliases": build_vendor_aliases,
    }

    rebuilt, summary = {}, []
    for sheet in EXPECTED_SHEETS:
        new_df = builders[sheet](current[sheet])
        rebuilt[sheet] = new_df
        n_add = sum(1 for c in changes if c["Sheet"] == sheet and c["Action"] == "added")
        n_drop = sum(1 for c in changes if c["Sheet"] == sheet and c["Action"] == "dropped")
        n_flag = sum(1 for c in changes if c["Sheet"] == sheet and c["Action"] == "flagged")
        summary.append((sheet, len(current[sheet]), len(new_df), n_add, n_drop, n_flag))
        print(f"  {sheet:<30} {len(current[sheet]):>4} -> {len(new_df):<4}"
              f"  (+{n_add} added, -{n_drop} dropped, {n_flag} flagged)")

    pd.DataFrame(changes).to_csv(CHANGES_CSV, index=False)
    print(f"\n  Change list ({len(changes)} rows): {CHANGES_CSV}")

    if not apply:
        print("\nDRY RUN — no backup made, nothing written. Re-run with --apply.")
        return

    backup_dir = op.month_subdir(ARCHIVE_DIR)
    backup = backup_dir / f"{op.dated_stem('wave_reference_lists')}.xlsx"
    shutil.copy(REF_PATH, backup)
    print(f"\n  Backup: {backup}")

    extra_sheets = {s: df for s, df in current.items() if s not in EXPECTED_SHEETS}
    with pd.ExcelWriter(REF_PATH, engine="openpyxl") as writer:
        for sheet in EXPECTED_SHEETS:
            rebuilt[sheet].to_excel(writer, sheet_name=sheet, index=False)
        for sheet, df in extra_sheets.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    print(f"  Written: {REF_PATH.name} ({len(EXPECTED_SHEETS)} rebuilt"
          f"{' + ' + str(len(extra_sheets)) + ' preserved' if extra_sheets else ''} sheets)")

    for sheet, before, after, n_add, n_drop, n_flag in summary:
        runlog.append(RUNLOG_TAB, RUNLOG_HEADERS, [
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "apply", sheet, before, after, n_add, n_drop, n_flag,
            f"backup={backup.name}",
        ])
    print("\n  Next: python sql\\refresh.py refs   (reload raw.ref_* from the new sheets)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true",
                    help="Back up the workbook and write the rebuilt sheets. Default: dry run.")
    with track_run("build_reference_lists.py"):
        main(apply=ap.parse_args().apply)
