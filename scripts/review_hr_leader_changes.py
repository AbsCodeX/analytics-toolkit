# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
review_hr_leader_changes.py

READ-ONLY preview of the leader changes that update_master_from_hr_safe.py WOULD
make on its next real run. Writes the full list to a CSV in the OneDrive
morning-review folder (data\\reports\\morning_review\\hr_leader_changes_preview.csv,
where build_morning_review.py folds it into Morning Review.xlsx) and prints an
on-screen table so you can verify BEFORE applying (i.e. before running
update_master_from_hr_safe.py with no flag, which applies by default).

HR source:
  - Reads from the newest CLEANED HR file (OneDrive data\\processed\\hr\\
    clean_hr_copies\\HR_cleaned_*.xlsx, "All Staff" tab, via onedrive_paths)
    produced by scripts/clean_hr.py — in which leader-chain Universal IDs have
    already been resolved to full names and leader spellings standardized. This
    preview does NOT reimplement any HR-loading logic; it imports HR_PATH and
    load_hr from update_master_from_hr_safe, so it shares the EXACT same source and
    logic as the real run (single source of truth).

Fidelity:
  - Reads the Master through Excel/xlwings (read_only=True) using the EXACT same
    extent detection and value reads as the real run, so the count here matches
    the real run cell-for-cell. (An earlier openpyxl version disagreed by one row
    because openpyxl reads cached values and a different sheet extent.)
  - Reading never risks the workbook package — only SAVING does, and this tool
    never saves. It refuses to run if the Master is open in Excel.
  - Imports the same logic + helpers from update_master_from_hr_safe so there is
    a single source of truth.

Run this before every real run:
    python scripts\review_hr_leader_changes.py
"""

import csv
import sys
from collections import Counter
from pathlib import Path

import xlwings as xw

# These imports must follow the sys.path.insert above so the sibling scripts/
# modules resolve when this file is run directly; hence the disable.
sys.path.insert(0, str(Path(__file__).parent))
import activity_log  # noqa: E402  pylint: disable=wrong-import-position
import onedrive_paths  # noqa: E402  pylint: disable=wrong-import-position
from leader_names import (  # noqa: E402  pylint: disable=wrong-import-position
    classify_leader_change,
    parse_hr_export_date,
)
from update_master_from_hr_safe import (  # noqa: E402  pylint: disable=wrong-import-position
    HR_PATH,
    LEADER_UID_OVERRIDES,
    MASTER_PATH,
    TARGET_COLS,
    already_open_in_excel,
    clean_leadership,
    clean_uid,
    compute_leaders,
    last_data_row,
    last_header_col,
    load_hr,
    office_lock_present,
    read_col,
)

# Saved with the rest of the morning-review surfaces on OneDrive —
# build_morning_review.py reads it from here into Morning Review.xlsx.
OUT_CSV = onedrive_paths.MORNING_REVIEW_DIR / "hr_leader_changes_preview.csv"

# Optional person-name column on the DATA sheet (first match wins; blank if none).
_NAME_CANDIDATES = ("Name", "Full Name", "Employee Name", "Worker", "Last Name")


def collect_changes():
    """Return (changes, hr_date_short) — the full list the real run would log."""
    if HR_PATH is None:
        sys.exit("ERROR: No cleaned HR file found. Run scripts/clean_hr.py first to generate it.")
    hr, sheet_name = load_hr(HR_PATH)
    hr_date_short = parse_hr_export_date(sheet_name)

    # Refuse to read while the Master is open in Excel — opening read-only could
    # attach to that instance, and the real run requires it closed anyway.
    if office_lock_present(MASTER_PATH) or already_open_in_excel(MASTER_PATH):
        sys.exit("ERROR: The Master is open in Excel. Close it completely and re-run.")

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(MASTER_PATH), update_links=False, read_only=True)
        if "DATA" not in [s.name for s in wb.sheets]:
            sys.exit("ERROR: 'DATA' sheet not found in the Master Wave File.")
        ws = wb.sheets["DATA"]

        # Same extent detection as the real run.
        last_col = last_header_col(ws)
        header_row = ws.range((1, 1), (1, last_col)).value
        if not isinstance(header_row, list):
            header_row = [header_row]
        col = {str(h).strip(): i + 1 for i, h in enumerate(header_row)
               if h is not None and str(h).strip() != ""}

        # Validate against the FULL set the real run requires (TARGET_COLS +
        # UniversalID), not just the columns this preview happens to read. If any
        # of them were missing, the real run would abort with zero changes, so the
        # preview must refuse too rather than showing changes that would never apply.
        required_cols = list(TARGET_COLS) + ["UniversalID"]
        missing_cols = [c for c in required_cols if c not in col]
        if missing_cols:
            sys.exit(f"ERROR: Master DATA sheet missing required column(s): {missing_cols}. "
                      "(The real run would abort too — same requirement.)")

        uid_col = col["UniversalID"]
        leaders_col = col["Leaders"]
        name_col = next((col[c] for c in _NAME_CANDIDATES if c in col), None)

        last_row = last_data_row(ws, uid_col)
        first = 2
        uid_vals = read_col(ws, uid_col, first, last_row)
        leaders_vals = read_col(ws, leaders_col, first, last_row)
        name_vals = (read_col(ws, name_col, first, last_row)
                     if name_col else [None] * len(uid_vals))
        wb.close()
    finally:
        app.quit()

    changes = []
    for r, uid_raw in enumerate(uid_vals):
        uid = clean_uid(uid_raw)
        if not uid:
            continue
        hr_row = hr.get(uid)
        if not hr_row:
            continue  # not in HR -> no leader change

        vp = clean_leadership(hr_row.get("VP"), canonicalize=True)
        avp = clean_leadership(hr_row.get("AVP"), canonicalize=True)
        svp = clean_leadership(hr_row.get("SVP"), canonicalize=True)

        # Blank senior chain -> leadership preserved, no change logged (matches main script).
        if vp is None and avp is None and svp is None:
            continue

        new_leaders = compute_leaders(vp, avp, svp)
        if uid in LEADER_UID_OVERRIDES:
            new_leaders = LEADER_UID_OVERRIDES[uid]

        old_leaders = leaders_vals[r]
        bucket = classify_leader_change(old_leaders, new_leaders)
        if not bucket:
            continue

        name = name_vals[r]
        changes.append({
            "UID": uid,
            "Name": "" if name is None else str(name),
            "Old Leader": "" if old_leaders is None else str(old_leaders),
            "New Leader": "" if new_leaders is None else str(new_leaders),
            "Change Type": bucket,
            "HR VP": vp or "",
            "HR AVP": avp or "",
            "HR SVP": svp or "",
        })

    # Group identical transitions together; FormerLeaderA->Lastname10 etc. land as one block.
    changes.sort(key=lambda c: (c["Old Leader"], c["New Leader"], c["UID"]))
    return changes, hr_date_short


def write_csv(changes):
    """Write the full change list to OUT_CSV (UTF-8 BOM for Excel)."""
    fields = ["UID", "Name", "Old Leader", "New Leader", "Change Type",
              "HR VP", "HR AVP", "HR SVP"]
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(changes)


def print_table(changes, hr_date_short):
    """Print the grouped-transition summary, by-type counts, and full change table."""
    print(f"\nHR export date: {hr_date_short}   |   Total leader changes: {len(changes)}\n")

    # --- Grouped summary: the fastest way to verify a mass reassignment --------
    pairs = Counter((c["Old Leader"] or "(blank)", c["New Leader"] or "(blank)")
                    for c in changes)
    print("Transitions grouped (Old -> New : count):")
    for (old, new), n in sorted(pairs.items(), key=lambda kv: (-kv[1], kv[0])):
        print(f"  {n:>4}   {old}  ->  {new}")

    by_type = Counter(c["Change Type"] for c in changes)
    print("\nBy change type:")
    for t, n in sorted(by_type.items(), key=lambda kv: -kv[1]):
        print(f"  {n:>4}   {t}")

    # --- Full row-by-row table -------------------------------------------------
    w_uid = max(3, max((len(c["UID"]) for c in changes), default=3))
    w_old = max(10, max((len(c["Old Leader"]) for c in changes), default=10))
    w_new = max(10, max((len(c["New Leader"]) for c in changes), default=10))
    w_typ = max(11, max((len(c["Change Type"]) for c in changes), default=11))

    header = (f"{'UID':<{w_uid}}  {'Old Leader':<{w_old}}  "
              f"{'New Leader':<{w_new}}  {'Change Type':<{w_typ}}")
    print(f"\nAll {len(changes)} changes:")
    print(header)
    print("-" * len(header))
    for c in changes:
        print(f"{c['UID']:<{w_uid}}  {c['Old Leader']:<{w_old}}  "
              f"{c['New Leader']:<{w_new}}  {c['Change Type']:<{w_typ}}")


def main():
    """Collect the preview changes, write the CSV, and print the review table."""
    changes, hr_date_short = collect_changes()
    write_csv(changes)
    print_table(changes, hr_date_short)
    print(f"\nCSV written: {OUT_CSV}")
    print("Review the CSV/table above, then apply with:\n"
          "    python scripts/update_master_from_hr_safe.py\n"
          "(this preview is read-only; the update script always applies — no flag needed.)")


if __name__ == "__main__":
    with activity_log.track_run("review_hr_leader_changes.py"):
        main()
