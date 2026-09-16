# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
update_master_email.py — put an Email address on every person in the Master's
DATA table, so the wave file and every roster built from it can reach the end
user (her ask, 2026-09-02: "for wave it is just emails").

    Email   <- MVP UserUPN, falling back to report.hr.Email

MVP IS THE SOURCE, HR ONLY THE FALLBACK. Per the analyst 2026-09-02: HR is used only
to map the leader chain (Manager -> SVP), never as a reporting source — the same
rule that already sends business unit and system-state fields to raw.mvp. UserUPN
is the YourOrg ACCOUNT address (8,503 of 8,508 @yourorg.edu); HR's Email is a
contact field carrying 49 vendor/affiliate domains (flexstaff.org, vendor.example,
health-roi.com), which Epic and LAVA provisioning cannot use. MVP also covers more
people, 8,508 to HR's 7,860.

UserUPN lives ONLY in the MVP CSV; raw.mvp has no such column, so this reads
data\\raw\\mvp\\User MappingsData.csv directly.

Rules honored: xlwings/COM only (never openpyxl on the Master), Master must be
closed, column located by header name and appended as the Table's last column so
every other writer is unaffected. Idempotent.

HAND-TYPED ADDRESSES ARE NEVER OVERWRITTEN. On a re-run, a cell holding an
address that matches neither HR nor MVP for that person is treated as her own
correction, kept as-is, and reported. Blanks and source-matching cells refresh
normally.

Usage:
    python scripts\\update_master_email.py            # preview only
    python scripts\\update_master_email.py --apply    # write
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

import activity_log                    # noqa: E402
import master_backup                   # noqa: E402
import onedrive_paths as op            # noqa: E402
from email_lookup import email_sources  # noqa: E402
from refresh import CONN               # noqa: E402
from sqlalchemy import create_engine    # noqa: E402

TABLE_NAME = "Table1"
DATA_SHEET = "DATA"
HEADER = "Email"

# Empty-load guard. Both feeds are big and stable; a run that resolves fewer
# addresses than this means a source failed to load, not that people lost their
# email. Floor is ~90% of the 8,516 that resolve today.
MIN_EMAILS = 7_600
MAX_DROP_PCT = 0.10

PREVIEW_CSV = op.LOCAL_MAIN_REPORTS_DIR / "master_email_preview.csv"


def load_emails() -> pd.DataFrame:
    """UniversalID -> the address to write, plus both sources for the audit.
    Resolution lives in email_lookup so the Master, Team File, Tracker, LAVA and
    Soft Live list can never disagree about a person's address."""
    df = email_sources(create_engine(CONN))
    both = df[df["HR"].notna() & df["MVP"].notna()]
    differ = both[~both["HR"].str.lower().eq(both["MVP"].str.lower())]
    print(f"HR emails: {df['HR'].notna().sum():,} | MVP UserUPN: {df['MVP'].notna().sum():,}"
          f" | resolved: {df['Email'].notna().sum():,}")
    print(f"both present: {len(both):,} | differ: {len(differ):,} (HR wins — name changes)")
    return df


def office_lock_present() -> bool:
    lock = op.MASTER_WAVE_PATH.parent / ("~$" + op.MASTER_WAVE_PATH.name)
    return lock.exists()


def main(apply: bool) -> None:
    if office_lock_present():
        sys.exit("ERROR: Master Wave File is open in Excel (~$ lock present). "
                 "Close it and re-run.")

    src = load_emails()
    resolved = int(src["Email"].notna().sum())
    if resolved < MIN_EMAILS:
        sys.exit(f"ABORT: only {resolved:,} address(es) resolved (floor {MIN_EMAILS:,}). "
                 f"Refusing to write — check that HR and the MVP User Mappings CSV "
                 f"both loaded.")

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

        hs = headers()
        if HEADER in hs:
            col_pos = hs.index(HEADER) + 1
        elif apply:
            lo.ListColumns.Add()
            col_pos = lo.ListColumns.Count
            lo.ListColumns(col_pos).Name = HEADER
            print(f"  added column '{HEADER}' (table col {col_pos})")
        else:
            col_pos = None
            print(f"  would add column '{HEADER}' (table col {lo.ListColumns.Count + 1})")

        first_row = lo.Range.Row + 1
        n_rows = lo.Range.Rows.Count - 1
        uid_col = lo.Range.Column + headers().index("UniversalID")
        uids = [str(u).strip().upper() if u is not None else ""
                for u in ws.range((first_row, uid_col),
                                  (first_row + n_rows - 1, uid_col)).value]

        existing = []
        if col_pos is not None:
            sheet_col = lo.Range.Column + col_pos - 1
            existing = ws.range((first_row, sheet_col),
                                (first_row + n_rows - 1, sheet_col)).value
            existing = [str(v).strip() if v not in (None, "") else ""
                        for v in (existing or [])]
        if len(existing) != n_rows:
            existing = [""] * n_rows

        lookup = src.to_dict("index")

        def source_of(row) -> set:
            out = set()
            for k in ("HR", "MVP"):
                v = row.get(k)
                if v is not None and not pd.isna(v):
                    out.add(str(v).strip().lower())
            return out

        values, kept, filled = [], [], 0
        for uid, cur in zip(uids, existing):
            row = lookup.get(uid) or {}
            new = row.get("Email")
            new = None if (new is None or pd.isna(new)) else str(new).strip()
            # Her rule: hand-typed cells are authoritative. A non-blank cell that
            # matches neither source is her own correction — keep it, report it.
            if cur and cur.lower() not in source_of(row):
                kept.append((uid, cur, new or ""))
                values.append(cur)
                filled += 1
                continue
            values.append(new)
            filled += new is not None

        prev_filled = sum(1 for v in existing if v)
        blank = [u for u, v in zip(uids, values) if not v]
        print(f"\nMaster DATA rows: {n_rows:,}")
        print(f"  Email filled     : {filled:,}  ({filled / n_rows:.1%})")
        print(f"  Email blank      : {len(blank):,}")
        print(f"  hand-typed kept  : {len(kept):,}")
        if kept:
            print("\n  manual overrides preserved (UniversalID, kept, source says):")
            for uid, cur, new in kept[:15]:
                print(f"    {uid:<14} {cur:<34} {new}")

        if prev_filled and filled < prev_filled * (1 - MAX_DROP_PCT):
            sys.exit(f"ABORT: Email would fall from {prev_filled:,} to {filled:,} filled "
                     f"rows (>{MAX_DROP_PCT:.0%} drop). Nothing written.")
        if prev_filled:
            print(f"  guard: Email {prev_filled:,} -> {filled:,} filled")

        manual = {k[0] for k in kept}
        audit = pd.DataFrame({"UniversalID": uids, "Email": values})
        audit["HR Email"] = audit["UniversalID"].map(src["HR"])
        audit["MVP UserUPN"] = audit["UniversalID"].map(src["MVP"])
        # Label by which source actually SUPPLIED the address, not by which
        # happens to match it — MVP wins outright, so a value equal to HR's is
        # still MVP-sourced whenever MVP has one.
        audit["Source"] = [
            "" if not v else
            ("manual" if u in manual else ("MVP" if isinstance(m, str) and m else "HR"))
            for u, v, m in zip(audit["UniversalID"], audit["Email"], audit["MVP UserUPN"])]
        PREVIEW_CSV.parent.mkdir(parents=True, exist_ok=True)
        audit.to_csv(PREVIEW_CSV, index=False)
        print(f"\nPreview CSV ({len(audit):,} rows): {PREVIEW_CSV}")
        print(audit["Source"].value_counts().to_string())

        if not apply:
            print("\nDRY RUN — nothing written. Re-run with --apply to write.")
            return

        master_backup.ensure_daily_full_backup(op.MASTER_WAVE_PATH)
        sheet_col = lo.Range.Column + col_pos - 1
        rng = ws.range((first_row, sheet_col), (first_row + n_rows - 1, sheet_col))
        rng.number_format = "@"        # plain text: no autolinking, no reformatting
        rng.value = [[v] for v in values]
        wb.save()
        print(f"\nWrote '{HEADER}' across {n_rows:,} rows; saved.")
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
    with activity_log.track_run("update_master_email.py"):
        main("--apply" in sys.argv[1:])
