# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
fill_master_participation_defaults.py — default the Master's two participation
flags to "No" wherever nobody has designated the person.

    FEC Participant? (Yes/No)        blank -> No
    Soft Live Participant (Y/N)      blank -> No

Her rule, 2026-09-02: "If they are not currently designated they should have no
in that field." This REPLACES the older reading (recorded while building the LAVA
list) that a blank meant "not captured, never No" — blank now means No, and only
an explicit Yes marks a participant.

BLANKS ONLY. A cell that already holds anything is left exactly as it is, so the
62 designations that exist today — and every hand-entered answer after them —
survive untouched. Re-running is a no-op once the column is full.

Nothing else on the pipeline writes these two columns; they are hand-maintained,
with wave change request forms as the other source of truth. That matters for
ordering: once the Master reads "No" everywhere, a LATER "Yes" arriving on a
change request must still win downstream. build_lava_list.py and
build_soft_live_list.py were updated the same day so an explicit Yes from the
change requests outranks a defaulted No here. (At the time of the fill all 49
change-request values already sat on the Master, so nothing was masked.)

Written with one bulk range write per column rather than cell-by-cell: this
touches ~17,000 cells, and a COM round-trip each would take many minutes.

Rules honored: xlwings/COM only (never openpyxl on the Master), Master must be
closed, columns located by header name, existing values never overwritten.

Usage:
    python scripts\\fill_master_participation_defaults.py            # preview
    python scripts\\fill_master_participation_defaults.py --apply    # write
"""

from __future__ import annotations

import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

import activity_log          # noqa: E402
import master_backup         # noqa: E402
import onedrive_paths as op  # noqa: E402

TABLE_NAME = "Table1"
DATA_SHEET = "DATA"
DEFAULT = "No"

COLUMNS = ["FEC Participant? (Yes/No)", "Soft Live Participant (Y/N)"]

# A cell holding any of these is treated as empty. Excel hands back None for a
# truly blank cell, but a stray space or a stringified null reads as data.
EMPTY = {"", "nan", "none", "nat", "null"}


def office_lock_present() -> bool:
    lock = op.MASTER_WAVE_PATH.parent / ("~$" + op.MASTER_WAVE_PATH.name)
    return lock.exists()


def is_blank(v) -> bool:
    return v is None or str(v).strip().lower() in EMPTY


def main(apply: bool) -> None:
    if office_lock_present():
        sys.exit("ERROR: Master Wave File is open in Excel (~$ lock present). "
                 "Close it and re-run.")

    import xlwings as xw

    app = wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = app.screen_updating = False
        wb = app.books.open(str(op.MASTER_WAVE_PATH), update_links=False)
        ws = wb.sheets[DATA_SHEET]
        lo = ws.tables[TABLE_NAME].api

        heads = [str(h).strip() if h is not None else ""
                 for h in lo.HeaderRowRange.Value[0]]
        missing = [c for c in COLUMNS if c not in heads]
        if missing:
            sys.exit(f"ERROR: column(s) not found on DATA: {missing}")

        first_row = lo.Range.Row + 1
        n_rows = lo.Range.Rows.Count - 1
        base_col = lo.Range.Column

        plan = {}
        for header in COLUMNS:
            col = base_col + heads.index(header)
            rng = ws.range((first_row, col), (first_row + n_rows - 1, col))
            cur = rng.value
            cur = list(cur) if isinstance(cur, list) else [cur]
            filled = [DEFAULT if is_blank(v) else v for v in cur]
            n_blank = sum(1 for v in cur if is_blank(v))
            kept = {}
            for v in cur:
                if not is_blank(v):
                    kept[str(v).strip()] = kept.get(str(v).strip(), 0) + 1
            plan[header] = (col, filled, n_blank, kept)
            print(f"{header}")
            print(f"   rows {n_rows:,} · blank -> '{DEFAULT}': {n_blank:,} · "
                  f"already set: {n_rows - n_blank:,} {kept}")

        if not apply:
            print("\nDRY RUN — nothing written. Re-run with --apply to write.")
            return

        master_backup.ensure_daily_full_backup(op.MASTER_WAVE_PATH)
        for header, (col, filled, n_blank, _k) in plan.items():
            if not n_blank:
                print(f"  {header}: nothing to fill")
                continue
            ws.range((first_row, col),
                     (first_row + n_rows - 1, col)).value = [[v] for v in filled]
            print(f"  {header}: wrote {n_blank:,} default(s)")
        wb.save()
        print("\nSaved.")
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
    with activity_log.track_run("fill_master_participation_defaults.py"):
        main("--apply" in sys.argv[1:])
