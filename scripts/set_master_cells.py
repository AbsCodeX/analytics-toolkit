# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
r"""
set_master_cells.py — set ONE column to ONE value for a list of Universal IDs on
the Master Wave File DATA sheet, with a dated note appended to the person's
"Users Wave Change Log" cell.

Built 2026-09-08 for the boss review: 12 Lastname05 HIM scanning staff carried
Training Needed = No but are registered for HIM classes; she asked for the flag
to be flipped to Yes. Kept generic so the next hand-edit does not need a new
Master writer.

    python scripts\set_master_cells.py --column "Training Needed (Yes/No)" --value Yes ^
        --uids CGRAFF DEMOSS ... --note "Training Needed No -> Yes (boss review)"
    ... add --apply to write (preview otherwise)

Rules honored: xlwings/COM only (never openpyxl on the Master); Master must be
closed (~$ lock check); columns located by header name inside Table1; daily
full backup + DATA snapshot around the write; activity log entry.
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
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
UID_HEADER = "UniversalID"
LOG_HEADER = "Users Wave Change Log"


def office_lock_present() -> bool:
    return (op.MASTER_WAVE_PATH.parent / ("~$" + op.MASTER_WAVE_PATH.name)).exists()


def main() -> None:
    ap = argparse.ArgumentParser(description="Set one Master DATA column for a list of UIDs.")
    ap.add_argument("--column", required=True, help="exact DATA header, e.g. 'Training Needed (Yes/No)'")
    ap.add_argument("--value", required=True)
    ap.add_argument("--uids", nargs="+", required=True)
    ap.add_argument("--note", default="", help="appended to 'Users Wave Change Log' as 'MM/DD/YY - <note>'")
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    if office_lock_present():
        sys.exit("ERROR: Master Wave File is open in Excel (~$ lock present). Close it and re-run.")

    want = {u.strip().upper() for u in args.uids}
    import xlwings as xw

    app = wb = None
    changed = 0
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = app.screen_updating = False
        wb = app.books.open(str(op.MASTER_WAVE_PATH), update_links=False)
        ws = wb.sheets[DATA_SHEET]
        lo = ws.tables[TABLE_NAME].api
        heads = [str(h).strip() if h is not None else "" for h in lo.HeaderRowRange.Value[0]]
        for h in (UID_HEADER, args.column, LOG_HEADER):
            if h not in heads:
                sys.exit(f"ERROR: column not found on DATA: {h!r}")
        base_col = lo.Range.Column
        first_row = lo.Range.Row + 1
        n_rows = lo.Range.Rows.Count - 1
        uid_col = base_col + heads.index(UID_HEADER)
        tgt_col = base_col + heads.index(args.column)
        log_col = base_col + heads.index(LOG_HEADER)

        uids = ws.range((first_row, uid_col), (first_row + n_rows - 1, uid_col)).value
        rows = {}
        for i, u in enumerate(uids):
            key = str(u).strip().upper() if u is not None else ""
            if key in want:
                rows[key] = first_row + i
        missing = sorted(want - set(rows))
        if missing:
            print(f"NOT on Master (skipped): {', '.join(missing)}")

        stamp = f"{date.today():%m/%d/%y} - {args.note}" if args.note else ""
        print(f"{'UniversalID':14s} {'row':>6s}  {'current':12s} -> {'new':12s}  change")
        plan = []
        for u in sorted(rows):
            r = rows[u]
            cur = ws.range((r, tgt_col)).value
            cur_s = "" if cur is None else str(cur).strip()
            same = cur_s == args.value
            print(f"{u:14s} {r:6d}  {cur_s[:12]:12s} -> {args.value[:12]:12s}  {'(no change)' if same else 'CHANGE'}")
            if not same:
                plan.append((u, r, cur_s))

        if not args.apply:
            print(f"\nDRY RUN — {len(plan)} cell(s) would change. Re-run with --apply to write.")
            return
        if not plan:
            print("\nNothing to write."); return

        master_backup.ensure_daily_full_backup(op.MASTER_WAVE_PATH)
        for u, r, cur_s in plan:
            ws.range((r, tgt_col)).value = args.value
            if stamp:
                old = ws.range((r, log_col)).value
                old_s = "" if old is None else str(old).strip()
                ws.range((r, log_col)).value = (old_s + "\n" if old_s else "") + f"{stamp} [{args.column}: {cur_s or 'blank'} -> {args.value}]"
            changed += 1
        wb.save()
        print(f"\nSaved. {changed} cell(s) set to {args.value!r} in {args.column!r}.")
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            if app is not None:
                app.quit()

    if args.apply and changed:
        snap = master_backup.save_data_snapshot(op.MASTER_WAVE_PATH)
        print(f"DATA snapshot: {snap}")


if __name__ == "__main__":
    with activity_log.track_run("set_master_cells.py"):
        main()
