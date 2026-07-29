# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
query.py  —  run a quick SQL pull against AnalyticsDB.

Print results to the screen, or export them to CSV/Excel for sharing.
Reuses the same connection as refresh.py (Windows auth, .\\SQLEXPRESS).

Examples:
    # print the first 50 rows
    python sql\\query.py "SELECT TOP 100 * FROM report.roster WHERE Wave='Wave 3'"

    # run a saved .sql file (one SELECT, no GO batches)
    python sql\\query.py -f sql\\my_pull.sql

    # export to Excel or CSV
    python sql\\query.py "SELECT * FROM report.roster WHERE IsInScope=1" --excel wave_roster.xlsx
    python sql\\query.py "SELECT * FROM report.leader_summary" --csv leaders.csv

Notes:
  - Give it ONE statement (a SELECT). For multi-batch scripts with GO, use
    sqlcmd:  sqlcmd -S ".\\SQLEXPRESS" -E -C -d AnalyticsDB -i script.sql
  - --csv / --excel with a bare filename saves into the shared OneDrive pulls
    folder (data\\reports\\sql_pulls\\). Give a full/relative path to override.
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

# Reuse refresh.py's connection string (single source of truth).
sys.path.insert(0, str(Path(__file__).resolve().parent))
from refresh import CONN, op                   # noqa: E402
from sqlalchemy import create_engine, text     # noqa: E402
import pandas as pd                            # noqa: E402


def out_path(arg: str) -> Path:
    """Bare filenames land in the shared OneDrive sql_pulls folder;
    anything with a directory in it is used as given."""
    p = Path(arg)
    if p.parent == Path("."):
        op.SQL_PULLS_DIR.mkdir(parents=True, exist_ok=True)
        return op.SQL_PULLS_DIR / p.name
    return p


def main() -> None:
    ap = argparse.ArgumentParser(description="Run a SQL pull against AnalyticsDB.")
    ap.add_argument("query", nargs="?", help="SQL text to run (quote the whole thing).")
    ap.add_argument("-f", "--file", help="Path to a .sql file to run instead of inline text.")
    ap.add_argument("--csv", help="Write results to this .csv instead of printing.")
    ap.add_argument("--excel", help="Write results to this .xlsx instead of printing.")
    ap.add_argument("-n", "--rows", type=int, default=50, help="Max rows to print (default 50).")
    args = ap.parse_args()

    if args.file:
        sql = Path(args.file).read_text(encoding="utf-8")
    elif args.query:
        sql = args.query
    else:
        ap.error("Provide a SQL string in quotes, or -f path\\to\\file.sql")

    engine = create_engine(CONN)
    try:
        df = pd.read_sql(text(sql), engine)
    finally:
        engine.dispose()

    if args.csv:
        dest = out_path(args.csv)
        df.to_csv(dest, index=False)
        print(f"{len(df):,} rows -> {dest}")
    elif args.excel:
        dest = out_path(args.excel)
        df.to_excel(dest, index=False)
        print(f"{len(df):,} rows -> {dest}")
    else:
        with pd.option_context("display.max_rows", args.rows,
                               "display.max_columns", None,
                               "display.width", 220):
            print(df.head(args.rows).to_string(index=False))
        if len(df) > args.rows:
            print(f"\n[showing {args.rows} of {len(df):,} rows — use -n N, --csv, or --excel]")
        else:
            print(f"\n[{len(df):,} rows]")


if __name__ == "__main__":
    main()
