# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_data_dictionary.py — generate sql\\DATA_DICTIONARY.md from the live DB.

One section per schema, one entry per table/view: purpose, row count (tables),
source-file provenance (raw tables), snapshot date range (history tables),
what each view reads (lineage), and the full column list with types.

Descriptions come from the section comments in 3_report_views.sql /
4_dimensions.sql (the comment block right before each CREATE OR ALTER VIEW),
plus generated text for the machine-made objects (raw loads, alias views).

Also SYNCS every object's description into the SQL Server MS_Description
extended property, so SSMS's Properties pane and the report.whats_what
catalog view show the same text. One source of truth: the comment header
above each CREATE OR ALTER VIEW — edit that, never the property.

Runs standalone any time:
    python sql\\build_data_dictionary.py

...and automatically: sql\\auto_refresh.py regenerates it at the end of any
cycle that loaded data or re-applied a view file, so it can't go stale.
The output file is AUTO-GENERATED — don't hand-edit it.
"""

from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path

from sqlalchemy import create_engine, text

sys.path.insert(0, str(Path(__file__).resolve().parent))
import refresh  # CONN / SERVER / DATABASE + loader metadata

SQL_DIR = Path(__file__).resolve().parent
OUT_PATH = SQL_DIR / "DATA_DICTIONARY.md"
VIEW_DEF_FILES = [SQL_DIR / "3_report_views.sql", SQL_DIR / "4_dimensions.sql",
                  SQL_DIR / "5_sandbox.sql"]

# Present schemas in pipeline order, not alphabetical.
SCHEMA_ORDER = ["raw", "history", "report", "dim", "pbi",
                "cornerstone", "epic", "hr", "mvp", "wave", "sandbox"]

# Fixed descriptions for objects no .sql file documents.
HISTORY_DESC = {
    "roster_daily": "One person-level copy of `report.roster` per calendar day "
                    "(same-day rerun replaces). The raw material for all trend "
                    "and day-over-day views.",
    "epic_not_in_hr_daily": "Daily copy of `report.epic_not_in_hr` — lastname13 the "
                            "FirstSeen/NewToday flags on the boss's gap report.",
}


# ----------------------------------------------------------------------
# descriptions from the .sql definition files
# ----------------------------------------------------------------------
_CREATE_RE = re.compile(
    r"CREATE\s+OR\s+ALTER\s+VIEW\s+\[?(\w+)\]?\.\[?(\w+)\]?", re.IGNORECASE)


def load_sql_file_descriptions() -> dict[str, str]:
    """{'schema.view': 'comment text'} — the /* ... */ block immediately before
    each CREATE OR ALTER VIEW in the managed .sql files."""
    descs: dict[str, str] = {}
    for path in VIEW_DEF_FILES:
        if not path.exists():
            continue
        txt = path.read_text(encoding="utf-8", errors="replace")
        last_comment = ""
        pos = 0
        for m in _CREATE_RE.finditer(txt):
            before = txt[pos:m.start()]
            blocks = re.findall(r"/\*(.*?)\*/", before, re.DOTALL)
            if blocks:
                last_comment = blocks[-1]
            key = f"{m.group(1).lower()}.{m.group(2).lower()}"
            descs[key] = _clean_comment(last_comment)
            pos = m.start()
    return descs


def _clean_comment(block: str) -> str:
    """Strip the ---------- ruling and numbering; collapse to one paragraph."""
    lines = []
    for ln in block.splitlines():
        ln = ln.strip().strip("-").strip()
        ln = re.sub(r"^\d+\.\s*", "", ln)
        if ln:
            lines.append(ln)
    return re.sub(r"\s+", " ", " ".join(lines)).replace(" :", ":")[:600]


# ----------------------------------------------------------------------
# database introspection
# ----------------------------------------------------------------------
def fetch_objects(con) -> list[dict]:
    rows = con.execute(text("""
        SELECT s.name AS sch, o.name AS obj,
               CASE o.type WHEN 'U' THEN 'table' ELSE 'view' END AS kind
        FROM sys.objects o
        JOIN sys.schemas s ON s.schema_id = o.schema_id
        WHERE o.type IN ('U', 'V') AND o.is_ms_shipped = 0
        ORDER BY s.name, o.name""")).fetchall()
    return [{"sch": r.sch, "obj": r.obj, "kind": r.kind} for r in rows]


def fetch_columns(con) -> dict[tuple[str, str], list[tuple]]:
    rows = con.execute(text("""
        SELECT s.name AS sch, o.name AS obj, c.column_id, c.name AS col,
               t.name AS typ, c.max_length, c.is_nullable
        FROM sys.columns c
        JOIN sys.objects o ON o.object_id = c.object_id
        JOIN sys.schemas s ON s.schema_id = o.schema_id
        JOIN sys.types t   ON t.user_type_id = c.user_type_id
        WHERE o.type IN ('U', 'V') AND o.is_ms_shipped = 0
        ORDER BY s.name, o.name, c.column_id""")).fetchall()
    out: dict[tuple[str, str], list[tuple]] = {}
    for r in rows:
        out.setdefault((r.sch, r.obj), []).append(
            (r.column_id, r.col, _fmt_type(r.typ, r.max_length), r.is_nullable))
    return out


def _fmt_type(typ: str, max_length) -> str:
    if typ in ("nvarchar", "varchar", "nchar", "char"):
        if max_length in (-1, None):
            return f"{typ}(max)"
        n = max_length // 2 if typ.startswith("n") else max_length
        return f"{typ}({n})"
    return typ


def fetch_row_counts(con) -> dict[tuple[str, str], int]:
    rows = con.execute(text("""
        SELECT s.name AS sch, o.name AS obj, SUM(p.rows) AS n
        FROM sys.objects o
        JOIN sys.schemas s ON s.schema_id = o.schema_id
        JOIN sys.partitions p ON p.object_id = o.object_id AND p.index_id IN (0, 1)
        WHERE o.type = 'U'
        GROUP BY s.name, o.name""")).fetchall()
    return {(r.sch, r.obj): int(r.n or 0) for r in rows}


def fetch_dependencies(con) -> dict[tuple[str, str], list[str]]:
    """view -> the tables/views it reads (direct references)."""
    rows = con.execute(text("""
        SELECT DISTINCT s.name AS sch, o.name AS obj,
               ISNULL(d.referenced_schema_name, 'dbo') AS rsch,
               d.referenced_entity_name AS robj
        FROM sys.sql_expression_dependencies d
        JOIN sys.objects o ON o.object_id = d.referencing_id
        JOIN sys.schemas s ON s.schema_id = o.schema_id
        WHERE o.type = 'V' AND d.referenced_id IS NOT NULL""")).fetchall()
    out: dict[tuple[str, str], list[str]] = {}
    for r in rows:
        out.setdefault((r.sch, r.obj), []).append(f"{r.rsch}.{r.robj}")
    return {k: sorted(set(v)) for k, v in out.items()}


def fetch_raw_provenance(con, tables: list[str]) -> dict[str, tuple[str, str]]:
    """raw table -> (_source_file, _loaded_at) from its top row, when present."""
    out = {}
    for t in tables:
        try:
            r = con.execute(text(
                f"SELECT TOP 1 [_source_file], [_loaded_at] FROM raw.[{t}]")).fetchone()
            if r:
                out[t] = (r[0], r[1])
        except Exception:
            pass  # table without the metadata columns — fine
    return out


def fetch_history_ranges(con, tables: list[str]) -> dict[str, tuple[str, str, int]]:
    out = {}
    for t in tables:
        try:
            r = con.execute(text(
                f"SELECT MIN(SnapshotDate), MAX(SnapshotDate), "
                f"COUNT(DISTINCT SnapshotDate) FROM history.[{t}]")).fetchone()
            if r and r[0]:
                out[t] = (str(r[0]), str(r[1]), int(r[2]))
        except Exception:
            pass
    return out


# ----------------------------------------------------------------------
# description resolution
# ----------------------------------------------------------------------
def describe(sch: str, obj: str, kind: str, sql_descs: dict[str, str],
             prov: dict, hist: dict) -> str:
    key = f"{sch}.{obj}".lower()
    if key in sql_descs and sql_descs[key]:
        return sql_descs[key]
    if sch == "raw":
        if obj.startswith("ref_"):
            return (f"Reference list — the `{obj[4:]}` sheet of "
                    "`data\\references\\wave_reference_lists.xlsx` "
                    "(reload: `python sql\\refresh.py refs`).")
        if (p := prov.get(obj)):
            return f"Faithful text copy of `{p[0]}`, rebuilt on every load."
        return "Faithful text copy of a source export, rebuilt on every load."
    if sch == "history":
        return HISTORY_DESC.get(obj, "Append-only daily snapshot table.")
    # per-source alias views created by refresh.py
    for schema, view, raw_table, stage in refresh.BUILTIN_SOURCE_ALIASES:
        if schema == sch and view == obj:
            return (f"Alias over `raw.{raw_table}` with a `_stage` column "
                    f"('{stage}') — re-created on every refresh.")
    if sch in refresh.SOURCE_SCHEMAS:
        return ("Registry or hand-built view in a per-source schema "
                "(see `SOURCES.md`; hand-built views are auto-backed-up daily).")
    if sch == "sandbox":
        return "Hand-built sandbox view (auto-backed-up daily to user_views_backup.sql)."
    return ""


# ----------------------------------------------------------------------
# MS_Description extended-property sync
# ----------------------------------------------------------------------
def sync_extended_properties(engine, desc_map: dict[tuple[str, str, str], str]) -> int:
    """Push each object's description into its MS_Description extended
    property (add or update). desc_map: (sch, obj, kind) -> text. Best-effort
    per object; returns how many were synced."""
    n = 0
    stmt = text("""
        IF EXISTS (SELECT 1 FROM sys.extended_properties
                   WHERE class = 1 AND major_id = OBJECT_ID(:full)
                     AND minor_id = 0 AND name = 'MS_Description')
            EXEC sp_updateextendedproperty @name = N'MS_Description', @value = :val,
                 @level0type = N'SCHEMA', @level0name = :sch,
                 @level1type = :lvl, @level1name = :obj;
        ELSE
            EXEC sp_addextendedproperty @name = N'MS_Description', @value = :val,
                 @level0type = N'SCHEMA', @level0name = :sch,
                 @level1type = :lvl, @level1name = :obj;
    """)
    for (sch, obj, kind), desc in desc_map.items():
        if not desc:
            continue
        try:
            with engine.begin() as con:
                con.execute(stmt, {"full": f"{sch}.{obj}", "val": desc[:1000],
                                   "sch": sch, "obj": obj,
                                   "lvl": "TABLE" if kind == "table" else "VIEW"})
            n += 1
        except Exception as e:
            print(f"   note: MS_Description sync failed for {sch}.{obj}: {e}")
    return n


# ----------------------------------------------------------------------
# render
# ----------------------------------------------------------------------
def build(engine=None, out_path: Path = OUT_PATH) -> Path:
    own_engine = engine is None
    if own_engine:
        engine = create_engine(refresh.CONN)
    sql_descs = load_sql_file_descriptions()
    with engine.connect() as con:
        objects = fetch_objects(con)
        columns = fetch_columns(con)
        counts = fetch_row_counts(con)
        deps = fetch_dependencies(con)
        raw_tables = [o["obj"] for o in objects if o["sch"] == "raw"]
        hist_tables = [o["obj"] for o in objects if o["sch"] == "history"]
        prov = fetch_raw_provenance(con, raw_tables)
        hist = fetch_history_ranges(con, hist_tables)
    # One description per object — drives the .md, MS_Description sync, and
    # (via the property) the report.whats_what catalog view.
    desc_map = {(o["sch"], o["obj"], o["kind"]):
                describe(o["sch"], o["obj"], o["kind"], sql_descs, prov, hist)
                for o in objects}
    synced = sync_extended_properties(engine, desc_map)

    if own_engine:
        engine.dispose()

    by_schema: dict[str, list[dict]] = {}
    for o in objects:
        by_schema.setdefault(o["sch"], []).append(o)
    schemas = [s for s in SCHEMA_ORDER if s in by_schema]
    schemas += sorted(s for s in by_schema if s not in SCHEMA_ORDER)

    n_tables = sum(1 for o in objects if o["kind"] == "table")
    n_views = len(objects) - n_tables

    L: list[str] = []
    L.append("# AnalyticsDB — Data Dictionary")
    L.append("")
    L.append(f"**AUTO-GENERATED** by `sql\\build_data_dictionary.py` on "
             f"{datetime.now():%Y-%m-%d %H:%M} — do not hand-edit. It regenerates "
             "automatically whenever the scheduled auto-refresh loads data or "
             "re-applies a view file; run `python sql\\build_data_dictionary.py` "
             "to rebuild on demand.")
    L.append("")
    L.append(f"`{refresh.SERVER}` · `{refresh.DATABASE}` · "
             f"**{n_tables} tables** · **{n_views} views** · {len(schemas)} schemas")
    L.append("")
    L.append("| Schema | Objects | What lives there |")
    L.append("|---|---|---|")
    schema_blurb = {
        "raw": "faithful text copies of the source files (rebuilt every load)",
        "history": "append-only daily snapshots (trend raw material)",
        "report": "the clean, joined views everything reads",
        "dim": "mapping/dimension views (canonical vendors, BUs, titles…)",
        "pbi": "the Power BI surface (thin aliases)",
        "cornerstone": "per-source schema (aliases + your own views)",
        "epic": "per-source schema (aliases + your own views)",
        "hr": "per-source schema (aliases + your own views)",
        "mvp": "per-source schema (aliases + your own views)",
        "wave": "per-source schema (aliases + your own views)",
        "sandbox": "your hand-built cross-source views",
    }
    for s in schemas:
        objs = by_schema[s]
        t = sum(1 for o in objs if o["kind"] == "table")
        v = len(objs) - t
        counts_txt = " + ".join(f"{n} {p}{'s' if n != 1 else ''}"
                                for p, n in (("table", t), ("view", v)) if n)
        L.append(f"| [{s}](#{s}) | {counts_txt} | {schema_blurb.get(s, '')} |")
    L.append("")

    for s in schemas:
        L.append("---")
        L.append("")
        L.append(f"## {s}")
        L.append("")
        for o in sorted(by_schema[s], key=lambda x: x["obj"]):
            sch, obj, kind = o["sch"], o["obj"], o["kind"]
            L.append(f"### `{sch}.{obj}`")
            L.append("")
            meta = [kind]
            if kind == "table" and (sch, obj) in counts:
                meta.append(f"{counts[(sch, obj)]:,} rows")
            if sch == "raw" and obj in prov:
                meta.append(f"loaded from `{prov[obj][0]}` at {prov[obj][1]}")
            if sch == "history" and obj in hist:
                lo, hi, days = hist[obj]
                meta.append(f"snapshots {lo} → {hi} ({days} day{'s' if days != 1 else ''})")
            if (sch, obj) in deps:
                meta.append("reads: " + ", ".join(f"`{d}`" for d in deps[(sch, obj)]))
            L.append(" · ".join(meta))
            L.append("")
            d = desc_map.get((sch, obj, kind), "")
            if d:
                L.append(d)
                L.append("")
            cols = columns.get((sch, obj), [])
            L.append(f"<details><summary>{len(cols)} columns</summary>")
            L.append("")
            L.append("| # | Column | Type | Nullable |")
            L.append("|---|---|---|---|")
            for cid, col, typ, nullable in cols:
                L.append(f"| {cid} | `{col}` | {typ} | {'yes' if nullable else ''} |")
            L.append("")
            L.append("</details>")
            L.append("")

    out_path.write_text("\n".join(L) + "\n", encoding="utf-8")
    print(f"Data dictionary: {len(objects)} objects ({n_tables} tables, "
          f"{n_views} views) -> {out_path}  "
          f"({synced} MS_Description properties synced)")
    return out_path


if __name__ == "__main__":
    build()
