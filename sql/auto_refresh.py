# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
auto_refresh.py  —  change-aware, SCHEDULABLE wrapper around sql\\refresh.py.

Run by the 'YourOrgSQLAutoRefresh' Windows scheduled task (weekdays at
9:00 / 11:00 / 13:00 / 14:30) — and safe to run by hand any time:

    python sql\\auto_refresh.py            # load only sources whose files changed
    python sql\\auto_refresh.py --force    # ignore state, reload everything

What it does each cycle:
  1. Skips the whole cycle if daily_refresh.py is mid-run (its .running lock).
  2. Compares each source's newest file(s) (same discovery logic refresh.py
     uses) against the last successfully-loaded state in
     data\\runlogs\\daily_refresh\\auto_refresh_state.json.
  3. Loads ONLY the changed sources by calling refresh.py's own loaders
     in-process. A file modified less than STABLE_SECONDS ago is WAITED ON
     (re-stat every STABLE_POLL s, up to STABLE_WAIT_MAX s) until it has
     settled, then loaded in the same cycle; only a file still changing after
     that is deferred to the next cycle (it may still be syncing through
     OneDrive). State is updated per source only after a SUCCESSFUL load, so
     locked/failed files retry automatically next cycle.
  4. If sql\\3_report_views.sql or sql\\4_dimensions.sql changed on disk,
     re-applies them via sqlcmd (before the snapshot — snapshot reads
     report.roster).
  5. If any source data actually loaded, re-stamps today's history snapshot
     (same-day rows are replaced — the day's snapshot converges to the latest
     state of the day) and reloads raw.activity_log / raw.metrics_summary.
  6. Logging: one row in activity_log.xlsx ONLY when something loaded /
     applied / failed; quiet no-change cycles just append a line to
     data\\runlogs\\daily_refresh\\auto_refresh_<YYYY-MM>.log.

Notes:
  * 'runlogs' is deliberately never a change TRIGGER — activity_log.xlsx
    changes every time any script (including this one) runs, which would
    self-trigger forever. It piggybacks on real loads instead.
  * Exit code: 0 = clean (including no-op), 1 = one or more failures
    (visible in Task Scheduler history), 2 = could not reach SQL at all.
  * Stage-2 of daily_refresh.py (Master writers) is NOT part of this and must
    never be scheduled.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine

sys.path.insert(0, str(Path(__file__).resolve().parent))
import refresh  # the daily loader — we call its loaders in-process

op = refresh.op  # scripts\onedrive_paths (already imported by refresh)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import activity_log  # noqa: E402  (one-row-per-run log; used conditionally)

SQL_DIR = Path(__file__).resolve().parent
LOCK_DIR = op.RUNLOGS_DIR / "daily_refresh"
STATE_PATH = LOCK_DIR / "auto_refresh_state.json"
DAILY_REFRESH_LOCK = LOCK_DIR / ".running"          # daily_refresh.py's lock
OWN_LOCK = LOCK_DIR / ".auto_refresh_running"
DAILY_LOCK_STALE_HOURS = 12                         # match daily_refresh.py
OWN_LOCK_STALE_MINUTES = 60
STABLE_SECONDS = 60      # a file younger than this may still be OneDrive-syncing
# 2026-09-04: a young file is no longer skipped outright — the cycle WAITS for
# it to settle (every file >= STABLE_SECONDS old, i.e. its mtime stopped
# moving), up to STABLE_WAIT_MAX seconds, and only then defers. Before this,
# prep's clean_hr step wrote HR_cleaned_<date>.xlsx seconds before sql_refresh
# ran, so the fresh cleaned HR was deferred on EVERY HR day (8/31, 9/4) and
# SQL's HR copy stayed one cycle behind until the 13:00 task — Morning
# Review's Master vs Sources / Exceptions sheets and the boss HR gap report
# were built on the previous HR. A finished file that is merely young settles
# in <= STABLE_SECONDS; a file still being written keeps refreshing its mtime
# and is deferred as before.
STABLE_WAIT_MAX = 150    # longest one cycle waits for a young file to settle
STABLE_POLL = 10         # seconds between re-stats while waiting

# View definition files re-applied automatically when edited (mtime change).
VIEW_FILES = [SQL_DIR / "3_report_views.sql", SQL_DIR / "4_dimensions.sql"]

_EPIC_STATUS_WAVE_RE = re.compile(r"-\s*(W\d)", re.IGNORECASE)


# ----------------------------------------------------------------------
# watch map — resolve the file(s) each loader would read right now,
# mirroring refresh.py's own newest-file discovery per source
# ----------------------------------------------------------------------
def _epic_status_files() -> list[Path]:
    """Newest file per wave by mtime (same rule as refresh.load_epic_status)."""
    best: dict[str, Path] = {}
    for f in op.RAW_EPIC_STATUS_DIR.glob("Curriculum Status Detail by User - W*.xlsx"):
        m = _EPIC_STATUS_WAVE_RE.search(f.name)
        if not m:
            continue
        wave = m.group(1).upper()
        if wave not in best or f.stat().st_mtime > best[wave].stat().st_mtime:
            best[wave] = f
    return [best[w] for w in sorted(best)]


def _registry_files() -> list[Path]:
    """source_registry.xlsx plus the newest match of every Enabled row, so a
    new drop for a registered ad-hoc export triggers the 'sources' loader."""
    files: list[Path] = []
    reg_path = op.SOURCE_REGISTRY_PATH
    if not reg_path.exists():
        return files
    files.append(reg_path)
    try:
        reg = pd.read_excel(reg_path, sheet_name=0, dtype=str).fillna("")
    except Exception:
        return files  # unreadable right now (locked/syncing) — retry next cycle
    for _, row in reg.iterrows():
        if str(row.get("Enabled", "")).strip().lower() != "yes":
            continue
        folder = op.ONEDRIVE_ROOT / "data" / str(row["Folder"]).strip().strip("\\/")
        pattern = str(row.get("FilePattern", "")).strip() or "*.xlsx"
        newest = refresh._newest_by_mtime(folder.glob(pattern))
        if newest is not None:
            files.append(newest)
    return files


def _single(path: Path | None) -> list[Path]:
    return [path] if path is not None and path.exists() else []


# loader name -> callable returning the file(s) it reads today.
# Order matters: loads run in this order (matches refresh.LOADERS).
# 'runlogs' and 'snapshot' are intentionally absent (see module docstring).
WATCH = {
    "master":               lambda: _single(op.MASTER_WAVE_PATH),
    "hr":                   lambda: _single(op.RAW_HR_PATH),
    "hr_cleaned":           lambda: _single(op.latest_hr_cleaned()),
    "mvp":                  lambda: _single(op.RAW_MVP_USER_MAPPINGS),
    "mvp_roles":            lambda: _single(op.RAW_MVP_ROLE_MAPPINGS),
    "mvp_job_categories":   lambda: _single(op.RAW_MVP_JOB_CATEGORIES),
    "cornerstone":          lambda: _single(refresh._newest_by_mtime(
                                op.RAW_CORNERSTONE_DIR.glob("Enterprise_Training_Report*.xlsx"))),
    "epic_status":          _epic_status_files,
    "epic_status_summary":  lambda: _single(op.RAW_EPIC_STATUS_SUMMARY),
    "epic_lookup":          lambda: _single(op.RAW_EPIC_TM_DIR / "Epic Team Member Lookup.xlsx"),
    "epic_class_schedule":  lambda: _single(refresh._newest_by_mtime(
                                op.RAW_EPIC_CLASS_SCHEDULES_DIR.glob("epic_class_schedule*.xlsx"))),
    "wave_change_requests": lambda: sorted(
                                p for p in op.RAW_WAVE_CHANGE_REQUESTS_DIR
                                .glob("Wave Change Request Form*.xlsx")
                                if not p.name.startswith("~$")),
    "refs":                 lambda: _single(op.REF_LISTS_PATH),
    "sources":              _registry_files,
}


# ----------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------
def log_line(msg: str) -> None:
    """Timestamped line to the monthly auto_refresh log + stdout."""
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(line)
    try:
        LOCK_DIR.mkdir(parents=True, exist_ok=True)
        logf = LOCK_DIR / f"auto_refresh_{datetime.now():%Y-%m}.log"
        with open(logf, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass  # logging must never break the run


def signature(paths: list[Path]) -> dict[str, list[float]] | None:
    """{path: [mtime, size]} for the current file set, or None if empty."""
    sig = {}
    for p in paths:
        try:
            st = p.stat()
        except OSError:
            continue
        sig[str(p)] = [st.st_mtime, st.st_size]
    return sig or None


def is_stable(sig: dict[str, list[float]]) -> bool:
    """False if any file was modified < STABLE_SECONDS ago (mid-sync risk)."""
    now = time.time()
    return all(now - mtime >= STABLE_SECONDS for mtime, _size in sig.values())


def settle(resolver) -> tuple[dict[str, list[float]] | None, int]:
    """Wait for a young source to finish landing.

    Re-stats the file set every STABLE_POLL seconds until every file is
    >= STABLE_SECONDS old (a file still being written / synced keeps
    refreshing its mtime, so it never qualifies). Returns (signature, seconds
    waited) once settled, or (None, seconds waited) if it was still changing
    after STABLE_WAIT_MAX — the caller defers it to the next cycle.
    """
    waited = 0
    sig = signature(resolver())
    while sig is not None and not is_stable(sig) and waited < STABLE_WAIT_MAX:
        time.sleep(STABLE_POLL)
        waited += STABLE_POLL
        sig = signature(resolver())
    if sig is None or not is_stable(sig):
        return None, waited
    return sig, waited


def lock_is_fresh(path: Path, max_age_hours: float) -> bool:
    try:
        age_h = (time.time() - path.stat().st_mtime) / 3600
    except OSError:
        return False
    return age_h < max_age_hours


def load_state() -> dict:
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(state: dict) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps(state, indent=1), encoding="utf-8")


def apply_view_file(path: Path) -> bool:
    """Re-run a view definition file via sqlcmd. True on success."""
    r = subprocess.run(
        ["sqlcmd", "-S", refresh.SERVER, "-E", "-C", "-d", refresh.DATABASE,
         "-b", "-i", str(path)],
        capture_output=True, text=True)
    tail = (r.stdout or "").strip().splitlines()
    log_line(f"views    {path.name}: {'applied' if r.returncode == 0 else 'FAILED'}"
             + (f" — {tail[-1]}" if tail else ""))
    if r.returncode != 0 and r.stderr:
        log_line(f"views    {path.name} stderr: {r.stderr.strip()[:300]}")
    return r.returncode == 0


# ----------------------------------------------------------------------
# main
# ----------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="Change-aware scheduled SQL refresh.")
    ap.add_argument("--force", action="store_true",
                    help="ignore saved state and reload every source")
    ap.add_argument("--from-daily", action="store_true",
                    help="invoked by daily_refresh.py (which holds the .running "
                         "lock): don't self-skip on it; wait out any mid-flight "
                         "auto cycle instead")
    ap.add_argument("--ensure-snapshot", action="store_true",
                    help="stamp today's history snapshot even when no source "
                         "files changed (daily prep guarantee)")
    args = ap.parse_args()

    # -- respect the daily_refresh orchestrator ---------------------------
    if args.from_daily:
        waited = 0
        while lock_is_fresh(OWN_LOCK, OWN_LOCK_STALE_MINUTES / 60) and waited < 900:
            if waited == 0:
                log_line("waiting: an auto_refresh cycle is mid-flight")
            time.sleep(15)
            waited += 15
    else:
        if lock_is_fresh(DAILY_REFRESH_LOCK, DAILY_LOCK_STALE_HOURS):
            log_line("skipped: daily_refresh.py is running (.running lock present)")
            return 0
        if lock_is_fresh(OWN_LOCK, OWN_LOCK_STALE_MINUTES / 60):
            log_line("skipped: another auto_refresh is running (.auto_refresh_running)")
            return 0
    LOCK_DIR.mkdir(parents=True, exist_ok=True)
    OWN_LOCK.write_text(f"{datetime.now():%Y-%m-%d %H:%M:%S}", encoding="utf-8")

    try:
        state = load_state()

        # -- what changed? -------------------------------------------------
        to_load: list[tuple[str, dict]] = []
        deferred: list[str] = []
        settled: list[str] = []
        for name, resolver in WATCH.items():
            sig = signature(resolver())
            if sig is None:
                continue  # nothing on disk for this source
            if not args.force and state.get(name) == sig:
                continue  # unchanged
            if not args.force and not is_stable(sig):
                # Young file: wait for it to finish landing instead of skipping
                # the cycle — the pipeline's own outputs (cleaned HR) are
                # always young when prep's sql_refresh runs.
                sig, waited = settle(resolver)
                if sig is None:
                    deferred.append(name)  # still changing after STABLE_WAIT_MAX
                    continue
                settled.append(f"{name} ({waited}s)")
                if state.get(name) == sig:
                    continue  # settled back to what is already loaded
            to_load.append((name, sig))

        view_changes: list[tuple[Path, dict]] = []
        for vf in VIEW_FILES:
            key = f"viewfile:{vf.name}"
            sig = signature([vf])
            if sig is None:
                continue
            if args.force or state.get(key) != sig:
                view_changes.append((vf, sig))

        if settled:
            log_line(f"settled (waited for a young file): {', '.join(settled)}")
        if deferred:
            log_line(f"deferred (still changing after {STABLE_WAIT_MAX}s): "
                     f"{', '.join(deferred)}")
        if not to_load and not view_changes:
            if args.ensure_snapshot:
                try:
                    engine = create_engine(refresh.CONN, fast_executemany=True)
                    latest = pd.read_sql(
                        "SELECT MAX(SnapshotDate) AS d FROM history.roster_daily",
                        engine)["d"].iloc[0]
                    if str(latest or "") < f"{datetime.now():%Y-%m-%d}":
                        refresh.snapshot(engine)
                        log_line("no source changes — stamped today's history snapshot")
                    engine.dispose()
                except Exception as e:
                    log_line(f"FAILED ensure-snapshot: {e}")
                    return 1
            log_line("no changes — nothing to load")
            return 0

        # -- connect -------------------------------------------------------
        try:
            engine = create_engine(refresh.CONN, fast_executemany=True)
            with engine.connect():
                pass
        except Exception as e:
            log_line(f"FAILED: cannot reach SQL Server — {e}")
            activity_log.log_run("auto_refresh.py", "Failed",
                                 f"Cannot reach SQL Server: {e}"[:300])
            return 2

        loaded: list[str] = []
        failed: list[str] = []
        start = datetime.now()
        log_line("changed: " + ", ".join(n for n, _ in to_load)
                 + (f" | view files: {', '.join(p.name for p, _ in view_changes)}"
                    if view_changes else ""))

        # -- load changed sources (state updates only on success) ----------
        for name, sig in to_load:
            try:
                refresh.COMMANDS[name](engine)
                state[name] = sig
                loaded.append(name)
            except Exception as e:
                failed.append(name)
                log_line(f"FAILED  {name}: {e}")

        # -- re-apply edited view files (before snapshot) -------------------
        views_applied: list[str] = []
        for vf, sig in view_changes:
            if apply_view_file(vf):
                state[f"viewfile:{vf.name}"] = sig
                views_applied.append(vf.name)
            else:
                failed.append(vf.name)

        # -- re-bind view metadata whenever a raw table was reloaded --------
        # A raw table can change SHAPE (new export column, renamed Master
        # column) and SELECT * alias views keep the old column list until
        # they're refreshed. refresh.py does this at the end of its own runs;
        # auto_refresh calls the loaders directly, so it must too.
        if loaded:
            try:
                n = refresh.refresh_view_metadata(engine)
                log_line(f"views re-bound: {n}")
            except Exception as e:
                log_line(f"FAILED  view re-bind: {e}")

        # -- snapshot + runlogs piggyback when data actually changed --------
        if loaded:
            try:
                refresh.snapshot(engine)
            except Exception as e:
                failed.append("snapshot")
                log_line(f"FAILED  snapshot: {e}")
            try:
                refresh.load_runlogs(engine)
            except Exception as e:
                log_line(f"note: runlogs reload failed ({e}) — non-fatal")

        # -- keep the data dictionary current (best-effort) -----------------
        if loaded or views_applied:
            try:
                import build_data_dictionary
                build_data_dictionary.build(engine)
                log_line("data dictionary regenerated: DATA_DICTIONARY.md")
            except Exception as e:
                log_line(f"note: data dictionary rebuild failed ({e}) — non-fatal")

        engine.dispose()
        save_state(state)

        secs = (datetime.now() - start).total_seconds()
        summary = (f"Auto-refresh: loaded {', '.join(loaded) if loaded else 'nothing'}"
                   + (f"; views re-applied: {', '.join(p.name for p, _ in view_changes)}"
                      if view_changes else "")
                   + (f"; FAILED: {', '.join(failed)}" if failed else "")
                   + f" ({secs:.0f}s)")
        log_line(summary)
        activity_log.log_run("auto_refresh.py",
                             "Needs Review" if failed else "Success", summary[:300])
        return 1 if failed else 0
    finally:
        try:
            OWN_LOCK.unlink(missing_ok=True)
        except OSError:
            pass


if __name__ == "__main__":
    sys.exit(main())
