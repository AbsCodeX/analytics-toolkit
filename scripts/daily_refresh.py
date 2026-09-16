# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
daily_refresh.py — THE daily pipeline, in two stages.

    python scripts\\daily_refresh.py prep     Stage 1: review inputs, no Master
                                              writes. Clean HR, refresh SQL
                                              (+ history snapshot), preview HR
                                              leader changes, build Morning
                                              Review.xlsx.

    (you review data\\reports\\Main Reports\\Morning Review.xlsx)

    python scripts\\daily_refresh.py apply    Stage 2: the Master writers, with
                                              Excel CLOSED — new-member adds
                                              (Epic Centralized + HR Lastname03
                                              org), MVP, HR — then EVERYTHING
                                              built from the Master, so reports
                                              show today's state: Epic TM
                                              lookup, gap reports, SQL reload +
                                              history re-stamp, boss + no-show/
                                              unregistered reports, metrics,
                                              team file.

    python scripts\\daily_refresh.py all      prep, pause for review, apply.

    python scripts\\daily_refresh.py auto     unattended full run (scheduled
                                              weekday-morning task): prep
                                              (Morning Review deferred) then
                                              apply, no prompt. Apply ends by
                                              rebuilding Morning Review so it
                                              reports the WHOLE run — what ran,
                                              errors, files updated, flags.

    Merged runbook 2026-07-08: this replaces the manual whole-day chain in
    OneDrive data\\.md\\DAILY_SCRIPT_ORDER.md. Gap reports now run AFTER the
    Master writers (post-update semantics: "who is STILL missing").
    Automated full run added 2026-07-28 (the analyst's call): the YourOrgDailyRefresh
    task runs `auto` weekday mornings so everything is refreshed by 10 AM; the
    Morning Review workbook is now the post-run report she checks first.

Other flags:
    --list            show the stage's steps and gates, then exit
    --from NAME       resume a stage from a step (after a failure)
    --only NAME       run a single step
    --skip NAME       skip a step (repeatable), e.g. --skip hr_apply
    --force           ignore gates / the prep-success guard

Stage 2 requires desktop Excel (xlwings + COM steps), so the scheduled `auto`
task must run in the logged-on user's session ("Run only when user is logged
on") — it cannot run as a session-0 background task. Stage 1 is safe to run
any time; it never writes the Master.
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import time
from datetime import date, datetime
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))

# Wrong-interpreter guard: `python` on PATH is often the bare Windows Store
# Python (no sqlalchemy/pandas). Re-invoke under the project venv instead of
# failing at preflight. The venv's pythonw.exe also counts as "in the venv"
# (the scheduled YourOrgDailyRefresh task uses it to run windowless) —
# re-execing it into python.exe would pop a console window.
VENV_PY = ROOT / "analytics_env" / "Scripts" / "python.exe"
if (VENV_PY.exists()
        and Path(sys.executable).resolve().parent != VENV_PY.resolve().parent):
    print(f"(re-running under the project venv: {VENV_PY})", flush=True)
    raise SystemExit(subprocess.call(
        [str(VENV_PY), str(Path(__file__).resolve()), *sys.argv[1:]]))

# Line-buffer stdout so a killed run still shows/keeps its progress. Under
# pythonw.exe (the scheduled windowless run) sys.stdout is None — print()
# becomes a no-op, which is fine because every step's output still goes to
# the daily log file via run_subprocess.
if sys.stdout is not None:
    sys.stdout.reconfigure(line_buffering=True)

import onedrive_paths as op  # noqa: E402
import runlog                # noqa: E402

PY = sys.executable
STATE_DIR = op.RUNLOGS_DIR / "daily_refresh"
LOCK_FILE = STATE_DIR / ".running"
PREP_MARKER = STATE_DIR / "last_prep_success.txt"
LOCK_STALE_HOURS = 12


# ---------------------------------------------------------------------------
# gates — (should_run: bool, reason: str)
# ---------------------------------------------------------------------------
def _raw_hr_date():
    if not op.RAW_HR_PATH.exists():
        return None
    return datetime.fromtimestamp(op.RAW_HR_PATH.stat().st_mtime).date()


def gate_always():
    return True, "daily"


def gate_clean_hr():
    latest = op.latest_hr_cleaned()
    if latest is None:
        return True, "no cleaned HR exists yet"
    raw_d = _raw_hr_date()
    if raw_d is None:
        return False, "raw HR.xlsx missing"
    m = re.search(r"(\d{4})\.(\d{2})\.(\d{2})", latest.name)
    cleaned_d = date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else date.min
    if raw_d > cleaned_d:
        return True, f"raw HR ({raw_d}) newer than cleaned ({cleaned_d})"
    return False, f"cleaned HR ({cleaned_d}) already covers the raw file ({raw_d})"


def gate_weekly_hr():
    if date.today().weekday() == 0:
        return True, "Monday"
    if _raw_hr_date() == date.today():
        return True, "fresh HR file today"
    return False, "weekly report: not Monday, no fresh HR file"


# ---------------------------------------------------------------------------
# inline steps
# ---------------------------------------------------------------------------
def master_locked() -> bool:
    lock = op.MASTER_WAVE_PATH.parent / ("~$" + op.MASTER_WAVE_PATH.name)
    return lock.exists()


def preflight(force: bool) -> bool:
    print(f"Today: {date.today()}  |  python: {PY}")
    for label, p in [("raw HR", op.RAW_HR_PATH),
                     ("raw MVP", op.RAW_MVP_USER_MAPPINGS),
                     ("Epic TM lookup", op.RAW_EPIC_TM_DIR / "Epic Team Member Lookup.xlsx"),
                     ("Master", op.MASTER_WAVE_PATH)]:
        if p.exists():
            m = datetime.fromtimestamp(p.stat().st_mtime)
            print(f"  {label:<16} {m:%Y-%m-%d %H:%M}")
        else:
            print(f"  {label:<16} MISSING: {p}")
    if master_locked():
        print("  NOTE: Master is open in Excel — fine for prep (SQL falls back "
              "to the newest DATA snapshot), but close it before apply.")
    try:
        sys.path.insert(0, str(ROOT / "sql"))
        from refresh import CONN
        from sqlalchemy import create_engine, text
        eng = create_engine(CONN)
        with eng.connect() as con:
            con.execute(text("SELECT 1"))
        eng.dispose()
        print("  SQL Server       reachable")
    except Exception as e:
        print(f"  SQL Server       UNREACHABLE: {e}")
        if not force:
            print("Prep needs SQL (refresh/boss reports/review workbook). "
                  "Start the SQLEXPRESS service or rerun with --force.")
            return False
    return True


def apply_guard(force: bool) -> bool:
    if master_locked():
        print("ERROR: the Master is open in Excel. Close it and rerun apply.")
        return False
    if PREP_MARKER.exists() and PREP_MARKER.read_text().strip() == str(date.today()):
        return True
    if force:
        print("WARNING: no successful prep recorded today — continuing (--force).")
        return True
    print("ERROR: no successful prep run recorded today. Run "
          "'python scripts\\daily_refresh.py prep' (and review) first, or use --force.")
    return False


# ---------------------------------------------------------------------------
# step tables
# ---------------------------------------------------------------------------
def _script(name: str, *args: str) -> list[str]:
    return [PY, str(SCRIPTS / name), *args]


def _sql(*args: str) -> list[str]:
    return [PY, str(ROOT / "sql" / "refresh.py"), *args]


# (name, label, argv-or-None-for-inline, gate)
#
# Ordering (merged runbook, 2026-07-08; trimmed 2026-07-27): report steps run
# in APPLY after the Master writers so everything published reflects TODAY'S
# Master. sql_post must precede boss_reports/tracker/metrics (their views join
# raw.master, and the boss report's NewToday flags need today's history stamp).
# hr_preview must precede hr_apply (it diffs Master vs HR). metrics reads the
# MVP runlog tab + SQL gap views (post-sql_post); team_file republishes the
# Master, so last.
#
# 2026-08-11: SQL updates AFTER the wave file. The Master writers run first,
# then sql_stage loads the finished roster so training_status reads a current
# report.users (same-day coverage for people added this run), then sql_post
# reloads + stamps history from the final workbook. Anything reading raw.master
# must sit after sql_post.
PREP_STEPS = [
    ("preflight",      "Preflight checks",                    None,                                      gate_always),
    ("clean_hr",       "Clean raw HR export",                 _script("clean_hr.py"),                    gate_clean_hr),
    # Change-aware since 2026-07-10: loads only sources whose files changed
    # (auto_refresh state), guarantees today's history snapshot either way.
    ("sql_refresh",    "SQL refresh (changed sources + snapshot)",
                       [PY, str(ROOT / "sql" / "auto_refresh.py"),
                        "--from-daily", "--ensure-snapshot"],                  gate_always),
    ("hr_preview",     "Preview HR leader changes",           _script("review_hr_leader_changes.py"),    gate_always),
    ("morning_review", "Build Morning Review workbook",       _script("build_morning_review.py"),        gate_always),
]

APPLY_STEPS = [
    ("guard",              "Apply-stage guard",               None,                                      gate_always),
    # Pre-rebuild safety copies (added 2026-07-28): yesterday's Tracker +
    # Daily Log get a .bak.<date> in Main Reports\Backups\, and the Master's
    # daily full backup is guaranteed even if no updater writes today.
    # At most one backup per file per day — re-runs are no-ops.
    ("backups",            "Daily safety copies (revert points)",
                           _script("backup_deliverables.py"),                 gate_always),
    # add_members runs FIRST among the writers so mvp_update/hr_apply align the
    # brand-new rows in the same run. Prints the full add-list + review CSV
    # before writing; sanity cap (>150 candidates) aborts the whole apply.
    ("add_members",        "Add missing people to Master",    _script("add_missing_to_master.py", "--apply"), gate_always),
    ("mvp_update",         "Master update from MVP",          _script("update_master_from_mvp.py"),      gate_always),
    ("hr_apply",           "Master update from HR (xlwings)", _script("update_master_from_hr_safe.py"),  gate_always),
    # Email onto the Master (2026-09-02, her ask "for wave it is just emails").
    # After hr_apply because HR is the primary source, and before sql_stage so
    # the column lands in raw.master on the same run. Resolution is shared with
    # the Team File, Tracker, LAVA and Soft Live list via email_lookup.py, so a
    # person's address is identical everywhere. Hand-typed cells are preserved.
    ("email",              "Email onto Master",               _script("update_master_email.py", "--apply"), gate_always),
    # FEC Participant? / Soft Live Participant blanks -> No (her rule 2026-09-09:
    # "those columns should not be blank. if there is no info for that, then put
    # no"). add_members leaves the two cells empty on brand-new rows, so this
    # runs after every Master writer and before sql_stage. Blanks only — a
    # hand-entered Yes/No is never touched. No-op once the columns are full.
    ("participation",      "Default blank FEC / Soft Live to No",
                           _script("fill_master_participation_defaults.py", "--apply"), gate_always),
    # (training_status moved below sql_stage 2026-08-11 — see the note there.)
    # (epic_lookup_update retired 2026-07-06: Epic-mirror columns removed from
    #  the Master — Epic data lives in SQL via epic.raw_team_member_lookup.)
    # (Retired 2026-07-27, per the analyst's pipeline-trim audit — outputs no
    #  longer consumed / superseded: epic_tm_lookup (dated TM Lookup workbook;
    #  SQL + the Master mirror read the raw export), gaps_epic + gaps_wave
    #  (Morning Review + metrics now read the SQL gap views directly),
    #  sql_reports (all-wave no-show/unregistered exports; tracker covers
    #  them), w3_reports (daily W3 package; tracker + Daily Log cover it).
    #  Scripts archived under scripts\_archive\. gaps_hr stays — the weekly
    #  HR gap workbook is still shared.)
    ("gaps_hr",            "Gap report: HR not on wave",      _script("missing_from_wave_hr.py"),        gate_weekly_hr),
    # SQL now updates AFTER the wave file is written (her call 2026-08-11).
    # sql_stage reloads raw.master so report.users — and everything built on it,
    # including report.tracker_training_status — reflects TODAY'S roster, adds
    # from add_members included. training_status then reads a current source, so
    # brand-new people get their status the same day instead of waiting a run.
    # sql_post re-reloads afterwards and stamps history, so raw.master + the
    # daily snapshot both carry the finished workbook. Master-only load, no
    # snapshot — cheap (~4s), and the history stamp belongs on the final state.
    ("sql_stage",          "SQL reload Master (pre training-status)",
                           _sql("master"),                                    gate_always),
    ("training_status",    "Mirror training status onto Master",
                           _script("update_master_training_status.py", "--apply"), gate_always),
    ("sql_post",           "SQL reload Master + re-snapshot", _sql("master", "snapshot"),                gate_always),
    # 2026-08-26: rebuild raw.lava_person (behind report.lava_list) from the
    # finished roster. It reads report.* views, so it cannot run during the raw
    # loads — it belongs here, after sql_post. ~10s. build_lava_list.py reads
    # the same sources, so the workbook and the view always agree.
    ("lava",               "Rebuild LAVA list (report.lava_list)",
                           _sql("lava"),                                      gate_always),
    # 2026-09-03 (her ask): the LAVA workbook is rebuilt daily too, BEFORE the
    # tracker, which copies its four sheets in as snapshots — so the two can
    # never disagree. ~30s; hand-typed review answers carry forward.
    ("lava_workbook",      "Build LAVA workbook",             _script("build_lava_list.py"),             gate_always),
    ("boss_reports",       "Publish boss gap reports",        _script("export_boss_reports.py"),         gate_always),
    # tracker rebuilds from SQL, which sql_post just reloaded — skip its own
    # SQL pass. Aborts cleanly if either tracker workbook is open in Excel.
    ("tracker",            "Rebuild RCM Training Tracker + Daily/Monthly Log",
                           _script("build_rcm_training_tracker.py", "--no-sql"), gate_always),
    ("metrics",            "Append metrics trend row",        _script("build_metrics_summary.py"),       gate_always),
    ("team_file",          "Publish RCM Wave Team File",      _script("build_team_wave_file.py"),        gate_always),
    # wave_file_copy (2026-09-03 → 2026-09-10) retired: she keeps ONE main copy
    # of the wave file and makes ad-hoc copies herself. Script in _archive.
    # Rebuild Morning Review LAST (added 2026-07-28) so it reports the whole
    # run: step results, errors, files updated, and post-apply gap state.
    ("review_refresh",     "Rebuild Morning Review (post-apply report)",
                           _script("build_morning_review.py"),                gate_always),
]

STAGES = {"prep": PREP_STEPS, "apply": APPLY_STEPS}
INLINE = {"preflight": preflight, "guard": apply_guard}


# ---------------------------------------------------------------------------
# runner
# ---------------------------------------------------------------------------
def run_subprocess(argv: list[str], log_fh) -> bool:
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    p = subprocess.Popen(argv, cwd=str(ROOT), stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT, text=True,
                         encoding="utf-8", errors="replace", env=env)
    for line in p.stdout:
        print(line, end="")
        log_fh.write(line)
        log_fh.flush()   # keep the log current even if this run is killed
    p.wait()
    return p.returncode == 0


def run_stage(stage: str, args) -> int:
    steps = STAGES[stage]
    names = [n for n, *_ in steps]

    if args.list:
        print(f"{stage} steps:")
        for name, label, argv, gate in steps:
            kind = "inline" if argv is None else Path(argv[1]).name
            print(f"  {name:<20} {label:<38} [{kind}]")
        return 0

    # An `all` run passes the same flags to both stages, so a step name may
    # belong to the other stage: validate against the union of both stages,
    # then resolve what the flag means for THIS stage below.
    all_names = [n for ss in STAGES.values() for n, *_ in ss]
    for flag, val in [("--from", args.start), ("--only", args.only),
                      *(("--skip", s) for s in args.skip)]:
        if val and val not in all_names:
            print(f"Unknown step for {flag}: {val}\nValid: {', '.join(all_names)}")
            return 2

    if args.only:
        todo = [s for s in steps if s[0] == args.only]
        if not todo:
            print(f"({stage}: '{args.only}' belongs to the other stage — nothing to run)")
            return 0
    elif args.start and args.start not in names:
        if stage == "prep":
            print(f"(prep: '--from {args.start}' is an apply step — skipping prep)")
            return 0
        todo = steps  # resuming apply from a prep-stage step: run apply in full
    else:
        start_i = names.index(args.start) if args.start else 0
        todo = steps[start_i:]

    STATE_DIR.mkdir(parents=True, exist_ok=True)
    if LOCK_FILE.exists():
        age_h = (time.time() - LOCK_FILE.stat().st_mtime) / 3600
        if age_h < LOCK_STALE_HOURS:
            print(f"ERROR: another daily_refresh run appears active "
                  f"({LOCK_FILE}, {age_h:.1f}h old). Delete it if that's wrong.")
            return 1
        print(f"Ignoring stale lockfile ({age_h:.1f}h old).")
    LOCK_FILE.write_text(f"{stage} started {datetime.now():%Y-%m-%d %H:%M:%S}")

    log_dir = op.month_subdir(STATE_DIR)
    log_path = log_dir / f"daily_refresh_{date.today():%Y.%m.%d}_{stage}.log"
    overall = time.time()
    completed, skipped = [], []

    try:
        # buffering=1 (line buffered): step headers and OK/X lines must reach
        # disk immediately — review_refresh reads this log to build Morning
        # Review, and unflushed lines made the run look stuck at the prior step.
        with open(log_path, "a", encoding="utf-8", buffering=1) as log_fh:
            log_fh.write(f"\n===== {stage} run {datetime.now():%Y-%m-%d %H:%M:%S} =====\n")
            for name, label, argv, gate in todo:
                if name in args.skip:
                    print(f"\n----- {name}: SKIPPED (--skip) -----")
                    skipped.append(name)
                    continue
                should_run, reason = (True, "forced") if args.force else gate()
                if not should_run:
                    print(f"\n----- {name}: SKIPPED ({reason}) -----")
                    log_fh.write(f"{name}: skipped ({reason})\n")
                    skipped.append(name)
                    continue

                print(f"\n===== {name}: {label} =====")
                log_fh.write(f"\n===== {name}: {label} =====\n")
                start = time.time()
                if argv is None:
                    ok = INLINE[name](args.force)
                else:
                    ok = run_subprocess(argv, log_fh)
                elapsed = time.time() - start
                mark = "OK " if ok else "X  "
                print(f"{mark}{name} ({elapsed:.1f}s)")
                log_fh.write(f"{mark}{name} ({elapsed:.1f}s)\n")
                if not ok:
                    print(f"\nSTOPPED at '{name}'. Fix the error above, then resume:")
                    print(f"    python scripts\\daily_refresh.py {stage} --from {name}")
                    _log_summary(stage, "Failed", completed, skipped, name)
                    return 1
                completed.append(name)
    finally:
        LOCK_FILE.unlink(missing_ok=True)

    if stage == "prep" and not args.only:
        PREP_MARKER.write_text(str(date.today()))
    _log_summary(stage, "Success", completed, skipped, None)
    print(f"\n{stage.upper()} complete in {time.time() - overall:.1f}s "
          f"({len(completed)} run, {len(skipped)} skipped). Log: {log_path}")
    if stage == "prep":
        print(f"Review: {op.MORNING_REVIEW_PATH}")
        print("Then:   python scripts\\daily_refresh.py apply")
    return 0


def _log_summary(stage, status, completed, skipped, failed_at):
    try:
        runlog.append(
            "Daily Refresh",
            ["Run Timestamp", "Stage", "Status", "Steps Run", "Steps Skipped", "Failed At"],
            [f"{datetime.now():%Y-%m-%d %H:%M:%S}", stage, status,
             ", ".join(completed), ", ".join(skipped), failed_at or ""],
        )
    except Exception as e:
        print(f"(runlog append failed: {e})")


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Run the daily pipeline: prep -> (you review) -> apply.")
    ap.add_argument("stage", choices=["prep", "apply", "all", "auto"])
    ap.add_argument("--list", action="store_true", help="show steps and exit")
    ap.add_argument("--from", dest="start", default=None, metavar="NAME",
                    help="resume the stage from this step")
    ap.add_argument("--only", default=None, metavar="NAME", help="run one step only")
    ap.add_argument("--skip", action="append", default=[], metavar="NAME",
                    help="skip a step (repeatable)")
    ap.add_argument("--force", action="store_true",
                    help="ignore gates and the prep-success guard")
    args = ap.parse_args()

    if args.stage == "auto":
        # Unattended full run (scheduled task). Prep's Morning Review build is
        # deferred — apply's final review_refresh step builds it once, after
        # everything, so it reports the finished run instead of the pre-run
        # state. No prompt between stages.
        args.skip = [*args.skip, "morning_review"]
        rc = run_stage("prep", args)
        if rc != 0 or args.list:
            return rc
        return run_stage("apply", args)

    if args.stage == "all":
        rc = run_stage("prep", args)
        if rc != 0 or args.list:
            return rc
        print(f"\nOpen and review: {op.MORNING_REVIEW_PATH}")
        answer = input("Type APPLY to continue with the Master updates, anything else to stop: ")
        if answer.strip().upper() != "APPLY":
            print("Stopped before apply. Run 'python scripts\\daily_refresh.py apply' when ready.")
            return 0
        return run_stage("apply", args)
    return run_stage(args.stage, args)


if __name__ == "__main__":
    raise SystemExit(main())
