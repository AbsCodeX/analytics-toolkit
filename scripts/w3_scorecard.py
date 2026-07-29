# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
w3_scorecard.py — print the Wave 3 leadership scorecard to the terminal.

Read-only (never touches the Master). One screen, screenshot-ready:
totals, the by-leader table, standing no-show people, and people with no
Epic curriculum assigned — all from report.w3_scorecard and friends, so it
always reflects whatever is loaded in SQL right now (run it any time after
a refresh, morning or afternoon).

Usage:
    python scripts\\w3_scorecard.py
"""

from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent

# Wrong-interpreter guard (same as daily_refresh.py): `python` on PATH is often
# the bare Windows Store Python (no sqlalchemy/pandas). Re-invoke under the
# project venv instead of failing on import.
VENV_PY = ROOT / "analytics_env" / "Scripts" / "python.exe"
if VENV_PY.exists() and Path(sys.executable).resolve() != VENV_PY.resolve():
    raise SystemExit(subprocess.call(
        [str(VENV_PY), str(Path(__file__).resolve()), *sys.argv[1:]]))

import pandas as pd                      # noqa: E402

sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

import source_freshness                  # noqa: E402
from refresh import CONN                 # noqa: E402
from sqlalchemy import create_engine     # noqa: E402

W = 86  # print width


def line(char: str = "-") -> None:
    print(char * W)


def main() -> None:
    engine = create_engine(CONN)
    sc = pd.read_sql("SELECT * FROM report.w3_scorecard", engine)
    noshow = pd.read_sql(
        "SELECT UniversalID, FullName, Leader, ClassTitle, "
        "CONVERT(varchar(16), LastNoShowSessionDate, 120) AS SessionDate, "
        "RegistrationAction "
        "FROM report.w3_noshow_status "
        "WHERE Resolution = 'No Show standing' AND IsInScope = 1 "
        "ORDER BY Leader, FullName", engine)
    rereg = pd.read_sql(
        "SELECT UniversalID, FullName, Leader, ClassTitle, "
        "CONVERT(varchar(16), LastNoShowSessionDate, 120) AS NoShowedOn, "
        "CONVERT(varchar(16), CurrentSessionDate, 120) AS NewSession, "
        "RegistrationAction "
        "FROM report.w3_noshow_status "
        "WHERE Resolution <> 'No Show standing' AND IsInScope = 1 "
        "ORDER BY Leader, FullName", engine)
    nocurr = pd.read_sql(
        "SELECT UniversalID, FullName, Leader, JobRole1 "
        "FROM report.w3_no_epic_curriculum ORDER BY Leader, FullName", engine)
    sources = pd.read_sql(
        "SELECT 'Cornerstone Enterprise Training Report' AS Feed, MAX(_source_file) AS SourceFile, MAX(_loaded_at) AS LoadedAt FROM raw.cornerstone "
        "UNION ALL SELECT 'Epic Curriculum Status Detail', MAX(_source_file), MAX(_loaded_at) FROM raw.epic_status "
        "UNION ALL SELECT 'Epic Team Member Lookup', MAX(_source_file), MAX(_loaded_at) FROM raw.epic_lookup "
        "UNION ALL SELECT 'Master Wave File', MAX(_source_file), MAX(_loaded_at) FROM raw.master "
        "UNION ALL SELECT 'HR (cleaned)', MAX(_source_file), MAX(_loaded_at) FROM raw.hr_cleaned "
        "UNION ALL SELECT 'MVP User Mappings', MAX(_source_file), MAX(_loaded_at) FROM raw.mvp", engine)
    engine.dispose()

    total = sc[sc.IsTotal == 1].iloc[0]
    leaders = sc[sc.IsTotal == 0].sort_values("W3Users", ascending=False)

    print()
    line("=")
    print(f"  WAVE 3 SCORECARD — {datetime.now():%A, %B %d, %Y %I:%M %p}  (leader-scoped)")
    line("=")

    stale = source_freshness.check(dict(zip(sources.Feed, sources.SourceFile)))
    if stale:
        print()
        for w_ in stale:
            print(f"  !! STALE DATA WARNING: {w_}")

    print(f"""
  W3 users                  {total.W3Users:>5}
  Fully registered          {total.FullyRegistered:>5}   ({total.PctRegistered}%)
  Fully trained             {total.FullyTrained:>5}   ({total.PctTrained}%)
  Unregistered              {total.UnregisteredUsers:>5} users / {total.UnregisteredSessions} unique sessions
  Standing no-shows         {total.NoShowUsers:>5} users / {total.NoShowSessions} sessions
  No-showed to date         {total.NoShowUsersToDate:>5} users / {total.NoShowSessionsToDate} no-show records ever
  No Epic curriculum        {total.NoEpicCurriculum:>5} users  (cannot register until assigned)
""")

    line()
    print("  BY LEADER")
    line()
    hdr = f"  {'Leader':<22}{'Users':>6}{'Reg':>6}{'Reg%':>7}{'Trained':>9}{'Unreg':>7}{'NoShow':>8}{'NoCurr':>8}"
    print(hdr)
    print("  " + "-" * (len(hdr) - 2))
    for _, r in leaders.iterrows():
        print(f"  {r.Leader:<22}{r.W3Users:>6}{r.FullyRegistered:>6}"
              f"{str(r.PctRegistered) + '%':>7}{r.FullyTrained:>9}"
              f"{r.UnregisteredUsers:>7}{r.NoShowUsers:>8}{r.NoEpicCurriculum:>8}")

    print()
    line()
    print(f"  STANDING NO-SHOWS ({noshow.UniversalID.nunique()} distinct users / {len(noshow)} total classes)")
    line()
    if noshow.empty:
        print("  none")
    else:
        ns_hdr = (f"  {'Universal ID':<14}{'Name':<26}{'Leader':<22}"
                  f"{'Class':<28}{'Session Date':<18}{'Action'}")
        print(ns_hdr)
        print("  " + "-" * (len(ns_hdr) - 2))
        for _, r in noshow.iterrows():
            print(f"  {(r.UniversalID or '')[:12]:<14}{(r.FullName or '')[:24]:<26}"
                  f"{(r.Leader or '')[:20]:<22}"
                  f"{(r.ClassTitle or '')[:26]:<28}{(r.SessionDate or ''):<18}"
                  f"{r.RegistrationAction or ''}")

    print()
    line()
    print(f"  RE-REGISTERED AFTER NO-SHOW ({rereg.UniversalID.nunique()} distinct users / {len(rereg)} total classes)")
    line()
    if rereg.empty:
        print("  none")
    else:
        rr_hdr = (f"  {'Universal ID':<14}{'Name':<26}{'Leader':<22}"
                  f"{'Class':<28}{'No-Showed On':<18}{'New Session':<18}{'Action'}")
        print(rr_hdr)
        print("  " + "-" * (len(rr_hdr) - 2))
        for _, r in rereg.iterrows():
            print(f"  {(r.UniversalID or '')[:12]:<14}{(r.FullName or '')[:24]:<26}"
                  f"{(r.Leader or '')[:20]:<22}"
                  f"{(r.ClassTitle or '')[:26]:<28}{(r.NoShowedOn or ''):<18}"
                  f"{(r.NewSession or ''):<18}{r.RegistrationAction or ''}")

    print()
    line()
    print(f"  NO EPIC CURRICULUM ASSIGNED ({len(nocurr)} users — blocked from registering)")
    line()
    if nocurr.empty:
        print("  none")
    else:
        nc_hdr = f"  {'Universal ID':<14}{'Name':<26}{'Leader':<22}{'Job Role'}"
        print(nc_hdr)
        print("  " + "-" * (len(nc_hdr) - 2))
        for _, r in nocurr.iterrows():
            print(f"  {(r.UniversalID or '')[:12]:<14}{(r.FullName or '')[:24]:<26}"
                  f"{(r.Leader or '')[:20]:<22}{r.JobRole1 or ''}")

    # short auto-summary in the leadership style
    gap = leaders.sort_values("UnregisteredUsers", ascending=False).iloc[0]
    top_tr = leaders.sort_values("FullyTrained", ascending=False).iloc[0]
    print()
    line("=")
    print("  SUMMARY")
    line("=")
    print(f"""  Wave 3 registration stands at {total.PctRegistered}% ({total.FullyRegistered} of {total.W3Users}).
  Largest remaining gap: {gap.Leader} ({gap.UnregisteredUsers} of {total.UnregisteredUsers} unregistered users).
  Training completions at {total.FullyTrained}, led by {top_tr.Leader} ({top_tr.FullyTrained}).
  No-shows: {total.NoShowUsers} users with {total.NoShowSessions} unresolved sessions
  ({total.NoShowUsersToDate} users / {total.NoShowSessionsToDate} no-show records to date).
  {total.NoEpicCurriculum} users still need an Epic curriculum before they can be booked.
""")

    line()
    print("  DATA SOURCES FOR THIS RUN  (export file -> loaded into SQL)")
    line()
    for _, r in sources.iterrows():
        loaded = "" if pd.isna(r.LoadedAt) else str(r.LoadedAt)[:16]
        print(f"  {r.Feed:<40}{(r.SourceFile or '')[:52]:<54}loaded {loaded}")

    # reference notes — for the analyst, crop these out of leadership screenshots
    print()
    line()
    print("  NOTES (internal reference — crop before sending)")
    line()
    print("""  Cornerstone is the system of record for training/registration (Epic feeds
  from it). Epic's unique contribution is capturing UNREGISTERED users, which
  Cornerstone has no filter for. Session counts are always unique classes per
  user. Standing no-show = unresolved, no newer booking on the transcript.
  W3 classes began 2026-07-13; earlier no-show rows are excluded as data errors.

  HOW THE NO-SHOW ACTION LABEL IS DETERMINED
  For every standing no-show, the exact person + class is looked up in Epic's
  Curriculum Status Detail export:

    What Epic's export shows for that person+class          Action label
    ----------------------------------------------          ------------
    Row exists, Registered = No                             Needs to be Re-Registered
    Row exists, Registered = Yes                            Re-Registered
    No row for that class, but the person has               Class Not Required - Review
      other required classes in Epic
    No rows for the person at all                           No Epic Curriculum - Review

  So "Class Not Required" means: the person no-showed a class that Epic no
  longer lists among their requirements, while Epic still tracks them for
  other classes. The usual cause is a curriculum remap after the no-show —
  the person's Epic job category changed, and the new category requires
  different classes. That's exactly what happened to Unger and Altner: they
  no-showed PB Claims / PB Insurance Follow-Up, then were remapped to
  "PB Coding - Vendor (Onshore)," which requires a different class set — so
  Epic dropped the old classes from their plates.

  Why it says "Review" instead of something definitive: the label trusts
  Epic's export, and we've already caught Epic disagreeing with itself once —
  SBOWEN6's ED Nurse class is gone from Epic's curriculum export, yet Epic's
  no-show scorecard still says she needs re-registering. So the honest read
  is "Epic says this class isn't required anymore — a human should confirm
  whether that's a real curriculum change (do nothing, or book their new
  classes instead) or an Epic data error (rebook the original class)." The
  registration team resolves that question; the label just routes it to them
  instead of guessing.""")
    print()


if __name__ == "__main__":
    main()
