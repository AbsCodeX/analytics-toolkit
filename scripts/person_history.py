# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
person_history.py — full wave / job-role change history for any user or leader.

Scrubs every history source we have for a person and prints a chat-style report:
  1. Current record                 report.users (SQL)
  2. Monthly snapshot history       OneDrive Main Reports/Wave_JobRole_Change_Report.xlsx
                                    (rebuilt monthly by wave_job_role_change_report.py, Dec 2025+)
  3. Daily snapshot diffs           history.roster_daily (SQL, 2026-07-06 onward)
  4. Wave change request forms      raw.wave_change_requests (SQL)
  5. Cornerstone transcript         raw.cornerstone (SQL, full history export)
  6. Epic TM Lookup current view    raw.epic_lookup (SQL)
  7. W3 GNF planning mentions       raw.wave_gnf_w3_assignments / _mapping (SQL)

Examples:
    # one user, full report
    python scripts\\person_history.py cjohnson82

    # several users
    python scripts\\person_history.py cjohnson82 jdoe12 asmith9

    # everyone under a leader — compact change summary (matches on Leaders col)
    python scripts\\person_history.py --leader "Lastname07"

    # leader mode, but full reports for the users that have changes
    python scripts\\person_history.py --leader "Lastname07" --full
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "sql"))
from refresh import CONN                        # noqa: E402
from sqlalchemy import create_engine, text      # noqa: E402
from onedrive_paths import MAIN_REPORTS_DIR     # noqa: E402

CHANGE_REPORT = MAIN_REPORTS_DIR / "Wave_JobRole_Change_Report.xlsx"

# Fields worth diffing across daily snapshots (label -> roster_daily column)
DAILY_FIELDS = {
    "Wave": "Wave",
    "Job Role 1": "JobRole1",
    "Job Role 2": "JobRole2",
    "Job Role 3": "JobRole3",
    "Job Role 4": "JobRole4",
    "Leader": "Leader",
    "User Type": "UserType",
    "Vendor Y/N": "VendorYN",
    "HR Job Title": "HR_JobTitle",
    "Departed": "Departed",
    "Terminated": "TerminatedFlag",
}

LINE = "=" * 72


def norm(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "null") else s


# ---------------------------------------------------------------- SQL pulls
def fetch_df(engine, sql: str, **params) -> pd.DataFrame:
    with engine.connect() as cx:
        return pd.read_sql(text(sql), cx, params=params or None)


def current_record(engine, uid: str) -> pd.Series | None:
    df = fetch_df(engine, """
        SELECT UniversalID, FullName, Leader, AVP, VP, SVP, Wave, UserType,
               VendorYN, JobRole1, JobRole2, JobRole3, JobRole4, TrainingNeeded,
               Departed, TerminatedFlag, OnLeaderList, IsInScope
        FROM report.users WHERE UPPER(UniversalID) = :uid""", uid=uid)
    return df.iloc[0] if len(df) else None


def daily_changes(engine, uid: str):
    """Diff consecutive daily snapshots; return (first_date, last_date, [changes])."""
    df = fetch_df(engine, f"""
        SELECT SnapshotDate, {", ".join(DAILY_FIELDS.values())}
        FROM history.roster_daily WHERE UPPER(UniversalID) = :uid
        ORDER BY SnapshotDate""", uid=uid)
    if df.empty:
        return None, None, []
    changes = []
    prev = df.iloc[0]
    for _, row in df.iloc[1:].iterrows():
        for label, col in DAILY_FIELDS.items():
            a, b = norm(prev[col]), norm(row[col])
            if a != b:
                changes.append((str(row["SnapshotDate"]), label, a or "(blank)", b or "(blank)"))
        prev = row
    return str(df.iloc[0]["SnapshotDate"]), str(df.iloc[-1]["SnapshotDate"]), changes


def wave_requests(engine, uid: str) -> pd.DataFrame:
    return fetch_df(engine, """
        SELECT Add_Edit_Remove_Select_from_the_dropdown AS Action,
               Current_Wave, Requested_Wave_Change,
               Staff_Training_Job_Role_1, Staff_Training_Job_Role_2,
               Direct_Manager, _source_file
        FROM raw.wave_change_requests WHERE UPPER(Universal_ID) = :uid""", uid=uid)


def cornerstone(engine, uid: str) -> pd.DataFrame:
    return fetch_df(engine, """
        SELECT Training_Title, Transcript_Status,
               Transcript_Registration_Date, Transcript_Completed_Date
        FROM raw.cornerstone
        WHERE UPPER(User_ID) = :uid AND Training_Provider = 'EPIC'
        ORDER BY Transcript_Registration_Date""", uid=uid)


def epic_lookup(engine, uid: str) -> pd.Series | None:
    df = fetch_df(engine, """
        SELECT Job_Title, Curriculum_Type, Wave, Epic_Job_Categories_MVP,
               Recommended_Job_Categories, Total_Curriculums_MVP,
               Epic_Curriculums_Cornerstone, Removed_Curriculums_in_MVP,
               Removed_Curriculums, Fully_Registered, Fully_Trained, Direct_Manager
        FROM raw.epic_lookup WHERE UPPER(Universal_ID) = :uid""", uid=uid)
    return df.iloc[0] if len(df) else None


def gnf_mentions(engine, uid: str) -> pd.DataFrame:
    return fetch_df(engine, """
        SELECT GO_AND_FIND, _source_file FROM raw.wave_gnf_w3_assignments
        WHERE UPPER(UID) = :uid""", uid=uid)


# ------------------------------------------- Jan-May monthly change report
def _sheet_with_header(xl: pd.ExcelFile, sheet: str) -> pd.DataFrame | None:
    """The report sheets have 2-3 title rows above the real header."""
    raw = xl.parse(sheet, header=None)
    for i in range(min(8, len(raw))):
        if any("universal" in str(v).lower() for v in raw.iloc[i].tolist()):
            df = xl.parse(sheet, header=i)
            df["_uid_upper"] = df["Universal ID"].astype(str).str.strip().str.upper()
            return df
    return None


class MonthlyHistory:
    """Lazy one-time load of Wave_JobRole_Change_Report.xlsx (Jan-May snapshots)."""

    def __init__(self):
        self.loaded = False
        self.waves = self.roles = self.adds = None

    def load(self):
        if self.loaded:
            return
        self.loaded = True
        if not CHANGE_REPORT.exists():
            return
        xl = pd.ExcelFile(CHANGE_REPORT, engine="calamine")
        self.waves = _sheet_with_header(xl, "Wave Changes by User")
        self.roles = _sheet_with_header(xl, "Job Role Changes")
        self.adds = _sheet_with_header(xl, "New Additions")

    def _row(self, df, uid):
        if df is None:
            return None
        hit = df[df["_uid_upper"] == uid]
        return hit.iloc[0] if len(hit) else None

    def wave_row(self, uid):
        self.load()
        return self._row(self.waves, uid)

    def role_row(self, uid):
        self.load()
        return self._row(self.roles, uid)

    def new_addition_row(self, uid):
        self.load()
        return self._row(self.adds, uid)


MONTHLY = MonthlyHistory()


# ---------------------------------------------------------------- reporting
def print_full_report(engine, uid: str) -> None:
    uid = uid.strip().upper()
    cur = current_record(engine, uid)

    print(f"\n{LINE}")
    if cur is None:
        print(f"{uid} — NOT FOUND in report.users (not on the Master wave file)")
        print(LINE)
    else:
        print(f"{uid} — {norm(cur['FullName'])}")
        print(f"Leader: {norm(cur['Leader'])} | AVP: {norm(cur['AVP'])} | "
              f"VP: {norm(cur['VP'])} | SVP: {norm(cur['SVP'])}")
        flags = []
        if norm(cur["Departed"]):
            flags.append(f"Departed: {norm(cur['Departed'])}")
        if norm(cur["TerminatedFlag"]) not in ("", "0", "False"):
            flags.append("TERMINATED")
        print(f"Current: {norm(cur['Wave']) or '(no wave)'} | "
              f"{norm(cur['UserType']) or '(no user type)'} | "
              f"Vendor: {norm(cur['VendorYN']) or '?'} | "
              f"In scope: {'Yes' if str(cur['IsInScope']) in ('1', 'True') else 'No'}"
              + (" | " + " | ".join(flags) if flags else ""))
        roles = [norm(cur[f"JobRole{i}"]) for i in range(1, 5)]
        print("Job Roles: " + (" | ".join(r for r in roles if r) or "(none)"))
        print(LINE)

    # --- monthly history (rebuilt monthly by wave_job_role_change_report.py)
    print("\nWAVE HISTORY (monthly User Mappings snapshots)")
    wrow = MONTHLY.wave_row(uid)
    if wrow is not None:
        for col in wrow.index:
            if str(col).startswith("Wave (") and norm(wrow[col]):
                print(f"  {col[6:-1]:<10} {norm(wrow[col])}")
        chg = [f"{c.replace('Wave Change as of ', '')}: {norm(wrow[c])}"
               for c in wrow.index
               if str(c).startswith("Wave Change as of ") and norm(wrow[c]) not in ("", "-", "—")]
        print("  Changes:   " + ("; ".join(chg) if chg else "none"))
    else:
        arow = MONTHLY.new_addition_row(uid)
        if arow is not None:
            print(f"  Not in January snapshot — first appeared: {norm(arow['First Snapshot'])}")
            for col in arow.index:
                if str(col).startswith("Wave (") and norm(arow[col]):
                    print(f"  {col[6:-1]:<10} {norm(arow[col])}")
        elif not CHANGE_REPORT.exists():
            print(f"  (change report not found: {CHANGE_REPORT})")
        else:
            print("  No wave changes (not listed in the change report)")

    print("\nJOB ROLE HISTORY (monthly User Mappings snapshots)")
    rrow = MONTHLY.role_row(uid)
    if rrow is not None:
        for n in range(1, 5):
            vals = [(col, norm(rrow[col])) for col in rrow.index
                    if str(col).startswith(f"Job Role {n} (") and norm(rrow[col])]
            for col, v in vals:
                print(f"  {col.replace(f'Job Role {n} ', ''):<12} JR{n}: {v}")
        if norm(rrow.get("What Changed")):
            print(f"  What changed: {norm(rrow['What Changed'])}")
    else:
        print("  No job role changes (not listed in the change report)")

    # --- daily snapshots
    first, last, changes = daily_changes(engine, uid)
    print(f"\nDAILY SNAPSHOTS (history.roster_daily)")
    if first is None:
        print("  Not present in any daily snapshot")
    else:
        print(f"  Covered {first} to {last}")
        if changes:
            for d, label, a, b in changes:
                print(f"  {d}  {label}: {a} -> {b}")
        else:
            print("  No changes across the covered window")

    # --- wave change request forms
    req = wave_requests(engine, uid)
    print("\nWAVE CHANGE REQUEST FORMS (raw.wave_change_requests)")
    if req.empty:
        print("  None on file")
    else:
        for _, r in req.iterrows():
            print(f"  {norm(r['Action'])}: {norm(r['Current_Wave'])} -> "
                  f"{norm(r['Requested_Wave_Change'])} "
                  f"(JR1 {norm(r['Staff_Training_Job_Role_1'])}; {norm(r['_source_file'])})")

    # --- Cornerstone transcript
    cs = cornerstone(engine, uid)
    print("\nCORNERSTONE TRANSCRIPT (EPIC trainings, full history)")
    if cs.empty:
        print("  No EPIC transcript rows")
    else:
        for _, r in cs.iterrows():
            reg = norm(r["Transcript_Registration_Date"])[:10]
            done = norm(r["Transcript_Completed_Date"])[:10]
            print(f"  {reg}  {norm(r['Transcript_Status']):<10} {norm(r['Training_Title'])}"
                  + (f"  (completed {done})" if done else ""))

    # --- Epic TM Lookup
    ep = epic_lookup(engine, uid)
    print("\nEPIC TM LOOKUP (current)")
    if ep is None:
        print("  Not in the Epic Team Member Lookup export")
    else:
        print(f"  Wave: {norm(ep['Wave'])} | Job Title: {norm(ep['Job_Title'])} | "
              f"Curriculum Type: {norm(ep['Curriculum_Type'])}")
        print(f"  Epic Job Category (MVP): {norm(ep['Epic_Job_Categories_MVP'])}")
        if norm(ep["Recommended_Job_Categories"]):
            print(f"  Recommended: {norm(ep['Recommended_Job_Categories'])}")
        print(f"  Fully Registered: {norm(ep['Fully_Registered']) or '-'} | "
              f"Fully Trained: {norm(ep['Fully_Trained']) or '-'}")
        print(f"  Curricula ({norm(ep['Total_Curriculums_MVP'])}): "
              f"{norm(ep['Epic_Curriculums_Cornerstone'])}")
        removed = norm(ep["Removed_Curriculums"]) or norm(ep["Removed_Curriculums_in_MVP"])
        if removed not in ("", "0"):
            print(f"  Removed curricula: {removed}")

    # --- W3 GNF planning mentions
    gnf = gnf_mentions(engine, uid)
    if not gnf.empty:
        print("\nW3 GNF PLANNING FILE mentions")
        for _, r in gnf.iterrows():
            print(f"  {norm(r['GO_AND_FIND'])}  ({norm(r['_source_file'])})")


# ---------------------------------------------------------------- leader mode
def leader_summary(engine, leader: str, full: bool) -> None:
    pat = f"%{leader.strip()}%"
    roster = fetch_df(engine, """
        SELECT UniversalID, FullName, Leader, Wave, JobRole1, Departed
        FROM report.users WHERE Leader LIKE :pat
        ORDER BY FullName""", pat=pat)
    if roster.empty:
        print(f"No users found with Leader LIKE '{leader}'. "
              "Leaders are 'Last, First' — try just the last name.")
        return
    names = roster["Leader"].dropna().unique()
    print(f"\n{len(roster)} users under leader(s): {', '.join(sorted(names))}")

    MONTHLY.load()

    # bulk daily-snapshot diff for the whole roster (one query, joined on leader)
    cols = ", ".join(f"rd.{c}" for c in DAILY_FIELDS.values())
    daily = fetch_df(engine, f"""
        SELECT UPPER(rd.UniversalID) AS UID, rd.SnapshotDate, {cols}
        FROM history.roster_daily rd
        JOIN report.users u ON UPPER(u.UniversalID) = UPPER(rd.UniversalID)
        WHERE u.Leader LIKE :pat
        ORDER BY UPPER(rd.UniversalID), rd.SnapshotDate""", pat=pat)
    daily_chg: dict[str, list] = {}
    for uid, grp in daily.groupby("UID"):
        prev = None
        for _, row in grp.iterrows():
            if prev is not None:
                for label, col in DAILY_FIELDS.items():
                    a, b = norm(prev[col]), norm(row[col])
                    if a != b:
                        daily_chg.setdefault(uid, []).append(
                            f"{str(row['SnapshotDate'])[:10]} {label}: {a or '(blank)'} -> {b or '(blank)'}")
            prev = row

    rows = []
    for _, r in roster.iterrows():
        uid = str(r["UniversalID"]).upper()
        wrow, rrow = MONTHLY.wave_row(uid), MONTHLY.role_row(uid)
        wave_chg = []
        if wrow is not None:
            wave_chg += [norm(wrow[c]) for c in wrow.index
                         if str(c).startswith("Wave Change as of ")
                         and norm(wrow[c]) not in ("", "-", "—")]
        role_chg = norm(rrow["What Changed"]) if rrow is not None else ""
        d = daily_chg.get(uid, [])
        if wave_chg or role_chg or d:
            rows.append((uid, norm(r["FullName"]), "; ".join(wave_chg), role_chg, "; ".join(d)))

    if not rows:
        print("No wave or job-role changes found for anyone under this leader "
              "(Jan-May monthly report + daily snapshots).")
        return

    print(f"\n{len(rows)} of {len(roster)} users have change history:\n")
    for uid, name, w, ro, d in rows:
        print(f"{uid} — {name}")
        if w:
            print(f"    Wave:      {w}")
        if ro:
            print(f"    Job role:  {ro}")
        if d:
            print(f"    Daily:     {d}")
    if full:
        for uid, *_ in rows:
            print_full_report(engine, uid)
    else:
        print("\nFull report for any of them: python scripts\\person_history.py <uid>")


def main() -> None:
    ap = argparse.ArgumentParser(description="Wave / job-role history for any user or leader.")
    ap.add_argument("uids", nargs="*", help="One or more Universal IDs.")
    ap.add_argument("--leader", help="Leader name (or part of it) — summarize their whole team.")
    ap.add_argument("--full", action="store_true",
                    help="With --leader: also print full reports for users that have changes.")
    args = ap.parse_args()
    if not args.uids and not args.leader:
        ap.error("Give at least one Universal ID, or --leader \"Last, First\".")

    engine = create_engine(CONN)
    if args.leader:
        leader_summary(engine, args.leader, args.full)
    for uid in args.uids:
        print_full_report(engine, uid)
    print()


if __name__ == "__main__":
    main()
