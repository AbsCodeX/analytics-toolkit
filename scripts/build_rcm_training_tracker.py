# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_rcm_training_tracker.py — rebuild "RCM Training Tracker.xlsx" AND
"RCM Training Daily Log.xlsx" from SQL.

Replaces the old formula-heavy Wave 2 tracker (archived as
"RCM Training Tracker W2 Archive.xlsx"). Every number, list, and
recommendation is computed here in pandas from the SQL warehouse; the
workbook keeps only a thin layer of lookup/filter formulas so the
Control-sheet leader dropdown and Include/Exclude toggle stay live with
instant recalc. No pasting, no Office Script, no external links.

Sheets:
    Control              leader dropdown + exclusion toggle + source freshness
    Exec Dashboard       KPI cards, weekly activity chart, leader roll-up
    Summary              KPI row, curriculum/class breakdowns, leader summary
    Leader Tracking      per-leader KPIs, class status table, manager table
    Manager Tracking     manager roll-up for the selected leader
    Class Tracking       class roll-up for the selected leader
    Classes to Schedule  classes with unregistered demand + open future seats
    Reg Recommendations  per-person placement options (top 3 future sessions)
    Registration Schedule  full session schedule + possible placements
    Unregistered Summary top-line unregistered / no-show numbers
    Data                 flattened person x event detail (AutoFilter)
    _Cube/_Long/_Lists   hidden precomputed lookup tables

Companion workbook "RCM Training Daily Log.xlsx" (same folder): Daily Totals
(newest-first, trend charts), Daily by Leader, Weekly Summary — rendered from
report.tracker_daily_log (history since 2026-07-06; unregistered/no-show
columns since 2026-07-15).

Sources (all SQL — kept current by YourOrgSQLAutoRefresh / daily_refresh):
    report.tracker_detail, report.tracker_class_schedule, report.users,
    report.training_status, report.unregistered, report.w3_noshow_status,
    report.w3_no_epic_curriculum, report.w3_scorecard

Usage:
    python scripts\\build_rcm_training_tracker.py            # SQL load + build
    python scripts\\build_rcm_training_tracker.py --no-sql   # build only
    python scripts\\build_rcm_training_tracker.py --excel-check  # + recalc QA

Safe to run any time, as often as needed (every 2 hours is fine). The
workbook must be CLOSED in Excel when this runs — an open copy aborts the
write with a clear message.

Local-first build (2026-07-27): both workbooks are written and QA'd in
build_staging\\ (local, not synced), then published to the OneDrive folder
as ONE atomic copy per file — OneDrive never sees a partial write or an
Excel recalc session, which is what corrupted/wedged sync on 07/27.
"""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from datetime import date, datetime, timedelta
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
import xlsxwriter                        # noqa: E402

sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

from onedrive_paths import ONEDRIVE_ROOT          # noqa: E402
from refresh import CONN                          # noqa: E402
from sqlalchemy import create_engine              # noqa: E402

# ---------------------------------------------------------------------------
# Config — W4 = change WAVE/WAVE_NUM/GO_LIVE (and the w3_* view names below
# once W4 twins exist).
# ---------------------------------------------------------------------------
WAVE = "Wave 3"
WAVE_NUM = 3
GO_LIVE = date(2026, 11, 7)
TREND_START = date(2026, 7, 1)   # exec chart hides weeks before this
TITLE = f"RCM Wave {WAVE_NUM}"
OUT_PATH = ONEDRIVE_ROOT / "data" / "reports" / "Main Reports" / "RCM Training Tracker.xlsx"
LOG_PATH = ONEDRIVE_ROOT / "data" / "reports" / "Main Reports" / "RCM Training Daily Log.xlsx"
# Local-first build (2026-07-27): workbooks are written and QA'd in this
# NON-synced staging folder, then published to OneDrive as one atomic copy
# per file. Keeps OneDrive from ever seeing a partial write / Excel-open
# churn (which corrupted the tracker to 0 bytes and wedged sync that day).
STAGING = ROOT / "build_staging"

# Optional-workshop classes excluded from metrics by default (the Control
# sheet toggle switches to Include). Mirrors report.tracker_detail.IsExcluded.
EXCL_RE = re.compile(r"ADVANCED REPORTING|CHARGE CAPTURE", re.IGNORECASE)

MAX_NAMES_PER_CELL = 15          # cap newline lists in Registration Schedule

# Notion-ish neutrals, one muted accent — matches the house report style
# while keeping the old tracker's card/emoji layout.
INK = "#37352F"
GRAY = "#787774"
LINE = "#E3E2DE"
ACCENT = "#2E6FB7"
ACCENT2 = "#A8A29B"


def norm(s) -> str:
    return str(s).strip().upper() if s is not None and s == s else ""


def week_start(ts):
    if pd.isna(ts):
        return pd.NaT
    d = pd.Timestamp(ts).normalize()
    return d - pd.Timedelta(days=d.dayofweek)


# ---------------------------------------------------------------------------
# Load
# ---------------------------------------------------------------------------
def load_frames(engine) -> dict:
    f = {}
    f["det"] = pd.read_sql(
        f"SELECT * FROM report.tracker_detail WHERE Wave = '{WAVE}'", engine)
    f["users"] = pd.read_sql(
        f"SELECT UniversalID, FullName, Leader, AVP, VP, UserType, VendorYN, JobRole1 "
        f"FROM report.users WHERE Wave = '{WAVE}' AND IsInScope = 1 "
        f"AND Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)", engine)
    f["ts"] = pd.read_sql(
        "SELECT UniversalID, IsFullyTrained, IsFullyRegistered "
        "FROM report.training_status", engine)
    f["unreg"] = pd.read_sql(
        f"SELECT * FROM report.unregistered WHERE Wave = '{WAVE}'", engine)
    f["nss"] = pd.read_sql(
        "SELECT * FROM report.w3_noshow_status WHERE IsInScope = 1", engine)
    f["nocurr"] = pd.read_sql(
        "SELECT UniversalID, FullName, Leader, JobRole1 "
        "FROM report.w3_no_epic_curriculum", engine)
    f["scorecard"] = pd.read_sql("SELECT * FROM report.w3_scorecard", engine)
    f["sched"] = pd.read_sql(
        "SELECT * FROM report.tracker_class_schedule", engine)
    f["log"] = pd.read_sql(
        f"SELECT * FROM report.tracker_daily_log WHERE Wave = '{WAVE}' "
        "ORDER BY SnapshotDate, Leader", engine)
    f["snap_live"] = pd.read_sql(
        "SELECT * FROM report.w3_leader_snapshot", engine)
    f["ref_leaders"] = pd.read_sql(
        "SELECT LTRIM(RTRIM(Leader)) AS Leader FROM raw.ref_leaders", engine)
    f["tstatus"] = pd.read_sql(
        f"SELECT * FROM report.tracker_training_status "
        f"WHERE Wave = '{WAVE}' AND OnLeaderList = 1 "
        "ORDER BY Leader, FullName", engine)
    f["noshow"] = pd.read_sql(
        "SELECT * FROM report.w3_noshow_status WHERE IsInScope = 1 "
        "ORDER BY CASE WHEN Resolution = 'No Show standing' THEN 0 ELSE 1 END, "
        "Leader, FullName, ClassTitle", engine)
    f["gnf"] = pd.read_sql(
        f"SELECT * FROM report.gnf_assignments "
        f"WHERE Wave = '{WAVE}' AND OnLeaderList = 1 "
        "ORDER BY Leader, FullName, GoAndFind", engine)
    f["audit"] = pd.read_sql(
        "SELECT * FROM report.registration_audit "
        "ORDER BY CheckName, Leader, FullName", engine)
    f["gnf_status"] = pd.read_sql(
        "WITH mapf AS (SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),UID)))) AS uid, "
        "  MAX(CASE WHEN LTRIM(RTRIM(ISNULL(GO_AND_FIND,''))) <> '' THEN 1 ELSE 0 END) AS has_mod "
        "  FROM raw.wave_gnf_w3_mapping "
        "  GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),UID))))) "
        "SELECT u.UniversalID, u.FullName, u.Leader, "
        "  CASE WHEN m.has_mod = 1 THEN 'Mapped' "
        "       WHEN m.has_mod = 0 THEN 'No modules assigned' "
        "       ELSE 'Pending mapping' END AS GnfStatus "
        "FROM report.users u "
        "LEFT JOIN mapf m ON m.uid = u.UniversalID "
        "WHERE u.Wave = 'Wave 3' AND u.IsInScope = 1 AND u.OnLeaderList = 1",
        engine)
    f["fresh"] = pd.read_sql(
        "SELECT 'Epic Curriculum Status Detail' AS Feed, MAX(_source_file) AS SourceFile, MAX(_loaded_at) AS LoadedAt FROM raw.epic_status "
        "UNION ALL SELECT 'Epic Class Schedule', MAX(_source_file), MAX(_loaded_at) FROM raw.epic_class_schedule "
        "UNION ALL SELECT 'Epic Unregistered Sessions', MAX(_source_file), MAX(_loaded_at) FROM raw.epic_unregistered "
        "UNION ALL SELECT 'Cornerstone Enterprise', MAX(_source_file), MAX(_loaded_at) FROM raw.cornerstone "
        "UNION ALL SELECT 'Master Wave File', MAX(_source_file), MAX(_loaded_at) FROM raw.master", engine)
    for c in ("EventDate", "RegistrationDate"):
        f["det"][c] = pd.to_datetime(f["det"][c], errors="coerce")
    f["sched"]["SessionDate"] = pd.to_datetime(f["sched"]["SessionDate"], errors="coerce")
    return f


# ---------------------------------------------------------------------------
# Compute
# ---------------------------------------------------------------------------
def reg_state(v: str) -> int:
    u = norm(v)
    if u.startswith("YES"):
        return 3
    if u.startswith("EXEMPT"):
        return 2
    if u.startswith("NO -NO SHOW"):
        return 1
    return 0


def fmt_d(ts) -> str:
    return "" if pd.isna(ts) else pd.Timestamp(ts).strftime("%-m/%-d") \
        if os.name != "nt" else pd.Timestamp(ts).strftime("%#m/%#d")


def loc_str(row) -> str:
    loc = str(row.get("Location") or "").strip()
    room = str(row.get("Room") or "").strip()
    if room and room.upper() != "NO_ROOM":
        return f"{loc} — {room}"
    return loc


def compute(f: dict) -> dict:
    today = pd.Timestamp(date.today())
    users = f["users"].copy()
    users["UniversalID"] = users["UniversalID"].map(norm)
    # canonical allowlist, not just leaders who happen to have wave members —
    # zero-member leaders still appear in the dropdown/roll-ups showing 0s
    leaders = sorted(f["ref_leaders"]["Leader"].dropna().unique())

    det = f["det"].copy()
    det["UniversalID"] = det["UniversalID"].map(norm)
    det["ClassKey"] = det["Event_Class"].map(norm)
    det["RegState"] = det["Event_Class_Registered"].map(reg_state)
    det["IsSession"] = det["Event_Class_Type"].fillna("").str.strip().str.upper() == "SESSION"
    det["IsCompleted"] = det["Event_Class_Status"].fillna("").str.strip().str.upper() == "COMPLETED"

    ts = f["ts"].copy()
    ts["UniversalID"] = ts["UniversalID"].map(norm)
    users = users.merge(ts, on="UniversalID", how="left")
    users[["IsFullyTrained", "IsFullyRegistered"]] = (
        users[["IsFullyTrained", "IsFullyRegistered"]].fillna(0).astype(int))

    unreg = f["unreg"].copy()
    unreg["UniversalID"] = unreg["UniversalID"].map(norm)
    unreg["ClassKey"] = unreg["Event_Class"].map(norm)

    nss = f["nss"].copy()
    nss["UniversalID"] = nss["UniversalID"].map(norm)
    nss["ClassKey"] = nss["ClassTitle"].map(norm)
    nss_standing = nss[nss["Resolution"] == "No Show standing"].copy()
    nss_rereg_needed = nss_standing[
        nss_standing["RegistrationAction"] == "Needs to be Re-Registered"].copy()

    nocurr = f["nocurr"].copy()
    nocurr["UniversalID"] = nocurr["UniversalID"].map(norm)

    sched = f["sched"].copy()
    sched["ClassKey"] = sched["EventName"].map(norm)
    sched = sched.sort_values(["SessionDate", "StartTime"], kind="stable")
    fut = sched[(sched["SessionDate"] >= today + pd.Timedelta(days=1))].copy()
    fut_seats = fut[fut["AvailableSeats"].fillna(0) > 0].copy()

    uid_leader = users.set_index("UniversalID")["Leader"]
    uid_name = users.set_index("UniversalID")["FullName"]

    # ---- demand: who still needs what class (unregistered + no-show re-reg)
    d1 = unreg[["UniversalID", "FullName", "Leader", "ClassKey", "Event_Class"]].copy()
    d1["Reason"] = "Unregistered"
    d2 = nss_rereg_needed[["UniversalID", "FullName", "Leader", "ClassKey", "ClassTitle"]].rename(
        columns={"ClassTitle": "Event_Class"}).copy()
    d2["Reason"] = "No-Show — Re-register"
    demand = pd.concat([d1, d2], ignore_index=True)
    demand = demand.drop_duplicates(subset=["UniversalID", "ClassKey"], keep="first")
    demand["IsExcludedClass"] = demand["Event_Class"].fillna("").str.contains(EXCL_RE)

    # ---- per-variant cube + long tables ------------------------------------
    cube_rows, t_curr, t_class, t_mgr, t_trend = [], [], [], [], []
    days_to_golive = (GO_LIVE - date.today()).days
    golive_label = "Live" if days_to_golive <= 0 else str(days_to_golive)

    for variant in ("Exclude", "Include"):
        dv = det[~det["IsExcluded"].astype(bool)] if variant == "Exclude" else det
        dmv = demand[~demand["IsExcludedClass"]] if variant == "Exclude" else demand
        nsv = nss_standing[~nss_standing["ClassTitle"].fillna("").str.contains(EXCL_RE)] \
            if variant == "Exclude" else nss_standing
        sess = dv[dv["IsSession"]]

        # per person x class rollup (this variant)
        pc = (sess.groupby(["UniversalID", "ClassKey"], as_index=False)
              .agg(Leader=("Leader", "first"),
                   Class=("Event_Class", "first"),
                   RegState=("RegState", "max"),
                   Completed=("IsCompleted", "max"),
                   LastDate=("EventDate", "max")))

        unreg_by_uid = dmv.groupby("UniversalID").size()
        ns_by_uid = nsv.groupby("UniversalID").size()

        def leader_slice(name):
            if name == "All":
                return (users, sess, pc, dmv, nsv, nocurr)
            return (users[users["Leader"] == name],
                    sess[sess["Leader"] == name],
                    pc[pc["Leader"] == name],
                    dmv[dmv["Leader"] == name],
                    nsv[nsv["Leader"] == name],
                    nocurr[nocurr["Leader"] == name])

        for name in ["All"] + leaders:
            u, s, p, dm, ns, nc = leader_slice(name)
            members = len(u)
            trained = int(u["IsFullyTrained"].sum())
            reg = int(u["IsFullyRegistered"].sum())
            unreg_users = dm["UniversalID"].nunique()
            unreg_sessions = len(dm)
            ns_users = ns["UniversalID"].nunique()
            ns_sessions = len(ns)
            comp_occ = int(p["Completed"].sum())
            req_occ = len(p)
            # scheduled final session — max over ALL session dates incl. future
            # (same meaning as the original tracker's MAXIFS "Last Class Date")
            dated = s[s["EventDate"].notna()]
            last_row = dated.loc[dated["EventDate"].idxmax()] if len(dated) else None
            cube_rows.append({
                "Key": f"{variant}|{name}", "Variant": variant, "Leader": name,
                "Members": members,
                "FullyTrained": trained,
                "PctTrained": trained / members if members else 0,
                "FullyRegistered": reg,
                "PctRegistered": reg / members if members else 0,
                "UnregUsers": unreg_users,
                "UnregSessions": unreg_sessions,
                "PctUnregUsers": unreg_users / members if members else 0,
                "NoShowUsers": ns_users,
                "NoShowSessions": ns_sessions,
                "PlaceUnregTMs": int(dm[dm["Reason"] == "Unregistered"]["UniversalID"].nunique()),
                "PlaceUnregSessions": int((dm["Reason"] == "Unregistered").sum()),
                "PlaceNoShowTMs": int(dm[dm["Reason"] != "Unregistered"]["UniversalID"].nunique()),
                "PlaceNoShowSessions": int((dm["Reason"] != "Unregistered").sum()),
                "NoEpicCurriculum": len(nc),
                "SessionsCompleted": comp_occ,
                "SessionsRequired": req_occ,
                "PctSessionsCompleted": comp_occ / req_occ if req_occ else 0,
                "Curricula": s["Curriculum"].nunique(),
                "Classes": p["ClassKey"].nunique(),
                "LastClass": "" if last_row is None else str(last_row["Event_Class"]),
                "LastClassDate": pd.NaT if last_row is None else last_row["EventDate"],
                "DaysToGoLive": golive_label,
                "GoLiveDate": pd.Timestamp(GO_LIVE),
            })

            # curriculum breakdown
            cb = (s.groupby("Curriculum", as_index=False)
                  .agg(Users=("UniversalID", "nunique"), LastSession=("EventDate", "max")))
            for r in cb.itertuples(index=False):
                t_curr.append((variant, name, r.Curriculum, int(r.Users), r.LastSession))

            # class breakdown
            if len(p):
                nxt = fut.groupby("ClassKey")["SessionDate"].min()
                cbk = (p.groupby("ClassKey", as_index=False)
                       .agg(Class=("Class", "first"),
                            Users=("UniversalID", "nunique"),
                            Registered=("RegState", lambda x: int((x >= 2).sum())),
                            Completed=("Completed", "sum"),
                            LastSession=("LastDate", "max")))
                dm_by_class = dm.groupby("ClassKey").size()
                for r in cbk.itertuples(index=False):
                    gaps = int(dm_by_class.get(r.ClassKey, 0))
                    nx = nxt.get(r.ClassKey, pd.NaT)
                    if int(r.Completed) >= int(r.Users):
                        status = "✅ Completed"
                    elif gaps and pd.isna(nx):
                        status = "🔴 Gaps, no future session"
                    elif gaps:
                        status = "🟠 Gaps to place"
                    else:
                        status = "🟢 On track"
                    t_class.append((variant, name, r.Class, int(r.Users),
                                    int(r.Registered), int(r.Completed), gaps,
                                    r.LastSession, nx, status))

            # manager roll-up
            if len(u):
                mgr_src = s.groupby("UniversalID")["Direct_Manager"].first()
                um = u.copy()
                um["Manager"] = um["UniversalID"].map(mgr_src).fillna("—")
                mg = (um.groupby("Manager", as_index=False)
                      .agg(Members=("UniversalID", "nunique"),
                           FullyTrained=("IsFullyTrained", "sum"),
                           FullyRegistered=("IsFullyRegistered", "sum")))
                for r in mg.itertuples(index=False):
                    uids = set(um[um["Manager"] == r.Manager]["UniversalID"])
                    t_mgr.append((variant, name, r.Manager, int(r.Members),
                                  int(r.FullyTrained), int(r.FullyRegistered),
                                  int(sum(1 for x in uids if unreg_by_uid.get(x, 0) > 0)),
                                  int(sum(1 for x in uids if ns_by_uid.get(x, 0) > 0))))

            # weekly trend: sessions attended (past) vs registered (future) —
            # same unit, one axis; the current week can carry both. Epic
            # repeats a class row under every curriculum that requires it, so
            # dedup to one row per person x class x session datetime.
            sess_keys = ["UniversalID", "ClassKey", "EventDate"]
            booked = s[s["EventDate"].notna() & (s["EventDate"] > today)
                       & (s["Event_Class_Registered"].fillna("").str.upper()
                          .str.startswith("YES"))].drop_duplicates(sess_keys)
            booked = booked.copy()
            booked["Week"] = booked["EventDate"].map(week_start)
            comps = s[s["IsCompleted"] & s["EventDate"].notna()
                      & (s["EventDate"] <= today)].drop_duplicates(sess_keys)
            comps = comps.copy()
            comps["Week"] = comps["EventDate"].map(week_start)
            wk = (pd.concat([
                booked.groupby("Week").size().rename("Booked"),
                comps.groupby("Week").size().rename("Completions")], axis=1)
                .fillna(0).astype(int).sort_index())
            wk = wk[wk.index >= pd.Timestamp(TREND_START)]
            for w, r in wk.iterrows():
                t_trend.append((variant, name, w, int(r["Booked"]),
                                int(r["Completions"])))

    cube = pd.DataFrame(cube_rows)
    long_curr = pd.DataFrame(t_curr, columns=[
        "Variant", "Leader", "Curriculum", "Users", "LastSession"])
    long_class = pd.DataFrame(t_class, columns=[
        "Variant", "Leader", "Class", "Users", "Registered", "Completed",
        "OpenGaps", "LastSession", "NextSession", "Status"])
    long_mgr = pd.DataFrame(t_mgr, columns=[
        "Variant", "Leader", "Manager", "Members", "FullyTrained",
        "FullyRegistered", "UnregUsers", "NoShowUsers"])
    long_trend = pd.DataFrame(t_trend, columns=[
        "Variant", "Leader", "Week", "Booked", "Completions"])
    weeks = sorted(long_trend["Week"].dropna().unique())

    # ---- Classes to Schedule ----------------------------------------------
    cts_rows = []
    for ck, grp in demand.groupby("ClassKey"):
        cls = grp["Event_Class"].iloc[0]
        uids = grp["UniversalID"].tolist()
        users_txt = "\n".join(
            f"{u} — {uid_name.get(u, grp[grp['UniversalID'] == u]['FullName'].iloc[0])}"
            for u in uids)
        leaders_txt = "\n".join(str(uid_leader.get(u, "")) for u in uids)
        options = fut_seats[fut_seats["ClassKey"] == ck]
        opts_txt = "\n".join(
            f"{fmt_d(r.SessionDate)} — {loc_str(r._asdict())} ({int(r.AvailableSeats)} seats)"
            for r in options.head(8).itertuples(index=False))
        if len(options) > 8:
            opts_txt += f"\n… +{len(options) - 8} more sessions"
        any_future = len(fut[fut["ClassKey"] == ck]) > 0
        if len(options):
            status = "🟢 Has future session"
            first = options.iloc[0]
            rec = f"{fmt_d(first['SessionDate'])} — {loc_str(first)} ({int(first['AvailableSeats'])} seats)"
        elif any_future:
            status = "🟠 Future sessions full"
            rec = "Waitlist / add seats"
        else:
            status = "🔴 Needs scheduling"
            rec = "No future session on the schedule"
        cts_rows.append({"Class": cls, "Demand": len(grp), "Status": status,
                         "Users": users_txt, "Leaders": leaders_txt,
                         "Options": opts_txt, "Recommended": rec})
    cts = pd.DataFrame(cts_rows)
    if len(cts):
        order = {"🔴 Needs scheduling": 0, "🟠 Future sessions full": 1,
                 "🟢 Has future session": 2}
        cts["_o"] = cts["Status"].map(order)
        cts = (cts.sort_values(["_o", "Demand", "Class"],
                               ascending=[True, False, True])
               .drop(columns="_o").reset_index(drop=True))

    # ---- Reg Recommendations ----------------------------------------------
    rec_rows = []
    for r in demand.itertuples(index=False):
        options = fut_seats[fut_seats["ClassKey"] == r.ClassKey].head(3)
        row = {"UniversalID": r.UniversalID,
               "FullName": uid_name.get(r.UniversalID, r.FullName),
               "Leader": r.Leader, "Reason": r.Reason, "Class": r.Event_Class}
        for i in range(3):
            if i < len(options):
                o = options.iloc[i]
                row[f"Opt{i+1}Date"] = o["SessionDate"]
                row[f"Opt{i+1}Loc"] = loc_str(o)
                row[f"Opt{i+1}Seats"] = int(o["AvailableSeats"])
            else:
                row[f"Opt{i+1}Date"] = pd.NaT
                row[f"Opt{i+1}Loc"] = ""
                row[f"Opt{i+1}Seats"] = None
        row["Assignment"] = "✅ Session 1" if len(options) else "⚠️ No future sessions"
        rec_rows.append(row)
    recs = pd.DataFrame(rec_rows)
    if len(recs):
        recs = recs.sort_values(["Leader", "FullName", "Class"]).reset_index(drop=True)

    # ---- Registration Schedule (schedule + placements) ---------------------
    rs = sched.copy()
    rs["Option"] = ""
    rs["PossibleUsers"] = ""
    rs["TheirLeaders"] = ""
    fut_idx = rs["SessionDate"] >= today + pd.Timedelta(days=1)
    rs.loc[fut_idx, "Option"] = (
        rs[fut_idx].groupby("ClassKey").cumcount().add(1).map("Session {}".format))
    demand_by_class = {ck: grp for ck, grp in demand.groupby("ClassKey")}
    for idx in rs[fut_idx & (rs["AvailableSeats"].fillna(0) > 0)].index:
        ck = rs.at[idx, "ClassKey"]
        grp = demand_by_class.get(ck)
        if grp is None:
            continue
        uids = grp["UniversalID"].tolist()
        shown = uids[:MAX_NAMES_PER_CELL]
        extra = len(uids) - len(shown)
        utxt = "\n".join(f"{u} — {uid_name.get(u, '')}" for u in shown)
        ltxt = "\n".join(str(uid_leader.get(u, "")) for u in shown)
        if extra > 0:
            utxt += f"\n… +{extra} more"
            ltxt += "\n…"
        rs.at[idx, "PossibleUsers"] = utxt
        rs.at[idx, "TheirLeaders"] = ltxt

    # ---- Unregistered Summary ----------------------------------------------
    unsum = {
        "TotalTMs": demand["UniversalID"].nunique(),
        "TotalSessions": len(demand),
        "UnregTMs": d1["UniversalID"].nunique(),
        "UnregSessions": len(demand[demand["Reason"] == "Unregistered"]),
        "NoShowTMs": d2["UniversalID"].nunique(),
        "NoShowSessions": len(demand[demand["Reason"] != "Unregistered"]),
    }

    # ---- Data sheet (UniversalID first, per house rule) --------------------
    data_cols = {
        "UniversalID": "Universal Id", "FullName": "Team Member",
        "Leader": "Leaders", "Direct_Manager": "Direct Manager",
        "CurriculumType": "Curriculum Type", "Fully_Trained": "Fully Trained?",
        "Fully_Registered": "Fully Registered?", "UserType": "User Type",
        "Facility_Serviceline": "Facility / Serviceline",
        "Department": "Department", "Epic_Job_Category": "Epic Job Category",
        "Curriculum": "Curriculum", "Curriculum_Status": "Curriculum Status",
        "Sequence": "Sequence", "Event_Class": "Event (Class)",
        "Event_Class_Type": "Event (Class) Type",
        "Event_Class_Registered": "Event (Class) Registered?",
        "Event_Class_Status": "Event (Class) Status",
        "DurationHours": "Duration Hours", "EventDate": "Event (Class) Date",
        "Event_Class_Location": "Location", "Event_Class_Room": "Room",
        "RegistrationDate": "Registration Date", "IsExcluded": "Is Excluded",
    }
    data_df = det[list(data_cols)].rename(columns=data_cols)
    data_df = data_df.sort_values(["Team Member", "Curriculum", "Sequence"],
                                  kind="stable").reset_index(drop=True)

    # next 10 upcoming sessions for classes this population needs
    needed = set(det[det["IsSession"]]["ClassKey"].unique())
    upcoming = fut[fut["ClassKey"].isin(needed)].nsmallest(10, "SessionDate")[
        ["SessionDate", "EventName", "Location", "Room", "StartTime",
         "AvailableSeats"]]

    # ---- Daily Log ---------------------------------------------------------
    log = f["log"].copy()
    log["SnapshotDate"] = pd.to_datetime(log["SnapshotDate"], errors="coerce")
    metric_cols = ["Members", "FullyRegistered", "FullyTrained",
                   "UnregisteredUsers", "UnregisteredSessions",
                   "NoShowUsersStanding", "NoShowSessionsStanding",
                   "NoShowUsersToDate", "NoShowSessionsToDate"]
    # if today's snapshot hasn't landed yet, append today's live numbers
    if len(log) == 0 or log["SnapshotDate"].max().date() < date.today():
        live = f["snap_live"].rename(columns={"W3Users": "Members"}).copy()
        live["SnapshotDate"] = today
        live["Wave"] = WAVE
        log = pd.concat([log, live[["SnapshotDate", "Wave", "Leader"]
                                   + metric_cols]], ignore_index=True)
    log = log.sort_values(["SnapshotDate", "Leader"]).reset_index(drop=True)
    delta_of = {"FullyRegistered": "ΔReg", "FullyTrained": "ΔTrained",
                "UnregisteredUsers": "ΔUnreg", "NoShowUsersStanding": "ΔNoShow"}
    log_leader = log[["SnapshotDate", "Leader"] + metric_cols].copy()
    for src, dst in delta_of.items():
        log_leader[dst] = (log_leader.sort_values("SnapshotDate")
                           .groupby("Leader")[src].diff())
    log_totals = (log.groupby("SnapshotDate", as_index=False)[metric_cols]
                  .sum(min_count=1))
    for src, dst in delta_of.items():
        log_totals[dst] = log_totals[src].diff()
    for df_ in (log_totals, log_leader):
        m = df_["Members"].where(df_["Members"] > 0)
        df_["RegPct"] = df_["FullyRegistered"] / m
        df_["TrainedPct"] = df_["FullyTrained"] / m

    # weekly roll-up: end-of-week snapshot values + week-over-week deltas
    def weekly(df, by_leader: bool):
        d2 = df.copy()
        d2["Week"] = d2["SnapshotDate"].map(week_start)
        keys = ["Leader"] if by_leader else []
        idx = d2.groupby(keys + ["Week"])["SnapshotDate"].idxmax()
        days = d2.groupby(keys + ["Week"])["SnapshotDate"].nunique()
        wk = d2.loc[sorted(idx)].copy()
        wk["DaysCaptured"] = wk.set_index(keys + ["Week"]).index.map(days)
        wk = wk.sort_values(keys + ["Week"]).reset_index(drop=True)
        for src, dst in delta_of.items():
            wk[dst] = (wk.groupby("Leader")[src].diff() if by_leader
                       else wk[src].diff())
        return wk

    log_week_totals = weekly(log_totals, by_leader=False)
    log_week_leader = weekly(log_leader, by_leader=True)

    # ---- Visuals sheet feed (all leaders, Exclude scope, static) -----------
    exc = cube[(cube["Variant"] == "Exclude") & (cube["Leader"] != "All")]
    act = exc[exc["Members"] > 0].sort_values("Members")     # asc → largest bar on top
    allx = cube[(cube["Variant"] == "Exclude") & (cube["Leader"] == "All")].iloc[0]
    partial = max(0, int(allx["Members"] - allx["FullyRegistered"]
                         - allx["UnregUsers"] - allx["NoEpicCurriculum"]))
    viz = {
        "leader_reg": pd.DataFrame({
            "Leader": act["Leader"],
            "Registered": act["FullyRegistered"].astype(int),
            "Remaining": (act["Members"] - act["FullyRegistered"]).astype(int)}),
        "leader_unreg": (act[act["UnregUsers"] > 0].sort_values("UnregUsers")
                         [["Leader", "UnregUsers"]].astype({"UnregUsers": int})),
        "mix": [(label, v) for label, v in
                [("Fully Registered", int(allx["FullyRegistered"])),
                 ("Unregistered Gaps", int(allx["UnregUsers"])),
                 ("No Epic Curriculum", int(allx["NoEpicCurriculum"])),
                 ("Partially Registered", partial)] if v > 0],
        "trend": log_totals[["SnapshotDate", "FullyRegistered", "FullyTrained"]],
        "weekly": (long_trend[(long_trend["Variant"] == "Exclude")
                              & (long_trend["Leader"] == "All")]
                   [["Week", "Booked", "Completions"]]),
        "all": allx,
        "deltas": (log_totals.iloc[-1][["ΔReg", "ΔTrained", "ΔUnreg", "ΔNoShow"]]
                   .to_dict() if len(log_totals) else {}),
    }

    # ---- Training Status / No Shows / Go and Finds -------------------------
    ts = f["tstatus"].copy()
    for c in ("LastAttendedDate", "FinalScheduledDate", "CompletionDate"):
        ts[c] = pd.to_datetime(ts[c], errors="coerce")
    ts["Attention"] = ""
    needs = (ts["TrainingStatus"].isin(["Not Started", "In Progress"])
             & (ts["FinalScheduledDate"].isna() | (ts["UnregisteredClasses"] > 0)))
    ts.loc[needs, "Attention"] = "⚠️ Not fully scheduled"
    ts.loc[ts["TrainingStatus"] == "No Epic Curriculum",
           "Attention"] = "⛔ Cannot register yet"
    STATUS_ORDER = ["Not Started", "In Progress", "Fully Trained",
                    "No Epic Curriculum"]
    tstat_rows = []
    for name in leaders + ["TOTAL"]:
        sub = ts if name == "TOTAL" else ts[ts["Leader"] == name]
        n = len(sub)
        row = {"Leader": name, "Members": n}
        for st in STATUS_ORDER:
            k = int((sub["TrainingStatus"] == st).sum())
            row[st] = k
            row[st + " %"] = (k / n) if n else None
        tstat_rows.append(row)
    tstat_leader = pd.DataFrame(tstat_rows)

    ns = f["noshow"].copy()
    for c in ("LastNoShowSessionDate", "CurrentSessionDate",
              "CurrentRegisteredDate"):
        ns[c] = pd.to_datetime(ns[c], errors="coerce")

    gnf = f["gnf"].copy()
    gnf_status = f["gnf_status"].copy()
    gnf_status["UniversalID"] = gnf_status["UniversalID"].map(norm)
    stat = (gnf_status.pivot_table(index="Leader", columns="GnfStatus",
                                   aggfunc="size", fill_value=0)
            .reindex(leaders).fillna(0).astype(int))
    for c in ("Mapped", "No modules assigned", "Pending mapping"):
        if c not in stat.columns:
            stat[c] = 0
    gnf_leader = (gnf.groupby("Leader")
                  .agg(ModuleAssignments=("GoAndFind", "size"),
                       DistinctModules=("GoAndFind", "nunique"))
                  .reindex(leaders).fillna(0).astype(int))
    gnf_leader["Mapped"] = stat["Mapped"]
    gnf_leader["NoModules"] = stat["No modules assigned"]
    gnf_leader["PendingMapping"] = stat["Pending mapping"]
    gnf_leader = gnf_leader.reset_index()
    tot_row = pd.DataFrame([{
        "Leader": "TOTAL",
        "ModuleAssignments": len(gnf),
        "DistinctModules": gnf["GoAndFind"].nunique(),
        "Mapped": int((gnf_status["GnfStatus"] == "Mapped").sum()),
        "NoModules": int((gnf_status["GnfStatus"] == "No modules assigned").sum()),
        "PendingMapping": int((gnf_status["GnfStatus"] == "Pending mapping").sum())}])
    gnf_leader = pd.concat([gnf_leader, tot_row], ignore_index=True)

    # feed for the Visuals status-by-leader 100% stacked bar (active leaders,
    # ascending members so the largest team lands on top of the bar chart)
    sl = tstat_leader[(tstat_leader["Leader"] != "TOTAL")
                      & (tstat_leader["Members"] > 0)].copy()
    viz["status_leader"] = (sl.sort_values("Members")
                            [["Leader", "Not Started", "In Progress",
                              "Fully Trained"]])

    audit = f["audit"].copy()
    audit["EventDate"] = pd.to_datetime(audit["EventDate"], errors="coerce")
    audit_block = audit[["Leader", "UniversalID", "FullName", "CheckName",
                         "TrainingTrack", "ClassTitle", "EventDate", "Detail"]].copy()
    # truly-blank cells spill as 0 through Excel FILTER — store "" instead
    for col in ("TrainingTrack", "ClassTitle", "Detail"):
        audit_block[col] = audit_block[col].fillna("")
    audit_block["EventDate"] = audit_block["EventDate"].astype(object).where(
        audit_block["EventDate"].notna(), "")

    unmatched = sorted(set(demand["ClassKey"]) - set(sched["ClassKey"]))
    return {"cube": cube, "long_curr": long_curr, "long_class": long_class,
            "audit": audit_block,
            "long_mgr": long_mgr, "long_trend": long_trend, "weeks": weeks,
            "cts": cts, "recs": recs, "regsched": rs, "unsum": unsum,
            "data": data_df, "leaders": leaders, "upcoming": upcoming,
            "fresh": f["fresh"], "scorecard": f["scorecard"],
            "log_totals": log_totals, "log_leader": log_leader,
            "log_week_totals": log_week_totals, "log_week_leader": log_week_leader,
            "viz": viz, "tstatus": ts, "tstat_leader": tstat_leader,
            "noshow": ns, "gnf": gnf, "gnf_leader": gnf_leader,
            "gnf_status": gnf_status,
            "unmatched_classes": unmatched}


# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------
class Fmt:
    """Shared workbook formats."""

    def __init__(self, wb):
        base = {"font_name": "Segoe UI", "font_color": INK}
        self.title = wb.add_format({**base, "font_size": 18, "bold": True})
        self.subtitle = wb.add_format({**base, "font_size": 10, "font_color": GRAY})
        self.h2 = wb.add_format({**base, "font_size": 12, "bold": True})
        self.label = wb.add_format({**base, "font_size": 9, "font_color": GRAY,
                                    "align": "center"})
        self.kpi = wb.add_format({**base, "font_size": 22, "bold": True,
                                  "align": "center"})
        self.kpi_pct = wb.add_format({**base, "font_size": 11, "font_color": GRAY,
                                      "align": "center", "num_format": "0.0%"})
        self.kpi_txt = wb.add_format({**base, "font_size": 16, "bold": True,
                                      "align": "center"})
        self.kpi_link = wb.add_format({"font_name": "Segoe UI", "font_size": 22,
                                       "bold": True, "align": "center",
                                       "font_color": ACCENT})
        self.kpi_txt_link = wb.add_format({"font_name": "Segoe UI",
                                           "font_size": 16, "bold": True,
                                           "align": "center",
                                           "font_color": ACCENT})
        self.th = wb.add_format({**base, "bold": True, "font_size": 10,
                                 "bottom": 1, "border_color": LINE})
        self.cell = wb.add_format({**base, "font_size": 10})
        self.cell_wrap = wb.add_format({**base, "font_size": 10, "text_wrap": True,
                                        "valign": "top"})
        self.num = wb.add_format({**base, "font_size": 10, "num_format": "#,##0"})
        self.pct = wb.add_format({**base, "font_size": 10, "num_format": "0.0%"})
        self.dt = wb.add_format({**base, "font_size": 10,
                                 "num_format": "m/d/yyyy"})
        self.dtm = wb.add_format({**base, "font_size": 10,
                                  "num_format": "m/d/yyyy h:mm AM/PM"})
        self.note = wb.add_format({**base, "font_size": 9, "font_color": GRAY,
                                   "italic": True})
        self.delta = wb.add_format({**base, "font_size": 10, "font_color": GRAY,
                                    "num_format": "+0;-0;0"})
        self.ctrl = wb.add_format({**base, "font_size": 11, "bold": True,
                                   "border": 1, "border_color": ACCENT,
                                   "bg_color": "#FFFFFF", "align": "center"})
        self.link = wb.add_format({"font_name": "Segoe UI", "font_size": 10,
                                   "font_color": ACCENT, "underline": 1})
        self.gterm = wb.add_format({**base, "font_size": 10, "bold": True,
                                    "valign": "top"})
        self.gdef = wb.add_format({**base, "font_size": 10, "text_wrap": True,
                                   "valign": "top"})


def _w(ws, r, c, v, f=None, fdt=None):
    """Type-aware cell write; blanks for NaN/None."""
    if v is None or (isinstance(v, float) and v != v) or v is pd.NaT:
        ws.write_blank(r, c, None, f)
    elif isinstance(v, (datetime, pd.Timestamp)):
        ws.write_datetime(r, c, pd.Timestamp(v).to_pydatetime(), fdt or f)
    elif isinstance(v, (int, float)):
        ws.write_number(r, c, v, f)
    else:
        ws.write_string(r, c, str(v), f)


def write_df(ws, df, fm, start_row=0, start_col=0, date_cols=(), num_cols=(),
             wrap_cols=(), autofilter=True, freeze=True, widths=None):
    cols = list(df.columns)
    for c, name in enumerate(cols):
        ws.write_string(start_row, start_col + c, str(name), fm.th)
    for r, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for c, v in enumerate(row):
            name = cols[c]
            f = fm.cell
            fdt = fm.dt
            if name in wrap_cols:
                f = fm.cell_wrap
            elif name in num_cols:
                f = fm.num
            _w(ws, r, start_col + c, v, f, fdt)
    if autofilter and len(cols):
        ws.autofilter(start_row, start_col, start_row + max(len(df), 1),
                      start_col + len(cols) - 1)
    if freeze:
        ws.freeze_panes(start_row + 1, 0)
    if widths:
        for c, wdt in enumerate(widths):
            ws.set_column(start_col + c, start_col + c, wdt)


def build_workbook(d: dict, path: Path) -> None:
    wb = xlsxwriter.Workbook(str(path), {"nan_inf_to_errors": True,
                                         "default_date_format": "m/d/yyyy"})
    fm = Fmt(wb)
    cube, weeks = d["cube"], d["weeks"]
    n_cube = len(cube)
    n_curr, n_class = len(d["long_curr"]), len(d["long_class"])
    n_mgr, n_trend = len(d["long_mgr"]), len(d["long_trend"])
    n_audit = len(d["audit"])
    built_at = datetime.now()
    asof = f"Data as of {built_at:%m/%d/%Y %I:%M %p}"

    def stamp_live(ws, row=3, col=1):
        # standard caption for sheets driven by the Control dropdown/toggle
        ws.write_formula(
            row, col,
            '="🔎 Control filter applies — showing: "&SelectedLeader'
            '&"  ·  "&ExclToggle&" scope  ·  ' + asof + '"', fm.note)

    def stamp_static(ws, row=3, col=1):
        # standard caption for sheets that always show every leader
        ws.write_string(
            row, col,
            "🌐 All leaders — the Control leader filter does not apply on "
            "this sheet  ·  " + asof, fm.note)

    # ---- hidden: _Lists ----------------------------------------------------
    ws_lists = wb.add_worksheet("_Lists")
    ws_lists.write_string(0, 0, "Leaders")
    leader_opts = ["All"] + d["leaders"]
    for i, name in enumerate(leader_opts, start=1):
        ws_lists.write_string(i, 0, name)
    ws_lists.write_string(0, 1, "Toggle")
    ws_lists.write_string(1, 1, "Exclude")
    ws_lists.write_string(2, 1, "Include")
    ws_lists.hide()

    # ---- hidden: _Cube -----------------------------------------------------
    ws_cube = wb.add_worksheet("_Cube")
    cube_cols = list(cube.columns)
    for c, name in enumerate(cube_cols):
        ws_cube.write_string(0, c, name)
    for r, row in enumerate(cube.itertuples(index=False), start=1):
        for c, v in enumerate(row):
            _w(ws_cube, r, c, v, fm.cell, fm.dt)
    # chart feed block: X=Week label, Y=Registrations, Z=Completions
    cX, cY, cZ = len(cube_cols) + 2, len(cube_cols) + 3, len(cube_cols) + 4
    ws_cube.write_string(0, cX, "Week")
    ws_cube.write_string(0, cY, "Booked")
    ws_cube.write_string(0, cZ, "Completions")
    tr_a, tr_b = 2, n_trend + 1     # _Long trend rows (AA..AE headers row 1)
    for i, wk in enumerate(weeks, start=1):
        ws_cube.write_datetime(i, cX, pd.Timestamp(wk).to_pydatetime(),
                               wb.add_format({"num_format": "m/d"}))
        for cc, src in ((cY, "AD"), (cZ, "AE")):
            ws_cube.write_formula(
                i, cc,
                f"=SUMIFS(_Long!${src}${tr_a}:${src}${tr_b},"
                f"_Long!$AA${tr_a}:$AA${tr_b},ExclToggle,"
                f"_Long!$AB${tr_a}:$AB${tr_b},SelectedLeader,"
                f"_Long!$AC${tr_a}:$AC${tr_b},$"
                f"{xlsxwriter.utility.xl_col_to_name(cX)}{i + 1})")
    ws_cube.hide()

    def cube_lookup(col_name: str) -> str:
        ci = cube_cols.index(col_name)
        col = xlsxwriter.utility.xl_col_to_name(ci)
        return (f"XLOOKUP(ExclToggle&\"|\"&SelectedLeader,"
                f"_Cube!$A$2:$A${n_cube + 1},"
                f"_Cube!${col}$2:${col}${n_cube + 1})")

    # KPI numbers that have a person-level detail sheet become clickable:
    # HYPERLINK keeps the live cube value AND jumps to the listed sheet.
    KPI_LINKS = {
        "Members":              "Training Status",
        "FullyTrained":         "Training Status",
        "FullyRegistered":      "Data",
        "UnregUsers":           "Reg Recommendations",
        "NoShowUsers":          "No Shows",
        "NoEpicCurriculum":     "Training Status",
        "Classes":              "Class Tracking",
        "UnregSessions":        "Classes to Schedule",
        "PlaceUnregSessions":   "Reg Recommendations",
        "PlaceNoShowSessions":  "No Shows",
        "PlaceUnregTMs":        "Reg Recommendations",
        "PlaceNoShowTMs":       "No Shows",
    }

    def kpi_formula(col: str) -> str:
        base = cube_lookup(col)
        tgt = KPI_LINKS.get(col)
        if not tgt:
            return f"={base}"
        return f"=HYPERLINK(\"#'{tgt}'!A1\",{base})"

    # ---- hidden: _Long -----------------------------------------------------
    ws_long = wb.add_worksheet("_Long")
    blocks = [
        (d["long_curr"], 0),    # A..E
        (d["long_class"], 6),   # G..P
        (d["long_mgr"], 17),    # R..Y
        (d["long_trend"], 26),  # AA..AE
        (d["audit"], 32),       # AG..AN (registration audit rows)
    ]
    for df, c0 in blocks:
        for c, name in enumerate(df.columns):
            ws_long.write_string(0, c0 + c, str(name))
        for r, row in enumerate(df.itertuples(index=False), start=1):
            for c, v in enumerate(row):
                _w(ws_long, r, c0 + c, v, fm.cell, fm.dt)
    ws_long.hide()

    def mask(a_col: str, b_col: str, n: int) -> str:
        return (f"(_Long!${a_col}$2:${a_col}${n + 1}=ExclToggle)*"
                f"(_Long!${b_col}$2:${b_col}${n + 1}=SelectedLeader)")

    # ---- hidden: _Comp (Completion Status learner rows) --------------------
    # One row per learner; column A = Leader for this sheet's OWN filter
    # (independent of the Control dropdown), trailing Leader copy for display.
    comp_cols = ["Leader", "UniversalID", "FullName", "FullyRegisteredYN",
                 "FullyTrainedYN", "LastRegisteredClass", "FinalScheduledDate",
                 "CompletionDate", "ClassesRegistered", "ClassesCompleted",
                 "UnregisteredClasses", "PendingAttendanceClasses"]
    comp = d["tstatus"][comp_cols].copy()
    # blank cells spill as 0 through Excel FILTER — store "" instead
    comp = comp.astype(object).where(comp.notna(), "")
    ws_comp = wb.add_worksheet("_Comp")
    for c, name in enumerate(comp_cols + ["LeaderTail"]):
        ws_comp.write_string(0, c, name)
    for r, crow in enumerate(comp.itertuples(index=False), start=1):
        for c, v in enumerate(crow):
            _w(ws_comp, r, c, v, fm.cell, fm.dt)
        _w(ws_comp, r, len(comp_cols), crow[0], fm.cell, fm.dt)
    ws_comp.hide()
    n_comp = len(comp)

    # ---- Control -----------------------------------------------------------
    ws = wb.add_worksheet("Control")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 3)
    ws.set_column(1, 1, 26)
    ws.set_column(2, 2, 30)
    ws.set_column(3, 5, 24)
    ws.write_string(1, 1, f"{TITLE} — Training Tracker", fm.title)
    ws.write_string(2, 1, "Pick a leader and scope below. Sheets marked 🔎 "
                          "follow this selection; sheets marked 🌐 always show "
                          "all leaders.", fm.subtitle)
    ws.write_string(4, 1, "Leader", fm.h2)
    ws.write_string(4, 2, "Class scope", fm.h2)
    ws.write_string(5, 1, "All", fm.ctrl)
    ws.write_string(5, 2, "Exclude", fm.ctrl)
    ws.data_validation(5, 1, 5, 1, {
        "validate": "list",
        "source": f"=_Lists!$A$2:$A${len(leader_opts) + 1}"})
    ws.data_validation(5, 2, 5, 2, {"validate": "list",
                                    "source": "=_Lists!$B$2:$B$3"})
    ws.write_string(6, 2, "Exclude = leave optional Advanced Reporting / "
                          "Charge Capture workshops out of the numbers.", fm.note)
    wb.define_name("SelectedLeader", "=Control!$B$6")
    wb.define_name("ExclToggle", "=Control!$C$6")

    NAV = [
        ("Exec Dashboard", "🔎 KPI cards, weekly sessions chart, leader roll-up."),
        ("Visuals", "🌐 One-page visual dashboard — all leaders, standard scope."),
        ("Summary", "🔎 KPIs plus curriculum and class breakdowns for the selected "
                    "leader (its Leader Summary table always shows all leaders)."),
        ("Leader Tracking", "🔎 One leader's class status and open items."),
        ("Manager Tracking", "🔎 Manager roll-up under the selected leader."),
        ("Class Tracking", "🔎 Class-by-class roll-up for the selected leader."),
        ("Classes to Schedule", "🌐 Classes people still need a seat in, worst first."),
        ("Reg Recommendations", "🌐 Who to register where — top 3 session options per person."),
        ("Registration Schedule", "🌐 Every scheduled session with seats and possible placements."),
        ("Unregistered Summary", "🔎 Top-line unregistered and no-show numbers."),
        ("Registration Audit", "🔎 Hygiene flags — out-of-scope seats, schedule "
                               "conflicts, duplicates, wrong-track curricula."),
        ("Training Status", "🌐 Not Started / In Progress / Fully Trained by leader + "
                            "per-learner last & final session dates."),
        ("Completion Status", "🎚 Per-person registered/trained Yes-No, last booked "
                              "class, expected completion — own leader filter on "
                              "the sheet."),
        ("No Shows", "🌐 Person-level no-show list — standing and resolved."),
        ("Go and Finds", "🌐 Assigned GNF modules by leader and learner."),
        ("Data", "🌐 Raw person × class detail — filter anything."),
        ("Glossary", "Reference — what each status means and how every number is calculated."),
    ]
    ws.write_string(8, 1, "Navigation", fm.h2)
    for i, (sheet, desc) in enumerate(NAV, start=9):
        ws.write_url(i, 1, f"internal:'{sheet}'!A1", fm.link, sheet)
        ws.write_string(i, 2, desc, fm.cell)
    r_nav = 9 + len(NAV)
    ws.write_url(r_nav, 1, "external:RCM Training Daily Log.xlsx", fm.link,
                 "RCM Training Daily Log.xlsx")
    ws.write_string(r_nav, 2, "Separate workbook, same folder — daily & weekly "
                              "metric history and the unregistered trend matrix.",
                    fm.cell)

    r8 = r_nav + 2
    ws.write_string(r8, 1, "Data as of", fm.h2)
    ws.write_string(r8 + 1, 1, "Feed", fm.th)
    ws.write_string(r8 + 1, 2, "Source file", fm.th)
    ws.write_string(r8 + 1, 3, "Loaded", fm.th)
    for i, r in enumerate(d["fresh"].itertuples(index=False), start=r8 + 2):
        ws.write_string(i, 1, str(r.Feed), fm.cell)
        ws.write_string(i, 2, Path(str(r.SourceFile or "")).name, fm.cell)
        _w(ws, i, 3, pd.to_datetime(r.LoadedAt, errors="coerce"), fm.dtm, fm.dtm)
    r0 = r8 + 2 + len(d["fresh"]) + 1
    ws.write_string(r0, 1, "Workbook rebuilt", fm.cell)
    ws.write_datetime(r0, 2, built_at, fm.dtm)
    ws.write_string(r0 + 1, 1, "Go-live", fm.cell)
    ws.write_datetime(r0 + 1, 2, datetime.combine(GO_LIVE, datetime.min.time()), fm.dt)
    ws.write_string(r0 + 3, 1,
                    "Refresh: run the tracker build script — this file is fully "
                    "regenerated from the latest exports; do not paste data here.",
                    fm.note)

    # ---- Glossary ----------------------------------------------------------
    ws = wb.add_worksheet("Glossary")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 2, 26)
    ws.set_column(3, 3, 88)
    ws.write_string(1, 1, "Glossary", fm.title)
    ws.write_string(2, 1, "What the status labels mean, and how every number "
                          "in this workbook is determined.", fm.subtitle)
    ws.write_string(3, 1, f"Reference — definitions only, no data  ·  {asof}",
                    fm.note)

    def gsection(r, heading, cols, rows_):
        ws.write_string(r, 1, heading, fm.h2)
        for c, h in enumerate(cols):
            ws.write_string(r + 1, 1 + c, h, fm.th)
        for i, row in enumerate(rows_, start=r + 2):
            for c, v in enumerate(row):
                ws.write_string(i, 1 + c, v, fm.gterm if c == 0 else fm.gdef)
        return r + 2 + len(rows_) + 1

    r = 4
    r = gsection(r, "STATUS LABELS", ["Label", "Where it appears", "Meaning"], [
        ("🟢 On track", "Class Status / Class Tracking",
         "Everyone who needs this class is registered or done — no open gaps."),
        ("🟠 Gaps to place", "Class Status / Class Tracking",
         "Some people still need this class, and future sessions exist to book them into."),
        ("🔴 Gaps, no future session", "Class Status / Class Tracking",
         "People still need this class and NOTHING is left on the schedule — escalate."),
        ("✅ Completed", "Class Status / Class Tracking",
         "Every user who needs this class has completed it."),
        ("🔴 Needs scheduling", "Classes to Schedule",
         "No future session of this class exists at all — a session must be added."),
        ("🟠 Future sessions full", "Classes to Schedule",
         "Future sessions exist but every seat is taken — waitlist or add seats."),
        ("🟢 Has future session", "Classes to Schedule",
         "At least one future session has open seats — bookable today."),
        ("Unregistered", "Reg Recommendations (Reason)",
         "Epic shows a required class with no registration for this person."),
        ("No-Show — Re-register", "Reg Recommendations (Reason)",
         "Person no-showed a session and has nothing newer booked — needs a new seat."),
        ("✅ Session 1", "Reg Recommendations (Assignment)",
         "A bookable option exists — register them into Option 1 (earliest open session)."),
        ("⚠️ No future sessions", "Reg Recommendations (Assignment)",
         "No future session with open seats — pair with Classes to Schedule."),
        ("✅ Yes / ⛔ No", "Fully Trained? / Fully Registered?",
         "Epic's own person-level flags from the Curriculum Status Detail report."),
        ("Live", "Days to Go-Live card",
         "The go-live date has been reached."),
    ])
    r = gsection(r, "HOW THE NUMBERS ARE CALCULATED",
                 ["Metric", "Source", "How it's determined"], [
        ("Team Members", "Master Wave File",
         "Distinct people on the wave file, in scope, whose Leader is on the "
         "canonical 16-leader list. Every metric in this workbook counts ONLY "
         "these people."),
        ("Fully Registered / Fully Trained", "Epic status detail",
         "Epic's person-level flags — counted as Yes when any of the person's "
         "rows says Yes."),
        ("Unregistered", "Epic status detail",
         "Required session classes where Event Registered = 'No'. Users = "
         "distinct people with at least one such class; Sessions = unique "
         "person × class gaps."),
        ("No-Shows (standing)", "Cornerstone transcripts",
         "Person × class where the NEWEST live transcript row is still 'No "
         "Show' — nothing newer was booked. Resolved no-shows drop out."),
        ("No Epic Curriculum", "Epic status detail",
         "On our roster but ZERO rows in Epic's status report — they cannot "
         "register until Epic assigns a curriculum."),
        ("Demand / Classes to Schedule", "Epic + Cornerstone",
         "Unregistered gaps PLUS standing no-shows that need re-registration, "
         "grouped by class."),
        ("Top 3 options / Recommended", "Class schedule export",
         "The three earliest future sessions (tomorrow onward) of that class "
         "with open seats, by date."),
        ("Last Class Date", "Epic status detail",
         "The latest session date attached to this population — past OR "
         "future — i.e. the scheduled final session."),
        ("Sessions Attended / Registered — Future Sessions", "Epic status detail",
         "The weekly chart. Dark = class sessions our people attended that "
         "week (Epic event status Completed). Light = future-dated sessions "
         "they are registered for. July 2026 onward. Each person's session is "
         "counted ONCE even when Epic lists the class under several "
         "curricula. 'Attended' counts sessions — a person is only Fully "
         "Trained once every class they need is attended. Sessions still "
         "needing registration are NOT on the chart (see the note under it)."),
        ("Class scope toggle (Exclude)", "Control sheet",
         "Exclude leaves the optional Advanced Reporting / Charge Capture "
         "workshops out of every metric; Include counts them."),
        ("Data as of", "Control sheet",
         "The exact export files and load times behind this build. The whole "
         "workbook is regenerated from them — nothing is typed in by hand."),
        ("Not Started / In Progress", "Training Status sheet",
         "Not Started = zero required classes attended yet. In Progress = "
         "attended at least one required class but not yet Fully Trained. "
         "Together with Fully Trained and No Epic Curriculum these four "
         "buckets partition every team member."),
        ("Final Scheduled (role completion date)", "Training Status sheet",
         "The learner's LAST session date — attended or booked. When all "
         "classes are attended it is their completion date; otherwise their "
         "projected finish. Blank = nothing scheduled. Learners finishing "
         "close to go-live are candidates to shift earlier."),
        ("Go and Finds (GNF)", "Go and Finds sheet",
         "Self-directed modules assigned per learner in the W3 GNF mapping "
         "file. Currently assignments only — completion tracking is added "
         "when that data becomes available."),
        ("Pending Attendance", "Completion Status sheet",
         "Classes the person is registered for whose session date has not "
         "happened yet — registered but not yet complete. Pending "
         "Registration = required classes with no booking at all. The "
         "Completion Status sheet has its own leader dropdown at the top, "
         "independent of the Control sheet."),
    ])
    ws.write_url(r, 1, "internal:'Control'!A1", fm.link, "← Back to Control")

    # ---- Exec Dashboard ----------------------------------------------------
    ws = wb.add_worksheet("Exec Dashboard")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 12, 15)
    ws.write_string(1, 1, f"{TITLE} — Executive Dashboard", fm.title)
    stamp_live(ws, 2)
    ws.write_string(3, 1, "Blue KPI numbers are clickable — they jump to the "
                          "sheet listing that population.", fm.note)
    cards = [
        ("TEAM MEMBERS", "Members", None, "In scope"),
        ("FULLY TRAINED", "FullyTrained", "PctTrained", "Completed all classes"),
        ("FULLY REGISTERED", "FullyRegistered", "PctRegistered", "Registered / pending"),
        ("UNREGISTERED", "UnregUsers", "PctUnregUsers", "Action needed"),
        ("NO-SHOWS", "NoShowUsers", None, "Standing no-shows"),
        ("NO EPIC CURRICULUM", "NoEpicCurriculum", None, "Cannot register yet"),
        ("DAYS TO GO-LIVE", "DaysToGoLive", None, GO_LIVE.strftime("%b %d, %Y")),
        ("LAST CLASS DATE", "LastClassDate", None, "Scheduled final session"),
    ]
    row_a, row_b = 4, 9
    for i, (label, col, pct_col, caption) in enumerate(cards):
        rr = row_a if i < 4 else row_b
        cc = 1 + (i % 4) * 3
        ws.merge_range(rr, cc, rr, cc + 1, label, fm.label)
        if col == "LastClassDate":
            ws.merge_range(
                rr + 1, cc, rr + 1, cc + 1,
                f"=IFERROR(TEXT({cube_lookup(col)},\"m/d/yyyy\"),\"—\")", fm.kpi_txt)
        else:
            linked = col in KPI_LINKS
            if col == "DaysToGoLive":
                f = fm.kpi_txt
            else:
                f = fm.kpi_link if linked else fm.kpi
            ws.merge_range(rr + 1, cc, rr + 1, cc + 1, kpi_formula(col), f)
        if pct_col:
            ws.merge_range(rr + 2, cc, rr + 2, cc + 1,
                           f"={cube_lookup(pct_col)}", fm.kpi_pct)
        else:
            ws.merge_range(rr + 2, cc, rr + 2, cc + 1, caption, fm.label)
        ws.set_row(rr + 1, 30)

    # progress donuts — % of team members, live with the leader dropdown
    ws.write_string(13, 1, "PROGRESS VS TOTAL POPULATION", fm.h2)
    ws.write_string(13, 5, "(% of team members in the current view)", fm.note)
    donuts = [
        ("Fully Registered", f"={cube_lookup('PctRegistered')}", ACCENT),
        ("Fully Trained", f"={cube_lookup('PctTrained')}", "#4E9A6F"),
        ("Unregistered", f"={cube_lookup('PctUnregUsers')}", "#E8A33D"),
        ("No-Show", f"=IFERROR({cube_lookup('NoShowUsers')}/"
                    f"{cube_lookup('Members')},0)", "#C4614D"),
    ]
    RING_REST = "#E8E7E4"
    for j, (label, pct_formula, color) in enumerate(donuts):
        fr = 14 + j            # feed row (0-based) in hidden cols U:V
        ws.write_formula(fr, 20, pct_formula)
        ws.write_formula(fr, 21, f"=MAX(0,1-U{fr + 1})")
        dch = wb.add_chart({"type": "doughnut"})
        dch.add_series({
            "values": f"='Exec Dashboard'!$U${fr + 1}:$V${fr + 1}",
            "points": [{"fill": {"color": color}, "border": {"none": True}},
                       {"fill": {"color": RING_REST}, "border": {"none": True}}],
        })
        dch.set_hole_size(62)
        dch.set_legend({"none": True})
        dch.set_chartarea({"border": {"none": True}})
        dch.set_plotarea({"layout": {"x": 0.08, "y": 0.08,
                                     "width": 0.84, "height": 0.84}})
        dch.set_size({"width": 165, "height": 165})
        dch.show_hidden_data()   # feed cells live in hidden columns U:V
        cc = 1 + j * 2
        ws.insert_chart(14, cc, dch, {"x_offset": 18, "y_offset": 4})
        pct_big = wb.add_format({"font_name": "Segoe UI", "font_size": 18,
                                 "bold": True, "font_color": color,
                                 "align": "center", "num_format": "0.0%"})
        ws.merge_range(23, cc, 23, cc + 1, f"=U{fr + 1}", pct_big)
        ws.merge_range(24, cc, 24, cc + 1, label, fm.label)
    ws.set_column(20, 21, None, None, {"hidden": True})

    # weekly training calendar: attended (past) + registered (future) sessions —
    # one unit, one axis; the current week carries both segments
    ws.write_string(26, 1, "TRAINING SESSIONS BY WEEK", fm.h2)
    if weeks:
        chart = wb.add_chart({"type": "column", "subtype": "stacked"})
        colX = xlsxwriter.utility.xl_col_to_name(cX)
        colY = xlsxwriter.utility.xl_col_to_name(cY)
        colZ = xlsxwriter.utility.xl_col_to_name(cZ)
        last = len(weeks) + 1
        chart.add_series({
            "name": "Sessions Attended",
            "categories": f"=_Cube!${colX}$2:${colX}${last}",
            "values": f"=_Cube!${colZ}$2:${colZ}${last}",
            "fill": {"color": ACCENT},
            "border": {"none": True},
            "gap": 35,
        })
        chart.add_series({
            "name": "Registered — Future Sessions",
            "categories": f"=_Cube!${colX}$2:${colX}${last}",
            "values": f"=_Cube!${colY}$2:${colY}${last}",
            "fill": {"color": "#B3CCE8"},
            "border": {"none": True},
        })
        chart.set_title({"name": "Training sessions by week — attended and "
                                 "registered for future sessions",
                         "name_font": {"name": "Segoe UI", "size": 11,
                                       "color": INK, "bold": False}})
        chart.set_legend({"position": "bottom",
                          "font": {"name": "Segoe UI", "size": 9}})
        chart.set_x_axis({"num_font": {"name": "Segoe UI", "size": 9},
                          "num_format": "m/d", "text_axis": True,
                          "line": {"color": LINE}})
        chart.set_y_axis({"num_font": {"name": "Segoe UI", "size": 9},
                          "min": 0,
                          "major_gridlines": {"visible": True,
                                              "line": {"color": LINE}}})
        chart.set_chartarea({"border": {"none": True}})
        chart.set_plotarea({"border": {"none": True}})
        chart.set_size({"width": 700, "height": 280})
        ws.insert_chart(27, 1, chart)
        ws.write_formula(
            41, 1,
            f"=\"Not shown: \"&{cube_lookup('UnregSessions')}&\" required "
            f"session(s) with no registration yet — see Classes to Schedule.\"",
            fm.note)

    # leader roll-up (static, Exclude scope)
    r0 = 43
    ws.write_string(r0, 1, "LEADER ROLL-UP", fm.h2)
    ws.write_string(r0, 4, "(standard scope — excludes optional workshops)", fm.note)
    excl = cube[(cube["Variant"] == "Exclude") & (cube["Leader"] != "All")]
    roll = excl[["Leader", "Members", "FullyTrained", "PctTrained",
                 "FullyRegistered", "PctRegistered", "UnregUsers",
                 "NoShowUsers", "LastClassDate"]].copy()
    roll.columns = ["Leader", "Members", "Trained", "Trained %", "Registered",
                    "Registered %", "Unregistered", "No-Shows", "Last Class"]
    roll = roll.sort_values("Members", ascending=False)
    hdr = list(roll.columns)
    for c, name in enumerate(hdr):
        ws.write_string(r0 + 1, 1 + c, name, fm.th)
    for i, row in enumerate(roll.itertuples(index=False), start=r0 + 2):
        for c, v in enumerate(row):
            f = fm.pct if "%" in hdr[c] else (fm.num if c else fm.cell)
            _w(ws, i, 1 + c, v, f, fm.dt)

    # next sessions
    # class names get a merged 3-cell range instead of a widened column —
    # widening col C would stretch the first progress donut's anchor above
    r1 = r0 + len(roll) + 4
    ws.write_string(r1, 1, "NEXT SESSIONS ON THE SCHEDULE", fm.h2)
    up = d["upcoming"].copy()
    up.columns = ["Date", "Class", "Location", "Room", "Start", "Open Seats"]
    cols_map = [1, 2, 5, 6, 7, 8]        # B, C:E merged, F, G, H, I
    for c, name in enumerate(up.columns):
        if c == 1:
            ws.merge_range(r1 + 1, 2, r1 + 1, 4, name, fm.th)
        else:
            ws.write_string(r1 + 1, cols_map[c], name, fm.th)
    for i, row in enumerate(up.itertuples(index=False), start=r1 + 2):
        for c, v in enumerate(row):
            if c == 1:
                ws.merge_range(i, 2, i, 4, "" if v is None else str(v), fm.cell)
            else:
                _w(ws, i, cols_map[c], v, fm.num if c == 5 else fm.cell, fm.dt)

    # ---- Visuals (one-page dashboard, all leaders, standard scope) ---------
    vz = d["viz"]
    ws_vf = wb.add_worksheet("_Viz")          # hidden static chart feed
    blocks = [(0, vz["leader_reg"]), (4, pd.DataFrame(vz["mix"],
                                                      columns=["Label", "Value"])),
              (7, vz["trend"]), (11, vz["weekly"]),
              (15, vz["leader_unreg"]), (18, vz["status_leader"])]
    for c0, df in blocks:
        for c, name in enumerate(df.columns):
            ws_vf.write_string(0, c0 + c, str(name))
        for r, row in enumerate(df.itertuples(index=False), start=1):
            for c, v in enumerate(row):
                _w(ws_vf, r, c0 + c, v, fm.cell, fm.dt)
    ws_vf.hide()
    n_lr, n_mx = len(vz["leader_reg"]), len(vz["mix"])
    n_tr, n_wk, n_un = len(vz["trend"]), len(vz["weekly"]), len(vz["leader_unreg"])
    n_st = len(vz["status_leader"])

    ws = wb.add_worksheet("Visuals")
    ws.hide_gridlines(2)
    PAGE_BG = "#F6F5F3"
    bgf = wb.add_format({"bg_color": PAGE_BG})
    ws.set_column(0, 0, 2, bgf)
    ws.set_column(1, 23, 9, bgf)
    ws.set_column(24, 40, 9, bgf)
    _white = {"bg_color": "#FFFFFF"}
    _edge_fmts = {}

    def box_fmt(top, bottom, left, right):
        key = (top, bottom, left, right)
        if key not in _edge_fmts:
            spec = dict(_white)
            if top:
                spec.update({"top": 1, "top_color": LINE})
            if bottom:
                spec.update({"bottom": 1, "bottom_color": LINE})
            if left:
                spec.update({"left": 1, "left_color": LINE})
            if right:
                spec.update({"right": 1, "right_color": LINE})
            _edge_fmts[key] = wb.add_format(spec)
        return _edge_fmts[key]

    tile_title = wb.add_format({"font_name": "Segoe UI", "font_size": 11,
                                "bold": True, "font_color": INK,
                                "bg_color": "#FFFFFF"})

    def tile(r1, c1, r2, c2, title=None):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.write_blank(r, c, None,
                               box_fmt(r == r1, r == r2, c == c1, c == c2))
        if title:
            ws.write_string(r1 + 1, c1 + 1, title, tile_title)

    pt_f = wb.add_format({"font_name": "Segoe UI", "font_size": 18, "bold": True,
                          "font_color": INK, "bg_color": PAGE_BG})
    ps_f = wb.add_format({"font_name": "Segoe UI", "font_size": 9,
                          "font_color": GRAY, "bg_color": PAGE_BG})
    ws.write_string(1, 1, f"{TITLE} — Visual Dashboard", pt_f)
    ws.write_string(2, 1, f"🌐 All leaders — the Control leader filter does "
                          f"not apply here · standard scope (optional workshops "
                          f"excluded) · {asof}", ps_f)

    # KPI band
    allx = vz["all"]
    dl = vz["deltas"]

    def delta_txt(key, invert=False):
        v = dl.get(key)
        if v is None or v != v:
            return ("—", GRAY)
        v = int(v)
        if v == 0:
            return ("· no change today", GRAY)
        good = (v < 0) if invert else (v > 0)
        arrow = "▲" if v > 0 else "▼"
        return (f"{arrow} {v:+d} today", "#4E9A6F" if good else "#C4614D")

    kpis = [
        ("TEAM MEMBERS", int(allx["Members"]), ("in scope", GRAY)),
        ("FULLY REGISTERED", int(allx["FullyRegistered"]), delta_txt("ΔReg")),
        ("FULLY TRAINED", int(allx["FullyTrained"]), delta_txt("ΔTrained")),
        ("UNREGISTERED", int(allx["UnregUsers"]), delta_txt("ΔUnreg", True)),
        ("NO-SHOWS", int(allx["NoShowUsers"]), delta_txt("ΔNoShow", True)),
        ("DAYS TO GO-LIVE", allx["DaysToGoLive"],
         (GO_LIVE.strftime("%b %d, %Y"), GRAY)),
    ]
    side = {"left": 1, "left_color": LINE, "right": 1, "right_color": LINE,
            "bg_color": "#FFFFFF", "font_name": "Segoe UI", "align": "center"}
    lab_f = wb.add_format({**side, "font_size": 8, "font_color": GRAY})
    val_f = wb.add_format({**side, "font_size": 20, "bold": True,
                           "font_color": INK})
    for i, (label, value, (dtxt, dcolor)) in enumerate(kpis):
        c1 = 1 + i * 4
        for c in range(c1, c1 + 3):
            ws.write_blank(4, c, None, box_fmt(True, False, c == c1, c == c1 + 2))
        ws.merge_range(5, c1, 5, c1 + 2, label, lab_f)
        if isinstance(value, (int, float)):
            ws.merge_range(6, c1, 6, c1 + 2, value, val_f)
        else:
            ws.merge_range(6, c1, 6, c1 + 2, str(value), val_f)
        df_ = wb.add_format({**side, "font_size": 9, "font_color": dcolor,
                             "bottom": 1, "bottom_color": LINE})
        ws.merge_range(7, c1, 7, c1 + 2, dtxt, df_)
        ws.set_row(6, 28)

    def style_chart(ch, legend="bottom"):
        ch.set_title({"none": True})
        if legend:
            ch.set_legend({"position": legend,
                           "font": {"name": "Segoe UI", "size": 9}})
        else:
            ch.set_legend({"none": True})
        ch.set_chartarea({"fill": {"color": "#FFFFFF"}, "border": {"none": True}})
        ch.set_plotarea({"border": {"none": True}})

    axis_font = {"num_font": {"name": "Segoe UI", "size": 9},
                 "line": {"color": LINE}}
    grid = {"major_gridlines": {"visible": True, "line": {"color": LINE}}}

    # T1 — registration progress by leader (stacked horizontal bars)
    tile(9, 1, 26, 8, "Registration Progress by Leader")
    ch = wb.add_chart({"type": "bar", "subtype": "stacked"})
    ch.add_series({"name": "Fully Registered",
                   "categories": f"=_Viz!$A$2:$A${n_lr + 1}",
                   "values": f"=_Viz!$B$2:$B${n_lr + 1}",
                   "fill": {"color": ACCENT}, "border": {"none": True},
                   "gap": 40,
                   "data_labels": {"value": True,
                                   "font": {"name": "Segoe UI", "size": 8,
                                            "color": "#FFFFFF"}}})
    ch.add_series({"name": "Not Yet Registered",
                   "categories": f"=_Viz!$A$2:$A${n_lr + 1}",
                   "values": f"=_Viz!$C$2:$C${n_lr + 1}",
                   "fill": {"color": RING_REST}, "border": {"none": True}})
    style_chart(ch)
    ch.set_x_axis({**axis_font, **grid, "min": 0})
    ch.set_y_axis({**axis_font})
    ch.set_size({"width": 500, "height": 300})
    ws.insert_chart(11, 1, ch, {"x_offset": 8, "y_offset": 2})

    # T2 — population status mix (donut)
    tile(9, 10, 26, 15, "Population Status Mix")
    ch = wb.add_chart({"type": "doughnut"})
    ch.add_series({
        "categories": f"=_Viz!$E$2:$E${n_mx + 1}",
        "values": f"=_Viz!$F$2:$F${n_mx + 1}",
        "points": [{"fill": {"color": ACCENT}, "border": {"none": True}},
                   {"fill": {"color": "#E8A33D"}, "border": {"none": True}},
                   {"fill": {"color": "#A8A29B"}, "border": {"none": True}},
                   {"fill": {"color": "#D8D5D0"}, "border": {"none": True}}],
        "data_labels": {"percentage": True,
                        "font": {"name": "Segoe UI", "size": 9,
                                 "color": INK}},
    })
    ch.set_hole_size(55)
    style_chart(ch, legend="right")
    ch.set_size({"width": 385, "height": 300})
    ws.insert_chart(11, 10, ch, {"x_offset": 6, "y_offset": 2})

    # T3 — daily registration trend (line)
    tile(9, 17, 26, 23, "Daily Trend — Registered & Trained")
    ch = wb.add_chart({"type": "line"})
    ch.add_series({"name": "Fully Registered",
                   "categories": f"=_Viz!$H$2:$H${n_tr + 1}",
                   "values": f"=_Viz!$I$2:$I${n_tr + 1}",
                   "line": {"color": ACCENT, "width": 2.5}})
    ch.add_series({"name": "Fully Trained",
                   "categories": f"=_Viz!$H$2:$H${n_tr + 1}",
                   "values": f"=_Viz!$J$2:$J${n_tr + 1}",
                   "line": {"color": "#4E9A6F", "width": 2}})
    style_chart(ch)
    ch.set_x_axis({**axis_font, "num_format": "m/d", "date_axis": True,
                   "major_unit": 3, "major_unit_type": "days"})
    ch.set_y_axis({**axis_font, **grid, "min": 0})
    ch.set_size({"width": 450, "height": 300})
    ws.insert_chart(11, 17, ch, {"x_offset": 6, "y_offset": 2})
    tile_note = wb.add_format({"font_name": "Segoe UI", "font_size": 8,
                               "italic": True, "font_color": GRAY,
                               "bg_color": "#FFFFFF"})
    ws.write_string(25, 18, "Mid-July dip = Cornerstone withdrawal incident "
                            "(recovered).", tile_note)

    # T4 — training sessions by week (static copy of the exec chart)
    tile(28, 1, 45, 13, "Training Sessions by Week — Attended & Registered Ahead")
    ch = wb.add_chart({"type": "column", "subtype": "stacked"})
    ch.add_series({"name": "Sessions Attended",
                   "categories": f"=_Viz!$L$2:$L${n_wk + 1}",
                   "values": f"=_Viz!$N$2:$N${n_wk + 1}",
                   "fill": {"color": ACCENT}, "border": {"none": True},
                   "gap": 35})
    ch.add_series({"name": "Registered — Future Sessions",
                   "categories": f"=_Viz!$L$2:$L${n_wk + 1}",
                   "values": f"=_Viz!$M$2:$M${n_wk + 1}",
                   "fill": {"color": "#B3CCE8"}, "border": {"none": True}})
    style_chart(ch)
    ch.set_x_axis({**axis_font, "num_format": "m/d", "text_axis": True})
    ch.set_y_axis({**axis_font, **grid, "min": 0})
    ch.set_size({"width": 830, "height": 300})
    ws.insert_chart(30, 1, ch, {"x_offset": 8, "y_offset": 2})

    # T5 — unregistered users by leader (bars)
    tile(28, 15, 45, 23, "Unregistered Users by Leader")
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"categories": f"=_Viz!$P$2:$P${n_un + 1}",
                   "values": f"=_Viz!$Q$2:$Q${n_un + 1}",
                   "fill": {"color": "#E8A33D"}, "border": {"none": True},
                   "gap": 40,
                   "data_labels": {"value": True,
                                   "font": {"name": "Segoe UI", "size": 9,
                                            "color": INK}}})
    style_chart(ch, legend=None)
    ch.set_x_axis({**axis_font, **grid, "min": 0})
    ch.set_y_axis({**axis_font})
    ch.set_size({"width": 550, "height": 300})
    ws.insert_chart(30, 15, ch, {"x_offset": 6, "y_offset": 2})

    # T6 — training status by leader (100% stacked)
    tile(47, 1, 64, 13, "Training Status by Leader — % Not Started / "
                        "In Progress / Fully Trained")
    ch = wb.add_chart({"type": "bar", "subtype": "percent_stacked"})
    for label, colL, color in (("Not Started", "T", "#D8D5D0"),
                               ("In Progress", "U", "#B3CCE8"),
                               ("Fully Trained", "V", "#4E9A6F")):
        ch.add_series({"name": label,
                       "categories": f"=_Viz!$S$2:$S${n_st + 1}",
                       "values": f"=_Viz!${colL}$2:${colL}${n_st + 1}",
                       "fill": {"color": color}, "border": {"none": True},
                       "gap": 40})
    style_chart(ch)
    ch.set_x_axis({**axis_font, "num_format": "0%", "line": {"color": LINE}})
    ch.set_y_axis({**axis_font})
    ch.set_size({"width": 830, "height": 300})
    ws.insert_chart(49, 1, ch, {"x_offset": 8, "y_offset": 2})

    # ---- Summary -----------------------------------------------------------
    ws = wb.add_worksheet("Summary")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 52)
    ws.set_column(2, 3, 14)
    ws.set_column(4, 4, 3)
    ws.set_column(5, 5, 52)
    ws.set_column(6, 7, 14)
    ws.set_column(8, 8, 3)
    ws.set_column(9, 14, 16)
    ws.write_string(1, 1, f"{TITLE} — Training Summary", fm.title)
    stamp_live(ws, 2)
    ws.write_string(3, 1, "Blue KPI numbers are clickable — they jump to the "
                          "sheet listing that population.", fm.note)
    kpis = [("Team Members", "Members"), ("Fully Trained", "FullyTrained"),
            ("Fully Registered", "FullyRegistered"),
            ("Unregistered", "UnregUsers"), ("No-Shows", "NoShowUsers"),
            ("Curriculums", "Curricula"), ("Classes", "Classes"),
            ("Sessions Attended", "SessionsCompleted")]
    for i, (label, col) in enumerate(kpis):
        ws.write_string(4, 1 + i, label, fm.label)
        f = fm.kpi_txt_link if col in KPI_LINKS else fm.kpi_txt
        ws.write_formula(5, 1 + i, kpi_formula(col), f)

    r0 = 7
    ws.write_string(r0, 1, "CURRICULUM BREAKDOWN", fm.h2)
    ws.write_string(r0 + 1, 1, "Curriculum", fm.th)
    ws.write_string(r0 + 1, 2, "# Users", fm.th)
    ws.write_string(r0 + 1, 3, "Last Session", fm.th)
    if n_curr:
        m1 = mask("A", "B", n_curr)
        ws.write_dynamic_array_formula(
            r0 + 2, 1, r0 + 2, 1,
            f"=IFERROR(SORT(FILTER(_Long!$C$2:$E${n_curr + 1},{m1},\"\"),2,-1),\"\")",
            fm.cell)
    ws.write_string(r0, 5, "CLASS BREAKDOWN", fm.h2)
    for c, name in enumerate(["Class", "# Users", "Last Session"]):
        ws.write_string(r0 + 1, 5 + c, name, fm.th)
    if n_class:
        m2 = mask("G", "H", n_class)
        ws.write_dynamic_array_formula(
            r0 + 2, 5, r0 + 2, 5,
            f"=IFERROR(SORT(FILTER(CHOOSECOLS(_Long!$I$2:$N${n_class + 1},1,2,6),"
            f"{m2},\"\"),2,-1),\"\")", fm.cell)
    ws.write_string(r0, 9, "LEADER SUMMARY — ALL LEADERS (ignores the filter)",
                    fm.h2)
    lsum = excl[["Leader", "Members", "FullyRegistered", "UnregUsers",
                 "NoShowUsers", "LastClassDate"]].copy()
    lsum.columns = ["Leader", "Members", "Fully Registered", "Unregistered",
                    "No-Shows", "Last Class Date"]
    for c, name in enumerate(lsum.columns):
        ws.write_string(r0 + 1, 9 + c, name, fm.th)
    for i, row in enumerate(lsum.sort_values("Members", ascending=False)
                            .itertuples(index=False), start=r0 + 2):
        for c, v in enumerate(row):
            _w(ws, i, 9 + c, v, fm.num if c else fm.cell, fm.dt)
    # date formats on spilled Last Session columns
    ws.set_column(3, 3, 14, fm.dt)
    ws.set_column(7, 7, 14, fm.dt)
    ws.set_column(14, 14, 16, fm.dt)

    # ---- Leader Tracking ---------------------------------------------------
    ws = wb.add_worksheet("Leader Tracking")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 56)
    ws.set_column(2, 10, 14)
    ws.write_formula(1, 1, "=SelectedLeader&\" — Completion Tracking\"", fm.title)
    ws.write_formula(
        2, 1,
        "=IF(SelectedLeader=\"All\",\"Viewing every leader — pick one on the "
        "Control sheet for a focused view.\","
        "IF(N5+N6=0,\"✅ On track — no registration gaps and no standing "
        "no-shows.\",\"⚠️ \"&N5&\" unregistered session(s) and \"&N6&\" standing "
        "no-show session(s) need action.\"))", fm.subtitle)
    stamp_live(ws, 3)
    # hidden helper numbers for the narrative
    ws.write_formula(4, 13, f"={cube_lookup('UnregSessions')}")   # N5
    ws.write_formula(5, 13, f"={cube_lookup('NoShowSessions')}")  # N6
    ws.set_column(13, 13, None, None, {"hidden": True})
    labels = [("Members", "Members"), ("Fully Trained", "FullyTrained"),
              ("Trained %", "PctTrained"), ("Fully Registered", "FullyRegistered"),
              ("Registered %", "PctRegistered"), ("Unreg. Users", "UnregUsers"),
              ("No-Show Users", "NoShowUsers"), ("Days to Go-Live", "DaysToGoLive")]
    for i, (label, col) in enumerate(labels):
        ws.write_string(4, 1 + i, label, fm.label)
        if col.startswith("Pct"):
            f = fm.kpi_pct
        else:
            f = fm.kpi_txt_link if col in KPI_LINKS else fm.kpi_txt
        ws.write_formula(5, 1 + i, kpi_formula(col), f)
    r0 = 7
    ws.write_string(r0, 1, "CLASS STATUS", fm.h2)
    heads = ["Class", "# Users", "Registered", "Completed", "Open Gaps",
             "Last Session", "Next Session", "Status"]
    for c, name in enumerate(heads):
        ws.write_string(r0 + 1, 1 + c, name, fm.th)
    if n_class:
        m2 = mask("G", "H", n_class)
        ws.write_dynamic_array_formula(
            r0 + 2, 1, r0 + 2, 1,
            f"=IFERROR(SORT(FILTER(_Long!$I$2:$P${n_class + 1},{m2},\"\"),5,-1),\"\")",
            fm.cell)
    ws.set_column(6, 7, 14, fm.dt)

    # ---- Manager Tracking --------------------------------------------------
    ws = wb.add_worksheet("Manager Tracking")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 7, 15)
    ws.write_formula(1, 1, "=SelectedLeader&\" — Manager Roll-Up\"", fm.title)
    stamp_live(ws, 2)
    heads = ["Manager", "Members", "Fully Trained", "Fully Registered",
             "Unreg. Users", "No-Show Users"]
    for c, name in enumerate(heads):
        ws.write_string(4, 1 + c, name, fm.th)
    if n_mgr:
        m3 = mask("R", "S", n_mgr)
        ws.write_dynamic_array_formula(
            5, 1, 5, 1,
            f"=IFERROR(SORT(FILTER(_Long!$T$2:$Y${n_mgr + 1},{m3},\"\"),2,-1),\"\")",
            fm.cell)

    # ---- Class Tracking ----------------------------------------------------
    ws = wb.add_worksheet("Class Tracking")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 56)
    ws.set_column(2, 8, 14)
    ws.write_formula(1, 1, "=SelectedLeader&\" — Class Roll-Up\"", fm.title)
    stamp_live(ws, 2)
    heads = ["Class", "# Users", "Registered", "Completed", "Open Gaps",
             "Last Session", "Next Session", "Status"]
    for c, name in enumerate(heads):
        ws.write_string(4, 1 + c, name, fm.th)
    if n_class:
        m2 = mask("G", "H", n_class)
        ws.write_dynamic_array_formula(
            5, 1, 5, 1,
            f"=IFERROR(SORT(FILTER(_Long!$I$2:$P${n_class + 1},{m2},\"\"),1,1),\"\")",
            fm.cell)
    ws.set_column(6, 7, 14, fm.dt)

    # ---- Classes to Schedule (static) --------------------------------------
    ws = wb.add_worksheet("Classes to Schedule")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 4)
    ws.write_string(1, 0, "Classes to Schedule", fm.title)
    ws.write_string(2, 0, "Every class someone still needs a seat in — sorted "
                          "so classes with no bookable future session come first.",
                    fm.subtitle)
    cts = d["cts"]
    n_users = d["unsum"]["TotalTMs"]
    ws.write_string(3, 0, f"Classes with demand: {len(cts)}    ·    "
                          f"User-class slots: {d['unsum']['TotalSessions']}    ·    "
                          f"Distinct users affected: {n_users}", fm.note)
    stamp_static(ws, 4, 0)
    if len(cts):
        out = cts.copy()
        out.insert(0, "#", range(1, len(out) + 1))
        out.columns = ["#", "Class (Event Name)", "Demand", "Status",
                       "Users (UID — Name)", "Their Leaders",
                       "Available Class Dates with Available Seats",
                       "Recommended registration"]
        write_df(ws, out, fm, start_row=5,
                 wrap_cols=["Users (UID — Name)", "Their Leaders",
                            "Available Class Dates with Available Seats",
                            "Recommended registration"],
                 num_cols=["#", "Demand"],
                 widths=[4, 52, 8, 22, 42, 24, 46, 34])
    else:
        ws.write_string(5, 0, "🎉 No open demand — everyone is placed.", fm.h2)

    # ---- Reg Recommendations (static) --------------------------------------
    ws = wb.add_worksheet("Reg Recommendations")
    ws.hide_gridlines(2)
    ws.write_string(1, 0, "Registration Recommendations", fm.title)
    ws.write_string(2, 0, "Top 3 future sessions with open seats for every "
                          "person who still needs a class. Filter any column.",
                    fm.subtitle)
    stamp_static(ws, 3, 0)
    recs = d["recs"]
    if len(recs):
        out = recs.copy()
        out.insert(0, "#", range(1, len(out) + 1))
        out.columns = ["#", "Universal Id", "Team Member", "Leader", "Reason",
                       "Class Needed",
                       "Opt 1 Date", "Opt 1 Location", "Opt 1 Seats",
                       "Opt 2 Date", "Opt 2 Location", "Opt 2 Seats",
                       "Opt 3 Date", "Opt 3 Location", "Opt 3 Seats",
                       "Assignment"]
        write_df(ws, out, fm, start_row=4,
                 num_cols=["#", "Opt 1 Seats", "Opt 2 Seats", "Opt 3 Seats"],
                 widths=[4, 14, 24, 20, 20, 46, 11, 30, 9, 11, 30, 9, 11, 30,
                         9, 20])
    else:
        ws.write_string(4, 0, "🎉 No one needs placement right now.", fm.h2)

    # ---- Registration Schedule (static) ------------------------------------
    ws = wb.add_worksheet("Registration Schedule")
    ws.hide_gridlines(2)
    ws.write_string(1, 0, "Registration Schedule", fm.title)
    ws.write_string(2, 0, "Every scheduled session with seats and possible "
                          "placements for people still needing the class.",
                    fm.subtitle)
    stamp_static(ws, 3, 0)
    rs = d["regsched"][["SessionDate", "Location", "Room", "Session_Type",
                        "Primary_Instructor", "Secondary_Instructor",
                        "EventName", "StartTime", "EndTime", "TotalSeats",
                        "SeatsTaken", "AvailableSeats", "Option",
                        "PossibleUsers", "TheirLeaders"]].copy()
    rs.columns = ["Date", "Training Location", "Room", "Session Type",
                  "Primary Instructor", "Secondary Instructor", "Event Name",
                  "Start Time", "End Time", "Total Seats", "Seats Taken",
                  "Available Seats", "Option",
                  "Unregistered Users Possible Placement", "Their Leaders"]
    write_df(ws, rs, fm, start_row=5,
             num_cols=["Total Seats", "Seats Taken", "Available Seats"],
             wrap_cols=["Unregistered Users Possible Placement", "Their Leaders"],
             widths=[11, 26, 24, 13, 26, 26, 46, 10, 10, 10, 10, 10, 10, 40, 22])

    # ---- Unregistered Summary (follows the Control filter) -----------------
    ws = wb.add_worksheet("Unregistered Summary")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 6, 22)
    ws.write_string(1, 1, f"WAVE {WAVE_NUM} — UNREGISTERED & NO-SHOW SESSIONS",
                    fm.title)
    ws.write_string(2, 1, f"Report as of {built_at:%m/%d/%Y %I:%M %p}", fm.subtitle)
    stamp_live(ws, 3)
    vals = [("Distinct Team Members Needing Placement", "UnregUsers"),
            ("Total Sessions To Place", "UnregSessions"),
            ("Unregistered Sessions", "PlaceUnregSessions"),
            ("No-Show Sessions To Re-Register", "PlaceNoShowSessions"),
            ("Unregistered Team Members", "PlaceUnregTMs"),
            ("No-Show Team Members", "PlaceNoShowTMs")]
    for i, (label, col) in enumerate(vals):
        f = fm.kpi_link if col in KPI_LINKS else fm.kpi
        ws.write_formula(4, 1 + i, kpi_formula(col), f)
        ws.write_string(5, 1 + i, label, fm.label)
    ws.write_string(7, 1, "BREAKDOWN", fm.h2)
    for c, name in enumerate(["", "Team Members", "% of Total", "Sessions",
                              "% of Total"]):
        ws.write_string(8, 1 + c, name, fm.th)
    for i, (label, tm_col, s_col) in enumerate([
            ("Unregistered", "PlaceUnregTMs", "PlaceUnregSessions"),
            ("No-Show (re-register)", "PlaceNoShowTMs", "PlaceNoShowSessions")],
            start=9):
        ws.write_string(i, 1, label, fm.cell)
        ws.write_formula(i, 2, f"={cube_lookup(tm_col)}", fm.num)
        ws.write_formula(i, 3, f"=IFERROR(C{i + 1}/$C$12,0)", fm.pct)
        ws.write_formula(i, 4, f"={cube_lookup(s_col)}", fm.num)
        ws.write_formula(i, 5, f"=IFERROR(E{i + 1}/$E$12,0)", fm.pct)
    ws.write_string(11, 1, "Total", fm.th)
    ws.write_formula(11, 2, f"={cube_lookup('UnregUsers')}", fm.num)
    ws.write_formula(11, 4, f"={cube_lookup('UnregSessions')}", fm.num)

    # ---- Registration Audit (follows the leader filter) --------------------
    ws = wb.add_worksheet("Registration Audit")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 2, 24)
    ws.set_column(3, 3, 26)
    ws.set_column(4, 4, 30)
    ws.set_column(5, 5, 44)
    ws.set_column(6, 6, 44)
    ws.set_column(7, 7, 12, fm.dt)
    ws.set_column(8, 8, 44)
    ws.write_string(1, 1, "Registration Audit", fm.title)
    ws.write_string(2, 1, "Hygiene checks over current registrations — a zero "
                          "means the check ran clean.", fm.subtitle)
    ws.write_formula(
        3, 1,
        '="🔎 Leader filter applies — showing: "&SelectedLeader'
        '&"  ·  scope toggle not used on this sheet  ·  ' + asof + '"', fm.note)
    audit_checks = [
        ("Out-of-scope registration",
         "Departed / out-of-scope people still holding future seats — free them."),
        ("Schedule conflict",
         "Registered in two overlapping sessions on the same day."),
        ("Duplicate registration",
         "More than one active booking for the same class (Epic future "
         "sessions or Cornerstone Registered transcripts, W2-era titles "
         "normalized)."),
        ("Curriculum not in role mapping",
         "Assigned Epic curriculum is outside the MVP training tracks for "
         "the person's job role(s)."),
        ("Registered for class not required",
         "Future registration for a class in NEITHER the person's Epic "
         "requirements NOR their MVP role tracks — verify with the leader; "
         "the Detail column carries the evidence. Withdrawn-then-"
         "re-registered rows are handled (only current registrations "
         "count)."),
    ]
    audit_ok = [
        ("OK - correct class, Epic gap",
         "NOT an issue. The class matches the person's role tracks — the "
         "registration is right; Epic just can't see them (on leave / "
         "Epic-ineligible, or Epic lag). Listed only to explain Epic-vs-"
         "Cornerstone count differences."),
    ]
    aA, aZ = 2, max(n_audit + 1, 2)

    def audit_count_row(row, name, desc):
        ws.write_string(row, 1, name, fm.cell)
        ws.write_formula(
            row, 2,
            f'=IF(SelectedLeader="All",'
            f'COUNTIF(_Long!$AJ${aA}:$AJ${aZ},B{row + 1}),'
            f'COUNTIFS(_Long!$AJ${aA}:$AJ${aZ},B{row + 1},'
            f'_Long!$AG${aA}:$AG${aZ},SelectedLeader))', fm.num)
        ws.merge_range(row, 3, row, 6, desc, fm.cell)

    ws.write_string(5, 1, "ISSUES TO REVIEW", fm.h2)
    ws.write_string(6, 1, "Check", fm.th)
    ws.write_string(6, 2, "Flagged", fm.th)
    ws.merge_range(6, 3, 6, 6, "What it means", fm.th)
    for i, (name, desc) in enumerate(audit_checks, start=7):
        audit_count_row(i, name, desc)
    r_ok = 7 + len(audit_checks) + 1
    ws.write_string(r_ok, 1, "EXPLAINED — NO ACTION NEEDED", fm.h2)
    ws.write_string(r_ok + 1, 1, "Category", fm.th)
    ws.write_string(r_ok + 1, 2, "Sessions", fm.th)
    ws.merge_range(r_ok + 1, 3, r_ok + 1, 6, "Why it is not an issue", fm.th)
    for i, (name, desc) in enumerate(audit_ok, start=r_ok + 2):
        audit_count_row(i, name, desc)
    r0 = r_ok + 2 + len(audit_ok) + 1
    ws.write_string(r0, 1, "FLAGGED DETAIL", fm.h2)
    audit_heads = ["Universal Id", "Full Name", "Leader", "Check",
                   "Curriculum / Track", "Class", "Date", "Detail"]
    for c, h in enumerate(audit_heads):
        ws.write_string(r0 + 1, 1 + c, h, fm.th)
    if n_audit:
        am = (f'(((_Long!$AG${aA}:$AG${aZ}=SelectedLeader)'
              f'+(SelectedLeader="All"))>0)')
        ws.write_dynamic_array_formula(
            r0 + 2, 1, r0 + 2, 1,
            f'=IFERROR(SORT(FILTER(CHOOSECOLS(_Long!$AG${aA}:$AN${aZ},'
            f'2,3,1,4,5,6,7,8),{am},"✓ Nothing flagged for this selection"),'
            f'4,1),"")', fm.cell)
    else:
        ws.write_string(r0 + 2, 1, "✓ All checks ran clean — nothing flagged.",
                        fm.cell)

    # ---- Training Status (static, all leaders) -----------------------------
    ws = wb.add_worksheet("Training Status")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 12, 11)
    ws.write_string(1, 1, "Training Status — Not Started / In Progress / "
                          "Fully Trained", fm.title)
    ws.write_string(
        2, 1,
        "All leaders, standard scope. In Progress = attended at least one "
        "required class; Fully Trained = attended every class needed (Epic "
        "flag). Final Scheduled = the learner's last booked session — their "
        "projected completion date; blank means nothing scheduled. Use it to "
        "spot who finishes too close to go-live and needs to shift earlier.",
        fm.note)
    stamp_static(ws, 3)
    ws.write_string(4, 1, "BY LEADER", fm.h2)
    heads = ["Leader", "Members"]
    for st in ("Not Started", "In Progress", "Fully Trained"):
        heads += [st, "%"]
    heads += ["No Epic Curriculum"]
    for c, h in enumerate(heads):
        ws.write_string(5, 1 + c, h, fm.th)
    tl = d["tstat_leader"]
    for i, (_, row) in enumerate(tl.iterrows(), start=6):
        vals = [row["Leader"], row["Members"],
                row["Not Started"], row["Not Started %"],
                row["In Progress"], row["In Progress %"],
                row["Fully Trained"], row["Fully Trained %"],
                row["No Epic Curriculum"]]
        for c, v in enumerate(vals):
            fmt = fm.cell if c == 0 else (fm.pct if heads[c] == "%" else fm.num)
            if row["Leader"] == "TOTAL" and c == 0:
                fmt = fm.th
            _w(ws, i, 1 + c, v, fmt)
    r0 = 6 + len(tl) + 2
    ws.write_string(r0, 1, "LEARNER DETAIL (filter any column)", fm.h2)
    det_cols = [("Universal Id", "UniversalID"), ("Full Name", "FullName"),
                ("Leader", "Leader"), ("Status", "TrainingStatus"),
                ("Classes Done", "ClassesCompleted"),
                ("Classes Needed", "ClassesRequired"),
                ("Last Attended", "LastAttendedDate"),
                ("Final Scheduled", "FinalScheduledDate"),
                ("Unreg Classes", "UnregisteredClasses"),
                ("Attention", "Attention")]
    for c, (h, _) in enumerate(det_cols):
        ws.write_string(r0 + 1, 1 + c, h, fm.th)
    tsd = d["tstatus"]
    for i, (_, row) in enumerate(tsd.iterrows(), start=r0 + 2):
        for c, (h, col) in enumerate(det_cols):
            fmt = fm.dt if "Attended" in h or "Scheduled" in h else \
                (fm.num if h.startswith(("Classes", "Unreg")) else fm.cell)
            _w(ws, i, 1 + c, row[col], fmt, fm.dt)
    ws.autofilter(r0 + 1, 1, r0 + 1 + len(tsd), len(det_cols))
    ws.set_column(2, 2, 24)
    ws.set_column(3, 3, 20)

    # ---- Completion Status (own leader filter) -----------------------------
    ws = wb.add_worksheet("Completion Status")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 15)          # Universal Id
    ws.set_column(2, 2, 26)          # Full Name
    ws.set_column(3, 4, 13)          # Yes/No flags
    ws.set_column(5, 5, 46)          # Last Registered Class
    ws.set_column(6, 7, 16, fm.dt)   # Expected / Completion dates
    ws.set_column(8, 11, 13)         # counts
    ws.set_column(12, 12, 22)        # Leader
    ws.write_string(1, 1, "Completion Status", fm.title)
    ws.write_string(
        2, 1,
        "Per-person registration & completion status (standard scope, "
        "distinct classes). Expected Training Completion = the LAST session "
        "date the person is booked into; Completion Date = last attended "
        "session once Epic marks them fully trained. # Pending Attendance = "
        "classes registered for whose session date is still ahead.", fm.note)
    wb.define_name("CompLeader", "='Completion Status'!$B$7")
    ws.write_formula(
        3, 1,
        '="🎚 This sheet has its OWN leader filter below (the Control sheet '
        'does not apply) — showing: "&CompLeader&"  ·  ' + asof + '"',
        fm.note)
    ws.write_string(5, 1, "Leader", fm.h2)
    ws.write_string(6, 1, "All", fm.ctrl)
    ws.data_validation(6, 1, 6, 1, {
        "validate": "list",
        "source": f"=_Lists!$A$2:$A${len(leader_opts) + 1}"})
    nc1 = n_comp + 1
    cmask = (f"((_Comp!$A$2:$A${nc1}=CompLeader)+(CompLeader=\"All\"))")
    kpi_hdrs = ["Members", "Fully Registered", "% Fully Registered",
                "Fully Trained", "% Fully Trained", "Not Fully Registered"]
    kpi_fmls = [
        f"=SUMPRODUCT(({cmask}>0)*1)",
        f"=SUMPRODUCT(({cmask}>0)*(_Comp!$D$2:$D${nc1}=\"Yes\"))",
        "=IFERROR(C10/B10,0)",
        f"=SUMPRODUCT(({cmask}>0)*(_Comp!$E$2:$E${nc1}=\"Yes\"))",
        "=IFERROR(E10/B10,0)",
        "=B10-C10",
    ]
    for c, (h, f_) in enumerate(zip(kpi_hdrs, kpi_fmls)):
        ws.write_string(8, 1 + c, h, fm.th)
        ws.write_formula(9, 1 + c, f_, fm.pct if "%" in h else fm.num)
    ws.set_row(8, 26)
    hdrs = ["Universal Id", "Full Name", "Fully Registered", "Fully Trained",
            "Last Registered Class", "Expected Completion", "Completion Date",
            "# Registered", "# Completed", "# Pending Registration",
            "# Pending Attendance", "Leader"]
    for c, h in enumerate(hdrs):
        ws.write_string(11, 1 + c, h, fm.th)
    ws.write_dynamic_array_formula(
        12, 1, 12, 1,
        f"=IFERROR(SORT(FILTER(_Comp!$B$2:$M${nc1},{cmask}>0,\"\"),2,1),\"\")",
        fm.cell)
    ws.freeze_panes(12, 0)

    # ---- No Shows (static) -------------------------------------------------
    ws = wb.add_worksheet("No Shows")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 13)
    ws.set_column(2, 2, 22)
    ws.set_column(3, 3, 20)
    ws.set_column(4, 4, 40)
    ws.set_column(5, 12, 13)
    ws.write_string(1, 1, "No Shows", fm.title)
    ws.write_string(
        2, 1,
        "One row per person × class with a no-show on record (Cornerstone "
        "transcripts). Resolution 'No Show standing' = the newest transcript "
        "row is still No Show — nothing newer booked; those with Registration "
        "Action 'Needs to be Re-Registered' also appear in Reg "
        "Recommendations. Resolved rows are kept for history.", fm.note)
    stamp_static(ws, 3)
    ns_cols = [("Universal Id", "UniversalID"), ("Full Name", "FullName"),
               ("Leader", "Leader"), ("Class", "ClassTitle"),
               ("No-Show Count", "NoShowCount"),
               ("No-Show Session", "LastNoShowSessionDate"),
               ("Current Status", "CurrentStatus"),
               ("Current Session", "CurrentSessionDate"),
               ("Registered per Epic", "RegisteredPerEpic"),
               ("Resolution", "Resolution"),
               ("Registration Action", "RegistrationAction")]
    for c, (h, _) in enumerate(ns_cols):
        ws.write_string(4, 1 + c, h, fm.th)
    nsd = d["noshow"]
    for i, (_, row) in enumerate(nsd.iterrows(), start=5):
        for c, (h, col) in enumerate(ns_cols):
            fmt = fm.dt if "Session" in h and "Status" not in h else \
                (fm.num if h == "No-Show Count" else fm.cell)
            _w(ws, i, 1 + c, row[col], fmt, fm.dt)
    ws.autofilter(4, 1, 4 + len(nsd), len(ns_cols))
    ws.freeze_panes(5, 0)

    # ---- Go and Finds (static) ---------------------------------------------
    ws = wb.add_worksheet("Go and Finds")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 5, 14)
    ws.write_string(1, 1, "Go and Finds — Module Assignments", fm.title)
    ws.write_string(
        2, 1,
        "Assigned Go-and-Find (self-directed) modules per learner, from the "
        "W3 GNF mapping file. Assignment mapping only — completion tracking "
        "will be added when that data becomes available. New mapping files "
        "drop into data\\raw\\wave\\gnf_mapping (newest file loads on "
        "refresh, any cadence).", fm.note)
    stamp_static(ws, 3)
    gs = d["gnf_status"]
    gd_all = d["gnf"]
    n_pop = len(gs)
    n_map = int((gs["GnfStatus"] == "Mapped").sum())
    kpis = [("Team Members", n_pop, "num"),
            ("Mapped", n_map, "num"),
            ("Mapped %", n_map / n_pop if n_pop else 0, "pct"),
            ("No Modules Assigned",
             int((gs["GnfStatus"] == "No modules assigned").sum()), "num"),
            ("Pending Mapping",
             int((gs["GnfStatus"] == "Pending mapping").sum()), "num"),
            ("Module Assignments", len(gd_all), "num"),
            ("Distinct Modules", int(gd_all["GoAndFind"].nunique()), "num")]
    for i, (label, v, kind) in enumerate(kpis):
        ws.write_string(5, 1 + i, label, fm.label)
        ws.write_number(6, 1 + i, v, fm.kpi_pct if kind == "pct" else fm.kpi)
    ws.write_string(
        7, 1,
        "Mapped = at least one GNF module in the W3 mapping file. No Modules "
        "Assigned = in the mapping file with zero modules (can be correct "
        "for the role). Pending Mapping = not in the mapping file yet. "
        "In-progress/completed metrics arrive when completion data becomes "
        "available.", fm.note)
    ws.write_string(9, 1, "BY LEADER", fm.h2)
    gl_heads = ["Leader", "Mapped", "No Modules", "Pending Mapping",
                "Module Assignments", "Distinct Modules"]
    for c, h in enumerate(gl_heads):
        ws.write_string(10, 1 + c, h, fm.th)
    gl = d["gnf_leader"]
    for i, (_, row) in enumerate(gl.iterrows(), start=11):
        for c, col in enumerate(["Leader", "Mapped", "NoModules",
                                 "PendingMapping", "ModuleAssignments",
                                 "DistinctModules"]):
            fmt = (fm.th if col == "Leader" and row["Leader"] == "TOTAL"
                   else (fm.cell if col == "Leader" else fm.num))
            _w(ws, i, 1 + c, row[col], fmt)
    r0 = 11 + len(gl) + 2
    ws.write_string(r0, 1, "ASSIGNMENT DETAIL (filter any column)", fm.h2)
    g_cols = [("Universal Id", "UniversalID"), ("Full Name", "FullName"),
              ("Leader", "Leader"), ("User Role", "UserRole"),
              ("Business Unit", "BusinessUnit"),
              ("Go and Find Module", "GoAndFind"),
              ("User Total GNFs", "UserGnfCount")]
    for c, (h, _) in enumerate(g_cols):
        ws.write_string(r0 + 1, 1 + c, h, fm.th)
    gd = d["gnf"]
    for i, (_, row) in enumerate(gd.iterrows(), start=r0 + 2):
        for c, (h, col) in enumerate(g_cols):
            _w(ws, i, 1 + c, row[col],
               fm.num if h == "User Total GNFs" else fm.cell)
    ws.autofilter(r0 + 1, 1, r0 + 1 + len(gd), len(g_cols))
    ws.set_column(2, 2, 24)
    ws.set_column(4, 4, 26)
    ws.set_column(5, 5, 26)
    ws.set_column(6, 6, 44)

    # ---- Data (static detail) ----------------------------------------------
    ws = wb.add_worksheet("Data")
    ws.write_string(0, 0, "Data — raw person × class detail (Epic Curriculum "
                          "Status)", fm.h2)
    stamp_static(ws, 1, 0)
    write_df(ws, d["data"], fm, start_row=3,
             num_cols=["Sequence", "Duration Hours", "Is Excluded"],
             widths=[14, 26, 20, 16, 24, 12, 13, 12, 24, 26, 26, 44, 16, 10,
                     44, 14, 12, 14, 10, 16, 22, 18, 18, 10])

    wb.close()


LOG_SPEC = [
    ("Members", "Members", "num"),
    ("Fully Registered", "FullyRegistered", "num"),
    ("Δ Reg", "ΔReg", "delta"),
    ("Reg %", "RegPct", "pct"),
    ("Fully Trained", "FullyTrained", "num"),
    ("Δ Trained", "ΔTrained", "delta"),
    ("Trained %", "TrainedPct", "pct"),
    ("Unreg Users", "UnregisteredUsers", "num"),
    ("Δ Unreg", "ΔUnreg", "delta"),
    ("Unreg Sessions", "UnregisteredSessions", "num"),
    ("No-Show Standing", "NoShowUsersStanding", "num"),
    ("Δ No-Show", "ΔNoShow", "delta"),
    ("No-Show Sessions", "NoShowSessionsStanding", "num"),
    ("No-Show To-Date", "NoShowSessionsToDate", "num"),
]


def _log_table(ws, fm, df, spec, r0, newest_first=True, autofilter=False,
               freeze=True):
    """Render a log table at row r0 / col B. spec = [(label, col, kind)]."""
    kinds = {"num": fm.num, "delta": fm.delta, "pct": fm.pct,
             "date": fm.dt, "text": fm.cell}
    for c, (label, _, _k) in enumerate(spec):
        ws.write_string(r0, 1 + c, label, fm.th)
    body = df.iloc[::-1] if newest_first else df
    for i, (_, row) in enumerate(body.iterrows(), start=r0 + 1):
        for c, (_, col, k) in enumerate(spec):
            _w(ws, i, 1 + c, row[col], kinds[k], fm.dt)
    if autofilter:
        ws.autofilter(r0, 1, r0 + len(df), len(spec))
    if freeze:
        ws.freeze_panes(r0 + 1, 0)
    return r0 + 1 + len(df)


def build_log_workbook(d: dict, path: Path) -> None:
    """The dedicated daily/weekly tracking workbook (RCM Training Daily Log)."""
    wb = xlsxwriter.Workbook(str(path), {"nan_inf_to_errors": True,
                                         "default_date_format": "m/d/yyyy"})
    fm = Fmt(wb)
    built_at = datetime.now()
    tot, ldr = d["log_totals"], d["log_leader"]
    wkt, wkl = d["log_week_totals"], d["log_week_leader"]

    # hidden ascending chart feed
    ws_cd = wb.add_worksheet("_ChartData")
    for c, col in enumerate(["SnapshotDate", "FullyRegistered", "FullyTrained",
                             "UnregisteredUsers", "RegPct"]):
        ws_cd.write_string(0, c, col)
        for i, v in enumerate(tot[col], start=1):
            _w(ws_cd, i, c, v, fm.num, fm.dt)
    ws_cd.hide()
    n = len(tot)

    # ---- Daily Totals ------------------------------------------------------
    ws = wb.add_worksheet("Daily Totals")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 11)
    ws.set_column(2, 15, 11)
    ws.write_string(1, 1, f"{TITLE} — Daily Log", fm.title)
    ws.write_string(
        2, 1,
        f"Newest first. One row per snapshot day (weekdays). Registered/Trained "
        f"tracked since 07/06/2026; unregistered & no-show tracking began "
        f"07/15/2026 (blank before). Mid-July registration dip = the Cornerstone "
        f"withdrawal incident. Today's row is live until the next daily snapshot. "
        f"Rebuilt {built_at:%m/%d/%Y %I:%M %p}.", fm.note)
    spec = [("Date", "SnapshotDate", "date")] + LOG_SPEC
    ws.write_string(3, 1, "DAILY TOTALS", fm.h2)
    _log_table(ws, fm, tot, spec, 4, newest_first=True)
    if n >= 2:
        for title, series, y_opts, pos in (
            ("Daily totals", (("Fully Registered", "B", ACCENT),
                              ("Fully Trained", "C", "#4E9A6F"),
                              ("Unregistered Users", "D", ACCENT2)),
             {}, (4, 17)),
            ("Fully Registered %", (("Reg %", "E", ACCENT),),
             {"num_format": "0%", "min": 0, "max": 1}, (18, 17)),
        ):
            ch = wb.add_chart({"type": "line"})
            for label, colL, color in series:
                ch.add_series({
                    "name": label,
                    "categories": f"=_ChartData!$A$2:$A${n + 1}",
                    "values": f"=_ChartData!${colL}$2:${colL}${n + 1}",
                    "line": {"color": color, "width": 2},
                })
            ch.set_title({"name": title,
                          "name_font": {"name": "Segoe UI", "size": 11,
                                        "color": INK, "bold": False}})
            ch.set_legend({"position": "bottom",
                           "font": {"name": "Segoe UI", "size": 9}})
            ch.set_x_axis({"num_font": {"name": "Segoe UI", "size": 9},
                           "num_format": "m/d", "line": {"color": LINE}})
            ch.set_y_axis({"num_font": {"name": "Segoe UI", "size": 9},
                           "major_gridlines": {"visible": True,
                                               "line": {"color": LINE}},
                           **y_opts})
            ch.set_chartarea({"border": {"none": True}})
            ch.set_size({"width": 540, "height": 250})
            ws.insert_chart(*pos, ch)

    # ---- Daily by Leader ---------------------------------------------------
    ws = wb.add_worksheet("Daily by Leader")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 11)
    ws.set_column(2, 2, 20)
    ws.set_column(3, 16, 11)
    ws.write_string(1, 1, "Daily Log — By Leader", fm.title)
    ws.write_string(2, 1, f"Newest first. Filter the Leader column for a "
                          f"single-leader trend. Δ columns compare to that "
                          f"leader's previous snapshot day. "
                          f"Rebuilt {built_at:%m/%d/%Y %I:%M %p}.", fm.note)
    spec = [("Date", "SnapshotDate", "date"), ("Leader", "Leader", "text")] + LOG_SPEC
    _log_table(ws, fm, ldr, spec, 4, newest_first=True, autofilter=True)

    # ---- Weekly Summary ----------------------------------------------------
    ws = wb.add_worksheet("Weekly Summary")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 11)
    ws.set_column(2, 17, 11)
    ws.write_string(1, 1, "Weekly Summary", fm.title)
    ws.write_string(
        2, 1,
        f"End-of-week values (last snapshot of each week, Monday-start). "
        f"Δ columns are week-over-week. Days = snapshot days captured that "
        f"week. Rebuilt {built_at:%m/%d/%Y %I:%M %p}.", fm.note)
    wspec = [("Week Of", "Week", "date"), ("Days", "DaysCaptured", "num")] + LOG_SPEC
    ws.write_string(3, 1, "WEEKLY TOTALS", fm.h2)
    end = _log_table(ws, fm, wkt, wspec, 4, newest_first=True, freeze=False)
    r1 = end + 2
    ws.write_string(r1, 1, "WEEKLY BY LEADER (filter any column)", fm.h2)
    wlspec = [("Week Of", "Week", "date"), ("Leader", "Leader", "text"),
              ("Days", "DaysCaptured", "num")] + LOG_SPEC
    _log_table(ws, fm, wkl, wlspec, r1 + 1, newest_first=True, autofilter=True,
               freeze=False)   # two stacked tables — freezing mid-sheet would
    ws.set_column(2, 2, 18)    # hide the totals table above

    # ---- Unregistered Trend (leader x day matrix, count + % adjustment) ----
    ws = wb.add_worksheet("Unregistered Trend")
    ws.hide_gridlines(2)
    base = {"font_name": "Segoe UI", "font_color": INK, "font_size": 10}
    hdr_rot = wb.add_format({**base, "bold": True, "rotation": 45,
                             "bottom": 1, "border_color": LINE})
    pct_dn = wb.add_format({**base, "num_format": "0.0%",
                            "bg_color": "#C6EFCE", "font_color": "#006100"})
    pct_up = wb.add_format({**base, "num_format": "0.0%",
                            "bg_color": "#FFC7CE", "font_color": "#9C0006"})
    ldr_f = wb.add_format({**base, "bold": True})
    tot_f = wb.add_format({**base, "bold": True, "top": 1,
                           "border_color": GRAY})
    piv = (ldr.pivot_table(index="Leader", columns="SnapshotDate",
                           values="UnregisteredUsers", aggfunc="first")
           .dropna(axis=1, how="all"))
    piv = piv[sorted(piv.columns)]
    dates = list(piv.columns)
    ws.set_row(0, 62)
    ws.set_column(0, 0, 22)
    ws.set_column(1, 1 + 2 * len(dates) + 2, 9.5)
    ws.write_string(0, 0, "Leaders", fm.th)
    for j, dt_ in enumerate(dates):
        ws.write_string(0, 1 + 2 * j, pd.Timestamp(dt_).strftime("%b %d, %Y"),
                        hdr_rot)
        ws.write_string(0, 2 + 2 * j, "% Adjustment", hdr_rot)

    def matrix_row(r, label, series, label_fmt):
        ws.write_string(r, 0, label, label_fmt)
        prev = None
        for j, dt_ in enumerate(dates):
            v = series.get(dt_)
            v = None if v is None or pd.isna(v) else int(v)
            if v is not None:
                ws.write_number(r, 1 + 2 * j, v, fm.num)
                if prev is not None:
                    if prev > 0:
                        pct = (v - prev) / prev
                        ws.write_number(r, 2 + 2 * j, pct,
                                        pct_dn if pct < 0 else pct_up)
                    elif v == 0:
                        ws.write_number(r, 2 + 2 * j, 0, pct_up)
                    # prev == 0 and v > 0: no % possible — leave blank
                prev = v
    for i, name in enumerate(d["leaders"], start=1):
        series = piv.loc[name] if name in piv.index else {}
        matrix_row(i, name, series, ldr_f)
    matrix_row(1 + len(d["leaders"]), "TOTAL", piv.sum(), tot_f)
    ws.freeze_panes(1, 1)
    r_note = 3 + len(d["leaders"])
    ws.write_string(
        r_note, 0,
        f"Unregistered USERS per leader per snapshot day (weekdays). "
        f"% Adjustment = change vs the previous captured day — green = fewer "
        f"unregistered. Tracking began 07/15/2026; earlier days were not "
        f"captured. Today's column is live until the next daily snapshot. "
        f"Rebuilt {built_at:%m/%d/%Y %I:%M %p}.", fm.note)

    wb.close()


# ---------------------------------------------------------------------------
# QA + main
# ---------------------------------------------------------------------------
def qa_crosscheck(d: dict) -> None:
    sc = d["scorecard"]
    tot = sc[sc["IsTotal"] == 1].iloc[0]
    cube = d["cube"]
    inc = cube[(cube["Variant"] == "Include") & (cube["Leader"] == "All")].iloc[0]
    print("\nQA cross-check vs report.w3_scorecard (TOTAL row):")
    checks = [("Members", "W3Users"), ("FullyRegistered", "FullyRegistered"),
              ("FullyTrained", "FullyTrained"),
              ("NoEpicCurriculum", "NoEpicCurriculum")]
    ok = True
    for cube_col, sc_col in checks:
        a, b = int(inc[cube_col]), int(tot[sc_col])
        flag = "OK " if a == b else "DIFF"
        if a != b:
            ok = False
        print(f"  [{flag}] {cube_col:18} tracker={a:5}  scorecard={b:5}")
    print(f"  [info] UnregSessions tracker(Include)={int(inc['UnregSessions'])} "
          f"scorecard={int(tot['UnregisteredSessions'])} "
          "(tracker adds no-show re-registrations)")
    if d["unmatched_classes"]:
        print(f"  [warn] {len(d['unmatched_classes'])} in-demand class(es) not on "
              f"the schedule at all: "
              + "; ".join(d["unmatched_classes"][:5]))
    if not ok:
        print("  !! numbers differ from the scorecard — investigate before sharing.")


def excel_check(path: Path) -> bool:
    """Open in Excel (invisible), force a recalc, scan for error values.
    Returns True when clean. Run against the STAGED copy so Excel never
    touches the OneDrive-synced files."""
    import xlwings as xw
    errs = []
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        book = app.books.open(str(path))
        app.calculate()
        for sht in book.sheets:
            used = sht.used_range
            if used.count == 1 and used.value is None:
                continue
            vals = used.options(ndim=2).value
            for ri, rowv in enumerate(vals):
                for ci, v in enumerate(rowv):
                    if isinstance(v, int) and v in range(-2146826288, -2146826200):
                        errs.append((sht.name, ri + used.row, ci + used.column, v))
        book.close()
    finally:
        app.quit()
    if errs:
        print(f"  !! Excel recalc found {len(errs)} error cell(s):")
        for e in errs[:20]:
            print("     ", e)
        return False
    print("  Excel recalc: no error values in any sheet.")
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("--no-sql", action="store_true",
                    help="skip the change-aware SQL load before building")
    ap.add_argument("--excel-check", action="store_true",
                    help="after building, open in Excel and scan for errors")
    args = ap.parse_args()

    if not args.no_sql:
        print("SQL load (change-aware)…")
        rc = subprocess.call([str(VENV_PY), str(ROOT / "sql" / "auto_refresh.py")])
        if rc != 0:
            print(f"  auto_refresh exited {rc} — building from what's loaded.")

    # refuse to clobber a workbook that's open in Excel
    for p in (OUT_PATH, LOG_PATH):
        if p.exists():
            try:
                with open(p, "r+b"):
                    pass
            except PermissionError:
                print(f"ABORT: {p.name} is open in Excel — close it and rerun.")
                return 1

    print("Querying SQL…")
    engine = create_engine(CONN)
    frames = load_frames(engine)
    engine.dispose()
    if not len(frames["det"]):
        print(f"ABORT: report.tracker_detail has no {WAVE} rows — check the "
              "Epic status load before building.")
        return 1
    stale = pd.to_datetime(frames["fresh"]["LoadedAt"], errors="coerce").max()
    if pd.notna(stale) and (datetime.now() - stale) > timedelta(days=2):
        print(f"  [warn] newest SQL load is {stale:%m/%d %H:%M} — data may be stale.")

    print("Computing…")
    d = compute(frames)

    print("Writing workbooks (local staging)…")
    STAGING.mkdir(exist_ok=True)
    stage_out = STAGING / OUT_PATH.name
    stage_log = STAGING / LOG_PATH.name
    build_workbook(d, stage_out)
    print(f"Built {stage_out.name}  ({stage_out.stat().st_size / 1e6:.1f} MB)")
    build_log_workbook(d, stage_log)
    print(f"Built {stage_log.name}  ({stage_log.stat().st_size / 1e6:.1f} MB)")

    qa_crosscheck(d)
    if args.excel_check:
        for p in (stage_out, stage_log):
            print(f"Excel recalc check (staged) — {p.name}…")
            if not excel_check(p):
                print("ABORT: recalc errors in the staged build — nothing "
                      "was published to OneDrive.")
                return 1

    # publish: ONE atomic copy per file onto the OneDrive path — the synced
    # folder never sees a partially-written workbook or an Excel session.
    import shutil
    for src, dst in ((stage_out, OUT_PATH), (stage_log, LOG_PATH)):
        tmp = dst.with_suffix(".publishing.xlsx")
        try:
            shutil.copy2(src, tmp)
            os.replace(tmp, dst)
        except PermissionError:
            print(f"ABORT: {dst.name} is open in Excel — close it and rerun. "
                  f"The verified build is kept at {src}.")
            return 1
        print(f"Published {dst}  ({dst.stat().st_size / 1e6:.1f} MB)")

    inc = d["cube"]
    row = inc[(inc["Variant"] == "Exclude") & (inc["Leader"] == "All")].iloc[0]
    print(f"\n{TITLE}: {int(row['Members'])} members · "
          f"{int(row['FullyRegistered'])} fully registered · "
          f"{int(row['FullyTrained'])} fully trained · "
          f"{int(row['UnregUsers'])} unregistered · "
          f"{int(row['NoShowUsers'])} standing no-shows · "
          f"{len(d['cts'])} classes to schedule")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
