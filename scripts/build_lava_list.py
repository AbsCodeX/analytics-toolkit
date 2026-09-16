# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_lava_list.py  —  build the LAVA Wave census (the working LAVA list).

Renamed 2026-09-08 (her ask): no longer "Preliminary" — this is the working
version. Output = "LAVA <wave> Census current as of <build date>.xlsx"; the
date is the report completion (build) date. Older "Preliminary User List"
files are swept into LAVA Archive like any other previous build.

Population = every wave-file person tagged for training in the target wave,
plus everyone in the SVC (Simple Visit Coding) provisioning export, plus every
Soft Live participant on the wave file whatever their wave (2026-09-03).

    python scripts\\build_lava_list.py                       # Wave 3, latest SVC export
    python scripts\\build_lava_list.py --wave "Wave 4"
    python scripts\\build_lava_list.py --svc-export "path\\to\\export.xlsx"
    python scripts\\build_lava_list.py --out "LAVA W3 for Lastname03.xlsx"
    python scripts\\build_lava_list.py --show-config        # print settings, build nothing

Re-running is safe. If the output file already exists it is backed up first, and
any Yes/No a person typed into a REVIEW column (Training Complete, FEC, Soft
Live, Workqueue Owner) is carried forward onto the refreshed rows — the data
refreshes, hand-entered confirmations survive. Use --no-preserve to start clean.

Sheets produced: Executive Summary · LAVA List · Guest House Job Roles ·
SVC Export (raw).

Customising it later:
  - new column        -> add one (Header, field, width) tuple to COLUMNS
  - who counts as SVC -> SVC_ROLE_PATTERN
  - who counts as a leader -> LEADER_TITLE_PATTERN
  - guest house        -> comes from MVP IsGuestHouse, no list to maintain
  - population rules   -> --leader-list-only / --include-role-svc, see CONFIG
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "sql"))
from refresh import CONN, op                                   # noqa: E402
from email_lookup import email_map                            # noqa: E402
from sqlalchemy import create_engine                           # noqa: E402
import pandas as pd                                            # noqa: E402
import openpyxl                                                # noqa: E402
from openpyxl.styles import Font, Border, Side, Alignment       # noqa: E402
from openpyxl.utils import get_column_letter                    # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation    # noqa: E402

# ---------------------------------------------------------------- config ----
WAVE = "Wave 3"
GO_LIVE = "11/07/2026"                       # shown on the summary page
# SVC population source. 2026-08-26 (the analyst): SVC is DERIVED, never read from
# the ad_hoc "Curriculum Status by User - SVC.xlsx" — that file was only this
# summary export pre-filtered, and it goes stale (it was missing 48 real SVC
# role holders). A person counts as SVC when they hold a Simple Visit Coding job
# role in MVP AND appear in the Epic Curriculum Status by User Summary export,
# i.e. someone we actually have Epic status data on. Same rule as
# build_soft_live_list.py and report.lava_list.
SVC_SUMMARY_EXPORT = op.RAW_EPIC_STATUS_SUMMARY
# Lands in Main Reports (shared deliverable folder) since 2026-08-26 —
# previously data\reports\sql_pulls\. Only the NEWEST build stays there;
# every older dated copy is swept into Main Reports\LAVA Archive\ at the end
# of each run, so the shared folder always shows exactly one LAVA file.
OUT_NAME = "LAVA {wave} Census current as of {date}.xlsx"
ARCHIVE_SUBDIR = "LAVA Archive"
# Classes that never count as someone's "last class" or completion date:
#   2026-09-08 (her rule): Cogito Advanced Reporting sessions, and anything that
#   is not a classroom Session (Online Class / Test = web-based training).
#   2026-09-11 (her list, after the compare with the trainer's Epic-only schedule):
#   Dolbey Fusion, PwC SMART, Charge Reconciliation Workshops and Epic Super
#   User observation sessions — vendor / workshop / observation sessions that
#   sit in Cornerstone under the Epic provider but are not core Epic classes.
#   Matched by family so every site / hours / wave variant is caught.
EXCLUDED_CLASS_PATTERNS = ["Advanced Reporting", "Dolbey Fusion", "PwC SMART",
                           "Charge Reconciliation Workshop", "Epic Super User"]
EXCLUDED_CLASS_SQL = "Training_Type='Session'" + "".join(
    f" AND Training_Title NOT LIKE '%{pat}%'" for pat in EXCLUDED_CLASS_PATTERNS)
EXCLUDED_TITLE_SQL = "".join(
    f" AND Training_Title NOT LIKE '%{pat}%'" for pat in EXCLUDED_CLASS_PATTERNS)


def is_excluded_title(title) -> bool:
    t = str(title).lower()
    return any(pat.lower() in t for pat in EXCLUDED_CLASS_PATTERNS)


def is_countable_class(title, ttype) -> bool:
    return str(ttype).strip() == "Session" and not is_excluded_title(title)

OUT_GLOB = "LAVA * Census current as of *.xlsx"
# Pre-2026-09-08 builds carried the old name; still swept into the archive.
LEGACY_GLOBS = ("LAVA * Preliminary User List *.xlsx",)

# Epic workqueue ownership census (boss ask 2026-09-02). The export names its
# owners by EMAIL ONLY, so they are resolved to Universal IDs through MVP
# UserUPN -> HR email -> email local-part (the local-part IS the UID at
# YourOrg). Both tabs feed the list: the census tab flags every owner, and the
# DNFB tab is carried as its own population the way SVC is, because DNFB owners
# are NOT all rev cycle under Lastname03 and so are mostly absent from the wave file.
# WAVE IS TAKEN FROM MVP, NOT THE WAVE FILE. Of the 71 DNFB owners, zero are
# Wave 3 on the wave file but 56 are Wave 3 in MVP — the wave file only carries
# rev-cycle staff, and these people are exactly the ones it does not cover.
# Leader values that mean the person has left rev cycle — never on the LAVA
# wave population (2026-09-09). Renamed 2026-06-05; both spellings still occur.
NOT_REV_CYCLE = {"No Longer Rev Cycle", "Not Rev Cycle"}
WQ_CENSUS_SHEET = "Epic Workqueue Ownership (83)"
WQ_DNFB_SHEET = "Wave 3 DNFB Owners"
WQ_OWNER_COL = "WQ Owner (Supervisor) (Input)"      # matched loosely, it has a trailing space
WQ_CORE_COL = "Core (Y is Rev Cycle (Dr Lastname03))"

# Which job categories mean "SVC" and which job titles mean "leader".
SVC_ROLE_PATTERN = r"simple visit coding"
# Supervisor and above. 'Lead <role>' is an individual contributor and is excluded
# on purpose; 'Management' must not match on 'Manager'.
LEADER_TITLE_PATTERN = r"\b(supervisor|spvr|superviser|manager|mgr|director|avp|vp|vice\s+president|chief)\b"

# Columns a human fills in; carried forward across rebuilds.
# Carried forward across rebuilds because nothing else records them. 'Training Complete?'
# is deliberately NOT here: Epic recomputes it every morning, so carrying a typed answer
# forward lets it silently outrank live data (74 stale 'No's survived that way on 08-25).
REVIEW_COLUMNS = ["FEC Access (Yes/No)", "Soft Live Access (Yes/No)"]
# 'Workqueue Owner' left REVIEW_COLUMNS on 2026-09-02: it now comes from the
# workqueue ownership census, so it is derived like SVC rather than typed. Same
# reasoning that keeps 'Training Complete?' out — carrying a typed answer
# forward lets it silently outrank the live source.
YESNO_COLUMNS = REVIEW_COLUMNS + ["Training Complete? (Yes/No)", "Leader (Yes/No)",
                                  "Vendor (Yes/No)", "Simple Visit Coding (Yes/No)",
                                  "Workqueue Owner (Yes/No)", "DNFB Owner (Yes/No)"]
EQUIV_LABEL = "Completed (Equivalent)"
# Larry's rule (via the analyst, 2026-09-15): only the CURRENT population is counted.
# Anyone not currently in the target wave (unless on the Wave 3 DNFB Owners
# list) and anyone on leave / terminated per HR or Epic moves to the
# "Not Current" sheet with a reason and is left out of every total.
NOT_CURRENT_STATUSES = {"pay leave", "leave and absent", "terminated", "severance",
                        "unaccounted for (zombie records)"}
NOT_CURRENT_SHEET = "Not Current"
HEADER_ROW = 5                                # table header lives on sheet row 5

# (Header, internal field, column width). Reorder or extend freely.
COLUMNS = [
    ("UniversalID", "UniversalID", 14), ("Full Name", "Full Name", 24),
    ("First Name", "First Name", 14), ("Last Name", "Last Name", 16),
    ("Population", "Population", 16), ("On Wave File", "On Wave File", 11),
    ("Wave", "Wave", 9), ("Leader", "Leader", 20), ("Director", "Director", 20),
    ("Sr. Director", "SeniorDirector", 20), ("AVP", "AVP", 20), ("VP", "VP", 18),
    ("User Type", "User Type", 11), ("Leader (Yes/No)", "Leader YN", 12),
    ("Vendor (Yes/No)", "Vendor YN", 12),
    ("Is Guesthouse (True/False)", "Is Guesthouse", 20),
    ("Simple Visit Coding (Yes/No)", "SVC YN", 15), ("SVC Job Role", "SVC Job Role", 36),
    ("Job Role 1", "JobRole1", 32), ("Job Role 2", "JobRole2", 30),
    ("Job Role 3", "JobRole3", 26), ("Job Role 4", "JobRole4", 22),
    ("Epic Job Category", "EpicJobCategory", 28), ("Curriculum Type", "Curriculum type", 24),
    ("Job Title (HR)", "JobTitle", 30), ("Department", "Department", 26),
    ("Business Unit (MVP)", "BusinessUnitDescription", 24), ("Email", "Email", 26),
    ("Direct Manager", "Direct Manager", 16),
    ("# Required Curricula", "# Required Cirrculums", 13),
    ("# Registered", "# Registered Curriculums", 12),
    ("# Completed", "# Completed Curriculums", 12),
    ("Training Status", "Training Status", 20),
    ("Est. Training Completion Date", "Est Completion", 24),
    ("Epic Status Fully Registered", "Fully Registered", 20),
    ("FEC Access (Yes/No)", "FEC", 13), ("Soft Live Access (Yes/No)", "SoftLive", 15),
    ("Workqueue Owner (Yes/No)", "Workqueue", 15),
    ("Training Complete? (Yes/No)", "Training Complete", 22),
    # Appended 2026-09-02, deliberately at the END. Everything above keeps the
    # column letter it had, because the ask referred to 'column AL' by letter --
    # inserting 'Last Training Class' next to the date it belongs with would have
    # pushed Workqueue Owner from AL to AM and broken that reference.
    ("Last Training Class", "Last Class", 34),
    ("DNFB Owner (Yes/No)", "DNFB", 13), ("# Workqueues Owned", "WQ Count", 13),
    ("Notes", "Notes", 52),
    # 2026-09-09 (VP's ask relayed by the analyst: "last date of training and class
    # for all wave 3 staff"): the last classroom session the person actually
    # sat, from Cornerstone. Sessions only — online classes, tests and Advanced
    # Reporting never count. Appended at the end for the same column-letter reason.
    ("Last Training Date", "Last Training Date", 18),
    # 2026-09-10 (her ask): a running, hand-maintained change log per person —
    # "9/10/26 Soft Live to Yes per email; 9/12/26 ...". Source = raw\wave\n    # LAVA Change Log.xlsx (see change_log()), never typed here: the census is
    # rebuilt every morning, so a typed note would not survive. Appended last.
    ("Change Log", "Change Log", 60),
]

# Off-wave DNFB owners (Wave 1 / Wave 2 staff activating with the Wave 3 DNFB
# workqueues) are already live in Epic from their own wave. VP's instruction
# 2026-09-09: "would be NA as they are in the system. They can just have a
# training completed indicator of yes."
PRIOR_WAVE_NA = "N/A"

INK, MUTED, RULE, ACCENT = "37352F", "787774", "E9E9E7", "2F6F6B"


def styles():
    """Notion-style formats: no fills, thin grey rules, one muted accent."""
    f = lambda **k: Font(name="Segoe UI", **k)                        # noqa: E731
    return {
        "title": f(size=14, bold=True, color=INK),
        "big": f(size=18, bold=True, color=INK),
        "kpi": f(size=26, bold=True, color=INK),
        "kpi_accent": f(size=26, bold=True, color=ACCENT),
        "sub": f(size=9, color=MUTED),
        "section": f(size=11, bold=True, color=INK),
        "head": f(size=9, bold=True, color=MUTED),
        "body": f(size=10, color=INK),
        "key": f(size=10, color=ACCENT),
        "bold": f(size=10, bold=True, color=INK),
        "rule": Border(bottom=Side(style="thin", color=RULE)),
    }


# ------------------------------------------------------------------ data ----
def svc_summary_export() -> Path:
    """The Epic Curriculum Status by User Summary export — the SVC population
    source since 2026-08-26 (see SVC_SUMMARY_EXPORT)."""
    if not SVC_SUMMARY_EXPORT.exists():
        sys.exit(
            f"Epic summary export not found: {SVC_SUMMARY_EXPORT}. Drop "
            "'Curriculum Status by User Summary.xlsx' there, or pass --svc-export.")
    return SVC_SUMMARY_EXPORT


def email_to_uid(eng) -> dict:
    """Email -> UniversalID. MVP UserUPN first (per-person address, the widest
    coverage), then HR's Email, then the local-part, which at YourOrg IS the
    Universal ID (jdoe@yourorg.example -> JDOE). UserUPN lives only in the MVP
    CSV — raw.mvp does not carry that column."""
    emap = {}
    upn = pd.read_csv(op.RAW_MVP_USER_MAPPINGS, usecols=["UniversalID", "UserUPN"],
                      dtype=str, on_bad_lines="skip")
    upn["e"] = upn["UserUPN"].astype(str).str.strip().str.lower()
    for e, u in zip(upn.loc[upn["e"].str.contains("@", na=False), "e"],
                    upn.loc[upn["e"].str.contains("@", na=False), "UniversalID"]):
        emap[e] = str(u).strip().upper()
    hr = pd.read_sql("SELECT UniversalID, Email FROM report.hr WHERE Email IS NOT NULL", eng)
    hr["e"] = hr["Email"].astype(str).str.strip().str.lower()
    for e, u in zip(hr.loc[hr["e"].str.contains("@", na=False), "e"],
                    hr.loc[hr["e"].str.contains("@", na=False), "UniversalID"]):
        emap.setdefault(e, str(u).strip().upper())
    return emap


def wq_owners(args, eng, known_uids):
    """Read the workqueue ownership census.

    DNFB owners are included REGARDLESS OF THEIR OWN WAVE (her boss, 2026-09-02:
    "although they are not part of Wave 3, they are newly activating with the
    Wave 3 DNFB workqueue, so they should be included"). 15 of the 71 are Wave 1
    or Wave 2 staff; they still need PROD LAVA access at Wave 3 go-live, and they
    carry a Notes entry so IT can pick them out.

    Returns a dict: census = owners whose own MVP wave matches the target wave
    (drives row inclusion); all_owners = every resolved owner at any wave (drives
    the Workqueue Owner flag, which is a fact about the person, not about a wave);
    dnfb = every DNFB owner at any wave; counts = workqueues owned; wave_of = each
    owner's MVP wave, for the Notes column; unresolved = emails with no match."""
    src = args.wq_export
    if src is None or not Path(src).exists():
        print("No workqueue ownership export found — Workqueue/DNFB columns left blank.")
        return dict(census=set(), all_owners=set(), dnfb=set(), counts={},
                    wave_of={}, unresolved=[])

    emap = email_to_uid(eng)
    mvpwave = pd.read_sql("SELECT UPPER(LTRIM(RTRIM(UniversalID))) _u, GoLiveWave"
                          " FROM raw.mvp", eng).drop_duplicates("_u").set_index("_u")["GoLiveWave"]
    mvpwave = mvpwave.to_dict()
    unresolved = []

    def uid(email):
        e = str(email).strip().lower()
        if "@" not in e:
            return None
        u = emap.get(e) or (e.split("@")[0].strip().upper() if
                            e.split("@")[0].strip().upper() in known_uids else None)
        if u is None:
            unresolved.append(e)
        return u

    def emails(sheet, col=None):
        """Owner emails off one tab. The DNFB tab is a bare column with no
        header, so its first row is data and must not be eaten as one."""
        if col is None:
            s = pd.read_excel(src, sheet_name=sheet, header=None, dtype=str,
                              engine="calamine").iloc[:, 0]
        else:
            df = pd.read_excel(src, sheet_name=sheet, dtype=str, engine="calamine")
            hit = next((c for c in df.columns if c.strip().lower() == col.strip().lower()), None)
            if hit is None:
                sys.exit(f"'{sheet}' has no '{col}' column. Columns: {list(df.columns)[:8]}")
            s = df[hit]
        return s.dropna().astype(str).str.strip().str.lower()

    census = emails(WQ_CENSUS_SHEET, WQ_OWNER_COL)
    counts_by_email = census.value_counts()
    census_map = {e: uid(e) for e in counts_by_email.index}
    dnfb_map = {e: census_map.get(e, uid(e)) for e in set(emails(WQ_DNFB_SHEET))}

    def in_wave(m):
        return {u for u in m.values() if u and mvpwave.get(u) == args.wave}

    census_ids = in_wave(census_map)
    all_owners = {u for u in census_map.values() if u} | {u for u in dnfb_map.values() if u}
    dnfb_ids = {u for u in dnfb_map.values() if u}      # every wave, deliberately
    counts = {}
    for e, n in counts_by_email.items():
        u = census_map.get(e)
        if u:
            counts[u] = counts.get(u, 0) + int(n)
    off_wave = {u for u in dnfb_ids if mvpwave.get(u) != args.wave}
    print(f"Workqueue census: {len(counts_by_email):,} owner emails "
          f"({len(unresolved)} unresolved) | {args.wave} owners: {len(census_ids)} "
          f"| DNFB owners: {len(dnfb_ids)} (off-wave, included anyway: {len(off_wave)})")
    return dict(census=census_ids, all_owners=all_owners, dnfb=dnfb_ids,
                counts=counts, wave_of={u: mvpwave.get(u) for u in all_owners},
                unresolved=sorted(set(unresolved)))


def change_log() -> dict:
    """UID -> "m/d/yy note; m/d/yy note" from raw/wave/LAVA Change Log.xlsx
    (sheet "Change Log": UniversalID, Date, Note, By). Oldest first. Missing or
    unreadable file = no notes, never a failure — the log is optional."""
    path = op.RAW_LAVA_CHANGE_LOG
    if not path.exists():
        return {}
    try:
        log = pd.read_excel(path, sheet_name="Change Log", dtype=str)
    except Exception as e:  # noqa: BLE001
        print(f"change log unreadable, skipped: {e}")
        return {}
    log = log.dropna(subset=["UniversalID", "Note"])
    log["_u"] = log["UniversalID"].astype(str).str.strip().str.upper()
    log["_d"] = pd.to_datetime(log["Date"], errors="coerce")
    log = log.sort_values(["_u", "_d"], na_position="last")
    out = {}
    for u, g in log.groupby("_u", sort=False):
        parts = [(f"{d.month}/{d.day}/{d:%y} " if pd.notna(d) else "") + str(n).strip()
                 for d, n in zip(g["_d"], g["Note"])]
        out[u] = "; ".join(parts)
    print(f"change log: {len(log)} entries for {len(out)} people")
    return out


def class_titles(eng, uids) -> dict:
    """(UniversalID, date) -> Epic class name, for the people on this list only.

    'Est. Training Completion Date' resolves to one of several dates depending on
    where the person is in training, so the class name is looked up by the date
    that actually won rather than recomputed — the two can never disagree.
    Scoped to the list's UIDs so this stays a small read of raw.cornerstone."""
    if not uids:
        return {}
    ids = ",".join("'" + u.replace("'", "''") + "'" for u in uids)
    df = pd.read_sql(f"""
        SELECT UPPER(LTRIM(RTRIM(User_ID))) _u, Training_Title, Training_Type, Transcript_Status,
               Training_Start_Date, Transcript_Completed_Date
        FROM raw.cornerstone
        WHERE Training_Provider='EPIC'
          AND UPPER(LTRIM(RTRIM(User_ID))) IN ({ids})
          AND Training_Title IS NOT NULL{EXCLUDED_TITLE_SQL}""", eng)
    out = {}
    # Start-date keys: classroom sessions only (web-based trainings and tests
    # have no start date anyway). Completed-date keys: equivalency credit rows
    # only, so an eLearning finished the same day never names itself the class.
    st = pd.to_datetime(df["Training_Start_Date"], errors="coerce").dt.normalize()
    for u, day, title, ttype in zip(df["_u"], st, df["Training_Title"], df["Training_Type"]):
        if pd.notna(day) and is_countable_class(title, ttype):
            out.setdefault((u, day), title)
    cd = pd.to_datetime(df["Transcript_Completed_Date"], errors="coerce").dt.normalize()
    for u, day, title, status in zip(df["_u"], cd, df["Training_Title"], df["Transcript_Status"]):
        if pd.notna(day) and "equivalent" in str(status).lower():
            out.setdefault((u, day), title)
    return out


def load(args):
    """Pull every source and return the assembled person-level frame."""
    eng = create_engine(CONN)
    # The Epic summary export: one row per person, everyone we have Epic status
    # data on. SVC is narrowed out of it below, by MVP job role.
    summary = pd.read_excel(args.svc_export, sheet_name=0, dtype=str, engine="calamine")
    uid_col = next((c for c in summary.columns
                    if c.strip().lower().replace(" ", "") == "universalid"), None)
    if uid_col is None:
        sys.exit(f"Epic summary export has no Universal Id column. Columns: {list(summary.columns)[:8]}")
    summary["_u"] = summary[uid_col].astype(str).str.strip().str.upper()
    summary = summary[summary["_u"].ne("") & summary["_u"].ne("NAN")].drop_duplicates("_u").set_index("_u")
    svc = summary        # joined below for its per-person Epic status columns

    users = pd.read_sql("SELECT * FROM report.users", eng)
    users["_u"] = users["UniversalID"].str.strip().str.upper()
    tagged = (users["Wave"] == args.wave) & (users["TrainingNeeded"] == "Yes")
    # Left rev cycle = left the RCM provisioning population (the analyst 2026-09-09:
    # "if no longer rev cycle as the leader should they be on lava report?" -> no).
    # HR moved them out of Dr Lastname03's org; they get access through their new
    # department's wave. Both spellings of the tag are synonyms.
    tagged &= ~users["Leader"].fillna("").str.strip().isin(NOT_REV_CYCLE)
    if args.leader_list_only:
        tagged &= (users["IsInScope"] == 1) & (users["OnLeaderList"] == 1)
    wave_ids = set(users.loc[tagged, "_u"])
    # Soft Go Live participants join whatever their wave — the one Soft Go Live group
    # sits on Wave 1/2 in MVP but activates with Wave 3 (the analyst, 2026-09-03).
    soft = users["SoftLiveParticipant"].fillna("").astype(str).str.strip().str.lower()
    soft_ids = set(users.loc[soft.str.startswith("y"), "_u"])

    jc = pd.read_sql("SELECT JobCategoryName, IsGuestHouse FROM raw.mvp_job_categories", eng)
    guest = {n.strip().upper() for n in
             jc.loc[jc["IsGuestHouse"].astype(str).str.strip().str.lower() == "true",
                    "JobCategoryName"].dropna()}
    svc_roles = {n.strip().upper() for n in
                 jc.loc[jc["JobCategoryName"].str.contains(SVC_ROLE_PATTERN, case=False, na=False),
                        "JobCategoryName"].dropna()}

    mvp = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(UniversalID))) _u, GoLiveWave,
        IndividualCategoryUpdate1Name JobRole1, IndividualCategoryUpdate2Name JobRole2,
        IndividualCategoryUpdate3Name JobRole3, IndividualCategoryUpdate4Name JobRole4,
        BusinessUnitDescription FROM raw.mvp""", eng).drop_duplicates("_u").set_index("_u")

    # SVC = holds a Simple Visit Coding job role in MVP AND we have Epic status
    # data on them (present in the summary export). Being in the summary export
    # alone is NOT SVC — that is the whole Epic population.
    jr = ["JobRole1", "JobRole2", "JobRole3", "JobRole4"]
    is_svc_role = mvp[jr].fillna("").apply(
        lambda r: any(str(v).strip().upper() in svc_roles for v in r if str(v).strip()), axis=1)
    svc_all = set(mvp.index[is_svc_role])
    with_data = svc_all & set(summary.index)
    svc_ids = set(with_data)
    if not args.all_waves_svc:
        svc_ids &= set(mvp.index[mvp["GoLiveWave"].fillna("").str.strip() == args.wave])
    print(f"SVC role holders in MVP: {len(svc_all):,} | with Epic status data: "
          f"{len(with_data):,} | in {args.wave}: {len(svc_ids):,}")

    # Change-request SVC (her ask 2026-09-15): a Wave Change Request Form that
    # assigns a Simple Visit Coding job role puts the person into the SVC
    # population the same day, before the MVP and Epic exports carry the new
    # role ("management need them on reports today"). The form's roles are
    # used until MVP shows them; once MVP catches up the person qualifies the
    # normal way and this path simply agrees. Wave = requested wave, else the
    # form's current wave, and it must be the target wave.
    wcr_roles = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(Universal_ID))) _u,
        Staff_Training_Job_Role_1 R1, Staff_Training_Job_Role_2 R2,
        Staff_Training_Job_Role_3 R3, Staff_Training_Job_Role_4 R4,
        Current_Wave, Requested_Wave_Change, _source_file
        FROM raw.wave_change_requests WHERE Universal_ID IS NOT NULL""", eng)
    svc_pat = re.compile(SVC_ROLE_PATTERN, re.I)
    wcr_roles["_wave"] = (wcr_roles["Requested_Wave_Change"].fillna("").astype(str).str.strip()
                          .replace("", pd.NA).fillna(wcr_roles["Current_Wave"].fillna("").astype(str).str.strip()))
    wcr_svc_roles = {}
    for _, r in wcr_roles.iterrows():
        roles = [str(r[c]).strip() for c in ("R1", "R2", "R3", "R4")
                 if pd.notna(r[c]) and str(r[c]).strip()]
        if str(r["_wave"]).strip() == args.wave and any(svc_pat.search(v) for v in roles):
            wcr_svc_roles.setdefault(r["_u"], [])
            wcr_svc_roles[r["_u"]] += [v for v in roles if v not in wcr_svc_roles[r["_u"]]]
    # Only IDs MVP knows: a mistyped Universal ID on a form must not become a
    # nameless census row (KROMAN on the HemOnc 6504 form, 2026-09-15).
    unknown = sorted(u for u in wcr_svc_roles if u not in mvp.index)
    if unknown:
        print(f"  skipped — Universal ID on a change request form not found in MVP: {', '.join(unknown)}")
        for u in unknown:
            wcr_svc_roles.pop(u, None)
    # Same gate as the MVP path: we must have Epic status data on the person.
    no_epic = sorted(u for u in wcr_svc_roles if u not in summary.index)
    if no_epic:
        print(f"  skipped — SVC role on a change request form but not in the Epic summary export: {', '.join(no_epic)}")
        for u in no_epic:
            wcr_svc_roles.pop(u, None)
    wcr_svc_ids = set(wcr_svc_roles) - svc_ids
    svc_ids |= wcr_svc_ids
    print(f"SVC via Wave Change Request Forms (role not yet in MVP): {len(wcr_svc_ids):,}"
          + (f" -> {', '.join(sorted(wcr_svc_ids))}" if wcr_svc_ids else ""))
    svc_raw = (summary.loc[sorted(svc_ids & set(summary.index))].reset_index() if svc_ids
               else summary.head(0).reset_index())

    hr = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(UniversalID))) _u, FullName HRName, Leader HRLeader,
        Director, SeniorDirector, AVP HRAVP, VP HRVP, Department, JobTitle, Email, WorkerType,
        AssignmentStatus HRStatus
        FROM report.hr""", eng).drop_duplicates("_u").set_index("_u")
    # Epic's HR status (Team Member Lookup) — catches leave / termination that
    # the bi-weekly HR export has not caught up with yet.
    eps = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(Universal_ID))) _u, MAX(Hr_Status_PDM) EpicHRStatus
        FROM raw.epic_lookup WHERE LTRIM(RTRIM(ISNULL(Universal_ID,''))) <> ''
        GROUP BY UPPER(LTRIM(RTRIM(Universal_ID)))""", eng).drop_duplicates("_u").set_index("_u")
    mstr = pd.read_sql("SELECT UPPER(LTRIM(RTRIM(UniversalID))) _u, FirstName, LastName"
                       " FROM raw.master", eng).drop_duplicates("_u").set_index("_u")
    track = pd.read_sql("""
        WITH equiv AS (SELECT UniversalID,
               MAX(CASE WHEN Event_Class_Status LIKE '%Equivalent%' THEN 1 ELSE 0 END) HasEquiv
            FROM report.tracker_detail WHERE Event_Class_Type='Session' AND IsExcluded=0
              AND (Event_Class_Registered LIKE '%Yes%' OR Event_Class_Status LIKE 'Completed%')
            GROUP BY UniversalID)
        SELECT UPPER(LTRIM(RTRIM(t.UniversalID))) _u, t.TrainingStatus, t.FinalScheduledDate,
               t.FullyRegisteredYN, t.FullyTrainedYN, t.LastRegisteredClass,
               CASE WHEN t.FinalScheduledDate IS NULL AND e.HasEquiv=1 THEN 1 ELSE 0 END EquivalencyOnly
        FROM report.tracker_training_status t LEFT JOIN equiv e ON e.UniversalID=t.UniversalID
        WHERE t.OnLeaderList=1""", eng).drop_duplicates("_u").set_index("_u")
    epic = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(Universal_Id))) _u,
        MAX(Epic_Job_Category) EpicJobCategory FROM raw.epic_status
        GROUP BY UPPER(LTRIM(RTRIM(Universal_Id)))""", eng).drop_duplicates("_u").set_index("_u")
    # Epic status for EVERYONE, not just wave-file people. report.tracker_*
    # joins report.users (IsInScope + canonical leader), so the SVC group came
    # back blank and these columns used to fall through to the summary export's
    # curriculum-level flags. report.epic_training_rollup is the same
    # aggregation the tracker runs, minus the wave-file gate — verified
    # identical to the tracker on all 6,914 on-list wave people. The tracker
    # still wins below; this fills in everyone it cannot see. (2026-08-26)
    roll = pd.read_sql("""SELECT UniversalID _u, TrainingStatus RollStatus,
        FullyRegisteredYN RollFR, FullyTrainedYN RollFT, ClassesRequired RollReq,
        ClassesRegistered RollReg, ClassesCompleted RollComp,
        FinalScheduledDate RollFinalDate, HasEquiv RollEquiv
        FROM report.epic_training_rollup""", eng).drop_duplicates("_u").set_index("_u")
    # LastCompletedClass = the last class they actually sat. Only real sessions carry a
    # Training_Start_Date; online guides and assessments have none, so requiring it keeps
    # optional coursework from masquerading as the final class. Epic dates no equivalency
    # row, so this is the only place a date exists for equivalency-credit people.
    # EquivCompleted dates the equivalency credit itself, for people who never sat a class:
    # the equivalency record's own completion date, else any completed transcript date.
    # LastSession = the last session they are actually booked for or sat. Withdrawn,
    # no-show and cancelled rows are excluded (2026-09-08: three "No Epic Curriculum"
    # people showed 9/22 and 10/05 dates they had withdrawn from).
    corn = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(User_ID))) _u,
        MAX(CASE WHEN Transcript_Status NOT LIKE 'Withdraw%' AND Transcript_Status NOT LIKE 'No Show%'
                  AND Transcript_Status NOT LIKE 'Cancel%' AND {EXCL} THEN Training_Start_Date END) LastSession,
        MAX(CASE WHEN Transcript_Status LIKE 'Completed%' AND Training_Start_Date IS NOT NULL
                  AND {EXCL} THEN Training_Start_Date END) LastCompletedClass,
        COALESCE(
          MAX(CASE WHEN Transcript_Status LIKE '%Equivalent%' THEN Transcript_Completed_Date END),
          MAX(CASE WHEN Transcript_Status LIKE 'Completed%' THEN Transcript_Completed_Date END)
        ) EquivCompleted
        FROM raw.cornerstone WHERE Training_Provider='EPIC'
        GROUP BY UPPER(LTRIM(RTRIM(User_ID)))""".replace("{EXCL}", EXCLUDED_CLASS_SQL),
        eng).drop_duplicates("_u").set_index("_u")
    # Epic's own date for equivalency credit (her rule 2026-09-15: a Fully Trained
    # person must show a date, never the label). Cornerstone carries no row at
    # all for people credited without sitting a class — Epic records the credit
    # in the Curriculum Status Detail with a Registration_Date; the latest one
    # across their classes is the date the credit was completed.
    epq = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(Universal_Id))) _u,
        MAX(TRY_CONVERT(datetime2, Registration_Date)) EpicEquivDate
        FROM raw.epic_status
        WHERE _snapshot_date = (SELECT MAX(_snapshot_date) FROM raw.epic_status)
          AND Event_Class_Status = 'Completed (Equivalent)'
          AND LTRIM(RTRIM(ISNULL(Universal_Id,''))) <> ''
        GROUP BY UPPER(LTRIM(RTRIM(Universal_Id)))""", eng).drop_duplicates("_u").set_index("_u")
    wcr = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(Universal_ID))) _u,
        MAX(FEC_Participant_Yes_No) WcrFEC, MAX(Soft_Live_Participant_Y_N) WcrSoft
        FROM raw.wave_change_requests WHERE Universal_ID IS NOT NULL
        GROUP BY UPPER(LTRIM(RTRIM(Universal_ID)))""", eng).drop_duplicates("_u").set_index("_u")

    wq = wq_owners(args, eng, set(mvp.index))
    # DNFB owners join the list whatever their own wave — they activate with the
    # Wave 3 DNFB workqueue (her boss, 2026-09-02).
    rows = sorted(wave_ids | svc_ids | soft_ids | wq["census"] | wq["dnfb"])
    titles = class_titles(eng, rows)
    emails = email_map(eng)
    print(f"{args.wave} tagged for training: {len(wave_ids)} | SVC: {len(svc_ids)} "
          f"| Soft Live: {len(soft_ids)} "
          f"| WQ owners: {len(wq['census'])} | DNFB: {len(wq['dnfb'])} "
          f"| total: {len(rows)}")

    keep = ["UniversalID", "FullName", "Wave", "Leader", "AVP", "VP", "UserType",
            "FECParticipant", "SoftLiveParticipant"]
    d = (pd.DataFrame(index=pd.Index(rows, name="_u"))
         .join(users.set_index("_u")[keep]).join(mstr).join(mvp).join(hr)
         .join(epic).join(corn).join(epq).join(track).join(roll).join(wcr).join(eps))
    svc_cols = [c for c in ["Team Member", "Fully Registered?", "Fully Trained?", "Curriculum type",
                            "Department / Unit", "Direct Manager", "# Required Cirrculums",
                            "# Registered Curriculums", "# Completed Curriculums"] if c in svc.columns]
    d = d.join(svc[svc_cols])
    d.attrs["wcr_svc_roles"] = wcr_svc_roles
    d.attrs["wcr_svc_ids"] = wcr_svc_ids
    return d, svc_raw, guest, svc_roles, wave_ids, svc_ids, soft_ids, wq, titles, emails


def derive(d, guest, svc_roles, wave_ids, svc_ids, soft_ids, wq, titles, emails,
           WAVE_LABEL=(WAVE,)):
    """Everything computed rather than looked up."""
    d["UniversalID"] = d["UniversalID"].fillna(pd.Series(d.index, index=d.index))
    d["Full Name"] = d["FullName"].fillna(d.get("HRName")).fillna(d.get("Team Member"))
    split = d["Full Name"].fillna("").str.split(",", n=1)
    d["First Name"] = d["FirstName"].fillna(split.str[-1].str.strip())
    d["Last Name"] = d["LastName"].fillna(split.str[0].str.strip())
    def population(u):
        """Every reason this person is on the list, most specific first. WQ
        owners are largely NOT rev cycle, so many carry that tag alone."""
        tags = [n for n, ids in (("Wave", wave_ids), ("SVC", svc_ids),
                                 ("Soft Live", soft_ids),
                                 ("WQ Owner", wq["census"])) if u in ids]
        if not tags and u in wq["dnfb"]:
            # Wave 1 / 2 staff carried only because they own a Wave 3 DNFB
            # workqueue (was blank, 2026-09-09).
            return "DNFB Owner (prior wave)"
        return {"Wave": "Wave Training", "SVC": "SVC Provisioning",
                "WQ Owner": "WQ Owner"}.get(tags[0], tags[0]) if len(tags) == 1             else " + ".join(tags)
    d["Population"] = [population(u) for u in d.index]
    d["On Wave File"] = ["Yes" if pd.notna(v) else "No" for v in d["FullName"]]
    d["Wave"] = d["Wave"].fillna(d["GoLiveWave"])
    d["Leader"] = d["Leader"].fillna(d["HRLeader"])
    d["AVP"] = d["AVP"].fillna(d["HRAVP"])
    d["VP"] = d["VP"].fillna(d["HRVP"])
    d["User Type"] = d["UserType"].fillna(
        d["WorkerType"].map({"VEN": "Vendor", "Employee": "YourOrg"}))
    # Email came from HR alone, so ~8% of the list — the SVC and WQ-owner people
    # with no HR row — showed blank. email_lookup adds the MVP UserUPN fallback
    # and is the same resolver the Master, Team File, Tracker and Soft Live list
    # use, so a person's address is identical in all of them. (2026-09-02)
    # The RESOLVER WINS. d["Email"] arrives off the HR attribute join, and HR is a
    # leader-mapping source only — its address is a contact field carrying vendor
    # domains (vendor.example, flexstaff.org) that Epic/LAVA cannot provision.
    # email_lookup is MVP UserUPN first with HR already behind it, so HR's column
    # is only a last resort here. Filling blanks alone left HR winning outright.
    filled = pd.Series(d.index.map(emails), index=d.index)
    d["Email"] = filled.fillna(d["Email"].replace("", pd.NA))

    jr = ["JobRole1", "JobRole2", "JobRole3", "JobRole4"]
    for c in jr:
        d[c] = d[c].fillna("").astype(str).str.strip()
    # Roles assigned on a Wave Change Request Form that MVP does not show yet
    # (2026-09-15): fill empty slots so SVC Job Role / guest house / SVC YN
    # read from the form until the MVP export carries the change.
    for uid, roles in d.attrs.get("wcr_svc_roles", {}).items():
        if uid not in d.index:
            continue
        have = {d.at[uid, c].upper() for c in jr if d.at[uid, c]}
        for role in roles:
            if role.upper() in have:
                continue
            slot = next((c for c in jr if not d.at[uid, c]), None)
            if slot is None:
                break
            d.at[uid, slot] = role
            have.add(role.upper())
    d["Is Guesthouse"] = d[jr].apply(lambda r: any(v.upper() in guest for v in r if v), axis=1)
    d["SVC Job Role"] = d[jr].apply(lambda r: next((v for v in r if v.upper() in svc_roles), ""), axis=1)
    d["SVC YN"] = ["Yes" if (u in svc_ids or role) else "No"
                   for u, role in zip(d.index, d["SVC Job Role"])]

    title = re.compile(LEADER_TITLE_PATTERN, re.I)
    d["Leader YN"] = d["JobTitle"].fillna("").apply(lambda t: "Yes" if title.search(t) else "No")
    d["Vendor YN"] = ["Yes" if str(w).strip().upper() == "VEN" or str(t).strip() in ("Onshore", "Offshore")
                      else "No" for w, t in zip(d["WorkerType"].fillna(""), d["UserType"].fillna(""))]

    def pick(*vals):
        """First real value wins — EXCEPT that an explicit Yes always beats a No.

        The Master defaults these flags to No for anyone not designated (her rule
        2026-09-02, fill_master_participation_defaults.py). Plain first-wins would
        then let that default outrank a genuine Yes arriving later on a wave change
        request, silently dropping a participant. Yes is the assertion; No is the
        absence of one, so Yes wins wherever any source says it.

        Never blank (her rule 2026-09-09: "those columns should not be blank. if
        there is no info for that, then put no"). Someone with no value in any
        source — typically an SVC or DNFB person who is not on the wave file —
        reads No."""
        seen = []
        for v in vals:
            s = str(v).strip()
            if s and s.lower() not in ("nan", "none", "nat"):
                seen.append("Yes" if s.lower().startswith("y")
                            else ("No" if s.lower().startswith("n") else s))
        if not seen:
            return "No"
        return "Yes" if "Yes" in seen else seen[0]

    d["FEC"] = [pick(a, b) for a, b in zip(d["FECParticipant"], d["WcrFEC"])]
    d["SoftLive"] = [pick(a, b) for a, b in zip(d["SoftLiveParticipant"], d["WcrSoft"])]
    # From the workqueue ownership census (2026-09-02), no longer hand-typed.
    # Owning a workqueue is a fact about the PERSON, so the flag ignores wave —
    # otherwise the off-wave DNFB owners we deliberately include would read "No".
    d["Workqueue"] = ["Yes" if u in wq["all_owners"] else "No" for u in d.index]
    d["DNFB"] = ["Yes" if u in wq["dnfb"] else "No" for u in d.index]
    d["WQ Count"] = [wq["counts"].get(u) for u in d.index]

    # Notes — her boss's ask (2026-09-02): "add a note for these individuals in
    # the Notes column so I can easily identify and highlight them for IT."
    # Terse category label, not a sentence.
    wcr_svc_ids = d.attrs.get("wcr_svc_ids", set())

    def note(u):
        if u in wq["dnfb"]:
            w = wq["wave_of"].get(u)
            if w and str(w).strip() != WAVE_LABEL[0]:
                return f"DNFB owner — {str(w).strip()} staff, activating with {WAVE_LABEL[0]} DNFB"
        if u in wcr_svc_ids:
            return "SVC role per Wave Change Request Form"
        return ""
    d["Notes"] = [note(u) for u in d.index]
    clog = change_log()
    d["Change Log"] = [clog.get(u, "") for u in d.index]

    # Epic-side guard: if the tracker's final scheduled class is an excluded
    # family (Advanced Reporting, Dolbey, PwC, workshops, observations), it does
    # not count — fall through to the Cornerstone-based dates below.
    if "LastRegisteredClass" in d.columns:
        adv = d["LastRegisteredClass"].map(is_excluded_title).astype(bool)
        d.loc[adv, "FinalScheduledDate"] = pd.NaT
        d.loc[adv, "LastRegisteredClass"] = None
    d["FinalScheduledDate"] = pd.to_datetime(
        d["FinalScheduledDate"].fillna(d["RollFinalDate"]), errors="coerce")
    d["LastSession"] = pd.to_datetime(d["LastSession"], errors="coerce")
    d["LastCompletedClass"] = pd.to_datetime(d["LastCompletedClass"], errors="coerce")
    d["EquivCompleted"] = pd.to_datetime(d["EquivCompleted"], errors="coerce").fillna(
        pd.to_datetime(d["EpicEquivDate"], errors="coerce"))
    # This column is always a real date. Already finished? Show the last class they actually
    # sat. Earned equivalency credit without ever sitting one? Show the date that credit was
    # recorded. Still training? The estimate is forward-looking, so keep falling back to their
    # next scheduled session — never rewrite it to a past date. EQUIV_LABEL survives only as a
    # guard for an equivalency person with no date anywhere in Cornerstone.
    # Tracker first, then the Epic rollup — resolved here because the
    # completion-date chain below keys off 'Fully Trained'. The summary export's
    # "Fully Registered?" / "Fully Trained?" are deliberately NOT used any more:
    # they are person-level Epic flags with no class counts, status or dates
    # behind them, so the SVC group read as registered with nothing to show for
    # it. The rollup carries the same flags PLUS the session detail. (2026-08-26)
    d["Fully Registered"] = d["FullyRegisteredYN"].fillna(d["RollFR"])
    d["Fully Trained"] = d["FullyTrainedYN"].fillna(d["RollFT"])
    d["EquivalencyOnly"] = d["EquivalencyOnly"].fillna(
        d["RollEquiv"].where(d["FinalScheduledDate"].isna(), 0))
    trained = d["Fully Trained"].astype(str).str.strip().eq("Yes")
    d["Est Completion"] = [f if pd.notna(f) else
                           (lc if (t and pd.notna(lc)) else
                            (eq if ((e == 1 or t) and pd.notna(eq)) else
                             (EQUIV_LABEL if e == 1 else (s if pd.notna(s) else None))))
                           for f, lc, t, e, eq, s in zip(d["FinalScheduledDate"], d["LastCompletedClass"],
                                                         trained, d["EquivalencyOnly"].fillna(0),
                                                         d["EquivCompleted"], d["LastSession"])]
    # No Epic record of any kind -> say so, the way the tracker and
    # report.lava_list both do. Blank would read as "not looked up".
    d["Training Status"] = (d["TrainingStatus"].fillna(d["RollStatus"])
                            .fillna("No Epic Curriculum"))
    # class counts: tracker's, else the rollup's, else the summary's
    for col, roll_col, sum_col in (
            ("# Required Cirrculums", "RollReq", "# Required Cirrculums"),
            ("# Registered Curriculums", "RollReg", "# Registered Curriculums"),
            ("# Completed Curriculums", "RollComp", "# Completed Curriculums")):
        have = d[sum_col] if sum_col in d.columns else pd.Series(index=d.index, dtype=object)
        d[col] = d[roll_col].fillna(have)
    # The class the date above belongs to. LastRegisteredClass is the tracker's own
    # name for the final scheduled session, so it wins when that date is the one
    # showing; otherwise the class is looked up by date in Cornerstone.
    def last_class(u, est, sched):
        if not isinstance(est, pd.Timestamp) or pd.isna(est):
            return ""
        if pd.notna(sched) and est == sched and isinstance(u, str):
            named = d.at[u, "LastRegisteredClass"] if "LastRegisteredClass" in d.columns else None
            if pd.notna(named) and str(named).strip():
                return str(named).strip()
        return titles.get((u, est.normalize()), "")
    d["Last Class"] = [last_class(u, e, f) for u, e, f
                       in zip(d.index, d["Est Completion"], d["FinalScheduledDate"])]
    # An equivalency date has no class title behind it: label it so the reader
    # knows the date is the credit, not a session (2026-09-15).
    eq_dated = [isinstance(e, pd.Timestamp) and pd.notna(q) and e.normalize() == q.normalize() and not c
                for e, q, c in zip(d["Est Completion"], d["EquivCompleted"], d["Last Class"])]
    d.loc[eq_dated, "Last Class"] = EQUIV_LABEL
    d["Training Complete"] = d["Fully Trained"].where(d["Fully Trained"].isin(["Yes", "No"]), "")
    # Last classroom session actually sat (Cornerstone, sessions only, Advanced
    # Reporting excluded) — for everyone, in progress or done. A Fully Trained
    # person credited by equivalency never sat one, so the credit date stands
    # in (her rule 2026-09-15: both date columns show their completion).
    d["Last Training Date"] = pd.to_datetime(d["LastCompletedClass"], errors="coerce").dt.normalize()
    eq_fill = trained & d["Last Training Date"].isna() & d["EquivCompleted"].notna()
    d.loc[eq_fill, "Last Training Date"] = d.loc[eq_fill, "EquivCompleted"].dt.normalize()
    # Prior-wave DNFB owners: live already, so training is not a question here.
    prior = [u for u in d.index if u in wq["dnfb"]
             and str(wq["wave_of"].get(u) or "").strip() not in ("", WAVE_LABEL[0])]
    if prior:
        d.loc[prior, "Training Status"] = [f"{PRIOR_WAVE_NA} - live ({str(wq['wave_of'].get(u)).strip()})"
                                           for u in prior]
        for col in ("Est Completion", "Last Class", "Last Training Date"):
            d[col] = d[col].astype(object)
            d.loc[prior, col] = PRIOR_WAVE_NA
        d.loc[prior, "Training Complete"] = "Yes"
    d["Department"] = d["Department"].fillna(d.get("Department / Unit"))
    return d.sort_values(["Population", "Leader", "Last Name", "First Name"],
                         key=lambda s: s.astype(str).str.upper())


def split_current(d, wq, wave):
    """Larry's rule (2026-09-15): tag everyone who is not part of the CURRENT
    population. Reasons are short category labels (shareable deliverable)."""
    hr_st = d["HRStatus"].fillna("").astype(str).str.strip().str.lower()
    ep_st = d["EpicHRStatus"].fillna("").astype(str).str.strip().str.lower()
    wave_now = d["Wave"].fillna("").astype(str).str.strip()
    dnfb = d.index.isin(wq["dnfb"])
    reason = pd.Series("", index=d.index, dtype=object)
    not_wave = wave_now.ne(wave) & ~dnfb
    reason[not_wave] = ["Not " + wave + (f" ({w})" if w else "") for w in wave_now[not_wave]]
    term = hr_st.isin({"terminated", "severance"}) | ep_st.isin({"terminated", "severance"})
    leave = (hr_st.isin(NOT_CURRENT_STATUSES) | ep_st.isin(NOT_CURRENT_STATUSES)) & ~term
    reason[leave & reason.eq("")] = "On leave"
    reason[term] = "No longer active"
    d["Not Current Reason"] = reason
    return d


def carry_forward(d, out: Path):
    """Keep hand-entered review answers from a previous build of the same file."""
    if not out.exists():
        return d, 0
    sheets = pd.ExcelFile(out).sheet_names
    prev = pd.concat([pd.read_excel(out, sheet_name=sh, header=HEADER_ROW - 1)
                      for sh in ("LAVA List", NOT_CURRENT_SHEET) if sh in sheets],
                     ignore_index=True)
    prev = prev[prev["UniversalID"].notna()]
    prev["_u"] = prev["UniversalID"].astype(str).str.strip().str.upper()
    kept = 0
    for header in REVIEW_COLUMNS:
        field = next((f for h, f, _w in COLUMNS if h == header), None)
        if field is None or header not in prev.columns:
            continue
        prior = prev.set_index("_u")[header].dropna().astype(str).str.strip()
        prior = prior[prior.ne("") & prior.ne("nan")]
        for uid, val in prior.items():
            if uid not in d.index or str(d.at[uid, field]).strip() == val:
                continue
            # Same precedence as pick(): Yes is the assertion, No is the absence
            # of one. A carried-forward No must never override a fresh Yes from
            # the Master / a change request (2026-09-11: 51 Front Desk people
            # flipped to Yes on the Master kept reading No on the census).
            if val.strip().upper() == "NO" and str(d.at[uid, field]).strip().upper() == "YES":
                continue
            d.at[uid, field] = val
            kept += 1
    return d, kept


# ----------------------------------------------------------------- write ----
def write_list_sheet(wb, d, st, args, sheet_name="LAVA List", subtitle=None, extra_col=None):
    """extra_col = (header, field) appended after the standard columns (used by
    the Not Current sheet for its reason)."""
    cols = list(COLUMNS) + ([(extra_col[0], extra_col[1], 30)] if extra_col else [])
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = (f"LAVA — {args.wave} Census · current as of {args.stamp:%m/%d/%Y}"
                if sheet_name == "LAVA List" else
                f"LAVA — {args.wave} Census · {sheet_name} · as of {args.stamp:%m/%d/%Y}")
    ws["A1"].font = st["title"]
    ws["A2"] = subtitle or (f"All {args.wave} staff tagged for training + all SVC (Simple Visit Coding) staff "
                            f"with provisioning data · {len(d):,} people · data as of {args.stamp:%m/%d/%Y %I:%M %p}")
    ws["A2"].font = st["sub"]
    ws["A3"] = ("Guest house = MVP IsGuestHouse on any assigned job role (see Guest House Job Roles "
                "sheet). Est. completion = last scheduled Epic classroom session (Advanced Reporting, "
                "vendor, workshop, observation and web-based sessions excluded). 'Training Complete?' is "
                "prefilled from Epic and editable. FEC / soft live read No unless the wave file, "
                "a wave change request or a hand-typed answer says Yes.")
    ws["A3"].font = st["sub"]

    for i, (label, _f, width) in enumerate(cols, start=1):
        c = ws.cell(row=HEADER_ROW, column=i, value=label)
        c.font, c.border = st["head"], st["rule"]
        c.alignment = Alignment(vertical="bottom", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[HEADER_ROW].height = 30

    centred = {"Leader YN", "Vendor YN", "SVC YN", "FEC", "SoftLive", "Workqueue",
               "Training Complete", "Is Guesthouse"}
    for r, (_, rec) in enumerate(d.iterrows(), start=HEADER_ROW + 1):
        for i, (_label, field, _w) in enumerate(cols, start=1):
            v = rec.get(field)
            c = ws.cell(row=r, column=i)
            c.font = st["key"] if i == 1 else st["body"]
            c.border = st["rule"]
            if field == "Is Guesthouse":
                c.value = "True" if v else "False"
            elif field in ("Est Completion", "Last Training Date") and isinstance(v, pd.Timestamp):
                c.value = v.to_pydatetime()
                c.number_format = "m/d/yyyy"
            else:
                c.value = "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v)
            if field in centred:
                c.alignment = Alignment(horizontal="center")

    last = HEADER_ROW + len(d)
    ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(len(COLUMNS))}{last}"
    ws.freeze_panes = f"C{HEADER_ROW + 1}"
    for label, _field, _w in COLUMNS:
        if label in YESNO_COLUMNS:
            col = get_column_letter([h for h, _f, _w2 in COLUMNS].index(label) + 1)
            dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{col}{HEADER_ROW + 1}:{col}{last}")


def write_guesthouse_sheet(wb, d, st):
    eng = create_engine(CONN)
    gh = pd.read_sql("""SELECT DISTINCT LTRIM(RTRIM(JobCategoryName)) JobRole, OwningApplication,
        TypeofJobCategory FROM raw.mvp_job_categories
        WHERE LOWER(LTRIM(RTRIM(IsGuestHouse)))='true' AND JobCategoryName IS NOT NULL""", eng)
    gh = gh.drop_duplicates("JobRole").sort_values("JobRole", key=lambda s: s.str.upper())
    jr = ["JobRole1", "JobRole2", "JobRole3", "JobRole4"]
    held = pd.Series([str(v).strip().upper() for c in jr for v in d[c] if str(v).strip()]).value_counts()

    ws = wb.create_sheet("Guest House Job Roles")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Guest House Job Roles"
    ws["A1"].font = st["title"]
    ws["A2"] = (f"Every MVP job category flagged IsGuestHouse = True ({len(gh)} roles). Anyone holding "
                f"one of these in Job Role 1-4 is marked Guesthouse on the LAVA List. "
                f"Source: MVP JobCategoriesData.")
    ws["A2"].font = st["sub"]
    for i, (h, w) in enumerate([("Job Role", 62), ("Owning Application", 26),
                                ("Category Type", 22), ("People in This List", 18)], start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font, c.border = st["head"], st["rule"]
        c.alignment = Alignment(vertical="bottom", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, (_, rec) in enumerate(gh.iterrows(), start=5):
        vals = [rec["JobRole"], rec["OwningApplication"], rec["TypeofJobCategory"],
                int(held.get(rec["JobRole"].upper(), 0))]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=i, value="" if pd.isna(v) else v)
            c.font, c.border = st["body"], st["rule"]
            if i == 4:
                c.alignment = Alignment(horizontal="center")
    ws.auto_filter.ref = f"A4:D{4 + len(gh)}"
    ws.freeze_panes = "A5"
    return len(gh)


def write_raw_sheet(wb, svc_raw, st, src: Path):
    ws = wb.create_sheet("SVC Export (raw)")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "SVC provisioning export — source file, unmodified"
    ws["A1"].font = st["title"]
    ws["A2"] = f"{len(svc_raw)} rows as received · {src.name}"
    ws["A2"].font = st["sub"]
    for i, col in enumerate(svc_raw.columns, start=1):
        c = ws.cell(row=4, column=i, value=str(col))
        c.font, c.border = st["head"], st["rule"]
        c.alignment = Alignment(vertical="bottom", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = 20
    for r, (_, rec) in enumerate(svc_raw.iterrows(), start=5):
        for i, col in enumerate(svc_raw.columns, start=1):
            v = rec[col]
            c = ws.cell(row=r, column=i,
                        value="" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v))
            c.font, c.border = st["body"], st["rule"]
    ws.auto_filter.ref = f"A4:{get_column_letter(len(svc_raw.columns))}{4 + len(svc_raw)}"
    ws.freeze_panes = "A5"


def write_summary_sheet(wb, d, st, args, n_guest_roles):
    """One-page executive view: 4 KPIs, four compact tables, no charts."""
    n = len(d)
    gh = d["Is Guesthouse"].astype(bool)
    vend = d["Vendor YN"].eq("Yes")
    lead = d["Leader YN"].eq("Yes")
    done = d["Training Complete"].astype(str).str.strip().eq("Yes")
    reg = d["Fully Registered"].astype(str).str.strip().eq("Yes")
    onfile = d["On Wave File"].eq("Yes")
    fec = d["FEC"].astype(str).str.strip().replace("nan", "")
    soft = d["SoftLive"].astype(str).str.strip().replace("nan", "")
    dates = pd.Series([v for v in d["Est Completion"] if isinstance(v, pd.Timestamp)])

    ws = wb.create_sheet("Executive Summary", 0)
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFG", [30, 11, 11, 4, 30, 11, 11]):
        ws.column_dimensions[col].width = w
    ws["A1"] = f"LAVA — {args.wave} Census · current as of {args.stamp:%m/%d/%Y}"
    ws["A1"].font = st["big"]
    ws["A2"] = (f"Executive summary · {n:,} users · {args.wave} go-live {GO_LIVE} · "
                f"data as of {args.stamp:%m/%d/%Y %I:%M %p}")
    ws["A2"].font = st["sub"]
    ws.row_dimensions[1].height = 26
    for col, val, label in [("A", f"{n:,}", "Total users needing LAVA"),
                            ("C", f"{int(gh.sum()):,}", "Guest house users"),
                            ("E", f"{int(vend.sum()):,}", "Vendor users"),
                            ("G", f"{done.sum()/n:.0%}", "Training complete")]:
        ws[f"{col}4"] = val
        ws[f"{col}4"].font = st["kpi_accent"] if col == "G" else st["kpi"]
        ws[f"{col}5"] = label
        ws[f"{col}5"].font = st["sub"]
    ws.row_dimensions[4].height = 34

    row = 7

    def block(title, headers, data_rows, note=None):
        nonlocal row
        ws.cell(row=row, column=1, value=title).font = st["section"]
        row += 1
        for i, h in enumerate(headers, start=1):
            c = ws.cell(row=row, column=i, value=h)
            c.font, c.border = st["head"], st["rule"]
            if i > 1:
                c.alignment = Alignment(horizontal="right")
        row += 1
        for rec in data_rows:
            bold = str(rec[0]).startswith("Total")
            for i, v in enumerate(rec, start=1):
                c = ws.cell(row=row, column=i, value=v)
                c.font = st["bold"] if bold else st["body"]
                c.border = st["rule"]
                if i > 1:
                    c.alignment = Alignment(horizontal="right")
                    if isinstance(v, float):
                        c.number_format = "0%"
            row += 1
        if note:
            c = ws.cell(row=row, column=1, value=note)
            c.font, c.alignment = st["sub"], Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=row, start_column=1, end_row=row + 1, end_column=3)
            row += 2
        row += 1

    # Every population label present, largest first, so the rows always add up
    # to the total (the fixed five-label list missed the workqueue / DNFB groups).
    pops = [(str(p), int(c), int(c) / n)
            for p, c in d["Population"].fillna("").value_counts().items() if p]
    block("Who is in scope", ["Group", "Users", "% of total"], pops + [("Total", n, 1.0)],
          note=f"{int((~onfile).sum()):,} users are not on the RCM wave file; they are included "
               f"for Simple Visit Coding or workqueue provisioning. "
               + (f"{args.n_not_current:,} people are listed on the Not Current sheet (not currently "
                  f"{args.wave}, on leave, or no longer active) and are not counted here."
                  if getattr(args, "n_not_current", 0) else ""))
    block("Training readiness", ["Status", "Users", "% of total"],
          [("Fully trained", int(done.sum()), done.sum() / n),
           ("Fully registered, not yet trained", int((reg & ~done).sum()), (reg & ~done).sum() / n),
           ("Not fully registered", int((~reg).sum()), (~reg).sum() / n),
           ("Total", n, 1.0)],
          note=(f"Estimated completion dates run {dates.min():%m/%d/%Y} to {dates.max():%m/%d/%Y} for "
                f"{len(dates):,} users; {n - len(dates):,} have no scheduled date yet."
                if len(dates) else "No estimated completion dates available."))
    block("Access profile", ["Attribute", "Users", "% of total"],
          [("Leaders (supervisor and above)", int(lead.sum()), lead.sum() / n),
           ("Vendor users", int(vend.sum()), vend.sum() / n),
           ("Guest house users", int(gh.sum()), gh.sum() / n),
           ("FEC access confirmed", int(fec.eq("Yes").sum()), fec.eq("Yes").sum() / n),
           ("Soft live access confirmed", int(soft.eq("Yes").sum()), soft.eq("Yes").sum() / n)],
          note="FEC and soft live access read No unless confirmed as Yes.")
    # By-leader breakdown (her ask 2026-09-10): one row per leader with the
    # provisioning and readiness counts VPs ask about, no commentary.
    leader = d["Leader"].fillna("(no leader on file)").astype(str)
    counts = leader.value_counts()
    dnfb = d["DNFB"].astype(str).str.strip().eq("Yes")
    svc = d["SVC YN"].astype(str).str.strip().eq("Yes")
    wq = d["Workqueue"].astype(str).str.strip().eq("Yes")
    metrics = [("DNFB", dnfb), ("FEC", fec.eq("Yes")), ("SVC", svc),
               ("Soft Live", soft.eq("Yes")), ("Guest House", gh), ("WQ Owners", wq),
               ("Registered", reg), ("Unregistered", ~reg), ("Fully Trained", done)]

    def leader_row(label, mask):
        return (label, int(mask.sum()), mask.sum() / n,
                *[int((m & mask).sum()) for _, m in metrics])

    top = [leader_row(str(k), leader.eq(str(k))) for k in counts.head(8).index]
    if len(counts) > 8:
        rest_mask = ~leader.isin([str(k) for k in counts.head(8).index])
        top.append(leader_row(f"All other leaders ({len(counts) - 8})", rest_mask))
    top.append(leader_row("Total", leader.notna()))
    for col in "HIJKL":
        ws.column_dimensions[col].width = 12
    block("Largest populations by leader",
          ["Leader", "Users", "% of total"] + [name for name, _ in metrics], top)

    # Last scheduled class by leader (her ask 2026-09-08): the latest Est. completion
    # date among each leader's people, and the class it belongs to. Classroom
    # sessions only — the EXCLUDED_CLASS_PATTERNS families and web-based
    # trainings never count.
    ws.cell(row=row, column=1, value="Last training class by leader").font = st["section"]
    row += 1
    for i, h in enumerate(["Leader", "Last class date", "Last training class"], start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font, c.border = st["head"], st["rule"]
        if i == 2:
            c.alignment = Alignment(horizontal="right")
    row += 1
    has_date = d["Est Completion"].apply(lambda v: isinstance(v, pd.Timestamp))
    dated = d[has_date & d["Leader"].notna() & d["Leader"].astype(str).str.strip().ne("")]
    for leader in sorted(dated["Leader"].astype(str).unique(), key=str.upper):
        grp = dated[dated["Leader"].astype(str) == leader]
        idx = grp["Est Completion"].astype("datetime64[ns]").idxmax()
        last_dt, last_cls = grp.at[idx, "Est Completion"], str(grp.at[idx, "Last Class"] or "")
        for i, v in enumerate([leader, last_dt.to_pydatetime(), last_cls], start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.font, c.border = st["body"], st["rule"]
            if i == 2:
                c.alignment = Alignment(horizontal="right")
                c.number_format = "mm/dd/yyyy"
        row += 1
    row += 1

    ws.cell(row=row, column=1,
            value=f"Detail: LAVA List sheet · {n_guest_roles} guest house definitions: Guest House Job "
                  f"Roles sheet · source export: SVC Export (raw) sheet.").font = st["sub"]
    ws.print_area = f"A1:G{row}"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


# ------------------------------------------------------------------ main ----
def sweep_old_builds(out: Path) -> int:
    """Keep only the newest LAVA workbook in the shared folder.

    Her rule (2026-08-26): Main Reports shows the current list and nothing else;
    every previous dated build belongs in Main Reports\\LAVA Archive\\. Runs after a
    successful write, so an archived file is only ever one we have replaced.
    Same-name collisions in the archive are kept, not clobbered.
    """
    archive = out.parent / ARCHIVE_SUBDIR
    moved = 0
    candidates = list(out.parent.glob(OUT_GLOB))
    for g in LEGACY_GLOBS:
        candidates += list(out.parent.glob(g))
    for old in sorted(set(candidates)):
        if old.resolve() == out.resolve() or ".bak." in old.name:
            continue
        archive.mkdir(parents=True, exist_ok=True)
        dest = archive / old.name
        if dest.exists():
            dest = archive / f"{old.stem}.{old.stat().st_mtime_ns}{old.suffix}"
        shutil.move(str(old), str(dest))
        print(f"  archived previous build: {ARCHIVE_SUBDIR}/{dest.name}")
        moved += 1
    return moved


def main():
    ap = argparse.ArgumentParser(description="Build the LAVA Wave census (working LAVA list).")
    ap.add_argument("--wave", default=WAVE, help=f"Wave to build (default {WAVE!r}).")
    ap.add_argument("--svc-export", type=Path,
                    help="Epic Curriculum Status by User Summary export "
                         "(default: raw/epic/epic_status_summary).")
    ap.add_argument("--out", help="Output filename or full path (default: dated name in sql_pulls).")
    ap.add_argument("--leader-list-only", action="store_true",
                    help="Restrict the wave population to canonical-leader people "
                         "(default: every tagged person, matching the original ask).")
    ap.add_argument("--all-waves-svc", action="store_true",
                    help="Include SVC role holders from every wave, not just --wave.")
    ap.add_argument("--wq-export",
                    help="Workqueue ownership export (default: newest in raw\\ad_hoc\\).")
    ap.add_argument("--no-preserve", action="store_true",
                    help="Do not carry hand-entered review answers forward from the previous build.")
    ap.add_argument("--show-config", action="store_true", help="Print settings and exit.")
    args = ap.parse_args()
    args.stamp = datetime.now()
    args.svc_export = args.svc_export or svc_summary_export()
    args.wq_export = Path(args.wq_export) if args.wq_export else op.latest_wq_owners()

    out = Path(args.out) if args.out else None
    if out is None:
        out = op.MAIN_REPORTS_DIR / OUT_NAME.format(wave=args.wave, date=f"{args.stamp:%Y-%m-%d}")
    elif out.parent == Path("."):
        out = op.MAIN_REPORTS_DIR / out.name

    if args.show_config:
        print(f"wave            : {args.wave}\nsvc export      : {args.svc_export}\n"
              f"wq export       : {args.wq_export}\n"
              f"output          : {out}\nleader-list only: {args.leader_list_only}\n"
              f"all waves SVC   : {args.all_waves_svc}\ncolumns         : {len(COLUMNS)}\n"
              f"review columns  : {', '.join(REVIEW_COLUMNS)}")
        return

    print(f"Epic summary export: {args.svc_export.name}  ({datetime.fromtimestamp(args.svc_export.stat().st_mtime):%m/%d/%Y %I:%M %p})")
    if args.wq_export:
        print(f"Workqueue export   : {args.wq_export.name}  "
              f"({datetime.fromtimestamp(args.wq_export.stat().st_mtime):%m/%d/%Y %I:%M %p})")
    d, svc_raw, guest, svc_roles, wave_ids, svc_ids, soft_ids, wq, titles, emails = load(args)
    d = derive(d, guest, svc_roles, wave_ids, svc_ids, soft_ids, wq, titles, emails, (args.wave,))

    kept = 0
    if not args.no_preserve:
        d, kept = carry_forward(d, out)
        if kept:
            print(f"carried forward {kept} hand-entered answer(s) from the previous build")
    if out.exists():
        # Backups live in Main Reports\Backups\<month>\, never beside the live
        # file — same convention as backup_deliverables.py.
        backup = (op.month_subdir(op.MAIN_REPORTS_BACKUPS_DIR, args.stamp.date())
                  / f"{out.stem}.bak.{args.stamp:%Y-%m-%d_%H%M}{out.suffix}")
        shutil.copy(str(out), str(backup))
        print(f"previous version backed up: {backup.parent.name}/{backup.name}")

    d = split_current(d, wq, args.wave)
    current = d[d["Not Current Reason"].eq("")]
    parked = d[d["Not Current Reason"].ne("")]
    args.n_not_current = len(parked)
    print(f"current population: {len(current):,} | Not Current sheet: {len(parked):,} "
          f"({', '.join(f'{k} {v}' for k, v in parked['Not Current Reason'].value_counts().items())})")

    st = styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_list_sheet(wb, current, st, args)
    n_guest = write_guesthouse_sheet(wb, current, st)
    write_raw_sheet(wb, svc_raw, st, args.svc_export)
    write_summary_sheet(wb, current, st, args, n_guest)
    write_list_sheet(
        wb, parked, st, args, sheet_name=NOT_CURRENT_SHEET,
        subtitle=(f"{len(parked):,} people not in the current {args.wave} population: not currently "
                  f"{args.wave} (Wave 3 DNFB Owners excepted), on leave, or no longer active per HR / Epic. "
                  f"Not counted in the Executive Summary. Data as of {args.stamp:%m/%d/%Y %I:%M %p}."),
        extra_col=("Not Current Reason", "Not Current Reason"))
    d = current
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)

    errors = {"#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "#SPILL!", "#CALC!"}
    chk = openpyxl.load_workbook(out, data_only=True)
    bad = sum(1 for s in chk.sheetnames for row in chk[s].iter_rows() for c in row
              if isinstance(c.value, str) and c.value.strip() in errors)
    dupes = int(d.index.duplicated().sum())
    archived = sweep_old_builds(out)

    print(f"\nwrote {len(d):,} rows x {len(COLUMNS)} columns -> {out}")
    if archived:
        print(f"  archived          : {archived} previous build(s) -> {ARCHIVE_SUBDIR}\\")
    print(f"  sheets            : {', '.join(chk.sheetnames)}")
    print(f"  guest house       : {int(d['Is Guesthouse'].astype(bool).sum())} of {len(d)}")
    print(f"  training complete : {int(d['Training Complete'].eq('Yes').sum())} Yes / "
          f"{int(d['Training Complete'].eq('No').sum())} No")
    print(f"  QA                : {bad} error value(s), {dupes} duplicate ID(s)")
    if bad or dupes:
        sys.exit("QA FAILED — do not send this file until it is fixed.")


if __name__ == "__main__":
    main()
