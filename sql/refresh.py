# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
refresh.py  —  RUN DAILY  (the whole job is: python sql\\refresh.py)

What it does:
  Reads the latest export from each source and (re)loads it into the
  matching raw.* table in SQL Server. Tables are created automatically,
  so there is no table DDL to maintain by hand.

  Sources (newest file of each type, read from OneDrive) -> raw tables:
    Master Wave File   -> raw.master        (DATA sheet)   = who's in / source of truth
    HR.xlsx  (raw)     -> raw.hr             = full 67-col HR detail
    HR_cleaned (proc.) -> raw.hr_cleaned     = resolved leader chain (AVP/VP/SVP/Leaders)
    MVP User Mappings  -> raw.mvp            = job roles
    MVP Role Mappings  -> raw.mvp_roles      = dept/job-code -> Epic job-category mapping
    MVP Job Categories -> raw.mvp_job_categories = Epic job-category reference
    Cornerstone        -> raw.cornerstone    = enterprise training transcripts
    Epic Curriculum    -> raw.epic_status    = per-user Fully Trained/Registered (per wave)
    Epic TM Lookup     -> raw.epic_lookup    = roster: Wave, training-needed, eligibility
    Epic Class Sched.  -> raw.epic_class_schedule = sessions: date/location/instructor/seats
    Wave Change Reqs   -> raw.wave_change_requests = add/edit/remove requests (form tab)

  Everything lands as text (NVARCHAR). The report views in
  3_report_views.sql do the typing, cleaning, and joining.
  File paths come from scripts/onedrive_paths.py (single source of truth).

  A full run ends with a `snapshot` step: today's report.roster and
  report.epic_not_in_hr are appended to history.* tables (one snapshot per
  day) and mirrored to dated CSVs on OneDrive, so day-over-day trend and
  change queries are possible and the history survives a lost/rebuilt DB.

Usage:
    python sql\\refresh.py                    # load everything + today's history snapshot
    python sql\\refresh.py mvp epic_status     # load only these
    python sql\\refresh.py snapshot            # just (re)stamp today's history
    python sql\\refresh.py backload_history    # rebuild history.* from the OneDrive CSVs
"""

from __future__ import annotations
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.types import NVARCHAR

# ----------------------------------------------------------------------
# CONFIG  — edit these if the server name or file locations change
# ----------------------------------------------------------------------
SERVER   = r".\SQLEXPRESS"
DATABASE = "AnalyticsDB"
ROOT     = Path(__file__).resolve().parent.parent   # yourorg_analytics/

# Reuse the OneDrive path constants + latest_*() helpers the other scripts use.
sys.path.insert(0, str(ROOT / "scripts"))
import onedrive_paths as op   # noqa: E402

CONN = (
    "mssql+pyodbc://@" + SERVER + "/" + DATABASE +
    "?driver=ODBC+Driver+18+for+SQL+Server"
    "&Trusted_Connection=yes&TrustServerCertificate=yes"
)

NVARLEN = 4000  # raw columns are NVARCHAR(4000); plenty for these files


# ----------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------
def clean_col(name: str) -> str:
    """Turn a messy header into a safe SQL column name."""
    s = re.sub(r"[^0-9a-zA-Z]+", "_", str(name).strip()).strip("_")
    if not s:
        s = "col"
    if s[0].isdigit():
        s = "c_" + s
    return s


def dedupe(cols):
    """Append _2, _3 ... to duplicate column names."""
    seen, out = {}, []
    for c in cols:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 1
            out.append(c)
    return out


def prepare(df: pd.DataFrame, source_file: str, extra: dict | None = None) -> pd.DataFrame:
    """Clean columns, add metadata, make everything text/None."""
    df = df.copy()
    df.columns = dedupe([clean_col(c) for c in df.columns])
    # drop completely empty rows
    df = df.dropna(how="all")
    # add metadata
    df["_source_file"] = Path(source_file).name
    df["_loaded_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if extra:
        for k, v in extra.items():
            df[k] = v
    # everything to clean strings (NaN -> None so SQL gets NULL)
    df = df.astype(object).where(pd.notna(df), None)
    df = df.map(lambda v: str(v).strip() if v is not None else None)
    return df


def write_table(engine, table: str, df: pd.DataFrame):
    # Stage-and-swap: load into raw._loading_<table>, then swap it in. An
    # interrupted load can never leave raw.<table> half-full — the old rows
    # survive until the (instant) swap.
    dtype = {c: NVARCHAR(NVARLEN) for c in df.columns}
    tmp = f"_loading_{table}"
    df.to_sql(
        tmp, engine, schema="raw", if_exists="replace", index=False,
        dtype=dtype, chunksize=1000, method=None,
    )
    with engine.begin() as con:
        con.execute(text(f"DROP TABLE IF EXISTS raw.[{table}]"))
        con.execute(text(f"EXEC sp_rename 'raw.[{tmp}]', '{table}'"))
    print(f"   -> raw.{table}: {len(df):,} rows x {len(df.columns)} cols")


# ----------------------------------------------------------------------
# per-source loaders
# ----------------------------------------------------------------------
def build_uids(engine, table: str, col: str):
    """Distinct clean UniversalIDs of raw.<table> -> raw.uids_<table>, with a
    unique clustered index. report.* anti-joins (e.g. report.exceptions) seek
    these tiny key tables instead of scanning the NVARCHAR(4000) heaps — the
    optimizer picks catastrophic nested-loop plans against the raw tables."""
    with engine.begin() as con:
        con.execute(text(f"DROP TABLE IF EXISTS raw.uids_{table}"))
        con.execute(text(
            f"SELECT DISTINCT CAST(UPPER(LTRIM(RTRIM([{col}]))) AS NVARCHAR(450)) AS UniversalID "
            f"INTO raw.uids_{table} FROM raw.[{table}] "
            f"WHERE LTRIM(RTRIM(ISNULL([{col}],''))) <> ''"))
        con.execute(text(
            f"CREATE UNIQUE CLUSTERED INDEX IX_uids_{table} ON raw.uids_{table}(UniversalID)"))


# ----------------------------------------------------------------------
# LAVA assembly — stage 3 of build_lava_list(). Kept as a module constant so
# the (long) column list stays out of the function body. Every derived column
# matches build_lava_list.py's derive(); see report.lava_list for what is and
# is not representable in SQL.
# ----------------------------------------------------------------------
LAVA_ASSEMBLE_SQL = """
SELECT
    p.u AS UniversalID,
    COALESCE(u.FullName, h.FullName, s.Team_Member) AS FullName,
    COALESCE(m.FirstName,
             LTRIM(SUBSTRING(COALESCE(u.FullName, h.FullName, s.Team_Member, ''),
                   NULLIF(CHARINDEX(',', COALESCE(u.FullName, h.FullName, s.Team_Member, '')), 0) + 1, 200))) AS FirstName,
    COALESCE(m.LastName,
             CASE WHEN CHARINDEX(',', COALESCE(u.FullName, h.FullName, s.Team_Member, '')) > 0
                  THEN LTRIM(RTRIM(LEFT(COALESCE(u.FullName, h.FullName, s.Team_Member, ''),
                       CHARINDEX(',', COALESCE(u.FullName, h.FullName, s.Team_Member, '')) - 1))) END) AS LastName,
    CASE WHEN p.InWave = 1 AND p.InSVC = 1 THEN 'Wave + SVC'
         WHEN p.InWave = 1                 THEN 'Wave Training'
         ELSE 'SVC Provisioning' END AS Population,
    CASE WHEN u.UniversalID IS NOT NULL THEN 'Yes' ELSE 'No' END AS OnWaveFile,
    COALESCE(u.Wave, v.GoLiveWave) AS Wave,
    COALESCE(u.Leader, h.Leader)   AS Leader,
    h.Director, h.SeniorDirector,
    COALESCE(u.AVP, h.AVP) AS AVP,
    COALESCE(u.VP,  h.VP)  AS VP,
    COALESCE(u.UserType, CASE h.WorkerType WHEN 'VEN' THEN 'Vendor'
                                           WHEN 'Employee' THEN 'YourOrg' END) AS UserType,
    -- Job title matched on whole words: punctuation becomes spaces and the
    -- string is padded, so ' manager ' cannot match 'management' and ' vp '
    -- cannot match 'avp'. Same intent as the script's \b(...)\b regex.
    CASE WHEN t2.TitlePad LIKE '% supervisor %' OR t2.TitlePad LIKE '% spvr %'
           OR t2.TitlePad LIKE '% superviser %' OR t2.TitlePad LIKE '% manager %'
           OR t2.TitlePad LIKE '% mgr %'        OR t2.TitlePad LIKE '% director %'
           OR t2.TitlePad LIKE '% avp %'        OR t2.TitlePad LIKE '% vp %'
           OR t2.TitlePad LIKE '% vice president %' OR t2.TitlePad LIKE '% chief %'
         THEN 'Yes' ELSE 'No' END AS LeaderYN,
    CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(h.WorkerType,'')))) = 'VEN'
           OR LTRIM(RTRIM(ISNULL(u.UserType,''))) IN ('Onshore','Offshore')
         THEN 'Yes' ELSE 'No' END AS VendorYN,
    CAST(ISNULL(rf.IsGuest, 0) AS bit) AS IsGuesthouse,
    CASE WHEN p.InSVC = 1 OR rf.SVCJobRole IS NOT NULL THEN 'Yes' ELSE 'No' END AS SVCYN,
    rf.SVCJobRole,
    rf.SVCSource,
    v.JobRole1, v.JobRole2, v.JobRole3, v.JobRole4,
    e.EpicJobCategory,
    s.Curriculum_type AS CurriculumType,
    h.JobTitle,
    COALESCE(h.Department, s.Department_Unit) AS Department,
    v.BusinessUnitDescription, h.Email,
    s.Direct_Manager AS DirectManager,
    COALESCE(CAST(t.ClassesRequired AS varchar(20)),
             CAST(tr.ClassesRequired AS varchar(20)), s.Required_Cirrculums)    AS RequiredCurricula,
    COALESCE(CAST(t.ClassesRegistered AS varchar(20)),
             CAST(tr.ClassesRegistered AS varchar(20)), s.Registered_Curriculums) AS RegisteredCurricula,
    COALESCE(CAST(t.ClassesCompleted AS varchar(20)),
             CAST(tr.ClassesCompleted AS varchar(20)), s.Completed_Curriculums)  AS CompletedCurricula,
    COALESCE(t.TrainingStatus,
             CASE WHEN tr.u IS NULL          THEN 'No Epic Curriculum'
                  WHEN tr.FT = 1             THEN 'Fully Trained'
                  WHEN tr.ClassesCompleted>0 THEN 'In Progress'
                  ELSE 'Not Started' END) AS TrainingStatus,
    -- Forward-looking while they are still training, factual once they finish:
    -- next scheduled session wins; else the last class actually sat (Fully
    -- Trained only); else the date the equivalency credit was recorded. NULL
    -- when an equivalency person has no date anywhere in Cornerstone, in which
    -- case EstCompletionNote carries the 'Completed (Equivalent)' label.
    CAST(COALESCE(COALESCE(t.FinalScheduledDate, tr.FinalScheduledDate),
                  CASE WHEN COALESCE(t.FullyTrainedYN,
                                     CASE WHEN tr.FT = 1 THEN 'Yes' END) = 'Yes'
                       THEN c.LastCompletedClass END,
                  CASE WHEN COALESCE(t.FinalScheduledDate, tr.FinalScheduledDate) IS NULL
                            AND (COALESCE(q.HasEquiv, tr.HasEquiv) = 1
                                 OR COALESCE(t.FullyTrainedYN, CASE WHEN tr.FT = 1 THEN 'Yes' END) = 'Yes')
                       THEN COALESCE(c.EquivCompleted, eq.EpicEquivDate) END,
                  c.LastSession) AS date) AS EstCompletionDate,
    CASE WHEN COALESCE(COALESCE(t.FinalScheduledDate, tr.FinalScheduledDate),
                       CASE WHEN COALESCE(t.FullyTrainedYN,
                                          CASE WHEN tr.FT = 1 THEN 'Yes' END) = 'Yes'
                            THEN c.LastCompletedClass END,
                       COALESCE(c.EquivCompleted, eq.EpicEquivDate), c.LastSession) IS NULL
               AND COALESCE(q.HasEquiv, tr.HasEquiv) = 1
              THEN 'Completed (Equivalent)' END AS EstCompletionNote,
    -- Epic rollup, never the summary export: its Fully Registered / Fully
    -- Trained describe whatever curriculum the person sits on, which for the
    -- SVC group is Non-Provider (Main Pool), not Wave 3 RCM training.
    COALESCE(t.FullyRegisteredYN, CASE WHEN tr.u IS NULL THEN NULL
                                       WHEN tr.FR = 1 THEN 'Yes' ELSE 'No' END) AS FullyRegistered,
    COALESCE(t.FullyTrainedYN,    CASE WHEN tr.u IS NULL THEN NULL
                                       WHEN tr.FT = 1 THEN 'Yes' ELSE 'No' END) AS FullyTrained,
    ISNULL(COALESCE(t.FullyTrainedYN, CASE WHEN tr.u IS NULL THEN NULL
                                           WHEN tr.FT = 1 THEN 'Yes' ELSE 'No' END), '') AS TrainingComplete,
    -- system-of-record values only; the workbook's typed answers live there.
    -- Yes if EITHER the Master or a wave change request says Yes (the Master
    -- defaults to No for anyone not designated, so first-non-empty would let
    -- that default mask a later Yes on a change request — same rule as
    -- build_lava_list.pick()); otherwise No. NEVER blank — her rule 2026-09-09:
    -- "those columns should not be blank. if there is no info for that, then put no".
    CASE WHEN LEFT(LOWER(LTRIM(RTRIM(ISNULL(u.FECParticipant,'')))),1) = 'y'
           OR LEFT(LOWER(LTRIM(RTRIM(ISNULL(w.WcrFEC,'')))),1) = 'y' THEN 'Yes'
         ELSE 'No' END AS FECAccess,
    CASE WHEN LEFT(LOWER(LTRIM(RTRIM(ISNULL(u.SoftLiveParticipant,'')))),1) = 'y'
           OR LEFT(LOWER(LTRIM(RTRIM(ISNULL(w.WcrSoft,'')))),1) = 'y' THEN 'Yes'
         ELSE 'No' END AS SoftLiveAccess,
    u.IsInScope, u.OnLeaderList,
    -- Larry's rule (2026-09-15): only the CURRENT population counts. Status
    -- from HR (bi-weekly) or Epic's Team Member Lookup (daily), whichever
    -- shows leave / termination; the wave test is applied at the call site
    -- (Wave = target wave, or OnWave3DNFBList = 'Yes').
    h.AssignmentStatus AS HRStatus,
    el.EpicHRStatus,
    CASE WHEN dn.UniversalID IS NOT NULL THEN 'Yes' ELSE 'No' END AS OnWave3DNFBList,
    CASE WHEN LOWER(LTRIM(RTRIM(ISNULL(h.AssignmentStatus,'')))) IN
              ('pay leave','leave and absent','terminated','severance','unaccounted for (zombie records)')
           OR LOWER(LTRIM(RTRIM(ISNULL(el.EpicHRStatus,'')))) IN
              ('pay leave','leave and absent','terminated','severance','unaccounted for (zombie records)')
         THEN 0 ELSE 1 END AS IsCurrent
INTO raw.lava_person
FROM #pop p
LEFT JOIN report.users u              ON u.UniversalID = p.u
LEFT JOIN raw.[master] m              ON m.UniversalID = p.u
LEFT JOIN raw.mvp_person v            ON v.UniversalID = p.u
LEFT JOIN report.hr h                 ON h.UniversalID = p.u
LEFT JOIN raw.epic_person e           ON e.UniversalID = p.u
LEFT JOIN raw.cornerstone_person c    ON c.UniversalID = p.u
LEFT JOIN #rf rf                      ON rf.u = p.u
LEFT JOIN report.tracker_training_status t
       ON t.UniversalID = p.u AND t.OnLeaderList = 1
LEFT JOIN #tr tr ON tr.u = p.u
LEFT JOIN (SELECT UPPER(LTRIM(RTRIM(Universal_Id))) AS u,
                  MAX(Team_Member) AS Team_Member, MAX(Curriculum_type) AS Curriculum_type,
                  MAX(Department_Unit) AS Department_Unit, MAX(Direct_Manager) AS Direct_Manager,
                  MAX(Required_Cirrculums) AS Required_Cirrculums,
                  MAX(Registered_Curriculums) AS Registered_Curriculums,
                  MAX(Completed_Curriculums) AS Completed_Curriculums,
                  MAX(Fully_Registered) AS Fully_Registered,
                  MAX(Fully_Trained) AS Fully_Trained
             FROM raw.epic_status_summary
            WHERE LTRIM(RTRIM(ISNULL(Universal_Id,''))) <> ''
            GROUP BY UPPER(LTRIM(RTRIM(Universal_Id)))) s ON s.u = p.u
LEFT JOIN (SELECT UniversalID AS u,
                  MAX(CASE WHEN Event_Class_Status LIKE '%Equivalent%' THEN 1 ELSE 0 END) AS HasEquiv
             FROM report.tracker_detail
            WHERE Event_Class_Type = 'Session' AND IsExcluded = 0
              AND (Event_Class_Registered LIKE '%Yes%' OR Event_Class_Status LIKE 'Completed%')
            GROUP BY UniversalID) q ON q.u = p.u
LEFT JOIN #el el ON el.u = p.u
LEFT JOIN #dn dn ON dn.UniversalID = p.u
LEFT JOIN (SELECT UPPER(LTRIM(RTRIM(Universal_Id))) AS u,
                  MAX(TRY_CONVERT(datetime2, Registration_Date)) AS EpicEquivDate
             FROM raw.epic_status
            WHERE _snapshot_date = (SELECT MAX(_snapshot_date) FROM raw.epic_status)
              AND Event_Class_Status = 'Completed (Equivalent)'
              AND LTRIM(RTRIM(ISNULL(Universal_Id,''))) <> ''
            GROUP BY UPPER(LTRIM(RTRIM(Universal_Id)))) eq ON eq.u = p.u
LEFT JOIN (SELECT UPPER(LTRIM(RTRIM(Universal_ID))) AS u,
                  MAX(FEC_Participant_Yes_No) AS WcrFEC,
                  MAX(Soft_Live_Participant_Y_N) AS WcrSoft
             FROM raw.wave_change_requests WHERE Universal_ID IS NOT NULL
            GROUP BY UPPER(LTRIM(RTRIM(Universal_ID)))) w ON w.u = p.u
OUTER APPLY (SELECT ' ' + REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(
                    LOWER(ISNULL(h.JobTitle,'')),
                    ',',' '),'.',' '),'/',' '),'-',' '),'(',' '),')',' ') + ' ' AS TitlePad) t2
"""


def build_lava_list(engine):
    """raw.lava_person — the LAVA Wave census (working LAVA list), assembled in stages.

    Why this is Python-staged instead of one view: as a single view the same
    logic NEVER RETURNED. Every join is individually fast (worst 1.7s, ~6s all
    told) but SQL Server could not plan 12 joins hanging off a UNION-derived
    driving set — the plan came back with 27 table scans and 27 sorts, and it
    would not prune the LEFT JOINs, so projecting even one column cost 41s.
    Materialising the driving set first and joining against it takes 8s.

    report.lava_list is a thin view over the table this writes, so the query
    surface is unchanged. Run after the report.* views exist:
        python sql
efresh.py lava

    Population = wave-training people UNION SVC people, where SVC means an MVP
    "Simple Visit Coding" job role AND presence in the Epic Curriculum Status by
    User Summary export (per the analyst 2026-08-26 — NOT the stale ad_hoc SVC file).
    """
    with engine.begin() as con:
        con.execute(text("DROP TABLE IF EXISTS #pop"))
        con.execute(text("DROP TABLE IF EXISTS #rf"))
        con.execute(text("DROP TABLE IF EXISTS #tr"))
        # stage 1 — per-person job-role flags off the indexed MVP table
        con.execute(text("""
            SELECT m.UniversalID AS u,
                   MAX(f.IsGuest) AS IsGuest,
                   MAX(f.IsSVC)   AS IsSVC,
                   MAX(CASE WHEN f.IsSVC = 1 THEN LTRIM(RTRIM(r.n)) END) AS SVCJobRole,
                   CAST(CASE WHEN MAX(f.IsSVC) = 1 THEN 'MVP' END AS varchar(20)) AS SVCSource
            INTO #rf
            FROM raw.mvp_person m
            CROSS APPLY (VALUES (m.JobRole1),(m.JobRole2),(m.JobRole3),(m.JobRole4)) r(n)
            JOIN raw.job_role_flags f ON f.RoleName = UPPER(LTRIM(RTRIM(r.n)))
            GROUP BY m.UniversalID"""))
        con.execute(text("CREATE UNIQUE CLUSTERED INDEX IX_rf ON #rf(u)"))
        # stage 1b — SVC roles assigned on a Wave Change Request Form that MVP
        # does not carry yet (her ask 2026-09-15: "management need them on
        # reports today"). Same rule as build_lava_list.py: the form's role
        # counts until MVP shows it; only IDs MVP knows (a mistyped ID on a
        # form must not become a nameless row). SVCSource says which.
        con.execute(text("DROP TABLE IF EXISTS #wcr"))
        con.execute(text("""
            SELECT UPPER(LTRIM(RTRIM(w.Universal_ID))) AS u,
                   MIN(LTRIM(RTRIM(r.n))) AS SVCJobRole
            INTO #wcr
            FROM raw.wave_change_requests w
            CROSS APPLY (VALUES (w.Staff_Training_Job_Role_1),(w.Staff_Training_Job_Role_2),
                                (w.Staff_Training_Job_Role_3),(w.Staff_Training_Job_Role_4)) r(n)
            WHERE w.Universal_ID IS NOT NULL AND r.n LIKE '%Simple Visit Coding%'
              AND EXISTS (SELECT 1 FROM raw.mvp_person m
                           WHERE m.UniversalID = UPPER(LTRIM(RTRIM(w.Universal_ID))))
            GROUP BY UPPER(LTRIM(RTRIM(w.Universal_ID)))"""))
        con.execute(text("""
            UPDATE rf SET IsSVC = 1,
                          SVCJobRole = COALESCE(rf.SVCJobRole, w.SVCJobRole),
                          SVCSource  = CASE WHEN rf.IsSVC = 1 THEN 'MVP' ELSE 'Change Request' END
            FROM #rf rf JOIN #wcr w ON w.u = rf.u"""))
        con.execute(text("""
            INSERT INTO #rf (u, IsGuest, IsSVC, SVCJobRole, SVCSource)
            SELECT w.u, 0, 1, w.SVCJobRole, 'Change Request'
            FROM #wcr w WHERE NOT EXISTS (SELECT 1 FROM #rf rf WHERE rf.u = w.u)"""))
        # stage 2 — the driving population, materialised so the optimizer has stats
        con.execute(text("""
            SELECT u, MAX(InWave) AS InWave, MAX(InSVC) AS InSVC
            INTO #pop
            FROM (
                SELECT UniversalID AS u, 1 AS InWave, 0 AS InSVC
                FROM report.users WHERE TrainingNeeded = 'Yes'
                UNION ALL
                SELECT rf.u, 0, 1
                FROM #rf rf
                JOIN (SELECT DISTINCT UPPER(LTRIM(RTRIM(Universal_Id))) AS u
                        FROM raw.epic_status_summary
                       WHERE LTRIM(RTRIM(ISNULL(Universal_Id,''))) <> '') es ON es.u = rf.u
                WHERE rf.IsSVC = 1
            ) z GROUP BY u"""))
        con.execute(text("CREATE UNIQUE CLUSTERED INDEX IX_pop ON #pop(u)"))
        # stage 3 — training status for EVERYONE in the population, computed
        # straight from raw.epic_status.
        #
        # Why this exists (her call 2026-08-26): report.tracker_training_status
        # only covers wave-file people — it joins report.users WHERE IsInScope=1
        # and Leader on the canonical list — so the SVC-provisioning group came
        # back empty and the columns fell through to the Epic *summary* export,
        # whose Fully Registered / Fully Trained describe whatever curriculum
        # that person is on (Non-Provider Main Pool for all 346 of them), NOT
        # Wave 3 RCM training. That read as "97% of SVC already registered",
        # which is wrong. A handful of SVC people don't report to Dr Lastname03 but
        # we still track them, so they get handled like anyone else: if Epic or
        # Cornerstone has a registration for them, it gets filled in.
        #
        # This is the SAME aggregation report.tracker_training_status runs, over
        # the SAME source (raw.epic_status), including the tracker's IsExcluded
        # rule (Advanced Reporting / Charge Capture are optional workshops). The
        # only difference is it is not gated on wave-file membership. Tracker
        # values still win in the assembly; this fills in everyone the tracker
        # cannot see, so the two can never disagree for wave people.
        con.execute(text("""
            SELECT r.UniversalID AS u, r.FT, r.FR, r.TrainingStatus,
                   r.FullyRegisteredYN, r.FullyTrainedYN, r.ClassesRequired,
                   r.ClassesCompleted, r.ClassesRegistered,
                   r.LastAttendedDate, r.FinalScheduledDate, r.HasEquiv
            INTO #tr
            FROM report.epic_training_rollup r
            WHERE EXISTS (SELECT 1 FROM #pop p WHERE p.u = r.UniversalID)"""))
        con.execute(text("CREATE UNIQUE CLUSTERED INDEX IX_tr ON #tr(u)"))
        # stage 3b — current-population lookups (Larry's rule 2026-09-15),
        # materialised: joining report.dnfb_workqueue_owners (an email-resolving
        # view over raw.mvp / raw.hr) inline into the assembly never returned.
        con.execute(text("DROP TABLE IF EXISTS #el"))
        con.execute(text("DROP TABLE IF EXISTS #dn"))
        con.execute(text("""
            SELECT UPPER(LTRIM(RTRIM(Universal_ID))) AS u, MAX(Hr_Status_PDM) AS EpicHRStatus
            INTO #el
            FROM raw.epic_lookup WHERE LTRIM(RTRIM(ISNULL(Universal_ID,''))) <> ''
            GROUP BY UPPER(LTRIM(RTRIM(Universal_ID)))"""))
        con.execute(text("CREATE UNIQUE CLUSTERED INDEX IX_el ON #el(u)"))
        con.execute(text("""
            SELECT DISTINCT UniversalID INTO #dn
            FROM report.dnfb_workqueue_owners WHERE OnWave3DNFBList = 'Yes'"""))
        con.execute(text("CREATE UNIQUE CLUSTERED INDEX IX_dn ON #dn(UniversalID)"))
        # stage 4 — assemble against the materialised driving set
        con.execute(text("DROP TABLE IF EXISTS raw.lava_person"))
        con.execute(text(LAVA_ASSEMBLE_SQL))
        con.execute(text(
            "CREATE UNIQUE CLUSTERED INDEX IX_lava_person ON raw.lava_person(UniversalID)"))
    with engine.connect() as con:
        n = con.execute(text("SELECT COUNT(*) FROM raw.lava_person")).scalar()
    print(f"LAVA     -> raw.lava_person: {n:,} rows")


def build_epic_person(engine):
    """raw.epic_status -> raw.epic_person: one indexed row per UniversalID.

    The curriculum-status detail export is one row per curriculum (51k rows,
    9.7k people). Views that only want a person-level fact off it — the Epic job
    category, say — were each doing their own GROUP BY. See build_mvp_person for
    why re-deriving that inside an inlined CTE is so expensive."""
    with engine.begin() as con:
        con.execute(text("DROP TABLE IF EXISTS raw.epic_person"))
        con.execute(text("""
            SELECT CAST(UPPER(LTRIM(RTRIM([Universal_Id]))) AS NVARCHAR(450)) AS UniversalID,
                   MAX([Epic_Job_Category]) AS EpicJobCategory
            INTO raw.epic_person
            FROM raw.epic_status
            WHERE LTRIM(RTRIM(ISNULL([Universal_Id],''))) <> ''
            GROUP BY CAST(UPPER(LTRIM(RTRIM([Universal_Id]))) AS NVARCHAR(450))"""))
        con.execute(text(
            "CREATE UNIQUE CLUSTERED INDEX IX_epic_person ON raw.epic_person(UniversalID)"))


def build_cornerstone_person(engine):
    """raw.cornerstone -> raw.cornerstone_person: one indexed row per person with
    the Epic-provider training dates every wave report asks for.

    raw.cornerstone is the biggest table in the warehouse (800k+ rows) and this
    same GROUP BY was being written out inside view after view. Column meanings:
      LastSession        last Epic training start date of any kind
      LastCompletedClass last class they ACTUALLY SAT — only real sessions carry
                         a Training_Start_Date, so requiring it keeps online
                         guides and assessments from masquerading as the final
                         class. Epic never dates an equivalency row, so this is
                         the only place a date exists for equivalency people.
      EquivCompleted     the date the equivalency credit itself was recorded,
                         for people who never sat a class.
    """
    with engine.begin() as con:
        con.execute(text("DROP TABLE IF EXISTS raw.cornerstone_person"))
        con.execute(text("""
            SELECT CAST(UPPER(LTRIM(RTRIM([User_ID]))) AS NVARCHAR(450)) AS UniversalID,
                   MAX([Training_Start_Date]) AS LastSession,
                   MAX(CASE WHEN [Transcript_Status] LIKE 'Completed%'
                             AND [Training_Start_Date] IS NOT NULL
                            THEN [Training_Start_Date] END) AS LastCompletedClass,
                   COALESCE(
                     MAX(CASE WHEN [Transcript_Status] LIKE '%Equivalent%'
                              THEN [Transcript_Completed_Date] END),
                     MAX(CASE WHEN [Transcript_Status] LIKE 'Completed%'
                              THEN [Transcript_Completed_Date] END)
                   ) AS EquivCompleted
            INTO raw.cornerstone_person
            FROM raw.cornerstone
            WHERE [Training_Provider] = 'EPIC'
              AND LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
            GROUP BY CAST(UPPER(LTRIM(RTRIM([User_ID]))) AS NVARCHAR(450))"""))
        con.execute(text(
            "CREATE UNIQUE CLUSTERED INDEX IX_cornerstone_person "
            "ON raw.cornerstone_person(UniversalID)"))


def build_mvp_person(engine):
    """raw.mvp -> raw.mvp_person: ONE indexed row per UniversalID.

    raw.mvp is 166k rows and every view that wants per-person MVP data was
    re-deriving it with ROW_NUMBER() ... ORDER BY LastImportedDate inside a CTE.
    CTEs are inlined, not materialised, so a view referencing that CTE twice
    re-scanned AND re-sorted all 166k rows twice — report.lava_list's first plan
    had 27 table scans and 27 sorts and never finished. Same reasoning as
    build_uids(): give the optimizer a small indexed table to seek.
    """
    with engine.begin() as con:
        con.execute(text("DROP TABLE IF EXISTS raw.mvp_person"))
        con.execute(text("""
            SELECT UniversalID, GoLiveWave, JobRole1, JobRole2, JobRole3, JobRole4,
                   BusinessUnitDescription, IsRCMUser
            INTO raw.mvp_person
            FROM (
                SELECT CAST(UPPER(LTRIM(RTRIM([UniversalID]))) AS NVARCHAR(450)) AS UniversalID,
                       [GoLiveWave],
                       [IndividualCategoryUpdate1Name] AS JobRole1,
                       [IndividualCategoryUpdate2Name] AS JobRole2,
                       [IndividualCategoryUpdate3Name] AS JobRole3,
                       [IndividualCategoryUpdate4Name] AS JobRole4,
                       [BusinessUnitDescription], [IsRCMUser],
                       ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM([UniversalID])))
                                          ORDER BY [LastImportedDate] DESC) AS rn
                FROM raw.mvp
                WHERE LTRIM(RTRIM(ISNULL([UniversalID],''))) <> ''
            ) z WHERE rn = 1"""))
        con.execute(text(
            "CREATE UNIQUE CLUSTERED INDEX IX_mvp_person ON raw.mvp_person(UniversalID)"))


def build_job_role_flags(engine):
    """raw.mvp_job_categories -> raw.job_role_flags: one indexed row per job-role
    name with the two flags every wave report asks for — is it a Guest House
    role, is it Simple Visit Coding. Small, but it was being GROUP BY-ed inside
    an inlined CTE and re-evaluated per row of a 664k-row CROSS APPLY."""
    with engine.begin() as con:
        con.execute(text("DROP TABLE IF EXISTS raw.job_role_flags"))
        con.execute(text("""
            SELECT CAST(UPPER(LTRIM(RTRIM([JobCategoryName]))) AS NVARCHAR(450)) AS RoleName,
                   MAX(CASE WHEN LOWER(LTRIM(RTRIM(ISNULL([IsGuestHouse],'')))) = 'true'
                            THEN 1 ELSE 0 END) AS IsGuest,
                   MAX(CASE WHEN LOWER([JobCategoryName]) LIKE '%simple visit coding%'
                            THEN 1 ELSE 0 END) AS IsSVC
            INTO raw.job_role_flags
            FROM raw.mvp_job_categories
            WHERE LTRIM(RTRIM(ISNULL([JobCategoryName],''))) <> ''
            GROUP BY CAST(UPPER(LTRIM(RTRIM([JobCategoryName]))) AS NVARCHAR(450))"""))
        con.execute(text(
            "CREATE UNIQUE CLUSTERED INDEX IX_job_role_flags ON raw.job_role_flags(RoleName)"))


def _newest(paths):
    """Newest path by name (dated/timestamped names sort chronologically), or None."""
    paths = sorted(paths)
    return paths[-1] if paths else None


def _newest_by_mtime(paths):
    """Newest path by modified time, skipping Excel lock files (~$...), or None.
    Use this for exports whose filenames don't sort chronologically (e.g.
    Cornerstone's 12-hour AM/PM timestamps)."""
    paths = [p for p in paths if not p.name.startswith("~$")]
    return max(paths, key=lambda p: p.stat().st_mtime, default=None)


def load_master(engine):
    """Master DATA sheet. If the live file is open in Excel (locked), fall back to
    the newest DATA_*.xlsx snapshot so the refresh still runs."""
    path, note = op.MASTER_WAVE_PATH, ""
    try:
        df = pd.read_excel(path, sheet_name="DATA", dtype=str, engine="calamine")
    except Exception:  # locked / corrupt / mid-sync — calamine + os errors vary
        snap = _newest(op.WAVE_REPOSITORY_DIR.glob("*/DATA_*.xlsx"))
        if snap is None:
            raise
        path, note = snap, "  (live master locked; used snapshot)"
        df = pd.read_excel(path, sheet_name="DATA", dtype=str, engine="calamine")
    print("MASTER  ", path.name, note)
    write_table(engine, "master", prepare(df, path.name))
    build_uids(engine, "master", "UniversalID")


def load_hr(engine):
    print("HR      ", op.RAW_HR_PATH.name)
    df = pd.read_excel(op.RAW_HR_PATH, sheet_name=0, dtype=str, engine="calamine")
    write_table(engine, "hr", prepare(df, op.RAW_HR_PATH.name))
    build_uids(engine, "hr", "UniversalID")


def load_hr_cleaned(engine):
    """Cleaned HR (resolved leader chain) — the 'All Staff' tab of the newest
    HR_cleaned_YYYY.MM.DD.xlsx produced by scripts/clean_hr.py."""
    path = op.latest_hr_cleaned()
    if path is None:
        print("HR_CLEAN (no HR_cleaned_*.xlsx found, skipped)")
        return
    print("HR_CLEAN", path.name)
    df = pd.read_excel(path, sheet_name="All Staff", dtype=str, engine="calamine")
    write_table(engine, "hr_cleaned", prepare(df, path.name))


def load_mvp(engine):
    print("MVP     ", op.RAW_MVP_USER_MAPPINGS.name)
    df = pd.read_csv(op.RAW_MVP_USER_MAPPINGS, dtype=str, encoding="utf-8-sig", low_memory=False)
    write_table(engine, "mvp", prepare(df, op.RAW_MVP_USER_MAPPINGS.name))
    build_mvp_person(engine)


def load_mvp_roles(engine):
    print("MVP_ROLE", op.RAW_MVP_ROLE_MAPPINGS.name)
    df = pd.read_csv(op.RAW_MVP_ROLE_MAPPINGS, dtype=str, encoding="utf-8-sig", low_memory=False)
    write_table(engine, "mvp_roles", prepare(df, op.RAW_MVP_ROLE_MAPPINGS.name))


def load_mvp_job_categories(engine):
    print("MVP_JOBC", op.RAW_MVP_JOB_CATEGORIES.name)
    df = pd.read_csv(op.RAW_MVP_JOB_CATEGORIES, dtype=str, encoding="utf-8-sig", low_memory=False)
    write_table(engine, "mvp_job_categories", prepare(df, op.RAW_MVP_JOB_CATEGORIES.name))
    build_job_role_flags(engine)


def load_cornerstone(engine):
    # Matches the fixed overwrite name (Enterprise_Training_Report.xlsx, since
    # 2026-08-04) and the older dated exports (Enterprise_Training_Report_<ts>.xlsx).
    path = _newest_by_mtime(op.RAW_CORNERSTONE_DIR.glob("Enterprise_Training_Report*.xlsx"))
    if path is None:
        print("CORNER   (no Enterprise_Training_Report*.xlsx found, skipped)")
        return
    print("CORNER  ", path.name)
    # Banner + filter block on top, and its height varies between exports —
    # locate the real header row ("User Full Name" in column A) instead of
    # hardcoding it.
    probe = pd.read_excel(path, sheet_name="Enterprise Training Report", header=None,
                          nrows=30, dtype=str, engine="calamine")
    hits = probe.index[probe[0].astype(str).str.strip() == "User Full Name"]
    if len(hits) == 0:
        raise ValueError(f"no 'User Full Name' header row in the first 30 rows of {path.name}")
    df = pd.read_excel(path, sheet_name="Enterprise Training Report", header=int(hits[0]),
                       dtype=str, engine="calamine")
    write_table(engine, "cornerstone", prepare(df, path.name))
    build_cornerstone_person(engine)


def load_epic_status(engine):
    """Curriculum Status Detail by User - W*.xlsx (one file per wave, no date in the
    name). Keep the newest file per wave by modified time; combine into raw.epic_status."""
    files = list(op.RAW_EPIC_STATUS_DIR.glob("Curriculum Status Detail by User - W*.xlsx"))
    if not files:
        print("EPIC_ST  (no Curriculum Status Detail files found, skipped)")
        return
    pat = re.compile(r"-\s*(W\d)", re.IGNORECASE)
    best: dict[str, Path] = {}
    for f in files:
        m = pat.search(f.name)
        if not m:
            continue
        wave = m.group(1).upper()
        if wave not in best or f.stat().st_mtime > best[wave].stat().st_mtime:
            best[wave] = f
    frames = []
    for wave, f in sorted(best.items()):
        snap = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d")
        print(f"EPIC_ST  {wave}  {f.name}")
        df = pd.read_excel(f, sheet_name="Export", dtype=str, engine="calamine")
        frames.append(prepare(df, f.name, extra={"_wave": wave, "_snapshot_date": snap}))
    write_table(engine, "epic_status", pd.concat(frames, ignore_index=True, sort=False))
    build_uids(engine, "epic_status", "Universal_Id")
    build_epic_person(engine)


def load_epic_status_summary(engine):
    """Curriculum Status by User Summary.xlsx — ONE row per person (the detail
    export is one row per curriculum). Loaded 2026-08-26 because it is the
    authoritative "who do we have Epic status data on" population: the retired
    ad_hoc "Curriculum Status by User - SVC.xlsx" was only this export
    pre-filtered, so SVC is derived as (MVP Simple Visit Coding job role) AND
    (present here) — see report.lava_list and build_soft_live_list.py."""
    path = op.RAW_EPIC_STATUS_SUMMARY
    if not path.exists():
        print("EPIC_SUM (no Curriculum Status by User Summary.xlsx found, skipped)")
        return
    print("EPIC_SUM", path.name)
    df = pd.read_excel(path, sheet_name="Export", dtype=str, engine="calamine")
    write_table(engine, "epic_status_summary", prepare(df, path.name))
    build_uids(engine, "epic_status_summary", "Universal_Id")


def load_epic_lookup(engine):
    path = op.RAW_EPIC_TM_DIR / "Epic Team Member Lookup.xlsx"
    print("EPIC_LK ", path.name)
    df = pd.read_excel(path, sheet_name="Export", dtype=str, engine="calamine")
    write_table(engine, "epic_lookup", prepare(df, path.name))


def load_epic_class_schedule(engine):
    """Newest class schedule export. Matches the fixed overwrite name
    (epic_class_schedule.xlsx, since 2026-08-04) and the older dated
    epic_class_schedule_*.xlsx copies (archive/ ignored by non-recursive glob)."""
    path = _newest_by_mtime(op.RAW_EPIC_CLASS_SCHEDULES_DIR.glob("epic_class_schedule*.xlsx"))
    if path is None:
        print("EPIC_CS  (no epic_class_schedule*.xlsx found, skipped)")
        return
    print("EPIC_CS ", path.name)
    df = pd.read_excel(path, sheet_name="Export", dtype=str, engine="calamine")
    write_table(engine, "epic_class_schedule", prepare(df, path.name))


def load_refs(engine):
    """Reference lists workbook (data/references/wave_reference_lists.xlsx) —
    every sheet loads as raw.ref_<sheetname>. Edit the workbook, rerun this."""
    path = op.REF_LISTS_PATH
    if not path.exists():
        print("REFS     (wave_reference_lists.xlsx not found, skipped)")
        return
    xl = pd.ExcelFile(path, engine="calamine")
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, dtype=str)
        table = "ref_" + clean_col(sheet).lower()
        write_table(engine, table, prepare(df, path.name))


# ----------------------------------------------------------------------
# per-source schemas (cornerstone / epic / mvp / hr / wave)
# ----------------------------------------------------------------------
# Every source system gets its own schema so its tables can be browsed —
# and joined — without hunting through raw.*. The alias views below are
# re-created on every run (so they always match the raw tables' columns);
# the analyst's own views in these schemas are left alone and are backed
# up daily alongside sandbox (see _backup_user_views).
SOURCE_SCHEMAS = ("cornerstone", "epic", "mvp", "hr", "wave")

# View names carry the data stage so raw exports and processed files can't be
# confused: raw_* = the export exactly as received (data\raw\...);
# processed_* = cleaned by the pipeline (data\processed\...); wave.master_data
# is the Master Wave File itself (data\reports\). Each alias also exposes a
# _stage column with the same value.
BUILTIN_SOURCE_ALIASES = [
    # (schema, view name, raw table, stage)
    ("cornerstone", "raw_enterprise_training",   "cornerstone",          "raw"),
    ("epic",        "raw_team_member_lookup",    "epic_lookup",          "raw"),
    ("epic",        "raw_curriculum_status",     "epic_status",          "raw"),
    ("epic",        "raw_curriculum_status_summary", "epic_status_summary", "raw"),
    ("epic",        "raw_class_schedule",        "epic_class_schedule",  "raw"),
    ("mvp",         "raw_user_mappings",         "mvp",                  "raw"),
    ("mvp",         "raw_role_mappings",         "mvp_roles",            "raw"),
    ("mvp",         "raw_job_categories",        "mvp_job_categories",   "raw"),
    ("hr",          "raw_workday_export",        "hr",                   "raw"),
    ("hr",          "processed_cleaned",         "hr_cleaned",           "processed"),
    ("wave",        "master_data",               "master",               "reports"),
    ("wave",        "raw_change_requests",       "wave_change_requests", "raw"),
]


def _ensure_schema(con, schema: str):
    con.execute(text(f"IF SCHEMA_ID('{schema}') IS NULL EXEC('CREATE SCHEMA [{schema}]')"))


def _probe_header_row(path, sheet: str) -> int:
    """0-based header row for exports whose header position moves. Cornerstone
    files carry a variable-length title/criteria block above the header (it
    grows or shrinks when the report's saved filters change), so a fixed
    HeaderRow breaks on the next filter edit — registry rows can say 'auto'
    instead. The header is the first row with 5+ non-empty cells; the
    metadata/criteria rows above it never have more than 2."""
    probe = pd.read_excel(path, sheet_name=sheet if sheet else 0, header=None,
                          nrows=30, engine="calamine")
    for i, row in probe.iterrows():
        if row.notna().sum() >= 5:
            return int(i)
    raise ValueError(f"no header row found in the first 30 rows of {path.name}")


def load_sources(engine):
    """Per-source schemas: (re)create the builtin alias views, then process the
    source registry (data/references/source_registry.xlsx) — each Enabled row
    loads the newest matching file into raw.<source>_<table> and creates the
    <source>.<table> view. Adding a new export needs NO code: drop the file in
    its folder and add one registry row."""
    with engine.begin() as con:
        for schema in SOURCE_SCHEMAS:
            _ensure_schema(con, schema)
        for schema, view, raw_table, stage in BUILTIN_SOURCE_ALIASES:
            if con.execute(text(f"SELECT OBJECT_ID('raw.[{raw_table}]')")).scalar():
                con.execute(text(
                    f"CREATE OR ALTER VIEW [{schema}].[{view}] AS "
                    f"/* {schema}.{view} — alias over raw.{raw_table} "
                    f"(stage: {stage}); auto-created by refresh.py, do not edit */ "
                    f"SELECT *, CONVERT(varchar(9), '{stage}') AS _stage "
                    f"FROM raw.[{raw_table}]"))
    print(f"SOURCES  {len(BUILTIN_SOURCE_ALIASES)} builtin alias views refreshed "
          f"({', '.join(SOURCE_SCHEMAS)})")

    path = op.SOURCE_REGISTRY_PATH
    if not path.exists():
        print("SOURCES  (source_registry.xlsx not found, registry skipped)")
        return
    reg = pd.read_excel(path, sheet_name=0, dtype=str).fillna("")
    for _, row in reg.iterrows():
        if str(row.get("Enabled", "")).strip().lower() != "yes":
            continue
        source = clean_col(str(row["Source"]).strip()).lower()
        table = clean_col(str(row["Table"]).strip()).lower()
        folder = op.ONEDRIVE_ROOT / "data" / str(row["Folder"]).strip().strip("\\/")
        pattern = str(row.get("FilePattern", "")).strip() or "*.xlsx"
        files = [p for p in folder.glob(pattern) if not p.name.startswith("~$")]
        if not files:
            print(f"SOURCES  {source}.{table}: nothing matches '{pattern}' in {folder} — skipped")
            continue
        newest = max(files, key=lambda p: p.stat().st_mtime)
        sheet = str(row.get("Sheet", "")).strip()
        hdr_txt = str(row.get("HeaderRow", "")).strip()
        if hdr_txt.lower() == "auto":
            hdr = _probe_header_row(newest, sheet)
        else:
            hdr = (int(float(hdr_txt)) - 1) if hdr_txt else 0
        if newest.suffix.lower() == ".csv":
            df = pd.read_csv(newest, dtype=str, encoding="utf-8-sig",
                             low_memory=False, header=hdr)
        else:
            df = pd.read_excel(newest, sheet_name=sheet if sheet else 0,
                               dtype=str, engine="calamine", header=hdr)
        # Stage comes from the first folder segment (raw\... or processed\...),
        # and prefixes the view name so the stage is visible at a glance.
        stage = str(row["Folder"]).strip().replace("/", "\\").split("\\")[0].lower() or "raw"
        view = f"{stage}_{table}"
        raw_table = f"{source}_{table}"
        print(f"SOURCES  {newest.name}")
        write_table(engine, raw_table, prepare(df, newest.name))
        with engine.begin() as con:
            _ensure_schema(con, source)
            con.execute(text(
                f"CREATE OR ALTER VIEW [{source}].[{view}] AS "
                f"/* {source}.{view} — registry export over raw.{raw_table} "
                f"(stage: {stage}); auto-created by refresh.py from "
                f"source_registry.xlsx, do not edit */ "
                f"SELECT *, CONVERT(varchar(9), '{stage}') AS _stage "
                f"FROM raw.[{raw_table}]"))
        print(f"         -> raw.{raw_table} + view {source}.{view}")


def load_runlogs(engine):
    """Run-history workbooks -> raw.activity_log / raw.metrics_summary, so run
    status and metric trends are queryable (and visible in Power BI)."""
    for table, path in [
        ("activity_log", op.RUNLOGS_DIR / "activity_log.xlsx"),
        ("metrics_summary", op.RUNLOGS_DIR / "metrics_summary.xlsx"),
    ]:
        if not path.exists():
            print(f"RUNLOGS  ({path.name} not found, skipped)")
            continue
        print("RUNLOGS ", path.name)
        df = pd.read_excel(path, sheet_name=0, dtype=str, engine="calamine")
        write_table(engine, table, prepare(df, path.name))


def load_wave_change_requests(engine):
    """ALL 'Wave Change Request Form*.xlsx' workbooks in the folder, combined —
    only their 'Wave Change Request' data tab (the Instructions / quick-guide /
    ref tabs are for the requesters). One form per requester accumulates here;
    archive a form (move to archive\\<YYYY-MM>\\) once its requests are handled
    and it drops out on the next load. The strict filename pattern keeps stray
    workbooks dropped in the folder from being read as request forms."""
    files = sorted(p for p in op.RAW_WAVE_CHANGE_REQUESTS_DIR
                   .glob("Wave Change Request Form*.xlsx")
                   if not p.name.startswith("~$"))
    if not files:
        print("WAVE_CR  (no Wave Change Request Form*.xlsx found, skipped)")
        return
    frames = []
    for f in files:
        print("WAVE_CR ", f.name)
        df = pd.read_excel(f, sheet_name="Wave Change Request", dtype=str, engine="calamine")
        frames.append(prepare(df, f.name))
    write_table(engine, "wave_change_requests",
                pd.concat(frames, ignore_index=True, sort=False))


# ----------------------------------------------------------------------
# daily history snapshot (for trend / day-over-day comparison queries)
# ----------------------------------------------------------------------
# history table -> the query whose result is today's state worth keeping.
# roster_daily keeps the PERSON-LEVEL state (not aggregates) so history can
# be sliced by leader, vendor, wave, or any future category after the fact.
HISTORY_TABLES = {
    "roster_daily": "SELECT * FROM report.roster",
    # Daily record of Epic team members with no HR row — lastname13 the
    # FirstSeen/NewToday flags on the boss's Epic-not-in-HR gap report.
    "epic_not_in_hr_daily": "SELECT * FROM report.epic_not_in_hr",
    # Wave 3 per-leader daily totals (registered/trained, unregistered,
    # no-shows standing & to date) — lastname13 report.w3_registration_summary
    # and the W3 registration look-back CSVs. Added 2026-07-15.
    "w3_leader_daily": "SELECT * FROM report.w3_leader_snapshot",
}


def _ensure_history_schema(engine):
    with engine.begin() as con:
        con.execute(text("IF SCHEMA_ID('history') IS NULL EXEC('CREATE SCHEMA history')"))


def _sync_history_columns(engine, table: str, df: pd.DataFrame):
    """Add any columns df has that history.<table> lacks (as nullable NVARCHAR),
    so the roster can grow new columns without breaking the daily append."""
    with engine.begin() as con:
        existing = {r[0] for r in con.execute(text(
            f"SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('history.{table}')"))}
        if not existing:
            return  # table doesn't exist yet; to_sql will create it
        for col in df.columns:
            if col not in existing:
                con.execute(text(f"ALTER TABLE history.{table} ADD [{col}] NVARCHAR(4000) NULL"))
                print(f"HISTORY  history.{table}: added new column [{col}]")


def _backup_user_views(engine):
    """Script every hand-built view in the analyst's schemas (sandbox + the
    per-source schemas) to ONE .sql file on OneDrive (overwritten each run;
    OneDrive keeps versions). If the local DB is ever lost/rebuilt, running
    that file restores them all."""
    schemas = "', '".join(("sandbox",) + SOURCE_SCHEMAS)
    with engine.connect() as con:
        rows = con.execute(text(f"""
            SELECT SCHEMA_NAME(v.schema_id) AS sch, v.name, m.definition
            FROM sys.views v
            JOIN sys.sql_modules m ON m.object_id = v.object_id
            WHERE SCHEMA_NAME(v.schema_id) IN ('{schemas}')
            ORDER BY SCHEMA_NAME(v.schema_id), v.name""")).fetchall()
    if not rows:
        return
    defs = []
    for _, _, d in rows:
        # normalize to CREATE OR ALTER so the restore file is rerunnable
        defs.append(re.sub(r"(?i)CREATE\s+(OR\s+ALTER\s+)?VIEW",
                           "CREATE OR ALTER VIEW", d.strip(), count=1))
    path = op.REFERENCES_DIR / "user_views_backup.sql"
    path.parent.mkdir(parents=True, exist_ok=True)
    schema_ddl = "\n".join(
        f"IF SCHEMA_ID('{s}') IS NULL EXEC('CREATE SCHEMA [{s}]');"
        for s in ("sandbox",) + SOURCE_SCHEMAS)
    header = (
        "-- Hand-built view definitions (sandbox + source schemas) — auto-exported\n"
        "-- by sql\\refresh.py (snapshot step). Restore them all on a fresh DB with:\n"
        "--   sqlcmd -S \".\\SQLEXPRESS\" -E -C -d AnalyticsDB -i \"data\\references\\user_views_backup.sql\"\n"
        + schema_ddl + "\nGO\n\n"
    )
    path.write_text(header + "\nGO\n\n".join(defs) + "\nGO\n", encoding="utf-8")
    print(f"HISTORY  user views backed up: {len(rows)} -> {path.name}")


def snapshot(engine):
    """Append today's state to the history.* tables — one snapshot per calendar
    day; rerunning on the same day replaces that day's rows. Each day's rows are
    also mirrored to a dated CSV on OneDrive (history_snapshots/), which is the
    durable copy: the local DB can always be rebuilt from those via
    `refresh.py backload_history`. Also backs up hand-built sandbox.* views."""
    _ensure_history_schema(engine)
    today = f"{date.today():%Y-%m-%d}"
    for table, src in HISTORY_TABLES.items():
        df = pd.read_sql(src, engine)
        if df.empty:
            print(f"HISTORY  history.{table}: source empty, skipped")
            continue
        df.insert(0, "SnapshotDate", today)
        _sync_history_columns(engine, table, df)
        with engine.begin() as con:
            if con.execute(text(f"SELECT OBJECT_ID('history.{table}')")).scalar():
                con.execute(text(f"DELETE FROM history.{table} WHERE SnapshotDate = :d"),
                            {"d": today})
        df.to_sql(table, engine, schema="history", if_exists="append", index=False)
        csv_dir = op.month_subdir(op.HISTORY_SNAPSHOTS_DIR / table)
        csv_path = csv_dir / f"{table}_{date.today():%Y.%m.%d}.csv"
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        print(f"HISTORY  history.{table}: {len(df):,} rows stamped {today}  (CSV: {csv_path.name})")
    _backup_user_views(engine)


def backload_history(engine):
    """Rebuild the history.* tables from the dated CSVs on OneDrive — for a new
    machine or a lost/rebuilt local DB. Replaces the SQL tables entirely."""
    _ensure_history_schema(engine)
    for table in HISTORY_TABLES:
        files = sorted((op.HISTORY_SNAPSHOTS_DIR / table).glob(f"*/{table}_*.csv"))
        if not files:
            print(f"BACKLOAD history.{table}: no CSVs found, skipped")
            continue
        df = pd.concat([pd.read_csv(f, dtype=str, encoding="utf-8-sig") for f in files],
                       ignore_index=True, sort=False)
        df.to_sql(table, engine, schema="history", if_exists="replace", index=False)
        print(f"BACKLOAD history.{table}: {len(df):,} rows from {len(files)} daily CSVs")


LOADERS = {
    "master": load_master,
    "hr": load_hr,
    "hr_cleaned": load_hr_cleaned,
    "mvp": load_mvp,
    "mvp_roles": load_mvp_roles,
    "mvp_job_categories": load_mvp_job_categories,
    "cornerstone": load_cornerstone,
    "epic_status": load_epic_status,
    "epic_status_summary": load_epic_status_summary,
    "epic_lookup": load_epic_lookup,
    "epic_class_schedule": load_epic_class_schedule,
    "wave_change_requests": load_wave_change_requests,
    "refs": load_refs,
    "runlogs": load_runlogs,
    "sources": load_sources,  # per-source schemas + the Excel source registry
    "snapshot": snapshot,   # keep last: a full run stamps history after all loads
}

# backload_history is deliberately NOT part of a full run — only when named.
# build_lava_list is NOT in LOADERS: it reads report.* views, so it must run
# AFTER the view files are applied, not during the raw loads.
COMMANDS = {**LOADERS, "backload_history": backload_history, "lava": build_lava_list}


def refresh_view_metadata(engine) -> int:
    """Re-bind every non-schemabound view to its base tables.

    SQL Server freezes a `SELECT *` view's column list at CREATE time, so when a
    raw table changes shape — a column added to an export, or a Master column
    renamed — the alias views keep serving the OLD names and can hand back
    mismatched data. Caught 2026-08-11: wave.master_data still exposed the two
    pre-rename Cornerstone_* columns after the Master gained its training-status
    block, because the alias is only re-created when the `sources` loader runs.
    sp_refreshview is idempotent and costs milliseconds across ~100 views, so it
    runs after every load rather than only on a full refresh.
    """
    refreshed = 0
    with engine.begin() as con:
        views = [r[0] for r in con.execute(text(
            "SELECT QUOTENAME(TABLE_SCHEMA)+'.'+QUOTENAME(TABLE_NAME) "
            "FROM INFORMATION_SCHEMA.VIEWS")).fetchall()]
        for v in views:
            try:
                con.execute(text(f"EXEC sp_refreshview '{v}'"))
                refreshed += 1
            except Exception as e:      # a genuinely broken view shouldn't fail the load
                print(f"   !! sp_refreshview {v}: {str(e)[:100]}")
    return refreshed


# ----------------------------------------------------------------------
# main
# ----------------------------------------------------------------------
def main():
    requested = [a.lower() for a in sys.argv[1:]] or list(LOADERS)
    unknown = [r for r in requested if r not in COMMANDS]
    if unknown:
        print("Unknown source(s):", unknown, "\nValid:", list(COMMANDS))
        sys.exit(1)

    engine = create_engine(CONN, fast_executemany=True)
    start = datetime.now()
    print(f"Loading {requested} into {DATABASE} at {start:%H:%M:%S}\n")
    failed = []
    for name in requested:
        try:
            COMMANDS[name](engine)
        except Exception as e:
            failed.append(name)
            print(f"   !! {name} FAILED: {e}")
    n = refresh_view_metadata(engine)
    print(f"\nVIEWS    {n} view(s) re-bound to their base tables (sp_refreshview)")
    engine.dispose()
    print(f"Done in {(datetime.now() - start).total_seconds():.0f}s.")
    print("Next: run 3_report_views.sql once (or after editing views) to refresh report.*")
    if failed:
        print(f"FAILED sources: {', '.join(failed)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
