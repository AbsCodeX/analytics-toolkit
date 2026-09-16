-- =========================================================================
-- SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
-- Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
-- fill in your own values before running. See README.md for the full map.
-- =========================================================================
/* ============================================================
   3_report_views.sql  —  RUN ONCE (and again any time you edit a view)
   Builds the report.* layer: clean, joined views over raw.*.

   Join key everywhere: UniversalID, normalized = UPPER(TRIM(id)).

   Source authority (who "wins" for each field):
     Master Wave  -> who is in scope / identity / wave
     HR           -> org rollup (AVP / VP / SVP)
     MVP          -> job role
     Cornerstone  -> latest registration + training
     Epic         -> HR + Cornerstone combined (cross-check)

   How to run:
     sqlcmd -S ".\SQLEXPRESS" -E -C -d AnalyticsDB -i "sql\3_report_views.sql"
   ============================================================ */
USE AnalyticsDB;
GO

/* ---------- 1. report.users — the clean Master roster ----------
   Shows      : one row per Master person — identity, wave, leader chain,
                job roles, vendor/user type + IsInScope (1 = active & needs training)
                + OnLeaderList (1 = Leader is on the canonical raw.ref_leaders
                list; the published aggregates below require BOTH flags).
   Built from : raw.master + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.users AS
SELECT
    UPPER(LTRIM(RTRIM(UniversalID)))            AS UniversalID,
    [Full_Name]                                 AS FullName,
    [Leaders]                                   AS Leader,
    [AVP]                                        AS AVP,
    [VP]                                         AS VP,
    [SVP]                                        AS SVP,
    [GoLiveWave]                                AS Wave,
    [User_Type_Offshore_Onshore_YourOrg]     AS UserType,
    [Vendor_Yes_No]                            AS VendorYN,
    [Job_Role_1]                               AS JobRole1,
    [Job_Role_2]                               AS JobRole2,
    [Job_Role_3]                               AS JobRole3,
    [Job_Role_4]                               AS JobRole4,
    [Training_Needed_Yes_No]                   AS TrainingNeeded,
    [IsDepartedInactive]                       AS Departed,
    -- Never blank (her rule 2026-09-09): a Master cell nobody has filled in
    -- reads No. fill_master_participation_defaults.py writes the same default
    -- onto the Master itself every apply run; this covers the gap in between.
    ISNULL(NULLIF(LTRIM(RTRIM([FEC_Participant_Yes_No])),''),'No')    AS FECParticipant,
    ISNULL(NULLIF(LTRIM(RTRIM([Soft_Live_Participant_Y_N])),''),'No') AS SoftLiveParticipant,
    [Terminated]                               AS TerminatedFlag,
    CASE WHEN LTRIM(RTRIM(ISNULL([Leaders],''))) IN
              (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
         THEN 1 ELSE 0 END                      AS OnLeaderList,
    CASE
        WHEN UPPER(LTRIM(RTRIM([Training_Needed_Yes_No]))) = 'YES'
         AND ISNULL(UPPER(LTRIM(RTRIM([IsDepartedInactive]))),'') NOT IN ('YES','Y','TRUE','1')
        THEN 1 ELSE 0
    END                                         AS IsInScope
FROM raw.[master]
WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> '';
GO

/* ---------- 2. report.training_status — Epic training flags ----------
   Shows      : one row per user — IsFullyTrained / IsFullyRegistered (1/0)
                from Epic's Curriculum Status export.
                Wave-file people only (default scope for all report views).
   Built from : raw.epic_status + report.users
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.training_status AS
SELECT
    UPPER(LTRIM(RTRIM([Universal_Id])))                            AS UniversalID,
    MAX([_wave])                                                   AS Wave,
    -- Epic exports the flags as "✅ Yes" / "⛔ No" (emoji-prefixed), so match on
    -- LIKE '%YES%' rather than = 'YES'.
    MAX(CASE WHEN UPPER([Fully_Trained])    LIKE '%YES%' THEN 1 ELSE 0 END) AS IsFullyTrained,
    MAX(CASE WHEN UPPER([Fully_Registered]) LIKE '%YES%' THEN 1 ELSE 0 END) AS IsFullyRegistered,
    MAX([_snapshot_date])                                          AS SnapshotDate
FROM raw.epic_status
WHERE LTRIM(RTRIM(ISNULL([Universal_Id],''))) <> ''
  AND UPPER(LTRIM(RTRIM([Universal_Id]))) IN (SELECT UniversalID FROM report.users)
GROUP BY UPPER(LTRIM(RTRIM([Universal_Id])));
GO

/* ---------- 3. report.cornerstone_status — Cornerstone rollup ----------
   Shows      : one row per user — registered/completed SESSION counts
                + latest registration date.
                Since 2026-07-17 the Enterprise export carries full transcript
                history (every attempt, all statuses) plus Online Class / Test
                rows, so this rolls up the newest LIVE row per person x class
                (Withdrawn / Cancelled / Denied / Waitlist Expired ignored),
                Sessions only — same meaning as the old latest-only export.
                Training Provider = EPIC only; wave-file people only.
   Built from : raw.cornerstone + report.users
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.cornerstone_status AS
WITH latest AS (
    SELECT UniversalID, ClassKey, [Transcript_Status], RegDate
    FROM (
        SELECT UPPER(LTRIM(RTRIM([User_ID])))        AS UniversalID,
               UPPER(LTRIM(RTRIM([Training_Title]))) AS ClassKey,
               [Transcript_Status],
               TRY_CONVERT(datetime, [Transcript_Registration_Date]) AS RegDate,
               ROW_NUMBER() OVER (
                   PARTITION BY UPPER(LTRIM(RTRIM([User_ID]))), UPPER(LTRIM(RTRIM([Training_Title])))
                   ORDER BY TRY_CONVERT(datetime, [Transcript_Registration_Date]) DESC) AS rn
        FROM raw.cornerstone
        WHERE LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
          AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
          AND [Training_Type] = 'Session'
          AND [Transcript_Status] NOT IN ('Withdrawn','Cancelled','Denied','Waitlist Expired')
          AND UPPER(LTRIM(RTRIM([User_ID]))) IN (SELECT UniversalID FROM report.users)
    ) x WHERE rn = 1
)
SELECT
    UniversalID,
    COUNT(*)                                                                   AS TotalRecords,
    SUM(CASE WHEN UPPER([Transcript_Status]) = 'REGISTERED' THEN 1 ELSE 0 END) AS RegisteredCount,
    SUM(CASE WHEN UPPER([Transcript_Status]) = 'COMPLETED'  THEN 1 ELSE 0 END) AS CompletedCount,
    MAX(RegDate)                                                               AS LatestRegistrationDate
FROM latest
GROUP BY UniversalID;
GO

/* ---------- 4. report.kpi_summary — top-line KPI cards ----------
   Shows      : ONE row — in-scope headcount, registered, trained + %s.
                Scoped to the canonical leader list (OnLeaderList = 1),
                same population rule as every other published report.
   Built from : report.users + report.training_status
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.kpi_summary AS
WITH base AS (
    SELECT u.UniversalID,
           ISNULL(t.IsFullyTrained,0)    AS Trained,
           ISNULL(t.IsFullyRegistered,0) AS Registered
    FROM report.users u
    LEFT JOIN report.training_status t ON t.UniversalID = u.UniversalID
    WHERE u.IsInScope = 1
      AND u.OnLeaderList = 1
)
SELECT
    COUNT(*)                                                          AS InScopeUsers,
    SUM(Registered)                                                   AS RegisteredUsers,
    SUM(Trained)                                                      AS TrainedUsers,
    CAST(100.0 * SUM(Registered) / NULLIF(COUNT(*),0) AS decimal(5,1)) AS PctRegistered,
    CAST(100.0 * SUM(Trained)    / NULLIF(COUNT(*),0) AS decimal(5,1)) AS PctTrained
FROM base;
GO

/* ---------- 5. report.leader_summary — rollup by leader ----------
   Shows      : one row per leader — in-scope, registered, trained, % trained.
                Canonical leaders only (OnLeaderList = 1) — off-list tags like
                'No Longer Rev Cycle' and blank leaders never publish here.
   Built from : report.users + report.training_status
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.leader_summary AS
SELECT
    u.Leader                                                          AS Leader,
    COUNT(*)                                                          AS InScopeUsers,
    SUM(ISNULL(t.IsFullyRegistered,0))                               AS RegisteredUsers,
    SUM(ISNULL(t.IsFullyTrained,0))                                  AS TrainedUsers,
    CAST(100.0 * SUM(ISNULL(t.IsFullyTrained,0)) / NULLIF(COUNT(*),0) AS decimal(5,1)) AS PctTrained
FROM report.users u
LEFT JOIN report.training_status t ON t.UniversalID = u.UniversalID
WHERE u.IsInScope = 1
  AND u.OnLeaderList = 1
GROUP BY u.Leader;
GO

/* ---------- 6. report.exceptions — data-quality watchlist ----------
   Shows      : one row per person per issue, two checks (2026-07-13, per
                user: the TM LOOKUP is the authoritative Epic population —
                it captures everyone, even recently terminated — so ALL
                population reconciliation runs against it. The Curriculum
                Status Detail is training/registration METRICS for people
                cleared to train and is never used for population checks.)
     1. 'Epic Centralized not on wave file' — THE hard invariant: anyone
        on the Epic Team Member Lookup with Curriculum Type 'Revenue
        Cycle - Centralized' must be on the Master. apply's add_members
        step closes this daily, so any row here means something slipped.
     2. 'In scope but missing from HR' — active trainee whose UID is not
        in the HR export: org rollup impossible, usually a bad UID.
   Built from : raw.epic_lookup + report.users + raw.uids_hr + raw.uids_master
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.exceptions AS
-- 1. Centralized on the Epic TM Lookup but absent from the Master.
--    DUPLICATE LOGIC, DELIBERATELY KEPT (reviewed 2026-08-26): this anti-join
--    is the same population report.epic_missing_from_wave returns, presented
--    differently — this sheet is the must-be-zero alarm, Missing Wave - Epic
--    is the 14-column detail list. It is NOT collapsed into that view because
--    it lives in 4_dimensions.sql, which applies AFTER this file; depending on
--    it here would break a fresh database build. If you change the anti-join
--    below, change it there too.
SELECT UPPER(LTRIM(RTRIM(e.Universal_ID))) AS UniversalID,
       e.Team_Member                       AS FullName,
       e.Direct_Manager                    AS Leader,
       'Epic Centralized not on wave file' AS Issue
FROM raw.epic_lookup e
-- Seek the indexed raw.uids_master key table instead of a correlated NOT
-- EXISTS over the wide NVARCHAR(4000) raw.master heap, which the optimizer
-- resolved as a per-row nested loop (30s for 2 rows). Same rationale as
-- raw.uids_hr below; verified row-for-row identical output.
LEFT JOIN raw.uids_master mu
       ON mu.UniversalID = UPPER(LTRIM(RTRIM(e.Universal_ID)))
WHERE LTRIM(RTRIM(e.Curriculum_Type)) = 'Revenue Cycle - Centralized'
  AND LTRIM(RTRIM(ISNULL(e.Universal_ID,''))) <> ''
  AND mu.UniversalID IS NULL
UNION ALL
-- 2. In-scope users with no HR org record (org rollup will be blank)
SELECT u.UniversalID, u.FullName, u.Leader,
       'In scope but missing from HR' AS Issue
FROM report.users u
WHERE u.IsInScope = 1
  AND NOT EXISTS (SELECT 1 FROM raw.uids_hr h
                  WHERE h.UniversalID = u.UniversalID);
GO

/* ---------- 7. report.master_vs_sources — the update preview ----------
   Shows      : one row per in-scope user — the Master's current value next
                to the authoritative source value, with *_Differs flags
                (HR wins org rollup, MVP wins job role). READ-ONLY: proposes
                changes, never writes the file.
   Built from : report.users + raw.hr_cleaned + raw.mvp
                + report.training_status + report.cornerstone_status
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.master_vs_sources AS
WITH hr1 AS (         -- one HR row per user (org authority)
    SELECT UniversalID, AVP, VP, SVP FROM (
        SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
               [AVP], [VP], [SVP],
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM(UniversalID))) ORDER BY (SELECT NULL)) rn
        FROM raw.hr_cleaned
        WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
    ) x WHERE rn = 1
),
mvp1 AS (             -- one MVP row per user (job-role authority)
    SELECT UniversalID, JobRole1 FROM (
        SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
               [IndividualCategoryUpdate1Name] AS JobRole1,
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM(UniversalID)))
                                  ORDER BY [LastImportedDate] DESC) rn
        FROM raw.mvp
        WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
    ) x WHERE rn = 1
)
SELECT
    u.UniversalID, u.FullName,
    -- org rollup: HR wins
    u.AVP AS Master_AVP, hr1.AVP AS HR_AVP,
    CASE WHEN ISNULL(UPPER(LTRIM(RTRIM(u.AVP))),'') <> ISNULL(UPPER(LTRIM(RTRIM(hr1.AVP))),'') THEN 1 ELSE 0 END AS AVP_Differs,
    u.VP  AS Master_VP,  hr1.VP  AS HR_VP,
    CASE WHEN ISNULL(UPPER(LTRIM(RTRIM(u.VP))),'')  <> ISNULL(UPPER(LTRIM(RTRIM(hr1.VP))),'')  THEN 1 ELSE 0 END AS VP_Differs,
    -- job role: MVP wins
    u.JobRole1 AS Master_JobRole1, mvp1.JobRole1 AS MVP_JobRole1,
    CASE WHEN ISNULL(UPPER(LTRIM(RTRIM(u.JobRole1))),'') <> ISNULL(UPPER(LTRIM(RTRIM(mvp1.JobRole1))),'') THEN 1 ELSE 0 END AS JobRole_Differs,
    -- training: Epic + Cornerstone (informational)
    ISNULL(t.IsFullyTrained,0)    AS Epic_Trained,
    ISNULL(t.IsFullyRegistered,0) AS Epic_Registered,
    ISNULL(c.RegisteredCount,0)   AS Cornerstone_Registrations
FROM report.users u
LEFT JOIN hr1  ON hr1.UniversalID  = u.UniversalID
LEFT JOIN mvp1 ON mvp1.UniversalID = u.UniversalID
LEFT JOIN report.training_status   t ON t.UniversalID = u.UniversalID
LEFT JOIN report.cornerstone_status c ON c.UniversalID = u.UniversalID
WHERE u.IsInScope = 1;
GO

/* ---------- 8. report.hr — resolved leaders + org detail ----------
   Shows      : one row per HR person (ALL staff, not just the wave) —
                standardized leader chain + department/title/email detail
   Built from : raw.hr_cleaned (leader chain) + raw.hr (org detail)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.hr AS
WITH hc AS (          -- cleaned HR: resolved leader chain (one row per user)
    SELECT UniversalID, Full_Name, Leaders, SeniorManager, Director, SeniorDirector,
           AVP, VP, SVP, EVP, IsRCM, Departed FROM (
        SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
               [Full_Name], [Leaders], [SeniorManager], [Director], [SeniorDirector],
               [AVP], [VP], [SVP], [EVP], [IsRCM], [IsDeparted_Inactive] AS Departed,
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM(UniversalID))) ORDER BY (SELECT NULL)) rn
        FROM raw.hr_cleaned
        WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
    ) x WHERE rn = 1
),
hd AS (               -- raw HR: richer org detail (one row per user)
    SELECT UniversalID, Department_Name, JobTitle, WorkerType, Job_Function,
           Service_Line, BusinessUnitDesc, Email_Address, ManagerName,
           Assignment_Status FROM (
        SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
               [Department_Name], [JobTitle], [WorkerType], [Job_Function],
               [Service_Line], [BusinessUnitDesc], [Email_Address], [ManagerName],
               [Assignment_Status],
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM(UniversalID))) ORDER BY (SELECT NULL)) rn
        FROM raw.hr
        WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
    ) x WHERE rn = 1
)
SELECT hc.UniversalID, hc.Full_Name AS FullName, hc.Leaders AS Leader,
       hc.SeniorManager, hc.Director, hc.SeniorDirector, hc.AVP, hc.VP, hc.SVP, hc.EVP,
       hd.Department_Name AS Department, hd.JobTitle, hd.WorkerType,
       hd.Job_Function AS JobFunction, hd.Service_Line AS ServiceLine,
       hd.BusinessUnitDesc AS BusinessUnit, hd.Email_Address AS Email,
       hd.ManagerName AS Manager, hd.Assignment_Status AS AssignmentStatus,
       hc.IsRCM, hc.Departed
FROM hc
LEFT JOIN hd ON hd.UniversalID = hc.UniversalID;
GO

/* ---------- 9. report.training — Epic + Cornerstone combined ----------
   Shows      : one row per user — registered/trained flags from BOTH
                systems side by side + IsRegisteredAnywhere
   Built from : report.training_status + report.cornerstone_status
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.training AS
SELECT
    COALESCE(e.UniversalID, c.UniversalID)                 AS UniversalID,
    e.Wave                                                 AS EpicWave,
    ISNULL(e.IsFullyTrained,0)                             AS IsFullyTrained,
    ISNULL(e.IsFullyRegistered,0)                          AS IsFullyRegistered,
    ISNULL(c.RegisteredCount,0)                            AS CornerstoneRegistered,
    ISNULL(c.CompletedCount,0)                             AS CornerstoneCompleted,
    c.LatestRegistrationDate                               AS CornerstoneLatestRegDate,
    e.SnapshotDate                                         AS EpicSnapshotDate,
    CASE WHEN ISNULL(e.IsFullyRegistered,0) = 1
              OR ISNULL(c.RegisteredCount,0) > 0 THEN 1 ELSE 0 END AS IsRegisteredAnywhere
FROM report.training_status e
FULL OUTER JOIN report.cornerstone_status c ON c.UniversalID = e.UniversalID;
GO

/* ---------- 10. report.roster — THE wide view for ad-hoc pulls ----------
   Shows      : one row per Master person (EVERYONE; IsInScope is a flag,
                not a filter) — identity, wave, leaders, job roles, training
                status, HR detail, Epic lookup. Add a WHERE, done.
   Built from : report.users + report.training + report.hr
                + raw.mvp (latest job role) + raw.epic_lookup
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.roster AS
WITH mvp1 AS (        -- MVP job role (latest row per user)
    SELECT UniversalID, JobRole1, JobRole2, JobRole3, JobRole4 FROM (
        SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
               [IndividualCategoryUpdate1Name] AS JobRole1,
               [IndividualCategoryUpdate2Name] AS JobRole2,
               [IndividualCategoryUpdate3Name] AS JobRole3,
               [IndividualCategoryUpdate4Name] AS JobRole4,
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM(UniversalID)))
                                  ORDER BY [LastImportedDate] DESC) rn
        FROM raw.mvp WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
    ) x WHERE rn = 1
),
lk1 AS (              -- Epic team-member lookup (one row per user)
    SELECT UniversalID, Wave, EpicTrainingNeeded, EpicEligible FROM (
        SELECT UPPER(LTRIM(RTRIM([Universal_ID]))) AS UniversalID,
               [Wave]                                        AS Wave,
               [Epic_Training_Needed_MVP]                    AS EpicTrainingNeeded,
               [Epic_Training_Eligible_Appears_on_Dashboard] AS EpicEligible,
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM([Universal_ID]))) ORDER BY (SELECT NULL)) rn
        FROM raw.epic_lookup WHERE LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> ''
    ) x WHERE rn = 1
)
SELECT
    u.UniversalID, u.FullName, u.Wave, u.Leader, u.AVP, u.VP, u.SVP,
    u.UserType, u.VendorYN,
    u.TrainingNeeded, u.Departed, u.IsInScope, u.OnLeaderList,
    u.FECParticipant, u.SoftLiveParticipant, u.TerminatedFlag,
    COALESCE(NULLIF(LTRIM(RTRIM(u.JobRole1)),''), mvp1.JobRole1) AS JobRole1,
    u.JobRole2, u.JobRole3, u.JobRole4,
    mvp1.JobRole1 AS MVP_JobRole1,
    t.IsFullyTrained, t.IsFullyRegistered, t.IsRegisteredAnywhere,
    t.CornerstoneRegistered, t.CornerstoneCompleted, t.CornerstoneLatestRegDate,
    hr.Department, hr.JobTitle AS HR_JobTitle, hr.WorkerType,
    hr.Manager AS HR_Manager, hr.Email AS HR_Email,
    hr.AssignmentStatus,
    lk1.EpicTrainingNeeded, lk1.EpicEligible
FROM report.users u
LEFT JOIN report.training t ON t.UniversalID = u.UniversalID
LEFT JOIN report.hr hr      ON hr.UniversalID = u.UniversalID
LEFT JOIN mvp1              ON mvp1.UniversalID = u.UniversalID
LEFT JOIN lk1               ON lk1.UniversalID = u.UniversalID;
GO

/* ---------- 11. report.registration_daily — the day-over-day trend ----------
   Shows      : one row per day x wave x leader x vendor x user type —
                people / in-scope / registered / trained counts. SUM further
                or WHERE to slice; other dims: query history.roster_daily.
                Canonical leaders only (raw.ref_leaders, evaluated against
                the CURRENT list) — matches kpi_summary / leader_summary.
   Built from : history.roster_daily + raw.ref_leaders (needs >= 1 snapshot)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.registration_daily AS
SELECT
    SnapshotDate,
    Wave, Leader, VendorYN, UserType,
    COUNT(*)                                                          AS People,
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1 THEN 1 ELSE 0 END) AS InScope,
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1
              AND TRY_CAST(IsFullyRegistered AS float) = 1 THEN 1 ELSE 0 END) AS Registered,
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1
              AND TRY_CAST(IsFullyTrained AS float) = 1 THEN 1 ELSE 0 END)    AS Trained
FROM history.roster_daily
WHERE LTRIM(RTRIM(ISNULL(Leader,''))) IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
GROUP BY SnapshotDate, Wave, Leader, VendorYN, UserType;
GO

/* ---------- 12. report.daily_changes — who flipped, day over day ----------
   Shows      : one row per person per snapshot where something changed vs
                their previous snapshot — became registered/trained, wave or
                leader moved (with the prior value). Aggregate as needed.
   Built from : history.roster_daily
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.daily_changes AS
WITH h AS (
    SELECT SnapshotDate, UniversalID, FullName, Leader, VendorYN, Wave,
           CASE WHEN TRY_CAST(IsFullyRegistered AS float) = 1 THEN 1 ELSE 0 END AS Reg,
           CASE WHEN TRY_CAST(IsFullyTrained    AS float) = 1 THEN 1 ELSE 0 END AS Trn
    FROM history.roster_daily
),
d AS (
    SELECT *,
           LAG(Reg)    OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevReg,
           LAG(Trn)    OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevTrn,
           LAG(Wave)   OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevWave,
           LAG(Leader) OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevLeader
    FROM h
)
SELECT
    SnapshotDate, UniversalID, FullName, Leader, VendorYN, Wave,
    CASE WHEN Reg = 1 AND PrevReg = 0 THEN 1 ELSE 0 END AS BecameRegistered,
    CASE WHEN Reg = 0 AND PrevReg = 1 THEN 1 ELSE 0 END AS LostRegistration,
    CASE WHEN Trn = 1 AND PrevTrn = 0 THEN 1 ELSE 0 END AS BecameTrained,
    CASE WHEN ISNULL(Wave,'')   <> ISNULL(PrevWave,'')   THEN PrevWave   END AS WaveChangedFrom,
    CASE WHEN ISNULL(Leader,'') <> ISNULL(PrevLeader,'') THEN PrevLeader END AS LeaderChangedFrom
FROM d
WHERE PrevReg IS NOT NULL          -- needs a previous snapshot to compare against
  AND (   Reg <> PrevReg
       OR Trn <> PrevTrn
       OR ISNULL(Wave,'')   <> ISNULL(PrevWave,'')
       OR ISNULL(Leader,'') <> ISNULL(PrevLeader,''));
GO

/* ---------- 13. report.person_events — the lifecycle event log ----------
   Shows      : one row per person per event — added to Master; wave / job
                role / leader / vendor / departed / HR status (LOA) changed;
                dropped off the Master. Replaces the in-cell "Users Wave/HR
                Change Log" columns for anything after 2026-07-06.
   Built from : history.roster_daily
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.person_events AS
WITH h AS (
    SELECT SnapshotDate, UniversalID, FullName, Wave, Leader, JobRole1,
           VendorYN, Departed, AssignmentStatus, TerminatedFlag
    FROM history.roster_daily
),
d AS (
    SELECT h.*,
        LAG(SnapshotDate)     OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevDate,
        LEAD(SnapshotDate)    OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS NextDate,
        LAG(Wave)             OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevWave,
        LAG(Leader)           OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevLeader,
        LAG(JobRole1)         OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevJobRole1,
        LAG(VendorYN)         OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevVendorYN,
        LAG(Departed)         OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevDeparted,
        LAG(AssignmentStatus) OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevStatus,
        LAG(TerminatedFlag)   OVER (PARTITION BY UniversalID ORDER BY SnapshotDate) AS PrevTerminated
    FROM h
)
SELECT SnapshotDate AS EventDate, UniversalID, FullName, Leader, Wave,
       'Added to Master' AS EventType,
       CAST(NULL AS nvarchar(4000)) AS FromValue, Wave AS ToValue
FROM d
WHERE PrevDate IS NULL
  AND SnapshotDate > (SELECT MIN(SnapshotDate) FROM history.roster_daily)
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'Wave changed', PrevWave, Wave
FROM d WHERE PrevDate IS NOT NULL AND ISNULL(Wave,'') <> ISNULL(PrevWave,'')
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'Job role changed', PrevJobRole1, JobRole1
FROM d WHERE PrevDate IS NOT NULL AND ISNULL(JobRole1,'') <> ISNULL(PrevJobRole1,'')
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'Leader changed', PrevLeader, Leader
FROM d WHERE PrevDate IS NOT NULL AND ISNULL(Leader,'') <> ISNULL(PrevLeader,'')
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'Vendor flag changed', PrevVendorYN, VendorYN
FROM d WHERE PrevDate IS NOT NULL AND ISNULL(VendorYN,'') <> ISNULL(PrevVendorYN,'')
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'Departed/Inactive changed', PrevDeparted, Departed
FROM d WHERE PrevDate IS NOT NULL AND ISNULL(Departed,'') <> ISNULL(PrevDeparted,'')
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'HR status changed (Active/LOA)', PrevStatus, AssignmentStatus
FROM d WHERE PrevDate IS NOT NULL AND ISNULL(AssignmentStatus,'') <> ISNULL(PrevStatus,'')
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'Terminated flag changed', PrevTerminated, TerminatedFlag
FROM d WHERE PrevDate IS NOT NULL AND ISNULL(TerminatedFlag,'') <> ISNULL(PrevTerminated,'')
UNION ALL
SELECT SnapshotDate, UniversalID, FullName, Leader, Wave,
       'Dropped off Master', Wave, CAST(NULL AS nvarchar(4000))
FROM d
WHERE NextDate IS NULL
  AND SnapshotDate < (SELECT MAX(SnapshotDate) FROM history.roster_daily);
GO

/* ---------- 14. report.epic_not_in_hr — daily boss gap report ----------
   Shows      : Epic Lookup people with NO row in HR — possible missed
                terminations or brand-new adds. "Possible new adds" =
                IsRCMCurriculum = 1 AND OnMaster = 0.
   Built from : raw.epic_lookup + raw.hr + raw.master
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.epic_not_in_hr AS
WITH ep AS (
    SELECT UPPER(LTRIM(RTRIM([Universal_ID])))            AS UniversalID,
           [Team_Member]                                  AS FullName,
           [Job_Title]                                    AS JobTitle,
           [Curriculum_Type]                              AS CurriculumType,
           [Wave]                                         AS Wave,
           [Worker_Type]                                  AS WorkerType,
           [Business_Unit]                                AS BusinessUnit,
           [Epic_Training_Needed_MVP]                     AS EpicTrainingNeeded,
           [Epic_Training_Eligible_Appears_on_Dashboard]  AS EpicEligible,
           [Hr_Status_PDM]                                AS HrStatusPDM,
           [Offshore_Vender_PDM]                          AS OffshoreVendorPDM,
           [Direct_Manager]                               AS DirectManager,
           [Team_Member_Hire_Date]                        AS HireDate,
           ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM([Universal_ID])))
                              ORDER BY (SELECT NULL)) rn
    FROM raw.epic_lookup
    WHERE LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> ''
),
h AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID
    FROM raw.hr WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
),
m AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID
    FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
)
SELECT ep.UniversalID, ep.FullName, ep.JobTitle, ep.CurriculumType, ep.Wave,
       ep.WorkerType, ep.BusinessUnit, ep.EpicTrainingNeeded, ep.EpicEligible,
       ep.HrStatusPDM, ep.OffshoreVendorPDM, ep.DirectManager, ep.HireDate,
       CASE WHEN m.UniversalID IS NOT NULL THEN 1 ELSE 0 END AS OnMaster,
       CASE WHEN ep.CurriculumType = 'Revenue Cycle - Centralized' THEN 1 ELSE 0 END AS IsRCMCurriculum
FROM ep
LEFT JOIN h ON h.UniversalID = ep.UniversalID
LEFT JOIN m ON m.UniversalID = ep.UniversalID
WHERE ep.rn = 1 AND h.UniversalID IS NULL;
GO

/* ---------- 15. report.noshow — Cornerstone STANDING no-shows, roster-enriched ----------
   Shows      : one row per STANDING no-show for OUR population only —
                wave-file people whose Leader is on the canonical list
                (raw.ref_leaders). Since 2026-07-17 the Enterprise export
                carries full transcript history, so a no-show row alone no
                longer means unresolved: this keeps only person x class pairs
                whose NEWEST live transcript row (Withdrawn / Cancelled /
                Denied / Waitlist Expired ignored) is still 'No Show'.
                Resolved no-shows (re-registered / completed after) live in
                report.w3_noshow_status.
   Built from : raw.cornerstone + report.roster + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.noshow AS
WITH latest AS (   -- newest live transcript row per person x class, Sessions only
    SELECT *
    FROM (
        SELECT
            UPPER(LTRIM(RTRIM(c.[User_ID])))    AS UniversalID,
            c.User_Full_Name, c.Training_Title, c.Training_Type,
            c.Training_Start_Date, c.Transcript_Registration_Date,
            c.Job_Title, c.Business_Unit, c.Transcript_Status,
            ROW_NUMBER() OVER (
                PARTITION BY UPPER(LTRIM(RTRIM(c.[User_ID]))),
                             UPPER(LTRIM(RTRIM(c.Training_Title)))
                ORDER BY TRY_CONVERT(datetime, c.Transcript_Registration_Date) DESC) AS rn
        FROM raw.cornerstone c
        WHERE UPPER(LTRIM(RTRIM(ISNULL(c.Training_Provider,'')))) = 'EPIC'
          AND c.Training_Type = 'Session'
          AND c.Transcript_Status NOT IN ('Withdrawn','Cancelled','Denied','Waitlist Expired')
          AND UPPER(c.Training_Title) NOT LIKE '%W2%'      -- match the old export's
          AND UPPER(c.Training_Title) NOT LIKE '%WAVE 2%'  -- W2-title exclusion
    ) x WHERE rn = 1
)
SELECT
    l.UniversalID,
    l.User_Full_Name                    AS FullName,
    r.Wave, r.Leader, r.AVP, r.VP, r.VendorYN, r.UserType, r.IsInScope,
    l.Training_Title, l.Training_Type,
    l.Training_Start_Date, l.Transcript_Registration_Date,
    l.Job_Title                         AS Cornerstone_JobTitle,
    l.Business_Unit                     AS Cornerstone_BusinessUnit
FROM latest l
JOIN report.roster r ON r.UniversalID = l.UniversalID
WHERE UPPER(LTRIM(RTRIM(l.Transcript_Status))) = 'NO SHOW'
  AND r.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders);
GO

/* ---------- 16. report.unregistered — unregistered sessions, roster-enriched ----------
   Shows      : ONE row per person per UNIQUE unregistered event class for OUR
                population only — wave-file people whose Leader is on the
                canonical list (raw.ref_leaders). Derived from the curriculum
                status export (Event_Class_Registered = 'No'). The same class
                can appear under several curricula/sequences in the export —
                deduped here (2026-07-20) so session counts are unique classes
                per person; the kept row is the earliest-dated one.
                Missing_Sessions = the person's unique unregistered classes.
   Built from : raw.epic_status + report.roster + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.unregistered AS
WITH src AS (
    SELECT
        UPPER(LTRIM(RTRIM(s.[Universal_Id])))  AS UniversalID,
        s.Team_Member                          AS FullName,
        r.Wave, r.Leader, r.AVP, r.VP, r.VendorYN, r.UserType, r.IsInScope,
        s.Direct_Manager,
        s.Department                           AS Department_Unit,
        s.Facility_Serviceline,
        s.Curriculum_type                      AS CurriculumType,
        s.Curriculum, s.Sequence, s.Event_Class, s.Event_Class_Type,
        s.Event_Class_Date, s.Event_Class_Status, s.Event_Class_Registered,
        s.Fully_Trained, s.Fully_Registered,
        ROW_NUMBER() OVER (
            PARTITION BY UPPER(LTRIM(RTRIM(s.[Universal_Id]))),
                         UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200), s.Event_Class))))
            ORDER BY CASE WHEN TRY_CONVERT(datetime, s.Event_Class_Date) IS NULL THEN 1 ELSE 0 END,
                     TRY_CONVERT(datetime, s.Event_Class_Date) ASC,
                     s.Sequence) AS rn
    FROM raw.epic_status s
    JOIN report.roster r ON r.UniversalID = UPPER(LTRIM(RTRIM(s.[Universal_Id])))
    WHERE UPPER(LTRIM(RTRIM(ISNULL(s.Event_Class_Registered,'')))) = 'NO'
      AND r.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
)
SELECT
    UniversalID, FullName, Wave, Leader, AVP, VP, VendorYN, UserType, IsInScope,
    Direct_Manager, Department_Unit, Facility_Serviceline, CurriculumType,
    Curriculum, Sequence, Event_Class, Event_Class_Type,
    Event_Class_Date, Event_Class_Status, Event_Class_Registered,
    COUNT(*) OVER (PARTITION BY UniversalID) AS Missing_Sessions,
    Fully_Trained, Fully_Registered
FROM src
WHERE rn = 1;
GO

/* ---------- 17. report.cornerstone_inscope — Cornerstone transcripts, in-scope only ----------
   Shows      : one row per Cornerstone transcript (ALL statuses, and since
                2026-07-17 full history — every attempt — plus Online Class
                and Test rows, not just Sessions) for OUR population only —
                wave-file people whose Leader is on the canonical list
                (raw.ref_leaders). The raw Cornerstone feed is org-wide
                (active users); never report from raw.cornerstone directly.
   Built from : raw.cornerstone + report.roster + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.cornerstone_inscope AS
SELECT
    UPPER(LTRIM(RTRIM(c.[User_ID])))    AS UniversalID,
    c.User_Full_Name                    AS FullName,
    r.Wave, r.Leader, r.AVP, r.VP, r.VendorYN, r.UserType, r.IsInScope,
    c.Training_Title, c.Training_Type, c.Training_Provider,
    c.Transcript_Status,
    c.Transcript_Registration_Date, c.Training_Start_Date,
    c.Transcript_Completed_Date, c.Transcript_Score,
    c.Job_Title                         AS Cornerstone_JobTitle,
    c.Business_Unit                     AS Cornerstone_BusinessUnit
FROM raw.cornerstone c
JOIN report.roster r ON r.UniversalID = UPPER(LTRIM(RTRIM(c.[User_ID])))
WHERE UPPER(LTRIM(RTRIM(ISNULL(c.Training_Provider,'')))) = 'EPIC'
  AND r.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders);
GO

/* ---------- 18. report.withdrawal_class_detail — exact classes: withdrawn / re-registered / pending ----------
   Shows      : one row per WAVE 3 person x class (our population), straightforward:
                WasWithdrawn, CornerstoneStatus (current transcript),
                RegisteredPerEpic, and ONE plain ClassStatus:
                Re-registered / Pending re-registration / Completed / No Show /
                Registered / Never registered.
                Sources: Cornerstone Enterprise export (full history since
                2026-07-17) = system of record (shows the withdrawn row AND
                the new registered row);
                Epic Curriculum Status Detail = the filtered reporting view
                (RegisteredPerEpic is what we report as registered/unregistered).
   Built from : raw.cornerstone + raw.epic_status
                + report.users + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.withdrawal_class_detail AS
WITH cs AS (   -- Cornerstone: one row per person x class, was it ever withdrawn
    SELECT
        UPPER(LTRIM(RTRIM([User_ID])))        AS UniversalID,
        UPPER(LTRIM(RTRIM([Training_Title]))) AS ClassKey,
        MAX([Training_Title])                 AS TrainingTitle,
        MAX(CASE WHEN [Transcript_Status] = 'Withdrawn' THEN 1 ELSE 0 END) AS WasWithdrawn
    FROM raw.cornerstone
    WHERE LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
      AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
      AND [Training_Type] = 'Session'
      AND UPPER([Training_Title]) NOT LIKE '%W2%'
      AND UPPER([Training_Title]) NOT LIKE '%WAVE 2%'
    GROUP BY UPPER(LTRIM(RTRIM([User_ID]))), UPPER(LTRIM(RTRIM([Training_Title])))
),
cur AS (       -- newest ACTIVE transcript (not Withdrawn/Cancelled/Denied) per person x class
    SELECT UniversalID, ClassKey, CornerstoneStatus, SessionDate, RegisteredDate
    FROM (
        SELECT UPPER(LTRIM(RTRIM([User_ID])))        AS UniversalID,
               UPPER(LTRIM(RTRIM([Training_Title]))) AS ClassKey,
               [Transcript_Status]                          AS CornerstoneStatus,
               TRY_CONVERT(datetime, [Training_Start_Date]) AS SessionDate,
               TRY_CONVERT(datetime, [Transcript_Registration_Date]) AS RegisteredDate,
               ROW_NUMBER() OVER (
                   PARTITION BY UPPER(LTRIM(RTRIM([User_ID]))), UPPER(LTRIM(RTRIM([Training_Title])))
                   ORDER BY TRY_CONVERT(datetime, [Transcript_Registration_Date]) DESC) AS rn
        FROM raw.cornerstone
        WHERE [Transcript_Status] NOT IN ('Withdrawn','Cancelled','Denied')
          AND LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
      AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
      AND [Training_Type] = 'Session'
      AND UPPER([Training_Title]) NOT LIKE '%W2%'
      AND UPPER([Training_Title]) NOT LIKE '%WAVE 2%'
    ) x WHERE rn = 1
),
ep AS (        -- Epic Curriculum Status Detail: required sessions + registered flag
    SELECT
        UPPER(LTRIM(RTRIM([Universal_Id]))) AS UniversalID,
        UPPER(LTRIM(RTRIM([Event_Class])))  AS ClassKey,
        MAX([Event_Class])                  AS EpicEventClass,
        CASE MAX(CASE WHEN UPPER([Event_Class_Registered]) LIKE 'YES%'    THEN 3
                      WHEN UPPER([Event_Class_Registered]) LIKE 'EXEMPT%' THEN 2
                      ELSE 1 END)
             WHEN 3 THEN 'Yes' WHEN 2 THEN 'Exempt' ELSE 'No' END AS RegisteredPerEpic
    FROM raw.epic_status
    WHERE [Event_Class_Type] = 'Session'
      AND LTRIM(RTRIM(ISNULL([Universal_Id],''))) <> ''
      AND LTRIM(RTRIM(ISNULL([Event_Class],''))) <> ''
      AND UPPER([Event_Class]) NOT LIKE '%W2%'
      AND UPPER([Event_Class]) NOT LIKE '%WAVE 2%'
    GROUP BY UPPER(LTRIM(RTRIM([Universal_Id]))), UPPER(LTRIM(RTRIM([Event_Class])))
),
allkeys AS (   -- every person x class either source knows about
    SELECT UniversalID, ClassKey FROM cs
    UNION
    SELECT UniversalID, ClassKey FROM ep
)
SELECT
    k.UniversalID, r.FullName, r.Wave, r.Leader, r.AVP, r.VP, r.IsInScope,
    COALESCE(cs.TrainingTitle, ep.EpicEventClass) AS ClassTitle,
    CASE WHEN ep.UniversalID IS NOT NULL THEN 1 ELSE 0 END AS RequiredPerEpic,
    ep.RegisteredPerEpic,
    ISNULL(cs.WasWithdrawn, 0) AS WasWithdrawn,
    cur.CornerstoneStatus, cur.SessionDate,
    cur.RegisteredDate AS CornerstoneRegisteredDate,
    CASE
        WHEN cur.CornerstoneStatus = 'Completed' THEN 'Completed'
        WHEN cur.CornerstoneStatus = 'No Show'   THEN 'No Show'
        WHEN ISNULL(cs.WasWithdrawn,0) = 1
             AND cur.CornerstoneStatus IS NOT NULL THEN 'Re-registered'
        WHEN ISNULL(cs.WasWithdrawn,0) = 1        THEN 'Pending re-registration'
        WHEN cur.CornerstoneStatus IS NOT NULL    THEN 'Registered'
        WHEN ep.RegisteredPerEpic = 'No'          THEN 'Never registered'
        ELSE 'Registered'
    END AS ClassStatus
FROM allkeys k
JOIN report.users r ON r.UniversalID = k.UniversalID
LEFT JOIN cs  ON cs.UniversalID  = k.UniversalID AND cs.ClassKey  = k.ClassKey
LEFT JOIN cur ON cur.UniversalID = k.UniversalID AND cur.ClassKey = k.ClassKey
LEFT JOIN ep  ON ep.UniversalID  = k.UniversalID AND ep.ClassKey  = k.ClassKey
WHERE r.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
  AND r.Wave = 'Wave 3';
GO

/* ---------- 19. report.withdrawal_person_status — the straight count ----------
   Shows      : one row per WAVE 3 person (our population) with ONE Status:
                Not affected / Re-registered / Pending re-registration.
                Affected = they lost registrations after the 2026-07-06
                baseline snapshot (the incident). Re-registered vs Pending is
                decided by the Epic reporting view: pending = still has
                required classes with RegisteredPerEpic = No (or withdrawn
                classes with no new registration in Cornerstone).
   Built from : history.roster_daily (baseline) + report.withdrawal_class_detail
                + report.users + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.withdrawal_person_status AS
WITH base AS (   -- registered before the incident (2026-07-06 snapshot)
    SELECT UniversalID,
        CASE WHEN TRY_CAST(IsFullyRegistered    AS float) = 1 THEN 1 ELSE 0 END AS WasFullyRegistered,
        CASE WHEN TRY_CAST(IsRegisteredAnywhere AS float) = 1
               OR TRY_CAST(CornerstoneRegistered AS float) > 0 THEN 1 ELSE 0 END AS WasRegistered
    FROM history.roster_daily
    WHERE SnapshotDate = '2026-07-06'
),
lost AS (        -- lost those registrations in any later snapshot = affected
    SELECT h.UniversalID, MAX(CASE
        WHEN b.WasFullyRegistered = 1 AND TRY_CAST(h.IsFullyRegistered AS float) = 0 THEN 1
        WHEN b.WasRegistered = 1 AND ISNULL(TRY_CAST(h.CornerstoneRegistered AS float),0) = 0 THEN 1
        ELSE 0 END) AS Affected
    FROM history.roster_daily h
    JOIN base b ON b.UniversalID = h.UniversalID
    WHERE h.SnapshotDate > '2026-07-06'
    GROUP BY h.UniversalID
),
cls AS (
    SELECT UniversalID,
        SUM(CASE WHEN WasWithdrawn = 1 THEN 1 ELSE 0 END)                    AS ClassesWithdrawn,
        SUM(CASE WHEN WasWithdrawn = 1 AND ClassStatus IN ('Re-registered','Completed')
                 THEN 1 ELSE 0 END)                                          AS ClassesReRegistered,
        SUM(CASE WHEN ClassStatus = 'Pending re-registration' THEN 1 ELSE 0 END) AS ClassesPendingReRegistration,
        SUM(CASE WHEN RequiredPerEpic = 1 AND RegisteredPerEpic = 'No' THEN 1 ELSE 0 END) AS ClassesUnregisteredPerEpic
    FROM report.withdrawal_class_detail
    GROUP BY UniversalID
)
SELECT
    r.UniversalID, r.FullName, r.Wave, r.Leader, r.AVP, r.VP, r.IsInScope,
    ISNULL(l.Affected, 0)                     AS Affected,
    ISNULL(c.ClassesWithdrawn, 0)             AS ClassesWithdrawn,
    ISNULL(c.ClassesReRegistered, 0)          AS ClassesReRegistered,
    ISNULL(c.ClassesPendingReRegistration, 0) AS ClassesPendingReRegistration,
    ISNULL(c.ClassesUnregisteredPerEpic, 0)   AS ClassesUnregisteredPerEpic,
    CASE
        WHEN ISNULL(l.Affected, 0) = 0 THEN 'Not affected'
        WHEN ISNULL(c.ClassesUnregisteredPerEpic, 0) = 0
         AND ISNULL(c.ClassesPendingReRegistration, 0) = 0 THEN 'Re-registered'
        ELSE 'Pending re-registration'
    END AS Status
FROM report.users r
LEFT JOIN base b ON b.UniversalID = r.UniversalID
LEFT JOIN lost l ON l.UniversalID = r.UniversalID
LEFT JOIN cls  c ON c.UniversalID = r.UniversalID
WHERE r.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
  AND r.Wave = 'Wave 3';
GO

/* ---------- 20. report.withdrawal_transcripts — transcript review, one row per user x course ----------
   Shows      : WAVE 3, our population. Original transcript (earliest row per
                user+course) vs current transcript (newest row), with dates:
                Original_Registration_Date, ReRegistration_Date (newest active
                row after a withdrawal), Completed_Date. Status is exactly:
                Affected (withdrawn, no new registration yet) / Re-registered /
                Not affected.
                Withdrawal_Date is NULL: Cornerstone exports do not carry the
                withdrawal timestamp (incident window = 2026-07-08/09 per the
                daily snapshots).
   Built from : raw.cornerstone + report.users + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.withdrawal_transcripts AS
WITH tx AS (
    SELECT UPPER(LTRIM(RTRIM([User_ID])))        AS UniversalID,
           UPPER(LTRIM(RTRIM([Training_Title]))) AS ClassKey,
           [Training_Title], [Training_Type], [Transcript_Status],
           TRY_CONVERT(datetime, [Transcript_Registration_Date]) AS RegDate,
           TRY_CONVERT(datetime, [Transcript_Completed_Date])    AS CompletedDate
    FROM raw.cornerstone
    WHERE LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
      AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
      AND [Training_Type] = 'Session'
      AND UPPER([Training_Title]) NOT LIKE '%W2%'
      AND UPPER([Training_Title]) NOT LIKE '%WAVE 2%'
),
first_row AS (   -- earliest transcript per user+course = the original
    SELECT * FROM (SELECT tx.*, ROW_NUMBER() OVER (
        PARTITION BY UniversalID, ClassKey ORDER BY RegDate ASC) AS rn FROM tx) a
    WHERE rn = 1
),
last_row AS (    -- newest transcript per user+course = the current one
    SELECT * FROM (SELECT tx.*, ROW_NUMBER() OVER (
        PARTITION BY UniversalID, ClassKey ORDER BY RegDate DESC) AS rn FROM tx) a
    WHERE rn = 1
),
last_active AS ( -- newest NON-withdrawn/cancelled/denied row = the re-registration
    SELECT * FROM (SELECT tx.*, ROW_NUMBER() OVER (
        PARTITION BY UniversalID, ClassKey ORDER BY RegDate DESC) AS rn
        FROM tx WHERE [Transcript_Status] NOT IN ('Withdrawn','Cancelled','Denied')) a
    WHERE rn = 1
),
wd AS (          -- was this user+course ever withdrawn?
    SELECT UniversalID, ClassKey,
           MAX(CASE WHEN [Transcript_Status] = 'Withdrawn' THEN 1 ELSE 0 END) AS WasWithdrawn
    FROM tx GROUP BY UniversalID, ClassKey
)
SELECT
    u.Leader,
    f.UniversalID,
    u.FullName,
    f.[Training_Title]                AS Course_Title,
    f.[Training_Type]                 AS Training_Type,
    f.[Transcript_Status]             AS Original_Status,
    f.RegDate                         AS Original_Registration_Date,
    CAST(NULL AS datetime)            AS Withdrawal_Date,   -- not in Cornerstone exports
    l.[Transcript_Status]             AS Current_Status,
    CASE WHEN w.WasWithdrawn = 1 AND a.RegDate > f.RegDate
         THEN a.RegDate END           AS ReRegistration_Date,
    l.CompletedDate                   AS Completed_Date,
    CASE WHEN w.WasWithdrawn = 1 AND a.UniversalID IS NOT NULL THEN 'Re-registered'
         WHEN w.WasWithdrawn = 1                               THEN 'Affected'
         ELSE 'Not affected' END      AS Status,
    u.Wave
FROM first_row f
JOIN last_row  l ON l.UniversalID = f.UniversalID AND l.ClassKey = f.ClassKey
JOIN wd        w ON w.UniversalID = f.UniversalID AND w.ClassKey = f.ClassKey
LEFT JOIN last_active a ON a.UniversalID = f.UniversalID AND a.ClassKey = f.ClassKey
JOIN report.users u ON u.UniversalID = f.UniversalID
WHERE u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
  AND u.Wave = 'Wave 3';
GO

/* ---------- 21. report.unregistered_vs_cornerstone — Epic unregistered bumped against Cornerstone ----------
   Shows      : WAVE 3, our population. One row per session Epic still reports
                as UNREGISTERED (same data as the Unregistered Sessions by User
                export, but live from raw.epic_status), with what Cornerstone
                says about it right now. Resolution:
                  Registered in Cornerstone - awaiting Epic sync  (already fixed)
                  Pending re-registration   (withdrawn, nothing new yet)
                  Never registered          (pre-incident backlog)
                  No Show
                As Epic absorbs Cornerstone re-registrations, rows leave this
                view. Track the daily count dropping.
   Built from : report.withdrawal_class_detail
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.unregistered_vs_cornerstone AS
SELECT
    Leader, UniversalID, FullName,
    ClassTitle        AS Course_Title,
    WasWithdrawn,
    CornerstoneStatus,                    -- NULL = no active Cornerstone transcript
    CornerstoneRegisteredDate AS ReRegistered_Date,   -- when Cornerstone took the (re-)registration
    SessionDate       AS Cornerstone_SessionDate,
    CASE
        WHEN CornerstoneStatus IN ('Registered','Completed')
            THEN 'Registered in Cornerstone - awaiting Epic sync'
        WHEN CornerstoneStatus = 'No Show' THEN 'No Show'
        WHEN WasWithdrawn = 1 THEN 'Pending re-registration'
        ELSE 'Never registered'
    END AS Resolution,
    Wave
FROM report.withdrawal_class_detail
WHERE RequiredPerEpic = 1 AND RegisteredPerEpic = 'No';
GO

/* ---------- 22. report.withdrawal_timeline — the drop-and-recovery curve ----------
   Shows      : one row per snapshot day x wave — the SAME wave-file people
                measured different ways: WavePeople (in scope), FullyRegistered
                (all required classes, per the Epic report), AnyClassRegistered
                (at least 1 class in either system), Cornerstone_AnyClassRegistered
                (at least 1 per Cornerstone), Cornerstone_TotalRegistrations
                (class seats, not people). The 2026-07-08/09 mass-withdrawal
                drop and the recovery since.
   Built from : history.roster_daily
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.withdrawal_timeline AS
SELECT
    SnapshotDate, Wave,
    -- Same wave-file people throughout; each column is one measure of them.
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1 THEN 1 ELSE 0 END) AS WavePeople,
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1
              AND TRY_CAST(IsFullyRegistered AS float) = 1 THEN 1 ELSE 0 END)    AS FullyRegistered,
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1
              AND TRY_CAST(IsRegisteredAnywhere AS float) = 1 THEN 1 ELSE 0 END) AS AnyClassRegistered,
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1
              AND TRY_CAST(CornerstoneRegistered AS float) > 0 THEN 1 ELSE 0 END) AS Cornerstone_AnyClassRegistered,
    SUM(CASE WHEN TRY_CAST(IsInScope AS float) = 1
             THEN ISNULL(TRY_CAST(CornerstoneRegistered AS float), 0) ELSE 0 END) AS Cornerstone_TotalRegistrations
FROM history.roster_daily
GROUP BY SnapshotDate, Wave;
GO

/* ---------- 23. report.w3_unregistered — Wave 3 unregistered sessions ----------
   Shows      : WAVE 3, our population. One row per person per unregistered
                event class (Event_Class_Registered = 'No' — plain unregistered
                only; no-show-driven gaps show 'No -No Show' in Epic and are
                reported in the no-show views instead, so the two never
                double-count). Published daily by export_w3_registration_reports.py.
   Built from : report.unregistered (already leader-scoped)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_unregistered AS
SELECT * FROM report.unregistered WHERE Wave = 'Wave 3';
GO

/* ---------- 24. report.w3_noshow — Wave 3 Cornerstone no-shows ----------
   Shows      : WAVE 3, our population. One row per NO-SHOW transcript.
                Published daily by export_w3_registration_reports.py.
   Built from : report.noshow (already leader-scoped)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_noshow AS
SELECT * FROM report.noshow WHERE Wave = 'Wave 3';
GO

/* ---------- 25. report.w3_noshow_status — no-show / re-registered, per class ----------
   Shows      : WAVE 3, our population. One row per person x class that has at
                least one Cornerstone NO SHOW transcript, resolved to ONE
                Resolution from the newest live transcript row (Withdrawn /
                Cancelled / Denied / Waitlist Expired rows are ignored — they
                can't resolve a no-show):
                  No Show standing        (nothing newer — still unresolved)
                  Re-registered           (newer Registered/Approved row)
                  Completed after no-show (newer Completed row)
                  Exempted                (newer Exempt row)
                RegisteredPerEpic cross-checks what the Epic reporting view
                says about the same class right now.
   Built from : raw.cornerstone + raw.epic_status
                + report.users + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_noshow_status AS
WITH ns AS (   -- person x class: every No Show since W3 classes began.
               -- Wave 3 classes started 2026-07-13 — no-show rows with earlier
               -- session dates are data errors (wrong-class registrations) and
               -- are excluded (per the analyst 2026-07-20).
    SELECT
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID]))))        AS UniversalID,
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title])))) AS ClassKey,
        MAX([Training_Title])                 AS TrainingTitle,
        COUNT(*)                              AS NoShowCount,
        MAX(TRY_CONVERT(datetime, [Training_Start_Date])) AS LastNoShowSessionDate
    FROM raw.cornerstone
    WHERE [Transcript_Status] = 'No Show'
      AND LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
      AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
      AND [Training_Type] = 'Session'
      AND TRY_CONVERT(datetime, [Training_Start_Date]) >= '2026-07-13'
      AND UPPER([Training_Title]) NOT LIKE '%W2%'
      AND UPPER([Training_Title]) NOT LIKE '%WAVE 2%'
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID])))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title]))))
),
cur AS (       -- newest LIVE transcript row per person x class (the no-show row
               -- itself competes here: if nothing newer exists, it wins = standing)
    SELECT UniversalID, ClassKey, CurrentStatus, SessionDate, RegisteredDate
    FROM (
        SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID]))))        AS UniversalID,
               UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title])))) AS ClassKey,
               [Transcript_Status]                          AS CurrentStatus,
               TRY_CONVERT(datetime, [Training_Start_Date]) AS SessionDate,
               TRY_CONVERT(datetime, [Transcript_Registration_Date]) AS RegisteredDate,
               ROW_NUMBER() OVER (
                   PARTITION BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID])))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title]))))
                   ORDER BY TRY_CONVERT(datetime, [Transcript_Registration_Date]) DESC,
                            CASE WHEN [Transcript_Status] IN
                                 ('Registered','Registered / Past Due','Approved','Completed')
                                 THEN 0 ELSE 1 END) AS rn
        FROM raw.cornerstone
        WHERE [Transcript_Status] NOT IN ('Withdrawn','Cancelled','Denied','Waitlist Expired')
          AND LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
          AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
          AND [Training_Type] = 'Session'
          AND UPPER([Training_Title]) NOT LIKE '%W2%'
          AND UPPER([Training_Title]) NOT LIKE '%WAVE 2%'
    ) x WHERE rn = 1
),
ep AS (        -- Epic reporting view's current word on the same class
    SELECT
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[Universal_Id])))) AS UniversalID,
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Event_Class]))))  AS ClassKey,
        CASE MAX(CASE WHEN UPPER([Event_Class_Registered]) LIKE 'YES%'    THEN 3
                      WHEN UPPER([Event_Class_Registered]) LIKE 'EXEMPT%' THEN 2
                      ELSE 1 END)
             WHEN 3 THEN 'Yes' WHEN 2 THEN 'Exempt' ELSE 'No' END AS RegisteredPerEpic
    FROM raw.epic_status
    WHERE [Event_Class_Type] = 'Session'
      AND LTRIM(RTRIM(ISNULL([Universal_Id],''))) <> ''
      AND LTRIM(RTRIM(ISNULL([Event_Class],''))) <> ''
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[Universal_Id])))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Event_Class]))))
)
-- Locator CTE removed 2026-07-27: the Cornerstone Roster Report (sole Locator
-- Number source) is PAUSED per user — manually maintained by someone else,
-- unverified filters, ambiguous user-ID columns. Locator columns are kept
-- (schema stability for the W3 report exporters) but NULL while paused.
SELECT
    ns.UniversalID, u.FullName, u.Wave, u.Leader, u.AVP, u.VP,
    u.VendorYN, u.UserType, u.IsInScope,
    ns.TrainingTitle          AS ClassTitle,
    ns.NoShowCount,
    ns.LastNoShowSessionDate,
    CAST(NULL AS bigint)      AS NoShowSessionLocator,
    cur.CurrentStatus,
    cur.SessionDate           AS CurrentSessionDate,
    cur.RegisteredDate        AS CurrentRegisteredDate,
    CAST(NULL AS bigint)      AS CurrentSessionLocator,
    ep.RegisteredPerEpic,
    CASE
        WHEN cur.CurrentStatus = 'No Show'   THEN 'No Show standing'
        WHEN cur.CurrentStatus = 'Completed' THEN 'Completed after no-show'
        WHEN cur.CurrentStatus IN ('Registered','Registered / Past Due','Approved')
                                             THEN 'Re-registered'
        WHEN cur.CurrentStatus = 'Exempt'    THEN 'Exempted'
        ELSE 'No Show standing'
    END AS Resolution,
    /* RegistrationAction — the registration team's action label. For standing
       no-shows it cross-checks Epic: class still required and unregistered =
       Needs to be Re-Registered; Epic already shows a booking = Re-Registered
       (Cornerstone lag); class gone from Epic = review cases (curriculum
       changed / no curriculum at all — don't rebook the old class blindly). */
    CASE
        WHEN cur.CurrentStatus IN ('Registered','Registered / Past Due','Approved')
                                             THEN 'Re-Registered'
        WHEN cur.CurrentStatus = 'Completed' THEN 'Completed'
        WHEN cur.CurrentStatus = 'Exempt'    THEN 'Exempted'
        WHEN ep.RegisteredPerEpic = 'No'     THEN 'Needs to be Re-Registered'
        WHEN ep.RegisteredPerEpic = 'Yes'    THEN 'Re-Registered'
        WHEN ep.RegisteredPerEpic = 'Exempt' THEN 'Exempted - Review'
        WHEN epu.UniversalID IS NULL         THEN 'No Epic Curriculum - Review'
        ELSE 'Class Not Required - Review'
    END AS RegistrationAction
FROM ns
JOIN report.users u ON u.UniversalID = ns.UniversalID
LEFT JOIN cur ON cur.UniversalID = ns.UniversalID AND cur.ClassKey = ns.ClassKey
LEFT JOIN ep  ON ep.UniversalID  = ns.UniversalID AND ep.ClassKey  = ns.ClassKey
LEFT JOIN (SELECT DISTINCT UPPER(LTRIM(RTRIM([Universal_Id]))) AS UniversalID
           FROM raw.epic_status
           WHERE [Event_Class_Type] = 'Session'
             AND LTRIM(RTRIM(ISNULL([Universal_Id],''))) <> '') epu
       ON epu.UniversalID = ns.UniversalID
WHERE u.Wave = 'Wave 3'
  AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders);
GO

/* ---------- 26. report.w3_leader_snapshot — TODAY's W3 rollup by leader ----------
   Shows      : one row per leader — Wave 3 in-scope users with today's counts:
                fully registered/trained, unregistered users & sessions,
                no-show users & sessions (standing = still unresolved, and
                to-date = every no-show ever). This is the LIVE view; the daily
                snapshot step appends it to history.w3_leader_daily, which is
                what report.w3_registration_summary trends over.
   Built from : report.users + report.training_status + report.w3_unregistered
                + report.w3_noshow_status + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_leader_snapshot AS
WITH pop AS (
    SELECT u.UniversalID, u.Leader,
           ISNULL(t.IsFullyRegistered,0) AS Reg,
           ISNULL(t.IsFullyTrained,0)    AS Trn
    FROM report.users u
    LEFT JOIN report.training_status t ON t.UniversalID = u.UniversalID
    WHERE u.Wave = 'Wave 3' AND u.IsInScope = 1
      AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
),
unreg AS (
    SELECT UniversalID, COUNT(*) AS Sessions
    FROM report.w3_unregistered GROUP BY UniversalID
),
nss AS (
    SELECT UniversalID,
           SUM(CASE WHEN Resolution = 'No Show standing' THEN 1 ELSE 0 END) AS Standing,
           SUM(NoShowCount) AS ToDate
    FROM report.w3_noshow_status GROUP BY UniversalID
)
SELECT
    p.Leader,
    COUNT(*)                                          AS W3Users,
    SUM(p.Reg)                                        AS FullyRegistered,
    SUM(p.Trn)                                        AS FullyTrained,
    SUM(CASE WHEN un.Sessions > 0 THEN 1 ELSE 0 END)  AS UnregisteredUsers,
    SUM(ISNULL(un.Sessions,0))                        AS UnregisteredSessions,
    SUM(CASE WHEN ns.Standing > 0 THEN 1 ELSE 0 END)  AS NoShowUsersStanding,
    SUM(ISNULL(ns.Standing,0))                        AS NoShowSessionsStanding,
    SUM(CASE WHEN ns.ToDate > 0 THEN 1 ELSE 0 END)    AS NoShowUsersToDate,
    SUM(ISNULL(ns.ToDate,0))                          AS NoShowSessionsToDate
FROM pop p
LEFT JOIN unreg un ON un.UniversalID = p.UniversalID
LEFT JOIN nss   ns ON ns.UniversalID = p.UniversalID
GROUP BY p.Leader;
GO

/* ---------- 26b. report.w3_no_epic_curriculum — Epic doesn't know them ----------
   Shows      : in-scope WAVE 3 people (our population) with ZERO rows in the
                Epic curriculum status report — they cannot register or train
                until Epic assigns a curriculum. The hidden blocker inside the
                not-fully-registered gap: they never appear in the unregistered
                view because they have nothing to register FOR.
   Built from : report.users + raw.epic_status + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_no_epic_curriculum AS
SELECT u.UniversalID, u.FullName, u.Leader, u.AVP, u.VP,
       u.JobRole1, u.UserType, u.VendorYN
FROM report.users u
WHERE u.Wave = 'Wave 3' AND u.IsInScope = 1
  AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
  AND NOT EXISTS (SELECT 1 FROM raw.epic_status s
                  WHERE UPPER(LTRIM(RTRIM(s.Universal_Id))) = u.UniversalID);
GO

/* ---------- 26c. report.w3_scorecard — leadership scorecard, one query ----------
   Shows      : the daily leadership numbers in one place — per-leader rows
                plus a TOTAL row (IsTotal = 1): W3 users, fully registered
                (+%), fully trained (+%), unregistered users & unique
                sessions, standing no-show users & sessions, and people with
                no Epic curriculum assigned. Live counts (not the history
                snapshot). scripts\w3_scorecard.py prints this + the person-
                level breakdowns in the screenshot format for leadership.
   Built from : report.w3_leader_snapshot + report.w3_no_epic_curriculum
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_scorecard AS
WITH nc AS (
    SELECT Leader, COUNT(*) AS NoEpicCurriculum
    FROM report.w3_no_epic_curriculum GROUP BY Leader
),
base AS (
    SELECT s.Leader, s.W3Users, s.FullyRegistered, s.FullyTrained,
           s.UnregisteredUsers, s.UnregisteredSessions,
           s.NoShowUsersStanding, s.NoShowSessionsStanding,
           s.NoShowUsersToDate, s.NoShowSessionsToDate,
           ISNULL(nc.NoEpicCurriculum, 0) AS NoEpicCurriculum
    FROM report.w3_leader_snapshot s
    LEFT JOIN nc ON nc.Leader = s.Leader
)
SELECT
    CASE WHEN GROUPING(b.Leader) = 1 THEN 'TOTAL' ELSE b.Leader END AS Leader,
    SUM(b.W3Users)               AS W3Users,
    SUM(b.FullyRegistered)       AS FullyRegistered,
    CAST(100.0 * SUM(b.FullyRegistered) / NULLIF(SUM(b.W3Users),0) AS decimal(5,1)) AS PctRegistered,
    SUM(b.FullyTrained)          AS FullyTrained,
    CAST(100.0 * SUM(b.FullyTrained) / NULLIF(SUM(b.W3Users),0) AS decimal(5,1))    AS PctTrained,
    SUM(b.UnregisteredUsers)     AS UnregisteredUsers,
    SUM(b.UnregisteredSessions)  AS UnregisteredSessions,
    SUM(b.NoShowUsersStanding)   AS NoShowUsers,
    SUM(b.NoShowSessionsStanding) AS NoShowSessions,
    SUM(b.NoShowUsersToDate)     AS NoShowUsersToDate,
    SUM(b.NoShowSessionsToDate)  AS NoShowSessionsToDate,
    SUM(b.NoEpicCurriculum)      AS NoEpicCurriculum,
    GROUPING(b.Leader)           AS IsTotal
FROM base b
GROUP BY GROUPING SETS ((b.Leader), ());
GO

/* ---------- 26d. report.w3_noshow_trend — daily no-show state since 2026-07-13 ----------
   Shows      : one row per snapshot day (weekdays from history.roster_daily)
                — Wave 3 no-show counts AS OF that day, reconstructed from the
                Enterprise export's full transcript history: a no-show exists
                once its session date has passed (the attempt row keeps status
                'No Show' forever); it stops standing when a NEWER live
                transcript row (re-registration / completion / exemption) has
                been booked by that day. ToDate = every no-show ever by then;
                Standing = still unresolved that day; Resolved = the difference.
                Backfill starts 2026-07-13 (start of the week the tracker
                began). Approximation note: departed users drop out of the
                Active-only export, so long-past days can undercount slightly.
   Built from : raw.cornerstone + report.users + raw.ref_leaders
                + history.roster_daily (day spine)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_noshow_trend AS
WITH days AS (
    SELECT DISTINCT TRY_CONVERT(date, SnapshotDate) AS AsOfDate
    FROM history.roster_daily
    WHERE TRY_CONVERT(date, SnapshotDate) >= '2026-07-13'
),
pop AS (
    SELECT u.UniversalID FROM report.users u
    WHERE u.Wave = 'Wave 3' AND u.IsInScope = 1
      AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
),
tx AS (        -- live transcript rows (Sessions, EPIC, non-W2) for our W3 people
    SELECT UPPER(LTRIM(RTRIM(c.[User_ID])))                              AS uid,
           UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200), c.Training_Title)))) AS ck,
           c.Transcript_Status                                           AS st,
           TRY_CONVERT(datetime, c.Transcript_Registration_Date)         AS regdt,
           TRY_CONVERT(datetime, c.Training_Start_Date)                  AS startdt
    FROM raw.cornerstone c
    JOIN pop p ON p.UniversalID = UPPER(LTRIM(RTRIM(c.[User_ID])))
    WHERE UPPER(LTRIM(RTRIM(ISNULL(c.Training_Provider,'')))) = 'EPIC'
      AND c.Training_Type = 'Session'
      AND c.Transcript_Status NOT IN ('Withdrawn','Cancelled','Denied','Waitlist Expired')
      AND UPPER(c.Training_Title) NOT LIKE '%W2%'
      AND UPPER(c.Training_Title) NOT LIKE '%WAVE 2%'
),
asof AS (      -- newest live row per person x class as of end of each day
    SELECT d.AsOfDate, t.uid, t.ck, t.st, t.startdt,
           ROW_NUMBER() OVER (
               PARTITION BY d.AsOfDate, t.uid, t.ck
               ORDER BY t.regdt DESC) AS rn
    FROM days d
    JOIN tx t ON t.regdt < DATEADD(day, 1, CAST(d.AsOfDate AS datetime))
),
todate AS (    -- person x class pairs with ANY no-show whose session passed by day d.
               -- W3 classes started 2026-07-13; earlier no-show rows are data
               -- errors (wrong-class registrations) and are excluded.
    SELECT d.AsOfDate, t.uid, t.ck
    FROM days d
    JOIN tx t ON t.st = 'No Show'
             AND t.startdt >= '2026-07-13'
             AND t.startdt < DATEADD(day, 1, CAST(d.AsOfDate AS datetime))
    GROUP BY d.AsOfDate, t.uid, t.ck
)
SELECT
    td.AsOfDate,
    COUNT(DISTINCT td.uid)                                   AS NoShowUsersToDate,
    COUNT(*)                                                 AS NoShowSessionsToDate,
    COUNT(DISTINCT CASE WHEN a.st = 'No Show'
                        AND a.startdt < DATEADD(day, 1, CAST(td.AsOfDate AS datetime))
                        THEN td.uid END)                     AS NoShowUsersStanding,
    SUM(CASE WHEN a.st = 'No Show'
             AND a.startdt < DATEADD(day, 1, CAST(td.AsOfDate AS datetime))
             THEN 1 ELSE 0 END)                              AS NoShowSessionsStanding,
    SUM(CASE WHEN a.st = 'No Show'
             AND a.startdt < DATEADD(day, 1, CAST(td.AsOfDate AS datetime))
             THEN 0 ELSE 1 END)                              AS SessionsResolved
FROM todate td
JOIN asof a ON a.AsOfDate = td.AsOfDate AND a.uid = td.uid AND a.ck = td.ck AND a.rn = 1
GROUP BY td.AsOfDate;
GO

/* history.w3_leader_daily — created here empty so the summary view below always
   compiles; the daily snapshot step (sql\refresh.py) appends one row per leader
   per day and mirrors each day to a dated CSV under
   processed\history_snapshots\w3_leader_daily\. */
IF SCHEMA_ID('history') IS NULL EXEC('CREATE SCHEMA history');
GO
IF OBJECT_ID('history.w3_leader_daily','U') IS NULL
CREATE TABLE history.w3_leader_daily (
    SnapshotDate           nvarchar(10),
    Leader                 nvarchar(200),
    W3Users                int,
    FullyRegistered        int,
    FullyTrained           int,
    UnregisteredUsers      int,
    UnregisteredSessions   int,
    NoShowUsersStanding    int,
    NoShowSessionsStanding int,
    NoShowUsersToDate      int,
    NoShowSessionsToDate   int
);
GO

/* ---------- 27. report.w3_registration_summary — the W3 daily trend by leader ----------
   Shows      : one row per snapshot day x leader — Wave 3 daily totals
                (unregistered, no-shows, registered, trained) accumulating
                from 2026-07-15 forward. TRY_CAST everywhere so the view also
                survives a backload_history rebuild (which types everything
                NVARCHAR).
   Built from : history.w3_leader_daily (stamped daily by the snapshot step)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_registration_summary AS
SELECT
    SnapshotDate,
    Leader,
    TRY_CAST(W3Users                AS int) AS W3Users,
    TRY_CAST(FullyRegistered        AS int) AS FullyRegistered,
    TRY_CAST(FullyTrained           AS int) AS FullyTrained,
    TRY_CAST(UnregisteredUsers      AS int) AS UnregisteredUsers,
    TRY_CAST(UnregisteredSessions   AS int) AS UnregisteredSessions,
    TRY_CAST(NoShowUsersStanding    AS int) AS NoShowUsersStanding,
    TRY_CAST(NoShowSessionsStanding AS int) AS NoShowSessionsStanding,
    TRY_CAST(NoShowUsersToDate      AS int) AS NoShowUsersToDate,
    TRY_CAST(NoShowSessionsToDate   AS int) AS NoShowSessionsToDate
FROM history.w3_leader_daily;
GO

/* ---------- 28. report.w3_registration_metrics — W3 registration metrics by leader ----------
   Shows      : one row per leader — Wave 3 in-scope users. Counts + percentages:
                fully registered / not, fully trained / not, unregistered
                sessions, no-shows standing & to date, and the >=50% / >=80%
                registered buckets. RegPct per user = registered-or-Exempt
                required sessions / all required sessions (Curriculum Status
                Detail, Event_Class_Type = 'Session'; 'Test' rows excluded) —
                a user with no session detail inherits 100 if the Epic
                Fully Registered flag is Yes, else 0.
   Built from : report.users + report.training_status + raw.epic_status
                + report.w3_unregistered + report.w3_noshow_status + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_registration_metrics AS
WITH pct AS (   -- per-user % of required sessions registered (Yes or Exempt)
    SELECT UniversalID,
           CAST(100.0 * SUM(RegOK) / NULLIF(COUNT(*),0) AS decimal(5,1)) AS RegPct
    FROM (
        SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[Universal_Id])))) AS UniversalID,
               UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Event_Class]))))  AS ClassKey,
               MAX(CASE WHEN UPPER([Event_Class_Registered]) LIKE 'YES%'
                          OR UPPER([Event_Class_Registered]) LIKE 'EXEMPT%'
                        THEN 1 ELSE 0 END) AS RegOK
        FROM raw.epic_status
        WHERE [Event_Class_Type] = 'Session'
          AND LTRIM(RTRIM(ISNULL([Universal_Id],''))) <> ''
          AND LTRIM(RTRIM(ISNULL([Event_Class],''))) <> ''
          AND UPPER(LTRIM(RTRIM(ISNULL([Event_Class_Registered],'')))) <> 'TEST'
        GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[Universal_Id])))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Event_Class]))))
    ) x GROUP BY UniversalID
),
unreg AS (
    SELECT UniversalID, COUNT(*) AS Sessions
    FROM report.w3_unregistered GROUP BY UniversalID
),
nss AS (
    SELECT UniversalID,
           SUM(CASE WHEN Resolution = 'No Show standing' THEN 1 ELSE 0 END) AS Standing,
           SUM(NoShowCount) AS ToDate
    FROM report.w3_noshow_status GROUP BY UniversalID
),
pop AS (
    SELECT u.UniversalID, u.Leader,
           ISNULL(t.IsFullyRegistered,0) AS Reg,
           ISNULL(t.IsFullyTrained,0)    AS Trn,
           COALESCE(p.RegPct, CASE WHEN ISNULL(t.IsFullyRegistered,0) = 1
                                   THEN 100.0 ELSE 0 END) AS RegPct,
           ISNULL(un.Sessions,0) AS UnregSessions,
           ISNULL(ns.Standing,0) AS NsStanding,
           ISNULL(ns.ToDate,0)   AS NsToDate
    FROM report.users u
    LEFT JOIN report.training_status t ON t.UniversalID = u.UniversalID
    LEFT JOIN pct   p  ON p.UniversalID  = u.UniversalID
    LEFT JOIN unreg un ON un.UniversalID = u.UniversalID
    LEFT JOIN nss   ns ON ns.UniversalID = u.UniversalID
    WHERE u.Wave = 'Wave 3' AND u.IsInScope = 1
      AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
)
SELECT
    Leader,
    COUNT(*)                                            AS TotalUsers,
    SUM(Reg)                                            AS FullyRegistered,
    COUNT(*) - SUM(Reg)                                 AS NotFullyRegistered,
    CAST(100.0 * SUM(Reg) / NULLIF(COUNT(*),0) AS decimal(5,1)) AS PctFullyRegistered,
    SUM(Trn)                                            AS FullyTrained,
    COUNT(*) - SUM(Trn)                                 AS NotFullyTrained,
    CAST(100.0 * SUM(Trn) / NULLIF(COUNT(*),0) AS decimal(5,1)) AS PctFullyTrained,
    SUM(CASE WHEN UnregSessions > 0 THEN 1 ELSE 0 END)  AS UnregisteredUsers,
    SUM(UnregSessions)                                  AS UnregisteredSessions,
    SUM(CASE WHEN NsStanding > 0 THEN 1 ELSE 0 END)     AS NoShowUsersStanding,
    SUM(NsStanding)                                     AS NoShowSessionsStanding,
    SUM(CASE WHEN NsToDate > 0 THEN 1 ELSE 0 END)       AS NoShowUsersToDate,
    SUM(NsToDate)                                       AS NoShowSessionsToDate,
    SUM(CASE WHEN RegPct >= 50 THEN 1 ELSE 0 END)       AS UsersReg50Plus,
    CAST(100.0 * SUM(CASE WHEN RegPct >= 50 THEN 1 ELSE 0 END)
               / NULLIF(COUNT(*),0) AS decimal(5,1))    AS PctReg50Plus,
    SUM(CASE WHEN RegPct >= 80 THEN 1 ELSE 0 END)       AS UsersReg80Plus,
    CAST(100.0 * SUM(CASE WHEN RegPct >= 80 THEN 1 ELSE 0 END)
               / NULLIF(COUNT(*),0) AS decimal(5,1))    AS PctReg80Plus,
    CAST(AVG(RegPct) AS decimal(5,1))                   AS AvgRegPct
FROM pop
GROUP BY Leader;
GO

/* ---------- 29. report.w3_registration_kpis — W3 top-line, one row ----------
   Shows      : ONE row, Leader = 'ALL WAVE 3' — the same columns as
                report.w3_registration_metrics summed across all leaders,
                percentages recomputed over the whole population. The exporter
                appends this row under the leader rows as the TOTAL line.
   Built from : report.w3_registration_metrics
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_registration_kpis AS
SELECT
    'ALL WAVE 3'                                        AS Leader,
    SUM(TotalUsers)                                     AS TotalUsers,
    SUM(FullyRegistered)                                AS FullyRegistered,
    SUM(NotFullyRegistered)                             AS NotFullyRegistered,
    CAST(100.0 * SUM(FullyRegistered) / NULLIF(SUM(TotalUsers),0) AS decimal(5,1)) AS PctFullyRegistered,
    SUM(FullyTrained)                                   AS FullyTrained,
    SUM(NotFullyTrained)                                AS NotFullyTrained,
    CAST(100.0 * SUM(FullyTrained) / NULLIF(SUM(TotalUsers),0) AS decimal(5,1))    AS PctFullyTrained,
    SUM(UnregisteredUsers)                              AS UnregisteredUsers,
    SUM(UnregisteredSessions)                           AS UnregisteredSessions,
    SUM(NoShowUsersStanding)                            AS NoShowUsersStanding,
    SUM(NoShowSessionsStanding)                         AS NoShowSessionsStanding,
    SUM(NoShowUsersToDate)                              AS NoShowUsersToDate,
    SUM(NoShowSessionsToDate)                           AS NoShowSessionsToDate,
    SUM(UsersReg50Plus)                                 AS UsersReg50Plus,
    CAST(100.0 * SUM(UsersReg50Plus) / NULLIF(SUM(TotalUsers),0) AS decimal(5,1))  AS PctReg50Plus,
    SUM(UsersReg80Plus)                                 AS UsersReg80Plus,
    CAST(100.0 * SUM(UsersReg80Plus) / NULLIF(SUM(TotalUsers),0) AS decimal(5,1))  AS PctReg80Plus,
    CAST(SUM(AvgRegPct * TotalUsers) / NULLIF(SUM(TotalUsers),0) AS decimal(5,1))  AS AvgRegPct
FROM report.w3_registration_metrics;
GO

/* ---------- 30. report.w3_noshow_narrative — the no-show story, one sentence per row ----------
   Shows      : WAVE 3, our population. Same grain as report.w3_noshow_status
                (one row per person x no-showed class) but enriched for
                leadership reading: who booked the new seat (Epic coordinator),
                the exact new session (date, time window, location, instructor,
                seats from the Epic class schedule), whether the person's other
                classes are still moving, and a plain-English Summary column
                that tells the whole story — e.g. "Camino, Firstname08 no-showed
                EPIC_NH_SBO_... on Tue 7/14 12:00 PM. Re-registered Tue 7/14
                3:27 PM by Coordlast, Coordfirst into the Mon 7/20 12:00–3:00 PM
                session (Remote Class 1; instructor Instructorlast, Instructorfirst...;
                60 of 120 seats taken)."
   Built from : report.w3_noshow_status + raw.epic_status
                + raw.epic_class_schedule + report.cornerstone_inscope
                + raw.cornerstone
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_noshow_narrative AS
WITH ep AS (   -- Epic's booking record for the re-registered class
    SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[Universal_Id])))) AS UniversalID,
           UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Event_Class]))))  AS ClassKey,
           MAX([Senior_Registration_Coordinator]) AS BookedBy,
           MAX([Event_Class_Location])            AS EpicLocation
    FROM raw.epic_status
    WHERE UPPER(ISNULL([Event_Class_Registered],'')) LIKE 'YES%'
      AND LTRIM(RTRIM(ISNULL([Universal_Id],''))) <> ''
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[Universal_Id])))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Event_Class]))))
),
done AS (      -- completion timestamp, if the class was finished after the no-show
    SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID]))))        AS UniversalID,
           UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title])))) AS ClassKey,
           MAX(TRY_CONVERT(datetime, [Transcript_Completed_Date])) AS CompletedOn
    FROM raw.cornerstone
    WHERE [Transcript_Status] = 'Completed'
      AND [Training_Type] = 'Session'
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID])))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title]))))
),
sched AS (     -- the exact session on the Epic class schedule (instructor list
               -- in the export is doubled "X; X" — halve it when it is)
    SELECT ClassKey, SessDate, StartTime, EndTime, Training_Location, Room,
           CASE WHEN Instructor IS NOT NULL AND LEN(Instructor) > 4
                 AND Instructor = LEFT(Instructor,(LEN(Instructor)-2)/2) + '; '
                                + LEFT(Instructor,(LEN(Instructor)-2)/2)
                THEN LEFT(Instructor,(LEN(Instructor)-2)/2) ELSE Instructor END AS Instructor,
           TotalSeats, SeatsTaken
    FROM (
        SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Event_Name))))            AS ClassKey,
               TRY_CONVERT(date, [Date])                  AS SessDate,
               TRY_CONVERT(time, Class_Event_Start_Time)  AS StartTime,
               TRY_CONVERT(time, Class_Event_End_Time)    AS EndTime,
               Training_Location, Room,
               Primary_Instructor                         AS Instructor,
               TRY_CAST(Total_Seats AS int)               AS TotalSeats,
               TRY_CAST(Total_Seats_Taken AS int)         AS SeatsTaken,
               ROW_NUMBER() OVER (
                   PARTITION BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Event_Name)))),
                                TRY_CONVERT(date, [Date]),
                                TRY_CONVERT(time, Class_Event_Start_Time),
                                Training_Location
                   ORDER BY (SELECT NULL)) AS rn
        FROM raw.epic_class_schedule
    ) x WHERE rn = 1
),
oth AS (       -- is the rest of their schedule still moving?
    SELECT UniversalID,
           COUNT(*) AS UpcomingOtherClasses,
           MIN(TRY_CONVERT(datetime, Training_Start_Date)) AS NextOtherClassDate
    FROM report.cornerstone_inscope
    WHERE Transcript_Status IN ('Registered','Registered / Past Due','Approved')
      AND TRY_CONVERT(datetime, Training_Start_Date) >= CAST(GETDATE() AS date)
    GROUP BY UniversalID
)
SELECT
    b.Leader, b.FullName, b.UniversalID, b.ClassTitle, b.Resolution,
    b.LastNoShowSessionDate                AS NoShowedOn,
    b.NoShowSessionLocator,
    b.CurrentRegisteredDate                AS ReRegisteredOn,
    ep.BookedBy,
    b.CurrentSessionDate                   AS NewSessionDate,
    b.CurrentSessionLocator                AS NewSessionLocator,
    CASE WHEN s.StartTime IS NOT NULL THEN
         FORMAT(CAST(s.StartTime AS datetime),'h:mm') + '-'
       + FORMAT(CAST(s.EndTime   AS datetime),'h:mm tt') END AS SessionTime,
    COALESCE(s.Room, s.Training_Location)  AS Location,
    s.Instructor,
    s.SeatsTaken, s.TotalSeats,
    d.CompletedOn,
    ISNULL(o.UpcomingOtherClasses, 0)      AS UpcomingOtherClasses,
    o.NextOtherClassDate,
    CONCAT(
        b.FullName, ' no-showed ', b.ClassTitle, ' on ',
        FORMAT(b.LastNoShowSessionDate, 'ddd M/d h:mm tt'),
        CASE WHEN b.NoShowSessionLocator IS NOT NULL
             THEN CONCAT(' (locator ', b.NoShowSessionLocator, ')') ELSE '' END,
        CASE b.Resolution
        WHEN 'Completed after no-show' THEN CONCAT(
            '. Re-registered ', FORMAT(b.CurrentRegisteredDate, 'ddd M/d h:mm tt'),
            CASE WHEN NULLIF(LTRIM(RTRIM(ep.BookedBy)),'') IS NOT NULL
                 THEN ' by ' + LTRIM(RTRIM(ep.BookedBy)) ELSE '' END,
            CASE WHEN b.CurrentSessionLocator IS NOT NULL
                 THEN CONCAT(' (locator ', b.CurrentSessionLocator, ')') ELSE '' END,
            ' and COMPLETED the class ',
            ISNULL(FORMAT(d.CompletedOn, 'ddd M/d h:mm tt'), ''), '.')
        WHEN 'Re-registered' THEN CONCAT(
            '. Re-registered ', FORMAT(b.CurrentRegisteredDate, 'ddd M/d h:mm tt'),
            CASE WHEN NULLIF(LTRIM(RTRIM(ep.BookedBy)),'') IS NOT NULL
                 THEN ' by ' + LTRIM(RTRIM(ep.BookedBy)) ELSE '' END,
            ' into the ', FORMAT(b.CurrentSessionDate, 'ddd M/d'),
            CASE WHEN s.StartTime IS NOT NULL
                 THEN ' ' + FORMAT(CAST(s.StartTime AS datetime),'h:mm') + '-'
                          + FORMAT(CAST(s.EndTime AS datetime),'h:mm tt') ELSE '' END,
            ' session',
            CASE WHEN COALESCE(s.Room, s.Training_Location) IS NOT NULL THEN CONCAT(
                 ' (',
                 CASE WHEN b.CurrentSessionLocator IS NOT NULL
                      THEN CONCAT('locator ', b.CurrentSessionLocator, '; ') ELSE '' END,
                 COALESCE(s.Room, s.Training_Location),
                 CASE WHEN s.Instructor IS NOT NULL THEN '; instructor ' + s.Instructor ELSE '' END,
                 CASE WHEN s.TotalSeats IS NOT NULL THEN CONCAT('; ', s.SeatsTaken, ' of ', s.TotalSeats, ' seats taken') ELSE '' END,
                 ')') ELSE '' END, '.')
        WHEN 'Exempted' THEN '. Marked Exempt in Cornerstone since - no seat needed.'
        ELSE CONCAT(
            '. NOT re-booked for this class yet - no-show still standing',
            CASE WHEN ISNULL(o.UpcomingOtherClasses,0) > 0
                 THEN CONCAT(', but their schedule is still moving: ',
                      o.UpcomingOtherClasses, ' upcoming registration(s) in other classes (next ',
                      FORMAT(o.NextOtherClassDate, 'ddd M/d'), ').')
                 ELSE ' and they have NO upcoming registrations in any class - needs attention.' END)
        END) AS Summary
FROM report.w3_noshow_status b
LEFT JOIN ep   ON ep.UniversalID = b.UniversalID AND ep.ClassKey = UPPER(LTRIM(RTRIM(b.ClassTitle)))
LEFT JOIN done d ON d.UniversalID = b.UniversalID AND d.ClassKey = UPPER(LTRIM(RTRIM(b.ClassTitle)))
LEFT JOIN sched s ON s.ClassKey = UPPER(LTRIM(RTRIM(b.ClassTitle)))
                 AND b.Resolution IN ('Re-registered','Completed after no-show')
                 AND s.SessDate  = CAST(b.CurrentSessionDate AS date)
                 AND s.StartTime = CAST(b.CurrentSessionDate AS time)
LEFT JOIN oth  o ON o.UniversalID = b.UniversalID;
GO

/* ---------- 31. report.w3_track_class_map — job role -> training track -> class, with progress ----------
   Shows      : WAVE 3, our population, in-scope users. One row per
                Epic Job Category x Curriculum (track) x Event Class (session),
                with how that class is going: users required / registered /
                unregistered / no-show gaps (Epic 'No -No Show') / completed
                per Cornerstone — plus whether there's capacity to fix it:
                upcoming sessions on the Epic class schedule, the next date,
                and total open seats. Sort by UsersUnregistered DESC to see
                the bottleneck classes; OpenSeats tells you if the fix is easy.
   Built from : raw.epic_status + report.users + raw.ref_leaders
                + raw.cornerstone + raw.epic_class_schedule
                (roster report PAUSED 2026-07-27 — completions now from the
                Enterprise export)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_track_class_map AS
WITH pop AS (
    SELECT u.UniversalID FROM report.users u
    WHERE u.Wave = 'Wave 3' AND u.IsInScope = 1
      AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
),
req AS (   -- one row per user x role x track x class, best registered state
    SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))  AS UniversalID,
           MAX(s.Epic_Job_Category)             AS JobRole,
           MAX(s.Curriculum)                    AS Track,
           UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))))   AS ClassKey,
           MAX(s.Event_Class)                   AS ClassTitle,
           MAX(CASE WHEN UPPER(s.Event_Class_Registered) LIKE 'YES%'    THEN 3
                    WHEN UPPER(s.Event_Class_Registered) LIKE 'EXEMPT%' THEN 2
                    WHEN UPPER(s.Event_Class_Registered) LIKE 'NO -NO SHOW%' THEN 1
                    ELSE 0 END)                 AS RegState   -- 3 reg, 2 exempt, 1 no-show gap, 0 unreg
    FROM raw.epic_status s
    JOIN pop p ON p.UniversalID = UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))
    WHERE s.Event_Class_Type = 'Session'
      AND LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id)))), s.Epic_Job_Category, s.Curriculum,
             UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))))
),
done AS (  -- distinct user x class completed per Cornerstone (Enterprise
           -- full-history export; roster report PAUSED 2026-07-27)
    SELECT DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID])))) AS UniversalID,
           UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title])))) AS ClassKey
    FROM raw.cornerstone
    WHERE [Transcript_Status] = 'Completed'
      AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
      AND [Training_Type] = 'Session'
),
sched AS ( -- future capacity per class
    SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Event_Name)))) AS ClassKey,
           COUNT(*)                        AS FutureSessions,
           MIN(TRY_CONVERT(date, [Date]))  AS NextSessionDate,
           SUM(TRY_CAST(Available_Seats AS int)) AS OpenSeats
    FROM raw.epic_class_schedule
    WHERE TRY_CONVERT(date, [Date]) >= CAST(GETDATE() AS date)
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Event_Name))))
)
SELECT
    r.JobRole, r.Track, MAX(r.ClassTitle) AS ClassTitle,
    COUNT(*)                                              AS UsersRequired,
    SUM(CASE WHEN r.RegState >= 2 THEN 1 ELSE 0 END)      AS UsersRegisteredOrExempt,
    SUM(CASE WHEN r.RegState = 1 THEN 1 ELSE 0 END)       AS UsersNoShowGap,
    SUM(CASE WHEN r.RegState = 0 THEN 1 ELSE 0 END)       AS UsersUnregistered,
    SUM(CASE WHEN d.ClassKey IS NOT NULL THEN 1 ELSE 0 END) AS UsersCompleted,
    CAST(100.0 * SUM(CASE WHEN r.RegState >= 2 THEN 1 ELSE 0 END)
               / NULLIF(COUNT(*),0) AS decimal(5,1))      AS PctRegistered,
    MAX(s.FutureSessions)                                 AS FutureSessions,
    MAX(s.NextSessionDate)                                AS NextSessionDate,
    MAX(s.OpenSeats)                                      AS OpenSeats
FROM req r
LEFT JOIN done  d ON d.UniversalID = r.UniversalID AND d.ClassKey = r.ClassKey
LEFT JOIN sched s ON s.ClassKey = r.ClassKey
GROUP BY r.JobRole, r.Track, r.ClassKey;
GO

/* ---------- 32. report.w3_training_progress — one person, all modalities ----------
   Shows      : WAVE 3, our population, in-scope users. One row per person:
                classroom sessions (required / registered / unregistered /
                no-show gaps per Epic; completed per Cornerstone) AND online
                CBT modules (required per Epic 'Online Class' rows; completed
                + average score per the Enterprise export's Online Class rows
                — the two systems' names match 1:1). OverallPctComplete counts
                a requirement done when the session/module is completed.
                Epic's own Fully Registered / Fully Trained flags ride along
                for cross-reference. (Epic 'Test' requirements not yet
                tracked here — the Enterprise export now carries Test rows,
                but needs 'OR Completed Date >= 01/01/2026' added to its date
                window before equivalency test-outs are complete.)
   Built from : report.users + report.training_status + raw.epic_status
                + raw.cornerstone (roster report PAUSED 2026-07-27)
                + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_training_progress AS
WITH sess AS (   -- classroom requirements, deduped person x class
    SELECT UniversalID, ClassKey, RegState,
           CASE WHEN d.ck IS NOT NULL THEN 1 ELSE 0 END AS Done
    FROM (
        SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id)))) AS UniversalID,
               UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))))  AS ClassKey,
               MAX(CASE WHEN UPPER(s.Event_Class_Registered) LIKE 'YES%'    THEN 3
                        WHEN UPPER(s.Event_Class_Registered) LIKE 'EXEMPT%' THEN 2
                        WHEN UPPER(s.Event_Class_Registered) LIKE 'NO -NO SHOW%' THEN 1
                        ELSE 0 END) AS RegState
        FROM raw.epic_status s
        WHERE s.Event_Class_Type = 'Session'
          AND LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
          AND LTRIM(RTRIM(ISNULL(s.Universal_Id,''))) <> ''
        GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id)))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))))
    ) x
    LEFT JOIN (SELECT DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID])))) uid,
                      UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title])))) ck
               FROM raw.cornerstone       -- roster report PAUSED 2026-07-27
               WHERE [Transcript_Status] = 'Completed'
                 AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
                 AND [Training_Type] = 'Session') d
      ON d.uid = x.UniversalID AND d.ck = x.ClassKey
),
onl AS (         -- online requirements, deduped person x module
    SELECT UniversalID, ModuleKey,
           CASE WHEN c.Done IS NOT NULL THEN 1 ELSE 0 END AS Done,
           c.Score
    FROM (
        SELECT DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id)))) AS UniversalID,
               UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))))           AS ModuleKey
        FROM raw.epic_status s
        WHERE s.Event_Class_Type = 'Online Class'
          AND LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
          AND LTRIM(RTRIM(ISNULL(s.Universal_Id,''))) <> ''
    ) x
    LEFT JOIN (SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),User_ID)))) uid,
                      UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Training_Title)))) t,
                      MAX(1) Done,
                      MAX(TRY_CAST(Transcript_Score AS float)) Score
               FROM raw.cornerstone
               WHERE Training_Type = 'Online Class'
                 AND Transcript_Status IN ('Completed','Completed (Equivalent)','Exempt')
               GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),User_ID)))), UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Training_Title))))) c
      ON c.uid = x.UniversalID AND c.t = x.ModuleKey
),
sagg AS (
    SELECT UniversalID,
           COUNT(*)                                        AS SessionsRequired,
           SUM(CASE WHEN RegState >= 2 THEN 1 ELSE 0 END)  AS SessionsRegisteredOrExempt,
           SUM(CASE WHEN RegState = 1 THEN 1 ELSE 0 END)   AS SessionsNoShowGap,
           SUM(CASE WHEN RegState = 0 THEN 1 ELSE 0 END)   AS SessionsUnregistered,
           SUM(Done)                                       AS SessionsCompleted
    FROM sess GROUP BY UniversalID
),
oagg AS (
    SELECT UniversalID,
           COUNT(*)          AS OnlineRequired,
           SUM(Done)         AS OnlineCompleted,
           CAST(AVG(CASE WHEN Score > 0 THEN Score END) AS decimal(5,1)) AS OnlineAvgScore
    FROM onl GROUP BY UniversalID
)
SELECT
    u.UniversalID, u.FullName, u.Leader, u.AVP, u.VP,
    u.JobRole1, u.VendorYN, u.UserType,
    ISNULL(sa.SessionsRequired,0)          AS SessionsRequired,
    ISNULL(sa.SessionsRegisteredOrExempt,0) AS SessionsRegisteredOrExempt,
    ISNULL(sa.SessionsUnregistered,0)      AS SessionsUnregistered,
    ISNULL(sa.SessionsNoShowGap,0)         AS SessionsNoShowGap,
    ISNULL(sa.SessionsCompleted,0)         AS SessionsCompleted,
    ISNULL(oa.OnlineRequired,0)            AS OnlineRequired,
    ISNULL(oa.OnlineCompleted,0)           AS OnlineCompleted,
    oa.OnlineAvgScore,
    CAST(100.0 * (ISNULL(sa.SessionsCompleted,0) + ISNULL(oa.OnlineCompleted,0))
               / NULLIF(ISNULL(sa.SessionsRequired,0) + ISNULL(oa.OnlineRequired,0),0)
         AS decimal(5,1))                  AS OverallPctComplete,
    ISNULL(t.IsFullyRegistered,0)          AS FullyRegisteredPerEpic,
    ISNULL(t.IsFullyTrained,0)             AS FullyTrainedPerEpic
FROM report.users u
LEFT JOIN report.training_status t ON t.UniversalID = u.UniversalID
LEFT JOIN sagg sa ON sa.UniversalID = u.UniversalID
LEFT JOIN oagg oa ON oa.UniversalID = u.UniversalID
WHERE u.Wave = 'Wave 3' AND u.IsInScope = 1
  AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders);
GO

/* ---------- 33. report.w3_watchlist — cross-system issues worth acting on ----------
   Shows      : WAVE 3, our population. One row per person per issue, each
                verified to have real signal on 2026-07-15 (count that day):
                  1. No Epic curriculum assigned            (32) - in scope,
                     needs training, zero Epic status rows: cannot register
                  2. Departed/out-of-scope holds upcoming seat (12) - seat waste
                  3. Epic registered, no Cornerstone transcript (38 classes) -
                     registration exists only in Epic; verify it's real
                  4. Cornerstone registered, Epic not synced (8 classes) -
                     the seat is booked; Epic report lags
                  5. Cornerstone complete, Epic not marked trained (3)
                  6. Epic trained, Cornerstone missing completions (11)
                  7. No-show standing, no future bookings - the chase-downs
                Detail carries the affected classes (capped at 300 chars).
   Built from : report.users + raw.epic_status + raw.cornerstone (roster report PAUSED 2026-07-27)
                + report.w3_noshow_status + report.w3_training_progress
                + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.w3_watchlist AS
WITH pop AS (
    SELECT u.UniversalID, u.FullName, u.Leader, u.IsInScope,
           CASE WHEN UPPER(CAST(u.Departed AS nvarchar(10)))='TRUE' THEN 1 ELSE 0 END AS Dep,
           CASE WHEN UPPER(CAST(u.TerminatedFlag AS nvarchar(10)))='TRUE' THEN 1 ELSE 0 END AS Term
    FROM report.users u
    WHERE u.Wave = 'Wave 3'
      AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
),
cs AS (        -- Cornerstone Enterprise sessions, normalized ONCE (roster
               -- report PAUSED 2026-07-27; EPIC provider = same population)
    SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID]))))            AS uid,
           UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title])))) AS ck,
           MAX([Training_Title])                 AS title,
           CONVERT(nvarchar(50), [Transcript_Status]) AS st,
           TRY_CONVERT(datetime, [Training_Start_Date]) AS startdt
    FROM raw.cornerstone
    WHERE LTRIM(RTRIM(ISNULL([User_ID],''))) <> ''
      AND UPPER(LTRIM(RTRIM(ISNULL([Training_Provider],'')))) = 'EPIC'
      AND [Training_Type] = 'Session'
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),[User_ID])))),
             UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),[Training_Title])))),
             CONVERT(nvarchar(50), [Transcript_Status]),
             TRY_CONVERT(datetime, [Training_Start_Date])
),
epuid AS (     -- everyone Epic's status report knows about
    SELECT DISTINCT UPPER(LTRIM(RTRIM(Universal_Id))) AS uid
    FROM raw.epic_status
    WHERE LTRIM(RTRIM(ISNULL(Universal_Id,''))) <> ''
),
cslive AS (    -- person x class pairs with a live/complete Cornerstone transcript
    SELECT DISTINCT uid, ck FROM cs
    WHERE st IN ('Registered','Registration Pending','Completed','Exempt','Incomplete')
)
-- 1. in scope but Epic doesn't know them
SELECT p.UniversalID, p.FullName, p.Leader,
       'No Epic curriculum assigned' AS Issue,
       'In scope and needs training but has NO rows in the Epic curriculum status report - cannot register or train until Epic assigns a curriculum' AS Detail
FROM pop p
LEFT JOIN epuid e ON e.uid = p.UniversalID
WHERE p.IsInScope = 1 AND e.uid IS NULL
UNION ALL
-- 2. departed / out of scope but still booked into future sessions
SELECT p.UniversalID, p.FullName, p.Leader,
       'Departed/out-of-scope holds upcoming seat',
       LEFT('Upcoming: ' + STRING_AGG(CONVERT(varchar(200),
            c.title + ' (' + CONVERT(varchar(10), c.startdt, 120) + ')'), '; '), 300)
FROM pop p
JOIN cs c ON c.uid = p.UniversalID
WHERE (p.Dep = 1 OR p.Term = 1 OR p.IsInScope = 0)
  AND c.st IN ('Registered','Registration Pending')
  AND c.startdt >= CAST(GETDATE() AS date)
GROUP BY p.UniversalID, p.FullName, p.Leader
UNION ALL
-- 3. Epic says registered but Cornerstone has no live transcript for the class
SELECT e.UniversalID, p.FullName, p.Leader,
       'Epic registered, no Cornerstone transcript',
       LEFT(CAST(COUNT(*) AS varchar(10)) + ' class(es): '
            + STRING_AGG(CONVERT(varchar(120), e.ClassKey), '; '), 300)
FROM (SELECT DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id)))) AS UniversalID,
             UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class)))) AS ClassKey
      FROM raw.epic_status s
      WHERE s.Event_Class_Type = 'Session'
        AND UPPER(s.Event_Class_Registered) LIKE 'YES%') e
JOIN pop p ON p.UniversalID = e.UniversalID AND p.IsInScope = 1
LEFT JOIN cslive l ON l.uid = e.UniversalID AND l.ck = e.ClassKey
WHERE l.uid IS NULL
GROUP BY e.UniversalID, p.FullName, p.Leader
UNION ALL
-- 4. Cornerstone has the seat but Epic still reports unregistered
SELECT wu.UniversalID, MAX(wu.FullName), MAX(wu.Leader),
       'Cornerstone registered, Epic not synced',
       LEFT(CAST(COUNT(DISTINCT wu.Event_Class) AS varchar(10)) + ' class(es): '
            + STRING_AGG(CONVERT(varchar(120), wu.Event_Class), '; '), 300)
FROM report.w3_unregistered wu
JOIN (SELECT DISTINCT uid, ck FROM cs
      WHERE st IN ('Registered','Registration Pending','Completed')) l
  ON l.uid = wu.UniversalID AND l.ck = UPPER(LTRIM(RTRIM(wu.Event_Class)))
WHERE wu.IsInScope = 1
GROUP BY wu.UniversalID
UNION ALL
-- 5/6. trained-flag disagreements between the systems
SELECT tp.UniversalID, tp.FullName, tp.Leader,
       CASE WHEN tp.FullyTrainedPerEpic = 0
            THEN 'Cornerstone complete, Epic not marked trained'
            ELSE 'Epic trained, Cornerstone missing completions' END,
       CONCAT('Cornerstone sessions completed ', tp.SessionsCompleted, ' of ',
              tp.SessionsRequired, '; Epic Fully Trained = ',
              CASE WHEN tp.FullyTrainedPerEpic = 1 THEN 'Yes' ELSE 'No' END)
FROM report.w3_training_progress tp
WHERE tp.SessionsRequired > 0
  AND ((tp.SessionsCompleted = tp.SessionsRequired AND tp.FullyTrainedPerEpic = 0)
    OR (tp.SessionsCompleted < tp.SessionsRequired AND tp.FullyTrainedPerEpic = 1))
UNION ALL
-- 7. standing no-shows with nothing else booked (the true chase-downs)
SELECT n.UniversalID, n.FullName, n.Leader,
       'No-show standing, no future bookings',
       LEFT('No-showed ' + n.ClassTitle + ' on '
            + CONVERT(varchar(10), n.NoShowedOn, 120)
            + ' and has no upcoming registrations in any class', 300)
FROM report.w3_noshow_narrative n
WHERE n.Resolution = 'No Show standing' AND n.UpcomingOtherClasses = 0;
GO

/* ---------- 33b. report.tracker_detail — RCM Training Tracker person x event detail ----------
   Shows      : our population (in-scope wave-file people whose Leader is on
                the canonical list), ALL waves — one row per Epic curriculum
                status export row (every event type), with typed dates/numbers
                and an IsExcluded flag for the optional-workshop classes the
                tracker leaves out of metrics (Advanced Reporting / Charge
                Capture). Wave filtering happens in the consumer.
   Built from : raw.epic_status + report.users + raw.ref_leaders
   Used by    : scripts\build_rcm_training_tracker.py (the Python-built
                "RCM Training Tracker.xlsx" — Data sheet, leader cube, trend)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.tracker_detail AS
SELECT
    u.UniversalID,
    s.Team_Member                                   AS FullName,
    u.Wave, u.Leader, u.AVP, u.VP, u.UserType, u.VendorYN,
    s.Direct_Manager,
    s.Curriculum_type                               AS CurriculumType,
    s.Fully_Trained, s.Fully_Registered,
    s.Facility_Serviceline, s.Department, s.Union_Description,
    s.Epic_Job_Category,
    s.Curriculum, s.Curriculum_Status,
    TRY_CAST(TRY_CAST(s.Sequence AS float) AS int)  AS Sequence,
    s.Event_Class, s.Event_Class_Type,
    s.Event_Class_Registered, s.Event_Class_Status,
    TRY_CAST(s.Event_Class_Duration_Hours AS float) AS DurationHours,
    TRY_CONVERT(datetime, s.Event_Class_Date)       AS EventDate,
    s.Event_Class_Location, s.Event_Class_Room,
    s.Senior_Registration_Coordinator,
    TRY_CONVERT(datetime, s.Registration_Date)      AS RegistrationDate,
    CASE WHEN UPPER(s.Event_Class) LIKE '%ADVANCED REPORTING%'
           OR UPPER(s.Event_Class) LIKE '%CHARGE CAPTURE%'
           OR UPPER(s.Curriculum)  LIKE '%ADVANCED REPORTING%'
           OR UPPER(s.Curriculum)  LIKE '%CHARGE CAPTURE%'
         THEN 1 ELSE 0 END                          AS IsExcluded
FROM raw.epic_status s
JOIN report.users u ON u.UniversalID = UPPER(LTRIM(RTRIM(s.Universal_Id)))
WHERE u.IsInScope = 1
  AND u.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders);
GO

/* ---------- 33c. report.tracker_class_schedule — typed Epic class schedule ----------
   Shows      : one row per scheduled session from the Epic class schedule
                export, with typed date and seat counts. Drops the export's
                trailing "Total" and "Applied filters" junk rows. The export
                is already filtered upstream to the relevant curricula.
   Built from : raw.epic_class_schedule
   Used by    : scripts\build_rcm_training_tracker.py (Registration Schedule,
                Classes to Schedule, Reg Recommendations sheets)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.tracker_class_schedule AS
SELECT
    TRY_CONVERT(date, cs.[Date])                        AS SessionDate,
    LTRIM(RTRIM(cs.Training_Location))                  AS Location,
    LTRIM(RTRIM(ISNULL(cs.Room,'')))                    AS Room,
    cs.Session_Type, cs.Primary_Instructor, cs.Secondary_Instructor,
    LTRIM(RTRIM(cs.Event_Name))                         AS EventName,
    cs.Class_Event_Start_Time                           AS StartTime,
    cs.Class_Event_End_Time                             AS EndTime,
    cs.Event_Type,
    TRY_CAST(TRY_CAST(cs.Total_Seats AS float) AS int)       AS TotalSeats,
    TRY_CAST(TRY_CAST(cs.Total_Seats_Taken AS float) AS int) AS SeatsTaken,
    TRY_CAST(TRY_CAST(cs.Available_Seats AS float) AS int)   AS AvailableSeats
FROM raw.epic_class_schedule cs
WHERE TRY_CONVERT(date, cs.[Date]) IS NOT NULL
  AND LTRIM(RTRIM(ISNULL(cs.Event_Name,''))) <> '';
GO

/* ---------- 33d. report.tracker_daily_log — daily metric log per leader ----------
   Shows      : one row per snapshot day per leader (our population, leader on
                the canonical list) — Members / FullyRegistered / FullyTrained
                from history.roster_daily (captured since 2026-07-06), plus
                Unregistered and No-Show counts from history.w3_leader_daily
                (captured since 2026-07-15; NULL before that — those columns
                simply weren't tracked earlier). Feeds the tracker's Daily Log
                sheet; day-over-day deltas are computed by the builder.
   Built from : history.roster_daily + history.w3_leader_daily + raw.ref_leaders
   Used by    : scripts\build_rcm_training_tracker.py
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.tracker_daily_log AS
WITH r AS (
    SELECT TRY_CONVERT(date, SnapshotDate)  AS SnapshotDate,
           Wave,
           LTRIM(RTRIM(Leader))             AS Leader,
           COUNT(*)                         AS Members,
           SUM(CASE WHEN TRY_CAST(IsFullyRegistered AS float) = 1
                    THEN 1 ELSE 0 END)      AS FullyRegistered,
           SUM(CASE WHEN TRY_CAST(IsFullyTrained AS float) = 1
                    THEN 1 ELSE 0 END)      AS FullyTrained
    FROM history.roster_daily
    WHERE TRY_CAST(IsInScope AS float) = 1
      AND LTRIM(RTRIM(ISNULL(Leader,''))) IN
          (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)
    GROUP BY TRY_CONVERT(date, SnapshotDate), Wave, LTRIM(RTRIM(Leader))
)
SELECT r.SnapshotDate, r.Wave, r.Leader,
       r.Members, r.FullyRegistered, r.FullyTrained,
       w.UnregisteredUsers, w.UnregisteredSessions,
       w.NoShowUsersStanding, w.NoShowSessionsStanding,
       w.NoShowUsersToDate, w.NoShowSessionsToDate
FROM r
LEFT JOIN history.w3_leader_daily w
       ON TRY_CONVERT(date, w.SnapshotDate) = r.SnapshotDate
      AND LTRIM(RTRIM(w.Leader)) = r.Leader
      AND r.Wave = 'Wave 3';
GO

/* ---------- 33e. report.tracker_training_status — per-learner status ----------
   Shows      : one row per in-scope wave person: training status bucket
                (Not Started / In Progress / Fully Trained / No Epic
                Curriculum), session counts, LastAttendedDate,
                FinalScheduledDate (their projected role-completion date),
                open unregistered-class count, plus completion-status fields
                (2026-07-29): Epic Fully Registered / Fully Trained Yes-No
                flags, the NAME of the last booked class, CompletionDate
                (last attended date once fully trained), distinct classes
                registered, and pending-attendance count (registered but the
                session date is still ahead). Distinct classes only —
                Epic repeats a class under every curriculum requiring it.
   Built from : report.users + report.tracker_detail
   Used by    : tracker "Training Status" + "Completion Status" sheets;
                leadership pulls
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.tracker_training_status AS
WITH agg AS (
    SELECT UniversalID,
           MAX(CASE WHEN Fully_Trained LIKE '%Yes%' THEN 1 ELSE 0 END) AS FT,
           MAX(CASE WHEN Fully_Registered LIKE '%Yes%' THEN 1 ELSE 0 END) AS FR,
           COUNT(DISTINCT CASE WHEN Event_Class_Type = 'Session'
                 AND IsExcluded = 0 THEN Event_Class END)           AS ClassesRequired,
           /* 'Completed (Equivalent)' = equivalency credit (no session date)
              — counts as completed, same as a plain 'Completed'. */
           COUNT(DISTINCT CASE WHEN Event_Class_Type = 'Session'
                 AND IsExcluded = 0 AND Event_Class_Status LIKE 'Completed%'
                 THEN Event_Class END)                              AS ClassesCompleted,
           COUNT(DISTINCT CASE WHEN Event_Class_Type = 'Session'
                 AND IsExcluded = 0
                 AND (Event_Class_Registered LIKE '%Yes%'
                      OR Event_Class_Status LIKE 'Completed%')
                 THEN Event_Class END)                              AS ClassesRegistered,
           MAX(CASE WHEN Event_Class_Type = 'Session' AND IsExcluded = 0
                 AND Event_Class_Status LIKE 'Completed%'
                 THEN EventDate END)                                AS LastAttendedDate,
           MAX(CASE WHEN Event_Class_Type = 'Session' AND IsExcluded = 0
                 AND (Event_Class_Registered LIKE '%Yes%'
                      OR Event_Class_Status LIKE 'Completed%')
                 THEN EventDate END)                                AS FinalScheduledDate,
           COUNT(DISTINCT CASE WHEN Event_Class_Type = 'Session'
                 AND IsExcluded = 0 AND Event_Class_Registered LIKE '%No%'
                 AND Event_Class_Registered NOT LIKE '%Show%'
                 THEN Event_Class END)                              AS UnregisteredClasses
    FROM report.tracker_detail
    GROUP BY UniversalID
),
lastclass AS (
    /* Dated rows win (newest first); date-less rows (equivalency credit)
       still yield a class name instead of a blank. */
    SELECT UniversalID, Event_Class AS LastRegisteredClass
    FROM (
        SELECT UniversalID, Event_Class,
               ROW_NUMBER() OVER (PARTITION BY UniversalID
                                  ORDER BY CASE WHEN EventDate IS NULL
                                                THEN 1 ELSE 0 END,
                                           EventDate DESC, Event_Class) AS rn
        FROM report.tracker_detail
        WHERE Event_Class_Type = 'Session' AND IsExcluded = 0
          AND (Event_Class_Registered LIKE '%Yes%'
               OR Event_Class_Status LIKE 'Completed%')
    ) x
    WHERE rn = 1
)
SELECT u.UniversalID, u.FullName, u.Wave, u.Leader, u.AVP, u.VP,
       u.UserType, u.VendorYN, u.OnLeaderList,
       CASE WHEN a.UniversalID IS NULL THEN 'No Epic Curriculum'
            WHEN a.FT = 1              THEN 'Fully Trained'
            WHEN a.ClassesCompleted > 0 THEN 'In Progress'
            ELSE 'Not Started' END        AS TrainingStatus,
       CASE WHEN a.FR = 1 THEN 'Yes' ELSE 'No' END AS FullyRegisteredYN,
       CASE WHEN a.FT = 1 THEN 'Yes' ELSE 'No' END AS FullyTrainedYN,
       lc.LastRegisteredClass,
       ISNULL(a.ClassesRequired, 0)       AS ClassesRequired,
       ISNULL(a.ClassesCompleted, 0)      AS ClassesCompleted,
       ISNULL(a.ClassesRegistered, 0)     AS ClassesRegistered,
       a.LastAttendedDate,
       a.FinalScheduledDate,
       CASE WHEN a.FT = 1 THEN a.LastAttendedDate END AS CompletionDate,
       ISNULL(a.UnregisteredClasses, 0)   AS UnregisteredClasses,
       ISNULL(a.ClassesRegistered, 0)
         - ISNULL(a.ClassesCompleted, 0)  AS PendingAttendanceClasses
FROM report.users u
LEFT JOIN agg a ON a.UniversalID = u.UniversalID
LEFT JOIN lastclass lc ON lc.UniversalID = u.UniversalID
WHERE u.IsInScope = 1;
GO

/* ---------- 33f. report.gnf_assignments — Go-and-Find module assignments ----------
   Shows      : one row per wave person x assigned Go-and-Find module, with
                Leader / role / track context. Assignment mapping only —
                completion data not yet available (cadence TBD; drop new
                mapping files in raw\wave\gnf_mapping, newest loads).
   Built from : raw.wave_gnf_w3_assignments + raw.wave_gnf_w3_mapping
                + report.users
   Used by    : tracker "Go and Finds" sheet
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.gnf_assignments AS
WITH ctx AS (
    SELECT UPPER(LTRIM(RTRIM(UID))) AS uid,
           MAX(USER_ROLE) AS UserRole, MAX(BUS_UNIT) AS BusinessUnit
    FROM raw.wave_gnf_w3_mapping
    GROUP BY UPPER(LTRIM(RTRIM(UID)))
)
SELECT u.UniversalID, u.FullName, u.Wave, u.Leader, u.OnLeaderList,
       c.UserRole, c.BusinessUnit,
       LTRIM(RTRIM(g.GO_AND_FIND))            AS GoAndFind,
       TRY_CAST(g.GNF_COUNT AS int)           AS UserGnfCount
FROM raw.wave_gnf_w3_assignments g
JOIN report.users u
  ON u.UniversalID = UPPER(LTRIM(RTRIM(g.UID))) AND u.IsInScope = 1
LEFT JOIN ctx c ON c.uid = u.UniversalID
WHERE LTRIM(RTRIM(ISNULL(g.GO_AND_FIND, ''))) <> '';
GO

/* ---------- 34. report.whats_what — the in-database catalog ----------
   Shows      : one row per table/view in this database — its description,
                what it reads, and column count. New to the database?
                Start with: SELECT * FROM report.whats_what ORDER BY ObjectName
   Built from : the system catalog (sys.objects / sys.extended_properties /
                sys.sql_expression_dependencies). Descriptions are synced
                into MS_Description from these view headers by
                sql\build_data_dictionary.py — edit the header, not the property.
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.whats_what AS
SELECT
    s.name + '.' + o.name AS ObjectName,
    CASE o.type WHEN 'U' THEN 'table' ELSE 'view' END AS Type,
    CAST(ep.value AS nvarchar(1000)) AS Description,
    (SELECT STRING_AGG(x.ref, ', ')
     FROM (SELECT DISTINCT ISNULL(d.referenced_schema_name, 'dbo')
                  + '.' + d.referenced_entity_name AS ref
           FROM sys.sql_expression_dependencies d
           WHERE d.referencing_id = o.object_id
             AND d.referenced_id IS NOT NULL) x) AS Reads,
    (SELECT COUNT(*) FROM sys.columns c WHERE c.object_id = o.object_id) AS ColumnCount
FROM sys.objects o
JOIN sys.schemas s ON s.schema_id = o.schema_id
LEFT JOIN sys.extended_properties ep
       ON ep.class = 1 AND ep.major_id = o.object_id AND ep.minor_id = 0
      AND ep.name = 'MS_Description'
WHERE o.type IN ('U', 'V') AND o.is_ms_shipped = 0;
GO

/* ---------- 35. report.role_track_class_map — job category -> template -> track -> class ----------
   Shows      : the reference mapping (no leaders, no counts): one row per
                Job Category x Training Track x Event Class (session).
                Carries the MVP linkable template (ID + name), the job
                category group, and the MVP description. Limited to job
                categories actually held by wave-file users (JobRole1-4);
                ClassTitle is NULL where Epic has no session rows for the
                track yet (or the track is 'No Training').
   Built from : raw.mvp_job_categories (TrainingTrack1-6 unpivoted)
                + raw.epic_status (Curriculum -> Event_Class, sessions only)
                + report.users (population filter)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.role_track_class_map AS
WITH pop_roles AS (   -- job categories actually held by wave-file users
    SELECT DISTINCT UPPER(LTRIM(RTRIM(jr.JobRole))) AS JobRoleKey
    FROM report.users u
    CROSS APPLY (VALUES (u.JobRole1),(u.JobRole2),(u.JobRole3),(u.JobRole4)) jr(JobRole)
    WHERE LTRIM(RTRIM(ISNULL(jr.JobRole,''))) <> ''
),
cat AS (              -- one row per job category x training-track slot
    SELECT
        jc.JobCategoryGroup,
        jc.JobCategoryID,
        jc.JobCategoryName,
        UPPER(LTRIM(RTRIM(jc.JobCategoryName)))  AS JobRoleKey,
        jc.LinkableTemplateID,
        jc.LinkableTemplateName,
        jc.Description,
        t.TrackSlot,
        LTRIM(RTRIM(t.Track))                    AS TrainingTrack,
        UPPER(LTRIM(RTRIM(t.Track)))             AS TrackKey
    FROM raw.mvp_job_categories jc
    CROSS APPLY (VALUES (1, jc.TrainingTrack1),(2, jc.TrainingTrack2),
                        (3, jc.TrainingTrack3),(4, jc.TrainingTrack4),
                        (5, jc.TrainingTrack5),(6, jc.TrainingTrack6)) t(TrackSlot, Track)
    WHERE LTRIM(RTRIM(ISNULL(t.Track,''))) <> ''
),
cls AS (              -- distinct track -> class (sessions) per Epic
    SELECT DISTINCT
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200), s.Curriculum))))  AS TrackKey,
        LTRIM(RTRIM(CONVERT(nvarchar(200), s.Event_Class)))        AS ClassTitle
    FROM raw.epic_status s
    WHERE s.Event_Class_Type = 'Session'
      AND LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
)
SELECT
    c.JobCategoryGroup,
    c.JobCategoryID,
    c.JobCategoryName,
    c.LinkableTemplateID,
    c.LinkableTemplateName,
    c.TrackSlot,
    c.TrainingTrack,
    cl.ClassTitle,
    c.Description
FROM cat c
LEFT JOIN cls cl ON cl.TrackKey = c.TrackKey
WHERE c.JobRoleKey IN (SELECT JobRoleKey FROM pop_roles);
GO

/* ---------- 36. report.leader_role_training_map — leader rollup of the training map ----------
   Shows      : our population, in-scope + on-leader-list users. One row per
                Leader x Wave x Job Category x Training Track x Event Class,
                with the MVP linkable template (ID + name), job category
                group, UserCount = distinct users under that leader/wave
                holding the job category (JobRole1-4), and the MVP
                description. Wave NULL shows as 'No Wave'; ClassTitle NULL =
                no Epic session rows for the track (or 'No Training').
   Built from : report.users (JobRole1-4 unpivoted)
                + raw.mvp_job_categories (TrainingTrack1-6 unpivoted)
                + raw.epic_status (Curriculum -> Event_Class, sessions only)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.leader_role_training_map AS
WITH user_roles AS (  -- one row per in-scope user x held job category
    SELECT DISTINCT
        u.UniversalID,
        u.Leader,
        ISNULL(NULLIF(LTRIM(RTRIM(u.Wave)),''),'No Wave') AS Wave,
        UPPER(LTRIM(RTRIM(jr.JobRole)))                   AS JobRoleKey
    FROM report.users u
    CROSS APPLY (VALUES (u.JobRole1),(u.JobRole2),(u.JobRole3),(u.JobRole4)) jr(JobRole)
    WHERE u.IsInScope = 1 AND u.OnLeaderList = 1
      AND LTRIM(RTRIM(ISNULL(jr.JobRole,''))) <> ''
),
cat AS (              -- one row per job category x training-track slot
    SELECT
        jc.JobCategoryGroup,
        jc.JobCategoryID,
        jc.JobCategoryName,
        UPPER(LTRIM(RTRIM(jc.JobCategoryName)))  AS JobRoleKey,
        jc.LinkableTemplateID,
        jc.LinkableTemplateName,
        jc.Description,
        LTRIM(RTRIM(t.Track))                    AS TrainingTrack,
        UPPER(LTRIM(RTRIM(t.Track)))             AS TrackKey
    FROM raw.mvp_job_categories jc
    CROSS APPLY (VALUES (1, jc.TrainingTrack1),(2, jc.TrainingTrack2),
                        (3, jc.TrainingTrack3),(4, jc.TrainingTrack4),
                        (5, jc.TrainingTrack5),(6, jc.TrainingTrack6)) t(TrackSlot, Track)
    WHERE LTRIM(RTRIM(ISNULL(t.Track,''))) <> ''
),
cls AS (              -- distinct track -> class (sessions) per Epic
    SELECT DISTINCT
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200), s.Curriculum))))  AS TrackKey,
        LTRIM(RTRIM(CONVERT(nvarchar(200), s.Event_Class)))        AS ClassTitle
    FROM raw.epic_status s
    WHERE s.Event_Class_Type = 'Session'
      AND LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
)
SELECT
    ur.Leader,
    ur.Wave,
    c.JobCategoryGroup,
    c.JobCategoryID,
    c.JobCategoryName,
    c.LinkableTemplateID,
    c.LinkableTemplateName,
    c.TrainingTrack,
    cl.ClassTitle,
    COUNT(DISTINCT ur.UniversalID) AS UserCount,
    MAX(c.Description)             AS Description
FROM user_roles ur
JOIN cat c  ON c.JobRoleKey = ur.JobRoleKey
LEFT JOIN cls cl ON cl.TrackKey = c.TrackKey
GROUP BY ur.Leader, ur.Wave, c.JobCategoryGroup, c.JobCategoryID, c.JobCategoryName,
         c.LinkableTemplateID, c.LinkableTemplateName, c.TrainingTrack, cl.ClassTitle;
GO

/* ============================================================
   37-40: the RCM training-catalog dimensions + map.
   Three dimension views (job categories, training tracks, classes) built
   from every source that knows about them, cross-checked with source
   flags, then one map view chaining category -> template -> track -> class.
   Scope everywhere: people under the canonical leaders (OnLeaderList = 1).
   Naming eras are normalized: EPIC_W2_ / EPIC_W3_ prefixes -> EPIC_NH_
   (same track/class, renamed between waves), so W2-era Cornerstone history
   lines up with current NH names.
   ============================================================ */

/* ---------- 37. report.rcm_job_categories — dimension: all RCM job categories ----------
   Shows      : one row per job category relevant to our leaders — held by an
                in-scope user under a canonical leader (Master JobRole1-4)
                OR assigned in the Epic TM Lookup (Epic_Job_Categories_MVP).
                Carries the MVP template + group + description, source flags
                (OnMasterRoles / InTMLookup / InMVPCatalog), and user counts
                per source. InMVPCatalog = 0 is a cross-check finding: a
                category in use that the MVP Job Categories export lacks.
   Built from : report.users + raw.epic_lookup + raw.mvp_job_categories
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.rcm_job_categories AS
WITH pop AS (
    SELECT UniversalID FROM report.users WHERE IsInScope = 1 AND OnLeaderList = 1
),
master_roles AS (   -- category as held on the Master (JobRole1-4)
    SELECT UPPER(LTRIM(RTRIM(jr.JobRole))) AS CatKey,
           MIN(LTRIM(RTRIM(jr.JobRole)))   AS CatName,
           COUNT(DISTINCT u.UniversalID)   AS MasterUserCount
    FROM report.users u
    JOIN pop p ON p.UniversalID = u.UniversalID
    CROSS APPLY (VALUES (u.JobRole1),(u.JobRole2),(u.JobRole3),(u.JobRole4)) jr(JobRole)
    WHERE LTRIM(RTRIM(ISNULL(jr.JobRole,''))) <> ''
    GROUP BY UPPER(LTRIM(RTRIM(jr.JobRole)))
),
tm_roles AS (       -- category as assigned in the Epic TM Lookup
    SELECT UPPER(LTRIM(RTRIM(x.value)))    AS CatKey,
           MIN(LTRIM(RTRIM(x.value)))      AS CatName,
           COUNT(DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),el.Universal_ID))))) AS TMLookupUserCount
    FROM raw.epic_lookup el
    JOIN pop p ON p.UniversalID = UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),el.Universal_ID))))
    CROSS APPLY STRING_SPLIT(CONVERT(nvarchar(max),el.Epic_Job_Categories_MVP),'|') x
    WHERE LTRIM(RTRIM(x.value)) <> ''
    GROUP BY UPPER(LTRIM(RTRIM(x.value)))
),
all_cats AS (
    SELECT ISNULL(m.CatKey, t.CatKey)   AS CatKey,
           ISNULL(m.CatName, t.CatName) AS CatName,
           ISNULL(m.MasterUserCount,0)  AS MasterUserCount,
           ISNULL(t.TMLookupUserCount,0) AS TMLookupUserCount
    FROM master_roles m
    FULL OUTER JOIN tm_roles t ON t.CatKey = m.CatKey
)
SELECT
    jc.JobCategoryGroup,
    jc.JobCategoryID,
    ISNULL(jc.JobCategoryName, ac.CatName)          AS JobCategoryName,
    jc.LinkableTemplateID,
    jc.LinkableTemplateName,
    CASE WHEN jc.JobCategoryName IS NULL THEN 0 ELSE 1 END AS InMVPCatalog,
    CASE WHEN ac.MasterUserCount   > 0 THEN 1 ELSE 0 END   AS OnMasterRoles,
    CASE WHEN ac.TMLookupUserCount > 0 THEN 1 ELSE 0 END   AS InTMLookup,
    ac.MasterUserCount,
    ac.TMLookupUserCount,
    jc.Description
FROM all_cats ac
LEFT JOIN raw.mvp_job_categories jc
       ON UPPER(LTRIM(RTRIM(jc.JobCategoryName))) = ac.CatKey;
GO

/* ---------- 38. report.rcm_training_tracks — dimension: all RCM training tracks ----------
   Shows      : one row per normalized training track (curriculum) relevant
                to our leaders, from any of three sources: the MVP Job
                Categories tracks (TrainingTrack1-6 of an RCM category),
                the TM Lookup per-person Epic_Curriculums_Cornerstone list
                (covers ALL waves), or the Epic Curriculum Status export
                (W3 only). Source flags + per-source counts cross-check
                coverage; ClassCount = distinct W3-statused classes.
   Built from : report.rcm_job_categories + raw.mvp_job_categories
                + raw.epic_lookup + raw.epic_status + report.users
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.rcm_training_tracks AS
WITH pop AS (
    SELECT UniversalID FROM report.users WHERE OnLeaderList = 1
),
mvp_tracks AS (     -- tracks of RCM job categories, era-normalized
    SELECT UPPER(nrm.Track) AS TrackKey, MIN(nrm.Track) AS TrackName,
           COUNT(DISTINCT jc.JobCategoryName) AS CategoryCount
    FROM raw.mvp_job_categories jc
    JOIN report.rcm_job_categories rc
      ON UPPER(LTRIM(RTRIM(rc.JobCategoryName))) = UPPER(LTRIM(RTRIM(jc.JobCategoryName)))
    CROSS APPLY (VALUES (jc.TrainingTrack1),(jc.TrainingTrack2),(jc.TrainingTrack3),
                        (jc.TrainingTrack4),(jc.TrainingTrack5),(jc.TrainingTrack6)) t(Track)
    CROSS APPLY (SELECT REPLACE(REPLACE(LTRIM(RTRIM(t.Track)),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Track) nrm
    WHERE LTRIM(RTRIM(ISNULL(t.Track,''))) <> ''
    GROUP BY UPPER(nrm.Track)
),
tm_tracks AS (      -- tracks assigned per person in the TM Lookup (all waves)
    SELECT UPPER(nrm.Track) AS TrackKey, MIN(nrm.Track) AS TrackName,
           COUNT(DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),el.Universal_ID))))) AS TMLookupUserCount
    FROM raw.epic_lookup el
    JOIN pop p ON p.UniversalID = UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),el.Universal_ID))))
    CROSS APPLY STRING_SPLIT(CONVERT(nvarchar(max),el.Epic_Curriculums_Cornerstone),'|') x
    CROSS APPLY (SELECT REPLACE(REPLACE(LTRIM(RTRIM(x.value)),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Track) nrm
    WHERE LTRIM(RTRIM(x.value)) <> ''
    GROUP BY UPPER(nrm.Track)
),
es_tracks AS (      -- tracks statused in Epic Curriculum Status (W3, our pop)
    SELECT UPPER(nrm.Track) AS TrackKey, MIN(nrm.Track) AS TrackName,
           COUNT(DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))) AS EpicStatusUserCount,
           COUNT(DISTINCT CASE WHEN LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
                 THEN UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class)))) END) AS ClassCount
    FROM raw.epic_status s
    JOIN pop p ON p.UniversalID = UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))
    CROSS APPLY (SELECT REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Curriculum))),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Track) nrm
    WHERE LTRIM(RTRIM(ISNULL(s.Curriculum,''))) <> ''
    GROUP BY UPPER(nrm.Track)
),
all_tracks AS (
    SELECT TrackKey FROM mvp_tracks
    UNION SELECT TrackKey FROM tm_tracks
    UNION SELECT TrackKey FROM es_tracks
)
SELECT
    COALESCE(m.TrackName, t.TrackName, e.TrackName) AS TrainingTrack,
    CASE WHEN m.TrackKey IS NOT NULL THEN 1 ELSE 0 END AS InMVPCategories,
    CASE WHEN t.TrackKey IS NOT NULL THEN 1 ELSE 0 END AS InTMLookup,
    CASE WHEN e.TrackKey IS NOT NULL THEN 1 ELSE 0 END AS InEpicStatus,
    ISNULL(m.CategoryCount,0)      AS CategoryCount,
    ISNULL(t.TMLookupUserCount,0)  AS TMLookupUserCount,
    ISNULL(e.EpicStatusUserCount,0) AS EpicStatusUserCount,
    ISNULL(e.ClassCount,0)         AS ClassCount
FROM all_tracks a
LEFT JOIN mvp_tracks m ON m.TrackKey = a.TrackKey
LEFT JOIN tm_tracks  t ON t.TrackKey = a.TrackKey
LEFT JOIN es_tracks  e ON e.TrackKey = a.TrackKey;
GO

/* ---------- 39. report.rcm_classes — dimension: all RCM classes ----------
   Shows      : one row per normalized class x type relevant to our leaders,
                from any of three sources: Epic Curriculum Status (W3 — the
                only source that links class -> track), Cornerstone
                Enterprise transcripts (Provider = EPIC; full history, so
                W1/W2 era classes live here), or the Epic Class Schedule.
                TrackCount > 0 means the class is track-mapped via W3 data;
                InCornerstone-only rows are the W2-era catalog.
   Built from : raw.epic_status + raw.cornerstone + raw.epic_class_schedule
                + report.users
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.rcm_classes AS
WITH pop AS (
    SELECT UniversalID FROM report.users WHERE OnLeaderList = 1
),
es_cls AS (         -- Epic Curriculum Status: class + type + track link (W3)
    SELECT UPPER(nrm.Cls) AS ClsKey, MIN(nrm.Cls) AS ClsName,
           MIN(CONVERT(nvarchar(50),s.Event_Class_Type)) AS ClassType,
           COUNT(DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))) AS EpicStatusUserCount,
           COUNT(DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Curriculum)))))  AS TrackCount
    FROM raw.epic_status s
    JOIN pop p ON p.UniversalID = UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))
    CROSS APPLY (SELECT REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Cls) nrm
    WHERE LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
    GROUP BY UPPER(nrm.Cls)
),
cs_cls AS (         -- Cornerstone Enterprise: what our people actually took
    SELECT UPPER(nrm.Cls) AS ClsKey, MIN(nrm.Cls) AS ClsName,
           MIN(CONVERT(nvarchar(50),c.Training_Type)) AS ClassType,
           COUNT(DISTINCT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),c.[User_ID]))))) AS CornerstoneUserCount
    FROM raw.cornerstone c
    JOIN pop p ON p.UniversalID = UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),c.[User_ID]))))
    CROSS APPLY (SELECT REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),c.Training_Title))),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Cls) nrm
    WHERE UPPER(LTRIM(RTRIM(ISNULL(c.Training_Provider,'')))) = 'EPIC'
      AND LTRIM(RTRIM(ISNULL(c.Training_Title,''))) <> ''
    GROUP BY UPPER(nrm.Cls)
),
sch_cls AS (        -- Epic Class Schedule: sessions offered
    SELECT UPPER(nrm.Cls) AS ClsKey, MIN(nrm.Cls) AS ClsName,
           COUNT(*) AS ScheduledSessions,
           MIN(CASE WHEN TRY_CONVERT(date,cs.[Date]) >= CAST(GETDATE() AS date)
                    THEN TRY_CONVERT(date,cs.[Date]) END) AS NextSessionDate
    FROM raw.epic_class_schedule cs
    CROSS APPLY (SELECT REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),cs.Event_Name))),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Cls) nrm
    WHERE LTRIM(RTRIM(ISNULL(cs.Event_Name,''))) <> ''
    GROUP BY UPPER(nrm.Cls)
),
all_cls AS (
    SELECT ClsKey FROM es_cls
    UNION SELECT ClsKey FROM cs_cls
    UNION SELECT ClsKey FROM sch_cls
)
SELECT
    COALESCE(e.ClsName, c.ClsName, s.ClsName) AS ClassTitle,
    COALESCE(e.ClassType, c.ClassType,
             CASE WHEN s.ClsKey IS NOT NULL THEN 'Session' END) AS ClassType,
    CASE WHEN e.ClsKey IS NOT NULL THEN 1 ELSE 0 END AS InEpicStatus,
    CASE WHEN c.ClsKey IS NOT NULL THEN 1 ELSE 0 END AS InCornerstone,
    CASE WHEN s.ClsKey IS NOT NULL THEN 1 ELSE 0 END AS OnClassSchedule,
    ISNULL(e.TrackCount,0)           AS TrackCount,
    ISNULL(e.EpicStatusUserCount,0)  AS EpicStatusUserCount,
    ISNULL(c.CornerstoneUserCount,0) AS CornerstoneUserCount,
    ISNULL(s.ScheduledSessions,0)    AS ScheduledSessions,
    s.NextSessionDate
FROM all_cls a
LEFT JOIN es_cls  e ON e.ClsKey = a.ClsKey
LEFT JOIN cs_cls  c ON c.ClsKey = a.ClsKey
LEFT JOIN sch_cls s ON s.ClsKey = a.ClsKey;
GO

/* ---------- 40. report.rcm_training_map — job category -> template -> track -> class ----------
   Shows      : the full chain over the three dimensions above: one row per
                RCM Job Category x Training Track x Class. Category -> track
                comes from MVP (TrainingTrack1-6, era-normalized); track ->
                class comes from Epic Curriculum Status (the only source with
                that link; W3 people, but class names are era-normalized so
                they match Cornerstone W2 history). MappingStatus: 'Mapped' /
                'No Training' / 'No class data' (track absent from the W3
                status export). ClassInCornerstone = 1 cross-checks that our
                people have transcripts for the class.
   Built from : report.rcm_job_categories + raw.mvp_job_categories
                + raw.epic_status + report.users + report.rcm_classes
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.rcm_training_map AS
WITH cat_tracks AS ( -- category -> track (MVP authority), era-normalized
    SELECT rc.JobCategoryGroup, rc.JobCategoryID, rc.JobCategoryName,
           rc.LinkableTemplateID, rc.LinkableTemplateName, rc.Description,
           t.TrackSlot,
           nrm.Track          AS TrainingTrack,
           UPPER(nrm.Track)   AS TrackKey
    FROM report.rcm_job_categories rc
    JOIN raw.mvp_job_categories jc
      ON UPPER(LTRIM(RTRIM(jc.JobCategoryName))) = UPPER(LTRIM(RTRIM(rc.JobCategoryName)))
    CROSS APPLY (VALUES (1,jc.TrainingTrack1),(2,jc.TrainingTrack2),(3,jc.TrainingTrack3),
                        (4,jc.TrainingTrack4),(5,jc.TrainingTrack5),(6,jc.TrainingTrack6)) t(TrackSlot,Track)
    CROSS APPLY (SELECT REPLACE(REPLACE(LTRIM(RTRIM(t.Track)),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Track) nrm
    WHERE LTRIM(RTRIM(ISNULL(t.Track,''))) <> ''
),
track_cls AS (       -- track -> class (Epic Curriculum Status, our pop), era-normalized
    SELECT DISTINCT
           UPPER(REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Curriculum))),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_')) AS TrackKey,
           REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))),
                 'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_')  AS ClassTitle,
           CONVERT(nvarchar(50),s.Event_Class_Type)             AS ClassType
    FROM raw.epic_status s
    JOIN report.users u
      ON u.UniversalID = UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))
     AND u.OnLeaderList = 1
    WHERE LTRIM(RTRIM(ISNULL(s.Event_Class,''))) <> ''
)
SELECT
    ct.JobCategoryGroup,
    ct.JobCategoryID,
    ct.JobCategoryName,
    ct.LinkableTemplateID,
    ct.LinkableTemplateName,
    ct.TrackSlot,
    ct.TrainingTrack,
    tc.ClassTitle,
    tc.ClassType,
    CASE WHEN UPPER(ct.TrainingTrack) = 'NO TRAINING' THEN 'No Training'
         WHEN tc.ClassTitle IS NULL                    THEN 'No class data'
         ELSE 'Mapped' END                                        AS MappingStatus,
    CASE WHEN cls.InCornerstone = 1 THEN 1 ELSE 0 END             AS ClassInCornerstone,
    CASE WHEN cls.OnClassSchedule = 1 THEN 1 ELSE 0 END           AS ClassOnSchedule,
    ct.Description
FROM cat_tracks ct
LEFT JOIN track_cls tc ON tc.TrackKey = ct.TrackKey
LEFT JOIN report.rcm_classes cls
       ON UPPER(cls.ClassTitle) = UPPER(tc.ClassTitle);
GO

/* ---------- 41. report.registration_audit — registration hygiene flags ----------
   Shows      : Wave 3, our leaders. One row per person x flagged issue,
                six checks: 'Out-of-scope registration' (departed/out-of-
                scope people still holding future seats), 'Schedule conflict'
                (two overlapping same-day registered sessions), 'Duplicate
                registration' (>1 CURRENT booking for the same class — Epic
                future sessions or Cornerstone Registered transcripts, W2/W3
                era names normalized), 'Curriculum not in role mapping'
                (assigned Epic curriculum outside the MVP tracks of the
                person's JobRole1-4 per report.rcm_training_map),
                'Registered for class not required' (future registration for
                a class in NEITHER the person's Epic requirements NOR their
                MVP role tracks — the actionable bucket; Detail carries the
                evidence), and 'OK - correct class, Epic gap' (class matches
                their role tracks but Epic shows no requirement — on leave /
                Epic-ineligible / lag; NOT an issue, listed only to explain
                Epic-vs-Cornerstone differences).
                Cornerstone-based checks count only CURRENT registrations:
                a Registered row superseded by a newer Withdrawn transcript
                for the same session (the 2026-07-08 mass-withdrawal
                pattern) is ignored. Zero rows for a check = it ran clean.
                Feeds the tracker's Registration Audit sheet.
   Built from : report.users + raw.epic_status + raw.epic_class_schedule
                + raw.cornerstone + report.rcm_training_map
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.registration_audit AS
WITH pop AS (
    SELECT u.UniversalID, u.FullName, u.Leader, u.Wave, u.IsInScope,
           u.JobRole1, u.JobRole2, u.JobRole3, u.JobRole4
    FROM report.users u
    WHERE u.Wave = 'Wave 3' AND u.OnLeaderList = 1
),
regall AS (  -- distinct registered FUTURE sessions per person x class x day
    SELECT DISTINCT
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id))))  AS UniversalID,
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class))))  AS ClassKey,
        LTRIM(RTRIM(CONVERT(nvarchar(200),s.Event_Class)))         AS ClassTitle,
        TRY_CONVERT(date, s.Event_Class_Date)                      AS EventDate
    FROM raw.epic_status s
    WHERE s.Event_Class_Type = 'Session'
      AND UPPER(ISNULL(s.Event_Class_Registered,'')) LIKE 'YES%'
      AND TRY_CONVERT(date, s.Event_Class_Date) >= CAST(GETDATE() AS date)
),
sched AS (   -- one time window per class x day
    SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Event_Name)))) AS ClassKey,
           TRY_CONVERT(date,[Date])    AS EventDate,
           MIN(Class_Event_Start_Time) AS StartTime,
           MAX(Class_Event_End_Time)   AS EndTime
    FROM raw.epic_class_schedule
    GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),Event_Name)))),
             TRY_CONVERT(date,[Date])
)
SELECT p.UniversalID, p.FullName, p.Leader, p.Wave,
       CAST('Out-of-scope registration' AS nvarchar(40)) AS CheckName,
       CAST(NULL AS nvarchar(200)) AS TrainingTrack,
       r.ClassTitle, r.EventDate,
       CAST('Departed / out of scope' AS nvarchar(200)) AS Detail
FROM pop p
JOIN regall r ON r.UniversalID = p.UniversalID
WHERE p.IsInScope = 0
UNION ALL
SELECT p.UniversalID, p.FullName, p.Leader, p.Wave,
       'Schedule conflict', NULL,
       a.ClassTitle, a.EventDate,
       'Overlaps ' + b.ClassTitle
FROM pop p
JOIN regall a ON a.UniversalID = p.UniversalID
JOIN regall b ON b.UniversalID = a.UniversalID
            AND b.EventDate = a.EventDate AND b.ClassKey > a.ClassKey
JOIN sched sa ON sa.ClassKey = a.ClassKey AND sa.EventDate = a.EventDate
JOIN sched sb ON sb.ClassKey = b.ClassKey AND sb.EventDate = b.EventDate
WHERE p.IsInScope = 1
  AND sa.StartTime < sb.EndTime AND sb.StartTime < sa.EndTime
UNION ALL
SELECT p.UniversalID, p.FullName, p.Leader, p.Wave,
       'Duplicate registration', NULL,
       MAX(r.ClassTitle), MAX(r.EventDate),
       CONVERT(nvarchar(20), COUNT(*)) + ' future sessions booked'
FROM pop p
JOIN regall r ON r.UniversalID = p.UniversalID
WHERE p.IsInScope = 1
GROUP BY p.UniversalID, p.FullName, p.Leader, p.Wave, r.ClassKey
HAVING COUNT(*) > 1
UNION ALL
SELECT p.UniversalID, p.FullName, p.Leader, p.Wave,
       'Duplicate registration', NULL,
       MAX(c.Title), NULL,
       CONVERT(nvarchar(20), COUNT(*)) + ' registered transcripts (Cornerstone)'
FROM pop p
JOIN (SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),c0.[User_ID])))) AS uid,
             UPPER(REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),c0.Training_Title))),
                   'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_')) AS ClsKey,
             REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),c0.Training_Title))),
                   'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_')  AS Title
      FROM raw.cornerstone c0
      WHERE UPPER(ISNULL(c0.Training_Provider,'')) = 'EPIC'
        AND c0.Training_Type = 'Session'
        AND c0.Transcript_Status = 'Registered'
        AND NOT EXISTS (      -- ignore stale rows superseded by a withdrawal
            SELECT 1 FROM raw.cornerstone w
            WHERE w.[User_ID] = c0.[User_ID]
              AND w.Training_Title = c0.Training_Title
              AND w.Training_Start_Date = c0.Training_Start_Date
              AND w.Transcript_Status = 'Withdrawn'
              AND TRY_CONVERT(datetime, w.Transcript_Registration_Date)
                  > TRY_CONVERT(datetime, c0.Transcript_Registration_Date))) c
  ON c.uid = p.UniversalID
WHERE p.IsInScope = 1
GROUP BY p.UniversalID, p.FullName, p.Leader, p.Wave, c.ClsKey
HAVING COUNT(*) > 1
UNION ALL
SELECT p.UniversalID, p.FullName, p.Leader, p.Wave,
       'Curriculum not in role mapping',
       asg.Track, NULL, NULL,
       'Not in MVP tracks for role(s)'
FROM pop p
JOIN (SELECT DISTINCT
             UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s.Universal_Id)))) AS uid,
             REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),s.Curriculum))),
                   'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_') AS Track
      FROM raw.epic_status s
      WHERE LTRIM(RTRIM(ISNULL(s.Curriculum,''))) <> '') asg
  ON asg.uid = p.UniversalID
WHERE p.IsInScope = 1
  AND NOT EXISTS (
      SELECT 1 FROM report.rcm_training_map m
      WHERE UPPER(m.TrainingTrack) = UPPER(asg.Track)
        AND UPPER(LTRIM(RTRIM(m.JobCategoryName))) IN (
            UPPER(LTRIM(RTRIM(ISNULL(p.JobRole1,'')))),
            UPPER(LTRIM(RTRIM(ISNULL(p.JobRole2,'')))),
            UPPER(LTRIM(RTRIM(ISNULL(p.JobRole3,'')))),
            UPPER(LTRIM(RTRIM(ISNULL(p.JobRole4,''))))))
UNION ALL
/* checks 5+6: future Cornerstone registrations for classes absent from the
   person's Epic requirement rows. Only CURRENT registrations count — a
   Registered row with a NEWER Withdrawn transcript for the same session
   (the 2026-07-08 mass-withdrawal pattern) is stale and ignored. Split:
   'Registered for class not required'  = class also absent from the MVP
       role->track->class mapping (verify with the leader; Detail notes
       when the person's Epic curriculum is marked 'No Training').
   'OK - correct class, Epic gap'       = class DOES match the person's
       role tracks — the registration is RIGHT; Epic just can't see them
       (on leave / dashboard-ineligible / requirement lag). Not an issue. */
SELECT p.UniversalID, p.FullName, p.Leader, p.Wave,
       CASE WHEN rm.RoleMatch = 1
            THEN 'OK - correct class, Epic gap'
            ELSE 'Registered for class not required' END,
       CAST(NULL AS nvarchar(200)),
       r.Title, r.StartDate,
       CASE
         WHEN rm.RoleMatch = 1 AND el.Eligible = 'No'
           THEN 'Class matches role tracks; Epic-ineligible ('
                + ISNULL(NULLIF(el.HrStatus,''),'status unknown') + ')'
         WHEN rm.RoleMatch = 1
           THEN 'Class matches role tracks; Epic requirement lag'
         WHEN EXISTS (SELECT 1 FROM raw.epic_status s4
                      WHERE UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s4.Universal_Id))))
                            = p.UniversalID
                        AND UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200),s4.Event_Class))))
                            = 'NO TRAINING')
           THEN 'Not in Epic/role tracks; an Epic curriculum is marked No Training'
         ELSE 'Not in Epic/role tracks; verify with leader' END
FROM pop p
JOIN (SELECT DISTINCT
             UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),c.[User_ID])))) AS uid,
             UPPER(REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),c.Training_Title))),
                   'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_')) AS ClsKey,
             REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),c.Training_Title))),
                   'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_')  AS Title,
             TRY_CONVERT(date, c.Training_Start_Date)             AS StartDate
      FROM raw.cornerstone c
      WHERE UPPER(ISNULL(c.Training_Provider,'')) = 'EPIC'
        AND c.Training_Type = 'Session'
        AND c.Transcript_Status = 'Registered'
        AND TRY_CONVERT(date, c.Training_Start_Date) >= CAST(GETDATE() AS date)
        AND NOT EXISTS (      -- ignore stale rows superseded by a withdrawal
            SELECT 1 FROM raw.cornerstone w
            WHERE w.[User_ID] = c.[User_ID]
              AND w.Training_Title = c.Training_Title
              AND w.Training_Start_Date = c.Training_Start_Date
              AND w.Transcript_Status = 'Withdrawn'
              AND TRY_CONVERT(datetime, w.Transcript_Registration_Date)
                  > TRY_CONVERT(datetime, c.Transcript_Registration_Date))) r
  ON r.uid = p.UniversalID
CROSS APPLY (SELECT CASE WHEN EXISTS (
        SELECT 1 FROM report.rcm_training_map m
        WHERE UPPER(m.ClassTitle) = r.ClsKey
          AND UPPER(LTRIM(RTRIM(m.JobCategoryName))) IN (
              UPPER(LTRIM(RTRIM(ISNULL(p.JobRole1,'')))),
              UPPER(LTRIM(RTRIM(ISNULL(p.JobRole2,'')))),
              UPPER(LTRIM(RTRIM(ISNULL(p.JobRole3,'')))),
              UPPER(LTRIM(RTRIM(ISNULL(p.JobRole4,''))))))
        THEN 1 ELSE 0 END AS RoleMatch) rm
LEFT JOIN (SELECT UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),Universal_ID)))) AS uid,
                  MAX(LTRIM(RTRIM(ISNULL(Epic_Training_Eligible_Appears_on_Dashboard,'')))) AS Eligible,
                  MAX(LTRIM(RTRIM(ISNULL(Hr_Status_PDM,''))))                               AS HrStatus
           FROM raw.epic_lookup
           GROUP BY UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),Universal_ID))))) el
  ON el.uid = p.UniversalID
WHERE p.IsInScope = 1
  AND NOT EXISTS (
      SELECT 1 FROM raw.epic_status s3
      WHERE UPPER(LTRIM(RTRIM(CONVERT(nvarchar(40),s3.Universal_Id))))
            = p.UniversalID
        AND UPPER(REPLACE(REPLACE(LTRIM(RTRIM(CONVERT(nvarchar(200),s3.Event_Class))),
              'EPIC_W2_','EPIC_NH_'),'EPIC_W3_','EPIC_NH_')) = r.ClsKey);
GO

/* ---------- pbi schema: the plug-and-play Power BI surface ----------
   One clean namespace to point Power BI at: Get Data -> SQL Server ->
   .\SQLEXPRESS / AnalyticsDB -> pick everything under "pbi".
   Thin aliases only — logic lives in report.*. Re-run this script after any
   underlying view/table gains columns so the aliases pick them up. */
IF SCHEMA_ID('pbi') IS NULL EXEC('CREATE SCHEMA pbi');
GO
/* pbi.roster — Power BI alias of report.roster */
CREATE OR ALTER VIEW pbi.roster             AS SELECT * FROM report.roster;
GO
/* pbi.kpi_summary — Power BI alias of report.kpi_summary */
CREATE OR ALTER VIEW pbi.kpi_summary        AS SELECT * FROM report.kpi_summary;
GO
/* pbi.leader_summary — Power BI alias of report.leader_summary */
CREATE OR ALTER VIEW pbi.leader_summary     AS SELECT * FROM report.leader_summary;
GO
/* pbi.registration_trend — Power BI alias of report.registration_daily */
CREATE OR ALTER VIEW pbi.registration_trend AS SELECT * FROM report.registration_daily;
GO
/* pbi.daily_changes — Power BI alias of report.daily_changes */
CREATE OR ALTER VIEW pbi.daily_changes      AS SELECT * FROM report.daily_changes;
GO
/* pbi.person_events — Power BI alias of report.person_events */
CREATE OR ALTER VIEW pbi.person_events      AS SELECT * FROM report.person_events;
GO
/* pbi.noshow — Power BI alias of report.noshow */
CREATE OR ALTER VIEW pbi.noshow             AS SELECT * FROM report.noshow;
GO
/* pbi.unregistered — Power BI alias of report.unregistered */
CREATE OR ALTER VIEW pbi.unregistered       AS SELECT * FROM report.unregistered;
GO
/* pbi.epic_not_in_hr — Power BI alias of report.epic_not_in_hr */
CREATE OR ALTER VIEW pbi.epic_not_in_hr     AS SELECT * FROM report.epic_not_in_hr;
GO
/* pbi.run_history — Power BI alias of raw.activity_log (pipeline run history) */
CREATE OR ALTER VIEW pbi.run_history        AS SELECT * FROM raw.activity_log;
GO
/* pbi.tracker_detail — Power BI alias of report.tracker_detail
   (RCM Training Tracker Data sheet: one row per Epic status export row) */
CREATE OR ALTER VIEW pbi.tracker_detail          AS SELECT * FROM report.tracker_detail;
GO
/* pbi.tracker_class_schedule — Power BI alias of report.tracker_class_schedule
   (tracker Registration Schedule: one row per scheduled session) */
CREATE OR ALTER VIEW pbi.tracker_class_schedule  AS SELECT * FROM report.tracker_class_schedule;
GO
/* pbi.tracker_daily_log — Power BI alias of report.tracker_daily_log
   (tracker/Daily Log trend: one row per snapshot day per leader) */
CREATE OR ALTER VIEW pbi.tracker_daily_log       AS SELECT * FROM report.tracker_daily_log;
GO
/* pbi.tracker_training_status — Power BI alias of report.tracker_training_status
   (tracker Training Status sheet: one row per in-scope wave person) */
CREATE OR ALTER VIEW pbi.tracker_training_status AS SELECT * FROM report.tracker_training_status;
GO
/* pbi.metrics_history — Power BI alias of raw.metrics_summary (metric trend) */
CREATE OR ALTER VIEW pbi.metrics_history    AS SELECT * FROM raw.metrics_summary;
GO

/* ---------- report.leader_business_units — BU headcount per leader ----------
   Shows      : one row per Leader + BusinessUnit with user count
   Built from : report.users (Leader) + raw.mvp BusinessUnit (the column the
                MVP updater writes to the Master — NOT BusinessUnitDescription,
                which is sparse), scoped to canonical leaders (raw.ref_leaders)
   Added      : 2026-08-03 (the analyst's ask — leader/BU pull on demand)
   --------------------------------------------------------------------------- */
CREATE OR ALTER VIEW report.leader_business_units AS
SELECT
    u.Leader,
    m.BusinessUnit,
    COUNT(*) AS Users
FROM report.users u
JOIN raw.mvp m
    ON UPPER(LTRIM(RTRIM(m.UniversalID))) = u.UniversalID
JOIN raw.ref_leaders l
    ON l.Leader = u.Leader
GROUP BY u.Leader, m.BusinessUnit;
GO

/* ---------- report.user_training_summary — per-user training rollup ----------
   Shows      : one row per Master user reporting to a canonical leader
                (raw.ref_leaders; departed included — see DepartedInactive):
                identity, leadership chain (SrDirector from HR), Epic TM
                Lookup training flags (per-MVP + dashboard-eligible), MVP
                departed flag, UserType, JobRoles 1-4, Fully Registered /
                Trained Yes-No (Epic dashboard flags).
                TrainingCompletionStatus = CORNERSTONE-driven (her call
                2026-08-04, Cornerstone = training source of truth):
                completed-class count (latest-live EPIC sessions, via
                report.cornerstone_status) vs Epic ClassesRequired.
                LatestCompletedClassDate = Cornerstone; Estimated date = Epic
                (future schedules only exist in Epic). All dates DATE-only.
                NULL status = outside training scope (e.g. not needed).
   Built from : report.roster + report.hr + report.tracker_training_status
                + report.cornerstone_status + raw.cornerstone
   Added      : 2026-08-04 (the analyst's ask — leadership-ready user list)
   --------------------------------------------------------------------------- */
CREATE OR ALTER VIEW report.user_training_summary AS
WITH corn_done AS (   -- latest completed EPIC session date per user (Cornerstone)
    SELECT UPPER(LTRIM(RTRIM(User_ID))) AS UniversalID,
           CAST(MAX(TRY_CONVERT(datetime, Transcript_Completed_Date)) AS date) AS LastCompletedDate
    FROM raw.cornerstone
    WHERE UPPER(LTRIM(RTRIM(ISNULL(Training_Provider,'')))) = 'EPIC'
      AND Training_Type = 'Session'
      AND Transcript_Status LIKE 'Completed%'
    GROUP BY UPPER(LTRIM(RTRIM(User_ID)))
)
SELECT
    r.UniversalID,
    r.FullName,
    r.Wave                    AS GoLiveWave,
    hr.SeniorDirector         AS SrDirector,
    r.AVP,
    r.VP,
    r.Leader                  AS Leaders,
    r.EpicTrainingNeeded      AS TrainingRequiredPerMVP,
    r.EpicEligible            AS TrainingRequiredDashboard,
    r.Departed                AS DepartedInactive,
    r.UserType,
    r.JobRole1, r.JobRole2, r.JobRole3, r.JobRole4,
    COALESCE(ts.FullyRegisteredYN,
             CASE WHEN r.IsFullyRegistered = 1 THEN 'Yes' ELSE 'No' END) AS FullyRegisteredYN,
    COALESCE(ts.FullyTrainedYN,
             CASE WHEN r.IsFullyTrained = 1 THEN 'Yes' ELSE 'No' END)    AS FullyTrainedYN,
    CASE WHEN ts.UniversalID IS NULL THEN NULL
         WHEN ISNULL(ts.ClassesRequired, 0) = 0            THEN 'No Epic Curriculum'
         WHEN ISNULL(cs.CompletedCount, 0) >= ts.ClassesRequired THEN 'Fully Trained'
         WHEN ISNULL(cs.CompletedCount, 0) > 0             THEN 'In Progress'
         WHEN ISNULL(cs.RegisteredCount, 0) > 0            THEN 'Registered'
         ELSE 'Not Started' END                            AS TrainingCompletionStatus,
    cd.LastCompletedDate                                   AS LatestCompletedClassDate,
    CAST(ts.FinalScheduledDate AS date)                    AS EstimatedTrainingCompletionDate
FROM report.roster r
JOIN raw.ref_leaders l ON l.Leader = r.Leader
LEFT JOIN report.hr hr ON hr.UniversalID = r.UniversalID
LEFT JOIN report.tracker_training_status ts ON ts.UniversalID = r.UniversalID
LEFT JOIN report.cornerstone_status cs ON cs.UniversalID = r.UniversalID
LEFT JOIN corn_done cd ON cd.UniversalID = r.UniversalID;
GO

PRINT 'report.* views created/updated.';
GO
