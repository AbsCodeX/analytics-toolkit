-- =========================================================================
-- SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
-- Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
-- fill in your own values before running. See README.md for the full map.
-- =========================================================================
/* ============================================================
   4_dimensions.sql — RUN ONCE (and again after editing)
   The dim.* layer: conformed dimension views for mapping and
   left/right joins across the five source vocabularies
   (Master/Wave, HR, MVP, Epic, Cornerstone).

   Backed by two things:
     - the OBSERVED values in raw.* (rebuilt every refresh), and
     - the EDITABLE mappings in data\references\wave_reference_lists.xlsx
       (loaded as raw.ref_* by refresh.py's `refs` loader).
   To fix a mapping: edit the workbook, run `python sql\refresh.py refs`.

   Join keys are always normalized UPPER(LTRIM(RTRIM(...))).

   How to run:
     sqlcmd -S ".\SQLEXPRESS" -E -C -d AnalyticsDB -i "sql\4_dimensions.sql"
   ============================================================ */
USE AnalyticsDB;
GO
IF SCHEMA_ID('dim') IS NULL EXEC('CREATE SCHEMA dim');
GO

/* ---------- dim.vendor_alias — variant spelling -> canonical vendor ----------
   Shows      : one row per known alias spelling with its ONE canonical
                vendor name — join any BusinessUnit through this
   Built from : raw.ref_vendor_aliases (the vendor_aliases sheet of
                wave_reference_lists.xlsx — edit there, then `refresh.py refs`)
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.vendor_alias AS
SELECT UPPER(LTRIM(RTRIM(Alias)))  AS Alias,
       LTRIM(RTRIM(CanonicalVendor)) AS CanonicalVendor
FROM raw.ref_vendor_aliases
WHERE LTRIM(RTRIM(ISNULL(Alias,''))) <> '';
GO

/* ---------- dim.vendor — canonical vendors + Master headcount ----------
   Shows      : one row per canonical vendor — alias count, Master people,
                offshore/onshore split
   Built from : dim.vendor_alias + raw.master
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.vendor AS
WITH m AS (
    SELECT va.CanonicalVendor,
           COUNT(*) AS MasterPeople,
           SUM(CASE WHEN UPPER(LTRIM(RTRIM(x.[User_Type_Offshore_Onshore_YourOrg]))) = 'OFFSHORE'
                    THEN 1 ELSE 0 END) AS Offshore,
           SUM(CASE WHEN UPPER(LTRIM(RTRIM(x.[User_Type_Offshore_Onshore_YourOrg]))) = 'ONSHORE'
                    THEN 1 ELSE 0 END) AS Onshore
    FROM raw.[master] x
    JOIN dim.vendor_alias va ON va.Alias = UPPER(LTRIM(RTRIM(x.BusinessUnit)))
    WHERE UPPER(LTRIM(RTRIM(x.[Vendor_Yes_No]))) = 'YES'
    GROUP BY va.CanonicalVendor
)
SELECT va.CanonicalVendor, COUNT(*) AS AliasCount,
       ISNULL(MAX(m.MasterPeople),0) AS MasterPeople,
       ISNULL(MAX(m.Offshore),0) AS Offshore,
       ISNULL(MAX(m.Onshore),0) AS Onshore
FROM dim.vendor_alias va
LEFT JOIN m ON m.CanonicalVendor = va.CanonicalVendor
GROUP BY va.CanonicalVendor;
GO

/* ---------- dim.business_unit — observed BU catalog across all sources ----------
   Shows      : one row per distinct BU value — In<source> flags, vendor
                mapping, NeedsMapping (1 = Master value not in the ref).
                Four vocabularies by nature: Master/ref = vendor & org names;
                HR = facilities; Cornerstone = cost centers; Epic = org groups.
   Built from : raw.master + raw.hr + raw.cornerstone + raw.epic_lookup
                + raw.ref_business_units + dim.vendor_alias
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.business_unit AS
WITH src AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(BusinessUnit))) AS bu, 'Master' AS s
    FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL(BusinessUnit,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(BusinessUnitDesc))), 'HR'
    FROM raw.hr WHERE LTRIM(RTRIM(ISNULL(BusinessUnitDesc,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(Business_Unit))), 'Cornerstone'
    FROM raw.cornerstone WHERE LTRIM(RTRIM(ISNULL(Business_Unit,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(Business_Unit))), 'EpicLookup'
    FROM raw.epic_lookup WHERE LTRIM(RTRIM(ISNULL(Business_Unit,''))) <> ''
),
agg AS (
    SELECT bu,
           MAX(CASE WHEN s='Master'      THEN 1 ELSE 0 END) AS InMaster,
           MAX(CASE WHEN s='HR'          THEN 1 ELSE 0 END) AS InHR,
           MAX(CASE WHEN s='Cornerstone' THEN 1 ELSE 0 END) AS InCornerstone,
           MAX(CASE WHEN s='EpicLookup'  THEN 1 ELSE 0 END) AS InEpicLookup
    FROM src GROUP BY bu
)
SELECT agg.bu AS BusinessUnit,
       agg.InMaster, agg.InHR, agg.InCornerstone, agg.InEpicLookup,
       ref.IsVendor, ref.UserType,
       va.CanonicalVendor,
       CASE WHEN agg.InMaster = 1 AND ref.BusinessUnit IS NULL THEN 1 ELSE 0 END AS NeedsMapping
FROM agg
LEFT JOIN raw.ref_business_units ref
       ON UPPER(LTRIM(RTRIM(ref.BusinessUnit))) = agg.bu
LEFT JOIN dim.vendor_alias va ON va.Alias = agg.bu;
GO

/* ---------- dim.job_title — observed title catalog + leader flag ----------
   Shows      : one row per distinct job title — In<source> flags +
                IsLeaderTitle (from the ref list)
   Built from : raw.master + raw.hr + raw.cornerstone + raw.epic_lookup
                + raw.ref_job_titles_leader_flag
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.job_title AS
WITH src AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(JobTitle))) AS jt, 'Master' AS s
    FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL(JobTitle,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(JobTitle))), 'HR'
    FROM raw.hr WHERE LTRIM(RTRIM(ISNULL(JobTitle,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(Job_Title))), 'Cornerstone'
    FROM raw.cornerstone WHERE LTRIM(RTRIM(ISNULL(Job_Title,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(Job_Title))), 'EpicLookup'
    FROM raw.epic_lookup WHERE LTRIM(RTRIM(ISNULL(Job_Title,''))) <> ''
),
agg AS (
    SELECT jt,
           MAX(CASE WHEN s='Master'      THEN 1 ELSE 0 END) AS InMaster,
           MAX(CASE WHEN s='HR'          THEN 1 ELSE 0 END) AS InHR,
           MAX(CASE WHEN s='Cornerstone' THEN 1 ELSE 0 END) AS InCornerstone,
           MAX(CASE WHEN s='EpicLookup'  THEN 1 ELSE 0 END) AS InEpicLookup
    FROM src GROUP BY jt
)
SELECT agg.jt AS JobTitle,
       agg.InMaster, agg.InHR, agg.InCornerstone, agg.InEpicLookup,
       CASE WHEN ref.JobTitle IS NOT NULL
            AND UPPER(LTRIM(RTRIM(ISNULL(ref.IsLeader,'')))) = 'YES'
            THEN 1 ELSE 0 END AS IsLeaderTitle
FROM agg
LEFT JOIN raw.ref_job_titles_leader_flag ref
       ON UPPER(LTRIM(RTRIM(ref.JobTitle))) = agg.jt;
GO

/* ---------- dim.department — observed department catalog ----------
   Shows      : one row per distinct department — InMaster/InHR/InMVP flags
   Built from : raw.master + raw.hr + raw.mvp
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.department AS
WITH src AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(DepartmentLocation))) AS d, 'Master' AS s
    FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL(DepartmentLocation,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(Department_Name))), 'HR'
    FROM raw.hr WHERE LTRIM(RTRIM(ISNULL(Department_Name,''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM(DepartmentLocation))), 'MVP'
    FROM raw.mvp WHERE LTRIM(RTRIM(ISNULL(DepartmentLocation,''))) <> ''
)
SELECT d AS Department,
       MAX(CASE WHEN s='Master' THEN 1 ELSE 0 END) AS InMaster,
       MAX(CASE WHEN s='HR'     THEN 1 ELSE 0 END) AS InHR,
       MAX(CASE WHEN s='MVP'    THEN 1 ELSE 0 END) AS InMVP
FROM src GROUP BY d;
GO

/* ---------- dim.leader — canonical leaders + everything observed ----------
   Shows      : one row per leader — IsCanonical, training preference,
                Master headcount, NeedsReview (observed but not canonical)
   Built from : raw.master + raw.ref_leaders + raw.ref_leader_training_preference
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.leader AS
WITH obs AS (
    SELECT UPPER(LTRIM(RTRIM(Leaders))) AS leader, COUNT(*) AS MasterPeople
    FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL(Leaders,''))) <> ''
    GROUP BY UPPER(LTRIM(RTRIM(Leaders)))
),
canon AS (
    SELECT UPPER(LTRIM(RTRIM(r.Leader))) AS leader,
           LTRIM(RTRIM(r.Leader)) AS LeaderName,
           p.TrainingPreference
    FROM raw.ref_leaders r
    LEFT JOIN raw.ref_leader_training_preference p
           ON UPPER(LTRIM(RTRIM(p.Leader))) = UPPER(LTRIM(RTRIM(r.Leader)))
    WHERE LTRIM(RTRIM(ISNULL(r.Leader,''))) <> ''
)
SELECT COALESCE(canon.LeaderName, obs.leader) AS Leader,
       CASE WHEN canon.leader IS NOT NULL THEN 1 ELSE 0 END AS IsCanonical,
       canon.TrainingPreference,
       ISNULL(obs.MasterPeople, 0) AS MasterPeople,
       CASE WHEN canon.leader IS NULL
             AND obs.leader <> 'NO LONGER REV CYCLE' THEN 1 ELSE 0 END AS NeedsReview
FROM canon
FULL OUTER JOIN obs ON obs.leader = canon.leader;
GO

/* ---------- dim.job_role — job-role catalog + no-training flag ----------
   Shows      : one row per distinct role (Master roles 1-4 + MVP role 1) —
                InMaster/InMVP + NoTrainingNeeded (from the ref list)
   Built from : raw.master + raw.mvp + raw.ref_no_training_job_roles
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.job_role AS
WITH src AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM([Job_Role_1]))) AS jr, 'Master' AS s FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL([Job_Role_1],''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM([Job_Role_2]))), 'Master' FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL([Job_Role_2],''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM([Job_Role_3]))), 'Master' FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL([Job_Role_3],''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM([Job_Role_4]))), 'Master' FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL([Job_Role_4],''))) <> ''
    UNION ALL
    SELECT DISTINCT UPPER(LTRIM(RTRIM([IndividualCategoryUpdate1Name]))), 'MVP' FROM raw.mvp WHERE LTRIM(RTRIM(ISNULL([IndividualCategoryUpdate1Name],''))) <> ''
)
SELECT jr AS JobRole,
       MAX(CASE WHEN s='Master' THEN 1 ELSE 0 END) AS InMaster,
       MAX(CASE WHEN s='MVP'    THEN 1 ELSE 0 END) AS InMVP,
       MAX(CASE WHEN nt.JobRole IS NOT NULL THEN 1 ELSE 0 END) AS NoTrainingNeeded
FROM src
LEFT JOIN raw.ref_no_training_job_roles nt
       ON UPPER(LTRIM(RTRIM(nt.JobRole))) = src.jr
GROUP BY jr;
GO

/* ---------- dim.epic_job_category — MVP's Epic job-category reference ----------
   Shows      : one row per Epic job category — ID, name, group, description
   Built from : raw.mvp_job_categories
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW dim.epic_job_category AS
SELECT [JobCategoryID]    AS JobCategoryID,
       [JobCategoryName]  AS JobCategoryName,
       [JobCategoryGroup] AS JobCategoryGroup,
       [Description]      AS Description
FROM raw.mvp_job_categories;
GO

/* ---------- report.mapping_gaps — actionable conformance issues ----------
   Shows      : one row per unmapped value (Master BU not in ref, vendor BU
                without canonical alias, non-canonical leader) with people
                counts. Fix: edit wave_reference_lists.xlsx (then `refresh.py
                refs`) or fix the Master value. Feeds Morning Review.
   Built from : raw.master + raw.ref_business_units + dim.vendor_alias
                + raw.ref_leaders
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.mapping_gaps AS
SELECT 'Master BU not in ref_business_units' AS Issue,
       x.bu AS Value, COUNT(*) AS People
FROM (SELECT UPPER(LTRIM(RTRIM(BusinessUnit))) AS bu FROM raw.[master]
      WHERE LTRIM(RTRIM(ISNULL(BusinessUnit,''))) <> '') x
LEFT JOIN raw.ref_business_units ref
       ON UPPER(LTRIM(RTRIM(ref.BusinessUnit))) = x.bu
WHERE ref.BusinessUnit IS NULL
GROUP BY x.bu
UNION ALL
SELECT 'Vendor BU without canonical alias (add to vendor_aliases)',
       x.bu, COUNT(*)
FROM (SELECT UPPER(LTRIM(RTRIM(BusinessUnit))) AS bu FROM raw.[master]
      WHERE UPPER(LTRIM(RTRIM([Vendor_Yes_No]))) = 'YES'
        AND LTRIM(RTRIM(ISNULL(BusinessUnit,''))) <> '') x
LEFT JOIN dim.vendor_alias va ON va.Alias = x.bu
WHERE va.Alias IS NULL
GROUP BY x.bu
UNION ALL
SELECT 'Leader value not canonical (check ref_leaders / fix Master)',
       x.leader, COUNT(*)
FROM (SELECT UPPER(LTRIM(RTRIM(Leaders))) AS leader FROM raw.[master]
      WHERE LTRIM(RTRIM(ISNULL(Leaders,''))) <> '') x
LEFT JOIN raw.ref_leaders r ON UPPER(LTRIM(RTRIM(r.Leader))) = x.leader
WHERE r.Leader IS NULL AND x.leader <> 'NO LONGER REV CYCLE'
GROUP BY x.leader;
GO

/* ---------- pbi aliases for the dimension layer ---------- */
/* pbi.dim_vendor — Power BI alias of dim.vendor */
CREATE OR ALTER VIEW pbi.dim_vendor        AS SELECT * FROM dim.vendor;
GO
/* pbi.dim_business_unit — Power BI alias of dim.business_unit */
CREATE OR ALTER VIEW pbi.dim_business_unit AS SELECT * FROM dim.business_unit;
GO
/* pbi.dim_leader — Power BI alias of dim.leader */
CREATE OR ALTER VIEW pbi.dim_leader        AS SELECT * FROM dim.leader;
GO

/* ======================================================================
   MASTER QA LAYER (added 2026-07-08)
   Review-in-SSMS / apply-in-Excel: these views RECOMMEND and VALIDATE —
   they never write. raw.master is replaced on every refresh, so an UPDATE
   here can never reach the Master file. Fix values in the Master workbook
   (or the reference workbook), refresh, and watch the counts drop.
   Runbook: sql\QUERIES.md "Master QA & gap views".
   ====================================================================== */

/* ---------- report.user_type_recommendations ----------
   Shows      : one row per Master person with the current User Type, the
                rule-derived recommendation, which rule fired, and confidence.
                High   = Epic Worker Type (YourOrg side) or BU-map vendor type
                Medium = BU map / MVP role tag without Epic confirmation
                Review = conflicting or missing signals (never guessed:
                         UserType_Recommended stays NULL)
   Built from : raw.master + raw.epic_lookup (Worker Type)
                + raw.ref_business_units (BU -> IsVendor/UserType)
                + raw.mvp (Offshore/Onshore tags in IndividualCategoryUpdate1-4Name)
   Note       : Vendor Yes/No in the Master is a live formula over User Type —
                fixing User Type fixes Vendor automatically.
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.user_type_recommendations AS
WITH ep AS (
    SELECT UniversalID, WorkerType, HrStatusPDM FROM (
        SELECT UPPER(LTRIM(RTRIM([Universal_ID])))  AS UniversalID,
               LTRIM(RTRIM([Worker_Type]))          AS WorkerType,
               LTRIM(RTRIM([Hr_Status_PDM]))        AS HrStatusPDM,
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM([Universal_ID])))
                                  ORDER BY (SELECT NULL)) AS rn
        FROM raw.epic_lookup
        WHERE LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> ''
    ) x WHERE rn = 1
),
mvp1 AS (
    SELECT UniversalID,
           CASE WHEN HasOff = 1 AND HasOn = 1 THEN 'CONFLICT'
                WHEN HasOff = 1 THEN 'Offshore'
                WHEN HasOn  = 1 THEN 'Onshore' END AS MvpRoleTag
    FROM (
        SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
               CASE WHEN UPPER(CONCAT(ISNULL([IndividualCategoryUpdate1Name],''),'|',
                                      ISNULL([IndividualCategoryUpdate2Name],''),'|',
                                      ISNULL([IndividualCategoryUpdate3Name],''),'|',
                                      ISNULL([IndividualCategoryUpdate4Name],'')))
                    LIKE '%OFFSHORE%' THEN 1 ELSE 0 END AS HasOff,
               CASE WHEN UPPER(CONCAT(ISNULL([IndividualCategoryUpdate1Name],''),'|',
                                      ISNULL([IndividualCategoryUpdate2Name],''),'|',
                                      ISNULL([IndividualCategoryUpdate3Name],''),'|',
                                      ISNULL([IndividualCategoryUpdate4Name],'')))
                    LIKE '%ONSHORE%' THEN 1 ELSE 0 END AS HasOn,
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM(UniversalID)))
                                  ORDER BY [LastImportedDate] DESC) AS rn
        FROM raw.mvp
        WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
    ) x WHERE rn = 1
),
bu AS (
    SELECT UPPER(LTRIM(RTRIM(BusinessUnit))) AS bu,
           LTRIM(RTRIM(IsVendor))            AS IsVendor,
           LTRIM(RTRIM(UserType))            AS BuUserType
    FROM raw.ref_business_units
    WHERE LTRIM(RTRIM(ISNULL(BusinessUnit,''))) <> ''
),
base AS (
    SELECT UPPER(LTRIM(RTRIM(m.UniversalID)))                    AS UniversalID,
           m.[Full_Name]                                         AS FullName,
           m.[GoLiveWave]                                        AS Wave,
           m.[Leaders]                                           AS Leader,
           m.[IsDepartedInactive]                                AS Departed,
           m.[BusinessUnit]                                      AS BusinessUnit,
           NULLIF(LTRIM(RTRIM(m.[User_Type_Offshore_Onshore_YourOrg])),'')
                                                                 AS UserType_Current,
           ep.WorkerType, ep.HrStatusPDM,
           CASE WHEN ep.UniversalID IS NOT NULL THEN 1 ELSE 0 END AS OnEpic,
           CASE WHEN UPPER(ep.WorkerType) LIKE '%YOURORG%'      THEN 'YourOrg'
                WHEN UPPER(ep.WorkerType) LIKE '%VENDOR%'
                  OR UPPER(ep.WorkerType) LIKE '%AGENCY%'         THEN 'Vendor'
           END                                                    AS WorkerTypeClass,
           bu.IsVendor, bu.BuUserType, mvp1.MvpRoleTag
    FROM raw.[master] m
    LEFT JOIN ep    ON ep.UniversalID   = UPPER(LTRIM(RTRIM(m.UniversalID)))
    LEFT JOIN bu    ON bu.bu            = UPPER(LTRIM(RTRIM(m.BusinessUnit)))
    LEFT JOIN mvp1  ON mvp1.UniversalID = UPPER(LTRIM(RTRIM(m.UniversalID)))
    WHERE LTRIM(RTRIM(ISNULL(m.UniversalID,''))) <> ''
),
ruled AS (
    SELECT b.*,
        CASE
          WHEN b.WorkerTypeClass = 'YourOrg'
           AND UPPER(ISNULL(b.IsVendor,'NO')) <> 'YES'            THEN 'EPIC_YOURORG'
          WHEN b.WorkerTypeClass = 'YourOrg'                    THEN 'CONFLICT_EPIC_VS_BU'
          WHEN b.WorkerTypeClass = 'Vendor'
            OR UPPER(ISNULL(b.IsVendor,'')) = 'YES' THEN
               CASE
                 WHEN b.BuUserType IN ('Offshore','Onshore')
                  AND b.MvpRoleTag IN ('Offshore','Onshore')
                  AND b.BuUserType <> b.MvpRoleTag                THEN 'CONFLICT_BU_VS_MVP'
                 WHEN b.BuUserType IN ('Offshore','Onshore')      THEN 'VENDOR_BU_TYPE'
                 WHEN b.MvpRoleTag IN ('Offshore','Onshore')      THEN 'VENDOR_MVP_TAG'
                 ELSE 'VENDOR_SHORE_UNKNOWN'
               END
          WHEN b.BuUserType = 'YourOrg'                         THEN 'BU_REF_YOURORG'
          WHEN b.MvpRoleTag IN ('Offshore','Onshore')             THEN 'MVP_TAG_ONLY'
          ELSE 'NO_SIGNAL'
        END AS RuleFired
    FROM base b
),
rec AS (
    SELECT r.*,
        CASE r.RuleFired
          WHEN 'EPIC_YOURORG'   THEN 'YourOrg'
          WHEN 'BU_REF_YOURORG' THEN 'YourOrg'
          WHEN 'VENDOR_BU_TYPE'   THEN r.BuUserType
          WHEN 'VENDOR_MVP_TAG'   THEN r.MvpRoleTag
          WHEN 'MVP_TAG_ONLY'     THEN r.MvpRoleTag
        END AS UserType_Recommended,
        CASE
          WHEN r.RuleFired IN ('EPIC_YOURORG','VENDOR_BU_TYPE')  THEN 'High'
          WHEN r.RuleFired IN ('BU_REF_YOURORG','VENDOR_MVP_TAG','MVP_TAG_ONLY')
                                                                   THEN 'Medium'
          ELSE 'Review'
        END AS Confidence
    FROM ruled r
)
SELECT UniversalID, FullName, Wave, Leader, Departed, BusinessUnit,
       UserType_Current, UserType_Recommended, RuleFired, Confidence,
       OnEpic, WorkerType, HrStatusPDM,
       IsVendor   AS BU_IsVendor,
       BuUserType AS BU_UserType,
       MvpRoleTag,
       CASE WHEN UserType_Current IS NULL THEN 1 ELSE 0 END AS CurrentBlank,
       CASE WHEN UserType_Recommended IS NOT NULL
             AND ISNULL(UPPER(UserType_Current),'') <> UPPER(UserType_Recommended)
            THEN 1 ELSE 0 END AS Differs
FROM rec;
GO

/* ---------- report.training_needed_check ----------
   Shows      : people whose Training Needed value trips a rule:
                - Master=Yes but departed / no-training job role /
                  Epic Hr Status Terminated or Pay Leave  -> Recommended 'No'
                - Master=No but Epic Training Needed=Yes  -> review flag only,
                  and ONLY when unexplained: rows whose mismatch is already
                  explained by a no-training role, departed flag, or Epic HR
                  status Terminated/Pay Leave are suppressed (user decision
                  2026-07-21 — Epic's Yes is a mapping artifact there).
                Blank Epic Hr Status = no signal (1,800+ vendors) — never a rule.
   Built from : raw.master + raw.epic_lookup + raw.ref_no_training_job_roles
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.training_needed_check AS
WITH ep AS (
    SELECT UniversalID, EpicTrainingNeeded, HrStatusPDM FROM (
        SELECT UPPER(LTRIM(RTRIM([Universal_ID])))        AS UniversalID,
               LTRIM(RTRIM([Epic_Training_Needed_MVP]))   AS EpicTrainingNeeded,
               LTRIM(RTRIM([Hr_Status_PDM]))              AS HrStatusPDM,
               ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM([Universal_ID])))
                                  ORDER BY (SELECT NULL)) AS rn
        FROM raw.epic_lookup
        WHERE LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> ''
    ) x WHERE rn = 1
),
m AS (
    SELECT UPPER(LTRIM(RTRIM(mm.UniversalID)))            AS UniversalID,
           mm.[Full_Name]                                 AS FullName,
           mm.[GoLiveWave]                                AS Wave,
           mm.[Leaders]                                   AS Leader,
           UPPER(LTRIM(RTRIM(mm.[Training_Needed_Yes_No]))) AS Master_TrainingNeeded,
           CASE WHEN UPPER(LTRIM(RTRIM(ISNULL(mm.[IsDepartedInactive],''))))
                     IN ('YES','Y','TRUE','1') THEN 1 ELSE 0 END AS IsDeparted,
           CASE WHEN EXISTS (
                    SELECT 1 FROM raw.ref_no_training_job_roles nt
                    WHERE UPPER(LTRIM(RTRIM(nt.JobRole))) IN
                          (UPPER(LTRIM(RTRIM(ISNULL(mm.[Job_Role_1],'')))),
                           UPPER(LTRIM(RTRIM(ISNULL(mm.[Job_Role_2],'')))),
                           UPPER(LTRIM(RTRIM(ISNULL(mm.[Job_Role_3],'')))),
                           UPPER(LTRIM(RTRIM(ISNULL(mm.[Job_Role_4],''))))))
                THEN 1 ELSE 0 END AS HasNoTrainingRole,
           mm.[Job_Role_1] AS JobRole1
    FROM raw.[master] mm
    WHERE LTRIM(RTRIM(ISNULL(mm.UniversalID,''))) <> ''
)
SELECT m.UniversalID, m.FullName, m.Wave, m.Leader,
       m.Master_TrainingNeeded, m.IsDeparted, m.HasNoTrainingRole, m.JobRole1,
       ep.EpicTrainingNeeded, ep.HrStatusPDM,
       CASE WHEN m.Master_TrainingNeeded = 'NO'
             AND UPPER(ISNULL(ep.EpicTrainingNeeded,'')) LIKE '%YES%'
             AND m.HasNoTrainingRole = 0 AND m.IsDeparted = 0
             AND UPPER(ISNULL(ep.HrStatusPDM,'')) NOT IN ('TERMINATED','PAY LEAVE')
            THEN 1 ELSE 0 END AS EpicYes_MasterNo,
       CASE WHEN m.Master_TrainingNeeded = 'YES'
             AND UPPER(ISNULL(ep.HrStatusPDM,'')) = 'TERMINATED'      THEN 1 ELSE 0 END AS TerminatedButYes,
       CASE WHEN m.Master_TrainingNeeded = 'YES'
             AND UPPER(ISNULL(ep.HrStatusPDM,'')) = 'PAY LEAVE'       THEN 1 ELSE 0 END AS PayLeaveButYes,
       CASE WHEN m.Master_TrainingNeeded = 'YES'
             AND m.HasNoTrainingRole = 1                              THEN 1 ELSE 0 END AS NoTrainingRoleButYes,
       CASE WHEN m.Master_TrainingNeeded = 'YES'
             AND m.IsDeparted = 1                                     THEN 1 ELSE 0 END AS DepartedButYes,
       CASE WHEN NULLIF(m.Master_TrainingNeeded,'') IS NULL           THEN 1 ELSE 0 END AS MasterBlank,
       CASE WHEN m.Master_TrainingNeeded = 'YES'
             AND (m.IsDeparted = 1 OR m.HasNoTrainingRole = 1
                  OR UPPER(ISNULL(ep.HrStatusPDM,'')) IN ('TERMINATED','PAY LEAVE'))
            THEN 'No' END AS Recommended,
       CASE WHEN m.IsDeparted = 1                                     THEN 'Departed/Inactive'
            WHEN m.HasNoTrainingRole = 1                              THEN 'No-training job role'
            WHEN UPPER(ISNULL(ep.HrStatusPDM,'')) = 'TERMINATED'      THEN 'Epic HR status: Terminated'
            WHEN UPPER(ISNULL(ep.HrStatusPDM,'')) = 'PAY LEAVE'       THEN 'Epic HR status: Pay Leave'
            WHEN m.Master_TrainingNeeded = 'NO'
             AND UPPER(ISNULL(ep.EpicTrainingNeeded,'')) LIKE '%YES%' THEN 'Epic says Yes — review'
       END AS Reason
FROM m
LEFT JOIN ep ON ep.UniversalID = m.UniversalID
WHERE (m.Master_TrainingNeeded = 'NO'  AND UPPER(ISNULL(ep.EpicTrainingNeeded,'')) LIKE '%YES%'
       AND m.HasNoTrainingRole = 0 AND m.IsDeparted = 0
       AND UPPER(ISNULL(ep.HrStatusPDM,'')) NOT IN ('TERMINATED','PAY LEAVE'))
   OR (m.Master_TrainingNeeded = 'YES' AND UPPER(ISNULL(ep.HrStatusPDM,'')) IN ('TERMINATED','PAY LEAVE'))
   OR (m.Master_TrainingNeeded = 'YES' AND m.HasNoTrainingRole = 1)
   OR (m.Master_TrainingNeeded = 'YES' AND m.IsDeparted = 1)
   OR NULLIF(m.Master_TrainingNeeded,'') IS NULL;
GO

/* ---------- report.hr_missing_from_wave ----------
   Shows      : HR staff whose COMPUTED leader (first non-blank of VP > AVP >
                SVP, exactly like missing_from_wave_hr.py) is on the
                ref_leaders allowlist — excluding Lastname09/Lastname02 (mixed
                RCM/clinical orgs) — with no Master row. SQL twin of the
                weekly hr_missing_from_wave xlsx. No departed filter (the
                Python report doesn't filter it either; column included).
   Built from : raw.hr_cleaned + raw.ref_leaders + raw.master
   Caveat     : canonical match is case-insensitive EXACT; the Python's
                normalize_leader alias layer (e.g. AliasLast->Lastname12) is not
                replicated — counts may differ by an alias edge case.
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.hr_missing_from_wave AS
-- 2026-08-26: mu now reads the pre-built indexed key table (raw.uids_master,
-- built by refresh.build_uids on every master load) rather than re-deriving
-- DISTINCT UPPER(TRIM()) over the raw.master heap each call. 19s -> <1s,
-- verified row-for-row identical.
WITH mu AS (
    SELECT UniversalID AS u FROM raw.uids_master
),
rl AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(Leader))) AS l FROM raw.ref_leaders
    WHERE LTRIM(RTRIM(ISNULL(Leader,''))) <> ''
),
h AS (
    SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
           [Full_Name] AS FullName, [Email], [GoLiveWave] AS Wave,
           [VP], [AVP], [SVP], [IsDeparted_Inactive], [IsRCM],
           COALESCE(
             CASE WHEN UPPER(LTRIM(RTRIM(ISNULL([VP],''))))  IN ('','NAN','NAN, NAN') THEN NULL
                  ELSE UPPER(LTRIM(RTRIM([VP])))  END,
             CASE WHEN UPPER(LTRIM(RTRIM(ISNULL([AVP],'')))) IN ('','NAN','NAN, NAN') THEN NULL
                  ELSE UPPER(LTRIM(RTRIM([AVP]))) END,
             CASE WHEN UPPER(LTRIM(RTRIM(ISNULL([SVP],'')))) IN ('','NAN','NAN, NAN') THEN NULL
                  ELSE UPPER(LTRIM(RTRIM([SVP]))) END
           ) AS computed_leader
    FROM raw.hr_cleaned
    WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
)
SELECT h.UniversalID, h.FullName, h.[Email], h.Wave,
       h.computed_leader AS LeadersComputed,
       h.[VP], h.[AVP], h.[SVP], h.[IsDeparted_Inactive] AS IsDepartedInactive, h.[IsRCM]
FROM h
JOIN rl      ON rl.l = h.computed_leader
LEFT JOIN mu ON mu.u = h.UniversalID
WHERE rl.l NOT IN ('LASTNAME09, FIRSTNAME09','LASTNAME02, FIRSTNAME02')
  AND mu.u IS NULL;
GO

/* ---------- report.epic_missing_from_wave ----------
   Shows      : Epic Team Member Lookup people on the RCM-Centralized
                curriculum with no Master row. SQL twin of
                missing_from_wave_epic.py / the not-in-wave review CSV.
   Built from : raw.epic_lookup + raw.uids_master
   Paired with: report.exceptions check 1 (3_report_views.sql) returns this
                same population as a must-be-zero alarm with 3 columns. The
                anti-join is written out in both places on purpose — this file
                applies after that one, so that view cannot depend on this.
                Change one, change the other.
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.epic_missing_from_wave AS
SELECT e.[Wave],
       UPPER(LTRIM(RTRIM(e.[Universal_ID]))) AS UniversalID,
       e.[Employee_ID]    AS EmployeeID,
       e.[Team_Member]    AS TeamMember,
       e.[Job_Title]      AS JobTitle,
       e.[Curriculum_Type] AS CurriculumType,
       e.[Worker_Type]    AS WorkerType,
       e.[Hr_Status_PDM]  AS HrStatusPDM,
       e.[Epic_Training_Needed_MVP] AS EpicTrainingNeeded,
       e.[Epic_Training_Eligible_Appears_on_Dashboard] AS EpicEligible,
       e.[Fully_Registered] AS FullyRegistered,
       e.[Fully_Trained]    AS FullyTrained,
       e.[Direct_Manager]   AS DirectManager,
       e.[Team_Member_Hire_Date] AS HireDate
FROM raw.epic_lookup e
-- 2026-08-26: indexed key-table anti-join instead of a correlated NOT EXISTS
-- over the raw.master heap. 31s -> 0.1s, verified row-for-row identical.
LEFT JOIN raw.uids_master mu
       ON mu.UniversalID = UPPER(LTRIM(RTRIM(e.[Universal_ID])))
WHERE LTRIM(RTRIM(ISNULL(e.[Universal_ID],''))) <> ''
  AND LTRIM(RTRIM(ISNULL(e.[Curriculum_Type],''))) = 'Revenue Cycle - Centralized'
  AND mu.UniversalID IS NULL;
GO

/* ---------- report.wave_missing_from_sources ----------
   Shows      : Master people missing from HR / Epic lookup / MVP (anti-join
                flags; one row per person with >=1 gap). SQL twin of
                wave_missing_from_hr_epic_mvp.py's three tabs.
   Built from : raw.master + raw.hr_cleaned + raw.epic_lookup + raw.mvp
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.wave_missing_from_sources AS
WITH hr AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(UniversalID))) AS u FROM raw.hr_cleaned
    WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
),
ep AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM([Universal_ID]))) AS u FROM raw.epic_lookup
    WHERE LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> ''
),
mv AS (
    SELECT DISTINCT UPPER(LTRIM(RTRIM(UniversalID))) AS u FROM raw.mvp
    WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
)
SELECT UPPER(LTRIM(RTRIM(mm.UniversalID)))       AS UniversalID,
       mm.[Full_Name]                            AS FullName,
       mm.[GoLiveWave]                           AS Wave,
       mm.[Leaders]                              AS Leader,
       mm.[Training_Needed_Yes_No]               AS TrainingNeeded,
       mm.[IsDepartedInactive]                   AS Departed,
       CASE WHEN hr.u IS NULL THEN 1 ELSE 0 END  AS NotInHR,
       CASE WHEN ep.u IS NULL THEN 1 ELSE 0 END  AS NotInEpic,
       CASE WHEN mv.u IS NULL THEN 1 ELSE 0 END  AS NotInMVP
FROM raw.[master] mm
LEFT JOIN hr ON hr.u = UPPER(LTRIM(RTRIM(mm.UniversalID)))
LEFT JOIN ep ON ep.u = UPPER(LTRIM(RTRIM(mm.UniversalID)))
LEFT JOIN mv ON mv.u = UPPER(LTRIM(RTRIM(mm.UniversalID)))
WHERE LTRIM(RTRIM(ISNULL(mm.UniversalID,''))) <> ''
  AND (hr.u IS NULL OR ep.u IS NULL OR mv.u IS NULL);
GO

/* ---------- report.master_validations ----------
   Shows      : the one-screen health check — every rule violation with a
                people count, including reference-list hygiene (the guard
                load_refs doesn't have) and the three population gaps.
   Built from : the QA views above + raw.ref_* + raw.mvp
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW report.master_validations AS
SELECT 'User Type differs — High confidence' AS Issue,
       RuleFired AS Value, COUNT(*) AS People,
       'report.user_type_recommendations' AS DetailView
FROM report.user_type_recommendations
WHERE Differs = 1 AND Confidence = 'High' GROUP BY RuleFired
UNION ALL
SELECT 'User Type differs — Medium confidence', RuleFired, COUNT(*),
       'report.user_type_recommendations'
FROM report.user_type_recommendations
WHERE Differs = 1 AND Confidence = 'Medium' GROUP BY RuleFired
UNION ALL
SELECT 'User Type needs review (conflict / no signal)', RuleFired, COUNT(*),
       'report.user_type_recommendations'
FROM report.user_type_recommendations
WHERE Confidence = 'Review' GROUP BY RuleFired
UNION ALL
SELECT 'User Type blank in Master', ISNULL(RuleFired,'(none)'), COUNT(*),
       'report.user_type_recommendations'
FROM report.user_type_recommendations
WHERE CurrentBlank = 1 GROUP BY RuleFired
UNION ALL
SELECT 'Training: Epic=Yes but Master=No', NULL, COUNT(*), 'report.training_needed_check'
FROM report.training_needed_check WHERE EpicYes_MasterNo = 1
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Training: Terminated (Epic) but Training=Yes', NULL, COUNT(*), 'report.training_needed_check'
FROM report.training_needed_check WHERE TerminatedButYes = 1
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Training: Pay Leave (Epic) but Training=Yes', NULL, COUNT(*), 'report.training_needed_check'
FROM report.training_needed_check WHERE PayLeaveButYes = 1
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Training: no-training role but Training=Yes', NULL, COUNT(*), 'report.training_needed_check'
FROM report.training_needed_check WHERE NoTrainingRoleButYes = 1
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Training: departed but Training=Yes', NULL, COUNT(*), 'report.training_needed_check'
FROM report.training_needed_check WHERE DepartedButYes = 1
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Gap: HR (tracked leaders) missing from wave', NULL, COUNT(*), 'report.hr_missing_from_wave'
FROM report.hr_missing_from_wave
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Gap: Epic Centralized missing from wave', NULL, COUNT(*), 'report.epic_missing_from_wave'
FROM report.epic_missing_from_wave
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Gap: wave people missing from a source', NULL, COUNT(*), 'report.wave_missing_from_sources'
FROM report.wave_missing_from_sources
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Ref hygiene: blank BusinessUnit key', NULL, COUNT(*), 'raw.ref_business_units'
FROM raw.ref_business_units WHERE LTRIM(RTRIM(ISNULL(BusinessUnit,''))) = ''
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Ref hygiene: BU needs mapping (IsVendor/UserType incomplete)',
       NULL, COUNT(*), 'raw.ref_business_units'
FROM raw.ref_business_units
WHERE LTRIM(RTRIM(ISNULL(IsVendor,''))) = '' OR LTRIM(RTRIM(ISNULL(UserType,''))) = ''
HAVING COUNT(*) > 0
UNION ALL
SELECT 'Ref hygiene: duplicate vendor alias', a.Alias, COUNT(*), 'raw.ref_vendor_aliases'
FROM (SELECT UPPER(LTRIM(RTRIM(Alias))) AS Alias FROM raw.ref_vendor_aliases
      WHERE LTRIM(RTRIM(ISNULL(Alias,''))) <> '') a
GROUP BY a.Alias HAVING COUNT(*) > 1
UNION ALL
SELECT 'Ref: MVP RCM BusinessUnit not in ref_business_units', v.bu, COUNT(DISTINCT v.uid), 'raw.mvp'
FROM (SELECT UPPER(LTRIM(RTRIM(BusinessUnit))) AS bu,
             UPPER(LTRIM(RTRIM(UniversalID)))  AS uid
      FROM raw.mvp
      WHERE UPPER(LTRIM(RTRIM(ISNULL(IsRCMUser,'')))) IN ('TRUE','YES','1')
        AND LTRIM(RTRIM(ISNULL(BusinessUnit,''))) <> '') v
LEFT JOIN raw.ref_business_units ref
       ON UPPER(LTRIM(RTRIM(ref.BusinessUnit))) = v.bu
WHERE ref.BusinessUnit IS NULL
GROUP BY v.bu;
GO

/* ---------- mvp.rcm_users — the clean RCM user list ----------
   Shows      : one row per MVP user with IsRCMUser=True (latest import),
                with BU, roles 1-4, and the parsed Offshore/Onshore tag.
   Built from : raw.mvp
   ---------------------------------------------------------------- */
CREATE OR ALTER VIEW mvp.rcm_users AS
SELECT UniversalID, FirstName, LastName, JobTitle, BusinessUnit,
       JobRole1, JobRole2, JobRole3, JobRole4, GoLiveWave, IsDepartedInactive,
       CASE WHEN HasOff = 1 AND HasOn = 1 THEN 'CONFLICT'
            WHEN HasOff = 1 THEN 'Offshore'
            WHEN HasOn  = 1 THEN 'Onshore' END AS ShoreTag,
       LastImportedDate
FROM (
    SELECT UPPER(LTRIM(RTRIM(UniversalID))) AS UniversalID,
           [FirstName], [LastName], [JobTitle], [BusinessUnit],
           [IndividualCategoryUpdate1Name] AS JobRole1,
           [IndividualCategoryUpdate2Name] AS JobRole2,
           [IndividualCategoryUpdate3Name] AS JobRole3,
           [IndividualCategoryUpdate4Name] AS JobRole4,
           [GoLiveWave], [IsDepartedInactive], [LastImportedDate],
           CASE WHEN UPPER(CONCAT(ISNULL([IndividualCategoryUpdate1Name],''),'|',
                                  ISNULL([IndividualCategoryUpdate2Name],''),'|',
                                  ISNULL([IndividualCategoryUpdate3Name],''),'|',
                                  ISNULL([IndividualCategoryUpdate4Name],'')))
                LIKE '%OFFSHORE%' THEN 1 ELSE 0 END AS HasOff,
           CASE WHEN UPPER(CONCAT(ISNULL([IndividualCategoryUpdate1Name],''),'|',
                                  ISNULL([IndividualCategoryUpdate2Name],''),'|',
                                  ISNULL([IndividualCategoryUpdate3Name],''),'|',
                                  ISNULL([IndividualCategoryUpdate4Name],'')))
                LIKE '%ONSHORE%' THEN 1 ELSE 0 END AS HasOn,
           ROW_NUMBER() OVER (PARTITION BY UPPER(LTRIM(RTRIM(UniversalID)))
                              ORDER BY [LastImportedDate] DESC) AS rn
    FROM raw.mvp
    WHERE UPPER(LTRIM(RTRIM(ISNULL(IsRCMUser,'')))) IN ('TRUE','YES','1')
      AND LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''
) x WHERE rn = 1;
GO

/* ---------- pbi aliases for the QA layer ---------- */
/* pbi.master_validations — Power BI alias of report.master_validations */
CREATE OR ALTER VIEW pbi.master_validations AS SELECT * FROM report.master_validations;
GO

PRINT 'dim.* views + report.mapping_gaps + Master QA layer created/updated.';
GO


/* ---------- report.epic_training_rollup — Epic status per person ------------
   Added      : 2026-08-26.
   Shows      : one row per Universal ID in raw.epic_status — the SAME
                aggregation report.tracker_training_status performs, including
                its IsExcluded rule (Advanced Reporting / Charge Capture are
                optional workshops), but WITHOUT the wave-file gate.
   Why        : report.tracker_training_status joins report.users WHERE
                IsInScope = 1 and Leader on the canonical list, so anyone not on
                the wave file gets no row at all. The LAVA SVC-provisioning
                group is exactly that: 356 people with real Epic registrations
                (348 of them on EPIC_NH_SIMPLE VISIT CODING ERROR RESOLUTION)
                who came back completely blank. Per the analyst 2026-08-26 they are
                tracked like anyone else — if Epic or Cornerstone has training
                for someone, it gets filled in; blank is only for people we
                truly have nothing on.
   Contract   : verified identical to report.tracker_training_status on all
                6,914 on-leader-list wave people for TrainingStatus,
                FullyRegisteredYN, FinalScheduledDate and the three class counts
                (counts match once its ISNULL(...,0) is applied). Consumers must
                prefer the tracker where it has a row and use this to fill the
                rest, so the two can never disagree.
   Used by    : refresh.build_lava_list() (report.lava_list) and
                scripts\build_lava_list.py — one definition, both consumers.
   ------------------------------------------------------------------------ */
CREATE OR ALTER VIEW report.epic_training_rollup AS
WITH src AS (
    SELECT UPPER(LTRIM(RTRIM(s.Universal_Id))) AS UniversalID,
           CASE WHEN s.Fully_Trained    LIKE '%Yes%' THEN 1 ELSE 0 END AS FT,
           CASE WHEN s.Fully_Registered LIKE '%Yes%' THEN 1 ELSE 0 END AS FR,
           s.Event_Class,
           s.Event_Class_Registered AS Registered,
           s.Event_Class_Status     AS Status,
           TRY_CONVERT(datetime, s.Event_Class_Date)  AS EventDate,
           TRY_CONVERT(datetime, s.Registration_Date) AS RegistrationDate,
           CASE WHEN s.Event_Class_Type = 'Session'
                     AND UPPER(s.Event_Class) NOT LIKE '%ADVANCED REPORTING%'
                     AND UPPER(s.Event_Class) NOT LIKE '%CHARGE CAPTURE%'
                     AND UPPER(ISNULL(s.Curriculum,'')) NOT LIKE '%ADVANCED REPORTING%'
                     AND UPPER(ISNULL(s.Curriculum,'')) NOT LIKE '%CHARGE CAPTURE%'
                THEN 1 ELSE 0 END AS IsSession
    FROM raw.epic_status s
    WHERE LTRIM(RTRIM(ISNULL(s.Universal_Id,''))) <> ''
)
SELECT UniversalID,
       MAX(FT) AS FT,
       MAX(FR) AS FR,
       CASE WHEN MAX(FT) = 1 THEN 'Fully Trained'
            WHEN COUNT(DISTINCT CASE WHEN IsSession = 1 AND Status LIKE 'Completed%'
                                     THEN Event_Class END) > 0 THEN 'In Progress'
            ELSE 'Not Started' END AS TrainingStatus,
       CASE WHEN MAX(FR) = 1 THEN 'Yes' ELSE 'No' END AS FullyRegisteredYN,
       CASE WHEN MAX(FT) = 1 THEN 'Yes' ELSE 'No' END AS FullyTrainedYN,
       COUNT(DISTINCT CASE WHEN IsSession = 1 THEN Event_Class END) AS ClassesRequired,
       COUNT(DISTINCT CASE WHEN IsSession = 1 AND Status LIKE 'Completed%'
                           THEN Event_Class END) AS ClassesCompleted,
       COUNT(DISTINCT CASE WHEN IsSession = 1
                            AND (Registered LIKE '%Yes%' OR Status LIKE 'Completed%')
                           THEN Event_Class END) AS ClassesRegistered,
       MAX(CASE WHEN IsSession = 1 AND Status LIKE 'Completed%'
                THEN EventDate END) AS LastAttendedDate,
       MAX(CASE WHEN IsSession = 1
                 AND (Registered LIKE '%Yes%' OR Status LIKE 'Completed%')
                THEN EventDate END) AS FinalScheduledDate,
       MAX(CASE WHEN IsSession = 1
                 AND (Registered LIKE '%Yes%' OR Status LIKE 'Completed%')
                 AND Status LIKE '%Equivalent%' THEN 1 ELSE 0 END) AS HasEquiv
FROM src
GROUP BY UniversalID;
GO

PRINT 'report.epic_training_rollup created/updated.';
GO

/* ---------- report.lava_list — LAVA Wave census (working LAVA list) ------------------
   Added      : 2026-08-26 (the analyst's ask — LAVA queryable daily without having
                to run the script).
   Population : the union of
                  (a) WAVE — tagged for training on the wave file
                             (report.users.TrainingNeeded = 'Yes'), and
                  (b) SVC  — holds a "Simple Visit Coding" job role in MVP AND
                             appears in the Epic "Curriculum Status by User
                             Summary" export, i.e. someone we actually have Epic
                             status data on.
                Filter by wave at the call site: WHERE Wave = 'Wave 3'.
   SVC source : per the analyst 2026-08-26, SVC is NOT read from the ad_hoc
                "Curriculum Status by User - SVC.xlsx" — that file was only the
                summary export pre-filtered, and it goes stale. Same derivation
                build_soft_live_list.py already uses.
   Built by   : refresh.build_lava_list() -> raw.lava_person, refreshed with
                    python sql
efresh.py lava
                It is assembled in stages in Python rather than expressed as one
                view because as a single view it NEVER RETURNED: every join is
                individually fast (~6s all told) but SQL Server could not plan 12
                joins off a UNION-derived driving set — 27 table scans, 27 sorts,
                and it refused to prune the LEFT JOINs, so projecting one column
                cost 41s. Materialising the driving set first takes 8s.
   NOT here   : FEC / Soft Live / Workqueue Owner as *typed* answers — those are
                hand-entered in the workbook and carried forward on rebuild, so
                no view can hold them. Only the system-of-record values
                (report.users + wave change requests) appear here; Workqueue
                Owner has no system of record at all.
   ------------------------------------------------------------------------ */
-- On a fresh database this file runs before anything has been loaded, so the
-- table may not exist yet. Create an empty stand-in so the view is always
-- creatable; refresh.py's sp_refreshview pass re-binds the view to the real
-- column list the first time `refresh.py lava` populates it.
IF OBJECT_ID('raw.lava_person') IS NULL
    EXEC('SELECT CAST(NULL AS NVARCHAR(450)) AS UniversalID INTO raw.lava_person WHERE 1 = 0');
GO

CREATE OR ALTER VIEW report.lava_list AS
SELECT * FROM raw.lava_person;
GO

PRINT 'report.lava_list created/updated.';
GO

/* ---------- report.svc_population — the SVC (Simple Visit Coding) user population -----
   Added      : 2026-09-15 (the analyst's ask — "is there a SQL view for just the SVC user
                population? if not can you create one").
   Grain      : ONE ROW PER PERSON, every wave. Filter at the call site, e.g.
                  WHERE Wave = 'Wave 3'
   Rule       : the census SVC rule — holds a Simple Visit Coding job role (in MVP, or
                assigned on a Wave Change Request Form that MVP does not carry yet —
                SVCSource says which) AND appears in the Epic Curriculum Status by
                User Summary export (we have Epic status data on them). Being in the
                wave file is NOT required: OnWaveFile / Population show who is.
   Source     : report.lava_list (raw.lava_person, rebuilt by `python sql/refresh.py lava`
                and by the daily `lava` step). Same rows the LAVA census workbook tags
                SVC Provisioning / SVC + WQ Owner / Wave + SVC.
   Current    : IsCurrent = 0 when HR or Epic shows leave / termination (Larry's rule
                2026-09-15 — the census counts only current people); add
                WHERE IsCurrent = 1 AND (Wave = 'Wave 3' OR OnWave3DNFBList = 'Yes')
                to match the census exactly.
   ------------------------------------------------------------------------ */
CREATE OR ALTER VIEW report.svc_population AS
SELECT UniversalID, FullName, Wave, Leader, AVP, VP, OnWaveFile, Population,
       SVCJobRole, SVCSource, JobRole1, JobRole2, JobRole3, JobRole4,
       IsGuesthouse, VendorYN, UserType, JobTitle, Department, BusinessUnitDescription,
       Email, DirectManager, TrainingStatus, FullyRegistered, FullyTrained,
       EstCompletionDate, FECAccess, SoftLiveAccess, IsInScope, OnLeaderList,
       HRStatus, EpicHRStatus, OnWave3DNFBList, IsCurrent
FROM report.lava_list
WHERE SVCYN = 'Yes';
GO

PRINT 'report.svc_population created/updated.';
GO

/* ---------- report.dnfb_workqueue_owners — who owns a DNFB workqueue ---------------
   Added      : 2026-09-09 (the analyst's ask — "an SQL view for all DNFB workqueue owners").
   Grain      : ONE ROW PER OWNER (person), every wave. Filter at the call site:
                  WHERE OnWave3DNFBList = 'Yes'     -> the boss's Wave 3 DNFB list (71)
                  WHERE DNFBWorkqueuesWave3 > 0     -> owns any Wave 3 DNFB workqueue (85)
   Source     : raw.epic_workqueue_ownership = the "Epic Workqueue Ownership (83)"
                tab of data\raw\ad_hoc\DNFB Owners for LAva*.xlsx (newest copy), loaded
                daily through source_registry.xlsx. A DNFB workqueue is a row whose
                Specialty Grouper = 'DNFB'; the owner is the email in
                "WQ Owner (Supervisor) (Input)" — the file carries NO Universal ID.
   Owner list : the file's hand-made "Wave 3 DNFB Owners" tab is NOT read. It is
                exactly the distinct owners of Wave 3 DNFB workqueues whose Workgroup
                Owner Name is not 'Simple Visit Coding' (verified identical, 71 = 71,
                2026-09-09); OnWave3DNFBList reproduces it so a fresh export
                keeps it current without anyone re-pasting the tab.
   Identity   : email -> UniversalID via MVP UserUPN (508 of 551), else HR
                Email_Address (the other 43 — local-part <> UID, e.g. a_goldberg2),
                else the email local-part. IdSource says which. Name / leader chain
                come from the Master (report.users) when the person is on the wave
                file, otherwise from HR (report.hr) — most DNFB owners are NOT rev
                cycle under Dr Lastname03, so most are HR-only.
   Waves      : WaveFileWave = the Master's GoLiveWave (wave-file people only);
                MVPWave = MVP GoLiveWave (everyone). Off-wave people still own
                Wave 3 DNFB workqueues on purpose (her boss, 2026-09-02).
   ------------------------------------------------------------------------ */
-- On a fresh database this file runs before anything has been loaded, so the
-- table may not exist yet. Create an empty stand-in with the columns the view
-- reads so the view is always creatable; `refresh.py sources` replaces it.
IF OBJECT_ID('raw.epic_workqueue_ownership') IS NULL
    EXEC('SELECT CAST(NULL AS NVARCHAR(4000)) AS WQ_Name, CAST(NULL AS NVARCHAR(4000)) AS Wave,
                 CAST(NULL AS NVARCHAR(4000)) AS Specialty_Grouper,
                 CAST(NULL AS NVARCHAR(4000)) AS Workgroup_Owner_Name,
                 CAST(NULL AS NVARCHAR(4000)) AS WQ_Owner_Supervisor_Input,
                 CAST(NULL AS NVARCHAR(4000)) AS Core_Y_is_Rev_Cycle_Dr_Brogan,
                 CAST(NULL AS NVARCHAR(4000)) AS Site,
                 CAST(NULL AS NVARCHAR(4000)) AS Sim_New_Grouper_4,
                 CAST(NULL AS NVARCHAR(4000)) AS _source_file
          INTO raw.epic_workqueue_ownership WHERE 1 = 0');
GO

CREATE OR ALTER VIEW report.dnfb_workqueue_owners AS
WITH wq AS (                    -- every DNFB workqueue with an owner email
    SELECT LOWER(LTRIM(RTRIM(WQ_Owner_Supervisor_Input)))            AS OwnerEmail,
           LTRIM(RTRIM(Wave))                                        AS WQWave,
           LTRIM(RTRIM(WQ_Name))                                     AS WQName,
           LTRIM(RTRIM(Workgroup_Owner_Name))                        AS Workgroup,
           LTRIM(RTRIM(Site))                                        AS Site,
           LTRIM(RTRIM(Sim_New_Grouper_4))                           AS OwningGroup,
           UPPER(LTRIM(RTRIM(Core_Y_is_Rev_Cycle_Dr_Brogan)))        AS Core,
           _source_file                                              AS SourceFile
    FROM raw.epic_workqueue_ownership
    WHERE UPPER(LTRIM(RTRIM(Specialty_Grouper))) = 'DNFB'
      AND WQ_Owner_Supervisor_Input LIKE '%@%'
),
owners AS (SELECT DISTINCT OwnerEmail FROM wq),
-- Email -> UniversalID. Join the ~550 owner emails INTO the big raw tables and
-- collapse afterwards; ROW_NUMBER over all of raw.mvp / raw.hr first is the
-- pattern that never finishes (see raw.mvp_person).
via_mvp AS (
    SELECT o.OwnerEmail, MIN(UPPER(LTRIM(RTRIM(m.UniversalID)))) AS UniversalID
    FROM owners o
    JOIN raw.mvp m ON LOWER(LTRIM(RTRIM(m.UserUPN))) = o.OwnerEmail
    WHERE LTRIM(RTRIM(ISNULL(m.UniversalID, ''))) <> ''
    GROUP BY o.OwnerEmail
),
via_hr AS (
    SELECT o.OwnerEmail, MIN(UPPER(LTRIM(RTRIM(h.UniversalID)))) AS UniversalID
    FROM owners o
    JOIN raw.hr h ON LOWER(LTRIM(RTRIM(h.Email_Address))) = o.OwnerEmail
    WHERE LTRIM(RTRIM(ISNULL(h.UniversalID, ''))) <> ''
    GROUP BY o.OwnerEmail
),
ids AS (
    SELECT o.OwnerEmail,
           COALESCE(mv.UniversalID, hr.UniversalID,
                    UPPER(LEFT(o.OwnerEmail, CHARINDEX('@', o.OwnerEmail) - 1))) AS UniversalID,
           CASE WHEN mv.UniversalID IS NOT NULL THEN 'MVP'
                WHEN hr.UniversalID IS NOT NULL THEN 'HR'
                ELSE 'Email local-part' END                                     AS IdSource
    FROM owners o
    LEFT JOIN via_mvp mv ON mv.OwnerEmail = o.OwnerEmail
    LEFT JOIN via_hr  hr ON hr.OwnerEmail = o.OwnerEmail
),
agg AS (
    SELECT OwnerEmail,
           COUNT(*)                                                    AS DNFBWorkqueues,
           SUM(CASE WHEN WQWave = 'Wave 1' THEN 1 ELSE 0 END)          AS DNFBWorkqueuesWave1,
           SUM(CASE WHEN WQWave = 'Wave 2' THEN 1 ELSE 0 END)          AS DNFBWorkqueuesWave2,
           SUM(CASE WHEN WQWave = 'Wave 3' THEN 1 ELSE 0 END)          AS DNFBWorkqueuesWave3,
           SUM(CASE WHEN WQWave = 'Wave 3'
                     AND Workgroup <> 'Simple Visit Coding' THEN 1 ELSE 0 END) AS Wave3NonSVC,
           MAX(CASE WHEN Core = 'Y' THEN 1 ELSE 0 END)                 AS CoreRevCycle,
           MAX(SourceFile)                                             AS SourceFile
    FROM wq GROUP BY OwnerEmail
),
wave_list AS (
    SELECT OwnerEmail, STRING_AGG(CAST(WQWave AS NVARCHAR(MAX)), ', ') WITHIN GROUP (ORDER BY WQWave) AS v
    FROM (SELECT DISTINCT OwnerEmail, WQWave FROM wq WHERE WQWave IS NOT NULL) d GROUP BY OwnerEmail
),
group_list AS (
    SELECT OwnerEmail, STRING_AGG(CAST(Workgroup AS NVARCHAR(MAX)), ', ') WITHIN GROUP (ORDER BY Workgroup) AS v
    FROM (SELECT DISTINCT OwnerEmail, Workgroup FROM wq WHERE Workgroup IS NOT NULL) d GROUP BY OwnerEmail
),
site_list AS (
    SELECT OwnerEmail, STRING_AGG(CAST(Site AS NVARCHAR(MAX)), ', ') WITHIN GROUP (ORDER BY Site) AS v
    FROM (SELECT DISTINCT OwnerEmail, Site FROM wq WHERE Site IS NOT NULL) d GROUP BY OwnerEmail
),
wq_list AS (
    SELECT OwnerEmail, STRING_AGG(CAST(WQName AS NVARCHAR(MAX)), ' | ') WITHIN GROUP (ORDER BY WQName) AS v
    FROM (SELECT DISTINCT OwnerEmail, WQName FROM wq WHERE WQName IS NOT NULL) d GROUP BY OwnerEmail
)
SELECT
    i.UniversalID,
    COALESCE(u.FullName, h.FullName)                              AS FullName,
    i.OwnerEmail                                                  AS Email,
    i.IdSource,
    CASE WHEN u.UniversalID IS NOT NULL THEN 'Yes' ELSE 'No' END  AS OnWaveFile,
    u.Wave                                                        AS WaveFileWave,
    mp.GoLiveWave                                                 AS MVPWave,
    COALESCE(u.Leader, h.Leader)                                  AS Leader,
    COALESCE(u.AVP, h.AVP)                                        AS AVP,
    COALESCE(u.VP, h.VP)                                          AS VP,
    h.JobTitle,
    h.Department,
    CASE WHEN a.CoreRevCycle = 1 THEN 'Y' ELSE 'N' END            AS CoreRevCycle,
    CASE WHEN a.Wave3NonSVC > 0 THEN 'Yes' ELSE 'No' END          AS OnWave3DNFBList,
    a.DNFBWorkqueues,
    a.DNFBWorkqueuesWave1,
    a.DNFBWorkqueuesWave2,
    a.DNFBWorkqueuesWave3,
    wl.v                                                          AS WorkqueueWaves,
    gl.v                                                          AS Workgroups,
    sl.v                                                          AS Sites,
    ql.v                                                          AS Workqueues,
    a.SourceFile
FROM ids i
JOIN agg a               ON a.OwnerEmail  = i.OwnerEmail
LEFT JOIN wave_list  wl  ON wl.OwnerEmail = i.OwnerEmail
LEFT JOIN group_list gl  ON gl.OwnerEmail = i.OwnerEmail
LEFT JOIN site_list  sl  ON sl.OwnerEmail = i.OwnerEmail
LEFT JOIN wq_list    ql  ON ql.OwnerEmail = i.OwnerEmail
LEFT JOIN report.users u ON u.UniversalID = i.UniversalID
LEFT JOIN report.hr    h ON h.UniversalID = i.UniversalID
LEFT JOIN raw.mvp_person mp ON mp.UniversalID = i.UniversalID;
GO

PRINT 'report.dnfb_workqueue_owners created/updated.';
GO
