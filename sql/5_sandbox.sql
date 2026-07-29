-- =========================================================================
-- SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
-- Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
-- fill in your own values before running. See README.md for the full map.
-- =========================================================================
/* ============================================================
   5_sandbox.sql — RUN ONCE. Your personal space for hand-built views.

   THE RULES (simple on purpose):
     1. Everything YOU create by hand goes in  sandbox.  — nothing else.
          CREATE OR ALTER VIEW sandbox.my_view AS SELECT ... ;
     2. You can SELECT from anything (raw.*, report.*, dim.*, history.*)
        and join freely — reading never breaks the pipeline.
     3. Never create objects in raw/report/dim/history/pbi by hand — those
        belong to refresh.py and the numbered .sql scripts.
     4. dbo stays empty. A guard below blocks accidental dbo creates with a
        reminder (old habits die hard).
     5. Your hand-built views (sandbox + the per-source schemas) are auto-backed
        up on every daily refresh to data\references\user_views_backup.sql
        (OneDrive) — if the local DB is ever rebuilt, run that file to restore
        them all.
     6. Source-SPECIFIC views are even better placed in the per-source schemas
        (epic. / mvp. / cornerstone. / hr. / wave.) — see sql\SOURCES.md.
        sandbox. is for cross-source experiments and one-offs.

   HOW TO CREATE A VIEW (either way works):
     - Azure Data Studio / SSMS: connect to .\SQLEXPRESS -> AnalyticsDB,
       paste your CREATE OR ALTER VIEW sandbox.<name> AS ... and run it.
     - Command line:
       sqlcmd -S ".\SQLEXPRESS" -E -C -d AnalyticsDB -Q "CREATE OR ALTER VIEW sandbox.x AS SELECT ..."

   TIPS:
     - Always use CREATE OR ALTER VIEW — rerunnable, no 'already exists' errors.
     - Views keep working after the nightly raw.* reloads (they re-resolve at
       query time). But avoid SELECT * in a view you keep — it freezes the
       column list at creation; name your columns.
     - If a sandbox view turns out to be a keeper everyone needs, graduate it:
       move the SQL into 3_report_views.sql as report.<name> so it's part of
       the managed set.

   How to run this file:
     sqlcmd -S ".\SQLEXPRESS" -E -C -d AnalyticsDB -i "sql\5_sandbox.sql"
   ============================================================ */
USE AnalyticsDB;
GO
-- The guard trigger uses XML methods; QUOTED_IDENTIFIER must be ON at its
-- CREATE time (the setting is saved with the module) — sqlcmd defaults it OFF.
SET QUOTED_IDENTIFIER ON;
SET ANSI_NULLS ON;
GO

IF SCHEMA_ID('sandbox') IS NULL EXEC('CREATE SCHEMA sandbox');
GO

/* ---------- Guard: nothing gets created in dbo by accident ----------
   If you ever hit this on purpose and want it gone:
     DROP TRIGGER trg_no_dbo_objects ON DATABASE;                       */
CREATE OR ALTER TRIGGER trg_no_dbo_objects ON DATABASE
FOR CREATE_TABLE, CREATE_VIEW, CREATE_PROCEDURE, CREATE_FUNCTION
AS
BEGIN
    IF EVENTDATA().value('(/EVENT_INSTANCE/SchemaName)[1]', 'sysname') = N'dbo'
    BEGIN
        RAISERROR('Nothing lives in dbo here. Your hand-built objects go in the sandbox schema:  CREATE OR ALTER VIEW sandbox.my_view AS ...   (see sql\5_sandbox.sql; to remove this guard: DROP TRIGGER trg_no_dbo_objects ON DATABASE)', 16, 1);
        ROLLBACK;
    END
END;
GO

/* ---------- A worked example — open it, change it, make it yours ----------
   In-scope people not yet fully registered, with leader + canonical vendor.
   Shows the two most useful join patterns:
     - roster -> raw.master (to reach a Master column the roster doesn't carry)
     - business unit -> dim.vendor_alias (variant spelling -> one vendor name) */
CREATE OR ALTER VIEW sandbox.example_not_registered AS
SELECT r.UniversalID,
       r.FullName,
       r.Wave,
       r.Leader,
       r.VendorYN,
       va.CanonicalVendor,
       r.JobRole1,
       r.CornerstoneRegistered
FROM report.roster r
LEFT JOIN raw.[master] m
       ON UPPER(LTRIM(RTRIM(m.UniversalID))) = r.UniversalID
LEFT JOIN dim.vendor_alias va
       ON va.Alias = UPPER(LTRIM(RTRIM(m.BusinessUnit)))
WHERE r.IsInScope = 1
  AND ISNULL(r.IsFullyRegistered, 0) = 0;
GO

PRINT 'sandbox schema ready: create your views as sandbox.<name> - dbo guard active.';
GO
