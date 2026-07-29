-- =========================================================================
-- SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
-- Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
-- fill in your own values before running. See README.md for the full map.
-- =========================================================================
/* ============================================================
   1_setup.sql  —  RUN ONCE
   Creates the AnalyticsDB database and the two schemas.

   What the two schemas mean:
     raw      = a faithful copy of your export files (HR, MVP,
                Cornerstone, Epic, Master Wave). Reloaded daily.
     report   = clean views that do the joining + cleaning.
                This is what Power BI / Excel reads.

   How to run:
     sqlcmd -S ".\SQLEXPRESS" -E -C -i "sql\1_setup.sql"
   ============================================================ */

IF DB_ID('AnalyticsDB') IS NULL
BEGIN
    CREATE DATABASE AnalyticsDB;
END
GO

USE AnalyticsDB;
GO

IF SCHEMA_ID('raw') IS NULL EXEC('CREATE SCHEMA raw');
GO
IF SCHEMA_ID('report') IS NULL EXEC('CREATE SCHEMA report');
GO

PRINT 'Setup complete: database AnalyticsDB with schemas [raw] and [report].';
GO
