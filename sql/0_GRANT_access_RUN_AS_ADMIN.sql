-- =========================================================================
-- SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
-- Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
-- fill in your own values before running. See README.md for the full map.
-- =========================================================================
/* ============================================================
   0_GRANT_access_RUN_AS_ADMIN.sql
   FOR IT / A SQL SERVER ADMINISTRATOR — run ONCE.

   the analyst (AzureAD\YourUser) needs to build a small local
   analytics database on the SQLEXPRESS instance but currently has
   only read access. This grants the minimum needed.

   Run it connected to .\SQLEXPRESS as a sysadmin
   (e.g. in SSMS, or:  sqlcmd -S ".\SQLEXPRESS" -E -C -i this_file.sql)
   ============================================================ */

-- Ensure a login exists for the account
IF SUSER_ID('AzureAD\YourUser') IS NULL
    CREATE LOGIN [AzureAD\YourUser] FROM WINDOWS;

-- Let her create and fully own her own databases (she becomes owner of
-- anything she creates; she does NOT get rights over other databases).
ALTER SERVER ROLE dbcreator ADD MEMBER [AzureAD\YourUser];

PRINT 'Granted: AzureAD\YourUser can now create/own databases on this instance.';

/* ---- ALTERNATIVE (even more restrictive) ----
   If IT prefers NOT to grant dbcreator, they can instead create the
   database themselves and make her owner of just that one DB:

   CREATE DATABASE AnalyticsDB;
   IF SUSER_ID('AzureAD\YourUser') IS NULL
       CREATE LOGIN [AzureAD\YourUser] FROM WINDOWS;
   USE AnalyticsDB;
   CREATE USER [AzureAD\YourUser] FOR LOGIN [AzureAD\YourUser];
   ALTER ROLE db_owner ADD MEMBER [AzureAD\YourUser];
*/
