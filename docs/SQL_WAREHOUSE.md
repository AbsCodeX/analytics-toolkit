# AnalyticsDB — local SQL Server warehouse

> New here or need a quick orientation? **`WHAT_IS_WHAT.md`** is the one-page
> map of every file and schema in this folder.

One place to pull training, registration, wave, team-member, and HR data — and to
build views / left-joins on demand. Feeds Power BI **and** ad-hoc questions.
**Two schemas, a daily load, and a handful of friendly views.**

```
OneDrive data\ ──(refresh.py)──▶ raw.*  ──(views join + clean)──▶ report.*  ──▶ Power BI / Excel / ad-hoc SQL
  HR, MVP,                        exact                           clean +
  Cornerstone,                    NVARCHAR                        typed +
  Epic, Master                    copies                          joined
```

Source files are read via the shared path helpers in
`scripts/onedrive_paths.py` (newest of each type) — never hand-build a path.
Two roots since 2026-08-11:

- **OneDrive** `C:\path\to\shared\workspace\data\`
  — every raw export (HR, MVP, Cornerstone, Epic) and everything shared.
- **Local** `C:\path\to\analytics\data\` — the Master
  Wave File and its backups/snapshots only (`op.MASTER_WAVE_PATH`).

## The schemas — what goes where

| Schema | What it is | You edit it? |
|---|---|---|
| `raw` | A faithful text copy of each source file. Dropped & rebuilt every refresh. | No — the loader owns it |
| `history` | Append-only daily snapshots (trend/change/event queries). | No — the snapshot step owns it |
| `report` | Clean views that type + join the raw data. What everything reads. | Yes — in `3_report_views.sql` |
| `dim` | Conformed dimensions for mapping/joins. | Yes — in `4_dimensions.sql` + the reference workbook |
| `pbi` | Thin aliases: the Power BI surface. | Rarely — in `3_report_views.sql`/`4_dimensions.sql` |
| `cornerstone` `epic` `mvp` `hr` `wave` | Per-source schemas: each system's exports under clean names + your own source-specific views. New exports self-register via the source registry — see `SOURCES.md`. | Yes — your views live here too |
| `sandbox` | **Yours.** Cross-source hand-built views, experiments, one-offs. Auto-backed up daily. | Yes — freely, any tool (`5_sandbox.sql` has the rules) |
| `dbo` | Deliberately empty; a guard trigger blocks accidental creates there. | Never |

## What loads (hybrid: raw exports + cleaned HR)

| raw table | Source (newest of type, from OneDrive) | Notes |
|---|---|---|
| `raw.master` | `data\reports\Main Reports\RCM Wave Data Master File.xlsx` (`DATA`) | roster of record; falls back to newest `DATA_*.xlsx` snapshot if the file is open in Excel |
| `raw.hr` | `data\raw\hr\HR.xlsx` | full 67-col HR detail |
| `raw.hr_cleaned` | newest `data\processed\hr\clean_hr_copies\HR_cleaned_*.xlsx` (`All Staff`) | resolved leader chain (AVP/VP/SVP/Leaders) |
| `raw.mvp` | `data\raw\mvp\User MappingsData.csv` | job roles |
| `raw.mvp_roles` | `data\raw\mvp\RoleMappingsData.csv` | dept/job-code → Epic job-category mapping |
| `raw.mvp_job_categories` | `data\raw\mvp\JobCategoriesData.csv` | Epic job-category reference |
| `raw.cornerstone` | newest (by modified time) `data\raw\cornerstone\...\Enterprise_Training_Report_*.xlsx` | transcripts (header row 17); `archive/` ignored |
| `raw.epic_status` | `data\raw\epic\epic_status_details\Curriculum Status Detail by User - W*.xlsx` | per-user Fully Trained/Registered, per wave |
| `raw.epic_lookup` | `data\raw\epic\epic_tm_lookup\Epic Team Member Lookup.xlsx` | Wave, training-needed, eligibility |
| `raw.epic_class_schedule` | newest `data\raw\epic\epic_class_schedules\epic_class_schedule_*.xlsx` | sessions: date/location/instructor/seats |
| `raw.wave_change_requests` | ALL `data\raw\wave\wave_change_requests\Wave Change Request Form*.xlsx` (`Wave Change Request` tab, combined) | add/edit/remove requests; archive a form once handled |

> **Retired 2026-07-10:** the "Unregistered Sessions by User" and "Epic Status
> Summary" exports — both were derived from the TM Lookup / status detail.
> `report.unregistered` now derives from `raw.epic_status`
> (`Event_Class_Registered = 'No'`).

> **Epic coverage:** whichever `Curriculum Status Detail by User - W*.xlsx` files are
> present get loaded (newest per wave). Drop in W1/W2/W4 and they load automatically —
> until then, Epic training numbers only cover the waves you have.

## How to run

**First time (run once):**
```powershell
sqlcmd -S ".\SQLEXPRESS" -E -C -i "sql\1_setup.sql"                                # create DB + schemas
python sql\refresh.py                                                              # load all sources -> raw.*
sqlcmd -S ".\SQLEXPRESS" -E -C -d AnalyticsDB -i "sql\3_report_views.sql"   # build views
```

**Every day after** — largely automatic since 2026-07-07: the
`YourOrgSQLAutoRefresh` scheduled task runs `sql\auto_refresh.py` weekdays at
9:00 / 11:00 / 13:00 / 14:30. It reloads only the sources whose files changed,
re-applies `3_report_views.sql` / `4_dimensions.sql` when you edit them, and
re-stamps today's history snapshot when data changed. Manual runs are still
fine any time (close the Master in Excel first, or it uses yesterday's snapshot):
```powershell
python sql\refresh.py             # full reload, all sources
python sql\auto_refresh.py        # what the scheduled task runs (only changes)
```

Load just one source: `python sql\refresh.py mvp epic_status`
Re-run `3_report_views.sql` yourself only if you don't want to wait for the
next auto-refresh cycle to pick up your edit.

**Object-level reference:** `DATA_DICTIONARY.md` — every table and view with
row counts, source provenance, lineage (what each view reads), and full column
lists. Auto-generated; regenerates with the auto-refresh, or on demand with
`python sql\build_data_dictionary.py`.

## History (daily trend & change tracking)

A full `refresh.py` run ends with a **snapshot** step that stamps today's state
into append-only `history.*` tables — the raw material for trend and
day-over-day change queries (`report.registration_daily`, `report.daily_changes`):

- `history.roster_daily` — the whole `report.roster`, person-level, one set of
  rows per day. Person-level (not aggregated) on purpose: history can be sliced
  by leader, vendor, wave, or any category you think of **later**.
- `history.epic_not_in_hr_daily` — Epic team members with no HR row, per day —
  lastname13 the FirstSeen/NewToday flags on the boss's gap report.

Rules of the road:

- **One snapshot per calendar day** — rerunning the refresh the same day just
  replaces that day's rows (last run wins). Days you don't run, history has a
  gap; the change views compare consecutive snapshots, so nothing breaks.
- Each day's rows are mirrored to a dated CSV on OneDrive
  (`data\processed\history_snapshots\<table>\<YYYY-MM>\`). That's the durable
  copy — the local DB stays disposable. Lost or rebuilt the DB?
  `python sql\refresh.py backload_history` rebuilds `history.*` from the CSVs.
- `python sql\refresh.py snapshot` re-stamps today without reloading sources.

## The views (what to query)

| View | Answers | Population |
|---|---|---|
| `report.roster` ⭐ | **The everything view** — identity, wave, leader, job role, training, registration, HR, Epic-lookup — one wide row per person. Add a `WHERE`. | everyone on the Master |
| `report.users` | Clean roster (identity + wave + scope flag) | everyone on the Master |
| `report.training` | Registered / trained per user (Epic + Cornerstone combined) | anyone with training data |
| `report.hr` | HR org info: resolved leaders + dept / title / worker type | all HR staff |
| `report.training_status` | Fully Trained / Registered per user (Epic only) | Epic users |
| `report.cornerstone_status` | Registrations & completions per user | Cornerstone users |
| `report.kpi_summary` | Top-line % cards (denominator = in-scope) | in-scope |
| `report.leader_summary` | Rollup by leader | in-scope |
| `report.exceptions` | Data-quality watchlist (in-scope but missing from Epic / HR) | in-scope |
| `report.master_vs_sources` | **Proposed Master updates** — Master value vs source, with `*_Differs` flags. Read-only preview. | in-scope |
| `report.registration_daily` | Daily in-scope / registered / trained counts by wave × leader × vendor × user type (from history) | one row set per snapshot day |
| `report.daily_changes` | Who flipped day-over-day: became registered/trained, wave or leader moved | people whose state changed |
| `report.person_events` | Lifecycle event log: added, wave/job-role/leader changed, departed/LOA, dropped off | one row per person per event |
| `report.epic_not_in_hr` | Epic Team Member Lookup people with no HR row (daily boss gap report) | Epic-not-in-HR population |
| `report.noshow` | Cornerstone No Show transcripts, roster-enriched | no-show rows |
| `report.unregistered` | Unregistered sessions per user, roster-enriched | unregistered rows |
| `report.w3_unregistered` / `report.w3_noshow` | Wave 3 cuts of the two above | W3 detail rows |
| `report.w3_noshow_status` | No-shows resolved: standing / re-registered / completed after | one row per W3 person × no-showed class |
| `report.w3_leader_snapshot` | TODAY's W3 rollup by leader (feeds the daily history stamp) | one row per leader |
| `report.w3_registration_summary` | W3 daily trend by leader from history.w3_leader_daily (2026-07-15 →) | one row per day × leader |
| `report.w3_registration_metrics` | W3 registration metrics by leader: counts, %s, ≥50%/≥80% registered buckets | one row per leader |
| `report.w3_registration_kpis` | The same, ONE row for all of Wave 3 (the TOTAL line) | one row |
| `report.w3_track_class_map` | Job role → training track → class with progress + future capacity (open seats) | one row per role × track × class |
| `report.w3_training_progress` | Person-level progress across BOTH modalities: classroom sessions + online CBT (with scores) | one row per W3 in-scope person |
| `report.w3_watchlist` | Cross-system issues worth acting on (7 checks: no curriculum, seat waste, Epic↔Cornerstone sync, trained-flag disagreements, stalled no-shows) | one row per person per issue |

The Wave 3 set is published daily by `scripts\export_w3_registration_reports.py`
(daily_refresh apply, step `w3_reports`) to `data\reports\registration_daily\`:
one workbook + the rolling `W3 Leader Daily History.csv` look-back.

There is also a **`pbi` schema** — thin aliases over the views above plus run
history, as the plug-and-play Power BI surface. See **`POWERBI.md`**.

## The dimension layer (`dim.*`) — `4_dimensions.sql`

Conformed dimensions for mapping and left/right joins across the five source
vocabularies. Backed by the observed `raw.*` values plus the **editable**
mappings in `data\references\wave_reference_lists.xlsx` (edit the workbook →
`python sql\refresh.py refs` → done).

| View | What it gives you |
|---|---|
| `dim.vendor_alias` | Variant spelling → canonical vendor (e.g. 8 Getix spellings → one name). Join any BU through it. |
| `dim.vendor` | Canonical vendors with Master headcount + offshore/onshore split |
| `dim.business_unit` | Every BU observed in Master/HR/Cornerstone/Epic with per-source flags, vendor mapping, `NeedsMapping` |
| `dim.job_title` | Every job title observed across sources + `IsLeaderTitle` |
| `dim.department` | Departments observed in Master/HR/MVP with per-source flags |
| `dim.leader` | Canonical leaders + training preference + Master headcount + `NeedsReview` for non-canonical values |
| `dim.job_role` | Master roles 1–4 + MVP roles + `NoTrainingNeeded` flag |
| `dim.epic_job_category` | MVP's Epic job-category reference, clean columns |
| `report.mapping_gaps` | The actionable conformance issues (feeds the Morning Review "Mapping Gaps" sheet) |

Note the four BU vocabularies: Master/ref = vendor & org names; HR = facility
names; Cornerstone = cost-center strings; Epic Lookup = org groupings. They are
different by nature — `dim.business_unit` catalogs all of them; only Master
values are expected to map to `ref_business_units`.

`IsInScope` is a **column, not a filter** on `roster`/`users` — so you can pull everyone
and narrow with `WHERE IsInScope = 1` when you want just the training population.

## Ad-hoc pulls — three ways

1. **`query.py`** — quick one-off from PowerShell:
   ```powershell
   python sql\query.py "SELECT * FROM report.roster WHERE Wave='Wave 3' AND IsFullyRegistered=0" --excel not_registered_w3.xlsx
   ```
   Bare `--csv`/`--excel` filenames save to the shared OneDrive pulls folder
   (`data\reports\sql_pulls\`); pass a full path to save elsewhere.
   See **`QUERIES.md`** for copy-paste examples.
2. **Azure Data Studio / SSMS** — connect to `.\SQLEXPRESS` → `AnalyticsDB`, write SQL against `report.*`.
3. **Excel / Power BI** — Get Data → SQL Server → Server `.\SQLEXPRESS`, Database `AnalyticsDB` → pick a `report.*` view. No more intermediate export files.

## Files

| File | When |
|---|---|
| `1_setup.sql` | once |
| `refresh.py` | daily |
| `3_report_views.sql` | once, and after editing views |
| `query.py` | anytime — ad-hoc pulls to screen / CSV / Excel |
| `QUERIES.md` | reference — ready-made pulls |

## Notes
- Join key everywhere is `UniversalID`, normalized `UPPER(TRIM(id))` — this reconciles the
  four source spellings (`UniversalID`, `Universal ID`, `Universal Id`, Cornerstone `User ID`).
- `raw` columns are all text (NVARCHAR) — loads never fail on messy data; the views do the typing.
- Epic exports its flags as `✅ Yes` / `⛔ No`, so the training view matches `LIKE '%YES%'`.
- Connecting tools need `-C` / `TrustServerCertificate=yes` (the server uses a self-signed cert).
- `report.master_vs_sources` **only proposes** changes; writing them back to the Master is a
  separate, reviewed step (not built here — by design). The warehouse is read-only reporting.
