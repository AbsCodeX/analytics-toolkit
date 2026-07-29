# Daily Refresh Runbook — RCM Wave Training Data

**This is THE runbook. Since 2026-07-28 the full run is AUTOMATED**: the
`YourOrgDailyRefresh` scheduled task runs `daily_refresh.py auto`
(prep → apply, no pause) weekdays at **9:35**, finishing ~9:50 — before the
10:00 `YourOrgW3Tracker` exec brief. The morning routine is now: open
`data\reports\Main Reports\Morning Review.xlsx` at 10 and read **Start
Here** — it reports the whole run (status, errors, files updated, flags,
rows to review) plus a clickable Directory of everything.

Manual runs still work exactly as before (and are the recovery path):

```powershell
cd C:\path\to\analytics

python scripts\daily_refresh.py prep     # 1) everything that does NOT touch the Master
#    ... open data\reports\Main Reports\Morning Review.xlsx and review ...
python scripts\daily_refresh.py apply    # 2) the Master updates + team file publish
```

(`python scripts\daily_refresh.py all` runs both with a type-APPLY pause between;
`auto` runs both unattended (Morning Review built once, at the END, as the
run report). `--list`, `--from <step>`, `--only <step>`, `--skip <step>`,
`--force` for control.)

---

## The rule that decides where logic lives

> **Python writes files; SQL answers questions.**
>
> - Anything that *changes* the Master or the Team File is a Python writer with
>   its guards (lock checks, daily backup, DATA snapshots, runlogs). Blocking
>   validations live inside those writers.
> - Anything needing *human review or reporting* — differs flags, gaps, trends,
>   history, KPIs — is a SQL view in `report.*`, exported to Excel. Never
>   re-implemented in pandas.
> - New Master column logic → edit the relevant updater script.
>   New report → new view in `sql\3_report_views.sql` + an export step.

---

## Step 0 — drop the fresh exports (manual, filenames must match)

| Export                                           | Drop into (OneDrive`data\raw\...`)         | Cadence   |
| ------------------------------------------------ | -------------------------------------------- | --------- |
| `HR.xlsx` (Workday)                            | `hr\`                                      | Weekly    |
| `User MappingsData.csv` (+ Role/JobCategories) | `mvp\`                                     | Daily     |
| `Epic Team Member Lookup.xlsx` (pre-combined)  | `epic\epic_tm_lookup\`                     | Daily     |
| `Curriculum Status Detail by User - W*.xlsx`   | `epic\epic_status_details\`                | Daily     |
| `epic_class_schedule_YYYYMMDD.xlsx`            | `epic\epic_class_schedules\`               | As needed |
| `Enterprise_Training_Report_*.xlsx`            | `cornerstone\enterprise_training_reports\` | Daily     |
| `Roster_Report_by_Event*.xlsx`                 | `cornerstone\roster_report\`               | PAUSED 2026-07-27 (last resort — see source_registry notes) |
| `Wave Change Request Form*.xlsx`               | `wave\wave_change_requests\`               | As needed |

(2026-07-17: the `Transcript_Status_*` and `Online_Training_Status*` exports
are retired — the Enterprise Training Report now carries full history + Online
Class + Test rows and replaces both. Two Cornerstone pulls a day, not four.)

Overwrite in place (dated/timestamped sources accumulate; sweep older copies to
their `archive\<YYYY-MM>\` subfolder).

## Stage 1 — `prep` (no Master writes; safe to run any time)

**Merged runbook 2026-07-08:** everything that *classifies against the Master*
(Epic TM lookup, gap reports) and every leadership-facing export moved to
`apply`, AFTER the Master writers — so published reports always reflect
today's Master (this was the manual DAILY_SCRIPT_ORDER.md ordering; that doc's
whole-day chain is now retired in favor of this script).

| Step               | What it does                                                                                                                                                  | Gate                                     |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------- |
| `preflight`      | Source freshness, SQL reachable, Master-lock warning                                                                                                          | —                                       |
| `clean_hr`       | Raw HR → dated`HR_cleaned_*.xlsx`                                                                                                                          | only if raw HR newer than latest cleaned |
| `sql_refresh`    | Change-aware (2026-07-10): loads only sources whose files changed since the last load (shared auto_refresh state); always guarantees today's history snapshot | daily                                    |
| `hr_preview`     | Read-only HR leader-change preview CSV                                                                                                                        | daily                                    |
| `morning_review` | Builds**Morning Review.xlsx**                                                                                                                           | daily                                    |

(Since 2026-07-27 EVERYTHING in Morning Review — including the Missing Wave /
Wave vs Sources sheets — reads the SQL views directly, fresh from this
morning's `sql_refresh`. The dated gap workbooks those sheets used to show
are retired, except the weekly HR one.)

## The review — one workbook

Open `data\reports\Main Reports\Morning Review.xlsx`. **Start Here** shows
source freshness and which sheets have rows. Work only the sheets with rows:

| Sheet                                        | What it means                                                  | Your action                                                                      |
| -------------------------------------------- | -------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| HR Leader Changes                            | Leader changes the HR apply step WILL make                     | scan; if wrong, fix HR/aliases before apply                                      |
| Master vs Sources                            | Master differs from HR (AVP/VP) or MVP (job role)              | spot-check; apply will align most                                                |
| Exceptions                                   | Centralized-not-on-wave (must be 0) / in-scope missing from HR | add to wave file / fix the UID                                                   |
| Missing Wave - HR / - Epic / Wave vs Sources | Gap reports between systems                                    | decide who to add / follow up                                                    |
| Mapping Gaps                                 | Master values not in the reference lists (BU/vendor/leader)    | edit`wave_reference_lists.xlsx` + `refresh.py refs`, or fix the Master value |
| Wave Change Requests                         | ALL open request forms in the folder (one per requester)       | disposition each; move a finished form to`archive\<YYYY-MM>\` and it drops off |
| Yesterday Changes                            | What actually changed at the last snapshot                     | sanity-check yesterday's apply                                                   |

## Stage 2 — `apply` (Excel CLOSED; needs desktop Excel)

First safety copies, then the writers, then everything derived from the Master:

`guard` (refuses without a successful prep today)
→ `backups` (2026-07-28: pre-rebuild `.bak.<date>` copies of the Tracker +
Daily Log into `Main Reports\Backups\<month>\`, and guarantees the Master's
daily full backup — the revert points; at most one per file per day)
→ `add_members`
(add_missing_to_master.py --apply: appends people missing from the Master —
Epic "Revenue Cycle - Centralized" + HR Lastname03-org actives — prints the full
add-list + review CSV first; a >150-candidate sanity cap aborts the apply)
→ `mvp_update` → `hr_apply` (xlwings/COM)
→ `gaps_hr` (weekly gate: the shared HR gap workbook)
→ `sql_post` (reload Master + re-stamp today's history — must precede the
report steps, whose views join `raw.master`)
→ `boss_reports`
→ `tracker` (rebuilds `RCM Training Tracker.xlsx` + `RCM Training Daily
Log.xlsx` — built and QA'd in local `build_staging\`, then published
atomically to `data\reports\Main Reports\`; both files must be CLOSED in
Excel or the step aborts cleanly; resume with `--from tracker`. Midday
refresh without the full pipeline: `python
scripts\build_rcm_training_tracker.py`, safe every 2 hours)
→ `metrics` (SQL-based; reads the gap views directly since 2026-07-27)
→ `team_file` (publishes `RCM Wave Team File.xlsx` + dated distribution copy)
→ `review_refresh` (2026-07-28: rebuilds **Morning Review.xlsx** LAST so it
reports the finished run — Today's Run step results, Files Updated flags,
Run History, Activity Log, clickable Directory, post-apply gap sheets. Sheets
named `My ...` are carried forward on every rebuild — personal notes live
there. If the workbook is open in Excel it retries, then saves a side-by-side
"(LATEST ...)" copy instead of failing the pipeline).

(Retired steps: `epic_lookup_update` 2026-07-09 — Epic-mirror columns removed
from the Master, Epic data lives in SQL. **Trimmed 2026-07-27** after the
pipeline audit: `epic_tm_lookup` (dated TM Lookup workbook — nothing consumed
it; SQL + the Master mirror read the raw export), `gaps_epic` + `gaps_wave`
(Morning Review + metrics now read the SQL gap views), `sql_reports`
(no-show/unregistered exports — the tracker covers them), `w3_reports` (daily
W3 package — the tracker + Daily Log cover it). Scripts archived under
`scripts\_archive\`. The 10:00 `YourOrgW3Tracker` task now runs
`w3_exec_brief.py` only — the old 6-tab W3 tracker workbook is retired, trend
questions go to the Daily Log. Apply runtime dropped ~5.5 minutes.)

Safety net (all pre-existing, all still active): lock-check aborts, once-daily
full backup in `data\reports\archive\<month>\`, values-only DATA snapshot after
every edit, header-integrity + empty-load guards, per-script runlogs.

## Where things land

- **Team:** `data\reports\Main Reports\RCM Wave Team File.xlsx` (share this, never the Master)
- **Boss:** `data\reports\gap_reports\` (both "(Latest)" reports + dated copies)
- **Tracker:** `data\reports\Main Reports\` — RCM Training Tracker + Daily Log
  (no-shows, unregistered, audits, trend all live here since 2026-07-27; the
  standalone `no_show\` / `unregistered\` exports are retired)
- **Ad-hoc pulls:** `python sql\query.py "SELECT ..." --excel name.xlsx` → `data\reports\sql_pulls\`
- **Power BI:** connect straight to SQL — see `sql\POWERBI.md`
- **History/audit:** `report.person_events`, `report.daily_changes`,
  `report.registration_daily` in SQL (see `sql\README.md`); run history in
  `data\runlogs\` and `pbi.run_history`

## Speed & robustness notes (2026-07-10)

- `daily_refresh.py` re-runs itself under `analytics_env` automatically — plain
  `python scripts\daily_refresh.py ...` is safe from any terminal.
- Loads are atomic (stage-and-swap): a killed run can no longer leave a
  half-loaded `raw.*` table, and Excel reads use the ~10x faster calamine engine.
- `refresh.py` builds `raw.uids_hr` / `raw.uids_epic_status` key tables at load
  time; `report.exceptions` seeks them (was the ~10-min query in Morning Review,
  now <1s). If those tables are ever missing, run `python sql\refresh.py hr epic_status`.

## Troubleshooting

- **A step fails:** fix the printed error, then
  `python scripts\daily_refresh.py <stage> --from <step>`. Full output is in
  `data\runlogs\daily_refresh\<YYYY-MM>\`.
- **Master open in Excel:** prep continues (SQL uses the newest DATA snapshot);
  apply refuses — close Excel and rerun.
- **HR landed late:** rerun `python scripts\daily_refresh.py prep --from clean_hr`.
- **"another run appears active":** a crashed run left
  `data\runlogs\daily_refresh\.running` — delete it if nothing is running.
- **Skip the HR apply on a stale-HR day:** `apply --skip hr_apply`.

## Scheduling

**Decision update 2026-07-28 (the analyst's call — supersedes the 2026-07-06
"apply stays manual" rule):** the FULL run is now scheduled. Weekday morning
order:

| Time | Task | What |
| --- | --- | --- |
| 8:30 | `YourOrgSQLAutoRefresh` | change-aware SQL load (early exports) |
| 9:35 | `YourOrgDailyRefresh` | `daily_refresh.py auto` — full prep+apply, ~15 min (after the MVP export lands ~9:00–9:30) |
| 10:00 | `YourOrgW3Tracker` | `w3_exec_brief.py` (reads the freshly loaded SQL) |
| 11:00 / 13:00 / 14:30 | `YourOrgSQLAutoRefresh` | midday change-aware SQL loads (catch late exports) |

`YourOrgDailyRefresh` runs `analytics_env\Scripts\pythonw.exe
scripts\daily_refresh.py auto` as "Run only when user is logged on" — the
xlwings/COM steps need the interactive session, so **the laptop must be on
and logged in** (locked is fine). A missed start (laptop off) runs when it
next can (`StartWhenAvailable`); or just run the two manual commands.
Revert points: every file the run overwrites has a same-day `.bak` /dated
copy — see the Morning Review **Directory** sheet's "Backups / revert" rows.

`YourOrgSQLAutoRefresh` (`sql\auto_refresh.py`) is the change-aware wrapper:
reloads only the `raw.*` sources whose files actually changed, re-applies
`3_report_views.sql` / `4_dimensions.sql` when those files are edited, and
re-stamps today's history snapshot when data changed. No-change cycles are
silent no-ops (log: `data\runlogs\daily_refresh\auto_refresh_<YYYY-MM>.log`;
state: `auto_refresh_state.json` in the same folder). It skips a cycle while
`daily_refresh.py` holds the `.running` lock, so the 9:35 full run and the
SQL task never collide. Manual `python sql\refresh.py` runs remain safe at
any time; the auto task simply notices there is then nothing new to load.

## History note

`report.person_events` (SQL) is the change history going forward (from
2026-07-06). The in-cell `Users Wave/HR Change Log` columns keep dual-running
until ~2026-08-03, then their write-blocks can be removed from
`update_master_from_mvp.py` / `update_master_from_hr_safe.py`.
