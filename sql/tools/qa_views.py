# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""QA suite for the session-built views: reconciliation + synthetic history test."""
import sys
sys.path.insert(0, r"C:\path\to\analytics\sql")
from refresh import CONN
import pandas as pd
from sqlalchemy import create_engine, text

eng = create_engine(CONN)
q = lambda sql: pd.read_sql(sql, eng)
one = lambda sql: q(sql).iloc[0, 0]

results = []

def check(name, ok, detail=""):
    results.append((name, ok, detail))
    print(f"{'PASS' if ok else 'FAIL'}  {name}  {detail}")

print("========== PART 1: reconciliation checks ==========")

master_rows = one("SELECT COUNT(*) FROM raw.[master] WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> ''")
users_rows = one("SELECT COUNT(*) FROM report.users")
check("users == master non-blank-UID rows", users_rows == master_rows, f"{users_rows} vs {master_rows}")

dup_uids = one("""SELECT COUNT(*) FROM (SELECT UPPER(LTRIM(RTRIM(UniversalID))) u FROM raw.[master]
 WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> '' GROUP BY UPPER(LTRIM(RTRIM(UniversalID))) HAVING COUNT(*)>1) x""")
print(f"INFO  duplicate UIDs in Master: {dup_uids}")

roster_rows = one("SELECT COUNT(*) FROM report.roster")
check("roster == users (no join fan-out)", roster_rows == users_rows, f"{roster_rows} vs {users_rows}")

pbi_roster = one("SELECT COUNT(*) FROM pbi.roster")
check("pbi.roster alias parity", pbi_roster == roster_rows, f"{pbi_roster}")

# Both reports are scoped to wave-file people under canonical leaders.
# Since 2026-07-17 raw.cornerstone carries full history, so report.noshow is
# STANDING no-shows only: newest live row per person x class still 'No Show'.
raw_noshow = one("""WITH latest AS (
  SELECT UPPER(LTRIM(RTRIM(c.[User_ID]))) AS uid, c.Transcript_Status,
         ROW_NUMBER() OVER (
             PARTITION BY UPPER(LTRIM(RTRIM(c.[User_ID]))), UPPER(LTRIM(RTRIM(c.Training_Title)))
             ORDER BY TRY_CONVERT(datetime, c.Transcript_Registration_Date) DESC) AS rn
  FROM raw.cornerstone c
  WHERE UPPER(LTRIM(RTRIM(ISNULL(c.Training_Provider,'')))) = 'EPIC'
    AND c.Training_Type = 'Session'
    AND c.Transcript_Status NOT IN ('Withdrawn','Cancelled','Denied','Waitlist Expired')
    AND UPPER(c.Training_Title) NOT LIKE '%W2%'
    AND UPPER(c.Training_Title) NOT LIKE '%WAVE 2%')
 SELECT COUNT(*) FROM latest l
 JOIN report.roster r ON r.UniversalID = l.uid
 WHERE l.rn = 1 AND UPPER(LTRIM(RTRIM(l.Transcript_Status))) = 'NO SHOW'
   AND r.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)""")
v_noshow = one("SELECT COUNT(*) FROM report.noshow")
check("noshow == leader-scoped STANDING No Shows (no fan-out)", v_noshow == raw_noshow, f"{v_noshow} vs {raw_noshow}")

# Since 2026-07-20 the view dedupes to UNIQUE event classes per person (the
# same class repeats across curricula/sequences in the export) — compare
# against distinct person x class, not raw 'No' rows.
raw_unreg = one("""SELECT COUNT(*) FROM (
 SELECT DISTINCT UPPER(LTRIM(RTRIM(s.[Universal_Id]))) AS uid,
        UPPER(LTRIM(RTRIM(CONVERT(nvarchar(200), s.Event_Class)))) AS ec
 FROM raw.epic_status s
 JOIN report.roster r ON r.UniversalID = UPPER(LTRIM(RTRIM(s.[Universal_Id])))
 WHERE UPPER(LTRIM(RTRIM(ISNULL(s.Event_Class_Registered,'')))) = 'NO'
   AND r.Leader IN (SELECT LTRIM(RTRIM(Leader)) FROM raw.ref_leaders)) x""")
v_unreg = one("SELECT COUNT(*) FROM report.unregistered")
check("unregistered == leader-scoped distinct person x class 'No' rows", v_unreg == raw_unreg, f"{v_unreg} vs {raw_unreg}")

viol = one("""SELECT COUNT(*) FROM report.epic_not_in_hr e
 JOIN (SELECT DISTINCT UPPER(LTRIM(RTRIM(UniversalID))) u FROM raw.hr
       WHERE LTRIM(RTRIM(ISNULL(UniversalID,''))) <> '') h ON h.u = e.UniversalID""")
check("epic_not_in_hr: zero rows actually in HR", viol == 0, f"{viol} violations")

indep = one("""SELECT COUNT(*) FROM
 (SELECT DISTINCT UPPER(LTRIM(RTRIM([Universal_ID]))) u FROM raw.epic_lookup
  WHERE LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> '') e
 WHERE NOT EXISTS (SELECT 1 FROM raw.hr h WHERE UPPER(LTRIM(RTRIM(h.UniversalID))) = e.u
                   AND LTRIM(RTRIM(ISNULL(h.UniversalID,''))) <> '')""")
v_gap = one("SELECT COUNT(*) FROM report.epic_not_in_hr")
check("epic_not_in_hr count == independent anti-join", v_gap == indep, f"{v_gap} vs {indep}")

kpi = q("SELECT * FROM report.kpi_summary").iloc[0]
today_trend = q("""SELECT SUM(InScope) i, SUM(Registered) r, SUM(Trained) t
 FROM report.registration_daily WHERE SnapshotDate = (SELECT MAX(SnapshotDate) FROM history.roster_daily)""").iloc[0]
check("registration_daily(today) InScope == kpi_summary", int(today_trend.i) == int(kpi.InScopeUsers),
      f"{int(today_trend.i)} vs {int(kpi.InScopeUsers)}")
check("registration_daily(today) Registered == kpi_summary", int(today_trend.r) == int(kpi.RegisteredUsers),
      f"{int(today_trend.r)} vs {int(kpi.RegisteredUsers)}")

# Re-baselined 2026-07-08 after build_reference_lists.py rebuilt the ref
# workbook: every Master BU is now seeded into ref_business_units, so the
# "BU not in ref" gap category must stay at zero by construction.
bu_gap = one("""SELECT COUNT(*) FROM report.mapping_gaps
 WHERE Issue LIKE 'Master BU not in ref%'""")
check("mapping_gaps: zero 'Master BU not in ref' rows", bu_gap == 0, f"{bu_gap}")
gaps = one("SELECT COUNT(*) FROM report.mapping_gaps")
print(f"INFO  mapping_gaps rows: {gaps} (data-dependent — review report.mapping_gaps when > 0)")
bu_need = one("SELECT COUNT(*) FROM dim.business_unit WHERE NeedsMapping = 1")
check("dim.business_unit NeedsMapping == 0 (all Master BUs seeded into ref)",
      bu_need == 0, f"{bu_need}")
ldr_rev = one("SELECT COUNT(*) FROM dim.leader WHERE NeedsReview = 1")
check("dim.leader NeedsReview == 0 (JLIM9 typo fixed by 2026-07-21)", ldr_rev == 0, f"{ldr_rev}")

vend_sum = one("SELECT SUM(MasterPeople) FROM dim.vendor")
master_vend = one("""SELECT COUNT(*) FROM raw.[master] m JOIN dim.vendor_alias va
 ON va.Alias = UPPER(LTRIM(RTRIM(m.BusinessUnit)))
 WHERE UPPER(LTRIM(RTRIM(m.[Vendor_Yes_No]))) = 'YES'""")
check("dim.vendor headcount == master vendor rows w/ alias", int(vend_sum) == int(master_vend),
      f"{int(vend_sum)} vs {int(master_vend)}")

sb = one("SELECT COUNT(*) FROM sandbox.example_not_registered")
sb_indep = one("""SELECT COUNT(*) FROM report.roster
 WHERE IsInScope = 1 AND ISNULL(IsFullyRegistered, 0) = 0""")
check("sandbox example == roster in-scope not-registered", sb == sb_indep, f"{sb} vs {sb_indep}")

print("\n========== PART 1b: Master QA layer (added 2026-07-08) ==========")

rec_rows = one("SELECT COUNT(*) FROM report.user_type_recommendations")
check("user_type_recommendations == users (no join fan-out)", rec_rows == users_rows,
      f"{rec_rows} vs {users_rows}")

bad_domain = one("""SELECT COUNT(*) FROM report.user_type_recommendations
 WHERE UserType_Recommended IS NOT NULL
   AND UserType_Recommended NOT IN ('YourOrg','Offshore','Onshore')""")
check("recommendations: values within domain", bad_domain == 0, f"{bad_domain}")

diff_null = one("""SELECT COUNT(*) FROM report.user_type_recommendations
 WHERE Differs = 1 AND UserType_Recommended IS NULL""")
check("recommendations: no Differs row without a recommendation", diff_null == 0, f"{diff_null}")

high_null = one("""SELECT COUNT(*) FROM report.user_type_recommendations
 WHERE Confidence = 'High' AND UserType_Recommended IS NULL""")
check("recommendations: High confidence always has a value", high_null == 0, f"{high_null}")

nw_wrong = one("""SELECT COUNT(*) FROM report.user_type_recommendations
 WHERE UPPER(ISNULL(WorkerType,'')) LIKE '%YOURORG%'
   AND UserType_Recommended IN ('Offshore','Onshore')""")
check("recommendations: YourOrg worker never recommended a vendor type", nw_wrong == 0, f"{nw_wrong}")

# Mirrors the 2026-07-21 view logic: EpicYes_MasterNo excludes departed,
# no-training job roles, and Epic HR status Terminated / Pay Leave. (The view
# picks one arbitrary epic_lookup row per UID; MAX() here can differ by an
# edge case if a UID has conflicting rows — none observed to date.)
tn_indep = one("""SELECT COUNT(*) FROM raw.[master] m
 JOIN (SELECT UPPER(LTRIM(RTRIM([Universal_ID]))) u,
              MAX(LTRIM(RTRIM([Epic_Training_Needed_MVP]))) e,
              MAX(LTRIM(RTRIM([Hr_Status_PDM]))) hs
       FROM raw.epic_lookup WHERE LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> ''
       GROUP BY UPPER(LTRIM(RTRIM([Universal_ID])))) ep
   ON ep.u = UPPER(LTRIM(RTRIM(m.UniversalID)))
 WHERE UPPER(LTRIM(RTRIM(m.[Training_Needed_Yes_No]))) = 'NO'
   AND UPPER(ISNULL(ep.e,'')) LIKE '%YES%'
   AND UPPER(ISNULL(ep.hs,'')) NOT IN ('TERMINATED','PAY LEAVE')
   AND UPPER(LTRIM(RTRIM(ISNULL(m.[IsDepartedInactive],'')))) NOT IN ('YES','Y','TRUE','1')
   AND NOT EXISTS (SELECT 1 FROM raw.ref_no_training_job_roles nt
                   WHERE UPPER(LTRIM(RTRIM(nt.JobRole))) IN
                         (UPPER(LTRIM(RTRIM(ISNULL(m.[Job_Role_1],'')))),
                          UPPER(LTRIM(RTRIM(ISNULL(m.[Job_Role_2],'')))),
                          UPPER(LTRIM(RTRIM(ISNULL(m.[Job_Role_3],'')))),
                          UPPER(LTRIM(RTRIM(ISNULL(m.[Job_Role_4],''))))))""")
tn_view = one("SELECT ISNULL(SUM(EpicYes_MasterNo),0) FROM report.training_needed_check")
check("training_needed_check: EpicYes_MasterNo == independent join", tn_view == tn_indep,
      f"{tn_view} vs {tn_indep}")

epic_gap_indep = one("""SELECT COUNT(*) FROM
 (SELECT DISTINCT UPPER(LTRIM(RTRIM([Universal_ID]))) u FROM raw.epic_lookup
  WHERE LTRIM(RTRIM(ISNULL([Curriculum_Type],''))) = 'Revenue Cycle - Centralized'
    AND LTRIM(RTRIM(ISNULL([Universal_ID],''))) <> '') e
 WHERE NOT EXISTS (SELECT 1 FROM raw.[master] m
                   WHERE UPPER(LTRIM(RTRIM(m.UniversalID))) = e.u)""")
epic_gap_view = one("SELECT COUNT(DISTINCT UniversalID) FROM report.epic_missing_from_wave")
check("epic_missing_from_wave == independent anti-join", epic_gap_view == epic_gap_indep,
      f"{epic_gap_view} vs {epic_gap_indep}")

# Ref hygiene — the generator writes these clean; hand edits must not regress them.
sys.path.insert(0, r"C:\path\to\analytics\scripts")
from leader_names import CANONICAL_NAMES  # noqa: E402
ref_ldr = one("SELECT COUNT(*) FROM raw.ref_leaders WHERE LTRIM(RTRIM(ISNULL(Leader,''))) <> ''")
check("ref_leaders == leader_names.CANONICAL_NAMES (single source in sync)",
      ref_ldr == len(CANONICAL_NAMES), f"{ref_ldr} vs {len(CANONICAL_NAMES)}")

blank_keys = one("""SELECT
   (SELECT COUNT(*) FROM raw.ref_business_units WHERE LTRIM(RTRIM(ISNULL(BusinessUnit,'')))='')
 + (SELECT COUNT(*) FROM raw.ref_leaders WHERE LTRIM(RTRIM(ISNULL(Leader,'')))='')
 + (SELECT COUNT(*) FROM raw.ref_no_training_job_roles WHERE LTRIM(RTRIM(ISNULL(JobRole,'')))='')
 + (SELECT COUNT(*) FROM raw.ref_vendor_aliases WHERE LTRIM(RTRIM(ISNULL(Alias,'')))='')""")
check("ref hygiene: zero blank keys across ref lists", blank_keys == 0, f"{blank_keys}")

dup_alias = one("""SELECT COUNT(*) FROM (SELECT UPPER(LTRIM(RTRIM(Alias))) a
 FROM raw.ref_vendor_aliases WHERE LTRIM(RTRIM(ISNULL(Alias,''))) <> ''
 GROUP BY UPPER(LTRIM(RTRIM(Alias))) HAVING COUNT(*) > 1) x""")
check("ref hygiene: zero duplicate vendor aliases", dup_alias == 0, f"{dup_alias}")

print("\n========== PART 2: synthetic history test ==========")
# Re-baselined 2026-07-08: history.roster_daily now accumulates REAL multi-day
# snapshots, so the test is baseline-relative — it clones the EARLIEST real
# day, inserts a mutated copy one day before it, and asserts only on that
# first transition (EventDate = earliest real day). Post-cleanup compares
# against captured pre-test counts, not zero.
earliest = str(one("SELECT MIN(SnapshotDate) FROM history.roster_daily"))[:10]
FAKE_DAY = str((pd.to_datetime(earliest) - pd.Timedelta(days=1)).date())
with eng.begin() as con:
    con.execute(text(f"DELETE FROM history.roster_daily WHERE SnapshotDate = '{FAKE_DAY}'"))

events_before = one("SELECT COUNT(*) FROM report.person_events")
days_before = one("SELECT COUNT(DISTINCT SnapshotDate) FROM history.roster_daily")
print(f"baseline: earliest real day {earliest}, fake day {FAKE_DAY}, "
      f"{events_before} real events, {days_before} real days")

today = q(f"SELECT * FROM history.roster_daily WHERE SnapshotDate = '{earliest}'")
print(f"earliest snapshot rows: {len(today)}")
y = today.copy()
y["SnapshotDate"] = FAKE_DAY

def pick(mask, taken):
    cand = y[mask & ~y["UniversalID"].isin(taken)]
    return cand.iloc[0]["UniversalID"] if len(cand) else None

taken = set()
mut = {}
mut["wave"] = pick(y["Wave"].astype(str).str.strip().eq("Wave 3"), taken); taken.add(mut["wave"])
y.loc[y["UniversalID"] == mut["wave"], "Wave"] = "Wave 2"

mut["leader"] = pick(y["Leader"].astype(str).ne("Lastname08, Firstname08") & y["Leader"].notna(), taken); taken.add(mut["leader"])
y.loc[y["UniversalID"] == mut["leader"], "Leader"] = "Lastname08, Firstname08"

reg_mask = pd.to_numeric(y["IsFullyRegistered"], errors="coerce").fillna(0) == 1
mut["reg"] = pick(reg_mask, taken); taken.add(mut["reg"])
y.loc[y["UniversalID"] == mut["reg"], "IsFullyRegistered"] = 0

mut["role"] = pick(y["JobRole1"].notna() & y["JobRole1"].astype(str).str.strip().ne(""), taken); taken.add(mut["role"])
y.loc[y["UniversalID"] == mut["role"], "JobRole1"] = "QA OLD ROLE"

mut["dep"] = pick(y["Departed"].astype(str).str.upper().eq("FALSE"), taken); taken.add(mut["dep"])
y.loc[y["UniversalID"] == mut["dep"], "Departed"] = "True"

loa_mask = y["AssignmentStatus"].astype(str).str.contains("Leave", na=False)
mut["loa"] = pick(loa_mask, taken); taken.add(mut["loa"])
y.loc[y["UniversalID"] == mut["loa"], "AssignmentStatus"] = "Active - Payroll Eligible"

mut["added"] = pick(y["Wave"].notna(), taken); taken.add(mut["added"])
y = y[y["UniversalID"] != mut["added"]]  # absent yesterday -> "Added to Master" today

ghost = today.iloc[[0]].copy()
ghost["SnapshotDate"] = FAKE_DAY
ghost["UniversalID"] = "ZZQATEST1"
ghost["FullName"] = "QA, Ghost"
y = pd.concat([y, ghost], ignore_index=True)

print("mutations:", {k: v for k, v in mut.items()})
y.to_sql("roster_daily", eng, schema="history", if_exists="append", index=False)

try:
    # Change events are dated on the LATER day of the transition; 'Dropped off
    # Master' is dated on the person's LAST-SEEN day (the fake day) — so the
    # synthetic transition spans both dates.
    ev = q(f"""SELECT EventDate, UniversalID, EventType, FromValue, ToValue
 FROM report.person_events WHERE EventDate IN ('{FAKE_DAY}', '{earliest}')
 ORDER BY EventType""")
    print(ev.to_string(index=False))
    check("person_events(first transition): exactly 7 events", len(ev) == 7, f"{len(ev)}")
    expect = {
        (mut["wave"], "Wave changed"), (mut["leader"], "Leader changed"),
        (mut["role"], "Job role changed"), (mut["dep"], "Departed/Inactive changed"),
        (mut["loa"], "HR status changed (Active/LOA)"), (mut["added"], "Added to Master"),
        ("ZZQATEST1", "Dropped off Master"),
    }
    got = set(zip(ev.UniversalID, ev.EventType))
    check("person_events(first transition): exact expected set", got == expect,
          f"missing={expect - got} extra={got - expect}")

    dc = q(f"""SELECT UniversalID, BecameRegistered, WaveChangedFrom, LeaderChangedFrom
 FROM report.daily_changes WHERE SnapshotDate = '{earliest}'""")
    check("daily_changes(first transition): exactly 3 rows (wave, leader, registration)",
          len(dc) == 3, f"{len(dc)}")
    check("daily_changes: BecameRegistered flagged for the right person",
          bool((dc[dc.UniversalID == mut["reg"]].BecameRegistered == 1).all()) and len(dc[dc.UniversalID == mut["reg"]]) == 1, "")

    days = one("SELECT COUNT(DISTINCT SnapshotDate) FROM report.registration_daily")
    check("registration_daily: fake day adds one snapshot day", days == days_before + 1,
          f"{days} vs {days_before}+1")
finally:
    with eng.begin() as con:
        n = con.execute(text(f"DELETE FROM history.roster_daily WHERE SnapshotDate = '{FAKE_DAY}'")).rowcount
    print(f"cleanup: deleted {n} fake-day rows")

post = one("SELECT COUNT(*) FROM report.person_events")
check("post-cleanup: person_events back to baseline", post == events_before,
      f"{post} vs {events_before}")
post_days = one("SELECT COUNT(DISTINCT SnapshotDate) FROM history.roster_daily")
check("post-cleanup: history day count back to baseline", post_days == days_before,
      f"{post_days} vs {days_before}")

print("\n========== SUMMARY ==========")
fails = [r for r in results if not r[1]]
print(f"{len(results) - len(fails)} passed, {len(fails)} failed")
for name, _, detail in fails:
    print(f"  FAIL: {name}  {detail}")
eng.dispose()
