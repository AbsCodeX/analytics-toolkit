# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_soft_live_list.py  —  Wave 3 soft live participant list.

Population (per Dr. Lastname03's 08/26 ask):
  * every Wave 3 Rev Cycle person with a training indicator of Yes, and
  * every SVC (Simple Visit Coding) person we have Epic status data on to date.

SVC is derived, not read from a hand-maintained export: a person counts as SVC when
they hold a "Simple Visit Coding" job role in MVP User Mappings AND appear in the Epic
"Curriculum Status by User Summary" export (which is what "we have data on" means).
The old ad_hoc "Curriculum Status by User - SVC.xlsx" is stale and deliberately unused.

    python scripts\\build_soft_live_list.py
    python scripts\\build_soft_live_list.py --wave "Wave 4"
    python scripts\\build_soft_live_list.py --all-waves-svc   # SVC regardless of wave
    python scripts\\build_soft_live_list.py --show-config

Leadership (Senior Manager -> Leaders) is filled from the HR file for the whole
population, per the analyst 08/26. Email comes from MVP UserUPN, which covers everyone;
HR and the Epic summary are fallbacks.

Sheets produced: Summary · Soft Live List · SVC Job Roles
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "sql"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from refresh import CONN, op                                   # noqa: E402
from sqlalchemy import create_engine                           # noqa: E402
import pandas as pd                                            # noqa: E402
import openpyxl                                                # noqa: E402
from openpyxl.styles import Alignment                          # noqa: E402
from openpyxl.utils import get_column_letter                   # noqa: E402

from build_lava_list import styles, SVC_ROLE_PATTERN           # noqa: E402

# ---------------------------------------------------------------- config ----
WAVE = "Wave 3"
GO_LIVE = "11/07/2026"
SUMMARY_EXPORT = (op.ONEDRIVE_ROOT / "data" / "raw" / "epic" / "epic_status_summary"
                  / "Curriculum Status by User Summary.xlsx")
MVP_CSV = op.ONEDRIVE_ROOT / "data" / "raw" / "mvp" / "User MappingsData.csv"
OUT_NAME = "{wave} Soft Live Participants {date}.xlsx"
HEADER_ROW = 5
JR = ["JobRole1", "JobRole2", "JobRole3", "JobRole4"]

# (Header, internal field, width)
COLUMNS = [
    ("UniversalID", "UniversalID", 14), ("Full Name", "Full Name", 24),
    ("First Name", "First Name", 14), ("Last Name", "Last Name", 16),
    ("Email", "Email", 30),
    ("Population", "Population", 22),
    ("SVC Staff (Yes/No)", "SVC YN", 12), ("SVC Job Role", "SVC Job Role", 42),
    ("Wave", "Wave", 9), ("Training Needed", "TrainingNeeded", 13),
    ("User Type", "User Type", 12), ("Vendor (Yes/No)", "Vendor YN", 11),
    ("Leaders", "Leaders", 22), ("Senior Manager", "SeniorManager", 22),
    ("Director", "Director", 22), ("Sr. Director", "SeniorDirector", 22),
    ("AVP", "AVP", 22), ("VP", "VP", 22), ("SVP", "SVP", 22),
    ("Job Role 1", "JobRole1", 34), ("Job Role 2", "JobRole2", 30),
    ("Job Role 3", "JobRole3", 26), ("Job Role 4", "JobRole4", 22),
    ("Job Title (HR)", "JobTitle", 30), ("Department (HR)", "Department", 26),
    ("Business Unit (MVP)", "BusinessUnitDescription", 24),
    ("# Required Curricula", "# Required Cirrculums", 13),
    ("# Registered", "# Registered Curriculums", 12),
    ("# Completed", "# Completed Curriculums", 12),
    ("Fully Registered", "Fully Registered", 14),
    ("Fully Trained", "Fully Trained", 14),
    ("Training Status", "Training Status", 22),
    ("Est. Training Completion Date", "Est Completion", 24),
]


# ------------------------------------------------------------------ data ----
def load(args):
    """Assemble the person-level frame from SQL plus the two source files."""
    eng = create_engine(CONN)

    users = pd.read_sql("SELECT * FROM report.users", eng)
    users["_u"] = users["UniversalID"].str.strip().str.upper()
    wave_ids = set(users.loc[(users["Wave"] == args.wave)
                             & (users["TrainingNeeded"] == "Yes")
                             & (users["IsInScope"] == 1)
                             & (users["OnLeaderList"] == 1), "_u"])

    # --- SVC: an MVP Simple Visit Coding role holder we have Epic status data for ---
    summary = pd.read_excel(args.summary_export, sheet_name=0, dtype=str)
    uid_col = next(c for c in summary.columns
                   if c.strip().lower().replace(" ", "") == "universalid")
    summary["_u"] = summary[uid_col].astype(str).str.strip().str.upper()
    summary = summary[summary["_u"].ne("") & summary["_u"].ne("NAN")]
    summary = summary.drop_duplicates("_u").set_index("_u")

    mvp = pd.read_csv(args.mvp_csv, dtype=str, low_memory=False, usecols=[
        "UniversalID", "UserUPN", "GoLiveWave", "BusinessUnitDescription",
        "IsDepartedInactive",
        "IndividualCategoryUpdate1Name", "IndividualCategoryUpdate2Name",
        "IndividualCategoryUpdate3Name", "IndividualCategoryUpdate4Name"])
    mvp = mvp.rename(columns={f"IndividualCategoryUpdate{i}Name": f"JobRole{i}"
                              for i in range(1, 5)})
    mvp["_u"] = mvp["UniversalID"].astype(str).str.strip().str.upper()
    mvp = mvp.drop_duplicates("_u").set_index("_u").drop(columns=["UniversalID"])

    is_svc_role = pd.Series(False, index=mvp.index)
    for c in JR:
        is_svc_role |= mvp[c].fillna("").str.contains(SVC_ROLE_PATTERN, case=False)
    svc_all = set(mvp.index[is_svc_role])
    svc_ids = svc_all & set(summary.index)          # "received data on to date"
    if not args.all_waves_svc:
        svc_ids &= set(mvp.index[mvp["GoLiveWave"].fillna("").str.strip() == args.wave])

    print(f"{args.wave} training-needed (in scope, on leader list): {len(wave_ids)}")
    print(f"SVC role holders in MVP: {len(svc_all)} | with Epic status data: {len(svc_ids)}"
          f" | overlap with wave group: {len(wave_ids & svc_ids)}")

    hr = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(UniversalID))) _u, FullName HRName,
        Leader HRLeader, SeniorManager, Director, SeniorDirector, AVP, VP, SVP,
        Department, JobTitle, Email HREmail, WorkerType
        FROM report.hr""", eng).drop_duplicates("_u").set_index("_u")
    mstr = pd.read_sql("SELECT UPPER(LTRIM(RTRIM(UniversalID))) _u, FirstName, LastName"
                       " FROM raw.master", eng).drop_duplicates("_u").set_index("_u")
    track = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(UniversalID))) _u, TrainingStatus,
        FinalScheduledDate, FullyRegisteredYN, FullyTrainedYN
        FROM report.tracker_training_status""", eng).drop_duplicates("_u").set_index("_u")
    corn = pd.read_sql("""SELECT UPPER(LTRIM(RTRIM(User_ID))) _u,
        MAX(Training_Start_Date) LastSession,
        MAX(CASE WHEN Transcript_Status LIKE 'Completed%' AND Training_Start_Date IS NOT NULL
                 THEN Training_Start_Date END) LastCompletedClass,
        COALESCE(
          MAX(CASE WHEN Transcript_Status LIKE '%Equivalent%' THEN Transcript_Completed_Date END),
          MAX(CASE WHEN Transcript_Status LIKE 'Completed%' THEN Transcript_Completed_Date END)
        ) EquivCompleted
        FROM raw.cornerstone WHERE Training_Provider='EPIC'
        GROUP BY UPPER(LTRIM(RTRIM(User_ID)))""", eng).drop_duplicates("_u").set_index("_u")

    rows = sorted(wave_ids | svc_ids)
    keep = ["UniversalID", "FullName", "Wave", "Leader", "UserType", "VendorYN",
            "TrainingNeeded"]
    have = [c for c in keep if c in users.columns]
    summary_cols = [c for c in ["# Required Cirrculums", "# Registered Curriculums",
                                "# Completed Curriculums", "Fully Registered?",
                                "Fully Trained?", "Email", "User Type"]
                    if c in summary.columns]
    d = (pd.DataFrame(index=pd.Index(rows, name="_u"))
         .join(users.set_index("_u")[have]).join(mstr).join(mvp).join(hr)
         .join(track).join(corn)
         .join(summary[summary_cols].add_prefix("sum_")))
    return d, wave_ids, svc_ids, mvp, summary


def derive(d, wave_ids, svc_ids, args):
    """Everything computed rather than looked up."""
    d["UniversalID"] = d["UniversalID"].fillna(pd.Series(d.index, index=d.index))
    d["Full Name"] = d["FullName"].fillna(d["HRName"])
    split = d["Full Name"].fillna("").str.split(",", n=1)
    d["First Name"] = d["FirstName"].fillna(split.str[-1].str.strip())
    d["Last Name"] = d["LastName"].fillna(split.str[0].str.strip())

    d["Population"] = [f"{args.wave} Training + SVC" if (u in wave_ids and u in svc_ids)
                       else (f"{args.wave} Training" if u in wave_ids else "SVC")
                       for u in d.index]
    d["SVC YN"] = ["Yes" if u in svc_ids else "No" for u in d.index]
    for c in JR:
        d[c] = d[c].fillna("").astype(str).str.strip()
    d["SVC Job Role"] = d[JR].apply(
        lambda r: "; ".join(sorted({v for v in r if v and
                                    SVC_ROLE_PATTERN in v.lower()})), axis=1)

    d["Wave"] = d["Wave"].fillna(d["GoLiveWave"])
    # Leadership chain comes from HR for the whole population (her call 08/26).
    d["Leaders"] = d["Leader"].fillna(d["HRLeader"])
    d["User Type"] = d["UserType"].fillna(d.get("sum_User Type"))
    d["Vendor YN"] = d["VendorYN"].fillna(
        d["WorkerType"].map({"VEN": "Yes"})).fillna("No")

    # MVP UserUPN FIRST, then HR, then the Epic summary — the shared order in
    # email_lookup.py. HR is only a leader-chain source, never a reporting one
    # (her rule, 2026-09-02), and its address is a contact field that hands out
    # vendor domains (flexstaff.org, vendor.example); UserUPN is the YourOrg
    # account, which is what Epic/LAVA provisioning needs.
    d["Email"] = (d["UserUPN"].replace("", pd.NA)
                  .fillna(d["HREmail"].replace("", pd.NA))
                  .fillna(d.get("sum_Email")))

    for src, dst in (("sum_# Required Cirrculums", "# Required Cirrculums"),
                     ("sum_# Registered Curriculums", "# Registered Curriculums"),
                     ("sum_# Completed Curriculums", "# Completed Curriculums")):
        d[dst] = d.get(src)
    d["Fully Registered"] = d["FullyRegisteredYN"].fillna(d.get("sum_Fully Registered?"))
    d["Fully Trained"] = d["FullyTrainedYN"].fillna(d.get("sum_Fully Trained?"))
    # tracker_training_status is scoped to the leader list, so SVC-only people have no row.
    # Fall back to the Epic summary's own Fully Trained? flag rather than leaving a third of
    # the file blank.
    d["Training Status"] = d["TrainingStatus"].fillna(
        d["Fully Trained"].map({"Yes": "Fully Trained", "No": "In Progress / Not Complete"}))

    for c in ("FinalScheduledDate", "LastCompletedClass", "EquivCompleted", "LastSession"):
        d[c] = pd.to_datetime(d[c], errors="coerce")
    trained = d["Fully Trained"].astype(str).str.strip().eq("Yes")
    # Same rule as the LAVA list: always a real date. Finished -> the last class they sat;
    # equivalency credit -> the date it was recorded; still training -> next scheduled session.
    d["Est Completion"] = [f if pd.notna(f) else
                           (lc if (t and pd.notna(lc)) else
                            (eq if pd.notna(eq) else (s if pd.notna(s) else None)))
                           for f, lc, t, eq, s in zip(
                               d["FinalScheduledDate"], d["LastCompletedClass"], trained,
                               d["EquivCompleted"], d["LastSession"])]

    return d.sort_values(["Population", "Leaders", "Last Name", "First Name"],
                         key=lambda s: s.astype(str).str.upper())


# ----------------------------------------------------------------- write ----
def write_list_sheet(wb, d, st, args):
    ws = wb.create_sheet("Soft Live List")
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{args.wave} Soft Live Participants"
    ws["A1"].font = st["title"]
    ws["A2"] = (f"All {args.wave} Rev Cycle staff with a training indicator of Yes, plus all SVC "
                f"(Simple Visit Coding) staff we have Epic status data on · {len(d):,} people · "
                f"go-live {GO_LIVE} · data as of {args.stamp:%m/%d/%Y %I:%M %p}")
    ws["A2"].font = st["sub"]
    ws["A3"] = ("SVC = holds a Simple Visit Coding job role in MVP and appears in the Epic "
                "Curriculum Status by User Summary export. Senior Manager through Leaders "
                "come from the HR file. Email is MVP UserUPN.")
    ws["A3"].font = st["sub"]

    for i, (label, _f, width) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=HEADER_ROW, column=i, value=label)
        c.font, c.border = st["head"], st["rule"]
        c.alignment = Alignment(vertical="bottom", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[HEADER_ROW].height = 30

    centred = {"SVC YN", "Vendor YN", "TrainingNeeded", "Fully Registered", "Fully Trained",
               "# Required Cirrculums", "# Registered Curriculums", "# Completed Curriculums"}
    for r, (_, rec) in enumerate(d.iterrows(), start=HEADER_ROW + 1):
        for i, (_label, field, _w) in enumerate(COLUMNS, start=1):
            v = rec.get(field)
            c = ws.cell(row=r, column=i)
            c.font = st["key"] if i == 1 else st["body"]
            c.border = st["rule"]
            if field == "Est Completion" and isinstance(v, pd.Timestamp):
                c.value = v.to_pydatetime()
                c.number_format = "m/d/yyyy"
            else:
                c.value = "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v)
            if field in centred:
                c.alignment = Alignment(horizontal="center")

    last = HEADER_ROW + len(d)
    ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(len(COLUMNS))}{last}"
    ws.freeze_panes = f"C{HEADER_ROW + 1}"


def write_svc_roles_sheet(wb, d, st):
    ws = wb.create_sheet("SVC Job Roles")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "SVC job roles in this population"
    ws["A1"].font = st["title"]
    ws["A2"] = ("Every distinct Simple Visit Coding job role held, from MVP User Mappings. "
                "A person can hold more than one.")
    ws["A2"].font = st["sub"]
    counts = {}
    for roles in d.loc[d["SVC YN"].eq("Yes"), "SVC Job Role"]:
        for role in str(roles).split(";"):
            role = role.strip()
            if role:
                counts[role] = counts.get(role, 0) + 1
    for i, label in enumerate(("SVC Job Role", "People"), start=1):
        c = ws.cell(row=4, column=i, value=label)
        c.font, c.border = st["head"], st["rule"]
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 10
    for r, (role, n) in enumerate(sorted(counts.items(), key=lambda kv: -kv[1]), start=5):
        ws.cell(row=r, column=1, value=role).font = st["body"]
        ws.cell(row=r, column=2, value=n).font = st["body"]
        for col in (1, 2):
            ws.cell(row=r, column=col).border = st["rule"]


def write_summary_sheet(wb, d, st, args):
    ws = wb.create_sheet("Summary", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{args.wave} Soft Live Participants"
    ws["A1"].font = st["big"]
    ws["A2"] = f"data as of {args.stamp:%m/%d/%Y %I:%M %p} · go-live {GO_LIVE}"
    ws["A2"].font = st["sub"]
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 14

    row = 4
    ws.cell(row=row, column=1, value="Total people").font = st["section"]
    ws.cell(row=row, column=2, value=len(d)).font = st["kpi_accent"]
    row += 2
    for title, series in (("By population", d["Population"]),
                          ("SVC staff", d["SVC YN"]),
                          ("Training status", d["Training Status"].fillna("(not reported)"))):
        ws.cell(row=row, column=1, value=title).font = st["section"]
        row += 1
        for k, v in series.value_counts().items():
            ws.cell(row=row, column=1, value=str(k)).font = st["body"]
            ws.cell(row=row, column=2, value=int(v)).font = st["body"]
            for col in (1, 2):
                ws.cell(row=row, column=col).border = st["rule"]
            row += 1
        row += 1
    ws.cell(row=row, column=1, value="Email present").font = st["section"]
    ws.cell(row=row, column=2, value=int(d["Email"].notna().sum())).font = st["body"]


def main():
    ap = argparse.ArgumentParser(description="Build the Wave 3 soft live participant list.")
    ap.add_argument("--wave", default=WAVE)
    ap.add_argument("--summary-export", type=Path, default=SUMMARY_EXPORT)
    ap.add_argument("--mvp-csv", type=Path, default=MVP_CSV)
    ap.add_argument("--out")
    ap.add_argument("--all-waves-svc", action="store_true",
                    help="Include SVC staff from any wave, not just the target wave.")
    ap.add_argument("--show-config", action="store_true")
    args = ap.parse_args()
    args.stamp = datetime.now()

    out = Path(args.out) if args.out else op.SQL_PULLS_DIR / OUT_NAME.format(
        wave=args.wave, date=f"{args.stamp:%Y-%m-%d}")
    if out.parent == Path("."):
        out = op.SQL_PULLS_DIR / out.name

    if args.show_config:
        print(f"wave           : {args.wave}\nsummary export : {args.summary_export}\n"
              f"mvp csv        : {args.mvp_csv}\noutput         : {out}\n"
              f"columns        : {len(COLUMNS)}")
        return

    d, wave_ids, svc_ids, _mvp, _summary = load(args)
    d = derive(d, wave_ids, svc_ids, args)

    if out.exists():
        archive = out.parent / "archive"
        archive.mkdir(exist_ok=True)
        backup = archive / f"{out.stem}.bak.{args.stamp:%Y-%m-%d_%H%M}.xlsx"
        backup.write_bytes(out.read_bytes())
        print(f"previous version backed up: archive\\{backup.name}")

    st = styles()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_list_sheet(wb, d, st, args)
    write_svc_roles_sheet(wb, d, st)
    write_summary_sheet(wb, d, st, args)
    wb.save(out)

    errs = {"#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "#SPILL!"}
    wbv = openpyxl.load_workbook(out)
    bad = sum(1 for s in wbv.sheetnames for r in wbv[s].iter_rows()
              for c in r if isinstance(c.value, str) and c.value in errs)
    print(f"\nwrote {len(d):,} rows x {len(COLUMNS)} columns -> {out}")
    print(f"  sheets       : {', '.join(wbv.sheetnames)}")
    print(f"  email present: {int(d['Email'].notna().sum()):,} of {len(d):,}")
    print(f"  SVC flagged  : {int(d['SVC YN'].eq('Yes').sum()):,}")
    print(f"  QA           : {bad} error value(s), "
          f"{int(d['UniversalID'].duplicated().sum())} duplicate ID(s)")


if __name__ == "__main__":
    main()
