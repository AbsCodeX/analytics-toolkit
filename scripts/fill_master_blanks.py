# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
fill_master_blanks.py

Fill BLANK User Type / Training Preference / Vendor Yes/No cells on the Master
DATA sheet. Blanks only — a populated cell is never touched, so hand-entered
values always win (see the "never overwrite manual edits" rule).

Why this exists (2026-08-26): people appended by add_missing_to_master.py land
with these three columns empty, and nothing downstream fills them. The MVP
updater maps User Type from MVP's "vendor (offshore/onshore)", so anyone not yet
in MVP stays blank — 34 of the 37 added on 2026-08-26 were in exactly that state.
Training Preference then stays blank too, because fill_training_preference()
derives Remote/Onsite FROM User Type or Vendor Yes/No. One root cause, three
empty columns.

Evidence used, best first — nothing is guessed:
  1. report.user_type_recommendations at Confidence='High'  -> its User Type
  2. HR WorkerType = 'Employee'                             -> YourOrg / Onsite
  3. HR WorkerType = 'VEN'                                  -> Training Preference
                                                               = Remote (the same
                                                               rule the MVP updater
                                                               applies to vendors)

Vendor Yes/No is NEVER written here: it is a formula column derived from User
Type. Fill User Type and it resolves itself.
A vendor's SHORE (Onshore vs Offshore) is deliberately left blank when no source
knows it: HR carries no business unit for them, they are absent from MVP, and
Epic has nothing, so user_type_recommendations returns NO_SIGNAL. Guessing there
would feed a wrong number straight into vendor headcounts. Those normally
resolve on their own once MVP picks the person up.

    python scripts\\fill_master_blanks.py            # preview + CSV, writes nothing
    python scripts\\fill_master_blanks.py --apply    # write (Excel must be closed)
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "sql"))
import master_backup                                    # noqa: E402
import onedrive_paths as op                             # noqa: E402
from refresh import CONN                                # noqa: E402
from sqlalchemy import create_engine, text              # noqa: E402

DATA_SHEET = "DATA"
TABLE_NAME = "Table1"
USER_TYPE = "User Type (Offshore/Onshore/YourOrg)"
TRAIN_PREF = "Training Preference"
VENDOR_YN = "Vendor Yes/No"
CSV_OUT = op.LOCAL_MAIN_REPORTS_DIR / "master_blank_fills_applied.csv"

REMOTE_TYPES = {"offshore", "onshore"}


def office_lock_present() -> bool:
    p = op.MASTER_WAVE_PATH
    return (p.parent / ("~$" + p.name)).exists()


def proposals() -> pd.DataFrame:
    """One row per person with a blank in any of the three columns, plus what
    the evidence supports. Blank proposal = leave the cell alone."""
    eng = create_engine(CONN)
    # raw.master column names are sanitised by the loader, so select explicitly
    df = pd.read_sql(text("""
        SELECT m.UniversalID,
               m.Full_Name                              AS FullName,
               m.Leaders,
               m.GoLiveWave,
               m.[User_Type_Offshore_Onshore_YourOrg] AS Cur_UserType,
               m.[Training_Preference]                  AS Cur_TrainPref,
               m.[Vendor_Yes_No]                        AS Cur_VendorYN,
               h.WorkerType                             AS HR_WorkerType,
               u.UserType_Recommended, u.Confidence
        FROM raw.[master] m
        LEFT JOIN report.hr h
               ON h.UniversalID = UPPER(LTRIM(RTRIM(m.UniversalID)))
        LEFT JOIN report.user_type_recommendations u
               ON u.UniversalID = UPPER(LTRIM(RTRIM(m.UniversalID)))
        WHERE LTRIM(RTRIM(ISNULL(m.[User_Type_Offshore_Onshore_YourOrg],''))) = ''
           OR LTRIM(RTRIM(ISNULL(m.[Training_Preference],'')))                  = ''
           OR LTRIM(RTRIM(ISNULL(m.[Vendor_Yes_No],'')))                        = ''
    """), eng)

    def row(r):
        wt = str(r.HR_WorkerType or "").strip().upper()
        rec = str(r.UserType_Recommended or "").strip()
        high = str(r.Confidence or "").strip() == "High"
        utype = pref = vend = ""
        why = "No signal"
        if high and rec:
            utype, why = rec, "Epic/BU rule"
            pref = "Remote" if rec.lower() in REMOTE_TYPES else "Onsite"
            vend = "Yes" if rec.lower() in REMOTE_TYPES else "No"
        elif wt == "EMPLOYEE":
            utype, pref, vend, why = "YourOrg", "Onsite", "No", "HR WorkerType"
        elif wt == "VEN":
            pref, vend, why = "Remote", "Yes", "HR vendor, shore unknown"
        return utype, pref, vend, why

    out = df.apply(row, axis=1, result_type="expand")
    out.columns = ["New_UserType", "New_TrainPref", "New_VendorYN", "Evidence"]
    df = pd.concat([df, out], axis=1)
    # only propose into cells that are actually blank
    for cur, new in ((("Cur_UserType"), "New_UserType"),
                     (("Cur_TrainPref"), "New_TrainPref"),
                     (("Cur_VendorYN"), "New_VendorYN")):
        filled = df[cur].fillna("").astype(str).str.strip() != ""
        df.loc[filled, new] = ""
    keep = ((df.New_UserType != "") | (df.New_TrainPref != ""))
    return df[keep].reset_index(drop=True)


def main(apply: bool) -> None:
    if office_lock_present():
        sys.exit("ERROR: Master Wave File is open in Excel. Close it and re-run.")

    prop = proposals()
    if prop.empty:
        print("Nothing to fill — no blank User Type / Training Preference / Vendor cells.")
        return

    print(f"People with a fillable blank: {len(prop)}\n")
    for col, label in ((("New_UserType"), "User Type"),
                       (("New_TrainPref"), "Training Preference")):
        vc = prop[col].replace("", pd.NA).dropna().value_counts()
        print(f"  {label:22s} {int(vc.sum()):>3} cell(s)  " +
              (", ".join(f"{k}={v}" for k, v in vc.items()) if len(vc) else "—"))
    print("\nEvidence:")
    print(prop["Evidence"].value_counts().to_string())
    show = ["UniversalID", "FullName", "Leaders", "New_UserType",
            "New_TrainPref", "Evidence"]
    pd.set_option("display.max_rows", None, "display.width", 220, "display.max_colwidth", 26)
    print("\nProposed fills:")
    print(prop[show].to_string(index=False))
    CSV_OUT.parent.mkdir(parents=True, exist_ok=True)
    prop.to_csv(CSV_OUT, index=False)
    print(f"\nReview CSV: {CSV_OUT}")

    if not apply:
        print("\nDRY RUN — nothing written. Re-run with --apply to write.")
        return

    import xlwings as xw

    # Vendor Yes/No is DELIBERATELY not written. It is a FORMULA column on the
    # DATA sheet — =IF(OR(X{r}="Onshore",X{r}="Offshore"),"Yes",
    # IF(X{r}="YourOrg","No","")) — derived from User Type. Writing a literal
    # into it replaces the formula for that row and asserts a value the Master's
    # own logic would not produce. Learned the hard way on 2026-08-26: 33 cells
    # were overwritten with "Yes" and had to be restored. Fill User Type and the
    # formula fills this itself. Before adding any column here, check whether it
    # holds a formula.
    targets = {USER_TYPE: "New_UserType", TRAIN_PREF: "New_TrainPref"}
    written = {k: 0 for k in targets}
    skipped_nonblank = 0

    master_backup.ensure_daily_full_backup(op.MASTER_WAVE_PATH)
    app = wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = app.screen_updating = False
        wb = app.books.open(str(op.MASTER_WAVE_PATH), update_links=False)
        ws = wb.sheets[DATA_SHEET]
        lo = ws.tables[TABLE_NAME].api
        heads = [str(h).strip() if h is not None else "" for h in lo.HeaderRowRange.Value[0]]
        first_row = lo.Range.Row + 1
        n_rows = lo.Range.Rows.Count - 1
        base_col = lo.Range.Column

        missing = [h for h in targets if h not in heads]
        if missing:
            sys.exit(f"ERROR: column(s) not found on DATA: {missing}")

        uid_col = base_col + heads.index("UniversalID")
        uids = ws.range((first_row, uid_col), (first_row + n_rows - 1, uid_col)).value
        uid_row = {}
        for i, u in enumerate(uids):
            if u:
                uid_row.setdefault(str(u).strip().upper(), first_row + i)

        for _, r in prop.iterrows():
            rn = uid_row.get(str(r.UniversalID).strip().upper())
            if rn is None:
                continue
            for header, field in targets.items():
                val = str(getattr(r, field) or "").strip()
                if not val:
                    continue
                col = base_col + heads.index(header)
                cell = ws.range((rn, col))
                if str(cell.value or "").strip():      # never overwrite
                    skipped_nonblank += 1
                    continue
                cell.value = val
                written[header] += 1
        wb.save()
    finally:
        try:
            if wb is not None:
                wb.close()
        finally:
            if app is not None:
                app.quit()

    print("\nWritten:")
    for k, v in written.items():
        print(f"  {k:40s} {v} cell(s)")
    if skipped_nonblank:
        print(f"  (skipped {skipped_nonblank} cell(s) that were not blank after all)")
    snap = master_backup.save_data_snapshot(op.MASTER_WAVE_PATH)
    print(f"DATA snapshot: {snap}")


if __name__ == "__main__":
    main("--apply" in sys.argv[1:])
