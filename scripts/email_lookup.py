# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
email_lookup.py — the ONE place an end-user email address is resolved.

Every person-level output (Master, Team File, Training Tracker, LAVA, Soft Live)
imports this, so the same person shows the same address everywhere. Her ask,
2026-09-02: "All of our output files should match with data and email addresses
for end user."

    Email = MVP UserUPN, falling back to report.hr.Email

MVP IS THE SOURCE; HR IS THE FALLBACK. Per the analyst 2026-09-02: **HR is used only
to map the leader chain (Manager -> SVP), never as a reporting source.** Same
rule that already governs business unit and system-state fields, which come from
raw.mvp and never from report.hr.

The addresses themselves show why. UserUPN is the person's YourOrg ACCOUNT —
8,503 of 8,508 on the Master are @yourorg.edu (the other 5 are legacy.example, the
legacy YourOrg domain). HR's Email is a contact address and hands out 49
non-YourOrg ones: flexstaff.org, vendor.example, health-roi.com, ikshealth.com,
pds-online.com. Those are vendor and affiliate domains — no use for Epic or LAVA
provisioning, which needs the YourOrg account. MVP also simply covers more
people: 8,508 against HR's 7,860.

Where the two differ (345 people on the Master), HR is often a different
identity entirely — JDOE2 is jdoe@vendor.example in HR and
jdoe@yourorg.example in MVP; JSMITH is jsmith@ in HR and jsmith@ in MVP.

UserUPN lives ONLY in the MVP CSV — raw.mvp has no such column — and that CSV is
~115 MB, so parsing it on every run would add a minute to the 2-hourly tracker
build. The UPN slice is cached to a small parquet beside the raw file and rebuilt
only when the CSV changes.

    from email_lookup import email_map, attach_email

    emails = email_map(eng)              # {UNIVERSALID: address}
    df = attach_email(df, eng)           # adds an 'Email' column, keyed on UniversalID
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import onedrive_paths as op   # noqa: E402

CACHE = op.RAW_MVP_DIR / ".user_upn_cache.parquet"


def _clean(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    return s.where(s.str.contains("@", na=False) & ~s.str.lower().isin(["nan", "none"]))


def _mvp_upn() -> pd.Series:
    """UniversalID -> UserUPN, cached against the CSV's mtime."""
    csv = op.RAW_MVP_USER_MAPPINGS
    if not csv.exists():
        return pd.Series(dtype=object, name="MVP")
    stamp = csv.stat().st_mtime
    if CACHE.exists() and CACHE.stat().st_mtime >= stamp:
        try:
            return pd.read_parquet(CACHE)["MVP"]
        except Exception:
            pass                      # unreadable cache -> just rebuild it
    df = pd.read_csv(csv, usecols=["UniversalID", "UserUPN"], dtype=str,
                     on_bad_lines="skip")
    df["uid"] = df["UniversalID"].astype(str).str.strip().str.upper()
    df["MVP"] = _clean(df["UserUPN"])
    out = df.dropna(subset=["MVP"]).drop_duplicates("uid").set_index("uid")[["MVP"]]
    try:
        out.to_parquet(CACHE)
    except Exception:
        pass                          # cache is an optimisation, never a requirement
    return out["MVP"]


def email_sources(eng) -> pd.DataFrame:
    """Both sources plus the resolved address, indexed by UniversalID."""
    hr = pd.read_sql("SELECT UPPER(LTRIM(RTRIM(UniversalID))) uid, Email"
                     " FROM report.hr WHERE Email IS NOT NULL", eng)
    hr["HR"] = _clean(hr["Email"])
    hr = hr.dropna(subset=["HR"]).drop_duplicates("uid").set_index("uid")["HR"]

    df = pd.concat([hr, _mvp_upn()], axis=1)
    df["Email"] = df["MVP"].fillna(df["HR"])
    df.index.name = "uid"
    return df


def email_map(eng) -> dict:
    """{UNIVERSALID: address} — the lookup most callers want."""
    s = email_sources(eng)["Email"].dropna()
    return s.to_dict()


def attach_email(df: pd.DataFrame, eng, uid_col: str = "UniversalID",
                 out_col: str = "Email") -> pd.DataFrame:
    """Add/refresh an email column on a person-level frame, keyed on uid_col.

    An address already present is kept — a source that has nothing for someone
    must never blank an address another source already supplied."""
    m = email_map(eng)
    keys = df[uid_col].astype(str).str.strip().str.upper()
    found = keys.map(m)
    df[out_col] = found if out_col not in df.columns else \
        _clean(df[out_col].astype(str)).fillna(found)
    return df
