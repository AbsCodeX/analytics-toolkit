# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
build_leader_job_role_docs.py

Builds one "Job Role Analysis Summary" Word document per leader — the same
format as the hand-maintained Nickname07 Lastname07 reference, but generated from the
warehouse so the job roles and class lists are current.

Source        report.leader_role_training_map  (Leader / Wave / JobCategoryName
              / TrainingTrack / ClassTitle), driven by raw.mvp_job_categories
              TrainingTrack1-6 and refreshed by the daily pipeline.
Scope         All waves merged — a leader's roles are deduplicated across
              Wave 1-4, matching the reference doc, which has no wave dimension.
Template      The Lastname07 .docx itself is opened as the template, so styles,
              theme, margins and the bullet numbering definition are byte-for-
              byte the ones already in use. Headings and bullets are produced by
              deep-copying real paragraphs out of it, not by restyling new ones.
Class naming  Where the reference doc names a class, that exact string is
              reused. Anything it never listed is rendered through the same
              convention (_patti_case) derived from the strings it does contain.
Output        <OneDrive>/data/processed/ad_hoc/job_role_analysis/
              "Job Role Analysis Summary_<First Last> - MM.DD.YY.docx"
              plus "Job Role Analysis - Reference.xlsx" (Leader Summary, All
              Roles, Role Classes, Class Lookup, Not in Docs, Change Log). The
              previous set is moved to archive/<its date>/ first and the Change
              Log rows are carried forward, so re-running is a full refresh
              (a same-day re-run just overwrites today's files).
Retired roles Lastname07, Lastname16 and Lastname15 (2026-08-14) drop roles ending
              in an application suffix such as "(Hospital Billing)" and roles
              containing the word "edit" or "enhancement" - see
              RETIRED_ROLE_LEADERS. Dropped roles appear on Not in Docs.

eLearnings    The reference doc also lists eLearning / attestation items that
              report.leader_role_training_map does not carry, because that view
              expands a track to classes through raw.epic_status filtered to
              Event_Class_Type = 'Session'. The same table holds them as
              'Online Class' / 'Video' / 'Material' rows against the same
              Curriculum, so this script reads them directly and appends them to
              the role that owns the track. Five of the six items the reference
              lists by hand come back this way, including "Epic_Overview of
              Release of Information - for NON-HIM" (23 occurrences there).
              Set INCLUDE_ELEARNINGS = False for classes only.

Known gap     45 training tracks across 56 job roles have no class rows of any
              type in raw.epic_status — the Epic status export does not span
              every curriculum. Those roles are omitted from the documents and
              counted in coverage_summary.csv under "Roles Skipped (No Class)",
              which is a data-coverage gap, NOT a statement that the role needs
              no training. Only 11 roles are genuinely tagged "No Training".
"""

import copy
import datetime
import io
import re
import sys
from pathlib import Path

# Windows console defaults to cp1252; force UTF-8 so class names print cleanly
if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import docx
import pandas as pd
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

BASE = Path(r"C:\path\to\analytics")
sys.path.insert(0, str(BASE / "scripts"))
sys.path.insert(0, str(BASE / "sql"))
from onedrive_paths import ONEDRIVE_ROOT          # noqa: E402
from leader_names import CANONICAL_NAMES          # noqa: E402

TEMPLATE_PATH = (ONEDRIVE_ROOT / "data" / "raw" / "ad_hoc" /
                 "Job Role Analysis Summary_Patti Lastname07 - 07.09.26 revised 1.docx")
OUT_DIR = ONEDRIVE_ROOT / "data" / "processed" / "ad_hoc" / "job_role_analysis"

# Older per-wave Epic exports, kept locally. Only ever read as a fallback for
# curricula the warehouse's current-wave export does not contain.
EPIC_EXPORT_DIR = BASE / "EXPORTS" / "EPIC EXPORTS"

# ---------------------------------------------------------------------------
# eLearning / WBT layer
# ---------------------------------------------------------------------------
# report.leader_role_training_map expands a track to classes through
# raw.epic_status filtered to Event_Class_Type = 'Session', which is why
# eLearnings never reach it. They are in the same table under the types below,
# hung off the same Curriculum (= TrainingTrack), so the role -> eLearning
# mapping is derivable rather than something that has to be supplied by hand.
ELEARNING_CLASS_TYPES = ("Online Class", "Video", "Material")

# Event_Class_Type values that count as a class the reference doc would list.
# 'Test' is excluded on purpose: those rows are the _ASSESSMENT and
# _ATTESTATION OF COMPLETION companions, which the reference never bullets
# separately (it folds the attestation into the eLearning's own line).
CLASS_TYPES_KEPT = ("Session",) + ELEARNING_CLASS_TYPES

# Set False to publish instructor-led classes only.
INCLUDE_ELEARNINGS = True

# Manual additions for anything Epic does not carry, keyed by JobCategoryName
# (matched case-insensitively) and appended after everything derived from data.
ROLE_ELEARNINGS: dict[str, list[str]] = {}

# Roles a leader asked to see that MVP does not define yet. Each entry prints
# on that leader's document with the class list of the role it mirrors, is
# listed on Role Classes with Source = "Manual", and is dropped from here once
# MVP carries the role (the data path then takes over automatically).
MANUAL_ROLES: dict[str, list[dict]] = {
    "Lastname05, Firstname05": [
        {"role": "MRC HIM Quality - Auditor Offshore Vendor",
         "mirror": "MRC HIM Quality - Auditor Onshore Vendor",
         "added": "2026-09-11",
         "note": "Not in MVP job categories yet; mirrors the Onshore role"},
    ],
}
SRC_MANUAL = "Manual"


# ---------------------------------------------------------------------------
# Class-name rendering
# ---------------------------------------------------------------------------
# Words the reference doc leaves lowercase inside a title ("Billing for CCMC
# Admin", "Prepayments in Credit Workqueues"). Never applied to the first word.
_SMALL_WORDS = {"a", "an", "and", "at", "by", "for", "in", "of", "on", "or",
                "the", "to", "with"}

# Tokens that stay fully uppercase — Epic application codes and domain acronyms
# that appear inside class titles rather than in the DEPT_ prefix.
_ACRONYMS = {
    "HB", "PB", "SBO", "HIM", "ROI", "CDI", "CCMC", "DPU", "CDO", "EVS",
    "HEDIS", "VBPM", "ED", "FA", "L&D", "EMP", "SER", "WBT", "WQ", "OGA",
    "WCNF", "SS", "VS", "CBO", "STARS", "PAS", "GI", "PC", "LDA", "ADL",
    "ADLS", "IV", "ASU", "CCRC", "MRC", "SPARCS", "PWC", "DPU", "HHHB",
}
# "OR" is deliberately absent — it is the conjunction in "(WITHOUT HEDIS OR
# VBPM)" far more often than it is the operating room.

# Product names Epic writes in caps that are not acronyms and are not plain
# title case either.
_SPECIAL_CASE = {
    "MYCHART": "MyChart",
    "MYYOURORG": "MyYourOrg",
    "SLICERDICER": "SlicerDicer",
    "EPIC": "Epic",
}


def _cap_word(word: str, first: bool) -> str:
    """Title-case one whitespace-delimited token, preserving acronyms."""
    # Split on characters the reference doc capitalizes across: FOLLOW-UP ->
    # Follow-Up, SELF-PAY -> Self-Pay, (HB/PB) stays (HB/PB).
    def cap_piece(piece: str, is_first_piece: bool) -> str:
        core = piece.strip("()")
        lead = "(" * (len(piece) - len(piece.lstrip("(")))
        trail = ")" * (len(piece) - len(piece.rstrip(")")))
        if not core:
            return piece
        # A slash group is all-acronym only when every side is an acronym
        if "/" in core and all(s in _ACRONYMS for s in core.split("/") if s):
            return lead + core + trail
        if core in _ACRONYMS:
            return lead + core + trail
        # Trailing punctuation ("LEAD:") must not defeat the lookups
        stripped = core.rstrip(":;,.")
        punct = core[len(stripped):]
        if stripped.upper() in _SPECIAL_CASE:
            return lead + _SPECIAL_CASE[stripped.upper()] + punct + trail
        if core.isdigit():
            return lead + core + trail
        # Course codes ("LTC200") keep their shape rather than becoming "Ltc200"
        if any(ch.isdigit() for ch in stripped) and any(ch.isalpha() for ch in stripped) \
                and stripped.isupper():
            return lead + core + trail
        lower = core.lower()
        if lower in _SMALL_WORDS and not (first and is_first_piece):
            return lead + lower + trail
        return lead + core.capitalize() + trail

    parts = word.split("-")
    return "-".join(cap_piece(p, i == 0) for i, p in enumerate(parts))


def _patti_case(class_title: str) -> str:
    """
    Render a raw Epic ClassTitle in the reference doc's convention.

    EPIC_NH_HB_HOSPITAL BILLER 1          ->  Epic_HB_Hospital Biller 1
    EPIC_NH_SBO_SELF-PAY CREDITS          ->  Epic_SBO_Self-Pay Credits
    EPIC_ACTIVATING PATIENTS FOR MYCHART  ->  Epic_Activating Patients for MyChart

    The EPIC_ / EPIC_NH_ prefix collapses to Epic_, an application code keeps its
    uppercase, and the descriptive tail is title-cased. The NH_ segment and the
    application code are both optional — eLearning titles carry neither.
    """
    s = (class_title or "").strip()
    if not s:
        return s
    m = re.match(
        r"^EPIC[_ ](?:(?:NH|W1|W2|W3|W4|ADJ)[_ ])*(?:(?P<dept>[A-Z0-9&]{2,6})_)?"
        r"(?P<rest>.+)$", s, re.I)
    if not m:
        # No recognizable prefix — title-case the whole thing, leave it intact
        # rather than guessing at a structure that isn't there.
        return " ".join(_cap_word(w, i == 0) for i, w in enumerate(s.split()))
    rest = m.group("rest")
    tail = " ".join(_cap_word(w, i == 0) for i, w in enumerate(rest.split()))
    dept = m.group("dept")
    return f"Epic_{dept.upper()}_{tail}" if dept else f"Epic_{tail}"


def _strip_wave_prefix(s: str) -> str:
    """
    EPIC_W2_HB_HOSPITAL BILLER 1 -> EPIC_HB_HOSPITAL BILLER 1

    Epic tags the same curriculum and class with the wave it was delivered for:
    MVP's job categories say EPIC_NH_, the W2 export says EPIC_W2_. Dropping
    that segment is what lets a track in one source join a class in another,
    and what keeps two spellings of one class from printing as two bullets.
    """
    # Repeating group: Epic stacks them, e.g. EPIC_ADJ_W2_PWC SMART OVERVIEW.
    return re.sub(r"^EPIC[_ ](?:(?:NH|W1|W2|W3|W4|ADJ)[_ ])+", "EPIC_",
                  (s or "").strip().upper())


def _match_key(s: str) -> str:
    """Normalized key for pairing reference-doc strings with Epic ClassTitles."""
    return re.sub(r"[^A-Z0-9]", "", _strip_wave_prefix(s))


def build_name_lexicon(template_path: Path) -> dict[str, str]:
    """
    Harvest ClassTitle -> exact display string from the reference document.

    Anchoring to the leader-facing strings already in circulation matters more
    than internal consistency: the reference spells the same class several ways
    ("Insurance Follow Up" / "Insurance Follow-Up"), so the most frequent
    spelling per class wins.
    """
    if not template_path.exists():
        print(f"  WARNING: reference doc not found at {template_path}")
        print("           falling back to derived naming for every class.")
        return {}
    from collections import Counter
    doc = docx.Document(template_path)
    counts: dict[str, Counter] = {}
    for para in doc.paragraphs:
        if para.style.name != "List Paragraph":
            continue
        text = para.text.strip().rstrip(",").strip()
        if not text:
            continue
        counts.setdefault(_match_key(text), Counter())[text] += 1
    return {k: c.most_common(1)[0][0] for k, c in counts.items()}


# ---------------------------------------------------------------------------
# Word document assembly
# ---------------------------------------------------------------------------

def _clone_after(body, proto_el, sect_pr):
    """Deep-copy a prototype paragraph element into the body before sectPr."""
    el = copy.deepcopy(proto_el)
    if sect_pr is not None:
        sect_pr.addprevious(el)
    else:
        body.append(el)
    return el


def _set_text(doc, el, text: str) -> None:
    """Replace a cloned paragraph's content with text, keeping its run format."""
    para = Paragraph(el, doc)
    runs = para.runs
    if runs:
        runs[0].text = text
        for extra in runs[1:]:
            extra._element.getparent().remove(extra._element)
    else:
        para.add_run(text)


def load_prototypes(template_path: Path):
    """Grab a clean Heading 1 and bullet paragraph to clone from."""
    doc = docx.Document(template_path)
    h1 = li = None
    for para in doc.paragraphs:
        if not para.text.strip():
            continue
        pPr = para._p.find(qn("w:pPr"))
        # Skip headings carrying a direct colour override (the reference marks
        # some in red) — the prototype has to be the unstyled baseline.
        has_override = pPr is not None and pPr.find(qn("w:rPr")) is not None and \
            pPr.find(qn("w:rPr")).find(qn("w:color")) is not None
        if h1 is None and para.style.name == "Heading 1" and not has_override:
            h1 = copy.deepcopy(para._p)
        if li is None and para.style.name == "List Paragraph":
            li = copy.deepcopy(para._p)
        if h1 is not None and li is not None:
            break
    if h1 is None or li is None:
        raise ValueError("Could not find Heading 1 / List Paragraph prototypes "
                         f"in {template_path.name}")
    return h1, li


def write_leader_doc(roles: list[tuple[str, list[str]]],
                     template_path: Path, out_path: Path,
                     h1_proto, li_proto) -> None:
    """Write one leader's document: a heading per role, its classes beneath.

    No in-document title — the reference doc carries the person's name in the
    filename only, and matching it was the explicit requirement.
    """
    doc = docx.Document(template_path)
    body = doc.element.body

    # Clear the reference content, keeping styles/numbering/section properties
    for para in list(doc.paragraphs):
        para._element.getparent().remove(para._element)
    sect_pr = body.find(qn("w:sectPr"))

    for role_name, classes in roles:
        el = _clone_after(body, h1_proto, sect_pr)
        _set_text(doc, el, f"{role_name}:")
        for class_name in classes:
            el = _clone_after(body, li_proto, sect_pr)
            _set_text(doc, el, class_name)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


# ---------------------------------------------------------------------------
# Retired roles (leader decision, 2026-08-14)
# ---------------------------------------------------------------------------
# Lastname07, Lastname16 and Lastname15 asked for retired roles to come off their
# documents: roles ending in one of these application suffixes, and roles whose
# name carries the word "edit" or "enhancement" ("Credit" does not match).
# Other leaders still use those roles, so the trim is per leader, and every
# dropped role is listed on the Reference workbook's "Not in Docs" sheet.
RETIRED_ROLE_LEADERS = {"Lastname07, Firstname07", "Lastname16, Firstname16", "Lastname15, Firstname15"}
RETIRED_SUFFIXES = ("(Hospital Billing)", "(Professional Billing)",
                    "(Single Billing Office)", "(Access Focus)")
_RETIRED_WORDS = re.compile(r"\b(edit|enhancement)\b", re.I)

REASON_SUFFIX = "Removed - retired application-suffix role"
REASON_WORDS = 'Removed - role name contains "edit" or "enhancement"'
REASON_NO_TRAINING = "Role is tagged No Training"
REASON_NO_CLASS = ("No class found in any Epic export - needs a "
                   "curriculum-to-class list")

REFERENCE_NAME = "Job Role Analysis - Reference.xlsx"

# Rolling curriculum-to-class memory. The Epic Curriculum Status Detail export
# is user-centric: a curriculum (and its classes) only appears while some
# current-wave user is enrolled in it, so a live class can vanish from the
# warehouse for a while. Every run unions what it sees into this file; a pair
# no longer in the export is kept only while Cornerstone still delivers the
# class (a session starting in the last CARRY_ACTIVE_DAYS or in the future).
# First run seeds it from the newest archived Reference workbook.
CLASS_MAP_PATH = BASE / "data" / "reference" / "job_role_analysis_class_map.csv"
CARRY_ACTIVE_DAYS = 180
SRC_EXPORT = "Epic export"
SRC_OLDER = "Older Epic export"
SRC_CARRIED = "Carried forward"
CHANGE_LOG_COLS = ["Date", "Raised By", "Leader", "Job Role", "Class",
                   "Change Type", "What Should Change", "Reason", "Status",
                   "Applied On"]


def retired_reason(role: str) -> str | None:
    """Why a role is trimmed for a RETIRED_ROLE_LEADERS leader, else None."""
    if role.strip().endswith(RETIRED_SUFFIXES):
        return REASON_SUFFIX
    if _RETIRED_WORDS.search(role):
        return REASON_WORDS
    return None


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------

def _engine():
    from sqlalchemy import create_engine
    from refresh import CONN
    return create_engine(CONN)


def load_map() -> pd.DataFrame:
    return pd.read_sql(
        "SELECT Leader, Wave, JobCategoryGroup, JobCategoryName, "
        "TrainingTrack, ClassTitle, UserCount "
        "FROM report.leader_role_training_map",
        _engine())


def load_supplemental_class_map() -> dict[str, list[tuple[str, str]]]:
    """
    Curriculum -> [(class, class type)] harvested from the Epic Curriculum
    Status Detail exports sitting on disk, keyed by wave-stripped curriculum.

    raw.epic_status holds only the current W3 export, and that export is
    user-centric: a curriculum appears in it only if some W3 user is enrolled.
    Tracks belonging to other waves therefore resolve to no classes at all. The
    older W1/W2/W3 exports still carry them, so they are read here and used
    strictly as a fallback for tracks the warehouse cannot cover.

    Joining across them needs the wave prefix stripped — MVP's
    EPIC_NH_PB HB: BILLING & INS is the W2 export's EPIC_W2_PB HB: BILLING & INS.
    """
    seen_files, frames = [], []
    for pattern in ("Curriculum Status Detail by User - *.xlsx",
                    "RevCycle Epic Curriculum Status Detail by User - *.xlsx",
                    "Curriculm Status Detail/*.xlsx"):
        for path in sorted(EPIC_EXPORT_DIR.glob(pattern)):
            try:
                frame = pd.read_excel(
                    path, engine="calamine",
                    usecols=["Curriculum", "Event (Class)", "Event (Class) Type"])
            except Exception as exc:                     # noqa: BLE001
                print(f"    skipped {path.name}: {exc}")
                continue
            frame.columns = ["Curriculum", "ClassTitle", "ClassType"]
            frames.append(frame)
            seen_files.append(path.name)
    if not frames:
        print("    no on-disk Curriculum Status Detail exports found")
        return {}

    df = pd.concat(frames, ignore_index=True).dropna(subset=["Curriculum", "ClassTitle"])
    df["ClassType"] = df["ClassType"].astype(str).str.strip()
    df = df[df["ClassType"].isin(CLASS_TYPES_KEPT)]
    df["key"] = df["Curriculum"].map(_strip_wave_prefix)
    df["ClassTitle"] = df["ClassTitle"].astype(str).str.strip()
    out: dict[str, dict[str, str]] = {}
    for key, title, ctype in (df[["key", "ClassTitle", "ClassType"]]
                              .drop_duplicates().itertuples(index=False)):
        out.setdefault(key, {}).setdefault(title, ctype)
    result = {k: sorted(v.items()) for k, v in out.items()}
    print(f"    {len(seen_files)} export files -> {len(result)} curricula")
    return result


def load_elearning_map() -> dict[str, list[tuple[str, str]]]:
    """
    TrainingTrack -> [(class, class type)] for eLearning / video / material
    rows in raw.epic_status.

    Keyed on the uppercased Curriculum, which is the same string the job-category
    TrainingTrack1-6 columns hold, so it joins straight onto a role's tracks.
    """
    placeholders = ", ".join(f"'{t}'" for t in ELEARNING_CLASS_TYPES)
    df = pd.read_sql(
        "SELECT DISTINCT UPPER(LTRIM(RTRIM(Curriculum))) AS TrackKey, "
        "LTRIM(RTRIM(Event_Class)) AS ClassTitle, "
        "LTRIM(RTRIM(Event_Class_Type)) AS ClassType "
        "FROM raw.epic_status "
        f"WHERE Event_Class_Type IN ({placeholders}) "
        "AND LTRIM(RTRIM(ISNULL(Event_Class, ''))) <> ''",
        _engine())
    out: dict[str, dict[str, str]] = {}
    for row in df.itertuples(index=False):
        out.setdefault(row.TrackKey, {}).setdefault(row.ClassTitle, row.ClassType)
    return {k: sorted(v.items()) for k, v in out.items()}


def leader_display(leader: str) -> str:
    """'Lastname07, Firstname07' -> 'Firstname07 Lastname07'"""
    last, _, first = leader.partition(",")
    return f"{first.strip()} {last.strip()}".strip()


def leader_filename(leader: str, run_date: datetime.date) -> str:
    """'Lastname07, Firstname07' -> 'Job Role Analysis Summary_Firstname07 Lastname07 - 08.10.26.docx'"""
    return (f"Job Role Analysis Summary_{leader_display(leader)} - "
            f"{run_date.strftime('%m.%d.%y')}.docx")


# ---------------------------------------------------------------------------
# Rolling curriculum-to-class memory
# ---------------------------------------------------------------------------

MEM_COLS = ["TrackKey", "ClassTitle", "ClassType", "FirstSeen", "LastSeen"]


def _seed_class_map_from_archive(out_dir: Path) -> pd.DataFrame:
    """Rebuild the memory from the newest archived Reference workbook."""
    candidates = sorted(p for p in (out_dir / "archive").glob(f"*/{REFERENCE_NAME}")
                        if not p.parent.name.endswith("_previous"))
    if not candidates:
        return pd.DataFrame(columns=MEM_COLS)
    src = candidates[-1]
    import openpyxl
    ws = openpyxl.load_workbook(src, read_only=True, data_only=True)["Role Classes"]
    rows = list(ws.iter_rows(values_only=True))
    idx = {str(h).strip(): i for i, h in enumerate(rows[0]) if h}
    stamp = src.parent.name[:10]
    out = []
    for r in rows[1:]:
        cur, cls, typ = (r[idx["Curriculum"]], r[idx["Class"]], r[idx["Class Type"]])
        if cur and cls:
            out.append((_strip_wave_prefix(str(cur)), str(cls).strip(),
                        str(typ or "Session").strip(), stamp, stamp))
    df = pd.DataFrame(out, columns=MEM_COLS).drop_duplicates(["TrackKey", "ClassTitle"])
    print(f"  class memory seeded from archive/{src.parent.name} ({len(df)} pairs)")
    return df


def load_cornerstone_active_classes(run_date: datetime.date) -> set[str]:
    """Match keys of Epic classes Cornerstone delivered recently or will deliver."""
    since = (run_date - datetime.timedelta(days=CARRY_ACTIVE_DAYS)).isoformat()
    df = pd.read_sql(
        "SELECT DISTINCT Training_Title FROM raw.cornerstone "
        f"WHERE Training_Provider = 'EPIC' AND Training_Start_Date >= '{since}'",
        _engine())
    return {_match_key(t) for t in df["Training_Title"].dropna().astype(str)}


def update_class_map(seen_today: dict[str, list[tuple[str, str]]],
                     run_date: datetime.date, out_dir: Path
                     ) -> tuple[dict[str, list[tuple[str, str]]], int, int]:
    """
    Union today's warehouse map into the memory file and return the pairs to
    carry forward: (TrackKey -> [(ClassTitle, ClassType)], kept, dropped).
    """
    today = run_date.isoformat()
    if CLASS_MAP_PATH.exists():
        mem = pd.read_csv(CLASS_MAP_PATH, dtype=str).fillna("")
    else:
        mem = _seed_class_map_from_archive(out_dir)
    # Keyed on the normalized class name so the seed's display names and the
    # export's raw titles ("Epic_HB_Hospital Biller 1" vs
    # "EPIC_NH_HB_HOSPITAL BILLER 1") are one entry; the raw title wins.
    rows: dict[tuple[str, str], list] = {}
    for r in mem.itertuples(index=False):
        rows.setdefault((r.TrackKey, _match_key(r.ClassTitle)),
                        [r.ClassTitle, r.ClassType, r.FirstSeen, r.LastSeen])

    seen_keys = set()
    for track, pairs in seen_today.items():
        for title, ctype in pairs:
            key = (track, _match_key(title))
            seen_keys.add(key)
            if key in rows:
                rows[key][0], rows[key][3] = title, today
            else:
                rows[key] = [title, ctype, today, today]

    active = load_cornerstone_active_classes(run_date)
    carried: dict[str, list[tuple[str, str]]] = {}
    kept = dropped = 0
    for key in list(rows):
        if key in seen_keys:
            continue
        track, class_key = key
        title, ctype = rows[key][0], rows[key][1] or "Session"
        if class_key in active:
            carried.setdefault(track, []).append((title, ctype))
            kept += 1
        else:
            del rows[key]
            dropped += 1

    out = pd.DataFrame([(t, *v) for (t, _), v in rows.items()], columns=MEM_COLS)
    out = out.sort_values(["TrackKey", "ClassTitle"])
    CLASS_MAP_PATH.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(CLASS_MAP_PATH, index=False)
    return carried, kept, dropped


# ---------------------------------------------------------------------------
# Previous outputs: carry the Change Log forward, archive everything else
# ---------------------------------------------------------------------------

def read_change_log(path: Path) -> list[list]:
    """Rows already written on the Change Log sheet (header excluded)."""
    if not path.exists():
        return []
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if "Change Log" not in wb.sheetnames:
        return []
    rows = [list(r) for r in wb["Change Log"].iter_rows(values_only=True)]
    rows = [r for r in rows if any(v is not None for v in r)]
    if not rows:
        return []
    header = [str(v).strip() if v is not None else "" for v in rows[0]]
    idx = {h: i for i, h in enumerate(header)}
    out = []
    for r in rows[1:]:
        out.append([r[idx[c]] if c in idx and idx[c] < len(r) else None
                    for c in CHANGE_LOG_COLS])
    return out


def archive_previous_outputs(out_dir: Path, run_date: datetime.date) -> Path | None:
    """
    Move the current documents and Reference workbook into
    out_dir/archive/<date of that set>/ so the folder only ever shows one set.
    """
    docs = sorted(out_dir.glob("Job Role Analysis Summary_*.docx"))
    extras = [p for p in (out_dir / REFERENCE_NAME, out_dir / "coverage_summary.csv")
              if p.exists()]
    if not docs and not extras:
        return None
    locked = [p.name for p in docs + extras if _is_locked(p)]
    if locked:
        sys.exit("ABORT: close these files (open in Excel/Word) and re-run - nothing "
                 "was moved:" + "".join(chr(10) + "  " + name for name in locked))
    stamp = None
    for p in docs:
        m = re.search(r" - (\d{2})\.(\d{2})\.(\d{2})\.docx$", p.name)
        if m:
            stamp = f"20{m.group(3)}-{m.group(1)}-{m.group(2)}"
            break
    if stamp is None:
        src = docs[0] if docs else extras[0]
        stamp = datetime.date.fromtimestamp(src.stat().st_mtime).isoformat()
    if stamp == run_date.isoformat():
        # Same-day re-run: today's files are simply overwritten below, and the
        # "previous set" is still whatever the archive already holds.
        existing = sorted(d for d in (out_dir / "archive").glob("*") if d.is_dir())
        print("  today's set is rebuilt in place (no archive copy)")
        return existing[-1] if existing else None
    dest = out_dir / "archive" / stamp
    dest.mkdir(parents=True, exist_ok=True)
    import shutil
    import time
    for p in docs + extras:
        target = dest / p.name
        for attempt in range(3):     # OneDrive briefly locks files it is syncing
            try:
                if target.exists():  # a previous run was interrupted mid-move
                    target.unlink()
                shutil.move(str(p), str(target))
                break
            except PermissionError:
                if attempt < 2:
                    time.sleep(2)
                    continue
                if p in extras:
                    # The workbook is rewritten in place below, so a copy into
                    # the archive is enough when the delete is blocked.
                    shutil.copy2(str(p), str(target))
                    print(f"  note: {p.name} is held open; archived a copy and "
                          "will overwrite it in place")
                else:
                    raise
    return dest


def _is_locked(path: Path) -> bool:
    """True when Excel/Word (or OneDrive) holds the file open for writing."""
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True


# ---------------------------------------------------------------------------
# Reference workbook
# ---------------------------------------------------------------------------

def write_reference_workbook(path: Path, run_date: datetime.date,
                             leader_rows: list[dict], role_class_rows: list[dict],
                             not_in_docs: list[dict], change_log: list[list],
                             archive_dir: Path | None) -> None:
    import xlsxwriter

    rc = pd.DataFrame(role_class_rows)
    rc["IsSession"] = rc["Class Type"].eq("Session")

    # ---- All Roles: one row per leader x role that made it into a document
    all_roles = (rc.groupby(["Leader", "Job Role Group", "Job Role", "Training Track",
                             "Users in Role", "Word Document"], sort=False)
                   .agg(**{"Classes in Doc": ("Class", "nunique"),
                           "Instructor-Led": ("IsSession", "sum"),
                           "eLearning / Other": ("IsSession", lambda s: int((~s).sum()))})
                   .reset_index())
    all_roles["In Document"] = "Yes"
    all_roles = all_roles.sort_values(["Leader", "Job Role Group", "Job Role"])[
        ["Leader", "Job Role Group", "Job Role", "Training Track", "Users in Role",
         "Classes in Doc", "Instructor-Led", "eLearning / Other", "In Document",
         "Word Document"]]

    # ---- Class Lookup: one row per class
    def _joined(s):
        return "; ".join(sorted(set(s)))
    lookup = (rc.groupby("Class", sort=True)
                .agg(**{"Class Type": ("Class Type", "first"),
                        "Curriculum": ("Curriculum", _joined),
                        "Job Roles": ("Job Role", "nunique"),
                        "Leaders": ("Leader", "nunique"),
                        "Job Roles With This Class": ("Job Role", _joined),
                        "Leaders With This Class": ("Leader", _joined)})
                .reset_index())

    nid = pd.DataFrame(not_in_docs, columns=[
        "Leader", "Job Role Group", "Job Role", "Training Track", "Users in Role",
        "Why It Is Not in the Document"])
    nid = nid.sort_values(["Leader", "Why It Is Not in the Document", "Job Role"])
    n_gap = int((nid["Why It Is Not in the Document"] == REASON_NO_CLASS).sum())

    ls = pd.DataFrame(leader_rows)[[
        "Leader", "Name on Document", "Job Roles in Doc", "Distinct Classes",
        "Total Class Lines", "Instructor-Led Lines", "eLearning / Other Lines",
        "Roles Not in Doc", "Word Document"]]

    rc_out = rc[["Leader", "Job Role Group", "Job Role", "Training Track", "Curriculum",
                 "Class", "Class Type", "Source", "Users in Role", "Word Document"]]
    n_carried = int((rc["Source"] == SRC_CARRIED).sum())

    previous = (f"The set this refresh replaced is in archive\\{archive_dir.name}."
                if archive_dir else "This is the first set in this folder.")
    readme = [
        ("Job Role Analysis", None, "title"),
        (None, None, None),
        ("What this folder is",
         "One Word document per leader listing each Epic job role their people hold "
         "and the classes attached to it. This workbook is the source behind those "
         "documents and the place to record any changes needed.", "section"),
        (None, None, None),
        ("The sheets", None, "section"),
        ("Leader Summary", "One row per leader: how many job roles and classes are in "
         "their document, and how many roles were left out.", "row"),
        ("All Roles", "Every leader-and-job-role pair that made it into a document, "
         "with its training track, headcount and class counts. Filter or copy this "
         "sheet for a full export of all jobs.", "row"),
        ("Role Classes", "The full grid - one row per leader, job role and class, with "
         "the class type and curriculum. Everything else on this workbook rolls up "
         "from here.", "row"),
        ("Class Lookup", "Search a class here to see every job role and leader that has "
         "it. Use the filter arrow on the Class column, or Ctrl+F. The last two "
         "columns list the roles and leaders in full.", "row"),
        ("Not in Docs", "Job roles held by a leader's people that are NOT printed in "
         "their document, and why. Most are missing because no Epic export lists a "
         "class for that curriculum - that is a data gap, not a statement that the "
         "role needs no training.", "row"),
        ("Change Log", "Record anything that should change here - a role to exclude for "
         "a leader, a class that is named wrong, a role that is missing. Rows carry "
         "forward from one refresh to the next; example rows can be overwritten or "
         "deleted.", "row"),
        (None, None, None),
        ("Worth knowing", None, "section"),
        ("Source", "Epic Curriculum Status Detail exports and the MVP job-category "
         f"training tracks, as loaded on {run_date.isoformat()} - the date on the "
         "document filenames.", "row"),
        ("Assessments", "Assessment and attestation items are deliberately not listed on "
         "the documents - they are folded into the class they belong to.", "row"),
        ("Carried forward", "The Epic export only lists a curriculum while a current-wave "
         "user is enrolled in it, so a live class can drop out of it for a while. "
         f"{n_carried} class lines marked Carried forward on Role Classes were on the "
         "previous documents and are still on the Cornerstone class schedule, so they "
         "stay. A class leaves once Cornerstone stops delivering it.", "row"),
        ("Coverage gap", f"{n_gap} leader-and-role combinations have no class in any "
         "available Epic export. Getting a complete curriculum-to-class list from "
         "Epic would close this.", "row"),
        ("Retired roles", "Retired roles are left off the Lastname07, Lastname16 and "
         "Lastname15 documents only - roles ending in an application suffix such as "
         "(Hospital Billing), and roles containing the word edit or enhancement. "
         "Other leaders still use those roles. Each one is listed on Not in Docs.",
         "row"),
        ("Manual roles", "Roles a leader asked for that MVP does not define yet are "
         "printed with the class list of the role they mirror and show Source = "
         "Manual on Role Classes (headcount 0 until MVP carries the role). Each one "
         "is logged on Change Log.", "row"),
        ("Previous set", previous, "row"),
    ]

    wb = xlsxwriter.Workbook(str(path))
    TEXT, MUTED, LINE = "#37352F", "#6B6B6B", "#E0E0E0"
    base = {"font_name": "Calibri", "font_size": 11, "font_color": TEXT,
            "valign": "top", "text_wrap": True}
    f_cell = wb.add_format(base)
    f_num = wb.add_format({**base, "num_format": "0", "align": "right"})
    f_date = wb.add_format({**base, "num_format": "mm/dd/yyyy"})
    f_muted = wb.add_format({**base, "font_color": MUTED})
    f_head = wb.add_format({**base, "bold": True, "font_color": MUTED,
                            "bottom": 1, "bottom_color": LINE})
    f_title = wb.add_format({**base, "bold": True, "font_size": 16, "text_wrap": False})
    f_section = wb.add_format({**base, "bold": True, "text_wrap": False})
    f_label = wb.add_format({**base, "text_wrap": False})
    f_note = wb.add_format({**base, "font_size": 10, "font_color": MUTED})

    # ---- Read Me
    ws = wb.add_worksheet("Read Me")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 26.7)
    ws.set_column("B:B", 96.7)
    for r, (a, b, kind) in enumerate(readme):
        if a is None:
            continue
        fmt = {"title": f_title, "section": f_section, "row": f_label}[kind]
        ws.write(r, 0, a, fmt)
        if b:
            ws.write(r, 1, b, f_note)

    def table(name, df, widths, muted_cols=(), num_cols=(), date_cols=()):
        ws = wb.add_worksheet(name)
        ws.hide_gridlines(2)
        cols = list(df.columns)
        for i, c in enumerate(cols):
            ws.set_column(i, i, widths.get(c, 18))
            ws.write(0, i, c, f_head)
        for r, row in enumerate(df.itertuples(index=False), start=1):
            for i, v in enumerate(row):
                c = cols[i]
                fmt = (f_muted if c in muted_cols else f_num if c in num_cols
                       else f_date if c in date_cols else f_cell)
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    ws.write_blank(r, i, None, fmt)
                elif isinstance(v, (datetime.date, datetime.datetime, pd.Timestamp)):
                    ws.write_datetime(r, i, pd.Timestamp(v).to_pydatetime(), f_date)
                elif isinstance(v, (int, float)) and c in num_cols:
                    ws.write_number(r, i, v, fmt)
                else:
                    ws.write(r, i, v, fmt)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(df), 1), len(cols) - 1)
        return ws

    counts = ("Job Roles in Doc", "Distinct Classes", "Total Class Lines",
              "Instructor-Led Lines", "eLearning / Other Lines", "Roles Not in Doc",
              "Users in Role", "Classes in Doc", "Instructor-Led", "eLearning / Other",
              "Job Roles", "Leaders")
    table("Leader Summary", ls,
          {"Leader": 22.7, "Name on Document": 20.7, "Word Document": 46.7,
           "Job Roles in Doc": 14.7, "Distinct Classes": 14.7, "Total Class Lines": 14.7,
           "Instructor-Led Lines": 16.7, "eLearning / Other Lines": 18.7,
           "Roles Not in Doc": 14.7},
          muted_cols=("Word Document",), num_cols=counts)
    table("All Roles", all_roles,
          {"Leader": 22.7, "Job Role Group": 16.7, "Job Role": 46.7,
           "Training Track": 40.7, "Users in Role": 12.7, "Classes in Doc": 12.7,
           "Instructor-Led": 12.7, "eLearning / Other": 14.7, "In Document": 12.7,
           "Word Document": 46.7},
          muted_cols=("Word Document",), num_cols=counts)
    table("Role Classes", rc_out,
          {"Leader": 22.7, "Job Role Group": 16.7, "Job Role": 46.7,
           "Training Track": 40.7, "Curriculum": 34.7, "Class": 46.7,
           "Class Type": 14.7, "Source": 18.7, "Users in Role": 12.7,
           "Word Document": 46.7},
          muted_cols=("Word Document",), num_cols=counts)
    table("Class Lookup", lookup,
          {"Class": 46.7, "Class Type": 14.7, "Curriculum": 40.7, "Job Roles": 10.7,
           "Leaders": 10.7, "Job Roles With This Class": 70.7,
           "Leaders With This Class": 50.7},
          num_cols=counts)
    table("Not in Docs", nid,
          {"Leader": 22.7, "Job Role Group": 16.7, "Job Role": 46.7,
           "Training Track": 40.7, "Users in Role": 12.7,
           "Why It Is Not in the Document": 60.7},
          num_cols=counts)

    cl = (pd.DataFrame(change_log, columns=CHANGE_LOG_COLS) if change_log
          else pd.DataFrame(columns=CHANGE_LOG_COLS))
    ws = table("Change Log", cl,
               {"Date": 12.7, "Raised By": 18.7, "Leader": 22.7, "Job Role": 40.7,
                "Class": 40.7, "Change Type": 16.7, "What Should Change": 50.7,
                "Reason": 40.7, "Status": 10.7, "Applied On": 12.7},
               date_cols=("Date", "Applied On"))
    # Leave the filter covering room for new rows
    ws.autofilter(0, 0, max(len(cl), 1) + 200, len(CHANGE_LOG_COLS) - 1)
    wb.close()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    run_date = datetime.date.today()

    if not TEMPLATE_PATH.exists():
        sys.exit(f"ABORT: reference template not found:\n  {TEMPLATE_PATH}")

    print("Loading class-name lexicon from the reference document...")
    lexicon = build_name_lexicon(TEMPLATE_PATH)
    print(f"  {len(lexicon)} distinct class names harvested")

    print("Loading report.leader_role_training_map...")
    df = load_map()
    print(f"  {len(df):,} rows  |  {df.Leader.nunique()} leaders")

    h1_proto, li_proto = load_prototypes(TEMPLATE_PATH)

    df["JobCategoryName"] = df["JobCategoryName"].fillna("").str.strip()
    df["ClassTitle"] = df["ClassTitle"].fillna("").str.strip()
    df["JobCategoryGroup"] = df["JobCategoryGroup"].fillna("").str.strip()
    df["TrainingTrack"] = df["TrainingTrack"].fillna("").str.strip()
    df["UserCount"] = df["UserCount"].fillna(0).astype(int)
    df = df[df["JobCategoryName"] != ""]
    # Two frames: every role a leader owns (for enumeration, so roles whose only
    # track is "No Training" can still be counted as omitted), and just the rows
    # carrying a real class. "None" is the view's spelling for "no class".
    df_cls = df[(df["ClassTitle"] != "") & (df["ClassTitle"].str.lower() != "none")]

    elearning_map: dict[str, list[tuple[str, str]]] = {}
    if INCLUDE_ELEARNINGS:
        elearning_map = load_elearning_map()
        print(f"  {len(elearning_map)} tracks carry eLearning / video / material "
              f"classes ({', '.join(ELEARNING_CLASS_TYPES)})")

    # Tracks the warehouse can already resolve to at least one class — anything
    # else falls through to the older on-disk exports.
    warehouse_tracks = set(df_cls["TrainingTrack"].str.upper())
    warehouse_tracks |= set(elearning_map)
    print("  Reading older Epic exports for tracks the warehouse cannot cover...")
    supplemental = load_supplemental_class_map()
    supplemented_tracks: set[str] = set()
    supp_n_total = 0

    # Everything the warehouse resolves today, keyed like the memory file
    seen_today: dict[str, list[tuple[str, str]]] = {}
    for track, title in (df_cls[["TrainingTrack", "ClassTitle"]].drop_duplicates()
                         .itertuples(index=False)):
        seen_today.setdefault(_strip_wave_prefix(track), []).append((title, "Session"))
    for track, pairs in elearning_map.items():
        seen_today.setdefault(_strip_wave_prefix(track), []).extend(pairs)
    carried_map, n_kept, n_dropped = update_class_map(seen_today, run_date, OUT_DIR)
    carried_lines = 0
    print(f"  class memory: {n_kept} pairs carried forward (still on the Cornerstone "
          f"schedule), {n_dropped} retired -> {CLASS_MAP_PATH.name}")

    elearn_lookup = {k.strip().lower(): v for k, v in ROLE_ELEARNINGS.items()}
    derived_names: set[str] = set()

    # Carry the Change Log forward, then move the previous set out of the way.
    change_log = read_change_log(OUT_DIR / REFERENCE_NAME)
    archive_dir = archive_previous_outputs(OUT_DIR, run_date)
    if archive_dir:
        print(f"\nPrevious set moved to {archive_dir}")
    print(f"  {len(change_log)} Change Log rows carried forward")

    leader_rows, role_class_rows, not_in_docs = [], [], []

    print(f"\nWriting documents to {OUT_DIR}...")
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    for leader in CANONICAL_NAMES:
        sub = df[df["Leader"] == leader]
        if sub.empty:
            print(f"  {leader:24s} — no rows in the training map, skipped")
            leader_rows.append({
                "Leader": leader, "Name on Document": leader_display(leader),
                "Job Roles in Doc": 0, "Distinct Classes": 0, "Total Class Lines": 0,
                "Instructor-Led Lines": 0, "eLearning / Other Lines": 0,
                "Roles Not in Doc": 0, "Word Document": ""})
            continue

        # All waves merged; group ordering mirrors the reference doc's habit of
        # keeping a function's roles together, without printing group headings.
        order = (sub.groupby("JobCategoryName")
                    .agg(grp=("JobCategoryGroup", "min")).reset_index()
                    .sort_values(["grp", "JobCategoryName"]))

        sub_cls = df_cls[df_cls["Leader"] == leader]
        fname = leader_filename(leader, run_date)
        roles, all_classes = [], set()
        n_lines = n_session = n_other = supp_n = 0

        def render(title: str) -> str:
            key = _match_key(title)
            if key in lexicon:
                return lexicon[key]
            derived = _patti_case(title)
            derived_names.add(f"{title}  ->  {derived}")
            return derived

        for role_name, group in order[["JobCategoryName", "grp"]].itertuples(index=False):
            role_rows = sub[sub["JobCategoryName"] == role_name]
            tracks = sorted(set(role_rows["TrainingTrack"]) - {""})
            track_str = "; ".join(tracks)
            # A user has one wave, so per-wave headcounts add up; within a wave
            # every track/class row repeats the same count.
            users = int(role_rows.groupby("Wave")["UserCount"].max().sum())

            def omit(reason: str) -> None:
                not_in_docs.append({
                    "Leader": leader, "Job Role Group": group, "Job Role": role_name,
                    "Training Track": track_str, "Users in Role": users,
                    "Why It Is Not in the Document": reason})

            if leader in RETIRED_ROLE_LEADERS:
                reason = retired_reason(role_name)
                if reason:
                    omit(reason)
                    continue

            # (rendered class, class type, curriculum, source) in display order
            lines: list[tuple[str, str, str, str]] = []
            seen: set[str] = set()

            def add(title: str, ctype: str, curriculum: str, source: str) -> bool:
                rendered = render(title)
                if rendered in seen:
                    return False
                seen.add(rendered)
                lines.append((rendered, ctype, curriculum, source))
                return True

            cls_rows = (sub_cls[sub_cls["JobCategoryName"] == role_name]
                        [["ClassTitle", "TrainingTrack"]].drop_duplicates()
                        .sort_values(["ClassTitle", "TrainingTrack"]))
            for title, track in cls_rows.itertuples(index=False):
                add(title, "Session", track, SRC_EXPORT)

            # Fallback for tracks the current-wave export cannot resolve at all
            for track in tracks:
                if track.lower() == "no training" or track.upper() in warehouse_tracks:
                    continue
                for title, ctype in supplemental.get(_strip_wave_prefix(track), []):
                    if add(title, ctype, track, SRC_OLDER):
                        supplemented_tracks.add(track)
                        supp_n += 1

            # Sessions the export no longer shows but Cornerstone still delivers
            for track in tracks:
                for title, ctype in sorted(carried_map.get(_strip_wave_prefix(track), [])):
                    if ctype == "Session" and add(title, ctype, track, SRC_CARRIED):
                        carried_lines += 1

            # eLearnings hang off the role's tracks, including tracks that
            # carry no instructor-led session at all.
            if INCLUDE_ELEARNINGS:
                extras_by_title: dict[str, tuple[str, str, str]] = {}
                for track in tracks:
                    for title, ctype in elearning_map.get(track.upper(), []):
                        extras_by_title.setdefault(title, (ctype, track, SRC_EXPORT))
                for track in tracks:
                    for title, ctype in carried_map.get(_strip_wave_prefix(track), []):
                        if ctype != "Session":
                            extras_by_title.setdefault(title, (ctype, track, SRC_CARRIED))
                for title in sorted(extras_by_title):   # one alphabetical run
                    ctype, track, source = extras_by_title[title]
                    if add(title, ctype, track, source) and source == SRC_CARRIED:
                        carried_lines += 1


            for extra in elearn_lookup.get(role_name.lower(), []):
                if extra not in seen:
                    seen.add(extra)
                    lines.append((extra, "Online Class", "", "Manual"))

            if not lines:
                only_no_training = bool(tracks) and all(
                    t.lower() == "no training" for t in tracks)
                omit(REASON_NO_TRAINING if only_no_training else REASON_NO_CLASS)
                continue

            roles.append((role_name, [name for name, *_ in lines]))
            for name, ctype, curriculum, source in lines:
                all_classes.add(name)
                n_lines += 1
                if ctype == "Session":
                    n_session += 1
                else:
                    n_other += 1
                role_class_rows.append({
                    "Leader": leader, "Job Role Group": group, "Job Role": role_name,
                    "Training Track": track_str, "Curriculum": curriculum,
                    "Class": name, "Class Type": ctype, "Source": source,
                    "Users in Role": users, "Word Document": fname})

        # Leader-requested roles MVP does not define yet
        present = {r for r, _ in roles}
        role_group = dict(order[["JobCategoryName", "grp"]].itertuples(index=False))
        for spec in MANUAL_ROLES.get(leader, []):
            if spec["role"] in present:
                print(f"    note: {spec['role']} now comes from MVP - remove it from "
                      "MANUAL_ROLES")
                continue
            mirror = next((names for r, names in roles if r == spec["mirror"]), None)
            if mirror is None:
                print(f"    WARNING: {spec['role']} skipped - mirror role "
                      f"{spec['mirror']} is not on this document")
                continue
            group = role_group.get(spec["mirror"], "")
            roles.append((spec["role"], list(mirror)))
            role_group[spec["role"]] = group
            src_rows = [r for r in role_class_rows
                        if r["Leader"] == leader and r["Job Role"] == spec["mirror"]]
            for r in src_rows:
                role_class_rows.append({**r, "Job Role": spec["role"],
                                        "Source": SRC_MANUAL, "Users in Role": 0})
                n_lines += 1
                if r["Class Type"] == "Session":
                    n_session += 1
                else:
                    n_other += 1
            all_classes.update(mirror)
            key = (leader, spec["role"])
            if not any((row[2], row[3]) == key for row in change_log):
                change_log.append([spec["added"], "Leader request", leader, spec["role"],
                                   None, "Add role", f"Print with the {spec['mirror']} "
                                   "classes", spec["note"], "Applied", spec["added"]])
        roles.sort(key=lambda rn: (role_group.get(rn[0], ""), rn[0]))

        write_leader_doc(roles, TEMPLATE_PATH, OUT_DIR / fname, h1_proto, li_proto)
        supp_n_total += supp_n
        n_omitted = sum(1 for r in not_in_docs if r["Leader"] == leader)
        print(f"  {leader:24s} {len(roles):3d} roles, {len(all_classes):3d} classes"
              f"  ({n_omitted} not in doc, {supp_n} lines from older exports)")
        leader_rows.append({
            "Leader": leader, "Name on Document": leader_display(leader),
            "Job Roles in Doc": len(roles), "Distinct Classes": len(all_classes),
            "Total Class Lines": n_lines, "Instructor-Led Lines": n_session,
            "eLearning / Other Lines": n_other, "Roles Not in Doc": n_omitted,
            "Word Document": fname})

    ref_path = OUT_DIR / REFERENCE_NAME
    write_reference_workbook(ref_path, run_date, leader_rows, role_class_rows,
                             not_in_docs, change_log, archive_dir)

    n_docs = sum(1 for r in leader_rows if r["Job Roles in Doc"] > 0)
    print("\n=== COMPLETE ===")
    print(f"  {n_docs} documents written")
    print(f"  Output: {OUT_DIR}")
    print(f"  Reference workbook: {ref_path.name}")

    if derived_names:
        print(f"\n  {len(derived_names)} classes had no name in the reference doc "
              "and were rendered by convention:")
        for line in sorted(derived_names):
            print(f"    {line}")

    if carried_lines:
        print(f"\n  {carried_lines} bullet lines carried forward from the previous set "
              "(class still on the Cornerstone schedule, missing from today's export)")

    if supplemented_tracks:
        print(f"\n  {len(supplemented_tracks)} tracks resolved from the older "
              f"Epic exports ({supp_n_total} bullet lines):")
        for track in sorted(supplemented_tracks):
            print(f"    {track}")

    from collections import Counter
    why = Counter(r["Why It Is Not in the Document"] for r in not_in_docs)
    if why:
        print("\n  Roles not in a document:")
        for reason, n in why.most_common():
            print(f"    {n:3d}  {reason}")


if __name__ == "__main__":
    main()
