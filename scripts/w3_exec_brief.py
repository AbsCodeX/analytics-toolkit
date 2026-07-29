# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
w3_exec_brief.py — the screenshot-ready W3 executive brief dashboard.

Read-only. Builds FOUR variants in the OneDrive Main Reports\\Dashboards\\
folder on every run — all leaders, and
with Lastname09 and/or Lastname02 excluded (EXCLUDABLE_LEADERS config below):
    RCM W3 Executive Brief.html
    RCM W3 Executive Brief - excl Lastname09.html
    RCM W3 Executive Brief - excl Lastname02.html
    RCM W3 Executive Brief - excl Lastname09 and Lastname02.html
Every number (headline, WoW deltas, donuts, funnel, trends, leader table)
recomputes for the filtered population; the subtitle names the exclusion so a
screenshot self-documents. A pill switcher sits ABOVE the white page area
(easy to crop; hidden when printing).

Refreshed by w3_tracker.py (10:00 AM task) or run any time:
    python scripts\\w3_exec_brief.py
"""

from __future__ import annotations

import subprocess
import sys
from datetime import date, datetime
from itertools import combinations
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent

# Wrong-interpreter guard (same as daily_refresh.py).
VENV_PY = ROOT / "analytics_env" / "Scripts" / "python.exe"
if VENV_PY.exists() and Path(sys.executable).resolve() != VENV_PY.resolve():
    raise SystemExit(subprocess.call(
        [str(VENV_PY), str(Path(__file__).resolve()), *sys.argv[1:]]))

import pandas as pd                      # noqa: E402

sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(ROOT / "sql"))

from refresh import CONN                 # noqa: E402
from sqlalchemy import create_engine     # noqa: E402
from w3_dashboard import BRIGHT_BLUE, RED, dual_line_chart  # noqa: E402

from onedrive_paths import DASHBOARDS_DIR  # noqa: E402
REPORTS = DASHBOARDS_DIR

# ── Manual config — EDIT when these change ──────────────────────────────────
# Go-live dates as provided to the analyst (two dates given; countdown uses the
# first). Correct here if the 11/7 vs 11/28 question resolves differently.
GO_LIVE_DATES = [date(2026, 11, 7), date(2026, 11, 28)]
CRB_TOTAL = 0
CRB_APPROVED = 0
CRB_NOT_APPROVED = 0
# Leaders that get exclusion variants (a file per combination is generated).
EXCLUDABLE_LEADERS = ["Lastname09, Firstname09", "Lastname02, Firstname02"]
# ────────────────────────────────────────────────────────────────────────────

DARK_BLUE = "#002092"
GREEN = "#16A34A"
LOW_PCT_FLAG = 50.0  # leader rows under this Reg % get the amber flag


def variant_name(excluded: tuple[str, ...]) -> str:
    if not excluded:
        return "RCM W3 Executive Brief.html"
    last = " and ".join(e.split(",")[0] for e in excluded)
    return f"RCM W3 Executive Brief - excl {last}.html"


def donut(title: str, pct: float, main_label: str, rem_label: str,
          invert: bool = False) -> str:
    """Slim executive donut (11px stroke), % centered, legend beneath."""
    r, c = 62, 2 * 3.14159265 * 62
    arc = c * max(0.0, min(100.0, pct)) / 100
    bg, fg = (BRIGHT_BLUE, DARK_BLUE) if invert else (DARK_BLUE, BRIGHT_BLUE)
    return f"""
        <div class="donut-box">
            <div class="donut-title">{title}</div>
            <svg viewBox="0 0 170 170" role="img" aria-label="{title}: {pct:.1f} percent">
                <circle cx="85" cy="85" r="{r}" fill="transparent" stroke="{bg}" stroke-width="11"/>
                <circle cx="85" cy="85" r="{r}" fill="transparent" stroke="{fg}" stroke-width="11"
                    stroke-dasharray="{arc:.1f} {c:.1f}" transform="rotate(-50 85 85)">
                    <title>{main_label}: {pct:.1f}%</title></circle>
                <text x="85" y="92" text-anchor="middle" font-size="22" font-weight="700" fill="#111">{pct:.1f}%</text>
            </svg>
            <div class="legend-row" style="margin-top:2px;">
                <span class="legend-item"><span class="legend-dot" style="background:{fg}"></span>{main_label}</span>
                <span class="legend-item"><span class="legend-dot" style="background:{bg}"></span>{rem_label}</span>
            </div>
        </div>"""


def delta_chip(delta: float, good_when_down: bool = False) -> str:
    """Small colored WoW arrow chip. Green = moving the right way."""
    if delta == 0:
        return '<span class="chip" style="color:#94a3b8;">&#8212; 0 vs last wk</span>'
    up = delta > 0
    good = (not up) if good_when_down else up
    color = GREEN if good else RED
    arrow = "&#9650;" if up else "&#9660;"
    return (f'<span class="chip" style="color:{color};">{arrow} '
            f'{"+" if up else ""}{delta:.0f} vs last wk</span>')


def render(excluded: tuple[str, ...], sc: pd.DataFrame, hist: pd.DataFrame,
           nst: pd.DataFrame, src: pd.DataFrame, nav: str, now: datetime) -> str:
    """Render one variant's full HTML for the given excluded-leader set."""
    today = now.date()
    leaders = (sc[(sc.IsTotal == 0) & (~sc.Leader.isin(excluded))]
               .sort_values("W3Users", ascending=False))
    h = hist[~hist.Leader.isin(excluded)].copy()

    # totals recomputed over the filtered population
    t = leaders[["W3Users", "FullyRegistered", "FullyTrained", "UnregisteredUsers",
                 "UnregisteredSessions", "NoShowUsers", "NoShowSessions",
                 "NoShowUsersToDate", "NoShowSessionsToDate"]].sum()
    pct_reg = round(100.0 * t.FullyRegistered / t.W3Users, 1) if t.W3Users else 0.0
    pct_trn = round(100.0 * t.FullyTrained / t.W3Users, 1) if t.W3Users else 0.0

    # week-over-week anchors: latest snapshot of the PRIOR ISO week (filtered)
    h["snap"] = pd.to_datetime(h.SnapshotDate)
    this_monday = pd.Timestamp(today - pd.Timedelta(days=today.weekday()))
    prior = h[h.snap < this_monday]
    prior_day = prior.snap.max() if not prior.empty else None
    if prior_day is not None:
        pt = prior[prior.snap == prior_day].agg(
            {"FullyRegistered": "sum", "FullyTrained": "sum",
             "UnregisteredUsers": "sum", "NoShowUsersStanding": "sum"})
        d_reg = float(t.FullyRegistered - pt.FullyRegistered)
        d_trn = float(t.FullyTrained - pt.FullyTrained)
        d_unreg = float(t.UnregisteredUsers - pt.UnregisteredUsers)
        d_ns = float(t.NoShowUsers - pt.NoShowUsersStanding)
        lead_prior = prior[prior.snap == prior_day].set_index("Leader").FullyRegistered
    else:
        d_reg = d_trn = d_unreg = d_ns = 0.0
        lead_prior = pd.Series(dtype=float)

    recovered = int(t.NoShowUsersToDate - t.NoShowUsers)
    gap = leaders.sort_values("UnregisteredUsers", ascending=False).iloc[0]
    days_to_golive = (GO_LIVE_DATES[0] - today).days
    golive_str = " &amp; ".join(f"{d:%b %d}" for d in GO_LIVE_DATES)
    excl_note = (" (excluding " + " &amp; ".join(e.split(",")[0] for e in excluded) + ")"
                 if excluded else "")

    reg_dir = ("up" if d_reg > 0 else "down" if d_reg < 0 else "flat")
    headline = (
        f"Registration stands at <strong>{pct_reg}%</strong> "
        f"({t.FullyRegistered:.0f} of {t.W3Users:.0f} users{excl_note}, {reg_dir} "
        f"{abs(d_reg):.0f} vs last week) with <strong>{days_to_golive} days</strong> to the "
        f"Nov 7 go-live. Largest gap: <strong>{gap.Leader}</strong> "
        f"({gap.UnregisteredUsers} of {t.UnregisteredUsers:.0f} unregistered users). "
        f"No-shows are controlled: <strong>{recovered} of {t.NoShowUsersToDate:.0f}</strong> "
        f"recovered, {t.NoShowUsers:.0f} standing.")

    kpis = [
        (f"{t.W3Users:.0f}", "users", "", "Wave 3 RCM Users" + excl_note.replace("&amp;", "&")),
        (f"{t.FullyRegistered:.0f}", f"users ({pct_reg}%)", delta_chip(d_reg), "Fully Registered"),
        (f"{t.FullyTrained:.0f}", f"users ({pct_trn}%)", delta_chip(d_trn), "Fully Trained"),
        (f"{t.UnregisteredUsers:.0f}", "users",
         delta_chip(d_unreg, good_when_down=True),
         f"Unregistered &mdash; {t.UnregisteredSessions:.0f} open sessions"),
        (f"{t.NoShowUsers:.0f}", "users",
         delta_chip(d_ns, good_when_down=True),
         f"Standing No-Shows &mdash; {t.NoShowSessions:.0f} sessions"),
        (f"{CRB_TOTAL}", "requests", "",
         f"Curriculum Review Board &mdash; {CRB_APPROVED} approved / {CRB_NOT_APPROVED} not"),
        (f"{days_to_golive}", "days", "", f"To W3 Go-Live &mdash; {golive_str}"),
    ]
    kpi_cards = "\n".join(
        f'<div class="kpi-card"><div><span class="kpi-value">{v}</span> '
        f'<span class="kpi-sub">{s}</span></div>{chip}'
        f'<div class="kpi-label">{lab}</div></div>'
        for v, s, chip, lab in kpis)

    funnel = f"""
        <div class="funnel-box">
            <div class="donut-title">No-Show Recovery</div>
            <div class="funnel-row">
                <div class="funnel-step"><div class="funnel-val">{t.NoShowUsersToDate:.0f}</div>
                    <div class="funnel-lab">users no-showed<br>to date</div></div>
                <div class="funnel-arrow">&#10132;</div>
                <div class="funnel-step" style="background:#e8f7ee;"><div class="funnel-val" style="color:{GREEN};">{recovered}</div>
                    <div class="funnel-lab">recovered<br>(re-registered / completed)</div></div>
                <div class="funnel-arrow">&#10132;</div>
                <div class="funnel-step" style="background:#fdeaea;"><div class="funnel-val" style="color:{RED};">{t.NoShowUsers:.0f}</div>
                    <div class="funnel-lab">still standing<br>({t.NoShowSessions:.0f} sessions)</div></div>
            </div>
        </div>"""

    # ── trend panels (daily; filtered from leader-level history) ──
    def dlab(s):
        d = pd.to_datetime(s)
        return f"{d.month}/{d.day}"

    daily = h.groupby("SnapshotDate", as_index=False).agg(
        Reg=("FullyRegistered", "sum"), Unreg=("UnregisteredUsers", "sum"),
        NS=("NoShowUsersStanding", "sum"))
    reg_d = {dlab(d): float(v) for d, v in zip(daily.SnapshotDate, daily.Reg)}
    unreg_d = {dlab(d): float(v) for d, v in zip(daily.SnapshotDate, daily.Unreg)}
    # no-show series: the 7/13-based reconstruction is org-wide for our
    # leaders; keep it whenever the excluded leaders have zero no-shows ever
    # (exact), else fall back to filtered leader-level history (starts 7/15).
    excl_ns = float(sc[(sc.IsTotal == 0) & (sc.Leader.isin(excluded))]
                    .NoShowUsersToDate.sum()) if excluded else 0.0
    if excl_ns == 0:
        ns_d = {dlab(d): float(v) for d, v in zip(nst.AsOfDate, nst.NoShowUsersStanding)}
    else:
        ns_d = {dlab(d): float(v) for d, v in zip(daily.SnapshotDate, daily.NS)}

    reg_dates = sorted(reg_d, key=lambda s: tuple(int(p) for p in s.split("/")))
    mix_dates = sorted(set(unreg_d) | set(ns_d),
                       key=lambda s: tuple(int(p) for p in s.split("/")))
    trend_reg = dual_line_chart(reg_dates, [("Registered Users", DARK_BLUE, reg_d)],
                                aria="Registered users daily trend")
    trend_mix = dual_line_chart(mix_dates, [
        ("Unregistered Users", BRIGHT_BLUE, unreg_d),
        ("Standing No-Show Users", RED, ns_d),
    ], aria="Unregistered and no-show users daily trend")

    rows = []
    for _, r in leaders.iterrows():
        pct = float(r.PctRegistered)
        dv = float(r.FullyRegistered) - float(lead_prior.get(r.Leader, r.FullyRegistered))
        dtxt = ('<span style="color:#94a3b8;">&#8212;</span>' if dv == 0 else
                f'<span style="color:{GREEN if dv > 0 else RED};font-weight:600;">'
                f'{"&#9650;" if dv > 0 else "&#9660;"} {dv:+.0f}</span>')
        flag = (' <span class="flag" title="Registration below '
                f'{LOW_PCT_FLAG:.0f}%">&#9888;</span>' if pct < LOW_PCT_FLAG else "")
        rows.append(
            f'<tr><td>{r.Leader}{flag}</td><td>{r.W3Users}</td><td>{r.FullyRegistered}</td>'
            f'<td class="pctcell"><span class="pctbar" style="width:{pct:.0f}%;"></span>'
            f'<span class="pcttext">{pct:.1f}%</span></td>'
            f'<td>{dtxt}</td><td>{r.FullyTrained}</td>'
            f'<td>{r.UnregisteredUsers}</td><td>{r.NoShowUsers}</td></tr>')
    leader_rows = "\n".join(rows)
    total_row = (f'<tr class="ldr-total"><td>TOTAL{excl_note.replace("&amp;", "&")}</td>'
                 f'<td>{t.W3Users:.0f}</td><td>{t.FullyRegistered:.0f}</td>'
                 f'<td class="pctcell"><span class="pcttext">{pct_reg}%</span></td>'
                 f'<td>{delta_chip(d_reg)}</td>'
                 f'<td>{t.FullyTrained:.0f}</td><td>{t.UnregisteredUsers:.0f}</td>'
                 f'<td>{t.NoShowUsers:.0f}</td></tr>')
    src_line = " &nbsp;&bull;&nbsp; ".join(
        f"{r.Feed}: {r.F} (loaded {str(r.T)[:16]})" for _, r in src.iterrows())

    subtitle = f"Wave 3 Executive Brief{excl_note} &mdash; {now:%A, %B %d, %Y %I:%M %p}"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>RCM W3 Executive Brief{excl_note.replace('&amp;', '&')}</title>
<style>
    :root {{
        --brand-light-blue: #A5D8FF;
        --brand-dark-blue: #002092;
        --brand-bright-blue: #3399FF;
        --brand-orange: #ED7D31;
        --brand-green: #16A34A;
        --text-main: #333333;
        --text-muted: #666666;
    }}
    *, *::before, *::after {{ box-sizing: border-box; }}
    body {{ background: #eef2f7; font-family: 'Segoe UI', sans-serif; color: var(--text-main); margin: 0; }}
    .variant-nav {{ max-width: 1560px; margin: 0 auto; padding: 8px 26px 6px; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }}
    .variant-nav .nav-label {{ font-size: 0.72rem; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; }}
    .variant-nav a {{ font-size: 0.78rem; text-decoration: none; color: var(--brand-dark-blue); background: white; border: 1px solid #cbd5e1; border-radius: 999px; padding: 3px 12px; }}
    .variant-nav a.active {{ background: var(--brand-dark-blue); color: white; border-color: var(--brand-dark-blue); font-weight: 700; }}
    .page {{ max-width: 1560px; margin: 0 auto; padding: 16px 26px 20px; background: white; }}
    header {{ text-align: center; margin-bottom: 10px; }}
    header h1 {{ font-size: 1.45rem; font-weight: 700; color: #1f2937; margin: 0; }}
    header p {{ font-style: italic; color: var(--text-muted); font-size: 0.92rem; margin: 2px 0 0; }}
    .headline {{ background: #eff6ff; border: 1px solid #bfdbfe; border-left: 5px solid var(--brand-dark-blue);
        border-radius: 4px; padding: 10px 16px; font-size: 0.98rem; line-height: 1.5; margin-bottom: 14px; }}
    .kpi-row {{ display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }}
    .kpi-card {{ background: var(--brand-light-blue); border-radius: 4px; padding: 9px 10px; height: 92px; flex: 1 1 140px; display: flex; flex-direction: column; justify-content: space-between; }}
    .kpi-value {{ font-size: 1.55rem; font-weight: 500; line-height: 1; }}
    .kpi-sub {{ font-size: 0.76rem; font-weight: 500; color: var(--text-main); white-space: nowrap; }}
    .chip {{ font-size: 0.74rem; font-weight: 700; }}
    .kpi-label {{ font-size: 0.72rem; font-weight: 600; color: var(--text-muted); line-height: 1.15; }}
    .mid-row {{ display: flex; gap: 20px; flex-wrap: wrap; align-items: stretch; justify-content: space-between; margin-bottom: 12px; }}
    .donut-box {{ flex: 0 1 240px; text-align: center; }}
    .donut-box svg {{ width: 148px; height: auto; }}
    .donut-title {{ font-weight: 600; font-size: 0.92rem; margin-bottom: 4px; }}
    .funnel-box {{ flex: 1 1 460px; text-align: center; align-self: center; }}
    .funnel-row {{ display: flex; align-items: center; justify-content: center; gap: 10px; margin-top: 12px; }}
    .funnel-step {{ background: #f1f5f9; border-radius: 6px; padding: 12px 18px; min-width: 128px; }}
    .funnel-val {{ font-size: 1.7rem; font-weight: 700; line-height: 1.1; }}
    .funnel-lab {{ font-size: 0.72rem; color: var(--text-muted); line-height: 1.25; margin-top: 3px; }}
    .funnel-arrow {{ font-size: 1.3rem; color: #94a3b8; }}
    .legend-row {{ display: flex; justify-content: center; gap: 12px; flex-wrap: wrap; margin-top: 4px; }}
    .legend-item {{ display: flex; align-items: center; font-size: 0.72rem; color: var(--text-muted); }}
    .legend-dot {{ width: 10px; height: 10px; border-radius: 50%; margin-right: 4px; flex-shrink: 0; display: inline-block; }}
    .trend-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 22px; margin-bottom: 12px; }}
    .panel-title {{ font-weight: 600; font-size: 0.92rem; text-align: center; margin-bottom: 4px; }}
    .trend-panel svg {{ width: 100%; height: auto; }}
    .ldr-table {{ width: 100%; border-collapse: collapse; font-size: 0.8rem; }}
    .ldr-table th {{ background: var(--brand-dark-blue); color: #fff; font-weight: 600; padding: 5px 10px; text-align: right; white-space: nowrap; }}
    .ldr-table th:first-child {{ text-align: left; }}
    .ldr-table td {{ padding: 4px 10px; border-bottom: 1px solid #e2e8f0; text-align: right; }}
    .ldr-table td:first-child {{ text-align: left; font-weight: 600; }}
    .ldr-table tbody tr:nth-child(even) {{ background: #f8fafc; }}
    .ldr-table tr.ldr-total {{ background: var(--brand-light-blue); font-weight: 700; }}
    .ldr-table tr.ldr-total td {{ border-bottom: none; }}
    .pctcell {{ position: relative; min-width: 120px; }}
    .pctbar {{ position: absolute; left: 4px; top: 4px; bottom: 4px; background: var(--brand-light-blue); border-radius: 3px; z-index: 0; max-width: calc(100% - 8px); }}
    .pcttext {{ position: relative; z-index: 1; }}
    .flag {{ color: var(--brand-orange); }}
    footer {{ margin-top: 10px; text-align: center; font-size: 0.65rem; color: #94a3b8; line-height: 1.6; }}
    @media (max-width: 1000px) {{ .trend-grid {{ grid-template-columns: 1fr; }} }}
    @media print {{ .variant-nav {{ display: none; }} .page {{ padding: 6px; }} body {{ zoom: 0.8; background: white; }} }}
</style>
</head>
<body>

<nav class="variant-nav">{nav}</nav>

<div class="page">

    <header>
        <h1>Revenue Cycle Training Operations - Executive Summary</h1>
        <p>{subtitle}</p>
    </header>

    <div class="headline">{headline}</div>

    <div class="kpi-row">
{kpi_cards}
    </div>

    <div class="mid-row">
        {donut("% Registration Completed", pct_reg, "Registered", "Unregistered")}
        {donut("% Training Completed", pct_trn, "Completed", "Not Completed", invert=True)}
        {funnel}
    </div>

    <div class="trend-grid">
        <div>
            <div class="panel-title">Registered Users &mdash; daily</div>
            {trend_reg}
        </div>
        <div>
            <div class="panel-title">Unregistered vs Standing No-Show Users &mdash; daily</div>
            {trend_mix}
        </div>
    </div>

    <table class="ldr-table">
        <thead><tr><th>Leader</th><th>W3 Users</th><th>Registered</th><th>Reg %</th>
            <th>&Delta; wk</th><th>Trained</th><th>Unregistered</th><th>Standing No-Shows</th></tr></thead>
        <tbody>
{leader_rows}
{total_row}
        </tbody>
    </table>

    <footer>
        Registered/Trained per Epic curriculum flags; no-show history per Cornerstone transcripts (W3 classes began 7/13).
        Go-live dates as provided: {golive_str}, 2026 &mdash; countdown uses the first. &#9888; = registration below {LOW_PCT_FLAG:.0f}%.<br>
        {src_line}<br>
        Auto-generated by scripts\\w3_exec_brief.py &mdash; rerun to refresh.
    </footer>

</div>
</body>
</html>
"""


def main() -> None:
    engine = create_engine(CONN)
    sc = pd.read_sql("SELECT * FROM report.w3_scorecard", engine)
    hist = pd.read_sql(
        "SELECT SnapshotDate, Leader, W3Users, FullyRegistered, FullyTrained, "
        "UnregisteredUsers, NoShowUsersStanding "
        "FROM history.w3_leader_daily ORDER BY SnapshotDate", engine)
    nst = pd.read_sql("SELECT * FROM report.w3_noshow_trend ORDER BY AsOfDate", engine)
    src = pd.read_sql(
        "SELECT 'Cornerstone Enterprise' AS Feed, MAX(_source_file) F, MAX(_loaded_at) T FROM raw.cornerstone "
        "UNION ALL SELECT 'Epic Curriculum Status', MAX(_source_file), MAX(_loaded_at) FROM raw.epic_status "
        "UNION ALL SELECT 'Master Wave File', MAX(_source_file), MAX(_loaded_at) FROM raw.master", engine)
    engine.dispose()

    now = datetime.now()
    combos: list[tuple[str, ...]] = [()]
    for k in range(1, len(EXCLUDABLE_LEADERS) + 1):
        combos.extend(combinations(EXCLUDABLE_LEADERS, k))

    def label(exc: tuple[str, ...]) -> str:
        return ("All Leaders" if not exc
                else "Excl. " + " &amp; ".join(e.split(",")[0] for e in exc))

    for excluded in combos:
        nav = '<span class="nav-label">View:</span>' + "".join(
            f'<a href="{variant_name(c)}"{" class=" + chr(34) + "active" + chr(34) if c == excluded else ""}>'
            f'{label(c)}</a>' for c in combos)
        out = REPORTS / variant_name(excluded)
        out.write_text(render(excluded, sc, hist, nst, src, nav, now), encoding="utf-8")
        print(f"written: {out.name}")


if __name__ == "__main__":
    main()
