# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
w3_dashboard.py — generate the auto-populated W3 executive scorecard dashboard.

Read-only. Rebuilds Main Reports\\Dashboards\\RCM W3 Scorecard Dashboard.html from SQL on every
run: headline KPI cards, the by-leader table, and two daily trend line charts
(unregistered users, standing no-show sessions). Same brand theme as the
executive dashboards (dark blue #002092 / bright blue #3399FF / light blue
#A5D8FF / orange #ED7D31, Segoe UI). Self-contained file — no CDN, no JS
frameworks — so it opens anywhere, offline.

Runs standalone (python scripts\\w3_dashboard.py) and is also refreshed by
w3_tracker.py, so the 10:00 AM scheduled task keeps it current daily.
"""

from __future__ import annotations

import subprocess
import sys
from datetime import datetime
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

from onedrive_paths import DASHBOARDS_DIR  # noqa: E402
from refresh import CONN                 # noqa: E402
from sqlalchemy import create_engine     # noqa: E402

OUT = DASHBOARDS_DIR / "RCM W3 Scorecard Dashboard.html"

BRIGHT_BLUE = "#3399FF"
RED = "#DC2626"


def dual_line_chart(dates: list[str], series: list[tuple[str, str, dict]],
                    aria: str = "Unregistered and no-show users trend",
                    labels: str = "all") -> str:
    """Two-series SVG line chart on ONE shared y-axis (user counts).
    Polished spec: dashed recessive grid, 2.5px lines with a soft area tint,
    white-ringed markers with hover tooltips, value labels with a white halo
    so they never collide with the lines, legend beneath.
    series = [(name, color, {date_label: value})].
    labels = "all" (a value on every point) or "ends" (first and last point of
    each series only — for two dense series the per-point labels collide;
    hover tooltips still cover every point)."""
    w, h = 960, 290
    ml, mr, mt, mb = 52, 28, 24, 40
    pw, ph = w - ml - mr, h - mt - mb
    allvals = [v for _, _, d in series for v in d.values()]
    y0, y1 = 0, max(allvals) * 1.2
    n = len(dates)
    base = mt + ph

    def x(i):
        return ml + (pw * i / max(n - 1, 1)) if n > 1 else ml + pw / 2

    def y(v):
        return mt + ph - ph * (v - y0) / (y1 - y0)

    grid = []
    for g in range(4):
        gv = y0 + (y1 - y0) * (g + 1) / 5
        gy = y(gv)
        grid.append(f'<line x1="{ml}" y1="{gy:.1f}" x2="{w - mr}" y2="{gy:.1f}" '
                    f'stroke="#e5eaf1" stroke-width="1" stroke-dasharray="3 4"/>')
        grid.append(f'<text x="{ml - 10}" y="{gy + 4:.1f}" text-anchor="end" '
                    f'font-size="11" fill="#94a3b8">{gv:.0f}</text>')

    xlabels = [f'<text x="{x(i):.1f}" y="{h - 14}" text-anchor="middle" '
               f'font-size="11.5" fill="#666666">{d}</text>' for i, d in enumerate(dates)]

    areas, lines, marks, vlabels = [], [], [], []
    for name, color, data in series:
        idx = [(i, data[d]) for i, d in enumerate(dates) if d in data]
        if not idx:
            continue
        pts = " ".join(f"{x(i):.1f},{y(v):.1f}" for i, v in idx)
        # soft area tint under the line, anchored to the baseline
        areas.append(
            f'<polygon points="{x(idx[0][0]):.1f},{base:.1f} {pts} '
            f'{x(idx[-1][0]):.1f},{base:.1f}" fill="{color}" fill-opacity="0.06"/>')
        lines.append(f'<polyline points="{pts}" fill="none" stroke="{color}" '
                     f'stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>')
        label_at = {idx[0][0], idx[-1][0]} if labels == "ends" else {i for i, _ in idx}
        for i, v in idx:
            cx, cy = x(i), y(v)
            tip = f'<title>{name} — {dates[i]}: {v:.0f} users</title>'
            marks.append(
                f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="11" fill="transparent">{tip}</circle>'
                f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="4" fill="{color}" '
                f'stroke="#ffffff" stroke-width="2">{tip}</circle>')
            if i not in label_at:
                continue
            # value label with a white halo (paint-order) so it reads over lines
            vlabels.append(
                f'<text x="{cx:.1f}" y="{cy - 11:.1f}" text-anchor="middle" font-size="12.5" '
                f'font-weight="600" fill="#333333" stroke="#ffffff" stroke-width="3.5" '
                f'paint-order="stroke" stroke-linejoin="round">{v:.0f}</text>')

    legend = "".join(
        f'<span class="legend-item"><span class="legend-dot" '
        f'style="background:{color}"></span>{name}</span>'
        for name, color, _ in series)

    return (f'<div class="trend-panel">'
            f'<svg viewBox="0 0 {w} {h}" role="img" aria-label="{aria}">'
            f'{"".join(grid)}'
            f'<line x1="{ml}" y1="{base}" x2="{w - mr}" y2="{base}" stroke="#cbd5e1" stroke-width="1"/>'
            f'{"".join(areas)}{"".join(lines)}{"".join(marks)}{"".join(vlabels)}{"".join(xlabels)}</svg>'
            f'<div class="legend-row">{legend}</div></div>')


def main() -> None:
    engine = create_engine(CONN)
    sc = pd.read_sql("SELECT * FROM report.w3_scorecard", engine)
    hist = pd.read_sql(
        "SELECT SnapshotDate, SUM(W3Users) W3Users, SUM(FullyRegistered) Reg, "
        "SUM(FullyTrained) Trn, SUM(UnregisteredUsers) UnregUsers "
        "FROM history.w3_leader_daily GROUP BY SnapshotDate ORDER BY SnapshotDate", engine)
    nst = pd.read_sql("SELECT * FROM report.w3_noshow_trend ORDER BY AsOfDate", engine)
    src = pd.read_sql(
        "SELECT 'Cornerstone Enterprise' AS Feed, MAX(_source_file) F, MAX(_loaded_at) T FROM raw.cornerstone "
        "UNION ALL SELECT 'Epic Curriculum Status', MAX(_source_file), MAX(_loaded_at) FROM raw.epic_status "
        "UNION ALL SELECT 'Master Wave File', MAX(_source_file), MAX(_loaded_at) FROM raw.master", engine)
    engine.dispose()

    total = sc[sc.IsTotal == 1].iloc[0]
    leaders = sc[sc.IsTotal == 0].sort_values("W3Users", ascending=False)
    now = datetime.now()

    def dlab(s):
        d = pd.to_datetime(s)
        return f"{d.month}/{d.day}"

    unreg = {dlab(d): float(v) for d, v in zip(hist.SnapshotDate, hist.UnregUsers)}
    noshow = {dlab(d): float(v) for d, v in zip(nst.AsOfDate, nst.NoShowUsersStanding)}
    dates = sorted(set(unreg) | set(noshow),
                   key=lambda s: tuple(int(p) for p in s.split("/")))
    chart = dual_line_chart(dates, [
        ("Unregistered Users", BRIGHT_BLUE, unreg),
        ("Standing No-Show Users", RED, noshow),
    ])

    leader_rows = "\n".join(
        f'<tr><td>{r.Leader}</td><td>{r.W3Users}</td><td>{r.FullyRegistered}</td>'
        f'<td>{r.PctRegistered}%</td><td>{r.FullyTrained}</td>'
        f'<td>{r.UnregisteredUsers}</td><td>{r.NoShowUsers}</td></tr>'
        for _, r in leaders.iterrows())
    total_row = (f'<tr class="ldr-total"><td>TOTAL</td><td>{total.W3Users}</td>'
                 f'<td>{total.FullyRegistered}</td><td>{total.PctRegistered}%</td>'
                 f'<td>{total.FullyTrained}</td><td>{total.UnregisteredUsers}</td>'
                 f'<td>{total.NoShowUsers}</td></tr>')
    src_rows = "\n".join(
        f'<tr><td>{r.Feed}</td><td>{r.F}</td><td>{str(r.T)[:16]}</td></tr>'
        for _, r in src.iterrows())

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>RCM W3 Scorecard Dashboard</title>
<style>
    :root {{
        --brand-light-blue: #A5D8FF;
        --brand-dark-blue: #002092;
        --brand-bright-blue: #3399FF;
        --brand-orange: #ED7D31;
        --text-main: #333333;
        --text-muted: #666666;
    }}
    *, *::before, *::after {{ box-sizing: border-box; }}
    body {{ background: #fcfcfc; font-family: 'Segoe UI', sans-serif; color: var(--text-main); margin: 0; }}
    .page {{ max-width: 1180px; margin: 0 auto; padding: 28px 32px 40px; background: white; min-height: 100vh; }}
    header {{ text-align: center; margin-bottom: 24px; }}
    header h1 {{ font-size: 1.7rem; font-weight: 700; color: #1f2937; margin: 0; }}
    header p {{ font-style: italic; color: var(--text-muted); font-size: 1.05rem; margin: 6px 0 0; }}
    .main-kpi {{ background: var(--brand-light-blue); border-radius: 4px; padding: 12px 16px; height: 110px; width: 260px; display: flex; flex-direction: column; justify-content: space-between; margin-bottom: 26px; }}
    .main-kpi .val {{ font-size: 2.9rem; font-weight: 500; line-height: 1; }}
    .kpi-row {{ display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 34px; }}
    .kpi-card {{ background: var(--brand-light-blue); border-radius: 4px; padding: 12px 10px; height: 100px; flex: 1 1 150px; display: flex; flex-direction: column; justify-content: space-between; }}
    .kpi-value {{ font-size: 2.0rem; font-weight: 500; line-height: 1; }}
    .kpi-sub {{ font-size: 0.95rem; font-weight: 500; color: var(--text-main); }}
    .kpi-label {{ font-size: 0.82rem; font-weight: 600; color: var(--text-muted); line-height: 1.2; }}
    .section-title {{ font-weight: 600; font-size: 1.1rem; margin: 34px 0 14px; text-align: center; }}
    .trend-row {{ display: flex; gap: 24px; flex-wrap: wrap; }}
    .trend-panel {{ flex: 1 1 100%; }}
    .trend-panel svg {{ width: 100%; height: auto; }}
    .legend-row {{ display: flex; justify-content: center; gap: 18px; margin-top: 4px; }}
    .legend-item {{ display: flex; align-items: center; font-size: 0.85rem; color: var(--text-muted); }}
    .legend-dot {{ width: 12px; height: 12px; border-radius: 50%; margin-right: 5px; flex-shrink: 0; display: inline-block; }}
    .ldr-table {{ width: 100%; border-collapse: collapse; font-size: 0.9rem; }}
    .ldr-table th {{ background: var(--brand-dark-blue); color: #fff; font-weight: 600; padding: 8px 12px; text-align: right; white-space: nowrap; }}
    .ldr-table th:first-child {{ text-align: left; }}
    .ldr-table td {{ padding: 7px 12px; border-bottom: 1px solid #e2e8f0; text-align: right; }}
    .ldr-table td:first-child {{ text-align: left; font-weight: 600; }}
    .ldr-table tbody tr:nth-child(even) {{ background: #f8fafc; }}
    .ldr-table tr.ldr-total {{ background: var(--brand-light-blue); font-weight: 700; }}
    .ldr-table tr.ldr-total td {{ border-bottom: none; }}
    .src-table {{ border-collapse: collapse; font-size: 0.75rem; color: var(--text-muted); margin: 10px auto 0; }}
    .src-table td {{ padding: 2px 14px 2px 0; }}
    footer {{ margin-top: 36px; text-align: center; font-size: 0.72rem; color: #94a3b8; }}
    @media print {{ .page {{ padding: 10px; }} }}
</style>
</head>
<body>
<div class="page">

    <header>
        <h1>Revenue Cycle Training Operations - Executive Summary</h1>
        <p>Wave 3 Scorecard &mdash; {now:%A, %B %d, %Y %I:%M %p}</p>
    </header>

    <div class="main-kpi">
        <div class="val">{total.W3Users}</div>
        <div class="kpi-label" style="font-size:1rem;">Wave 3 Revenue Cycle Employees</div>
    </div>

    <div class="kpi-row">
        <div class="kpi-card"><div><span class="kpi-value">{total.FullyRegistered}</span>
            <span class="kpi-sub">users ({total.PctRegistered}%)</span></div>
            <div class="kpi-label">Fully Registered</div></div>
        <div class="kpi-card"><div><span class="kpi-value">{total.FullyTrained}</span>
            <span class="kpi-sub">users ({total.PctTrained}%)</span></div>
            <div class="kpi-label">Fully Trained</div></div>
        <div class="kpi-card"><div><span class="kpi-value">{total.UnregisteredUsers}</span>
            <span class="kpi-sub">users</span></div>
            <div class="kpi-label">Unregistered &mdash; {total.UnregisteredSessions} open class sessions</div></div>
        <div class="kpi-card"><div><span class="kpi-value">{total.NoShowUsers}</span>
            <span class="kpi-sub">users</span></div>
            <div class="kpi-label">Standing No-Shows &mdash; {total.NoShowSessions} missed sessions unresolved</div></div>
        <div class="kpi-card"><div><span class="kpi-value">{total.NoShowUsersToDate}</span>
            <span class="kpi-sub">users</span></div>
            <div class="kpi-label">No-Shows to Date &mdash; {total.NoShowSessionsToDate} no-show records</div></div>
    </div>

    <div class="section-title">Daily Trend &mdash; Unregistered vs No-Show Users (since Wave 3 classes began 7/13)</div>
    <div class="trend-row">
        {chart}
    </div>

    <div class="section-title">Wave 3 &mdash; By Leader Breakdown</div>
    <div style="overflow-x:auto;">
        <table class="ldr-table">
            <thead><tr><th>Leader</th><th>W3 Users</th><th>Registered</th><th>Reg %</th>
                <th>Trained</th><th>Unregistered</th><th>Standing No-Shows</th></tr></thead>
            <tbody>
{leader_rows}
{total_row}
            </tbody>
        </table>
    </div>

    <footer>
        Data sources for this run (export file &rarr; loaded into SQL)
        <table class="src-table">
{src_rows}
        </table>
        Auto-generated by scripts\\w3_dashboard.py &mdash; rerun to refresh. Leader-scoped, in-scope Wave 3 population.
    </footer>

</div>
</body>
</html>
"""
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(html, encoding="utf-8")
    print(f"Dashboard written: {OUT}")


if __name__ == "__main__":
    main()
