# =========================================================================
# SANITIZED TEMPLATE — org-specific values were replaced with placeholders.
# Search for: YourOrg, YourUser, Lastname01..NN, C:\path\to\...  and
# fill in your own values before running. See README.md for the full map.
# =========================================================================
"""
source_freshness.py — is what's loaded in SQL the newest Cornerstone export on disk?

The loaders always pick the newest file by modified time, but SQL only refreshes
on the auto-refresh cycle (weekdays 9:30/11:00/13:00/14:30) or a manual run —
so an export pulled between cycles sits on disk unloaded until the next one.
The scorecard and tracker call check() so that gap is always visible.

Only the two dated-filename Cornerstone feeds are checked; the Epic exports
overwrite fixed filenames, so a name comparison can't say anything about them.
"""

from __future__ import annotations

from pathlib import Path

import onedrive_paths as op

FEEDS = [
    ("Cornerstone Enterprise Training Report",
     op.RAW_CORNERSTONE_DIR, "Enterprise_Training_Report*.xlsx"),
    # NOTE 2026-08-04: Enterprise moved to a fixed overwrite filename
    # (Enterprise_Training_Report.xlsx). Once the dated copies are gone the
    # name comparison goes inert (same name every day) — mtime-based change
    # detection in auto_refresh still catches every overwrite.
    # Cornerstone Roster Report PAUSED 2026-07-27 (last-resort source —
    # manually maintained, unverified filters, ambiguous user-ID columns).
    # Re-add here + re-enable in source_registry.xlsx to resume.
]

FIX_HINT = ("run 'python sql\\refresh.py cornerstone sources' "
            "or wait for the next auto-refresh (9:30/11:00/13:00/14:30)")


def _newest_on_disk(folder: Path, pattern: str) -> Path | None:
    files = [p for p in folder.glob(pattern) if not p.name.startswith("~$")]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def check(loaded: dict[str, str]) -> list[str]:
    """loaded: {feed name: _source_file currently in SQL}. Returns one warning
    line per feed whose newest on-disk export is NOT what SQL holds."""
    warnings = []
    for feed, folder, pattern in FEEDS:
        newest = _newest_on_disk(folder, pattern)
        loaded_name = (loaded.get(feed) or "").strip()
        if newest is not None and loaded_name and newest.name != loaded_name:
            warnings.append(
                f"{feed}: newer export on disk ({newest.name}) than loaded "
                f"({loaded_name}) — {FIX_HINT}")
    return warnings
