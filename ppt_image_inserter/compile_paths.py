"""
Resolve the latest dated compile directory.

Compiles land in folders whose names end in an 8-digit date suffix, e.g.
``CTL_nuc_MT_granules_3v12min_bothdays_ncdn05_20260724``. When the pipeline
rebuilds them (renamed suffix, new date), scripts that hard-code the folder
name silently point at a stale compile.

``resolve_latest_compile`` finds the newest match under a parent directory by
the trailing ``_YYYYMMDD`` alone — glob what you want to match, and the newest
date wins.

Usage::

    from ppt_image_inserter import resolve_latest_compile

    ROOT = resolve_latest_compile(
        "L:/FF/Nucleus_granules/CTL_fixed/compiled_results",
        "CTL_nuc_MT_granules_3v12min_*",
    )
    # -> Path(".../CTL_nuc_MT_granules_3v12min_reposCent_20260725")
"""

import fnmatch
import os
import re
from pathlib import Path


_DATE_TAIL = re.compile(r"_(\d{8})$")


def _iter_dated_children(parent):
    """Yield (Path, date_str) for each subdir of *parent* ending in _YYYYMMDD.

    Uses os.listdir (not glob) so Windows long-path parents work.
    """
    parent = Path(parent)
    try:
        names = os.listdir(str(parent))
    except (FileNotFoundError, NotADirectoryError):
        return
    for name in names:
        m = _DATE_TAIL.search(name)
        if not m:
            continue
        p = parent / name
        if p.is_dir():
            yield p, m.group(1)


def resolve_latest_compile(parent, glob="*", *, must_exist=True):
    """Return the newest dated subdir of *parent* matching *glob*.

    Only subdirectories whose names end in ``_YYYYMMDD`` (8 digits) are
    considered. The one with the largest date wins; ties break on name.

    Parameters
    ----------
    parent : str or Path
        Directory to look inside.
    glob : str
        fnmatch-style pattern the folder name must match (default ``"*"``).
        Matches the *folder name*, not the full path. The date suffix does
        not need to be in the pattern — every candidate already ends in
        ``_YYYYMMDD``.
    must_exist : bool
        If True (default), raise ``FileNotFoundError`` when nothing matches.
        If False, return ``None`` instead.

    Returns
    -------
    pathlib.Path
        Absolute path to the newest matching compile directory.
    """
    candidates = [
        (date, p) for p, date in _iter_dated_children(parent)
        if fnmatch.fnmatchcase(p.name, glob)
    ]
    if not candidates:
        if must_exist:
            raise FileNotFoundError(
                "No dated compile matching {!r} under {}".format(glob, parent)
            )
        return None
    candidates.sort(key=lambda dp: (dp[0], dp[1].name))
    return candidates[-1][1].resolve()


def compile_date_from(path):
    """Extract the trailing ``YYYYMMDD`` date from a compile directory name.

    Returns ``"YYYY-MM-DD"`` for use in slide subtitles/footers. Raises
    ``ValueError`` if the name has no 8-digit date suffix.
    """
    name = Path(path).name
    m = _DATE_TAIL.search(name)
    if not m:
        raise ValueError("No trailing _YYYYMMDD in {!r}".format(name))
    d = m.group(1)
    return "{}-{}-{}".format(d[:4], d[4:6], d[6:8])
