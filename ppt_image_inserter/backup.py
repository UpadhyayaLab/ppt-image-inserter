"""
Backup system for PowerPoint presentations.
"""

import os
import re
import shutil
import hashlib
import datetime
from pathlib import Path
from typing import Dict, Optional


def subset_output_path(output_path, **subset) -> Path:
    """Output path for a SUBSET run (--cells / --since / --fovs / --only ...) when no
    --output was given: '<stem> (subset <name> <value>; ...)<suffix>' beside the full deck.

    Subset runs used to write the full deck's default path, so a quick '--cells 1,2'
    check replaced the whole deck (bug hunt 2026-10-24). Returns output_path unchanged
    when every subset value is empty.
    """
    output_path = Path(output_path)
    parts = [f"{k} {v}" for k, v in subset.items() if v not in (None, "", [], (), set())]
    if not parts:
        return output_path
    full = "; ".join(parts)
    tag = re.sub(r"[^A-Za-z0-9,;._ -]+", "-", full)
    if len(tag) > 60:
        # A plain cut made '--cells 1..39' and '--cells 1..40' one file; keep a hash of the
        # whole subset so different subsets never share a name (bug hunt 2026-10-24, second pass).
        tag = f"{tag[:51].rstrip()}~{hashlib.sha1(full.encode('utf-8')).hexdigest()[:8]}"
    tag = tag.strip()
    out = output_path.with_name(f"{output_path.stem} (subset {tag}){output_path.suffix}")
    print(f"Subset run -> {out} (the full deck {output_path.name} is left as is)")
    return out


def backup_presentation(ppt_path: str, backup_base: Optional[str] = None) -> Dict[str, str]:
    """
    Create backups in multiple time-interval categories (one per category).

    This function implements a smart backup strategy that maintains ONE backup
    per presentation in each time interval category. When a backup is created,
    it overwrites the existing backup for that presentation in that category.
    Interval backups are only created if the time threshold has been exceeded
    for that presentation.

    Args:
        ppt_path (str): Path to the PowerPoint file to backup
        backup_base (str, optional): Base directory for backups. Default: a 'backups' folder
            beside the presentation (was the relative 'PPT/backups', i.e. wherever the
            script happened to run, often the repo).

    Returns:
        dict: Dictionary mapping category names to backup file paths for backups created
              Example: {'latest': '<deck folder>/backups/latest/file.pptx'}

    Categories and time intervals:
        - 'latest': Always creates a backup (0 seconds threshold)
        - 'previous': The 'latest' copy it replaces (so two quick reruns still leave the
          deck as it was before the first one)
        - '5min': Creates backup if >5 minutes since last backup in this category
        - '10min': Creates backup if >10 minutes since last backup in this category
        - '30min': Creates backup if >30 minutes since last backup in this category
        - 'hourly': Creates backup if >1 hour since last backup in this category

    Directory structure:
        <deck folder>/backups/
        ├── latest/
        ├── previous/
        ├── 5min/
        ├── 10min/
        ├── 30min/
        └── hourly/

    Example:
        >>> backups = backup_presentation('presentation.pptx')
        >>> print(f"Created backups: {list(backups.keys())}")
        Created backups: ['latest', '5min', '10min', '30min', 'hourly']
    """
    # Validate input file
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")
    if backup_base is None:
        backup_base = os.path.join(os.path.dirname(os.path.abspath(ppt_path)), 'backups')

    # Time intervals in seconds
    intervals = {
        'latest': 0,      # Always backup
        '5min': 300,      # 5 minutes = 300 seconds
        '10min': 600,     # 10 minutes = 600 seconds
        '30min': 1800,    # 30 minutes = 1800 seconds
        'hourly': 3600    # 1 hour = 3600 seconds
    }

    created_backups = {}
    timestamp = datetime.datetime.now()
    filename = os.path.basename(ppt_path)

    for category, threshold_seconds in intervals.items():
        # Create category directory
        category_dir = os.path.join(backup_base, category)
        os.makedirs(category_dir, exist_ok=True)
        backup_path = os.path.join(category_dir, filename)

        # Check the existing backup for this presentation in this category
        existing_backups = [backup_path] if os.path.exists(backup_path) else []
        should_backup = True

        if existing_backups and threshold_seconds > 0:
            # Get most recent backup by modification time
            latest_backup = max(existing_backups, key=os.path.getmtime)
            backup_time = datetime.datetime.fromtimestamp(os.path.getmtime(latest_backup))
            time_diff = (timestamp - backup_time).total_seconds()

            # Only backup if threshold exceeded
            should_backup = time_diff >= threshold_seconds

        if should_backup:
            if category == 'latest' and os.path.exists(backup_path):
                # Keep the copy being replaced: 'latest' alone held only the newest run, so a
                # second bad run inside 5 min left the good deck only in the slower tiers.
                previous_dir = os.path.join(backup_base, 'previous')
                os.makedirs(previous_dir, exist_ok=True)
                previous_path = os.path.join(previous_dir, filename)
                os.replace(backup_path, previous_path)   # keeps the time it was made
                created_backups['previous'] = previous_path
            shutil.copy2(ppt_path, backup_path)
            # Stamp the backup with the time it was MADE. copy2 keeps the source's mtime,
            # so the tier test above measured the deck's content age instead: after two
            # quick reruns of a deck last saved >1 h ago every tier looked stale and all
            # were overwritten with the newest copy (bug hunt 2026-10-24).
            now_ts = timestamp.timestamp()
            os.utime(backup_path, (now_ts, now_ts))
            created_backups[category] = backup_path

    return created_backups
