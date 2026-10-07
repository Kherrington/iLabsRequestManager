#!/usr/bin/env python3
"""
Utility to find and consolidate duplicate cache files.

This script helps identify all copies of ilab_requests_cache.csv on your system
and merges them into a single authoritative cache file. It's useful when the app
has been run from multiple locations creating duplicate caches.

Usage:
    python consolidate_cache.py          # Interactive mode: scan & show duplicates
    python consolidate_cache.py --help   # Show options
"""

import csv
import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional


ALL_COLS = [
    "request_id", "name", "state", "created_at",
    "start_on", "end_on", "completed_on",
    "owner_name", "owner_email", "pi_name", "pi_email",
    "service_name", "last_synced",
    "assigned_to", "labels", "local_notes",
    "core_lab", "microscope",
    "training_date", "training_time", "training_day",
    "class_taken",
    "wf_laser_safety",
    "wf_emailed", "wf_class_scheduled", "wf_not_required",
    "wf_training_scheduled",
    "wf_post_email", "wf_post_listserve",
    "wf_post_approved", "wf_post_confirmed",
    "schedule_exported", "local_only",
    "form_data", "milestones_data",
]

BACKUP_DIR_NAME = "cache_conflict_backup"


def _prefs_data_file() -> Optional[Path]:
    """The cache path configured in prefs.json (data_file), if any."""
    try:
        with open(Path(__file__).resolve().parent / "prefs.json", encoding="utf-8") as fh:
            p = json.load(fh).get("data_file", "")
        return Path(p) if p and not p.lower().startswith("http") else None
    except Exception:
        return None


def is_cache_name(name: str) -> bool:
    """ilab_requests_cache.csv plus OneDrive conflict copies such as
    ilab_requests_cache-NIC-6D-2-10.csv (but not .backup / _mirror files)."""
    return bool(re.fullmatch(r"ilab_requests_cache(-.+)?\.csv", name))


def find_cache_files(search_paths: Optional[List[Path]] = None) -> List[Path]:
    """Find the cache file and all its conflict copies in the given paths."""
    if search_paths is None:
        search_paths = [
            Path(__file__).resolve().parent,
            Path.home() / "OneDrive - UCSF",
            Path.home() / "Documents",
            Path.home() / "Desktop",
        ]
        target = _prefs_data_file()
        if target is not None:
            search_paths.append(target.parent)

    found = set()
    for base in search_paths:
        if not base.exists():
            continue
        for f in base.rglob("ilab_requests_cache*.csv"):
            if (f.is_file() and is_cache_name(f.name)
                    and BACKUP_DIR_NAME not in f.parts
                    and "__pycache__" not in f.parts):
                found.add(f)
    return sorted(found)


def load_cache_records(filepath: Path) -> Dict[str, dict]:
    """Load all records from a cache CSV file."""
    records = {}
    try:
        with open(filepath, newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                if row.get("request_id"):
                    records[row["request_id"]] = row
    except Exception as e:
        print(f"  ⚠ Error reading {filepath}: {e}")
    return records


def _sync_key(cache_path: Path, record: dict):
    """Sort key: newest last_synced wins; file mtime breaks ties."""
    t = record.get("last_synced", "")
    if t in ("", "manual"):
        t = "0"
    try:
        mtime = cache_path.stat().st_mtime
    except OSError:
        mtime = 0
    return (t, mtime)


def merge_records(all_records: Dict[Path, Dict[str, dict]],
                  preferred: Optional[Path] = None) -> Dict[str, dict]:
    """Merge records from many cache files; newest last_synced wins per
    request_id (ties: *preferred* file, then newest file mtime)."""
    merged: Dict[str, dict] = {}
    best: Dict[str, tuple] = {}
    for cache_path, records in all_records.items():
        for req_id, record in records.items():
            key = _sync_key(cache_path, record) + (cache_path == preferred,)
            if req_id not in best or key > best[req_id]:
                best[req_id] = key
                merged[req_id] = record
    return merged


def print_summary(cache_files: List[Path], all_records: Dict[str, Dict[str, dict]]) -> None:
    """Print a summary of found cache files and their contents."""
    print(f"\n{'='*80}")
    print(f"Found {len(cache_files)} cache file(s):")
    print(f"{'='*80}\n")

    total_records = 0
    for cache_path in cache_files:
        records = all_records.get(cache_path, {})
        size_kb = cache_path.stat().st_size / 1024
        modified = datetime.fromtimestamp(cache_path.stat().st_mtime)
        print(f"📄 {cache_path}")
        print(f"   Size: {size_kb:.1f} KB")
        print(f"   Modified: {modified.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"   Records: {len(records)}")

        # Show unique request IDs in this file
        if records:
            sample_ids = sorted(records.keys())[:3]
            print(f"   Sample IDs: {', '.join(sample_ids)}")
            if len(records) > 3:
                print(f"             (... and {len(records) - 3} more)")
        print()
        total_records += len(records)

    print(f"{'─'*80}")
    print(f"Total unique records across all files: {total_records}")
    print(f"{'='*80}\n")


def consolidate_interactive(argv: Optional[List[str]] = None) -> None:
    """Merge every cache file / conflict copy into the prefs.json data_file."""
    argv = argv or []
    dry_run = "--dry-run" in argv
    assume_yes = "--yes" in argv

    target_file = _prefs_data_file() or (Path(__file__).resolve().parent / "ilab_requests_cache.csv")
    print(f"\n🎯 Target cache: {target_file}")
    print("🔍 Scanning for cache files and conflict copies...")
    cache_files = find_cache_files()
    if target_file not in cache_files and target_file.exists():
        cache_files.append(target_file)
    sources = [f for f in cache_files if f != target_file]

    if not sources:
        print("✓ Nothing to consolidate.")
        return

    all_records = {}
    for i, cache_path in enumerate(cache_files, 1):
        if i % 500 == 0:
            print(f"  read {i}/{len(cache_files)}...")
        all_records[cache_path] = load_cache_records(cache_path)

    merged = merge_records(all_records, preferred=target_file)
    own = len(all_records.get(target_file, {}))
    print(f"\nFound {len(sources)} other cache file(s) next to/besides the target.")
    print(f"Target currently has {own} record(s); merged result has {len(merged)}.")

    if dry_run:
        print("\n--dry-run: no changes made.")
        return
    if not assume_yes and input("\nWrite merged cache and move the copies into "
                                f"'{BACKUP_DIR_NAME}'? (y/N): ").strip().lower() != "y":
        print("Cancelled. No changes made.")
        return

    # Keep a backup of the target before overwriting it
    backup_dir = target_file.parent / BACKUP_DIR_NAME
    backup_dir.mkdir(exist_ok=True)
    if target_file.exists():
        shutil.copy2(target_file, backup_dir / f"{target_file.name}.pre_consolidate")

    with open(target_file, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=ALL_COLS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(merged.values())
    print(f"✓ Wrote {len(merged)} records to {target_file}")

    # Move (not delete) the copies, so everything is recoverable
    moved = 0
    for cache_path in sources:
        dest = backup_dir / cache_path.name
        if dest.exists():
            dest = backup_dir / f"{cache_path.parent.name}__{cache_path.name}"
        try:
            shutil.move(str(cache_path), str(dest))
            moved += 1
        except Exception as e:
            print(f"  ⚠ Could not move {cache_path}: {e}")
    print(f"✓ Moved {moved} file(s) to {backup_dir}")
    print("\n💡 Close the app on both computers before running this, and make sure")
    print("   both use the same data_file path in prefs.json.")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # emoji on cp1252 consoles
    consolidate_interactive(sys.argv[1:])
