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
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional


def find_cache_files(search_paths: Optional[List[Path]] = None) -> List[Path]:
    """Find all ilab_requests_cache.csv files in the given paths."""
    if search_paths is None:
        # Default: search user's home, OneDrive, Desktop, Documents
        search_paths = [
            Path.home(),
            Path.home() / "OneDrive - UCSF",
            Path.home() / "Documents",
            Path.home() / "Desktop",
            Path.home() / "Workspace",
        ]

    found = []
    for base in search_paths:
        if not base.exists():
            continue
        for cache_file in base.rglob("ilab_requests_cache.csv"):
            if cache_file.is_file():
                found.append(cache_file)

    return sorted(set(found))


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


def merge_records(all_records: Dict[str, Dict[str, dict]]) -> Dict[str, dict]:
    """
    Merge records from multiple cache files, preferring newest last_synced time.

    Returns a dict of request_id -> record (with the newest version of each request).
    """
    merged = {}
    for cache_path, records in all_records.items():
        for req_id, record in records.items():
            existing = merged.get(req_id)
            if existing is None:
                record["_source"] = str(cache_path)
                merged[req_id] = record
            else:
                # Prefer the record with the most recent last_synced time
                existing_time = existing.get("last_synced", "")
                new_time = record.get("last_synced", "")

                # Handle "manual" or empty timestamps
                if existing_time in ("", "manual"):
                    existing_time = "0"
                if new_time in ("", "manual"):
                    new_time = "0"

                try:
                    if new_time > existing_time:
                        record["_source"] = str(cache_path)
                        merged[req_id] = record
                except TypeError:
                    pass  # if timestamps can't be compared, keep existing

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


def consolidate_interactive() -> None:
    """Interactive consolidation process."""
    print("\n🔍 Scanning for cache files...")
    cache_files = find_cache_files()

    if not cache_files:
        print("✓ No duplicate cache files found.")
        return

    if len(cache_files) == 1:
        print(f"✓ Single cache file found: {cache_files[0]}")
        print("  No consolidation needed.")
        return

    print(f"\n⚠ Found {len(cache_files)} cache file(s) — potential duplicates.\n")

    # Load all records
    all_records = {}
    for cache_path in cache_files:
        print(f"Reading {cache_path}...")
        all_records[cache_path] = load_cache_records(cache_path)

    # Print summary
    print_summary(cache_files, all_records)

    # Ask user which file to keep
    print("\n📋 OPTIONS:")
    print(f"  1. Keep the app directory cache:")
    print(f"     {Path(__file__).parent / 'ilab_requests_cache.csv'}")
    print(f"  2. Keep the OneDrive cache (if found)")
    print(f"  3. Keep a specific file")
    print(f"  4. Cancel (do nothing)\n")

    choice = input("Choose an option (1-4): ").strip()

    if choice == "4" or not choice:
        print("Cancelled. No changes made.")
        return

    # Determine target file
    target_file = None
    if choice == "1":
        target_file = Path(__file__).parent / "ilab_requests_cache.csv"
    elif choice == "2":
        # Find OneDrive cache
        for f in cache_files:
            if "OneDrive" in str(f):
                target_file = f
                break
        if not target_file:
            print("No OneDrive cache found.")
            return
    elif choice == "3":
        print("\nAvailable files:")
        for i, f in enumerate(cache_files, 1):
            print(f"  {i}. {f}")
        file_choice = input("\nChoose file number: ").strip()
        try:
            target_file = cache_files[int(file_choice) - 1]
        except (ValueError, IndexError):
            print("Invalid selection.")
            return
    else:
        print("Invalid choice.")
        return

    # Merge and write
    print(f"\n📝 Consolidating into: {target_file}")
    merged = merge_records(all_records)

    # Write consolidated cache
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
        "form_data", "milestones_data", "_source",
    ]

    try:
        with open(target_file, "w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=ALL_COLS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(merged.values())
        print(f"✓ Consolidated {len(merged)} records into {target_file}\n")
    except Exception as e:
        print(f"✗ Error writing consolidated cache: {e}")
        return

    # Optional: backup and delete old files
    print("\n🗑️  What to do with the old cache files?")
    print("  1. Keep them (safe, but leaves duplicates)")
    print("  2. Rename them to .backup (preserves them but won't be used)")
    print("  3. Delete them (irreversible!)")

    cleanup = input("\nChoose (1-3): ").strip()

    if cleanup in ("2", "3"):
        for cache_path in cache_files:
            if cache_path == target_file:
                continue
            try:
                if cleanup == "2":
                    backup_path = cache_path.with_suffix(".csv.backup")
                    cache_path.rename(backup_path)
                    print(f"  ✓ Backed up: {cache_path} → {backup_path}")
                elif cleanup == "3":
                    cache_path.unlink()
                    print(f"  ✓ Deleted: {cache_path}")
            except Exception as e:
                print(f"  ⚠ Error handling {cache_path}: {e}")

    print("\n✓ Consolidation complete!")
    print(f"\n💡 TIP: To prevent future duplicates, both users should set the same")
    print(f"   cache file path in prefs.json (data_file setting).")
    print(f"   Recommended shared path:")
    print(f"   C:\\Users\\NIC-ADMIN4\\OneDrive - UCSF\\Documents - CALM\\ilab_requests_cache.csv")


if __name__ == "__main__":
    consolidate_interactive()
