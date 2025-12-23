#!/usr/bin/env python3
"""
Remove remaining Confluence breadcrumbs and author lines from all markdown files.
"""

import re
from pathlib import Path

def clean_file(file_path):
    """Clean a single markdown file.

    Removes Confluence breadcrumbs, author information, and excessive whitespace.

    Args:
        file_path: Path to the markdown file to clean

    Returns:
        True if file was modified, False otherwise
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        original_content = content

        # Remove breadcrumb lines - can span multiple lines
        # Matches: "2.  [CALM Microscopy Wiki\n    Home](CALM-Microscopy-Wiki-Home_512554980.html)"
        content = re.sub(r'\d+\.\s+\[CALM Microscopy Wiki[\s\S]*?Home\]\([^)]+\)', '', content)

        # Remove "Created by" lines that may span multiple lines
        # Matches: "Created by [ Delaine Larsen, last updated by [ Herrington,\nKari on Sep 26, 2025"
        content = re.sub(r'Created by \[[\s\S]*?(?:on|by)\s+[A-Z][a-z]+\s+\d+,\s+\d{4}', '', content)

        # Remove excessive blank lines (more than 2 consecutive)
        content = re.sub(r'\n{3,}', '\n\n', content)

        # Only write if content changed
        if content != original_content:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        return False
    except Exception as e:
        print(f"  [ERROR] Failed to clean {file_path.name}: {e}")
        return False

def main():
    """Process all markdown files in pages directories."""
    pages_dir = Path("wiki/pages")

    # Validate pages directory exists
    if not pages_dir.exists():
        print(f"[ERROR] Pages directory not found: {pages_dir}")
        return

    modified_count = 0
    total_count = 0
    error_count = 0

    # Process all .md files in all category subdirectories
    for md_file in pages_dir.rglob("*.md"):
        if md_file.name == "_file_mapping.txt":
            continue

        total_count += 1
        result = clean_file(md_file)
        if result is True:
            modified_count += 1
            print(f"Cleaned: {md_file.relative_to(pages_dir)}")
        elif result is False and md_file.exists():
            # File processed but not modified
            pass
        else:
            error_count += 1

    print(f"\n[SUMMARY]")
    print(f"Total files processed: {total_count}")
    print(f"Files modified: {modified_count}")
    print(f"Files unchanged: {total_count - modified_count - error_count}")
    if error_count > 0:
        print(f"Errors: {error_count}")

if __name__ == "__main__":
    main()
