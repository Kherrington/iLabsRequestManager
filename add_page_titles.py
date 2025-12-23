#!/usr/bin/env python3
"""
Add H1 page titles to all wiki pages based on their front matter title.
"""

import re
from pathlib import Path


def add_page_title(file_path):
    """Add H1 title to markdown file if it doesn't already have one.

    Args:
        file_path: Path to the markdown file

    Returns:
        True if file was modified, False otherwise
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Extract title from front matter
        title_match = re.search(r'^title:\s*(.+)$', content, re.MULTILINE)
        if not title_match:
            print(f"  [SKIP] No title in front matter: {file_path.name}")
            return False

        title = title_match.group(1).strip()

        # Check if content already starts with H1 title
        # Split by front matter end (---)
        parts = content.split('---', 2)
        if len(parts) < 3:
            print(f"  [ERROR] Invalid front matter: {file_path.name}")
            return False

        front_matter = parts[1]
        body_content = parts[2].strip()

        # Check if body already starts with H1
        if body_content.startswith(f'# {title}'):
            print(f"  [SKIP] Already has title: {file_path.name}")
            return False

        # Check if body already has any H1 at the start
        if re.match(r'^#\s+', body_content):
            # Replace existing H1 with title from front matter
            body_content = re.sub(r'^#\s+.+$', f'# {title}', body_content, count=1, flags=re.MULTILINE)
            new_content = f'---{front_matter}---\n\n{body_content}'
        else:
            # Add new H1 title at the top
            new_content = f'---{front_matter}---\n\n# {title}\n\n{body_content}'

        # Write updated content
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)

        return True

    except Exception as e:
        print(f"  [ERROR] Failed to process {file_path.name}: {e}")
        return False


def main():
    """Process all markdown files in the wiki pages directory."""
    pages_dir = Path("wiki/pages")

    if not pages_dir.exists():
        print(f"[ERROR] Pages directory not found: {pages_dir}")
        return

    print("Adding H1 page titles to wiki pages...")
    print("=" * 60)

    modified_count = 0
    skipped_count = 0
    error_count = 0
    total_count = 0

    # Process all .md files recursively
    for md_file in sorted(pages_dir.rglob("*.md")):
        # Skip mapping file
        if md_file.name == "_file_mapping.txt":
            continue

        total_count += 1
        result = add_page_title(md_file)

        if result is True:
            modified_count += 1
            print(f"[+] Added title: {md_file.relative_to(pages_dir)}")
        elif result is False:
            skipped_count += 1
        else:
            error_count += 1

    print("\n" + "=" * 60)
    print(f"[SUMMARY]")
    print(f"  Total files processed: {total_count}")
    print(f"  Files modified:        {modified_count}")
    print(f"  Files skipped:         {skipped_count}")
    if error_count > 0:
        print(f"  Errors:                {error_count}")
    print("\n[OK] Title addition complete!")


if __name__ == "__main__":
    main()
