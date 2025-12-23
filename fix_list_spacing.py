#!/usr/bin/env python3
"""
Fix list spacing in microscope pages - convert double-spaced lists to single-spaced.

Double-spaced lists have blank lines between items, single-spaced lists don't.
"""

import re
from pathlib import Path


def fix_list_spacing(content):
    """Remove extra blank lines between list items.

    Args:
        content: Markdown content

    Returns:
        Content with single-spaced lists
    """
    lines = content.split('\n')
    result_lines = []
    in_list = False
    prev_was_list_item = False

    for i, line in enumerate(lines):
        # Check if current line is a list item (numbered or bulleted)
        is_list_item = bool(re.match(r'^\s*(\d+\.|-|\*)\s+', line))

        # Check if line is blank
        is_blank = line.strip() == ''

        # If we're in a list and encounter a blank line before another list item, skip it
        if in_list and is_blank and prev_was_list_item:
            # Look ahead to see if next line is a list item
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                next_is_list_item = bool(re.match(r'^\s*(\d+\.|-|\*)\s+', next_line))
                if next_is_list_item:
                    # Skip this blank line - it's between list items
                    continue

        # Add the line
        result_lines.append(line)

        # Track list state
        if is_list_item:
            in_list = True
            prev_was_list_item = True
        elif is_blank:
            prev_was_list_item = False
        else:
            # Non-blank, non-list line
            if not is_list_item and line.strip() != '':
                # If it's not indented (continuation of list item), we're out of the list
                if not line.startswith('  ') and not line.startswith('\t'):
                    in_list = False
            prev_was_list_item = False

    return '\n'.join(result_lines)


def process_file(file_path):
    """Process a single microscope file.

    Args:
        file_path: Path to markdown file

    Returns:
        True if modified, False otherwise
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            original = f.read()

        fixed = fix_list_spacing(original)

        if fixed != original:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(fixed)
            return True

        return False

    except Exception as e:
        print(f"  [ERROR] {file_path.name}: {e}")
        return None


def main():
    """Main function."""
    microscopes_dir = Path("wiki/pages/microscopes")

    if not microscopes_dir.exists():
        print(f"[ERROR] Directory not found: {microscopes_dir}")
        return

    print("Fixing list spacing in microscope pages...")
    print("=" * 60)
    print("Converting double-spaced lists to single-spaced")
    print("=" * 60)

    modified = 0
    unchanged = 0
    errors = 0

    for md_file in sorted(microscopes_dir.glob("*.md")):
        result = process_file(md_file)

        if result is True:
            modified += 1
            print(f"[+] Fixed: {md_file.name}")
        elif result is False:
            unchanged += 1
        else:
            errors += 1

    print("\n" + "=" * 60)
    print(f"[SUMMARY]")
    print(f"  Total:     {modified + unchanged + errors}")
    print(f"  Modified:  {modified}")
    print(f"  Unchanged: {unchanged}")
    if errors > 0:
        print(f"  Errors:    {errors}")
    print("\n[OK] Complete!")


if __name__ == "__main__":
    main()
