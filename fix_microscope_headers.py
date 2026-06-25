#!/usr/bin/env python3
"""
Fix header formatting in microscope pages to consistent levels.

Rules:
- H1: Page title
- H2: FPBASE, Optics
- H3: All subsections (Objectives, Cameras, Lasers, Filter Turret, etc.)
"""

import re
from pathlib import Path


def fix_headers(content):
    """Fix header levels in microscope page.

    Args:
        content: Markdown content

    Returns:
        Fixed content
    """
    lines = content.split('\n')
    in_front_matter = False
    front_matter_count = 0
    result_lines = []

    for line in lines:
        # Track front matter
        if line.strip() == '---':
            front_matter_count += 1
            result_lines.append(line)
            if front_matter_count == 2:
                in_front_matter = False
            elif front_matter_count == 1:
                in_front_matter = True
            continue

        if front_matter_count < 2:
            result_lines.append(line)
            continue

        # Process headers outside front matter
        if line.startswith('#'):
            # Count the number of # symbols
            hash_count = len(line) - len(line.lstrip('#'))
            rest_of_line = line[hash_count:].strip()

            # H1: Page title (keep as is)
            if hash_count == 1:
                result_lines.append(line)

            # H2 or more: Check if it's FPBASE or Optics
            elif 'FPBASE' in rest_of_line.upper() or 'FPBASE' in rest_of_line:
                result_lines.append(f'## {rest_of_line}')

            elif rest_of_line.strip() == 'Optics':
                result_lines.append(f'## Optics')

            # H3 or more: Convert all other headers to H3
            elif hash_count >= 3:
                # Common subsection headers that should be H3
                subsection_keywords = [
                    'Objective', 'Filter', 'Camera', 'Laser', 'Excitation', 'Emission',
                    'Sample', 'Interface', 'Software', 'Computer', 'Stage', 'Illumination',
                    'Detector', 'Light', 'Dichroic', 'Control', 'Focus', 'Autofocus',
                    'Motorized', 'Feature', 'Configuration', 'Accessor', 'Note',
                    'Tip', 'Protocol', 'Training', 'Maintenance', 'Troubleshoot', 'Wheel',
                    'Turret', 'holder', 'Requirements'
                ]

                # Check if this is a subsection header
                is_subsection = any(keyword.lower() in rest_of_line.lower() for keyword in subsection_keywords)

                if is_subsection or hash_count >= 4:
                    result_lines.append(f'### {rest_of_line}')
                else:
                    # Keep H3 as H3
                    result_lines.append(line)
            else:
                result_lines.append(line)
        else:
            result_lines.append(line)

    return '\n'.join(result_lines)


def process_file(file_path):
    """Process a single file.

    Args:
        file_path: Path to markdown file

    Returns:
        True if modified, False otherwise
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            original = f.read()

        fixed = fix_headers(original)

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

    print("Fixing microscope page headers...")
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
