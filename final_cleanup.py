#!/usr/bin/env python3
"""
Final cleanup script to remove all remaining Confluence artifacts from markdown files.
This script removes external-link classes and other remaining HTML attributes.
"""

import re
from pathlib import Path


def clean_confluence_artifacts(content):
    """Remove all remaining Confluence artifacts from markdown content.

    Args:
        content: Markdown file content as string

    Returns:
        Cleaned content with artifacts removed
    """
    # Remove complex external-link patterns with multiple attributes (within table cells)
    # Example: "text"){.external-link | other | attributes
    content = re.sub(r'"\)\{\.external-link[^\n]*?\|', '")', content)

    # Remove external-link class with rel attribute
    # Example: {.external-link rel="nofollow"}
    content = re.sub(r'\{\.external-link\s+rel="nofollow"\}', '', content)

    # Remove standalone external-link class
    content = re.sub(r'\{\.external-link\}', '', content)

    # Remove external-link with pipe delimiter patterns
    # Example: {.external-link |
    content = re.sub(r'\{\.external-link\s*\|', '', content)

    # Remove Confluence content-wrapper divs (both standalone and inline)
    # Example: ::: content-wrapper
    content = re.sub(r'^:+\s*content-wrapper\s*$', '', content, flags=re.MULTILINE)
    content = re.sub(r'\|:+\s*content-wrapper\s*', '| ', content)
    content = re.sub(r'\|:+\s*', '| ', content)
    content = re.sub(r'^:+\s*$', '', content, flags=re.MULTILINE)

    # Remove Confluence embedded classes
    # Example: .confluence-embedded-manual-size}
    content = re.sub(r'\.confluence-[\w-]+\}', '', content)

    # Remove linked-resource attributes
    # Example: linked-resource-container-id="517182120"
    content = re.sub(r'linked-resource-[\w-]+="[^"]*"\s*', '', content)
    content = re.sub(r'linked-resource-[\w-]+="[^"]*"', '', content)

    # Remove any remaining class attributes in curly braces
    # Example: {.some-class}
    content = re.sub(r'\{\.[\w-]+\}', '', content)

    # Remove any remaining attribute blocks with style
    # Example: {style="..."}
    content = re.sub(r'\{style="[^"]*"\}', '', content)

    # Remove any remaining complex attribute blocks
    # Example: {.class .other-class attribute="value"}
    content = re.sub(r'\{\.[\w\s.-]+(?:\s+[\w-]+="[^"]*")*\}', '', content)

    # Clean up any double spaces created by removals
    content = re.sub(r'  +', ' ', content)

    # Clean up space before punctuation
    content = re.sub(r' +([.,;:!?])', r'\1', content)

    # Remove excessive blank lines (more than 2 consecutive)
    content = re.sub(r'\n{3,}', '\n\n', content)

    # Strip trailing whitespace from each line
    lines = content.split('\n')
    lines = [line.rstrip() for line in lines]
    content = '\n'.join(lines)

    return content.strip() + '\n'


def process_file(file_path):
    """Process a single markdown file.

    Args:
        file_path: Path to markdown file

    Returns:
        True if file was modified, False otherwise
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            original_content = f.read()

        cleaned_content = clean_confluence_artifacts(original_content)

        if cleaned_content != original_content:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(cleaned_content)
            return True

        return False

    except Exception as e:
        print(f"  [ERROR] Failed to process {file_path.name}: {e}")
        return None


def main():
    """Process all markdown files in the wiki pages directory."""
    pages_dir = Path("wiki/pages")

    if not pages_dir.exists():
        print(f"[ERROR] Pages directory not found: {pages_dir}")
        return

    print("Cleaning Confluence artifacts from wiki pages...")
    print("=" * 60)

    modified_count = 0
    unchanged_count = 0
    error_count = 0
    total_count = 0

    # Process all .md files recursively
    for md_file in sorted(pages_dir.rglob("*.md")):
        # Skip mapping file
        if md_file.name == "_file_mapping.txt":
            continue

        total_count += 1
        result = process_file(md_file)

        if result is True:
            modified_count += 1
            print(f"[+] Cleaned: {md_file.relative_to(pages_dir)}")
        elif result is False:
            unchanged_count += 1
        else:
            error_count += 1

    print("\n" + "=" * 60)
    print(f"[SUMMARY]")
    print(f"  Total files processed: {total_count}")
    print(f"  Files modified:        {modified_count}")
    print(f"  Files unchanged:       {unchanged_count}")
    if error_count > 0:
        print(f"  Errors:                {error_count}")
    print("\n[OK] Cleanup complete!")


if __name__ == "__main__":
    main()
