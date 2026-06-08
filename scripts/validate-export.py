#!/usr/bin/env python3
"""Validate export output from outlook-cli.

Usage:
    python validate-export.py <export-dir> [--format json|markdown]

Checks:
    - Files exist and are readable
    - JSON files are valid JSON
    - Markdown files have expected frontmatter
    - No empty files
    - Reports file count and total size
"""

import argparse
import json
import sys
from pathlib import Path


def validate_json_file(filepath: Path) -> tuple[bool, str]:
    """Validate a JSON file."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not data:
            return False, "Empty JSON object"
        return True, "Valid"
    except json.JSONDecodeError as e:
        return False, f"Invalid JSON: {e}"
    except Exception as e:
        return False, f"Error reading file: {e}"


def validate_markdown_file(filepath: Path) -> tuple[bool, str]:
    """Validate a markdown file."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        if not content.strip():
            return False, "Empty file"
        if content.startswith('---'):
            # Has frontmatter
            parts = content.split('---', 2)
            if len(parts) < 3:
                return False, "Invalid frontmatter"
        return True, "Valid"
    except Exception as e:
        return False, f"Error reading file: {e}"


def main():
    parser = argparse.ArgumentParser(description='Validate outlook-cli export')
    parser.add_argument('export_dir', help='Export directory to validate')
    parser.add_argument('--format', choices=['json', 'markdown'],
                       default='markdown', help='Export format to validate')
    args = parser.parse_args()

    export_path = Path(args.export_dir)
    if not export_path.exists():
        print(f"Error: Directory not found: {export_path}")
        sys.exit(1)

    if not export_path.is_dir():
        print(f"Error: Not a directory: {export_path}")
        sys.exit(1)

    # Find files
    if args.format == 'json':
        files = list(export_path.glob('*.json'))
        validator = validate_json_file
    else:
        files = list(export_path.glob('*.md'))
        validator = validate_markdown_file

    if not files:
        print(f"No {args.format} files found in {export_path}")
        sys.exit(1)

    # Validate each file
    valid_count = 0
    invalid_count = 0
    total_size = 0
    errors = []

    for filepath in sorted(files):
        is_valid, message = validator(filepath)
        total_size += filepath.stat().st_size

        if is_valid:
            valid_count += 1
        else:
            invalid_count += 1
            errors.append(f"  {filepath.name}: {message}")

    # Report
    print(f"Export Validation Report")
    print(f"========================")
    print(f"Directory: {export_path}")
    print(f"Format: {args.format}")
    print(f"Files found: {len(files)}")
    print(f"Valid: {valid_count}")
    print(f"Invalid: {invalid_count}")
    print(f"Total size: {total_size / 1024:.1f} KB")

    if errors:
        print(f"\nErrors:")
        for error in errors:
            print(error)
        sys.exit(1)
    else:
        print(f"\nAll files valid!")
        sys.exit(0)


if __name__ == '__main__':
    main()
