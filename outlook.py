#!/usr/bin/env python
"""
Outlook CLI 0.3.0 - Complete command-line interface for Microsoft Outlook

Search, view, export, send emails and manage calendar events from your terminal.

Usage:
    python outlook.py folders                    # List all folders
    python outlook.py search --folder Inbox      # Search & display in terminal
    python outlook.py export --output ./md       # Export to markdown
    python outlook.py send --to user@email.com   # Send email
    python outlook.py cal list                   # List calendar events
    python outlook.py read <message-id>          # Read single email

Full documentation: https://github.com/ob-cheng/outlook-cli-skill
"""

import sys


def _check_wsl_and_exit_if_needed():
    """Check WSL environment and print helpful error if misconfigured."""
    try:
        from outlook_cli.core.wsl import check_wsl_environment
        ok, error = check_wsl_environment()
        if not ok:
            print(error, file=sys.stderr)
            sys.exit(1)
    except ImportError:
        pass


if __name__ == '__main__':
    _check_wsl_and_exit_if_needed()
    from outlook_cli.cli import main
    sys.exit(main())
