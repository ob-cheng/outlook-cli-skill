"""Command-line interface for Outlook Email CLI."""

import argparse
import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path


def _check_send_allowed() -> tuple[bool, str]:
    """Check if direct sending is allowed.

    Checks config file first (send_mode), then falls back to env var.
    Config takes priority.

    Returns:
        tuple: (allowed: bool, error_message: str or None)
    """
    # Config file takes priority
    from .core.config import ConfigManager
    cfg = ConfigManager()
    config_mode = cfg.get('send_mode')
    if config_mode == 'send':
        return True, None

    # Fallback: env var
    if os.environ.get('OUTLOOK_CLI_ALLOW_SEND', '').strip() == '1':
        return True, None
    return False, (
        "Direct sending is disabled for safety. Emails are saved as drafts by default.\n"
        "To enable direct sending, set: python outlook.py config set send_mode send\n"
        "Or set environment variable: OUTLOOK_CLI_ALLOW_SEND=1\n"
        "Or remove --send flag to save as draft."
    )

from . import __version__
from .core.connection import connect_to_outlook
from .services.search import SearchService
from .services.viewer import ViewerService
from .services.compose import ComposeService
from .services.calendar import CalendarService
from .utils.formatting import parse_date


DEFAULT_LOOKBACK_DAYS = 7


def _output_json(data: dict) -> None:
    """Output data as formatted JSON."""
    print(json.dumps(data, indent=2, ensure_ascii=False))


def _json_error(error: str, code: str = "error") -> dict:
    """Create a JSON error response."""
    return {"success": False, "error": error, "code": code}


def _json_success(data: dict) -> dict:
    """Create a JSON success response."""
    return {"success": True, **data}


def create_parser() -> argparse.ArgumentParser:
    """Create the argument parser with subcommands."""
    parser = argparse.ArgumentParser(
        prog='outlook',
        description='Outlook Email CLI - Search, View, and Export emails',
    )
    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')

    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # =========================================================================
    # folders command
    # =========================================================================
    folders_parser = subparsers.add_parser(
        'folders',
        help='List all available Outlook folders',
    )
    folders_parser.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # search command
    # =========================================================================
    search_parser = subparsers.add_parser(
        'search',
        help='Search emails and display in terminal',
    )
    _add_search_args(search_parser)
    search_parser.add_argument(
        '--export', '-e',
        type=str,
        metavar='DIR',
        help='Also export results to markdown files in DIR',
    )
    search_parser.add_argument(
        '--no-view',
        action='store_true',
        help='Skip terminal display (useful with --export)',
    )
    search_parser.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # export command
    # =========================================================================
    export_parser = subparsers.add_parser(
        'export',
        help='Export emails to Obsidian markdown files',
    )
    _add_search_args(export_parser)
    export_parser.add_argument(
        '--output', '-o',
        type=str,
        required=True,
        help='Output directory for markdown files',
    )
    export_parser.add_argument(
        '--no-threads',
        action='store_true',
        help='Export each email as separate file (don\'t group threads)',
    )
    export_parser.add_argument(
        '--no-overwrite',
        action='store_true',
        help='Skip files that already exist',
    )
    export_parser.add_argument(
        '--incremental',
        action='store_true',
        help='Only export emails since last run (saves state to extraction_state.json)',
    )
    export_parser.add_argument(
        '--format',
        type=str,
        choices=['markdown', 'json'],
        default='markdown',
        help='Output format: markdown (default) or json',
    )
    export_parser.add_argument(
        '--batch',
        action='store_true',
        help='For JSON format: combine all emails into a single file (more token-efficient)',
    )
    # --stdout and --json are mutually exclusive: --stdout outputs email data,
    # --json outputs an export summary. Using both is ambiguous.
    export_output_group = export_parser.add_mutually_exclusive_group()
    export_output_group.add_argument(
        '--stdout',
        action='store_true',
        help='Output email data JSON to terminal instead of files (for agent/pipeline use)',
    )
    export_output_group.add_argument(
        '--json',
        action='store_true',
        help='Output export summary as JSON (file count, status, etc.)',
    )

    # =========================================================================
    # read command
    # =========================================================================
    read_parser = subparsers.add_parser(
        'read',
        help='Read emails by message ID',
    )
    read_parser.add_argument(
        'message_ids',
        nargs='+',
        help='One or more message IDs (EntryID) to read',
    )
    read_parser.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # send command
    # =========================================================================
    send_parser = subparsers.add_parser(
        'send',
        help='Send a new email',
    )
    send_parser.add_argument(
        '--to', '-t',
        type=str,
        required=True,
        help='Recipient email addresses (comma-separated)',
    )
    send_parser.add_argument(
        '--subject', '-s',
        type=str,
        required=True,
        help='Email subject',
    )
    send_parser.add_argument(
        '--body', '-b',
        type=str,
        required=True,
        help='Email body text',
    )
    send_parser.add_argument(
        '--cc',
        type=str,
        help='CC email addresses (comma-separated)',
    )
    send_parser.add_argument(
        '--bcc',
        type=str,
        help='BCC email addresses (comma-separated)',
    )
    send_parser.add_argument(
        '--attach', '-a',
        type=str,
        action='append',
        help='File path to attach (can specify multiple)',
    )
    send_parser.add_argument(
        '--html',
        action='store_true',
        help='Body is HTML format',
    )
    send_parser.add_argument(
        '--send',
        action='store_true',
        help='Send immediately (requires OUTLOOK_CLI_ALLOW_SEND=1 env var)',
    )
    send_parser.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # reply command
    # =========================================================================
    reply_parser = subparsers.add_parser(
        'reply',
        help='Reply to an email',
    )
    reply_parser.add_argument(
        'message_id',
        help='The message ID (EntryID) to reply to',
    )
    reply_parser.add_argument(
        '--body', '-b',
        type=str,
        required=True,
        help='Reply message body',
    )
    reply_parser.add_argument(
        '--all',
        action='store_true',
        help='Reply to all recipients',
    )
    reply_parser.add_argument(
        '--attach', '-a',
        type=str,
        action='append',
        help='File path to attach (can specify multiple)',
    )
    reply_parser.add_argument(
        '--html',
        action='store_true',
        help='Body is HTML format',
    )
    reply_parser.add_argument(
        '--send',
        action='store_true',
        help='Send immediately (requires OUTLOOK_CLI_ALLOW_SEND=1 env var)',
    )
    reply_parser.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # forward command
    # =========================================================================
    forward_parser = subparsers.add_parser(
        'forward',
        help='Forward an email',
    )
    forward_parser.add_argument(
        'message_id',
        help='The message ID (EntryID) to forward',
    )
    forward_parser.add_argument(
        '--to', '-t',
        type=str,
        required=True,
        help='Recipient email addresses (comma-separated)',
    )
    forward_parser.add_argument(
        '--body', '-b',
        type=str,
        help='Optional message to add',
    )
    forward_parser.add_argument(
        '--cc',
        type=str,
        help='CC email addresses (comma-separated)',
    )
    forward_parser.add_argument(
        '--bcc',
        type=str,
        help='BCC email addresses (comma-separated)',
    )
    forward_parser.add_argument(
        '--attach', '-a',
        type=str,
        action='append',
        help='Additional file paths to attach (can specify multiple)',
    )
    forward_parser.add_argument(
        '--html',
        action='store_true',
        help='Body is HTML format',
    )
    forward_parser.add_argument(
        '--send',
        action='store_true',
        help='Send immediately (requires OUTLOOK_CLI_ALLOW_SEND=1 env var)',
    )
    forward_parser.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # cal command (calendar)
    # =========================================================================
    cal_parser = subparsers.add_parser('cal', help='Calendar management')
    cal_subparsers = cal_parser.add_subparsers(dest='cal_command', help='Calendar commands')

    # cal list
    cal_list = cal_subparsers.add_parser('list', help='List calendar events')
    cal_list.add_argument('--start', type=str, help='Start date (YYYY-MM-DD, default: today)')
    cal_list.add_argument('--end', type=str, help='End date (YYYY-MM-DD, default: 7 days from start)')
    cal_list.add_argument('--subject', type=str, help='Filter by subject (substring match)')
    cal_list.add_argument('--location', type=str, help='Filter by location (substring match)')
    cal_list.add_argument('--organizer', type=str, help='Filter by organizer email')
    cal_list.add_argument('--all-day', action='store_true', help='All-day events only')
    cal_list.add_argument('--recurring', action='store_true', help='Recurring events only')
    cal_list.add_argument('--json', action='store_true', help='Output as JSON')

    # cal read
    cal_read = cal_subparsers.add_parser('read', help='Read event details')
    cal_read.add_argument('event_id', help='Event ID (EntryID)')
    cal_read.add_argument('--json', action='store_true', help='Output as JSON')

    # cal create
    cal_create = cal_subparsers.add_parser('create', help='Create a new event')
    cal_create.add_argument('--subject', '-s', type=str, required=True, help='Event subject')
    cal_create.add_argument('--start', type=str, required=True, help='Start date/time (YYYY-MM-DD HH:MM)')
    cal_create.add_argument('--end', type=str, required=True, help='End date/time (YYYY-MM-DD HH:MM)')
    cal_create.add_argument('--location', '-l', type=str, help='Event location')
    cal_create.add_argument('--body', '-b', type=str, help='Event description')
    cal_create.add_argument('--required', type=str, help='Required attendees (comma-separated)')
    cal_create.add_argument('--optional', type=str, help='Optional attendees (comma-separated)')
    cal_create.add_argument('--reminder', type=int, help='Reminder minutes before event (default: 15)')
    cal_create.add_argument('--no-reminder', action='store_true', help='No reminder')
    cal_create.add_argument('--json', action='store_true', help='Output as JSON')

    # cal delete
    cal_delete = cal_subparsers.add_parser('delete', help='Delete an event')
    cal_delete.add_argument('event_id', help='Event ID (EntryID)')
    cal_delete.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # tasks command
    # =========================================================================
    tasks_parser = subparsers.add_parser('tasks', help='Task/todo management')
    tasks_subparsers = tasks_parser.add_subparsers(dest='tasks_command', help='Task commands')

    # tasks list
    tasks_list = tasks_subparsers.add_parser('list', help='List tasks')
    tasks_list.add_argument('--status', type=str, choices=['not_started', 'in_progress', 'completed', 'waiting', 'deferred'], help='Filter by status')
    tasks_list.add_argument('--all', action='store_true', help='Include completed tasks')
    tasks_list.add_argument('--due-before', type=str, help='Tasks due before date (YYYY-MM-DD)')
    tasks_list.add_argument('--due-after', type=str, help='Tasks due after date (YYYY-MM-DD)')
    tasks_list.add_argument('--priority', type=str, choices=['low', 'normal', 'high'], help='Filter by priority')
    tasks_list.add_argument('--category', type=str, help='Filter by category')
    tasks_list.add_argument('--json', action='store_true', help='Output as JSON')

    # tasks read
    tasks_read = tasks_subparsers.add_parser('read', help='Read task details')
    tasks_read.add_argument('task_id', help='Task ID (EntryID)')
    tasks_read.add_argument('--json', action='store_true', help='Output as JSON')

    # tasks create
    tasks_create = tasks_subparsers.add_parser('create', help='Create a new task')
    tasks_create.add_argument('--subject', '-s', type=str, required=True, help='Task subject')
    tasks_create.add_argument('--due', type=str, help='Due date (YYYY-MM-DD)')
    tasks_create.add_argument('--start', type=str, help='Start date (YYYY-MM-DD)')
    tasks_create.add_argument('--priority', type=str, choices=['low', 'normal', 'high'], default='normal', help='Priority (default: normal)')
    tasks_create.add_argument('--body', '-b', type=str, help='Task description')
    tasks_create.add_argument('--category', type=str, help='Category name')
    tasks_create.add_argument('--reminder', type=str, help='Reminder date/time (YYYY-MM-DD HH:MM)')
    tasks_create.add_argument('--json', action='store_true', help='Output as JSON')

    # tasks complete
    tasks_complete = tasks_subparsers.add_parser('complete', help='Mark task as complete')
    tasks_complete.add_argument('task_id', help='Task ID (EntryID)')
    tasks_complete.add_argument('--json', action='store_true', help='Output as JSON')

    # tasks delete
    tasks_delete = tasks_subparsers.add_parser('delete', help='Delete a task')
    tasks_delete.add_argument('task_id', help='Task ID (EntryID)')
    tasks_delete.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # notes command
    # =========================================================================
    notes_parser = subparsers.add_parser('notes', help='Notes management')
    notes_subparsers = notes_parser.add_subparsers(dest='notes_command', help='Notes commands')

    # notes list
    notes_list = notes_subparsers.add_parser('list', help='List notes')
    notes_list.add_argument('--color', type=str, choices=['blue', 'green', 'pink', 'yellow', 'white'], help='Filter by color')
    notes_list.add_argument('--category', type=str, help='Filter by category')
    notes_list.add_argument('--keyword', '-k', type=str, help='Search keyword in subject/body')
    notes_list.add_argument('--limit', '-n', type=int, help='Limit number of results')
    notes_list.add_argument('--json', action='store_true', help='Output as JSON')

    # notes read
    notes_read = notes_subparsers.add_parser('read', help='Read note details')
    notes_read.add_argument('note_id', help='Note ID (EntryID)')
    notes_read.add_argument('--json', action='store_true', help='Output as JSON')

    # notes create
    notes_create = notes_subparsers.add_parser('create', help='Create a new note')
    notes_create.add_argument('--body', '-b', type=str, required=True, help='Note content')
    notes_create.add_argument('--color', type=str, choices=['blue', 'green', 'pink', 'yellow', 'white'], default='yellow', help='Note color (default: yellow)')
    notes_create.add_argument('--category', type=str, help='Category name')
    notes_create.add_argument('--json', action='store_true', help='Output as JSON')

    # notes delete
    notes_delete = notes_subparsers.add_parser('delete', help='Delete a note')
    notes_delete.add_argument('note_id', help='Note ID (EntryID)')
    notes_delete.add_argument('--json', action='store_true', help='Output as JSON')

    # =========================================================================
    # config command
    # =========================================================================
    config_parser = subparsers.add_parser(
        'config',
        help='View or change configuration (send mode, drafting rules, humanizer)',
    )
    config_subparsers = config_parser.add_subparsers(dest='config_command', help='Config commands')

    # config show
    config_show = config_subparsers.add_parser('show', help='Show all settings')

    # config set
    config_set = config_subparsers.add_parser('set', help='Set a config value')
    config_set.add_argument(
        'key',
        type=str,
        choices=['send_mode', 'draft_instructions', 'humanizer_enabled'],
        help='Config key to set',
    )
    config_set.add_argument('value', type=str, help='Value to set')

    # config clear
    config_clear = config_subparsers.add_parser('clear', help='Reset a config key to default')
    config_clear.add_argument(
        'key',
        type=str,
        choices=['send_mode', 'draft_instructions', 'humanizer_enabled'],
        help='Config key to reset',
    )

    return parser


def _add_search_args(parser: argparse.ArgumentParser) -> None:
    """Add common search/filter arguments to a parser."""
    parser.add_argument(
        '--folder', '-F',
        type=str,
        action='append',
        help='Folder to search (can specify multiple). Default: Inbox + Sent Items',
    )
    parser.add_argument(
        '--days', '-d',
        type=int,
        default=DEFAULT_LOOKBACK_DAYS,
        help=f'Number of days to look back (default: {DEFAULT_LOOKBACK_DAYS})',
    )
    parser.add_argument(
        '--from-date',
        type=str,
        help='Start date (YYYY-MM-DD). Overrides --days.',
    )
    parser.add_argument(
        '--to-date',
        type=str,
        help='End date (YYYY-MM-DD). Default: now.',
    )
    parser.add_argument(
        '--unread', '-u',
        action='store_true',
        help='Only include unread messages',
    )
    parser.add_argument(
        '--filter-email', '-f',
        type=str,
        action='append',
        help='Filter by email address (From/To/CC). Can specify multiple.',
    )
    parser.add_argument(
        '--filter-domain', '-D',
        type=str,
        action='append',
        help='Filter by email domain. Can specify multiple.',
    )
    parser.add_argument(
        '--keyword', '-k',
        type=str,
        help='Search for keyword in subject/body',
    )


def _get_date_range(args) -> tuple[datetime | None, datetime | None]:
    """Parse date range from arguments."""
    since_date = None
    until_date = None

    if hasattr(args, 'from_date') and args.from_date:
        since_date = parse_date(args.from_date)
        if hasattr(args, 'to_date') and args.to_date:
            until_date = parse_date(args.to_date)
            until_date = until_date.replace(hour=23, minute=59, second=59)
    elif hasattr(args, 'days'):
        since_date = datetime.now() - timedelta(days=args.days)

    return since_date, until_date


def cmd_folders(args) -> int:
    """Handle 'folders' command."""
    from .core.folders import list_all_folders

    if not getattr(args, 'json', False):
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    folders = list_all_folders(namespace)

    if getattr(args, 'json', False):
        _output_json(_json_success({"folders": folders}))
    else:
        viewer = ViewerService()
        viewer.print_folders(namespace)
        print("\nUse: outlook search --folder <name>")

    return 0


def cmd_search(args) -> int:
    """Handle 'search' command."""
    json_mode = getattr(args, 'json', False)

    if not json_mode:
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    since_date, until_date = _get_date_range(args)

    if not json_mode:
        print(f"\nSearching emails...")
        if since_date:
            print(f"  From: {since_date.strftime('%Y-%m-%d')}")
        if until_date:
            print(f"  To: {until_date.strftime('%Y-%m-%d')}")
        if args.folder:
            print(f"  Folders: {', '.join(args.folder)}")
        if args.unread:
            print(f"  Unread only: Yes")
        if args.filter_email:
            print(f"  Filter emails: {', '.join(args.filter_email)}")
        if args.filter_domain:
            print(f"  Filter domains: {', '.join(args.filter_domain)}")
        if args.keyword:
            print(f"  Keyword: {args.keyword}")

    # Search
    search = SearchService(namespace)
    emails = search.search(
        folders=args.folder,
        since_date=since_date,
        until_date=until_date,
        unread_only=args.unread,
        filter_emails=args.filter_email,
        filter_domains=args.filter_domain,
        filter_keyword=args.keyword,
    )

    if json_mode:
        _output_json(_json_success({
            "count": len(emails),
            "emails": [e.to_dict(include_body=False) for e in emails],
        }))
        return 0

    print(f"\nFound {len(emails)} email(s)")

    # Display results
    if not args.no_view:
        viewer = ViewerService()
        viewer.print_email_table(emails)
        viewer.print_summary(emails)

    # Export if requested
    if args.export:
        from .services.export import ExportService
        print(f"\nExporting to {args.export}...")
        exporter = ExportService(args.export)
        result = exporter.export_emails(emails)
        print(f"  Created {result['files_created']} file(s)")

    return 0


def cmd_export(args) -> int:
    """Handle 'export' command."""
    json_mode = getattr(args, 'json', False)
    incremental = getattr(args, 'incremental', False)
    stdout_mode = getattr(args, 'stdout', False)

    # Stdout mode implies JSON format and suppresses other output
    quiet_mode = json_mode or stdout_mode

    if not quiet_mode:
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    # For incremental mode, check last run and override since_date
    from .services.export import ExportService
    exporter = ExportService(args.output) if not stdout_mode else None
    since_date, until_date = _get_date_range(args)

    if incremental and exporter:
        last_run = exporter.get_last_run()
        if last_run:
            since_date = last_run
            if not quiet_mode:
                print(f"Incremental mode: fetching emails since {last_run.strftime('%Y-%m-%d %H:%M')}")
        else:
            if not quiet_mode:
                print("Incremental mode: no previous run found, using default date range")

    if not quiet_mode:
        print(f"\nExporting emails to: {args.output}")
        if since_date:
            print(f"  From: {since_date.strftime('%Y-%m-%d %H:%M')}")
        if until_date:
            print(f"  To: {until_date.strftime('%Y-%m-%d')}")

    # Search
    search = SearchService(namespace)
    emails = search.search(
        folders=args.folder,
        since_date=since_date,
        until_date=until_date,
        unread_only=args.unread,
        filter_emails=args.filter_email,
        filter_domains=args.filter_domain,
        filter_keyword=args.keyword,
    )

    if not quiet_mode:
        print(f"Found {len(emails)} email(s)")

    # Stdout mode: output JSON directly to terminal
    if stdout_mode:
        from .services.export import ExportService as ES
        temp_exporter = ES(Path.cwd())  # Just need the methods, not the directory
        data = temp_exporter.to_json_data(emails, group_threads=not args.no_threads)
        _output_json(data)
        return 0

    if not emails:
        exporter.save_state()  # Always save state to track last export attempt
        if json_mode:
            _output_json(_json_success({
                "emails_found": 0,
                "files_created": 0,
                "files_skipped": 0,
                "output_directory": args.output,
                "incremental": incremental,
            }))
        else:
            print("No emails to export.")
        return 0

    # Export to files (always save state for tracking last export)
    result = exporter.export_emails(
        emails,
        group_threads=not args.no_threads,
        no_overwrite=args.no_overwrite,
        save_state=True,
        output_format=args.format,
        batch=args.batch,
    )

    if json_mode:
        _output_json(_json_success({
            "emails_found": len(emails),
            "files_created": result['files_created'],
            "files_skipped": result.get('files_skipped', 0),
            "output_directory": args.output,
            "format": args.format,
            "batch": args.batch,
            "incremental": incremental,
        }))
    else:
        print(f"\nExport complete:")
        print(f"  Format: {args.format}")
        print(f"  Files created: {result['files_created']}")
        if result['files_skipped'] > 0:
            print(f"  Files skipped: {result['files_skipped']}")
        print(f"  Output directory: {args.output}")
        if args.batch:
            print("  Mode: batch (all emails in single file)")
        if incremental:
            print("  State saved for next incremental run")

    return 0


def cmd_read(args) -> int:
    """Handle 'read' command."""
    json_mode = getattr(args, 'json', False)

    if not json_mode:
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    search = SearchService(namespace)

    emails = []
    not_found = []
    for message_id in args.message_ids:
        email = search.get_message_by_id(message_id)
        if email:
            emails.append(email)
        else:
            not_found.append(message_id)

    if json_mode:
        if not_found and not emails:
            _output_json(_json_error(f"Messages not found: {', '.join(not_found)}", "not_found"))
            return 1
        _output_json(_json_success({
            "count": len(emails),
            "emails": [e.to_dict(include_body=True) for e in emails],
            "not_found": not_found if not_found else None,
        }))
        return 0 if not not_found else 1

    viewer = ViewerService()
    for i, email in enumerate(emails):
        if i > 0:
            print("\n" + "=" * 80 + "\n")
        viewer.print_email_detail(email)

    if not_found:
        print(f"\nMessages not found: {len(not_found)}")
        for mid in not_found:
            print(f"  - {mid}")
        return 1

    return 0


def cmd_send(args) -> int:
    """Handle 'send' command."""
    json_mode = getattr(args, 'json', False)

    # Determine if we should send immediately
    send_immediately = False
    if getattr(args, 'send', False):
        allowed, error_msg = _check_send_allowed()
        if not allowed:
            if json_mode:
                _output_json(_json_error(error_msg, "send_not_allowed"))
            else:
                print(f"✗ {error_msg}")
            return 1
        send_immediately = True

    if not json_mode:
        from .core.config import ConfigManager
        for tag in ConfigManager().status_tags():
            print(tag)
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    # Parse recipients
    to = [e.strip() for e in args.to.split(',')]
    cc = [e.strip() for e in args.cc.split(',')] if args.cc else None
    bcc = [e.strip() for e in args.bcc.split(',')] if args.bcc else None

    # Send
    compose = ComposeService(namespace)
    success, message = compose.send_email(
        to=to,
        subject=args.subject,
        body=args.body,
        cc=cc,
        bcc=bcc,
        attachments=args.attach,
        html=args.html,
        send_immediately=send_immediately,
    )

    if json_mode:
        if success:
            _output_json(_json_success({"message": message, "draft": not send_immediately}))
        else:
            _output_json(_json_error(message, "send_failed"))
        return 0 if success else 1

    if success:
        print(f"✓ {message}")
        return 0
    else:
        print(f"✗ {message}")
        return 1


def cmd_reply(args) -> int:
    """Handle 'reply' command."""
    json_mode = getattr(args, 'json', False)

    # Determine if we should send immediately
    send_immediately = False
    if getattr(args, 'send', False):
        allowed, error_msg = _check_send_allowed()
        if not allowed:
            if json_mode:
                _output_json(_json_error(error_msg, "send_not_allowed"))
            else:
                print(f"✗ {error_msg}")
            return 1
        send_immediately = True

    if not json_mode:
        from .core.config import ConfigManager
        for tag in ConfigManager().status_tags():
            print(tag)
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    compose = ComposeService(namespace)
    success, message = compose.reply(
        message_id=args.message_id,
        body=args.body,
        reply_all=args.all,
        attachments=args.attach,
        html=args.html,
        send_immediately=send_immediately,
    )

    if json_mode:
        if success:
            _output_json(_json_success({"message": message, "reply_all": args.all, "draft": not send_immediately}))
        else:
            _output_json(_json_error(message, "reply_failed"))
        return 0 if success else 1

    if success:
        print(f"✓ {message}")
        return 0
    else:
        print(f"✗ {message}")
        return 1


def cmd_forward(args) -> int:
    """Handle 'forward' command."""
    json_mode = getattr(args, 'json', False)

    # Determine if we should send immediately
    send_immediately = False
    if getattr(args, 'send', False):
        allowed, error_msg = _check_send_allowed()
        if not allowed:
            if json_mode:
                _output_json(_json_error(error_msg, "send_not_allowed"))
            else:
                print(f"✗ {error_msg}")
            return 1
        send_immediately = True

    if not json_mode:
        from .core.config import ConfigManager
        for tag in ConfigManager().status_tags():
            print(tag)
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    # Parse recipients
    to = [e.strip() for e in args.to.split(',')]
    cc = [e.strip() for e in args.cc.split(',')] if args.cc else None
    bcc = [e.strip() for e in args.bcc.split(',')] if args.bcc else None

    compose = ComposeService(namespace)
    success, message = compose.forward(
        message_id=args.message_id,
        to=to,
        body=args.body,
        cc=cc,
        bcc=bcc,
        attachments=args.attach,
        html=args.html,
        send_immediately=send_immediately,
    )

    if json_mode:
        if success:
            _output_json(_json_success({"message": message, "to": to, "draft": not send_immediately}))
        else:
            _output_json(_json_error(message, "forward_failed"))
        return 0 if success else 1

    if success:
        print(f"✓ {message}")
        return 0
    else:
        print(f"✗ {message}")
        return 1


def cmd_config(args) -> int:
    """Handle 'config' command."""
    from .core.config import ConfigManager

    cfg = ConfigManager()

    if not args.config_command:
        data = cfg.show()
        print(json.dumps(data, indent=2))
        return 0

    if args.config_command == 'show':
        data = cfg.show()
        print(json.dumps(data, indent=2))
        return 0

    if args.config_command == 'set':
        key = args.key
        value: str | bool = args.value

        if key == 'humanizer_enabled':
            if value.lower() in ('true', '1', 'yes'):
                value = True
            elif value.lower() in ('false', '0', 'no'):
                value = False
            else:
                print(f"Invalid value for {key}: expected true/false, got '{value}'")
                return 1

        if key == 'send_mode':
            if value not in ('draft', 'send'):
                print(f"Invalid value for send_mode: expected 'draft' or 'send', got '{value}'")
                return 1

        cfg.set(key, value)
        print(f"✓ {key} = {json.dumps(value)}")
        return 0

    if args.config_command == 'clear':
        cfg.clear(args.key)
        print(f"✓ {args.key} reset to default")
        return 0

    return 0


def _parse_datetime(dt_str: str) -> datetime:
    """Parse datetime string in format YYYY-MM-DD HH:MM."""
    return datetime.strptime(dt_str, "%Y-%m-%d %H:%M")


def cmd_cal(args) -> int:
    """Handle 'cal' command."""
    if not args.cal_command:
        print("Usage: outlook cal {list,read,create,delete}")
        return 1

    json_mode = getattr(args, 'json', False)

    if not json_mode:
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    calendar = CalendarService(namespace)

    if args.cal_command == 'list':
        start_date = parse_date(args.start) if args.start else None
        end_date = parse_date(args.end) if args.end else None
        if end_date:
            end_date = end_date.replace(hour=23, minute=59, second=59)

        if not json_mode:
            print("\nSearching calendar events...")
            if start_date:
                print(f"  From: {start_date.strftime('%Y-%m-%d')}")
            if end_date:
                print(f"  To: {end_date.strftime('%Y-%m-%d')}")

        events = calendar.list_events(
            start_date=start_date,
            end_date=end_date,
            subject_filter=args.subject,
            location_filter=args.location,
            organizer_filter=args.organizer,
            all_day_only=args.all_day,
            recurring_only=args.recurring,
        )

        if json_mode:
            _output_json(_json_success({
                "count": len(events),
                "events": [e.to_dict() for e in events],
            }))
        else:
            print(f"\nFound {len(events)} event(s)\n")
            viewer = ViewerService()
            viewer.print_events_table(events)

    elif args.cal_command == 'read':
        event = calendar.get_event(args.event_id)
        if not event:
            if json_mode:
                _output_json(_json_error(f"Event not found: {args.event_id}", "not_found"))
            else:
                print(f"Event not found: {args.event_id}")
            return 1

        if json_mode:
            _output_json(_json_success({"event": event.to_dict()}))
        else:
            viewer = ViewerService()
            viewer.print_event_detail(event)

    elif args.cal_command == 'create':
        start = _parse_datetime(args.start)
        end = _parse_datetime(args.end)

        required = [e.strip() for e in args.required.split(',')] if args.required else None
        optional = [e.strip() for e in args.optional.split(',')] if args.optional else None

        reminder = None if args.no_reminder else (args.reminder if args.reminder else 15)

        success, message = calendar.create_event(
            subject=args.subject,
            start=start,
            end=end,
            location=args.location,
            body=args.body,
            required_attendees=required,
            optional_attendees=optional,
            reminder_minutes=reminder,
        )

        if json_mode:
            if success:
                _output_json(_json_success({"event_id": message, "subject": args.subject}))
            else:
                _output_json(_json_error(message, "create_failed"))
            return 0 if success else 1

        if success:
            print(f"✓ Event created (ID: {message})")
            return 0
        else:
            print(f"✗ {message}")
            return 1

    elif args.cal_command == 'delete':
        success, message = calendar.delete_event(args.event_id)

        if json_mode:
            if success:
                _output_json(_json_success({"message": message}))
            else:
                _output_json(_json_error(message, "delete_failed"))
            return 0 if success else 1

        if success:
            print(f"✓ {message}")
            return 0
        else:
            print(f"✗ {message}")
            return 1

    return 0


def _parse_date(date_str: str) -> datetime:
    """Parse date string in format YYYY-MM-DD."""
    return datetime.strptime(date_str, "%Y-%m-%d")


def cmd_tasks(args) -> int:
    """Handle 'tasks' command."""
    if not args.tasks_command:
        print("Usage: outlook tasks {list,read,create,complete,delete}")
        return 1

    json_mode = getattr(args, 'json', False)

    if not json_mode:
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    from .services.tasks import TaskService
    tasks_service = TaskService(namespace)

    if args.tasks_command == 'list':
        due_before = _parse_date(args.due_before) if getattr(args, 'due_before', None) else None
        due_after = _parse_date(args.due_after) if getattr(args, 'due_after', None) else None

        if not json_mode:
            print("\nSearching tasks...")

        tasks = tasks_service.list_tasks(
            status=args.status,
            include_completed=getattr(args, 'all', False),
            due_before=due_before,
            due_after=due_after,
            priority=args.priority,
            category=args.category,
        )

        if json_mode:
            _output_json(_json_success({
                "count": len(tasks),
                "tasks": [t.to_dict() for t in tasks],
            }))
        else:
            print(f"\nFound {len(tasks)} task(s)\n")
            viewer = ViewerService()
            viewer.print_tasks_table(tasks)

    elif args.tasks_command == 'read':
        task = tasks_service.get_task(args.task_id)
        if not task:
            if json_mode:
                _output_json(_json_error(f"Task not found: {args.task_id}", "not_found"))
            else:
                print(f"Task not found: {args.task_id}")
            return 1

        if json_mode:
            _output_json(_json_success({"task": task.to_dict()}))
        else:
            viewer = ViewerService()
            viewer.print_task_detail(task)

    elif args.tasks_command == 'create':
        due_date = _parse_date(args.due) if args.due else None
        start_date = _parse_date(args.start) if args.start else None
        reminder_date = _parse_datetime(args.reminder) if args.reminder else None
        categories = [args.category] if args.category else None

        success, message = tasks_service.create_task(
            subject=args.subject,
            due_date=due_date,
            start_date=start_date,
            priority=args.priority,
            body=args.body,
            categories=categories,
            reminder_date=reminder_date,
        )

        if json_mode:
            if success:
                _output_json(_json_success({"task_id": message, "subject": args.subject}))
            else:
                _output_json(_json_error(message, "create_failed"))
            return 0 if success else 1

        if success:
            print(f"✓ Task created (ID: {message})")
            return 0
        else:
            print(f"✗ {message}")
            return 1

    elif args.tasks_command == 'complete':
        success, message = tasks_service.complete_task(args.task_id)

        if json_mode:
            if success:
                _output_json(_json_success({"message": message, "task_id": args.task_id}))
            else:
                _output_json(_json_error(message, "complete_failed"))
            return 0 if success else 1

        if success:
            print(f"✓ {message}")
            return 0
        else:
            print(f"✗ {message}")
            return 1

    elif args.tasks_command == 'delete':
        success, message = tasks_service.delete_task(args.task_id)

        if json_mode:
            if success:
                _output_json(_json_success({"message": message}))
            else:
                _output_json(_json_error(message, "delete_failed"))
            return 0 if success else 1

        if success:
            print(f"✓ {message}")
            return 0
        else:
            print(f"✗ {message}")
            return 1

    return 0


def cmd_notes(args) -> int:
    """Handle 'notes' command."""
    if not args.notes_command:
        print("Usage: outlook notes {list,read,create,delete}")
        return 1

    json_mode = getattr(args, 'json', False)

    if not json_mode:
        print("Connecting to Outlook...")
    _, namespace = connect_to_outlook()

    from .services.notes import NotesService
    notes_service = NotesService(namespace)

    if args.notes_command == 'list':
        if not json_mode:
            print("\nSearching notes...")

        notes = notes_service.list_notes(
            color=args.color,
            category=args.category,
            keyword=args.keyword,
            limit=args.limit,
        )

        if json_mode:
            _output_json(_json_success({
                "count": len(notes),
                "notes": [n.to_dict() for n in notes],
            }))
        else:
            print(f"\nFound {len(notes)} note(s)\n")
            viewer = ViewerService()
            viewer.print_notes_table(notes)

    elif args.notes_command == 'read':
        note = notes_service.get_note(args.note_id)
        if not note:
            if json_mode:
                _output_json(_json_error(f"Note not found: {args.note_id}", "not_found"))
            else:
                print(f"Note not found: {args.note_id}")
            return 1

        if json_mode:
            _output_json(_json_success({"note": note.to_dict()}))
        else:
            viewer = ViewerService()
            viewer.print_note_detail(note)

    elif args.notes_command == 'create':
        categories = [args.category] if args.category else None

        success, message = notes_service.create_note(
            body=args.body,
            color=args.color,
            categories=categories,
        )

        if json_mode:
            if success:
                _output_json(_json_success({"note_id": message}))
            else:
                _output_json(_json_error(message, "create_failed"))
            return 0 if success else 1

        if success:
            print(f"Note created (ID: {message})")
            return 0
        else:
            print(f"Failed: {message}")
            return 1

    elif args.notes_command == 'delete':
        success, message = notes_service.delete_note(args.note_id)

        if json_mode:
            if success:
                _output_json(_json_success({"message": message}))
            else:
                _output_json(_json_error(message, "delete_failed"))
            return 0 if success else 1

        if success:
            print(f"{message}")
            return 0
        else:
            print(f"Failed: {message}")
            return 1

    return 0


def main() -> int:
    """Main entry point."""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 0

    try:
        if args.command == 'folders':
            return cmd_folders(args)
        elif args.command == 'search':
            return cmd_search(args)
        elif args.command == 'export':
            return cmd_export(args)
        elif args.command == 'read':
            return cmd_read(args)
        elif args.command == 'send':
            return cmd_send(args)
        elif args.command == 'reply':
            return cmd_reply(args)
        elif args.command == 'forward':
            return cmd_forward(args)
        elif args.command == 'config':
            return cmd_config(args)
        elif args.command == 'cal':
            return cmd_cal(args)
        elif args.command == 'tasks':
            return cmd_tasks(args)
        elif args.command == 'notes':
            return cmd_notes(args)
        else:
            parser.print_help()
            return 1
    except KeyboardInterrupt:
        print("\nAborted.")
        return 130
    except Exception as e:
        print(f"Error: {e}")
        return 1


if __name__ == '__main__':
    sys.exit(main())
