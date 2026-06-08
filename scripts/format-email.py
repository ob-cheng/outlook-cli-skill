#!/usr/bin/env python3
"""Format email JSON for display or processing.

Usage:
    python format-email.py [--format summary|full|thread] < email.json
    python outlook.py read <id> --json | python format-email.py --format summary
    python outlook.py export --output . --stdout --days 7 | python format-email.py --format thread

Formats:
    summary - One line per email: date | sender | subject
    full    - Complete email with headers and body
    thread  - Grouped by conversation thread (for export --stdout output)

Handles both 'read --json' output (Email.to_dict) and 'export --stdout' output (_email_to_json).
"""

import argparse
import json
import sys
from datetime import datetime


def format_date(date_str: str) -> str:
    """Format ISO date to readable format."""
    if not date_str:
        return ""
    try:
        dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d %H:%M')
    except (ValueError, TypeError):
        return date_str


def _normalize_sender(email: dict) -> str:
    """Get sender name from either read or export format."""
    # read format: sender = "Name <email>" or sender_clean = "Name"
    # export format: from = "Name"
    if email.get('from'):
        return email['from']
    if email.get('sender_clean'):
        return email['sender_clean']
    if email.get('sender'):
        sender = email['sender']
        # Strip email portion if present: "Name <email>" -> "Name"
        if '<' in sender:
            sender = sender.split('<')[0].strip()
        return sender
    return 'Unknown'


def _normalize_sender_email(email: dict) -> str:
    """Get sender email from either format."""
    if email.get('from_email'):
        return email['from_email']
    if email.get('sender_smtp'):
        return email['sender_smtp']
    return ''


def _normalize_to(email: dict) -> list[str]:
    """Get recipients as a list from either format."""
    if 'to' in email:
        to_val = email['to']
        if isinstance(to_val, list):
            return to_val
        if isinstance(to_val, str) and to_val and to_val != 'N/A':
            # read format: "Name <email>; Name2 <email2>"
            return [a.strip() for a in to_val.split(';') if a.strip()]
    return []


def _normalize_cc(email: dict) -> list[str]:
    """Get CC as a list from either format."""
    if 'cc' in email:
        cc_val = email['cc']
        if isinstance(cc_val, list):
            return cc_val or []
        if isinstance(cc_val, str) and cc_val and cc_val != 'N/A':
            return [a.strip() for a in cc_val.split(';') if a.strip()]
    return []


def _normalize_body(email: dict) -> str:
    """Get body text from either format."""
    # export format uses 'body'; read format uses 'text_body'
    if email.get('body'):
        return email['body']
    if email.get('text_body'):
        return email['text_body']
    if email.get('html_body'):
        return email['html_body']
    return '(no body)'


def _normalize_has_attachments(email: dict) -> bool:
    """Check if email has attachments."""
    # export format doesn't have this field by default
    return email.get('has_attachments', False)


def format_summary(emails: list[dict]) -> str:
    """Format emails as one-line summaries."""
    lines = []
    for email in emails:
        date = format_date(email.get('date', ''))
        sender = _normalize_sender(email)
        subject = email.get('subject', '(no subject)')
        if len(subject) > 60:
            subject = subject[:57] + '...'
        lines.append(f"{date} | {sender:30} | {subject}")
    return '\n'.join(lines)


def format_full(emails: list[dict]) -> str:
    """Format emails with full details."""
    output = []
    for i, email in enumerate(emails):
        if i > 0:
            output.append('\n' + '=' * 80 + '\n')

        output.append(f"Subject: {email.get('subject', '(no subject)')}")
        sender = _normalize_sender(email)
        sender_email = _normalize_sender_email(email)
        if sender_email:
            output.append(f"From: {sender} <{sender_email}>")
        else:
            output.append(f"From: {sender}")

        to_list = _normalize_to(email)
        output.append(f"To: {', '.join(to_list) if to_list else 'N/A'}")

        cc_list = _normalize_cc(email)
        if cc_list:
            output.append(f"CC: {', '.join(cc_list)}")

        output.append(f"Date: {format_date(email.get('date', ''))}")

        if _normalize_has_attachments(email):
            output.append("Attachments: Yes")

        output.append('')
        output.append(_normalize_body(email))

    return '\n'.join(output)


def format_thread(data: dict) -> str:
    """Format as conversation threads (for export --stdout output)."""
    output = []

    # Handle export format with threads
    if 'threads' in data:
        for thread in data['threads']:
            subject = thread.get('subject', '(no subject)')
            participants = thread.get('participants', [])
            msg_count = thread.get('message_count', 0)
            date_start = format_date(thread.get('date_start', ''))
            date_end = format_date(thread.get('date_end', ''))

            output.append(f"Thread: {subject}")
            output.append(f"Participants: {', '.join(participants) if participants else 'N/A'}")
            if date_start:
                date_range = date_start
                if date_end and date_end != date_start:
                    date_range += f" → {date_end}"
                output.append(f"Date range: {date_range}")
            output.append(f"Messages: {msg_count}")
            output.append('-' * 40)

            for msg in thread.get('messages', []):
                date = format_date(msg.get('date', ''))
                sender = _normalize_sender(msg)
                output.append(f"\n[{date}] {sender}:")
                output.append(_normalize_body(msg))

            output.append('\n' + '=' * 80 + '\n')

    # Handle regular email list (e.g. from read --json with multiple emails)
    elif 'emails' in data:
        # Group by subject
        threads: dict = {}
        for email in data['emails']:
            key = email.get('subject', 'other')
            if key not in threads:
                threads[key] = []
            threads[key].append(email)

        for subject, msgs in threads.items():
            output.append(f"Thread: {subject}")
            output.append(f"Messages: {len(msgs)}")
            output.append('-' * 40)

            sorted_msgs = sorted(msgs, key=lambda x: x.get('date', ''))

            for msg in sorted_msgs:
                date = format_date(msg.get('date', ''))
                sender = _normalize_sender(msg)
                output.append(f"\n[{date}] {sender}:")
                output.append(_normalize_body(msg))

            output.append('\n' + '=' * 80 + '\n')

    return '\n'.join(output)


def main():
    parser = argparse.ArgumentParser(description='Format email JSON')
    parser.add_argument('--format', choices=['summary', 'full', 'thread'],
                        default='summary', help='Output format')
    args = parser.parse_args()

    # Read JSON from stdin
    try:
        data = json.load(sys.stdin)
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON input: {e}", file=sys.stderr)
        sys.exit(1)

    # Check for success wrapper (read/search commands)
    if isinstance(data, dict) and data.get('success') is False:
        print(f"Error: {data.get('error', 'Unknown error')}", file=sys.stderr)
        sys.exit(1)

    # Extract emails from various formats
    emails: list[dict] = []

    # Direct export output (no success wrapper) — has threads or emails at top level
    if 'threads' in data:
        if args.format == 'thread':
            print(format_thread(data))
            return
        # Flatten threads for summary/full
        for thread in data.get('threads', []):
            emails.extend(thread.get('messages', []))

    elif 'emails' in data:
        # Can be either direct export or read/search success wrapper
        emails = data.get('emails', [])
        # Also handle reading from export format with subjects as keys
        if not emails and isinstance(data.get('emails'), dict):
            for msgs in data['emails'].values():
                if isinstance(msgs, list):
                    emails.extend(msgs)

    # Check 'event' or 'events' for calendar (not email but handle gracefully)
    if not emails:
        if data.get('event'):
            print("Calendar event data — use cal list/read for calendar output")
            return
        if data.get('events'):
            print("Calendar events data — use cal list for calendar output")
            return
        if data.get('tasks'):
            print("Task data — use tasks list/read for task output")
            return

    if not emails:
        print("No emails found in input", file=sys.stderr)
        sys.exit(1)

    # Format output
    if args.format == 'summary':
        print(format_summary(emails))
    elif args.format == 'full':
        print(format_full(emails))
    elif args.format == 'thread':
        print(format_thread(data))


if __name__ == '__main__':
    main()
