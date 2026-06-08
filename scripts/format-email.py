#!/usr/bin/env python3
"""Format email JSON for display or processing.

Usage:
    python format-email.py [--format summary|full|thread] < email.json
    python outlook.py read <id> --json | python format-email.py --format summary

Formats:
    summary - One line per email: date | sender | subject
    full    - Complete email with headers and body
    thread  - Grouped by conversation thread
"""

import argparse
import json
import sys
from datetime import datetime


def format_date(date_str: str) -> str:
    """Format ISO date to readable format."""
    try:
        dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return dt.strftime('%Y-%m-%d %H:%M')
    except:
        return date_str


def format_summary(emails: list) -> str:
    """Format emails as one-line summaries."""
    lines = []
    for email in emails:
        date = format_date(email.get('date', ''))
        sender = email.get('sender', 'Unknown')
        subject = email.get('subject', '(no subject)')
        # Truncate long subjects
        if len(subject) > 60:
            subject = subject[:57] + '...'
        lines.append(f"{date} | {sender:30} | {subject}")
    return '\n'.join(lines)


def format_full(emails: list) -> str:
    """Format emails with full details."""
    output = []
    for i, email in enumerate(emails):
        if i > 0:
            output.append('\n' + '=' * 80 + '\n')

        output.append(f"Subject: {email.get('subject', '(no subject)')}")
        output.append(f"From: {email.get('sender_name', '')} <{email.get('sender', '')}>")
        output.append(f"To: {', '.join(email.get('to', []))}")
        if email.get('cc'):
            output.append(f"CC: {', '.join(email.get('cc', []))}")
        output.append(f"Date: {format_date(email.get('date', ''))}")

        if email.get('has_attachments'):
            attachments = email.get('attachments', [])
            if attachments:
                att_names = [a.get('name', 'unknown') for a in attachments]
                output.append(f"Attachments: {', '.join(att_names)}")
            else:
                output.append("Attachments: Yes")

        output.append('')
        body = email.get('body', '(no body)')
        output.append(body)

    return '\n'.join(output)


def format_thread(data: dict) -> str:
    """Format as conversation threads."""
    output = []

    # Handle export format with threads
    if 'threads' in data:
        for thread in data['threads']:
            output.append(f"Thread: {thread.get('subject', '(no subject)')}")
            output.append(f"Participants: {', '.join(thread.get('participants', []))}")
            output.append(f"Messages: {thread.get('message_count', 0)}")
            output.append('-' * 40)

            for msg in thread.get('messages', []):
                date = format_date(msg.get('date', ''))
                sender = msg.get('sender', 'Unknown')
                output.append(f"\n[{date}] {sender}:")
                output.append(msg.get('body', '(no body)'))

            output.append('\n' + '=' * 80 + '\n')

    # Handle regular email list
    elif 'emails' in data:
        # Group by conversation_id or subject
        threads = {}
        for email in data['emails']:
            key = email.get('conversation_id', email.get('subject', 'other'))
            if key not in threads:
                threads[key] = []
            threads[key].append(email)

        for subject, emails in threads.items():
            output.append(f"Thread: {subject}")
            output.append(f"Messages: {len(emails)}")
            output.append('-' * 40)

            for email in sorted(emails, key=lambda x: x.get('date', '')):
                date = format_date(email.get('date', ''))
                sender = email.get('sender', 'Unknown')
                output.append(f"\n[{date}] {sender}:")
                output.append(email.get('body', email.get('subject', '')))

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

    # Check for success
    if data.get('success') is False:
        print(f"Error: {data.get('error', 'Unknown error')}", file=sys.stderr)
        sys.exit(1)

    # Get emails from various formats
    emails = data.get('emails', [])
    if not emails and 'threads' in data:
        # Export format
        if args.format == 'thread':
            print(format_thread(data))
            return
        # Flatten threads for other formats
        for thread in data.get('threads', []):
            emails.extend(thread.get('messages', []))

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
