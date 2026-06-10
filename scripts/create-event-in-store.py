#!/usr/bin/env python
"""
Create a calendar event in a specific Outlook store's Calendar folder.

Use when `cal create` puts the event in the wrong store (default store)
and you need to target a secondary account (Alcon, Exchange, etc.).

Usage:
  python scripts/create-event-in-store.py \\
    --store "tianen.cheng@alcon.com" \\
    --subject "Team Standup" \\
    --start "2026-06-11 11:00" \\
    --end "2026-06-11 13:00" \\
    [--no-reminder]
"""

import win32com.client as win32
import argparse
from datetime import datetime


def find_store(namespace, store_hint: str):
    """Find a store by name (substring match, case-insensitive)."""
    for i in range(1, namespace.Folders.Count + 1):
        s = namespace.Folders.Item(i)
        if store_hint.lower() in s.Name.lower():
            return s
    return None


def find_calendar(folder):
    """Recursively find Calendar subfolder."""
    if "calendar" in folder.Name.lower() or "kalender" in folder.Name.lower():
        return folder
    for i in range(1, folder.Folders.Count + 1):
        try:
            result = find_calendar(folder.Folders.Item(i))
            if result:
                return result
        except Exception:
            pass
    return None


def main():
    parser = argparse.ArgumentParser(description="Create event in specific Outlook store")
    parser.add_argument("--store", required=True, help="Store name substring (e.g., 'alcon')")
    parser.add_argument("--subject", required=True, help="Event subject")
    parser.add_argument("--start", required=True, help="Start datetime: YYYY-MM-DD HH:MM")
    parser.add_argument("--end", required=True, help="End datetime: YYYY-MM-DD HH:MM")
    parser.add_argument("--no-reminder", action="store_true", help="Disable reminder")
    parser.add_argument("--location", help="Event location")
    parser.add_argument("--body", help="Event description")
    parser.add_argument("--json", action="store_true", help="JSON output")
    args = parser.parse_args()

    outlook = win32.Dispatch('Outlook.Application')
    ns = outlook.GetNamespace("MAPI")

    store = find_store(ns, args.store)
    if not store:
        available = [ns.Folders.Item(i).Name for i in range(1, ns.Folders.Count + 1)]
        err = f"Store '{args.store}' not found. Available: {', '.join(available)}"
        if args.json:
            import json; print(json.dumps({"success": False, "error": err}))
        else:
            print(f"ERROR: {err}")
        return 1

    calendar_folder = find_calendar(store)
    if not calendar_folder:
        err = f"Calendar folder not found in store '{store.Name}'"
        if args.json:
            import json; print(json.dumps({"success": False, "error": err}))
        else:
            print(f"ERROR: {err}")
        return 1

    appt = calendar_folder.Items.Add(1)  # 1 = olAppointmentItem
    appt.Subject = args.subject
    appt.Start = datetime.strptime(args.start, "%Y-%m-%d %H:%M")
    appt.End = datetime.strptime(args.end, "%Y-%m-%d %H:%M")

    if args.location:
        appt.Location = args.location
    if args.body:
        appt.Body = args.body
    if args.no_reminder:
        appt.ReminderSet = False

    appt.Save()

    if args.json:
        import json
        print(json.dumps({
            "success": True,
            "subject": appt.Subject,
            "start": str(appt.Start),
            "end": str(appt.End),
            "store": store.Name,
            "entry_id": appt.EntryID,
        }))
    else:
        print(f"Created '{appt.Subject}' in {store.Name} Calendar")
        print(f"  {appt.Start} → {appt.End}")


if __name__ == "__main__":
    exit(main())
