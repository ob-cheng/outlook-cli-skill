# JSON Output Schemas

All commands support `--json` for structured output. This reference documents the actual JSON schemas produced by the CLI.

## Response Wrapper

All commands that return JSON use this pattern:

**Success:**
```json
{
  "success": true,
  ...data fields
}
```

**Error:**
```json
{
  "success": false,
  "error": "Error message",
  "code": "error_code"
}
```

Common error codes: `not_found`, `send_failed`, `reply_failed`, `forward_failed`, `create_failed`, `delete_failed`, `complete_failed`, `send_not_allowed`

---

## Email Schemas

### search --json

Returns a summary (no body content):

```json
{
  "success": true,
  "count": 3,
  "emails": [
    {
      "message_id": "00000000ABC123...",
      "subject": "Meeting tomorrow",
      "sender": "Alice Smith <alice@company.com>",
      "sender_clean": "Alice Smith",
      "sender_smtp": "alice@company.com",
      "to": "bob@company.com",
      "to_emails": ["bob@company.com"],
      "cc": null,
      "cc_emails": [],
      "date": "2026-06-08T09:30:00",
      "is_sent": false,
      "is_read": false,
      "importance": "normal",
      "has_attachments": true
    }
  ]
}
```

### read --json

Returns full email with body content:

```json
{
  "success": true,
  "count": 1,
  "emails": [
    {
      "message_id": "00000000ABC123...",
      "subject": "Meeting tomorrow",
      "sender": "Alice Smith <alice@company.com>",
      "sender_clean": "Alice Smith",
      "sender_smtp": "alice@company.com",
      "to": "bob@company.com; carol@company.com",
      "to_emails": ["bob@company.com", "carol@company.com"],
      "cc": null,
      "cc_emails": [],
      "date": "2026-06-08T09:30:00",
      "is_sent": false,
      "is_read": true,
      "importance": "normal",
      "has_attachments": true,
      "text_body": "Hi Bob,\n\nLet's meet tomorrow at 2pm...",
      "html_body": "<html><body>Hi Bob,<br><br>Let's meet tomorrow at 2pm...</body></html>"
    }
  ],
  "not_found": null
}
```

### send/reply/forward --json

```json
{
  "success": true,
  "message": "Email saved to Drafts (ID: 00000000DRAFT123...)",
  "draft": true
}
```

When sent directly (`--send` with env var enabled):
```json
{
  "success": true,
  "message": "Email sent successfully",
  "draft": false
}
```

### export --json (summary)

Returns export operation summary (not email content):

```json
{
  "success": true,
  "emails_found": 15,
  "files_created": 12,
  "files_skipped": 3,
  "output_directory": "./exports",
  "format": "json",
  "batch": true,
  "incremental": false
}
```

### export --stdout (direct JSON to terminal)

Returns email data directly (no `success` wrapper). This is the raw output for agent/pipeline consumption:

```json
{
  "export_date": "2026-06-08T10:30:00",
  "thread_count": 3,
  "email_count": 15,
  "threads": [
    {
      "subject": "Project Update",
      "message_count": 3,
      "date_start": "2026-06-07T09:00:00",
      "date_end": "2026-06-08T14:00:00",
      "participants": ["Alice Smith", "Bob Jones"],
      "messages": [
        {
          "id": "00000000ABC123...",
          "subject": "Re: Project Update",
          "from": "Alice Smith",
          "from_email": "alice@company.com",
          "to": ["bob@company.com"],
          "cc": null,
          "date": "2026-06-08T14:00:00",
          "direction": "received",
          "body": "Updated status for the project..."
        }
      ]
    }
  ]
}
```

With `--no-threads`:
```json
{
  "export_date": "2026-06-08T10:30:00",
  "email_count": 15,
  "emails": [
    {
      "id": "00000000ABC123...",
      "subject": "Project Update",
      "from": "Alice Smith",
      "from_email": "alice@company.com",
      "to": ["bob@company.com"],
      "cc": null,
      "date": "2026-06-08T14:00:00",
      "direction": "received",
      "body": "Updated status for the project..."
    }
  ]
}
```

### folders --json

```json
{
  "success": true,
  "folders": [
    {
      "name": "user@example.com",
      "path": "user@example.com",
      "level": 0,
      "count": 0,
      "is_store": true
    },
    {
      "name": "Inbox",
      "path": "user@example.com/Inbox",
      "level": 1,
      "count": 150
    },
    {
      "name": "Sent Items",
      "path": "user@example.com/Sent Items",
      "level": 1,
      "count": 500
    }
  ]
}
```

**Note:** `count` is the total item count in the folder (not unread count). Store-level entries have `is_store: true`.

---

## Calendar Schemas

### cal list --json

```json
{
  "success": true,
  "count": 5,
  "events": [
    {
      "subject": "Team Standup",
      "start": "2026-06-08T09:00:00",
      "end": "2026-06-08T09:30:00",
      "location": "Room A",
      "body": "Weekly team sync",
      "organizer": "Alice Smith",
      "required_attendees": ["bob@company.com"],
      "optional_attendees": [],
      "is_all_day": false,
      "is_recurring": true,
      "recurrence_pattern": "Weekly",
      "entry_id": "00000000DEF456...",
      "categories": ["Work"],
      "importance": "normal",
      "reminder_minutes": 15
    }
  ]
}
```

### cal read --json

Same schema as list but for a single event, wrapped in `event`:

```json
{
  "success": true,
  "event": {
    "subject": "Team Standup",
    "start": "2026-06-08T09:00:00",
    "end": "2026-06-08T09:30:00",
    "location": "Room A",
    "body": "Weekly team sync meeting",
    "organizer": "Alice Smith",
    "required_attendees": ["bob@company.com", "carol@company.com"],
    "optional_attendees": [],
    "is_all_day": false,
    "is_recurring": true,
    "recurrence_pattern": "Weekly",
    "entry_id": "00000000DEF456...",
    "categories": ["Work"],
    "importance": "normal",
    "reminder_minutes": 15
  }
}
```

### cal create --json

```json
{
  "success": true,
  "event_id": "00000000NEW789...",
  "subject": "New Meeting"
}
```

### cal delete --json

```json
{
  "success": true,
  "message": "Event deleted successfully"
}
```

---

## Task Schemas

### tasks list --json

```json
{
  "success": true,
  "count": 8,
  "tasks": [
    {
      "entry_id": "00000000TASK123...",
      "subject": "Review PR #42",
      "status": "not_started",
      "due_date": "2026-06-15T00:00:00",
      "start_date": null,
      "completed_date": null,
      "percent_complete": 0,
      "priority": "high",
      "body": "Check the implementation details",
      "categories": ["Work"],
      "reminder_date": "2026-06-14T09:00:00"
    }
  ]
}
```

### tasks read --json

Same schema but single task wrapped in `task`:

```json
{
  "success": true,
  "task": {
    "entry_id": "00000000TASK123...",
    "subject": "Review PR #42",
    "status": "not_started",
    "due_date": "2026-06-15T00:00:00",
    "start_date": null,
    "completed_date": null,
    "percent_complete": 0,
    "priority": "high",
    "body": "Check the implementation details",
    "categories": ["Work"],
    "reminder_date": "2026-06-14T09:00:00"
  }
}
```

### tasks create --json

```json
{
  "success": true,
  "task_id": "00000000NEWTASK...",
  "subject": "Review PR #42"
}
```

### tasks complete --json

```json
{
  "success": true,
  "message": "Task marked as complete",
  "task_id": "00000000TASK123..."
}
```

### tasks delete --json

```json
{
  "success": true,
  "message": "Task deleted successfully"
}
```

---

## Notes Schemas

### notes list --json

```json
{
  "success": true,
  "count": 3,
  "notes": [
    {
      "entry_id": "00000000NOTE123...",
      "subject": "(No Subject)",
      "body": "Remember to follow up on Q2 report",
      "color": "yellow",
      "created": "2026-06-08T10:00:00",
      "modified": "2026-06-08T10:30:00",
      "categories": []
    }
  ]
}
```

### notes read --json

```json
{
  "success": true,
  "note": {
    "entry_id": "00000000NOTE123...",
    "subject": "(No Subject)",
    "body": "Remember to follow up on Q2 report",
    "color": "yellow",
    "created": "2026-06-08T10:00:00",
    "modified": "2026-06-08T10:30:00",
    "categories": []
  }
}
```

### notes create --json

```json
{
  "success": true,
  "note_id": "00000000NEWNOTE..."
}
```

### notes delete --json

```json
{
  "success": true,
  "message": "Note deleted successfully"
}
```
