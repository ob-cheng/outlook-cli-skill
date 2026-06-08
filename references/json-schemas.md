# JSON Output Schemas

All commands support `--json` for structured output. This reference documents the JSON schemas.

## Response Wrapper

All responses follow this pattern:

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

Common error codes: `not_found`, `send_failed`, `reply_failed`, `forward_failed`, `create_failed`, `delete_failed`, `complete_failed`

---

## Email Schemas

### search --json

```json
{
  "success": true,
  "count": 3,
  "emails": [
    {
      "message_id": "00000000ABC123...",
      "subject": "Meeting tomorrow",
      "sender": "alice@company.com",
      "sender_name": "Alice Smith",
      "to": ["bob@company.com"],
      "cc": [],
      "date": "2026-06-08T09:30:00",
      "is_read": false,
      "has_attachments": true,
      "folder": "Inbox"
    }
  ]
}
```

### read --json

```json
{
  "success": true,
  "count": 1,
  "emails": [
    {
      "message_id": "00000000ABC123...",
      "subject": "Meeting tomorrow",
      "sender": "alice@company.com",
      "sender_name": "Alice Smith",
      "to": ["bob@company.com"],
      "cc": [],
      "date": "2026-06-08T09:30:00",
      "is_read": true,
      "has_attachments": true,
      "folder": "Inbox",
      "body": "Full email body text here...",
      "body_html": "<html>...</html>",
      "attachments": [
        {
          "name": "document.pdf",
          "size": 102400
        }
      ]
    }
  ],
  "not_found": null
}
```

### send/reply/forward --json

```json
{
  "success": true,
  "message": "Email sent successfully",
  "draft": false
}
```

### export --json (summary)

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

### export --stdout (direct JSON)

```json
{
  "exported_at": "2026-06-08T10:30:00",
  "count": 15,
  "threads": [
    {
      "thread_id": "thread_abc123",
      "subject": "Project Update",
      "participants": ["alice@co.com", "bob@co.com"],
      "message_count": 3,
      "messages": [
        {
          "message_id": "...",
          "subject": "Project Update",
          "sender": "alice@co.com",
          "date": "2026-06-08T09:00:00",
          "body": "..."
        }
      ]
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
      "name": "Inbox",
      "path": "\\\\Inbox",
      "unread_count": 5,
      "total_count": 150
    },
    {
      "name": "Sent Items",
      "path": "\\\\Sent Items",
      "unread_count": 0,
      "total_count": 500
    }
  ]
}
```

---

## Calendar Schemas

### cal list --json

```json
{
  "success": true,
  "count": 5,
  "events": [
    {
      "event_id": "00000000DEF456...",
      "subject": "Team Standup",
      "start": "2026-06-08T09:00:00",
      "end": "2026-06-08T09:30:00",
      "location": "Room A",
      "organizer": "alice@company.com",
      "is_all_day": false,
      "is_recurring": true,
      "response_status": "accepted"
    }
  ]
}
```

### cal read --json

```json
{
  "success": true,
  "event": {
    "event_id": "00000000DEF456...",
    "subject": "Team Standup",
    "start": "2026-06-08T09:00:00",
    "end": "2026-06-08T09:30:00",
    "location": "Room A",
    "body": "Weekly team sync meeting",
    "organizer": "alice@company.com",
    "required_attendees": ["bob@company.com", "carol@company.com"],
    "optional_attendees": [],
    "is_all_day": false,
    "is_recurring": true,
    "recurrence_pattern": "Weekly",
    "reminder_minutes": 15,
    "response_status": "accepted"
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
  "message": "Event deleted"
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
      "task_id": "00000000TASK123...",
      "subject": "Review PR #42",
      "status": "not_started",
      "priority": "high",
      "due_date": "2026-06-15",
      "start_date": null,
      "percent_complete": 0,
      "categories": ["Work"]
    }
  ]
}
```

### tasks read --json

```json
{
  "success": true,
  "task": {
    "task_id": "00000000TASK123...",
    "subject": "Review PR #42",
    "body": "Check the implementation details",
    "status": "not_started",
    "priority": "high",
    "due_date": "2026-06-15",
    "start_date": null,
    "date_completed": null,
    "percent_complete": 0,
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
  "message": "Task deleted"
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
      "note_id": "00000000NOTE123...",
      "subject": "Quick note",
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
    "note_id": "00000000NOTE123...",
    "subject": "Quick note",
    "body": "Remember to follow up on...",
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
  "message": "Note deleted"
}
```
