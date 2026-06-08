# Command Reference

Complete reference for all outlook-cli commands and options.

All commands use `${SKILL_DIR}/outlook.py` as the entry point.

## Email Commands

### search - Find emails

```bash
python "${SKILL_DIR}/outlook.py" search [options]
```

| Option | Description |
|--------|-------------|
| `--folder NAME` | Folder to search (can specify multiple) |
| `--days N` | Days to look back (default: 7) |
| `--from-date DATE` | Start date (YYYY-MM-DD), overrides --days |
| `--to-date DATE` | End date (YYYY-MM-DD) |
| `--unread` | Only unread messages |
| `--filter-email ADDRESS` | Filter by email address (can specify multiple) |
| `--filter-domain DOMAIN` | Filter by domain (can specify multiple) |
| `--keyword TEXT` | Search in subject/body |
| `--export DIR` | Also export results to markdown |
| `--no-view` | Skip terminal display |
| `--json` | Output as JSON |

### read - View email content

```bash
python outlook.py read <message-id> [<message-id>...] [--json]
```

Get message IDs from search results. Can read multiple emails at once.

### send - Compose and send email

```bash
python outlook.py send --to ADDRESS --subject TEXT --body TEXT [options]
```

> **Draft by default.** No flag needed — all emails are saved as drafts unless `--send` is passed.

| Option | Description |
|--------|-------------|
| `--to ADDRESS` | Recipient (required, comma-separated for multiple) |
| `--subject TEXT` | Subject line (required) |
| `--body TEXT` | Message body (required) |
| `--cc ADDRESS` | CC recipients (comma-separated) |
| `--bcc ADDRESS` | BCC recipients (comma-separated) |
| `--attach PATH` | File attachment (can specify multiple) |
| `--html` | Body is HTML formatted |
| `--send` | Send immediately (requires `send_mode: send` in config) |
| `--json` | Output as JSON |

### reply - Reply to email

```bash
python outlook.py reply <message-id> --body TEXT [options]
```

> **Draft by default.** Replies are saved as drafts unless `--send` is passed.

| Option | Description |
|--------|-------------|
| `--body TEXT` | Reply message (required) |
| `--all` | Reply to all recipients |
| `--attach PATH` | Add attachment (can specify multiple) |
| `--html` | Body is HTML |
| `--send` | Send immediately (requires `send_mode: send` in config) |
| `--json` | Output as JSON |

### forward - Forward email

```bash
python outlook.py forward <message-id> --to ADDRESS [options]
```

> **Draft by default.** Forwards are saved as drafts unless `--send` is passed.

| Option | Description |
|--------|-------------|
| `--to ADDRESS` | Forward recipient (required, comma-separated) |
| `--body TEXT` | Optional message to add |
| `--cc ADDRESS` | CC recipients |
| `--bcc ADDRESS` | BCC recipients |
| `--attach PATH` | Additional attachment |
| `--html` | Body is HTML |
| `--send` | Send immediately (requires `send_mode: send` in config) |
| `--json` | Output as JSON |

### export - Export emails to files or stdout

```bash
python outlook.py export --output DIR [options]
```

| Option | Description |
|--------|-------------|
| `--output DIR` | Output directory (required, use `.` with --stdout) |
| `--days N` | Days to look back (default: 7) |
| `--folder NAME` | Folder to export (can specify multiple) |
| `--filter-email ADDRESS` | Filter by participant |
| `--filter-domain DOMAIN` | Filter by domain |
| `--keyword TEXT` | Filter by keyword |
| `--format FORMAT` | Output format: `markdown` (default) or `json` |
| `--batch` | For JSON: combine all emails into single file |
| `--stdout` | Output JSON to terminal (no files written) |
| `--no-threads` | Export each email separately |
| `--no-overwrite` | Skip existing files |
| `--incremental` | Only export since last run |
| `--json` | Output summary as JSON |

### folders - List all folders

```bash
python outlook.py folders [--json]
```

---

## Calendar Commands

### cal list - List events

```bash
python outlook.py cal list [options]
```

| Option | Description |
|--------|-------------|
| `--start DATE` | Start date (YYYY-MM-DD, default: today) |
| `--end DATE` | End date (YYYY-MM-DD, default: 7 days) |
| `--subject TEXT` | Filter by subject |
| `--location TEXT` | Filter by location |
| `--organizer EMAIL` | Filter by organizer |
| `--all-day` | All-day events only |
| `--recurring` | Recurring events only |
| `--json` | Output as JSON |

### cal read - View event details

```bash
python outlook.py cal read <event-id> [--json]
```

### cal create - Create event

```bash
python outlook.py cal create --subject TEXT --start DATETIME --end DATETIME [options]
```

| Option | Description |
|--------|-------------|
| `--subject TEXT` | Event subject (required) |
| `--start DATETIME` | Start "YYYY-MM-DD HH:MM" (required) |
| `--end DATETIME` | End "YYYY-MM-DD HH:MM" (required) |
| `--location TEXT` | Location |
| `--body TEXT` | Description |
| `--required EMAILS` | Required attendees (comma-separated) |
| `--optional EMAILS` | Optional attendees |
| `--reminder N` | Reminder minutes (default: 15) |
| `--no-reminder` | No reminder |
| `--json` | Output as JSON |

### cal delete - Delete event

```bash
python outlook.py cal delete <event-id> [--json]
```

---

## Task Commands

### tasks list - List tasks

```bash
python outlook.py tasks list [options]
```

| Option | Description |
|--------|-------------|
| `--status STATUS` | Filter: `not_started`, `in_progress`, `completed`, `waiting`, `deferred` |
| `--all` | Include completed tasks |
| `--due-before DATE` | Tasks due before date (YYYY-MM-DD) |
| `--due-after DATE` | Tasks due after date |
| `--priority PRIORITY` | Filter: `low`, `normal`, `high` |
| `--category NAME` | Filter by category |
| `--json` | Output as JSON |

### tasks read - View task details

```bash
python outlook.py tasks read <task-id> [--json]
```

### tasks create - Create task

```bash
python outlook.py tasks create --subject TEXT [options]
```

| Option | Description |
|--------|-------------|
| `--subject TEXT` | Task subject (required) |
| `--due DATE` | Due date (YYYY-MM-DD) |
| `--start DATE` | Start date (YYYY-MM-DD) |
| `--priority PRIORITY` | Priority: `low`, `normal` (default), `high` |
| `--body TEXT` | Task description |
| `--category NAME` | Category name |
| `--reminder DATETIME` | Reminder (YYYY-MM-DD HH:MM) |
| `--json` | Output as JSON |

### tasks complete - Mark complete

```bash
python outlook.py tasks complete <task-id> [--json]
```

### tasks delete - Delete task

```bash
python outlook.py tasks delete <task-id> [--json]
```

---

## Notes Commands

### notes list - List notes

```bash
python outlook.py notes list [options]
```

| Option | Description |
|--------|-------------|
| `--color COLOR` | Filter: `blue`, `green`, `pink`, `yellow`, `white` |
| `--category NAME` | Filter by category |
| `--keyword TEXT` | Search in subject/body |
| `--limit N` | Limit results |
| `--json` | Output as JSON |

### notes read - View note

```bash
python outlook.py notes read <note-id> [--json]
```

### notes create - Create note

```bash
python outlook.py notes create --body TEXT [options]
```

| Option | Description |
|--------|-------------|
| `--body TEXT` | Note content (required) |
| `--color COLOR` | Color: `blue`, `green`, `pink`, `yellow` (default), `white` |
| `--category NAME` | Category name |
| `--json` | Output as JSON |

### notes delete - Delete note

```bash
python outlook.py notes delete <note-id> [--json]
```

---

## People Commands

### people list - List directory

```bash
python outlook.py people list [--json]
```

| Option | Description |
|--------|-------------|
| `--json` | Output as JSON |

### people lookup - Find person

```bash
python outlook.py people lookup <name-or-email>
```

### people add - Add to directory

```bash
python outlook.py people add "Full Name" email@domain.com
```

---

## Multi-Account Behavior

| Command | Scope |
|---------|-------|
| `search`, `export`, `read` | Default account (use `--folder "Account/Inbox"` to target another) |
| `cal`, `tasks`, `notes` | All accounts — every configured store is searched automatically |
