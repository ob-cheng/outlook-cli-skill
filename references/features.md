# Feature Reference

Complete documentation of all behaviors, edge cases, and internal logic. Load this when you need to understand exactly how a feature works beyond the basic command reference.

---

## Search & Filter

### Default scope
If no `--folder` is specified, search covers **Inbox + Sent Items** only.

### Filter logic
- `--filter-email` matches against **From, To, and CC** (any match includes the email)
- `--filter-domain` matches email domain (e.g. `--filter-domain company.com` matches `alice@company.com`)
- Filters are **additive within type** but used together: if both `--filter-email` and `--filter-domain` are set, an email must match at least one filter
- `--keyword` searches subject + text body for whole-word matches (regex `\bkeyword\b`, case-insensitive)
- `--unread` filters by UnRead flag in Outlook

### Mass distribution guard
Emails with more than **20 recipients** (To + CC combined) are automatically excluded when any email/domain filter is active. This prevents matching mailing lists and mass distributions.

### Auto-reply suppression
Emails matching these subject patterns are **silently excluded**:
- `Automatic reply:`, `Out of Office:`, `Auto:`, `Autoreply:`, `Auto-reply:`, `OOO:`, `Absence:`

### Date filtering
- `--days N`: filter to last N days (default: 7). Uses `ReceivedTime` for inbox, `SentOn` for Sent Items
- `--from-date` / `--to-date`: explicit range, overrides `--days`

### Folder resolution
- Simple name match (e.g. `Inbox`, `Archive`)
- Path notation with `/` for nested folders (e.g. `Parent/Child`)
- Can specify `--folder` multiple times for cross-folder search

### Read by ID
Uses Outlook's `GetItemFromID(EntryID)`. **Note:** EntryIDs can change if an email is moved between folders. Re-search if a stored ID stops working.

---

## Compose & Send

### Draft-only by default
Every send/reply/forward command creates a draft. **No flag needed.** There is no `--draft` flag — drafts are always the default.

### Direct sending (--send)
Requires `OUTLOOK_CLI_ALLOW_SEND=1` env var + `--send` flag. Without the env var, `--send` fails with a clear error. Without the `--send` flag, the email is saved as a draft — even with the env var set. See [references/direct-send.md](direct-send.md) for setup.

### Reply behavior
- Default: reply to sender only
- `--all` flag: reply to all recipients (uses Outlook's ReplyAll)
- Body is **prepended** to the original message's quoted text

### Forward behavior
- Requires `--to` recipient(s)
- Optional `--body` is prepended to the forwarded content
- Optional `--cc` and `--bcc`

### Attachments
- `--attach` can be specified **multiple times** for multiple files
- Paths are resolved relative to the current working directory
- Non-existent paths are silently skipped

### HTML vs plain text
- `--html` flag: body is treated as HTML, set as `HTMLBody`
- Without `--html`: body is plain text, set as `Body`
- In reply/forward with `--html`, the body is inserted as HTML before the quoted content

---

## Export

### Formats

| Format | `--format` | Default | Notes |
|--------|-----------|---------|-------|
| Markdown | `markdown` | ✅ | Obsidian-flavored markdown with YAML frontmatter (title, date, tags, participants) |
| JSON | `json` | — | Structured JSON, optimized for AI ingestion |

### Output modes

| Mode | Flag | Behavior |
|------|------|----------|
| File export | `--output DIR` | Writes files to directory, returns summary |
| Stdout | `--stdout` | Emits JSON to stdout (for pipelines), no files written |
| Search + export | `search --export DIR` | Runs search, displays results, and exports matched emails |

### Thread grouping (default)
Emails are grouped by **normalized subject** (strips Re:/Fwd: prefixes). Each thread produces one file with all messages sorted chronologically. Disable with `--no-threads`.

### Batch JSON mode
With `--format json --batch`, all emails are combined into a single timestamped JSON file instead of one file per thread. More token-efficient for AI processing.

### Incremental export
See [Incremental Export lifecycle](#incremental-export-lifecycle) below.

### Content processing
- **HTML body**: Converted to markdown via BeautifulSoup + markdownify
  - `<script>`, `<style>`, `<meta>`, `<head>` stripped
  - `<blockquote>` removed (forwarded/replied content)
  - Reply separator patterns stripped (e.g. `--- From:... Subject:...`)
- **Text body**: Used as-is

### Obsidian markdown output
Each file includes:
```markdown
---
title: "Email Subject"
date: 2026-06-08
participants:
  - "Alice <alice@co.com>"
tags:
  - email/thread
  - email/received
---

> [!info] Thread Summary
> **3 message(s)** from [[2026-06-08]] to [[2026-06-10]]
> **Participants:** Alice, Bob, Carol

## Message 1 (RECEIVED) ^msg-1

**From:** Alice
**Date:** [[2026-06-08]] 14:30
#email/received

Body content...
```
AI processing hints are embedded in Obsidian `%%` comments:
```markdown
%%
ai-hints:
  thread_type: email_conversation
  message_count: 3
  direction: mixed
%%
```

### JSON output schema (per-email)

```json
{
  "id": "EntryID string",
  "subject": "Re: Project Update",
  "from": "Alice Johnson",
  "from_email": "alice@company.com",
  "to": ["bob@company.com"],
  "cc": ["carol@company.com"],
  "date": "2026-06-08T14:30:00",
  "direction": "received",
  "body": "Cleaned markdown body..."
}
```

Threaded JSON wraps emails in:
```json
{
  "subject": "Project Update",
  "message_count": 3,
  "date_start": "2026-06-08T10:00:00",
  "date_end": "2026-06-10T16:30:00",
  "participants": ["Alice", "Bob"],
  "messages": [ ...emails... ]
}
```

---

## Incremental Export Lifecycle

Tracks what's been exported so each run only fetches new emails.

### State file
`${SKILL_DIR}/extraction_state.json` — lives in the skill directory, **not** the output directory. This means:
- You can point `--output` at different directories per run
- All incremental runs share one state
- The state persists across output directory cleanups

### State content
```json
{"last_run": "2026-06-08T14:30:00.123456"}
```

### Run lifecycle

1. **First run** — no state file exists → falls back to `--days N` window (default 7 days)
2. **Post-export** — writes current timestamp to state file
3. **Subsequent runs** — reads `last_run`, overrides date filter to `[last_run → now]`
4. **Zero emails** — still saves the timestamp (prevents re-scanning the same window)

### Reset
Delete `${SKILL_DIR}/extraction_state.json` to force a full re-export.

### Compatibility
Also reads `last_extraction` key from older state file format.

### Limitations
- Does not work with `--stdout` mode (no state path)
- Date precision is second-level (ISO 8601), so emails received in the same second as `last_run` could be missed

---

## Calendar

### Default account
Calendar commands use Outlook's **default account** (File → Account Settings → Data Files → Set as Default). In multi-account setups, the default may not be the one with your calendar data.

### Event creation
- Required: `--subject`, `--start` (YYYY-MM-DD HH:MM), `--end` (YYYY-MM-DD HH:MM)
- Attendees use `--required` and `--optional` (comma-separated emails)
- Invitations are sent automatically when attendees are included
- Reminder defaults to 15 minutes; use `--no-reminder` to disable

### Event filtering
- Filter by `--subject` (substring), `--location` (substring), `--organizer` (email)
- Filter by `--all-day` or `--recurring`

---

## Tasks

### Status lifecycle
| Status | Description |
|--------|-------------|
| `not_started` | Default for new tasks |
| `in_progress` | Actively being worked on |
| `completed` | Finished (use `tasks complete` command) |
| `waiting` | Blocked or delegated |
| `deferred` | Postponed |

### Task completion
Use `tasks complete <task-id>` — this sets status to `completed` and marks percent complete to 100.

### Filtering
- `--status` filter accepts the statuses above
- `--all` includes completed tasks (hidden by default)
- `--due-before` / `--due-after` for date range filtering
- `--priority` filter: `low`, `normal`, `high`

---

## Notes

### Creation
- Required: `--body` (note content)
- Optional: `--color` (blue, green, pink, yellow, white; default: yellow), `--category`

### Filtering
- `--color`, `--category`, `--keyword` (searches body)

---

## Multi-Account Behavior

| Command | Store used |
|---------|-----------|
| Email search/read/export | **Default store** (primary mailbox). Use `--folder` to target a specific account's folders. |
| Calendar (`cal`) | Default account. If no data, set the right account as default in Outlook settings. |
| Tasks | Default account (same as calendar). |
| Notes | Default account (same as calendar). |

To fix multi-account issues: File → Account Settings → Data Files → select the account with your data → Set as Default → restart Outlook.
