# Feature Reference

Internal behaviors, edge cases, and undocumented logic. Load this when you need to understand exactly how a feature works beyond the basic command reference.

---

## Search & Filter

### Default scope
If no `--folder` specified: **Inbox + Sent Items** only.

### Filter logic
- `--filter-email` matches From, To, and CC
- `--filter-domain` matches email domain portion
- `--filter-name` matches the sender display name (substring, case-insensitive) — use when you know the name but not the address
- Filters are additive: with several set, an email must match at least one
- `--keyword` does whole-word regex search on subject + text body (case-insensitive)

### Mass distribution guard
Emails with >20 recipients (To + CC) are auto-excluded when any email/domain filter is active. Prevents matching mailing lists.

### Auto-reply suppression
Subjects matching these are silently excluded: `Automatic reply:`, `Out of Office:`, `Auto:`, `Autoreply:`, `Auto-reply:`, `OOO:`, `Absence:`

### Date filtering
- `--days N` (default: 7) — uses `ReceivedTime` for inbox, `SentOn` for Sent Items
- `--from-date` / `--to-date` overrides `--days`

### Folder resolution
- Name match (e.g. `Inbox`, `Archive`)
- Path notation for nested: `Parent/Child`
- Multiple `--folder` flags supported

### Read by ID
Uses `GetItemFromID(EntryID)`. EntryIDs can change when emails are moved — re-search if a stored ID fails.

---

## Compose & Send

### Draft-only default
All commands create drafts. No flag needed. No `--draft` flag exists.

### Direct sending
Requires `send_mode: send` in config + `--send` flag. See [direct-send.md](direct-send.md).

### Reply behavior
- Default: reply to sender only
- `--all`: ReplyAll
- Body is prepended to original quoted text

### Forward behavior
- Requires `--to`
- Optional `--body` prepended to forwarded content
- Optional `--cc`, `--bcc`

### Attachments
- Multiple `--attach` flags supported
- Non-existent paths silently skipped

### HTML vs plain text
- HTML is the default: body set as `HTMLBody`, inserted inside `<body>` so it inherits the signature's font/style
- `--plain`: plain `Body` instead
- Plain-text bodies passed while HTML is active have `\n` converted to `<br>` so paragraphs still show as separate lines
- In reply/forward, body inserted before quoted content

---

## Export

### Formats
- `markdown` (default): Obsidian-flavored with YAML frontmatter
- `json`: Structured, AI-optimized

### Output modes
- File export: `--output DIR`
- Stdout: `--stdout` (JSON to terminal, no files)
- Search+export: `search --export DIR`

### Thread grouping (default)
Normalized subject groups emails into thread files. Disable with `--no-threads`.

### Batch JSON
`--format json --batch`: one timestamped file instead of one per thread.

### Content processing
- HTML → markdown via BeautifulSoup + markdownify
- `<script>`, `<style>`, `<meta>`, `<head>` stripped
- `<blockquote>` removed
- Reply separators stripped
- Text body used as-is

### Incremental export
State file at `${SKILL_DIR}/extraction_state.json` (not in output dir).
- First run: uses `--days` window
- Saves `{"last_run": "ISO timestamp"}`
- Subsequent runs: only fetch since last_run
- Reset: delete the state file

---

## Calendar

### Account scope
Calendar, Tasks, and Notes search **all stores** accessible in Outlook — no configuration needed. Email search defaults to the primary store but accepts path-based folder targeting (e.g., `--folder "AccountName/Inbox"`) to reach any account.

### Event creation
- Required: `--subject`, `--start`, `--end` (YYYY-MM-DD HH:MM)
- Attendees: `--required`, `--optional` (comma-separated)
- Invitations sent automatically with attendees
- Default reminder: 15 min; use `--no-reminder` to disable

### ⚠️ Multi-account creation limitation
`cal create` calls `outlook.CreateItem(1)` internally, which saves to Outlook's **default delivery store** (`namespace.GetDefaultFolder(9)`). This is often *not* the default send account — they are separate concepts in Outlook's COM model. In a multi-account profile, new events land in whichever mailbox is configured as the primary delivery location.

There is no `--store` flag yet. Tracked upstream: `ob-cheng/outlook-cli-2.0` issue #2.

**Workaround** — target a specific store's calendar by walking `namespace.Folders` and using `folder.Items.Add(1)` instead:

```python
import win32com.client as win32
from datetime import datetime

outlook = win32.Dispatch('Outlook.Application')
ns = outlook.GetNamespace("MAPI")

# Find the store by name
target = next((ns.Folders.Item(i) for i in range(1, ns.Folders.Count + 1)
               if "alcon" in ns.Folders.Item(i).Name.lower()), None)
# Find Calendar subfolder
def find_cal(f):
    return f if "calendar" in f.Name.lower() else next(
        (find_cal(f.Folders.Item(j)) for j in range(1, f.Folders.Count + 1)), None)
cal = next((find_cal(target.Folders.Item(j)) for j in range(1, target.Folders.Count + 1)), None)
appt = cal.Items.Add(1)
appt.Subject, appt.Start, appt.End = "Meeting", datetime(2026,6,11,11,0,0), datetime(2026,6,11,13,0,0)
appt.Save()
```

### Filters
`--subject`, `--location`, `--organizer`, `--all-day`, `--recurring`

---

## Tasks

### Statuses
`not_started`, `in_progress`, `completed` (via `tasks complete`), `waiting`, `deferred`

### Task complete
Sets status to `completed` and percent complete to 100.

### Filters
`--status`, `--all` (include completed), `--due-before`/`--due-after`, `--priority`, `--category`

---

## Notes

- Required: `--body`
- Optional: `--color` (blue/green/pink/yellow/white, default yellow), `--category`
- Filters: `--color`, `--category`, `--keyword`, `--limit`

---

## Multi-Account

| Command | Store |
|---------|-------|
| Email search/read/export | Default store (use `--folder "Account/Inbox"` to target a specific account) |
| Calendar, Tasks, Notes | **All stores** — every configured account is searched automatically |
