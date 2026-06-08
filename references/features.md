# Feature Reference

Internal behaviors, edge cases, and undocumented logic. Load this when you need to understand exactly how a feature works beyond the basic command reference.

---

## Search & Filter

### Default scope
If no `--folder` specified: **Inbox + Sent Items** only.

### Filter logic
- `--filter-email` matches From, To, and CC
- `--filter-domain` matches email domain portion
- Filters are additive: with both set, an email must match at least one
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
Requires `OUTLOOK_CLI_ALLOW_SEND=1` + `--send` flag. See [direct-send.md](direct-send.md).

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
- `--html`: body set as `HTMLBody`
- Without: plain `Body`
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

### Default account
Uses Outlook's default account (File → Account Settings → Data Files). Multi-account setups may need to change default.

### Event creation
- Required: `--subject`, `--start`, `--end` (YYYY-MM-DD HH:MM)
- Attendees: `--required`, `--optional` (comma-separated)
- Invitations sent automatically with attendees
- Default reminder: 15 min; use `--no-reminder` to disable

### Filters
`--subject`, `--location`, `--organizer`, `--all-day`, `--recurring`

---

## Tasks

### Statuses
`not_started`, `in_progress`, `completed` (via `tasks complete`), `waiting`, `deferred`

### Task complete
Sets status to `completed` and percent complete to 100.

### Filters
`--status`, `--all` (include completed), `--due-before`/`--due-after`, `--priority`

---

## Notes

- Required: `--body`
- Optional: `--color` (blue/green/pink/yellow/white, default yellow), `--category`
- Filters: `--color`, `--category`, `--keyword`

---

## Multi-Account

| Command | Store |
|---------|-------|
| Email search/read/export | Default store (use `--folder` to target specific account) |
| Calendar, Tasks, Notes | Default account |
