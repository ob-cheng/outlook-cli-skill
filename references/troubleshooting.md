# Troubleshooting Guide

Common issues and solutions for outlook-cli.

## Connection Issues

### "Could not connect to Outlook"

**Cause:** Outlook is not running or COM automation is blocked.

**Solutions:**
1. Ensure Outlook desktop app is running
2. Run as the same user that owns the Outlook profile
3. Check if Outlook is in safe mode (restart normally)

```bash
# Test connection
python "${SKILL_DIR}/outlook.py" folders --json
```

### "Access denied" or permission errors

**Cause:** Security software blocking COM automation.

**Solutions:**
1. Add Python to antivirus exceptions
2. Run from a trusted directory
3. Check Windows Security settings

---

## Search Issues

### No results returned

**Possible causes:**
- Date range too narrow
- Folder not specified correctly
- Filter too restrictive

```bash
# Broaden search
python "${SKILL_DIR}/outlook.py" search --days 30 --json

# Check available folders
python "${SKILL_DIR}/outlook.py" folders --json

# Search specific folder
python "${SKILL_DIR}/outlook.py" search --folder "Inbox" --days 30 --json
```

### Wrong folder searched

**Default behavior:** Searches Inbox + Sent Items if no folder specified.

```bash
# Search all folders explicitly
python "${SKILL_DIR}/outlook.py" search --folder "Inbox" --folder "Sent Items" --folder "Archive" --days 7

# List all available folders first
python "${SKILL_DIR}/outlook.py" folders
```

### Message ID not found

**Cause:** Message IDs (EntryIDs) can change if email is moved between folders.

**Solutions:**
1. Re-search to get current message ID
2. Search in all folders

```bash
# Search by subject/sender instead
python "${SKILL_DIR}/outlook.py" search --keyword "subject text" --filter-email sender@co.com --json
```

---

## Export Issues

### "Permission denied" writing files

**Cause:** Output directory doesn't exist or no write permission.

```bash
# Create directory first
mkdir -p ./exports

# Then export
python "${SKILL_DIR}/outlook.py" export --output ./exports --days 7
```

### Incremental export not working

**Cause:** State file missing or corrupted.

```bash
# Check state file exists (stored in the skill directory, not the export directory)
ls ${SKILL_DIR}/extraction_state.json

# Reset by deleting state file
rm ${SKILL_DIR}/extraction_state.json

# Run fresh export
python "${SKILL_DIR}/outlook.py" export --output ./exports --days 30 --incremental
```

### Large export taking too long

**Solutions:**
1. Narrow date range
2. Add filters
3. Use batch mode for JSON

```bash
# Narrow scope
python "${SKILL_DIR}/outlook.py" export --output ./exports \
  --days 7 \
  --filter-email important@co.com \
  --format json --batch
```

---

## Calendar Issues

### "No events found" or empty results

**Cause:** Date range may be too narrow, or calendar data is sparse. Calendar commands search all stores by default — multi-account is handled automatically.

If you're still seeing no results, widen the date range:
```bash
python "${SKILL_DIR}/outlook.py" cal list --start 2026-01-01 --end 2026-12-31 --json
```

### Event not created

**Possible causes:**
- Invalid datetime format
- Missing required fields
- Calendar permissions

```bash
# Correct datetime format: "YYYY-MM-DD HH:MM"
python "${SKILL_DIR}/outlook.py" cal create \
  --subject "Test Event" \
  --start "2026-06-10 14:00" \
  --end "2026-06-10 15:00" \
  --json
```

### Attendees not receiving invites

**Cause:** Event created but not sent.

**Note:** Events with attendees automatically send invitations when created. If not receiving:
1. Check email addresses are valid
2. Check Outlook is online (not in offline mode)
3. Verify attendees aren't blocked

### Can't delete recurring event

**Behavior:** Deleting a recurring event instance may only delete that occurrence.

**Solution:** Delete from Outlook UI if you need to delete entire series.

---

## Task Issues

### Task status not updating

**Valid statuses:**
- `not_started`
- `in_progress`
- `completed`
- `waiting`
- `deferred`

```bash
# Mark complete
python "${SKILL_DIR}/outlook.py" tasks complete <task-id> --json

# Verify
python "${SKILL_DIR}/outlook.py" tasks read <task-id> --json
```

### Tasks not appearing in list

**Cause:** Completed tasks hidden by default.

```bash
# Include completed tasks
python "${SKILL_DIR}/outlook.py" tasks list --all --json
```

---

## JSON Output Issues

### Parsing JSON output

**Use jq for filtering:**

```bash
# Get just email subjects
python "${SKILL_DIR}/outlook.py" search --json | jq '.emails[].subject'

# Filter by condition
python "${SKILL_DIR}/outlook.py" tasks list --json | jq '.tasks[] | select(.priority == "high")'

# Count results
python "${SKILL_DIR}/outlook.py" search --unread --json | jq '.count'
```

### Special characters in output

**Cause:** Email body contains special characters.

**Solutions:**
1. Use `--json` flag (properly escapes)
2. Pipe through encoding filter

```bash
# JSON output handles encoding
python "${SKILL_DIR}/outlook.py" read <id> --json | jq -r '.emails[0].body'
```

---

## Performance Tips

### Verifying the folder cache is working

The folder cache is the biggest performance optimization (16-35s → 0.45s). If `folders` is still slow, the cache may not be populated:

```bash
# Check if cache file exists
ls ~/.outlook-cli/folder-cache.json

# Force populate (one-time, takes 10-15s)
python outlook.py folders --refresh

# Verify subsequent calls are fast (should be < 1s)
time python outlook.py folders --json
```

Cache auto-invalidates when store topology changes (accounts added/removed) and refreshes every 24h. Use `--refresh` to force a full re-walk at any time.

### Slow searches

1. Narrow date range (`--days`)
2. Specify folder (`--folder`)
3. Add filters (`--filter-email`, `--keyword`)

### Memory issues with large exports

```bash
# Use stdout mode for streaming
python "${SKILL_DIR}/outlook.py" export --output . --stdout --days 7 | process-tool

# Or use incremental mode
python "${SKILL_DIR}/outlook.py" export --output ./exports --incremental
```

### Reduce API calls

```bash
# Batch operations where possible
python "${SKILL_DIR}/outlook.py" read <id1> <id2> <id3> --json  # One call for multiple

# Use search with export flag
python "${SKILL_DIR}/outlook.py" search --keyword "topic" --export ./output  # Search + export in one
```

---

## Error Codes Reference

| Code | Meaning | Solution |
|------|---------|----------|
| `not_found` | Message/event/task ID invalid | Re-search to get current ID |
| `send_failed` | Email send failed | Check recipients, Outlook online status |
| `reply_failed` | Reply failed | Verify original message exists |
| `forward_failed` | Forward failed | Check recipient address valid |
| `create_failed` | Create failed | Check required fields, permissions |
| `delete_failed` | Delete failed | Verify item exists, check permissions |
| `complete_failed` | Task complete failed | Verify task exists and not already complete |
| `error` | Generic error | Check error message for details |

---

## Getting Help

If issues persist:

1. Check Outlook is running and responsive
2. Verify Python environment: `python --version`
3. Test basic command: `python "${SKILL_DIR}/outlook.py" folders --json`
4. Check error message in JSON output
5. Review Outlook's own error dialogs
