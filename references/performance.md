# Outlook CLI Performance Characteristics

Benchmarked 2026-06-10, WSL → Windows Python 3.13, Outlook 365 with 3 accounts (iCloud, alcon.com, outlook.com). All times are `real` from `time`.

## Component Costs

| Component | Time | Notes |
|-----------|------|-------|
| Python import (`from outlook_cli.cli import main`) | 0.34s | Pure import, no COM |
| COM `Dispatch('Outlook.Application')` | 0.22s | Lazy — doesn't actually connect until used |
| Combined cold-start per command | **~0.5s** | Python + import + argparse |

## Before → After (2026-06-10 optimizations)

| Command | Before | After | Speedup | Fix |
|---------|--------|-------|---------|-----|
| `folders` | 16-35s | **0.45s** | 35-75× | Folder tree cache |
| `cal list` (28 events) | 6.0s | **0.73s** | 8× | Summary mode (skip body/attendees/recurrence) |
| `tasks list` | 1.43s | 1.24s | 15% | Summary mode (skip body) |
| `search` (0 results) | 0.48s | 0.42s | — | Already fast |
| Batch (search+tasks+cal) | 2.64s | **1.60s** | 39% | Single-process batch dispatch |

## Key Findings

### 1. Cold-start is NOT the bottleneck
The actual per-command startup is ~0.5s. The claim of "5-10 second cold-start tax" would require 10-20 commands in sequence, which isn't realistic. A daemon would save very little.

### 2. The real bottlenecks are COM data iteration
- `folders`: 16-35s walking the namespace across 3 accounts. Gets WORSE on warm runs — COM state degradation.
- `cal list`: 6s for 28 events = ~200ms per calendar item. COM property access per item is expensive.

### 3. Optimizations implemented (shipped to all users)

**Folder tree cache** (`outlook_cli/core/folder_cache.py`):
- Serializes namespace walk to `~/.outlook-cli/folder-cache.json`
- Fingerprint invalidation: store names + count (cheap — no walk needed)
- 24h auto-refresh even if fingerprint matches
- `--refresh` flag forces full walk + cache repopulation

**COM summary mode** (calendar + tasks):
- `from_appointment(summary_only=True)` skips Body, Recipients iteration, GetRecurrencePattern, ReminderMinutesBeforeStart
- `from_task_item(summary_only=True)` skips Body
- `cal list` / `tasks list` use summary mode automatically
- `cal read` / `tasks read` use full mode (all properties)

**Batch subcommand**:
- `outlook.py batch --commands '[[cmd, ...args], ...]'`
- Runs all commands in a single Python process
- Saves ~0.34s per command after the first (Python startup)
- Returns JSON envelope with per-command results

## Test Methodology

```bash
PY="/mnt/c/Users/its_t/AppData/Local/Programs/Python/Python313/python.exe"
SKILL="/home/its_t/.hermes/skills/email/outlook-cli-skill"

# Single command
time "$PY" "$SKILL/outlook.py" cal list --json > /dev/null 2>&1

# Batch
time "$PY" "$SKILL/outlook.py" batch --commands '[
  ["search", "--unread", "--days", "1", "--limit", "5", "--json"],
  ["tasks", "list", "--json"],
  ["cal", "list", "--json"]
]' > /dev/null 2>&1
```
