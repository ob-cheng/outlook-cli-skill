---
name: outlook-cli
description: "Outlook email, calendar, tasks, and notes via CLI. Use when: checking inbox, finding/sending/replying emails, scheduling meetings, checking calendar, managing tasks/notes, exporting emails to markdown/JSON for AI processing. Supports --stdout for direct JSON output to AI agents. Windows-only (COM automation)."
version: 0.4.0
author: ob-cheng
license: MIT
# Platform restriction removed — the skill's description communicates the real requirement (Windows+COM).
# WSL users: follow docs/wsl.md to set up the Windows Python bridge.
# Categorization (Hermes discovery)
domain: email
tags:
  - email
  - calendar
  - outlook
  - tasks
  - notes
  - office
# Activation rules (works with Claude Code, Hermes, OpenClaw, and other agents)
when_to_use: |
  Use when:
  - User asks to check email, find messages, send/reply/forward
  - User wants to export emails to markdown/JSON for AI processing
  - User asks about calendar, schedule meetings, or check availability
  - User mentions tasks, todos, or notes in Outlook
  - Trigger phrases: "check my inbox", "emails from X", "send email to", "schedule meeting", "my calendar", "create task", "my todos", "export emails"
  
  Do NOT use for:
  - Web-based email (Gmail, Office 365 web) - this is COM only
  - macOS or Linux - Windows only
  - Email providers other than Outlook
# Tool permissions (agent-specific, most agents auto-detect)
allowed-tools:
  - Bash
  - Read
  - shell
  - execute
argument-hint: "an email action like 'find unread emails' or 'send email to alice@co.com'"
# Metadata
metadata:
  author: ob-cheng
  last-updated: "2026-06-11"
# Reference documents (on-demand loading)
references:
  - docs/install.md
  - references/commands.md
  - references/workflows.md
  - references/json-schemas.md
  - references/troubleshooting.md
  - docs/wsl.md
  - references/direct-send.md
  - docs/scripts.md
  - references/features.md
  - docs/structure.md
  - references/config.md
  - references/agent-ergonomics.md
  - references/performance.md
  # - testing.md omitted — dev-only, not user-facing
# Hermes-specific config (other agents use instructions in Safety Rules section)
metadata.hermes:
  config:
    - key: outlook.draft_only_mode
      help: "Draft-only mode — when enabled, emails are saved as drafts instead of being sent"
      default: true
      prompt: Enable draft-only mode for email safety?
---

# Outlook CLI Skill

AI-friendly CLI for Microsoft Outlook. Works via COM automation - no Azure setup, no OAuth, no API keys.

## Strategy

**Search -> Read -> Act**. Find emails first, then read details, then reply/forward/export.

Always add `--json` for structured output when processing programmatically.

> **Installation & setup:** See [docs/install.md](docs/install.md) for agent instructions on installing this skill and enabling direct sending.

## Updating This Skill

There are **two repos** — don't confuse them:

| Repo | Purpose | Location |
|------|---------|----------|
| **outlook-cli-skill** | Hermes skill wrapper (this file) | `~/.hermes/skills/email/outlook-cli-skill` — `github.com/ob-cheng/outlook-cli-skill` |
| **outlook-cli-2.0** | The actual CLI tool binary | `/mnt/c/Users/its_t/AppData/Local/outlook-cli-2.0` — `github.com/ob-cheng/outlook-cli-2.0` |

**To update the Hermes skill** (this SKILL.md + reference files):

```bash
cd ~/.hermes/skills/email/outlook-cli-skill
git pull
```

If `git pull` fails due to local modifications:
```bash
git stash && git pull && git stash pop
```

If `git stash pop` produces merge conflicts, resolve them manually — edit the conflicted file, remove conflict markers, `git add` it.

> **CRITICAL:** Never `git commit --amend` after a commit has been pushed. It rewrites history and diverges from origin. Always create a new fix commit instead.

**To update the CLI tool** (only needed when new features were released):

```bash
cd /mnt/c/Users/its_t/AppData/Local/outlook-cli-2.0
git pull
pip install -r requirements.txt
python outlook.py --help > /dev/null && echo "CLI OK"
```

Configuration (`~/.outlook-cli/`) lives outside both repos — updates never touch settings or people data.

> **WSL / non-Windows setup:** See [docs/wsl.md](docs/wsl.md) for the `OUTLOOK_CLI_PYTHON` workaround.
>
> **Feature deep-dive:** See [references/features.md](references/features.md) for internal behavior, export lifecycle, content processing, multi-account handling, and undocumented logic.
>
> **Structure & efficiency:** See [docs/structure.md](docs/structure.md) for directory layout and token efficiency.

## Quick Reference

Run all commands using `${OUTLOOK_CLI_PYTHON:-python}` (set only in WSL; falls back to `python` on native Windows):

| Intent | Command |
|--------|---------|
| Find emails | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" search [options]` |
| Read email | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" read <id> --json` |
| Send email | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z` |
| Reply | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" reply <id> --body "text" [--all] [--cc X]` |
| Forward | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" forward <id> --to X` |
| Calendar | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" cal list/read/create/delete` |
| Tasks | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" tasks list/read/create/complete/delete` |
| Notes | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" notes list/read/create/delete` |
| Export | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" export --output DIR [--format json] [--batch] [--stdout]` |
| Folders | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" folders [--json] [--refresh]` |
| People | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" people list/lookup/add` |
| Config | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" config show/set/clear` |
| Batch | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" batch --commands '[...]'` |

## Draft Workflow

Before every send/reply/forward:

1. Run `outlook.py config show` to read all settings (see [references/config.md](references/config.md))
2. Compose the email body
3. If `draft_instructions` is set → follow them while drafting
4. If `humanizer_enabled` is true → load `humanizer` skill and run the pattern checklist
5. Pass the final body to the CLI

The CLI prints status tags so skipped steps are visible in the output.

## People Directory

The CLI auto-tracks every person encountered in email interactions — sender, recipients, CC — via `python outlook.py read`. Unknown people are automatically added to `~/.outlook-cli/people.json`.

**After every email interaction** (any time you have an email in context — search results, read output, a composed send/reply/forward), scan the participants and run:

```bash
# Check if this person is already known
python outlook.py people lookup "Name"
python outlook.py people lookup email@domain.com

# If not found, add them (but only for people you haven't already handled via cmd_read)
python outlook.py people add "Full Name" email@domain.com
```

**The `read` command handles this automatically** — unknown participants are added and reported. For send/reply/forward, the agent should manually check since those involve the recipients you're sending to.

Whenever the user refers to someone by name (e.g. "email Alice about the report"), look them up in the directory:

```bash
python outlook.py people lookup "Alice"
```

If found, you have their email. If not found, ask the user for the email and save it.

To view all known people:

```bash
python outlook.py people list [--json]
```

## Common Patterns

These cover 80% of agent tasks without needing extra reference files.

### Quick inbox scan
```bash
# First N unread, last day only (fast on large inboxes)
python outlook.py search --unread --days 1 --limit 10 --json
```

### Find emails from someone
```bash
python outlook.py search --filter-email "alice@co.com" --days 7 --json
```

### Find emails from someone (name only, unknown email)
```bash
# --filter-email matches SMTP addresses only. When you only know the name:
# Step 1: discover accounts
python outlook.py folders --json
# Step 2: search broadly in the right account's inbox, then inspect sender_clean
python outlook.py search --folder "work@domain.com/Inbox" --days 7 --json
# Step 3: or narrow by domain when you know the organization
python outlook.py search --folder "work@domain.com/Inbox" --days 7 --filter-domain "alcon.com" --json
```

### Reply with extra CC
```bash
python outlook.py reply <id> --body "My reply" --cc "newcomer@co.com,support@co.com"
```

### Export thread to markdown
```bash
python outlook.py export --output ./inbox-export --filter-email "client@co.com" --days 7
```

### Export to JSON for AI processing
```bash
# Single JSON file, token-efficient
python outlook.py export --output ./data --format json --batch --days 30

# Direct JSON to stdout (no files, best for AI pipelines)
python outlook.py export --output . --stdout --days 7
```

### Task management
```bash
# List pending tasks
python outlook.py tasks list --json

# Create with due date and priority
python outlook.py tasks create --subject "Review PR" --due 2026-05-15 --priority high

# Mark complete
python outlook.py tasks complete <task-id>
```

### Calendar today
```bash
python outlook.py cal list --json
```

### Multi-account: always discover folders first
```bash
# See what accounts are connected (cached after first run — instant thereafter)
python outlook.py folders
# Force a full refresh if you've added/removed accounts
python outlook.py folders --refresh
# Then target a specific account
python outlook.py search --folder "work@domain.com/Inbox" --filter-email "sender@co.com"
```

### Batch mode (multi-command single process)
```bash
# Run multiple commands in one Python process — avoids 0.34s cold-start per extra command.
# Each inner array is [command, ...args]. Output is a JSON envelope with per-command results.
python outlook.py batch --commands '[
  ["search", "--unread", "--days", "1", "--limit", "5", "--json"],
  ["tasks", "list", "--json"],
  ["cal", "list", "--json"]
]'
```
Prefer batch when you know the full command pipeline upfront (e.g., search → read → reply). Saves ~40% on 3-command workflows vs. separate invocations.

### Date filtering options
- `--days N` — last N days (default 7)
- `--from-date YYYY-MM-DD --to-date YYYY-MM-DD` — exact range (overrides --days)
- `--limit N` — stop after N matches (for huge inboxes)

## Known Pitfalls

### Large inbox search can be slow
COM iterates every item in the folder. On 1000+ emails, narrow with `--days 1` or `--limit 20`. The CLI now shows a progress indicator in non-JSON mode.

### Per-command cold-start (~0.5s) and what's fast/slow

Every CLI invocation pays ~0.5s overhead (Python import: 0.34s + COM Dispatch: 0.22s). Three optimizations are baked in to reduce this and the larger COM data-iteration bottlenecks:

- **Folder cache** (`~/.outlook-cli/folder-cache.json`): First `folders` call walks all stores (16-35s). Subsequent calls return from cache instantly (0.45s). Invalidates automatically when store topology changes. `--refresh` forces a full re-walk.
- **Calendar/tasks summary mode**: `cal list` and `tasks list` skip expensive COM properties (body, attendees, recurrence details) automatically. `cal list`: 6s → 0.73s. `cal read` still returns full detail.
- **Batch mode**: Run multiple commands in one Python process. Cuts the 0.5s per-extra-command cold start. 3-command workflows run ~40% faster. Prefer batch when you know the full pipeline upfront (e.g., search → read → reply).

### cal create always uses the default delivery store in multi-account profiles

`cal create` calls `outlook.CreateItem(1)` which saves to Outlook's **default delivery store** — not necessarily the account the user considers primary. The default send account and default delivery store are separate concepts in COM. No `--store` flag exists yet. See `references/features.md` for workaround. Tracked: `ob-cheng/outlook-cli-2.0` #2.

### Multi-account: search defaults to primary account
Search/send/export target the default account. Run `folders` first to see what's connected, then use `--folder "AccountName/Inbox"` to reach another account.

### EntryIDs can change when emails move
If a stored message ID fails to load, re-search to get the current ID — Outlook regenerates EntryIDs on move/archive operations.

### Draft-only is the default
All compose commands create drafts. Direct sending requires both `send_mode: send` in config AND the `--send` flag.

### git stash pop can produce merge conflicts
If both your local modifications AND the upstream changed the same files (common for SKILL.md during feature updates), `git stash pop` after `git pull` will produce conflicts. Resolve them manually (edit conflicted file, remove markers, `git add` it) instead of trying to force through.

### Windows Python missing dependencies (first WSL setup)
First run from WSL often fails because the Windows Python doesn't have the tool's dependencies installed (`rich`, `markdownify`, `beautifulsoup4`, `pywin32`). The auto-detection finds the Python binary fine, then crashes with `ModuleNotFoundError`. Fix — install deps via the Windows Python that `OUTLOOK_CLI_PYTHON` points to:

```bash
"${OUTLOOK_CLI_PYTHON}" -m pip install -r "${SKILL_DIR}/requirements.txt"
```

For example, if auto-detected Python313:
```bash
"/mnt/c/Users/its_t/AppData/Local/Programs/Python/Python313/python.exe" -m pip install rich markdownify beautifulsoup4 pywin32
```

This is a one-time setup step. After install, all commands work.

### Testing from WSL (no Outlook COM) — developers only
When making code changes from WSL, COM won't work. The `tests/` directory has pytest suites that run without COM. See `references/testing.md` in the repo for mock patterns and CI setup.

## Safety Rules

**Draft-Only Mode:** The CLI defaults to draft mode. All send/reply/forward commands create drafts by default.

- Tell the user: "I've saved this as a draft. You can review and send it from Outlook."
- To enable direct sending, set config `send_mode: send` — see [references/config.md](references/config.md) for the behavior matrix
- **`--send` is the only way to direct-send.** Without it, even with `send_mode: send`, the email saves as a draft

**Always confirm before:** deleting events/tasks/notes
**Never:** delete without confirmation, forward sensitive emails without verification

> **Workflows & patterns:** See [references/workflows.md](references/workflows.md) for common email, calendar, and task workflows.
>
> **Utility scripts & SKILL_DIR:** See [docs/scripts.md](docs/scripts.md) for script docs and directory variable resolution.
>
> **Command details:** See [references/commands.md](references/commands.md) when you need all flags/options beyond the Quick Reference table.
> **Config reference:** See [references/config.md](references/config.md) for draft instructions, humanizer, and send mode settings.
>
> **JSON schemas:** See [references/json-schemas.md](references/json-schemas.md) when parsing output programmatically and you need the exact schema.
>
> **Troubleshooting:** See [references/troubleshooting.md](references/troubleshooting.md) when commands fail or produce unexpected results.
>
> **Agent ergonomics:** See [references/agent-ergonomics.md](references/agent-ergonomics.md) for known agent-side friction points, workarounds, and GitHub issues for missing features.