---
name: outlook-cli
description: Outlook email, calendar, tasks, and notes via CLI. Use when: checking inbox, finding/sending/replying emails, scheduling meetings, checking calendar, managing tasks/notes, exporting emails. Windows-only (COM automation).
version: 0.3.0
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
  last-updated: "2026-06-09"
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
  - references/testing.md
# Hermes-specific config (other agents use instructions in Safety Rules section)
metadata.hermes:
  config:
    - key: outlook.draft_only_mode
      description: "[Config] Draft-only mode — when enabled, emails are saved as drafts instead of being sent"
      default: true
      prompt: Enable draft-only mode for email safety?
---

# Outlook CLI Skill

AI-friendly CLI for Microsoft Outlook. Works via COM automation - no Azure setup, no OAuth, no API keys.

## Strategy

**Search -> Read -> Act**. Find emails first, then read details, then reply/forward/export.

Always add `--json` for structured output when processing programmatically.

> **Installation & setup:** See [docs/install.md](docs/install.md) for agent instructions on installing this skill and enabling direct sending.

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
| Export | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" export --output DIR [--stdout]` |
| Folders | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" folders` |
| People | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" people list/lookup/add` |
| Config | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" config show/set/clear` |

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

### Reply with extra CC
```bash
python outlook.py reply <id> --body "My reply" --cc "newcomer@co.com,support@co.com"
```

### Export thread to markdown
```bash
python outlook.py export --output ./inbox-export --filter-email "client@co.com" --days 7
```

### Calendar today
```bash
python outlook.py cal list --json
```

### Multi-account: always discover folders first
```bash
# See what accounts are connected
python outlook.py folders
# Then target a specific account
python outlook.py search --folder "work@domain.com/Inbox" --filter-email "sender@co.com"
```

### Date filtering options
- `--days N` — last N days (default 7)
- `--from-date YYYY-MM-DD --to-date YYYY-MM-DD` — exact range (overrides --days)
- `--limit N` — stop after N matches (for huge inboxes)

## Known Pitfalls

### Large inbox search can be slow
COM iterates every item in the folder. On 1000+ emails, narrow with `--days 1` or `--limit 20`. The CLI now shows a progress indicator in non-JSON mode.

### Multi-account: search defaults to primary account
Search/send/export target the default account. Run `folders` first to see what's connected, then use `--folder "AccountName/Inbox"` to reach another account.

### EntryIDs can change when emails move
If a stored message ID fails to load, re-search to get the current ID — Outlook regenerates EntryIDs on move/archive operations.

### Draft-only is the default
All compose commands create drafts. Direct sending requires both `send_mode: send` in config AND the `--send` flag.

### Testing from WSL (no Outlook COM)
When making code changes from WSL, COM won't work — but you can still run meaningful tests:

1. **Arg parsing:** `python3 outlook.py <command> --help` — verify new flags appear
2. **Import check:** `python3 -c "from outlook_cli.services.compose import ComposeService"` — catches syntax/import errors
3. **Isolated argparse:** Recreate the parser in a standalone script and test edge cases (empty strings, comma-separated lists, default values)
4. **Simulation:** For stateful logic like progress dots, run the function standalone with sample inputs

**Never** ship after `py_compile` alone. At minimum, exercise argparse and imports.

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
> **Testing (WSL/Linux):** See [references/testing.md](references/testing.md) for mock patterns, pytest setup, and the monkey-patch approach for COM-free unit testing.
