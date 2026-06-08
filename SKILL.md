---
name: outlook-cli
description: Search, send, and manage Outlook emails, calendar, tasks, and notes via CLI. Requires Outlook desktop running on Windows with COM automation.
version: 0.1.0
author: ob-cheng
license: MIT
# Platform restriction (Hermes uses 'platforms', Claude Code uses 'compatibility')
platforms:
  - windows
compatibility:
  - windows
# Categorization (Hermes discovery)
domain: productivity
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
  last-updated: "2026-06-08"
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
# Hermes-specific config (other agents use instructions in Safety Rules section)
metadata.hermes:
  config:
    - key: outlook.draft_only_mode
      description: When true, only create drafts; never send emails directly
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
| Reply | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" reply <id> --body "text"` |
| Forward | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" forward <id> --to X` |
| Calendar | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" cal list/read/create/delete` |
| Tasks | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" tasks list/read/create/complete/delete` |
| Notes | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" notes list/read/create/delete` |
| Export | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" export --output DIR [--stdout]` |
| Folders | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" folders` |
| Config | `${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" config show/set/clear` |

## Draft Workflow

Before every send/reply/forward:

1. Run `outlook.py config show` to read all settings (see [references/config.md](references/config.md))
2. Compose the email body
3. If `draft_instructions` is set → follow them while drafting
4. If `humanizer_enabled` is true → load `humanizer` skill and run the pattern checklist
5. Pass the final body to the CLI

The CLI prints status tags so skipped steps are visible in the output.

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
