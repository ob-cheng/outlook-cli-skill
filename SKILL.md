---
name: outlook-cli
description: Search, send, and manage Outlook emails, calendar, tasks, and notes via CLI. Requires Outlook desktop running on Windows with COM automation.
version: 2.0.0
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
  - references/wsl.md
  - references/direct-send.md
  - references/scripts.md
  - references/features.md
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

> **WSL / non-Windows setup:** See [references/wsl.md](references/wsl.md) for the `OUTLOOK_CLI_PYTHON` workaround.
>
> **Feature deep-dive:** See [references/features.md](references/features.md) for internal behavior, export lifecycle, content processing, multi-account handling, and undocumented logic.

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

## Safety Rules

**Draft-Only Mode (enforced at script level):**
The CLI defaults to draft mode. All send/reply/forward commands create drafts by default — **no flag needed.**

- Default behavior: emails saved as drafts
- Tell the user: "I've saved this as a draft. You can review and send it from Outlook."
- There is no `--draft` flag. Drafts are the default — always.

**Behavior matrix:**

| Env var set? | `--send` passed? | Result |
|---|---|---|
| — | — | ✅ Saved as draft |
| — | ✅ | ❌ Error: "env var not set" |
| ✅ | — | ✅ **Saved as draft** |
| ✅ | ✅ | ✅ Sent immediately |

Key rule: **`--send` is the only way to send.** Without it, even with the env var set, the email is saved as a draft. There's no way to accidentally send.

> **Direct sending:** See [references/direct-send.md](references/direct-send.md) for enabling and using `--send`.

**Always confirm before:**

- Deleting events/tasks/notes

**Never:**

- Delete without user confirmation
- Forward potentially sensitive emails without verification

> **Workflows & patterns:** See [references/workflows.md](references/workflows.md) for common email, calendar, and task workflows.
>
> **Utility scripts & SKILL_DIR:** See [references/scripts.md](references/scripts.md) for script docs and directory variable resolution.
>
> **Command details:** See [references/commands.md](references/commands.md) when you need all flags/options beyond the Quick Reference table.
>
> **JSON schemas:** See [references/json-schemas.md](references/json-schemas.md) when parsing output programmatically and you need the exact schema.
>
> **Troubleshooting:** See [references/troubleshooting.md](references/troubleshooting.md) when commands fail or produce unexpected results.
