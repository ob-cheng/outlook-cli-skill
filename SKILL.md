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

## Quick Reference

Run all commands using the skill directory path:

| Intent | Command |
|--------|---------|
| Find emails | `python "${SKILL_DIR}/outlook.py" search [options]` |
| Read email | `python "${SKILL_DIR}/outlook.py" read <id> --json` |
| Send email | `python "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z` |
| Reply | `python "${SKILL_DIR}/outlook.py" reply <id> --body "text"` |
| Forward | `python "${SKILL_DIR}/outlook.py" forward <id> --to X` |
| Calendar | `python "${SKILL_DIR}/outlook.py" cal list/read/create/delete` |
| Tasks | `python "${SKILL_DIR}/outlook.py" tasks list/read/create/complete/delete` |
| Notes | `python "${SKILL_DIR}/outlook.py" notes list/read/create/delete` |
| Export | `python "${SKILL_DIR}/outlook.py" export --output DIR [--stdout]` |
| Folders | `python "${SKILL_DIR}/outlook.py" folders` |

> **Note on multi-account setups:** Calendar (`cal`), Tasks (`tasks`), and Notes (`notes`) commands use Outlook's default account (the one marked as default under File → Account Settings → Data Files). If you have multiple accounts (e.g., iCloud + Exchange), the default may not be the one with your data. To fix this, set the desired account as default in Outlook (File → Account Settings → Data Files → Set as Default). Email `search`/`read`/`export` search the **default store** (your primary mailbox). Use `--folder` to target a specific account's folder.

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

**Always confirm before:**

- Deleting events/tasks/notes

**Never:**

- Delete without user confirmation
- Forward potentially sensitive emails without verification

## Direct Sending (Agent Instructions)

Direct sending requires BOTH:

1. Environment variable `OUTLOOK_CLI_ALLOW_SEND=1` must be set
2. Use `--send` flag with the command

**When user asks to send an email directly:**

1. First attempt with `--send` flag — if env var is not set, command will fail with clear error
2. If it fails, inform user they need to enable direct sending
3. Provide setup instructions based on their agent:

```
To enable direct sending, configure the environment variable:

- Claude Code: Add to .claude/settings.json:
  {"env": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}

- Hermes: Add to ~/.hermes/config.yaml:
  env:
    OUTLOOK_CLI_ALLOW_SEND: "1"

- OpenClaw: Add to ~/.openclaw/openclaw.json:
  {env: {vars: {OUTLOOK_CLI_ALLOW_SEND: "1"}}}

- Windows: Set system environment variable OUTLOOK_CLI_ALLOW_SEND=1
```

**When direct sending IS enabled:**

1. Show the user a summary of the email (To, Subject, Body preview)
2. Ask for explicit confirmation: "Ready to send this email?"
3. Only after confirmation, run with `--send` flag:

```bash
python "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z --send
python "${SKILL_DIR}/outlook.py" reply <id> --body "text" --send
python "${SKILL_DIR}/outlook.py" forward <id> --to X --send
```

## Decision Tree

```
User wants to...
├── Find emails ──────────> search --days N --keyword X --filter-email Y
├── Read specific email ──> read <message-id> --json
├── Compose new email ────> send --to --subject --body  (drafts by default)
├── Reply to email ───────> reply <id> --body [--all for reply-all]
├── Forward email ────────> forward <id> --to
├── Check calendar ───────> cal list [--start --end]
├── Schedule meeting ─────> cal create --subject --start --end [--required]
├── Manage tasks ─────────> tasks list/create/complete/delete
├── Take notes ───────────> notes list/create/read/delete
├── Export for AI ────────> export --output . --stdout
└── Export to files ──────> export --output DIR [--format json/markdown]
```

## Common Patterns

**Triage inbox:**
```bash
python "${SKILL_DIR}/outlook.py" search --unread --json
```

**Find emails from someone:**
```bash
python "${SKILL_DIR}/outlook.py" search --filter-email boss@company.com --days 14 --json
```

**Export for AI processing:**
```bash
python "${SKILL_DIR}/outlook.py" export --output . --stdout --days 7
```

**Schedule meeting with attendees:**
```bash
python "${SKILL_DIR}/outlook.py" cal create --subject "Team Sync" --start "2026-06-10 14:00" --end "2026-06-10 15:00" --required "alice@co.com,bob@co.com"
```

**Task workflow:**
```bash
python "${SKILL_DIR}/outlook.py" tasks create --subject "Review PR" --due 2026-06-15 --priority high
python "${SKILL_DIR}/outlook.py" tasks complete <task-id>
```

## Reference Files

For detailed documentation, load these on demand:

- `docs/install.md` - Installation, setup, and enabling direct sending
- `references/commands.md` - Full command reference with all options
- `references/json-schemas.md` - JSON output formats for all commands
- `references/workflows.md` - Common workflow patterns and examples
- `references/troubleshooting.md` - Error handling and common issues

## Skill Directory

This skill is self-contained. Commands use `${SKILL_DIR}` which resolves to the skill's installation directory.

- **Claude Code**: `${CLAUDE_SKILL_DIR}` (aliased to `${SKILL_DIR}`)
- **Hermes**: `${SKILL_DIR}`
- **OpenClaw**: `${SKILL_DIR}`
- **Other agents**: Most support `${SKILL_DIR}` per the Agent Skills Spec

**Utility scripts:**
```bash
# Validate export output
python "${SKILL_DIR}/scripts/validate-export.py" <export-dir> --format json

# Format email JSON for display
python "${SKILL_DIR}/outlook.py" read <id> --json | python "${SKILL_DIR}/scripts/format-email.py" --format summary
```
