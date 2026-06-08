<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://img.shields.io/badge/outlook--cli-0.1.0-blue?style=flat-square&logo=microsoftoutlook&logoColor=white&labelColor=1a1a2e">
    <img alt="Outlook CLI" src="https://img.shields.io/badge/outlook--cli-0.1.0-blue?style=flat-square&logo=microsoftoutlook&logoColor=white&labelColor=f0f0f0">
  </picture>
  <img alt="License" src="https://img.shields.io/github/license/ob-cheng/outlook-cli-skill?style=flat-square&color=success">
  <img alt="Python" src="https://img.shields.io/badge/python-3.10+-informational?style=flat-square&logo=python&logoColor=white">
  <img alt="Platform" src="https://img.shields.io/badge/platform-windows%20%7C%20wsl-lightgrey?style=flat-square">
  <img alt="Agent Skills" src="https://img.shields.io/badge/Agent%20Skills-compatible-8A2BE2?style=flat-square">
</p>

---

# Outlook CLI Skill

**A proper CLI for Outlook.** Search, read, send, and export emails. Manage your calendar, tasks, and notes — all from the command line. No Azure app registration. No OAuth. No API keys.

Works with any AI agent that supports the [Agent Skills Spec](https://agentskills.io) — Claude Code, Hermes, OpenClaw, Codex CLI, and everything else with a `SKILL.md` loader.

---

## What This Does

| Feature | Description |
|---|---|
| **Email** | Search by folder, date, sender, domain, keywords. Read, send, reply, forward. Export to markdown or JSON. |
| **Calendar** | List events, check schedules, create meetings. |
| **Tasks** | Create, list, update, and complete Outlook tasks. |
| **Notes** | Create and manage Outlook notes from the terminal. |

Under the hood it uses **COM automation** against the running Outlook desktop client — the same API Outlook itself uses. That means zero cloud dependencies, zero setup ceremonies, zero rate limits.

## Requirements

- **Windows** with Outlook desktop app installed (any edition from 2016+)
- **Python 3.10+** (use the Windows Python interpreter if running from WSL)
- **Outlook must be running** when the CLI executes

## Quick Start

```bash
# Clone the repo
git clone https://github.com/ob-cheng/outlook-cli-skill.git

# Install dependencies
pip install -r outlook-cli-skill/requirements.txt

# Verify
python outlook-cli-skill/outlook.py --version
# → outlook 0.1.0
```

Your AI agent can do all of this for you with a single command — just ask it to install the skill.

## Draft-Only Mode (Default)

**Every send, reply, and forward command creates a draft by default.** Nothing goes out without a human looking at it first.

The agent will produce output like: *"I've saved this as a draft. Review and send it from Outlook's Drafts folder."*

This is enforced at the script level. No configuration needed. No footguns.

### Enable Direct Sending (Optional)

If you want your agent to send emails without the draft step, set `OUTLOOK_CLI_ALLOW_SEND=1`. Even with this enabled, the agent will still ask for confirmation before each send — you always get a chance to say no.

| Platform | How |
|---|---|
| **Hermes** | `env.OUTLOOK_CLI_ALLOW_SEND: "1"` in `~/.hermes/config.yaml` |
| **Claude Code** | `{"env": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}` in `.claude/settings.json` |
| **OpenClaw** | `{"env": {"vars": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}}` in `~/.openclaw/openclaw.json` |
| **Any agent (Windows)** | System environment variable `OUTLOOK_CLI_ALLOW_SEND=1` |

## Feature Overview

- **Search** — folder, date range, sender, domain, keywords, unread filter
- **Export** — emails to markdown or JSON, with optional body truncation
- **Incremental tracking** — only export what's new since last run
- **Calendar** — list events, create meetings, manage your schedule
- **Tasks** — full CRUD for Outlook tasks
- **Notes** — create and manage Outlook notes
- **Multi-account** — works across multiple Outlook profiles

Detailed walkthroughs for each feature: **[docs/features.md](docs/features.md)**

## Agent Integration

This repo follows the [Agent Skills Spec](https://agentskills.io). When loaded by an AI agent:

| Resource | What it's for |
|---|---|
| **[SKILL.md](SKILL.md)** | Primary instruction set — activation rules, commands, safety, workflows |
| **[references/commands.md](references/commands.md)** | Full flag reference for every command |
| **[references/workflows.md](references/workflows.md)** | Common email, calendar, and task workflows |
| **[references/features.md](references/features.md)** | Internal behavior, edge cases, export lifecycle |
| **[references/troubleshooting.md](references/troubleshooting.md)** | Debugging when commands fail |
| **[reference/direct-send.md](reference/direct-send.md)** | Direct send mode setup and behavior |
| **[docs/install.md](docs/install.md)** | Installation across all platforms |
| **[docs/wsl.md](docs/wsl.md)** | WSL-specific setup |
| **[docs/structure.md](docs/structure.md)** | Directory layout and token efficiency notes |

## Command Quick Reference

```bash
# Email
python outlook.py search --folder Inbox --days 7
python outlook.py read <message-id>
python outlook.py send --to addr@example.com --subject "Hello" --body "..."

# Calendar
python outlook.py cal list --days 7
python outlook.py cal create --subject "Standup" --start "2025-01-15 09:00"

# Tasks
python outlook.py task list
python outlook.py task create --subject "Review PR"

# Notes
python outlook.py note list
python outlook.py note create --subject "Ideas"
```

## Project Structure

```
outlook-cli-skill/
├── SKILL.md              # Agent activation instructions
├── outlook.py            # Entry point
├── requirements.txt      # Python dependencies
├── README.md             # You are here
├── docs/                 # Human & agent-facing documentation
│   ├── features.md       # Feature walkthroughs
│   ├── install.md        # Installation guide
│   ├── wsl.md            # WSL setup
│   ├── structure.md      # Directory layout
│   └── scripts.md        # Utility scripts
└── references/           # Agent-facing reference (loaded on demand)
    ├── commands.md       # Full command reference
    ├── workflows.md      # Common workflows
    ├── features.md       # Internal edge cases & undocumented logic
    ├── troubleshooting.md
    └── direct-send.md    # Direct send mode
```

## License

MIT — do what you want with it.

## Contributing

Issues, feature requests, and pull requests are welcome. If something doesn't work or could be better, [open an issue](https://github.com/ob-cheng/outlook-cli-skill/issues).

## Related

- [Agent Skills Spec](https://agentskills.io) — the open standard this skill follows
- [pywin32](https://github.com/mhammond/pywin32) — Python for Windows extensions (what makes COM work)
