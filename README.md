[![Version](https://img.shields.io/badge/version-0.3.0-blue?style=flat-square)](https://github.com/ob-cheng/outlook-cli-skill)
[![License](https://img.shields.io/badge/license-MIT-success?style=flat-square)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.10+-informational?style=flat-square&logo=python&logoColor=white)](https://www.python.org)
[![Platform](https://img.shields.io/badge/platform-windows%20%7C%20wsl-lightgrey?style=flat-square)]()
[![Agent Skills](https://img.shields.io/badge/Agent%20Skills-compatible-8A2BE2?style=flat-square)](https://agentskills.io)

---

# Outlook CLI Skill

**A proper CLI for Outlook.** Search, read, send, and export emails. Manage your calendar, tasks, and notes — all from the command line. No Azure app registration. No OAuth. No API keys.

Works with any AI agent that supports the [Agent Skills Spec](https://agentskills.io) — Claude Code, Hermes, OpenClaw, Codex CLI, and everything else with a `SKILL.md` loader.

---

## What This Does

| Feature | Description |
|---|---|
| **Email** | Search by folder, date, sender, domain, keywords. Read, send, reply, forward. |
| **Calendar** | List events, check schedules, create meetings with attendees. |
| **Tasks** | Create, list, update, and complete Outlook tasks with priority, due dates, categories. |
| **Notes** | Create and manage Outlook notes from the terminal with color-coding. |
| **Export** | Export emails to markdown (Obsidian-ready) or JSON. Incremental tracking, thread grouping, batch mode, stdout. |
| **Folders** | List all Outlook folders by name and path. |
| **People Directory** | Auto-tracked contact list — every email read adds unknown senders/recipients. Lookup, add, list. |
| **Config** | Draft-only vs direct send, custom drafting instructions, humanizer integration. |

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
# → outlook 0.3.0
```
<br>
<blockquote><b>Copy-paste this to your AI agent:</b><br>
<code>install the outlook-cli skill from https://github.com/ob-cheng/outlook-cli-skill</code>
</blockquote>

## Draft-Only Mode (Default)

**Every send, reply, and forward command creates a draft by default.** Nothing goes out without a human looking at it first.

The agent will produce output like: *"I've saved this as a draft. Review and send it from Outlook's Drafts folder."*

This is enforced at the script level. No configuration needed. No footguns.

### Configuration

All settings are stored in `~/.outlook-cli/config.json`. Manage them via the CLI:

```bash
# View current settings
python outlook.py config show

# Enable direct sending
python outlook.py config set send_mode send

# Set custom drafting instructions (agent follows these)
python outlook.py config set draft_instructions "Keep it under 3 sentences, mention next steps"

# Enable humanizer processing
python outlook.py config set humanizer_enabled true

# Reset to default
python outlook.py config clear send_mode
```

**Settings reference:**

| Key | Default | Description |
|---|---|---|
| `send_mode` | `draft` | `draft` or `send` — controls whether `--send` flag is allowed |
| `draft_instructions` | `""` | Free-text instructions for the agent when drafting emails |
| `humanizer_enabled` | `false` | Whether to run the humanizer skill on email bodies |

### Behavior Matrix

| send_mode | `--send` flag? | Result |
|---|---|---|
| `draft` | — | ✅ Saved as draft |
| `draft` | ✅ | ❌ Rejected |
| `send` | — | ✅ Saved as draft |
| `send` | ✅ | ✅ Sent immediately |

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
| **[references/config.md](references/config.md)** | Draft instructions, humanizer, send mode settings |
| **[references/workflows.md](references/workflows.md)** | Common email, calendar, and task workflows |
| **[references/features.md](references/features.md)** | Internal behavior, edge cases, export lifecycle |
| _(developer docs: see `docs/developers.md` in the repo)_ | Testing patterns, mock setup, pytest |
| **[references/troubleshooting.md](references/troubleshooting.md)** | Debugging when commands fail |
| **[references/direct-send.md](references/direct-send.md)** | Direct send mode setup and behavior |
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
python outlook.py tasks list
python outlook.py tasks create --subject "Review PR"

# Notes
python outlook.py notes list
python outlook.py notes create --subject "Ideas"

# Folders
python outlook.py folders

# Config
python outlook.py config show
python outlook.py config set send_mode send
python outlook.py config set draft_instructions "Keep it brief"
python outlook.py config set humanizer_enabled true
python outlook.py config clear humanizer_enabled

# People directory
python outlook.py people list [--json]
python outlook.py people lookup "Alice"
python outlook.py people add "Alice Smith" alice@co.com
```

## Project Structure

```
outlook-cli-skill/
├── LICENSE               # MIT
├── SKILL.md              # Agent activation instructions
├── outlook.py            # Entry point
├── requirements.txt      # Python dependencies
├── README.md             # You are here
├── scripts/              # Utility scripts
│   ├── validate-export.py
│   └── format-email.py
├── docs/                 # Human & agent-facing documentation
│   ├── install.md        # Installation guide
│   ├── wsl.md            # WSL setup
│   ├── structure.md      # Directory layout
│   └── scripts.md        # Utility scripts
└── references/           # Agent-facing reference (loaded on demand)
    ├── commands.md       # Full command reference
    ├── config.md         # Draft instructions, humanizer, send mode
    ├── workflows.md      # Common workflows
    ├── features.md       # Internal edge cases & undocumented logic
    ├── troubleshooting.md
    └── direct-send.md    # Direct send mode

~/.outlook-cli/           # User data (auto-created, not in repo)
├── config.json
└── people.json
```

## Updating

```bash
cd path/to/outlook-cli-skill
git pull
pip install -r requirements.txt   # only needed if dependencies changed

# Quick smoke check (no test scripts needed)
python outlook.py --help > /dev/null && echo "CLI OK"
```

Your configuration (`~/.outlook-cli/config.json`) and people directory (`~/.outlook-cli/people.json`) live outside the repo — `git pull` will never touch them.

If a `git pull` fails due to local changes (e.g. you modified a reference doc), stash first:

```bash
git stash && git pull && git stash pop
```

## For Developers

This is a user-facing README. Developer resources live in the repo but aren't
advertised here so users don't trip over them:

- **`tests/`** — pytest suite (no COM needed, runs on Linux/WSL)
- **`scripts/smoke-test.py`** — Windows COM smoke test (needs Outlook)
- **`references/testing.md`** — mock patterns, monkey-patching, CI setup

## License

MIT — do what you want with it. See [LICENSE](LICENSE) for the legalese.

## Contributing

Issues, feature requests, and pull requests are welcome. If something doesn't work or could be better, [open an issue](https://github.com/ob-cheng/outlook-cli-skill/issues).

## Related

- [Agent Skills Spec](https://agentskills.io) — the open standard this skill follows
- [pywin32](https://github.com/mhammond/pywin32) — Python for Windows extensions (what makes COM work)
