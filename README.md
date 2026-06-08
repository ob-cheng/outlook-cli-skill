# Outlook CLI Skill

A self-contained agent skill for managing Microsoft Outlook via command line. Compatible with any AI agent that supports the [Agent Skills Spec](https://agentskills.io) (Claude Code, Hermes, OpenClaw, Codex CLI, etc.).

---

## For Humans

### What This Skill Does

Provides AI-friendly access to Outlook for:

- **Email**: Search, read, send, reply, forward, export
- **Calendar**: List events, create meetings, manage schedule
- **Tasks**: Create todos, track progress, mark complete
- **Notes**: Create and manage Outlook notes

Works via COM automation - no Azure setup, OAuth, or API keys required.

### Requirements

- Windows with Outlook desktop app installed (or WSL with Windows Python)
- Python 3.10+ (Windows Python if running from WSL)
- Outlook running when using the skill

### Installation

**Option 1: Ask your agent to install it**

If your agent supports skill installation, just ask:

> "Install the outlook-cli skill from github.com/ob-cheng/outlook-cli-skill"

Your agent will clone the repo, install dependencies, and configure everything.

**Option 2: Manual installation**

```bash
# Clone the repo
git clone https://github.com/ob-cheng/outlook-cli-skill.git

# For Hermes Agent — clone directly into the skills directory:
git clone https://github.com/ob-cheng/outlook-cli-skill.git ~/.hermes/skills/outlook-cli-skill

# For Claude Code — clone into your skills directory:
git clone https://github.com/ob-cheng/outlook-cli-skill.git /path/to/.claude/skills/outlook-cli-skill
```

Then install Python dependencies:

```bash
pip install -r outlook-cli-skill/requirements.txt
```

Verify it works:

```bash
python outlook-cli-skill/outlook.py --version
# Should output: outlook 2.0.0
```

### Safety: Draft-Only Mode

The CLI **enforces draft-only mode at the script level**. No configuration needed — it just works.

- All send/reply/forward commands create drafts by default
- The agent will tell you: "I've saved this as a draft. You can review and send it from Outlook."
- You can review and send manually from Outlook's Drafts folder

### Enable Direct Sending (Optional)

By default, all emails are saved as drafts for safety. To allow your agent to send emails directly:

**Option 1: Ask your agent to enable it**

> "Enable direct sending for the outlook-cli skill"

Your agent will configure the environment variable and confirm before each send.

**Option 2: Manual configuration**

**Hermes Agent** — add to `~/.hermes/config.yaml`:
```yaml
env:
  OUTLOOK_CLI_ALLOW_SEND: "1"
```

**Claude Code** — add to `.claude/settings.json`:
```json
{"env": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}
```

**OpenClaw** — add to `~/.openclaw/openclaw.json`:
```json
{"env": {"vars": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}}
```

**Windows (any agent)** — set system environment variable `OUTLOOK_CLI_ALLOW_SEND=1`

When enabled, the agent will still ask for your confirmation before sending each email.

### Feature Guide

For a detailed walkthrough of how search, export, incremental tracking, calendar, tasks, notes, and multi-account work under the hood, see **[docs/features.md](docs/features.md)**.

### License

MIT

### Related

- [Agent Skills Spec](https://agentskills.io) - The open standard this skill follows

---

## For Agents

Agent-facing documentation is split into separate files under `references/` — load them on demand:

- **[references/commands.md](../references/commands.md)** — Full command reference with all options
- **[references/direct-send.md](../references/direct-send.md)** — Enabling and using `--send`
- **[references/features.md](../references/features.md)** — Internal behavior, export lifecycle, schemas
- **[references/json-schemas.md](../references/json-schemas.md)** — JSON output formats for all commands
- **[references/scripts.md](../references/scripts.md)** — Utility scripts and SKILL_DIR variable resolution
- **[references/structure.md](../references/structure.md)** — Skill directory layout and token efficiency
- **[references/troubleshooting.md](../references/troubleshooting.md)** — Error handling and common issues
- **[references/workflows.md](../references/workflows.md)** — Common workflow patterns and examples
- **[references/wsl.md](../references/wsl.md)** — WSL setup guide

For installation and setup, load **[docs/install.md](docs/install.md)**.

