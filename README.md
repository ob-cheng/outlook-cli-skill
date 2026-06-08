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

For a detailed walkthrough of search, export, incremental tracking, calendar, tasks, notes, and multi-account, see **[docs/features.md](docs/features.md)**.

### License

MIT

### Contributing

Issues, feature requests, and pull requests are welcome. If something doesn't work or could be better, [open an issue](https://github.com/ob-cheng/outlook-cli-skill/issues) — it helps everyone.

### Related

- [Agent Skills Spec](https://agentskills.io) — The open standard this skill follows

---

## Feature Guide

> **👤 Humans:** [docs/features.md](docs/features.md) — Search, export, calendar, tasks, notes, and multi-account explained for non-technical readers.
>
> **🤖 Agents:** [references/features.md](references/features.md) — Internal behavior, edge cases, undocumented logic, and export lifecycle details.

---

## For Agents

When activated, you load **`SKILL.md`** as your primary instruction set. Pull reference files from `references/` on demand — they cover commands, troubleshooting, workflows, direct sending, and internal behaviors.

`docs/` has additional agent-facing pages for installation, WSL, utility scripts, and directory layout. Load them when you need them.

Quick links:
- **[SKILL.md](SKILL.md)** — Activation instructions, commands, safety rules, workflows (load this first)
- **[references/commands.md](references/commands.md)** — Full flag reference for every command
- **[references/workflows.md](references/workflows.md)** — Common email, calendar, and task workflows
- **[references/features.md](references/features.md)** — Internal behavior, edge cases, and undocumented logic
- **[references/troubleshooting.md](references/troubleshooting.md)** — Debugging when commands fail
- **[docs/install.md](docs/install.md)** — Installation and setup
- **[docs/wsl.md](docs/wsl.md)** — WSL setup
- **[docs/structure.md](docs/structure.md)** — Directory layout and token efficiency

