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

- Windows with Outlook desktop app installed
- Python 3.10+
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

### License

MIT

### Related

- [Agent Skills Spec](https://agentskills.io) - The open standard this skill follows

---

## For Agents

This section contains usage instructions for AI agents. For installation and setup, load **[docs/install.md](docs/install.md)** — it covers cloning the repo, installing dependencies, verifying installation, and enabling direct sending for Hermes, Claude Code, OpenClaw, and other agents.

### Quick Command Reference

| Intent | Command |
| ------ | ------- |
| Find emails | `python "${SKILL_DIR}/outlook.py" search [options] --json` |
| Read email | `python "${SKILL_DIR}/outlook.py" read <id> --json` |
| Send email | `python "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z` |
| Reply | `python "${SKILL_DIR}/outlook.py" reply <id> --body "text"` |
| Forward | `python "${SKILL_DIR}/outlook.py" forward <id> --to X` |
| Calendar | `python "${SKILL_DIR}/outlook.py" cal list/read/create/delete --json` |
| Tasks | `python "${SKILL_DIR}/outlook.py" tasks list/read/create/complete/delete --json` |
| Notes | `python "${SKILL_DIR}/outlook.py" notes list/read/create/delete --json` |
| Export | `python "${SKILL_DIR}/outlook.py" export --output DIR [--stdout]` |
| Folders | `python "${SKILL_DIR}/outlook.py" folders --json` |

**Note:** All send/reply/forward commands create drafts by default.

### Direct Sending (When Enabled by User)

By default, all email commands create drafts. If the user has enabled direct sending (see [docs/install.md](docs/install.md)), you can send immediately:

```bash
python "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z --send
python "${SKILL_DIR}/outlook.py" reply <id> --body "text" --send
python "${SKILL_DIR}/outlook.py" forward <id> --to X --send
```

**Before using `--send`, you MUST:**
1. Check if the env var is set (attempt will fail with clear error if not)
2. Show the user a summary of what will be sent
3. Ask for explicit confirmation before sending

### Directory Variable

- **Claude Code**: `${CLAUDE_SKILL_DIR}` (aliased to `${SKILL_DIR}`)
- **Hermes/OpenClaw/Others**: `${SKILL_DIR}`

### Reference Files

For detailed documentation, load these on demand:

- `${SKILL_DIR}/docs/install.md` — Installation, setup, enabling direct sending
- `${SKILL_DIR}/references/commands.md` — Full command reference with all options
- `${SKILL_DIR}/references/json-schemas.md` — JSON output formats for all commands
- `${SKILL_DIR}/references/workflows.md` — Common workflow patterns and examples
- `${SKILL_DIR}/references/troubleshooting.md` — Error handling and common issues

### Skill Structure

```text
outlook-cli-skill/
├── SKILL.md              # Entry point - triggers & core instructions
├── outlook.py            # CLI entry point
├── outlook_cli/          # CLI implementation
│   ├── cli.py
│   ├── core/
│   ├── services/
│   └── utils/
├── docs/
│   └── install.md        # Agent install & setup instructions
├── references/
│   ├── commands.md
│   ├── json-schemas.md
│   ├── workflows.md
│   └── troubleshooting.md
├── scripts/
│   ├── validate-export.py
│   └── format-email.py
├── assets/
│   └── email-template.md
├── requirements.txt
└── README.md
```

### Token Efficiency

- SKILL.md: ~120 lines (loaded when triggered)
- References: Loaded on-demand only when needed
- Scripts: Code never enters context — only output
- CLI code: Never enters context — runs via shell
