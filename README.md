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

Clone or download from GitHub:

```bash
git clone https://github.com/ob-cheng/outlook-cli-skill.git
```

Then point your agent to the skill directory (see "For Agents" section for agent-specific instructions).

Or, if your agent supports it, just ask:

> "Install the outlook-cli skill from github.com/ob-cheng/outlook-cli-skill"

### Safety: Draft-Only Mode

The CLI **enforces draft-only mode at the script level**. No configuration needed — it just works.

- All send/reply/forward commands create drafts by default
- The agent will tell you: "I've saved this as a draft. You can review and send it from Outlook."
- You can review and send manually from Outlook's Drafts folder

### Enable Direct Sending (Optional)

If you want the agent to send emails directly (not just drafts), ask your agent:

> "Enable direct sending for the outlook-cli skill"

The agent will configure the required environment variable for you. When enabled, the agent will still ask for your confirmation before sending each email.

### License

MIT

### Related

- [Agent Skills Spec](https://agentskills.io) - The open standard this skill follows

---

## For Agents

This section contains installation, verification, and usage instructions for AI agents.

### Install the Skill

#### Claude Code

```bash
# Clone to your skills directory
git clone https://github.com/ob-cheng/outlook-cli-skill.git /path/to/.claude/skills/outlook-cli-skill
```

Add to `.claude/settings.json`:
```json
{"skills": ["outlook-cli-skill"]}
```

#### Hermes Agent

```bash
# Clone or symlink into your skills directory
git clone https://github.com/ob-cheng/outlook-cli-skill.git ~/.hermes/skills/outlook-cli-skill
```

The skill will be auto-discovered on next run. Or install via CLI:
```
hermes skills install outlook-cli-skill
```

#### OpenClaw / Others

Clone the repo and point your agent's skill directory to the cloned path:

```bash
git clone https://github.com/ob-cheng/outlook-cli-skill.git
```

Then configure your agent's skill search path to include the parent directory (see your agent's documentation for details).

### Verify Installation

Run this command to verify the skill is installed correctly:

```bash
python "${SKILL_DIR}/outlook.py" --version
```

If this fails, check:

1. Python 3.10+ is installed
2. The skill directory path is correct
3. Dependencies are installed: `pip install -r "${SKILL_DIR}/requirements.txt"`

### Install Dependencies

```bash
pip install -r "${SKILL_DIR}/requirements.txt"
```

### Verify Outlook Connection

```bash
python "${SKILL_DIR}/outlook.py" folders --json
```

This should return a list of Outlook folders. If it fails:

- Ensure Outlook desktop app is running
- This skill requires Windows with Outlook installed (COM automation)

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

By default, all email commands create drafts. If the user has enabled direct sending (by setting `OUTLOOK_CLI_ALLOW_SEND=1` env var), you can send immediately:

```bash
# Send immediately (only works if env var is set)
python "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z --send
python "${SKILL_DIR}/outlook.py" reply <id> --body "text" --send
python "${SKILL_DIR}/outlook.py" forward <id> --to X --send
```

**Before using `--send`, you MUST:**

1. Check if the env var is set (attempt will fail with clear error if not)
2. Show the user a summary of what will be sent
3. Ask for explicit confirmation before sending

If user wants to enable direct sending, instruct them to configure the environment variable:

- **Claude Code**: Add `{"env": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}` to `.claude/settings.json`
- **Hermes**: Add `env:\n  OUTLOOK_CLI_ALLOW_SEND: "1"` to `~/.hermes/config.yaml`
- **OpenClaw**: Add `{env: {vars: {OUTLOOK_CLI_ALLOW_SEND: "1"}}}` to `~/.openclaw/openclaw.json`
- **Windows**: Set system environment variable `OUTLOOK_CLI_ALLOW_SEND=1`

### Directory Variable

- **Claude Code**: `${CLAUDE_SKILL_DIR}` (aliased to `${SKILL_DIR}`)
- **Hermes/OpenClaw/Others**: `${SKILL_DIR}`

### Reference Files

For detailed documentation, load these on demand:

- `${SKILL_DIR}/references/commands.md` - Full command reference with all options
- `${SKILL_DIR}/references/json-schemas.md` - JSON output formats for all commands
- `${SKILL_DIR}/references/workflows.md` - Common workflow patterns and examples
- `${SKILL_DIR}/references/troubleshooting.md` - Error handling and common issues

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
- Scripts: Code never enters context - only output
- CLI code: Never enters context - runs via shell
