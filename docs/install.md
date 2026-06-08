# Installation & Setup

Instructions for AI agents. When the user asks to install this skill or enable direct sending, follow these steps.

## Installing the Skill

### 1. Clone the repo

| Agent | Clone target | Notes |
|---|---|---|
| **Hermes Agent** | `~/.hermes/skills/outlook-cli-skill` | Auto-discovered on next session |
| **Claude Code** | `<skills-dir>/outlook-cli-skill` | Add `"skills": ["outlook-cli-skill"]` to `.claude/settings.json` |
| **OpenClaw** | Anywhere | Point skill search path to parent dir |
| **Other** | Anywhere | Point your agent's skill directory to the parent |

```bash
# Hermes Agent
git clone https://github.com/ob-cheng/outlook-cli-skill.git ~/.hermes/skills/outlook-cli-skill

# Claude Code / OpenClaw / other
git clone https://github.com/ob-cheng/outlook-cli-skill.git
```

### 2. Install Python dependencies

```bash
pip install -r "${SKILL_DIR}/requirements.txt"
```

If `${SKILL_DIR}` is not available, use the absolute path:

```bash
pip install -r outlook-cli-skill/requirements.txt
```

### 3. Verify installation

```bash
python "${SKILL_DIR}/outlook.py" --version
# Expected output: outlook 2.0.0
```

### 4. Verify Outlook connectivity (optional — only if user wants to test now)

```bash
python "${SKILL_DIR}/outlook.py" folders --json
```

This requires Outlook desktop to be running on Windows. If it fails, the skill is still installed correctly — it just can't reach Outlook right now.

---

## Enabling Direct Sending

By default, all emails are saved as drafts for safety. To allow direct sending, set the environment variable `OUTLOOK_CLI_ALLOW_SEND=1`.

### Agent-specific configuration

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

### Post-configuration

After configuring, tell the user:

> "Direct sending is now enabled. I'll still ask for your confirmation before sending each email."

The agent must still confirm with the user before using `--send`. If the env var is not set, the `--send` flag will fail with a clear error message.

---

## Related

- [SKILL.md](../SKILL.md) — skill trigger rules and command reference
- [README.md](../README.md) — human-facing overview
- [references/commands.md](../references/commands.md) — full command reference
- [references/troubleshooting.md](../references/troubleshooting.md) — common issues
