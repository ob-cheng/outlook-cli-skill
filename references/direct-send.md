# Direct Sending

By default, all emails are saved as drafts for safety. This doc covers enabling and using direct sending.

## Requirements

Direct sending requires **both**:

1. Environment variable `OUTLOOK_CLI_ALLOW_SEND=1` must be set
2. Use `--send` flag with the command

## Procedure

When the user asks to send an email directly:

1. First attempt with `--send` flag — if env var is not set, command fails with a clear error
2. If it fails, inform the user direct sending needs to be enabled
3. Provide setup instructions (see [Agent Configuration](#agent-configuration) below)
4. If `--send` works, show a summary (To, Subject, Body preview)
5. Ask for explicit confirmation: "Ready to send this email?"
6. Only after confirmation, run with `--send` flag:

```bash
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z --send
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" reply <id> --body "text" --send
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" forward <id> --to X --send
```

## Agent Configuration

### Hermes Agent

Add to `~/.hermes/config.yaml`:

```yaml
env:
  OUTLOOK_CLI_ALLOW_SEND: "1"
```

### Claude Code

Add to `.claude/settings.json`:

```json
{"env": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}
```

### OpenClaw

Add to `~/.openclaw/openclaw.json`:

```json
{"env": {"vars": {"OUTLOOK_CLI_ALLOW_SEND": "1"}}}
```

### Windows (any agent)

Set system environment variable `OUTLOOK_CLI_ALLOW_SEND=1`.

## Post-configuration

After configuring, tell the user:

> "Direct sending is now enabled. I'll still ask for your confirmation before sending each email."

The agent must still confirm with the user before using `--send`. If the env var is not set, the `--send` flag fails with a clear error message.
