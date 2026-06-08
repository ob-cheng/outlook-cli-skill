# Direct Sending

By default, all emails are saved as drafts for safety. This doc covers enabling and using direct sending.

## Requirements

Direct sending requires **both**:

1. Config `send_mode` set to `send` — run: `outlook.py config set send_mode send`
2. Use `--send` flag with the command

## Procedure

When the user asks to send an email directly:

1. First attempt with `--send` flag — if config is not set, command fails with a clear error
2. If it fails, inform the user direct sending needs to be enabled
3. Tell them: `python outlook.py config set send_mode send`
4. If `--send` works, show a summary (To, Subject, Body preview)
5. Ask for explicit confirmation: "Ready to send this email?"
6. Only after confirmation, run with `--send` flag:

```bash
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" send --to X --subject Y --body Z --send
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" reply <id> --body "text" --send
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" forward <id> --to X --send
```

## Post-configuration

After configuring, tell the user:

> "Direct sending is now enabled. I'll still ask for your confirmation before sending each email."
