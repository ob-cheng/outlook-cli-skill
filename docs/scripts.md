# Utility Scripts

The skill ships utility scripts under `${SKILL_DIR}/scripts/`.

## SKILL_DIR resolution

| Agent | Variable |
|-------|----------|
| Hermes | `${SKILL_DIR}` |
| Claude Code | `${CLAUDE_SKILL_DIR}` (aliased to `${SKILL_DIR}`) |
| OpenClaw | `${SKILL_DIR}` |
| Other | `${SKILL_DIR}` per Agent Skills Spec |

## Available scripts

### validate-export.py

Validates exported email files:

```bash
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/scripts/validate-export.py" <export-dir> --format json
```

### format-email.py

Formats email JSON for human display:

```bash
${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/outlook.py" read <id> --json | ${OUTLOOK_CLI_PYTHON:-python} "${SKILL_DIR}/scripts/format-email.py" --format summary
```
