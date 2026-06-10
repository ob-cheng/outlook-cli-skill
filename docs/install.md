# Installation & Setup

Instructions for AI agents. When the user asks to install this skill, follow these steps.

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
# Expected: outlook 0.2.0
```

### 4. Verify Outlook connectivity (optional — only if user wants to test now)

```bash
python "${SKILL_DIR}/outlook.py" folders --json
```

This requires Outlook desktop to be running on Windows. If it fails, the skill is still installed correctly — it just can't reach Outlook right now.

---

## WSL Setup (One-liner)

If running from WSL, set `OUTLOOK_CLI_PYTHON` to your Windows Python path. The skill auto-detects common paths and suggests the correct one if missing.

**Quick setup** — add to your shell profile (`~/.bashrc` or `~/.zshrc`):

```bash
export OUTLOOK_CLI_PYTHON="/mnt/c/Users/<your-username>/AppData/Local/Programs/Python/Python312/python.exe"
```

**For Hermes Agent** — add to `~/.hermes/config.yaml`:

```yaml
env:
  OUTLOOK_CLI_PYTHON: "/mnt/c/Users/<your-username>/AppData/Local/Programs/Python/Python312/python.exe"
```

**For Claude Code** — add to `.claude/settings.json`:

```json
{"env": {"OUTLOOK_CLI_PYTHON": "/mnt/c/Users/<your-username>/AppData/Local/Programs/Python/Python312/python.exe"}}
```

The skill will auto-detect your Windows Python path and suggest the correct value if you run it without the env var set.

> **Detailed WSL guide:** See [docs/wsl.md](wsl.md) for more options.

### WSL: Windows Python missing dependencies

First run from WSL may fail with `ModuleNotFoundError` if the Windows Python doesn't have the tool's dependencies. Fix:

```bash
"${OUTLOOK_CLI_PYTHON}" -m pip install rich markdownify beautifulsoup4 pywin32
```

Or with the explicit path:
```bash
"/mnt/c/Users/<username>/AppData/Local/Programs/Python/Python313/python.exe" -m pip install rich markdownify beautifulsoup4 pywin32
```

This is a one-time setup step. After install, all commands work.

---

## First-Time Setup (Ask the User)

After installation, before drafting the first email, ask the user:

> **"Do you want to set custom drafting instructions? (e.g. 'Keep it brief', 'Always mention next steps')"**

- If no → skip (instructions are empty by default)
- If yes → ask the user to provide their instructions, then save them **exactly as given — verbatim, no rewording, no interpretation**:

  ```
  outlook.py config set draft_instructions "<exact text the user provided>"
  ```

> **"Do you want to enable humanizer processing for email drafts?"**

- If yes → run: `outlook.py config set humanizer_enabled true`
- If no → skip (disabled by default)

> **"Do you want to allow direct sending instead of draft-only mode?"**

- If yes → run: `outlook.py config set send_mode send`
  Then explain: *"I'll still ask for confirmation before sending each email."*
- If no → skip (draft mode is default)

These settings are stored in `~/.outlook-cli/config.json` and printed as status tags on every send/reply/forward so the workflow is visible.

---

## Configuration Reference

All settings managed via the CLI:

```bash
# View current settings
python outlook.py config show

# Enable direct sending
python outlook.py config set send_mode send

# Set custom drafting instructions
python outlook.py config set draft_instructions "Keep it under 3 sentences"

# Enable humanizer processing
python outlook.py config set humanizer_enabled true

# Reset to default
python outlook.py config clear send_mode
```

| Key | Default | Description |
|---|---|---|
| `send_mode` | `draft` | `draft` or `send` — allows `--send` flag |
| `draft_instructions` | `""` | Free-text instructions for the agent when drafting |
| `humanizer_enabled` | `false` | Run humanizer skill on email bodies |

When `send_mode: send` is set and `--send` is used, the user must still confirm first.

---

## Updating the Skill

To pull the latest version:

```bash
cd "${SKILL_DIR}"
git pull
pip install -r "${SKILL_DIR}/requirements.txt"
```

Configuration and user data live in `~/.outlook-cli/` — outside the repo — so `git pull` won't touch settings, people data, or any personal state.

If `git pull` fails due to local modifications:

```bash
git stash && git pull && git stash pop
```

---

## Related

- [SKILL.md](../SKILL.md) — skill trigger rules and command reference
- [README.md](../README.md) — human-facing overview
- [references/commands.md](../references/commands.md) — full command reference
- [references/troubleshooting.md](../references/troubleshooting.md) — common issues
- [references/wsl.md](../references/wsl.md) — WSL setup guide
