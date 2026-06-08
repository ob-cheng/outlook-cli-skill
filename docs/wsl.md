# WSL Setup

If running this skill from **inside WSL** (Windows Subsystem for Linux), WSL's Python cannot use COM automation (`pywin32`). The CLI must be run via Windows Python instead.

## How it works

The skill uses `${OUTLOOK_CLI_PYTHON:-python}` in all commands. Set the `OUTLOOK_CLI_PYTHON` env var to your Windows Python interpreter path — the fallback `:-python` ensures native Windows users are unaffected.

## Setup

### 1. Find your Windows Python

```bash
# Likely locations
/mnt/c/Users/your-username/AppData/Local/Programs/Python/Python312/python.exe
/mnt/c/Users/your-username/AppData/Local/Programs/Python/Python313/python.exe
/mnt/c/Program Files/Python312/python.exe
```

If unsure, from WSL:

```bash
which python.exe          # if on PATH
cmd.exe /c "where python" # Windows-native lookup
```

### 2. Install dependencies

```bash
/path/to/windows/python.exe -m pip install -r "${SKILL_DIR}/requirements.txt"
```

### 3. Configure the env var

**Hermes Agent** — `~/.hermes/config.yaml`:

```yaml
env:
  OUTLOOK_CLI_PYTHON: "/mnt/c/Users/your-username/AppData/Local/Programs/Python/Python312/python.exe"
```

**Claude Code** — `.claude/settings.json`:

```json
{"env": {"OUTLOOK_CLI_PYTHON": "/mnt/c/Users/your-username/AppData/Local/Programs/Python/Python312/python.exe"}}
```

**Shell (any agent)** — `~/.bashrc` or `~/.zshrc`:

```bash
export OUTLOOK_CLI_PYTHON="/mnt/c/Users/your-username/AppData/Local/Programs/Python/Python312/python.exe"
```

### 4. Verify

```bash
"${OUTLOOK_CLI_PYTHON:-python}" "${SKILL_DIR}/outlook.py" --version
# Expected: outlook 0.1.0
```

## Known-good config (this machine)

- **Windows Python**: `/mnt/c/Users/its_t/AppData/Local/Programs/Python/Python312/python.exe`
- **Deps**: `pywin32`, `beautifulsoup4`, `markdownify`, `rich` — all installed
- **Hermes config**: `~/.hermes/config.yaml` env block
