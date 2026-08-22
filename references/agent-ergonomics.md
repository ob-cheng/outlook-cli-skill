# Agent Ergonomics — efficient paths & known gaps

Tips for driving the CLI efficiently as an agent. File issues at
https://github.com/ob-cheng/outlook-cli-skill

## 1. Search by sender display name — use `--filter-name`

When the user names someone but you don't know their SMTP address ("the email
from Babu"), `--filter-email "Babu"` returns nothing — it matches addresses only.
Use `--filter-name` instead (substring match on the sender display name, repeatable):

```bash
python outlook.py search --filter-name "Babu" --days 7 --json
```

Fallbacks if the name is ambiguous: `--filter-domain "alcon.com"` to narrow by
org, or scan a specific account with `--folder "Account/Inbox" --days N` and
inspect `sender_clean` in the results.

## 2. Long email bodies — use `--text-only` and `--max-body-lines`

A single `read` on a threaded email can return 50K+ chars of HTML. Don't pay for
that when you only need the text:

```bash
# Plain text only (omits html_body entirely)
python outlook.py read <id> --text-only --json
# Cap the body length on long threads
python outlook.py read <id> --text-only --max-body-lines 40 --json
```

Parse `text_body` from the JSON. Only reach for the full `html_body` when you
actually need markup.

## 3. Message IDs are long — use `--last` instead of copy-pasting

Message IDs are 100+ hex chars. After a `search`, you don't need to copy them:
`read`, `reply`, and `forward` accept `--last [N]`, which targets the Nth result
from the most recent search (default 1 = most recent), backed by
`~/.outlook-cli/` search caching.

```bash
python outlook.py search --unread --days 1 --json
python outlook.py read --last --json          # most recent hit
python outlook.py reply --last 2 --body "..." # 2nd hit from that search
```

When you do have explicit IDs (e.g. reading several at once), pass them
positionally: `read <id1> <id2> <id3> --json`.

## 4. WSL: `OUTLOOK_CLI_PYTHON` must be set manually

**Issue:** [#7](https://github.com/ob-cheng/outlook-cli-skill/issues/7)

On WSL, every command needs the Windows Python path. Without the env var set,
`python outlook.py` runs WSL's Python which lacks `win32com`.

**Workaround (current):** Override per-command with:
```bash
/path/to/windows/python.exe /path/to/skill/outlook.py <command>
```
Or set env per shell session:
```bash
export OUTLOOK_CLI_PYTHON="/mnt/c/Users/<you>/AppData/Local/Programs/Python/Python312/python.exe"
```

**Proper fix:** Add to `~/.hermes/config.yaml`:
```yaml
env:
  OUTLOOK_CLI_PYTHON: "/mnt/c/Users/<you>/AppData/Local/Programs/Python/Python312/python.exe"
```
