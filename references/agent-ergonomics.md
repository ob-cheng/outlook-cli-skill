# Agent Ergonomics — known gaps & workarounds

File issues for these at https://github.com/ob-cheng/outlook-cli-skill

## 1. Search by sender display name (`--filter-name` missing)

**Issue:** [#4](https://github.com/ob-cheng/outlook-cli-skill/issues/4)

`--filter-email` matches SMTP addresses only. If the user says "the email from Babu" and you don't know `babu.ningappa@alcon.com`, `--filter-email "Babu"` returns nothing.

**Workaround:** Use `--folder "AccountName/Inbox" --days N` to scan the right account broadly, then inspect sender_clean in results. Or use `--filter-domain "alcon.com"` to narrow by domain when you know the org.

## 2. Read returns full HTML body (token-heavy)

**Issue:** [#5](https://github.com/ob-cheng/outlook-cli-skill/issues/5)

Single `read` on a threaded email can return 50K+ chars of HTML. The `text_body` field is included but mixed with the full HTML version.

**Workaround:** Parse `text_body` from the JSON output — it's the plain-text version. For long threads, the token cost is unavoidable since the CLI returns the full email object. Consider piping through `jq .text_body` if you only need the text.

## 3. Message IDs are 100+ hex chars, no shorthand

**Issue:** [#6](https://github.com/ob-cheng/outlook-cli-skill/issues/6)

Every `reply` / `read` / `forward` requires copy-pasting a long message_id from search output.

**Workaround:** Keep the message_id string in your active context. Search results return IDs in the `message_id` field of each email object. When a row was just found, the ID is in the prior terminal output — pipe search into `--json` and reference the field.

## 4. WSL: `OUTLOOK_CLI_PYTHON` must be set manually

**Issue:** [#7](https://github.com/ob-cheng/outlook-cli-skill/issues/7)

On WSL, every command needs the Windows Python path. Without the env var set, `python outlook.py` runs WSL's Python which lacks `win32com`.

**Workaround (current):** Override per-command with:
```bash
/path/to/windows/python.exe /path/to/skill/outlook.py <command>
```
Or set env per shell session:
```bash
export OUTLOOK_CLI_PYTHON="/mnt/c/Users/its_t/AppData/Local/Programs/Python/Python312/python.exe"
```

**Proper fix:** Add to `~/.hermes/config.yaml`:
```yaml
env:
  OUTLOOK_CLI_PYTHON: "/mnt/c/Users/its_t/AppData/Local/Programs/Python/Python312/python.exe"
```
