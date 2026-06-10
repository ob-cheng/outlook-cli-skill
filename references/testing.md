# Testing outlook-cli from WSL/Linux

The CLI requires Windows COM (pywin32). When making changes from WSL
where Outlook isn't reachable, don't stop at `py_compile` — that only
catches syntax errors, not logic bugs.

## What you CAN test (no COM needed)

### 1. Argparse validation
Rebuild the parser in isolation (don't import the full CLI module —
it pulls win32com at import time):

```python
import argparse
parser = build_parser()  # copy from tests/test_argparse.py
args = parser.parse_args(["reply", "abc123", "--body", "test", "--cc", "a@x.com,b@x.com"])
# Assert cc == ["a@x.com", "b@x.com"]
```

### 2. Import integrity
Lazy-import the service modules directly (bypass `__init__.py`):

```python
from outlook_cli.services.search import SearchService  # no win32com needed
from outlook_cli.core.folders import list_all_folders  # no win32com needed
```

`ComposeService` is lazy via `get_compose_service(namespace)` — don't import it
at module level.

### 3. Monkey-patch COM conversion
`_extract_from_folder` iterates COM items and calls
`Email.from_mail_item(mail, is_sent)`. To test filtering logic
without COM, monkey-patch `from_mail_item` to return pre-built emails:

```python
from unittest.mock import patch
from outlook_cli.core.models import Email

with patch.object(Email, "from_mail_item") as mock_from:
    mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
    result = svc.search(limit=3)
```

The mock COM item (built via `build_mock_items()` in conftest.py) carries
a `_test_email` attribute that the side_effect returns.

### 4. Folder traversal
`list_all_folders()` walks `MagicMock` folder trees fine — no COM needed.
Use `PropertyMock(side_effect=Exception(...))` to test exception safety.

### 5. Progress dots
Copy `_print_progress_dot` from cli.py and test with `io.StringIO`:

```python
def _print_progress_dot(count, out=None):
    if count % 20 == 0:
        if count == 20:
            print("  Scanning", end="", flush=True, file=out)
        print(".", end="", flush=True, file=out)
```

## Running the full test suite

```bash
cd outlook-cli-skill
python3 -m venv .venv
.venv/bin/pip install pytest rich markdownify beautifulsoup4
.venv/bin/python -m pytest tests/ -v
```

All 52 tests should pass on Linux/WSL without pywin32 installed.

## Windows smoke test

On a Windows machine with Outlook running:

```bash
python scripts/smoke-test.py
```

This exercises: folders JSON (no count field), search --limit, reply --cc
(draft), --help output for all new flags, and a real reply draft with CC.

## Pitfalls

### Don't use GetDefaultFolder.return_value with side_effect
If the mock namespace uses `side_effect` for `GetDefaultFolder`,
`.return_value` returns a `MagicMock` that ISN'T the folder returned
by the side_effect function. Expose the real folder via an attribute
(`ns._inbox_folder = inbox_folder`) so tests can inject items.

### Shared mock folders double results
Default search scans both Inbox AND Sent Items. If both
`GetDefaultFolder(6)` and `GetDefaultFolder(5)` return the same mock,
results double. Use separate mocks — inbox with test data, sent empty.

### count=0 fires progress_dot
`0 % 20 == 0` is True, so `_print_progress_dot(0)` prints ".". In
practice, `scanned` starts at 0 and increments to 1 before the callback,
so count=0 never reaches the function — but the edge case is real.
