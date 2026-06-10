#!/usr/bin/env python3
"""Windows COM Smoke Tests for outlook-cli.

Run this on a Windows machine with Outlook open to verify the CLI
works end-to-end against real Outlook data.

Usage:
    python scripts/smoke-test.py

Requirements:
    - Windows with Outlook installed and running
    - outlook-cli installed (pip install -e .)
    - At least one configured email account with some emails

Exit code: 0 = all pass, 1 = one or more failures.
"""

import json
import subprocess
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
OUTLOOK_PY = SKILL_DIR / "outlook.py"

PASS = 0
FAIL = 0


def run(cmd_args, description):
    """Run outlook.py with given args, return (success, stdout)."""
    global PASS, FAIL
    full_cmd = [sys.executable, str(OUTLOOK_PY)] + cmd_args
    print(f"\n  [{description}]")
    print(f"  $ outlook {' '.join(cmd_args)}")

    result = subprocess.run(full_cmd, capture_output=True, text=True, timeout=120)

    if result.returncode != 0:
        print(f"  FAIL (exit {result.returncode})")
        if result.stderr:
            print(f"  stderr: {result.stderr[:200]}")
        FAIL += 1
        return False, result.stdout

    PASS += 1
    return True, result.stdout


def check(condition, message):
    """Assert a condition, counting pass/fail."""
    global PASS, FAIL
    if condition:
        PASS += 1
        print(f"    ✓ {message}")
    else:
        FAIL += 1
        print(f"    ✗ {message}")


# ---------------------------------------------------------------------------
# Test 1: folders
# ---------------------------------------------------------------------------
ok, out = run(["folders", "--json"], "List folders (JSON)")
if ok:
    try:
        data = json.loads(out)
        check("folders" in data, "JSON has 'folders' key")
        check(len(data["folders"]) > 0, "At least one folder found")
        # Verify no 'count' field (perf fix from issue #3)
        for f in data["folders"]:
            check("count" not in f, f"Folder '{f.get('name')}' has no 'count' field")
    except json.JSONDecodeError:
        check(False, "Output is valid JSON")
        FAIL += 1

# ---------------------------------------------------------------------------
# Test 2: search --limit
# ---------------------------------------------------------------------------
ok, out = run(["search", "--limit", "3", "--json"], "Search with --limit 3")
if ok:
    try:
        data = json.loads(out)
        check("count" in data, "JSON has 'count' key")
        check(data["count"] <= 3, f"Count <= 3 (got {data.get('count')})")
    except json.JSONDecodeError:
        check(False, "Output is valid JSON")

# ---------------------------------------------------------------------------
# Test 3: search without limit (verify it works)
# ---------------------------------------------------------------------------
ok, out = run(["search", "--days", "1", "--json"], "Search last 1 day")
if ok:
    try:
        data = json.loads(out)
        check("count" in data, "JSON has 'count' key")
        check(isinstance(data.get("count"), int), "Count is integer")
    except json.JSONDecodeError:
        check(False, "Output is valid JSON")

# ---------------------------------------------------------------------------
# Test 4: reply --help (verifies --cc/--bcc flags exist)
# ---------------------------------------------------------------------------
ok, out = run(["reply", "--help"], "Reply --help")
if ok:
    check("--cc" in out, "--cc flag in help output")
    check("--bcc" in out, "--bcc flag in help output")

# ---------------------------------------------------------------------------
# Test 5: search --help (verifies --limit flag exists)
# ---------------------------------------------------------------------------
ok, out = run(["search", "--help"], "Search --help")
if ok:
    check("--limit" in out, "--limit flag in help output")
    check("-N" in out, "-N short flag in help output")

# ---------------------------------------------------------------------------
# Test 6: export --help (verifies --limit on export)
# ---------------------------------------------------------------------------
ok, out = run(["export", "--help"], "Export --help")
if ok:
    check("--limit" in out, "--limit flag in export help output")

# ---------------------------------------------------------------------------
# Test 7: folders plain text (no --json) — verifies no timeout
# ---------------------------------------------------------------------------
ok, out = run(["folders"], "Folders plain text")
if ok:
    check(len(out) > 0, "Plain text folders output is non-empty")

# ---------------------------------------------------------------------------
# Test 8: reply creates draft with --cc (requires valid message ID from search)
# ---------------------------------------------------------------------------
ok, out = run(["search", "--days", "1", "--limit", "1", "--json"], "Get message ID for reply test")
message_id = None
if ok:
    try:
        data = json.loads(out)
        if data.get("emails"):
            message_id = data["emails"][0].get("message_id")
            print(f"    Got message ID: {message_id[:30]}...")
    except (json.JSONDecodeError, KeyError, IndexError):
        pass

if message_id:
    ok, out = run(
        ["reply", message_id, "--body", "Smoke test reply", "--cc", "smoke-test@example.com", "--json"],
        "Reply with --cc (creates draft)",
    )
    if ok:
        try:
            data = json.loads(out)
            check(data.get("draft") is True, "Reply saved as draft (not sent)")
        except json.JSONDecodeError:
            check(False, "Output is valid JSON")
else:
    print("\n  [Reply with --cc] SKIPPED — no messages found in last day")

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
print(f"\n{'=' * 50}")
print(f"Results: {PASS} passed, {FAIL} failed")
if FAIL > 0:
    print("SOME TESTS FAILED")
    sys.exit(1)
else:
    print("All smoke tests passed")
    sys.exit(0)
