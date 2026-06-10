"""Last search result cache for --last shorthand."""

import json
from pathlib import Path
from datetime import datetime

CONFIG_DIR = Path.home() / ".outlook-cli"
LAST_SEARCH_FILE = CONFIG_DIR / "last_search.json"


def save_last_search(emails: list) -> None:
    """Save the most recent search results for --last lookups.

    Args:
        emails: List of Email objects from a search.
    """
    if not emails:
        return

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    data = {
        "timestamp": datetime.now().isoformat(),
        "count": len(emails),
        "results": [
            {
                "message_id": e.message_id,
                "subject": e.subject,
                "sender_clean": e.sender_clean,
                "date": e.date.isoformat() if e.date else None,
            }
            for e in emails[:50]  # Cap at 50 to avoid huge files
        ],
    }

    with open(LAST_SEARCH_FILE, "w") as f:
        json.dump(data, f, indent=2)


def get_last_message_id(index: int = 0) -> str | None:
    """Get a message ID from the last search results.

    Args:
        index: Which result to return (0 = first/most recent).

    Returns:
        Message ID string, or None if not found.
    """
    if not LAST_SEARCH_FILE.exists():
        return None

    try:
        with open(LAST_SEARCH_FILE) as f:
            data = json.load(f)

        results = data.get("results", [])
        if index < len(results):
            return results[index].get("message_id")
    except (json.JSONDecodeError, OSError, KeyError):
        pass

    return None


def get_last_search_info() -> dict | None:
    """Get info about the last search for display purposes.

    Returns:
        Dict with timestamp, count, and first few results, or None.
    """
    if not LAST_SEARCH_FILE.exists():
        return None

    try:
        with open(LAST_SEARCH_FILE) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None
