"""Folder tree cache to avoid expensive COM namespace walks.

The COM namespace walk (list_all_folders) can take 16-35 seconds on
multi-account setups with deeply nested folders. This cache serializes
the folder tree to disk after the first walk and reuses it on subsequent
calls, invalidating only when the store topology changes.

Cache location: ~/.outlook-cli/folder-cache.json
Fingerprint: sorted store names + store count — cheap to compute, catches
  store additions/removals/renames without walking the full tree.
"""

import json
import time
from pathlib import Path
from typing import Optional

CACHE_DIR = Path.home() / ".outlook-cli"
CACHE_FILE = CACHE_DIR / "folder-cache.json"

# Auto-refresh after 24 hours even if fingerprint matches.
# Outlook folders rarely change, but a periodic refresh guards against
# silent drift (e.g. Outlook auto-created folders during an update).
MAX_CACHE_AGE_SECONDS = 86400  # 24 hours


def _compute_fingerprint(namespace) -> str:
    """Compute a lightweight store fingerprint from namespace.Folders.

    Only accesses store names and count — no recursive walk.
    This is the fast path that determines whether the cache is valid.
    """
    store_names = []
    try:
        for i in range(1, namespace.Folders.Count + 1):
            try:
                store = namespace.Folders.Item(i)
                store_names.append(store.Name)
            except Exception:
                continue
    except Exception:
        pass
    return "|".join(sorted(store_names)) + f"|count={len(store_names)}"


def load_cache(namespace) -> Optional[list[dict]]:
    """Load cached folder tree if fingerprint matches and age is acceptable.

    Args:
        namespace: Outlook MAPI namespace (for fingerprint check).

    Returns:
        List of folder dicts if cache is valid, None otherwise.
    """
    if not CACHE_FILE.exists():
        return None

    try:
        with open(CACHE_FILE) as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return None

    # Check fingerprint
    current_fp = _compute_fingerprint(namespace)
    if data.get("fingerprint") != current_fp:
        return None

    # Check age
    cached_at = data.get("cached_at", 0)
    if time.time() - cached_at > MAX_CACHE_AGE_SECONDS:
        return None

    return data.get("folders")


def save_cache(folders: list[dict], namespace) -> None:
    """Save folder tree to cache with current fingerprint.

    Args:
        folders: List of folder dicts from list_all_folders.
        namespace: Outlook MAPI namespace (for fingerprint).
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    data = {
        "fingerprint": _compute_fingerprint(namespace),
        "cached_at": time.time(),
        "folders": folders,
    }
    with open(CACHE_FILE, "w") as f:
        json.dump(data, f, indent=2)
        f.write("\n")


def invalidate_cache() -> None:
    """Force-delete the cache file. Used by --refresh flag."""
    try:
        CACHE_FILE.unlink(missing_ok=True)
    except Exception:
        pass
