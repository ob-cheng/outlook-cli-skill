"""Configuration manager for outlook-cli.

Stores user preferences in ~/.outlook-cli/config.json.
Settings are read by the agent workflow and printed as tags on
send/reply/forward so the agent can't claim it followed rules it didn't.
"""

import json
import os
from pathlib import Path
from typing import Any

CONFIG_DIR = Path.home() / ".outlook-cli"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULTS = {
    "send_mode": "draft",           # "draft" or "send"
    "draft_instructions": "",       # free text, empty = skip
    "humanizer_enabled": False,     # true/false
}


class ConfigManager:
    """Read/write outlook-cli configuration."""

    def __init__(self, config_path: Path | None = None):
        self.config_path = config_path or CONFIG_FILE

    def _ensure_dir(self) -> None:
        self.config_path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> dict[str, Any]:
        """Load config, merging with defaults for any missing keys."""
        if not self.config_path.exists():
            return dict(DEFAULTS)
        try:
            with open(self.config_path) as f:
                data = json.load(f)
            return {**DEFAULTS, **data}
        except (json.JSONDecodeError, OSError):
            return dict(DEFAULTS)

    def save(self, data: dict[str, Any]) -> None:
        """Save config, preserving existing keys not in data."""
        current = self.load()
        current.update(data)
        self._ensure_dir()
        with open(self.config_path, "w") as f:
            json.dump(current, f, indent=2)
            f.write("\n")

    def get(self, key: str) -> Any:
        """Get a single config value."""
        return self.load().get(key, DEFAULTS.get(key))

    def set(self, key: str, value: Any) -> None:
        """Set a single config value."""
        self.save({key: value})

    def clear(self, key: str) -> None:
        """Reset a single key to its default."""
        default = DEFAULTS.get(key)
        self.save({key: default})

    def show(self) -> dict[str, Any]:
        """Return the full config dict."""
        return self.load()

    def status_tags(self) -> list[str]:
        """Return human-readable status tags for agent output.

        Printed as part of send/reply/forward so the agent
        sees what's active and can't silently skip steps.
        """
        cfg = self.load()
        tags = []

        if cfg.get("send_mode") == "send":
            tags.append("[send mode: direct]")
        else:
            tags.append("[send mode: draft]")

        if cfg.get("draft_instructions"):
            tags.append("[draft instructions enabled]")

        if cfg.get("humanizer_enabled"):
            tags.append("[humanizer enabled]")

        return tags
