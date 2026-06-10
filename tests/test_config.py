"""Tests for ConfigManager — file I/O, defaults, merging, status tags.

These tests use temporary files to avoid touching real config.
"""

import json
import pytest
from pathlib import Path

from outlook_cli.core.config import ConfigManager, DEFAULTS


class TestConfigManagerLoad:
    """Tests for loading configuration."""

    def test_load_returns_defaults_when_no_file(self, tmp_path):
        """When config file doesn't exist, returns defaults."""
        config_path = tmp_path / "config.json"
        mgr = ConfigManager(config_path=config_path)

        result = mgr.load()

        assert result == DEFAULTS
        assert result["send_mode"] == "draft"
        assert result["draft_instructions"] == ""
        assert result["humanizer_enabled"] is False

    def test_load_merges_with_defaults(self, tmp_path):
        """Existing config merges with defaults for missing keys."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "send"}')
        mgr = ConfigManager(config_path=config_path)

        result = mgr.load()

        assert result["send_mode"] == "send"
        assert result["draft_instructions"] == ""  # from defaults
        assert result["humanizer_enabled"] is False  # from defaults

    def test_load_handles_invalid_json(self, tmp_path):
        """Invalid JSON falls back to defaults."""
        config_path = tmp_path / "config.json"
        config_path.write_text("not valid json {{{")
        mgr = ConfigManager(config_path=config_path)

        result = mgr.load()

        assert result == DEFAULTS

    def test_load_handles_empty_file(self, tmp_path):
        """Empty file falls back to defaults."""
        config_path = tmp_path / "config.json"
        config_path.write_text("")
        mgr = ConfigManager(config_path=config_path)

        result = mgr.load()

        assert result == DEFAULTS

    def test_load_preserves_extra_keys(self, tmp_path):
        """Config with extra keys preserves them."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "send", "custom_key": "custom_value"}')
        mgr = ConfigManager(config_path=config_path)

        result = mgr.load()

        assert result["send_mode"] == "send"
        assert result["custom_key"] == "custom_value"


class TestConfigManagerSave:
    """Tests for saving configuration."""

    def test_save_creates_file_and_dir(self, tmp_path):
        """Save creates parent directories and file."""
        config_path = tmp_path / "subdir" / "config.json"
        mgr = ConfigManager(config_path=config_path)

        mgr.save({"send_mode": "send"})

        assert config_path.exists()
        data = json.loads(config_path.read_text())
        assert data["send_mode"] == "send"

    def test_save_preserves_existing_keys(self, tmp_path):
        """Save merges with existing config."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "draft", "custom": "value"}')
        mgr = ConfigManager(config_path=config_path)

        mgr.save({"humanizer_enabled": True})

        data = json.loads(config_path.read_text())
        assert data["send_mode"] == "draft"  # preserved
        assert data["custom"] == "value"  # preserved
        assert data["humanizer_enabled"] is True  # new

    def test_save_overwrites_existing_key(self, tmp_path):
        """Save updates existing keys."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "draft"}')
        mgr = ConfigManager(config_path=config_path)

        mgr.save({"send_mode": "send"})

        data = json.loads(config_path.read_text())
        assert data["send_mode"] == "send"


class TestConfigManagerGetSet:
    """Tests for get/set single values."""

    def test_get_returns_value(self, tmp_path):
        """Get returns the value for a key."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "send"}')
        mgr = ConfigManager(config_path=config_path)

        assert mgr.get("send_mode") == "send"

    def test_get_returns_default_for_missing_key(self, tmp_path):
        """Get returns default when key not in file."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{}')
        mgr = ConfigManager(config_path=config_path)

        assert mgr.get("send_mode") == "draft"

    def test_get_returns_none_for_unknown_key(self, tmp_path):
        """Get returns None for keys not in defaults."""
        config_path = tmp_path / "config.json"
        mgr = ConfigManager(config_path=config_path)

        assert mgr.get("nonexistent_key") is None

    def test_set_creates_key(self, tmp_path):
        """Set creates a new key."""
        config_path = tmp_path / "config.json"
        mgr = ConfigManager(config_path=config_path)

        mgr.set("send_mode", "send")

        assert mgr.get("send_mode") == "send"

    def test_set_updates_existing_key(self, tmp_path):
        """Set updates an existing key."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "draft"}')
        mgr = ConfigManager(config_path=config_path)

        mgr.set("send_mode", "send")

        assert mgr.get("send_mode") == "send"


class TestConfigManagerClear:
    """Tests for clearing (resetting) keys."""

    def test_clear_resets_to_default(self, tmp_path):
        """Clear resets a key to its default value."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "send"}')
        mgr = ConfigManager(config_path=config_path)

        mgr.clear("send_mode")

        assert mgr.get("send_mode") == "draft"

    def test_clear_unknown_key_sets_none(self, tmp_path):
        """Clear unknown key sets it to None."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"custom": "value"}')
        mgr = ConfigManager(config_path=config_path)

        mgr.clear("custom")

        data = json.loads(config_path.read_text())
        assert data["custom"] is None


class TestConfigManagerShow:
    """Tests for show() method."""

    def test_show_returns_full_config(self, tmp_path):
        """Show returns the complete config dict."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "send", "draft_instructions": "Be formal"}')
        mgr = ConfigManager(config_path=config_path)

        result = mgr.show()

        assert result["send_mode"] == "send"
        assert result["draft_instructions"] == "Be formal"
        assert result["humanizer_enabled"] is False  # from defaults


class TestConfigManagerStatusTags:
    """Tests for status_tags() — agent-visible status indicators."""

    def test_status_tags_draft_mode(self, tmp_path):
        """Draft mode produces draft tag."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "draft"}')
        mgr = ConfigManager(config_path=config_path)

        tags = mgr.status_tags()

        assert "[send mode: draft]" in tags
        assert "[send mode: direct]" not in tags

    def test_status_tags_send_mode(self, tmp_path):
        """Send mode produces direct tag."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"send_mode": "send"}')
        mgr = ConfigManager(config_path=config_path)

        tags = mgr.status_tags()

        assert "[send mode: direct]" in tags
        assert "[send mode: draft]" not in tags

    def test_status_tags_draft_instructions_enabled(self, tmp_path):
        """Draft instructions produces tag when set."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"draft_instructions": "Be formal and professional"}')
        mgr = ConfigManager(config_path=config_path)

        tags = mgr.status_tags()

        assert "[draft instructions enabled]" in tags

    def test_status_tags_draft_instructions_empty(self, tmp_path):
        """Empty draft instructions produces no tag."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"draft_instructions": ""}')
        mgr = ConfigManager(config_path=config_path)

        tags = mgr.status_tags()

        assert "[draft instructions enabled]" not in tags

    def test_status_tags_humanizer_enabled(self, tmp_path):
        """Humanizer enabled produces tag."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"humanizer_enabled": true}')
        mgr = ConfigManager(config_path=config_path)

        tags = mgr.status_tags()

        assert "[humanizer enabled]" in tags

    def test_status_tags_humanizer_disabled(self, tmp_path):
        """Humanizer disabled produces no tag."""
        config_path = tmp_path / "config.json"
        config_path.write_text('{"humanizer_enabled": false}')
        mgr = ConfigManager(config_path=config_path)

        tags = mgr.status_tags()

        assert "[humanizer enabled]" not in tags

    def test_status_tags_all_features_enabled(self, tmp_path):
        """All features enabled produces all tags."""
        config_path = tmp_path / "config.json"
        config_path.write_text(json.dumps({
            "send_mode": "send",
            "draft_instructions": "Be formal",
            "humanizer_enabled": True
        }))
        mgr = ConfigManager(config_path=config_path)

        tags = mgr.status_tags()

        assert len(tags) == 3
        assert "[send mode: direct]" in tags
        assert "[draft instructions enabled]" in tags
        assert "[humanizer enabled]" in tags
