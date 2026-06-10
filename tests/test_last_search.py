"""Tests for last search result cache.

Uses temporary directories and mock Email objects.
"""

import json
import pytest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch, MagicMock

from outlook_cli.core import last_search
from outlook_cli.core.models import Email


def make_test_email(
    message_id="test-id-123",
    subject="Test Subject",
    sender_clean="Alice",
    date=None,
):
    """Create a minimal test Email for caching."""
    return Email(
        message_id=message_id,
        subject=subject,
        sender=f"{sender_clean} <alice@example.com>",
        sender_clean=sender_clean,
        sender_domain="example.com",
        sender_smtp="alice@example.com",
        to="bob@example.com",
        to_names=["Bob"],
        to_emails=["bob@example.com"],
        cc=None,
        cc_emails=[],
        date=date or datetime(2024, 3, 15, 10, 30),
        html_body=None,
        text_body="Test",
    )


class TestSaveLastSearch:
    """Tests for save_last_search()."""

    def test_saves_search_results(self, tmp_path):
        """Saves search results to JSON file."""
        cache_file = tmp_path / "last_search.json"
        emails = [
            make_test_email(message_id="id-1", subject="Email 1"),
            make_test_email(message_id="id-2", subject="Email 2"),
        ]

        with patch.object(last_search, 'CONFIG_DIR', tmp_path):
            with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
                last_search.save_last_search(emails)

        assert cache_file.exists()
        data = json.loads(cache_file.read_text())
        assert data["count"] == 2
        assert len(data["results"]) == 2
        assert data["results"][0]["message_id"] == "id-1"

    def test_does_not_save_empty_list(self, tmp_path):
        """Does not save when email list is empty."""
        cache_file = tmp_path / "last_search.json"

        with patch.object(last_search, 'CONFIG_DIR', tmp_path):
            with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
                last_search.save_last_search([])

        assert not cache_file.exists()

    def test_caps_at_50_results(self, tmp_path):
        """Caps saved results at 50."""
        cache_file = tmp_path / "last_search.json"
        emails = [make_test_email(message_id=f"id-{i}") for i in range(100)]

        with patch.object(last_search, 'CONFIG_DIR', tmp_path):
            with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
                last_search.save_last_search(emails)

        data = json.loads(cache_file.read_text())
        assert data["count"] == 100  # Original count
        assert len(data["results"]) == 50  # Capped

    def test_creates_config_dir(self, tmp_path):
        """Creates config directory if it doesn't exist."""
        config_dir = tmp_path / "new_dir"
        cache_file = config_dir / "last_search.json"
        emails = [make_test_email()]

        with patch.object(last_search, 'CONFIG_DIR', config_dir):
            with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
                last_search.save_last_search(emails)

        assert config_dir.exists()

    def test_includes_timestamp(self, tmp_path):
        """Saves timestamp with results."""
        cache_file = tmp_path / "last_search.json"
        emails = [make_test_email()]

        with patch.object(last_search, 'CONFIG_DIR', tmp_path):
            with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
                last_search.save_last_search(emails)

        data = json.loads(cache_file.read_text())
        assert "timestamp" in data
        # Should be parseable as ISO datetime
        datetime.fromisoformat(data["timestamp"])


class TestGetLastMessageId:
    """Tests for get_last_message_id()."""

    def test_returns_first_message_id(self, tmp_path):
        """Returns first message ID by default."""
        cache_file = tmp_path / "last_search.json"
        cache_file.write_text(json.dumps({
            "timestamp": "2024-03-15T10:30:00",
            "count": 2,
            "results": [
                {"message_id": "first-id", "subject": "First"},
                {"message_id": "second-id", "subject": "Second"},
            ]
        }))

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_message_id()

        assert result == "first-id"

    def test_returns_indexed_message_id(self, tmp_path):
        """Returns message ID at specified index."""
        cache_file = tmp_path / "last_search.json"
        cache_file.write_text(json.dumps({
            "results": [
                {"message_id": "first-id"},
                {"message_id": "second-id"},
                {"message_id": "third-id"},
            ]
        }))

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_message_id(index=2)

        assert result == "third-id"

    def test_returns_none_when_no_file(self, tmp_path):
        """Returns None when cache file doesn't exist."""
        cache_file = tmp_path / "nonexistent.json"

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_message_id()

        assert result is None

    def test_returns_none_for_invalid_index(self, tmp_path):
        """Returns None when index is out of range."""
        cache_file = tmp_path / "last_search.json"
        cache_file.write_text(json.dumps({
            "results": [{"message_id": "only-one"}]
        }))

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_message_id(index=5)

        assert result is None

    def test_returns_none_for_invalid_json(self, tmp_path):
        """Returns None for invalid JSON."""
        cache_file = tmp_path / "last_search.json"
        cache_file.write_text("not valid json")

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_message_id()

        assert result is None

    def test_returns_none_for_missing_results_key(self, tmp_path):
        """Returns None when results key is missing."""
        cache_file = tmp_path / "last_search.json"
        cache_file.write_text(json.dumps({"timestamp": "2024-03-15"}))

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_message_id()

        assert result is None


class TestGetLastSearchInfo:
    """Tests for get_last_search_info()."""

    def test_returns_full_data(self, tmp_path):
        """Returns full search data."""
        cache_file = tmp_path / "last_search.json"
        test_data = {
            "timestamp": "2024-03-15T10:30:00",
            "count": 3,
            "results": [
                {"message_id": "id-1", "subject": "Email 1"},
                {"message_id": "id-2", "subject": "Email 2"},
            ]
        }
        cache_file.write_text(json.dumps(test_data))

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_search_info()

        assert result == test_data

    def test_returns_none_when_no_file(self, tmp_path):
        """Returns None when cache file doesn't exist."""
        cache_file = tmp_path / "nonexistent.json"

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_search_info()

        assert result is None

    def test_returns_none_for_invalid_json(self, tmp_path):
        """Returns None for invalid JSON."""
        cache_file = tmp_path / "last_search.json"
        cache_file.write_text("not valid json")

        with patch.object(last_search, 'LAST_SEARCH_FILE', cache_file):
            result = last_search.get_last_search_info()

        assert result is None
