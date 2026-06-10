"""Tests for search service — limit, progress callback, filtering logic.

These use monkey-patching on Email.from_mail_item to avoid COM dependencies.
"""

import pytest
from unittest.mock import MagicMock, patch, call

from outlook_cli.services.search import SearchService
from outlook_cli.core.models import Email


# ---------------------------------------------------------------------------
# Factory for test Email objects
# ---------------------------------------------------------------------------


def make_email(
    subject="Test Subject",
    sender="Alice <alice@example.com>",
    sender_smtp="alice@example.com",
    sender_domain="example.com",
    text_body="Test body",
    to_emails=None,
    cc_emails=None,
    is_sent=False,
):
    """Create an Email with the minimum fields needed for search filtering."""
    return Email(
        subject=subject,
        sender=sender,
        sender_clean="Alice",
        sender_domain=sender_domain,
        sender_smtp=sender_smtp,
        to="Bob <bob@example.com>",
        cc="",
        to_names=["Bob"],
        to_emails=to_emails or ["bob@example.com"],
        cc_emails=cc_emails or [],
        date=None,
        html_body=None,
        text_body=text_body,
        is_read=False,
        importance="normal",
        has_attachments=False,
        message_id="test-id",
        is_sent=is_sent,
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def build_mock_items(mail_items, count=None):
    """Build a mock Items collection that returns pre-built Email objects.

    Uses monkey-patching so Email.from_mail_item isn't actually called —
    instead we return pre-built Email objects directly from Item(i).

    Args:
        mail_items: List of Email objects to return from Item(i)
        count: Override Items.Count (defaults to len(mail_items))
    """
    items = MagicMock()
    items.Count = count if count is not None else len(mail_items)

    def item_side_effect(i):
        # Return a mock COM object with Class=43 so the filter passes
        mock = MagicMock()
        mock.Class = 43
        # Store the email for the monkeypatch to return
        mock._test_email = mail_items[i - 1]
        return mock

    items.Item.side_effect = item_side_effect
    items.Restrict.return_value = items
    items.Sort = MagicMock()
    return items


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


class TestSearchLimit:
    """Tests for --limit functionality."""

    def test_limit_stops_at_n(self, mock_namespace):
        """With limit=3 and 10 matching emails, should return exactly 3."""
        emails = [make_email(subject=f"Email {i}") for i in range(10)]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            def from_mail_side_effect(mail, is_sent=False):
                return mail._test_email
            mock_from.side_effect = from_mail_side_effect

            result = svc.search(limit=3)
            assert len(result) == 3

    def test_limit_none_returns_all(self, mock_namespace):
        """Without limit, should return all matching emails."""
        emails = [make_email() for _ in range(5)]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            result = svc.search(limit=None)
            assert len(result) == 5

    def test_limit_zero_returns_empty(self, mock_namespace):
        """limit=0 should return empty list immediately."""
        svc = SearchService(mock_namespace)
        result = svc.search(limit=0)
        assert result == []

    def test_limit_larger_than_results(self, mock_namespace):
        """If limit exceeds available emails, return all of them."""
        emails = [make_email() for _ in range(2)]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            result = svc.search(limit=100)
            assert len(result) == 2


class TestProgressCallback:
    """Tests for progress callback during search."""

    def test_callback_fires_for_each_item(self, mock_namespace):
        """Progress callback should be called once per scanned item."""
        emails = [make_email() for _ in range(5)]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)
        calls = []

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            svc.search(progress_callback=lambda c: calls.append(c))

        # Should be called with 1, 2, 3, 4, 5
        assert calls == [1, 2, 3, 4, 5]

    def test_callback_stops_with_limit(self, mock_namespace):
        """Progress callback should only fire for items up to the limit."""
        emails = [make_email() for _ in range(10)]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)
        calls = []

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            svc.search(limit=3, progress_callback=lambda c: calls.append(c))

        # Only 3 items should be scanned before stopping
        assert calls == [1, 2, 3]


class TestSearchFiltering:
    """Tests for search filtering logic."""

    def test_auto_reply_skipped(self, mock_namespace):
        """Auto-reply subjects should be filtered out."""
        emails = [
            make_email(subject="Automatic reply: Out of Office"),
            make_email(subject="Real Email"),
            make_email(subject="Out of Office: Vacation"),
        ]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            result = svc.search()

        assert len(result) == 1
        assert result[0].subject == "Real Email"

    def test_keyword_filter_matches_subject(self, mock_namespace):
        """Keyword filter should match in subject (whole word)."""
        emails = [
            make_email(subject="Invoice #1234"),
            make_email(subject="Meeting notes"),
        ]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            result = svc.search(filter_keyword="Invoice")

        assert len(result) == 1
        assert result[0].subject == "Invoice #1234"

    def test_keyword_filter_matches_body(self, mock_namespace):
        """Keyword filter should match in text body."""
        emails = [
            make_email(subject="Hello", text_body="Please find the invoice attached"),
            make_email(subject="Hi", text_body="Just checking in"),
        ]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            result = svc.search(filter_keyword="invoice")

        assert len(result) == 1

    def test_email_filter_matches_sender(self, mock_namespace):
        """Filter by email should match sender."""
        emails = [
            make_email(sender_smtp="alice@co.com"),
            make_email(sender_smtp="bob@other.com"),
        ]
        inbox = mock_namespace._inbox_folder
        inbox.Items = build_mock_items(emails)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
            result = svc.search(filter_emails=["alice@co.com"])

        assert len(result) == 1
        assert result[0].sender_smtp == "alice@co.com"


class TestSearchMultiFolder:
    """Tests for multi-folder search with named folders."""

    def test_remaining_tracks_across_folders(self, mock_namespace, mock_folder):
        """When using multiple --folder flags, remaining limit carries over."""
        emails_a = [make_email(subject="A1"), make_email(subject="A2")]
        emails_b = [make_email(subject="B1"), make_email(subject="B2")]

        folder_a = MagicMock()
        folder_a.Items = build_mock_items(emails_a)

        folder_b = MagicMock()
        folder_b.Items = build_mock_items(emails_b)

        svc = SearchService(mock_namespace)

        with patch.object(Email, "from_mail_item") as mock_from:
            mock_from.side_effect = lambda mail, is_sent=False: mail._test_email

            # Mock _find_folder to return our test folders
            with patch.object(svc, "_find_folder") as mock_find:
                mock_find.side_effect = [folder_a, folder_b]
                result = svc.search(folders=["FolderA", "FolderB"], limit=3)

        assert len(result) == 3


class TestSearchServiceClass:
    """Test SearchService class-level concerns."""

    def test_imports_callable(self):
        """Verify Callable was added to imports."""
        from outlook_cli.services.search import Callable  # noqa: F811
        assert Callable is not None
