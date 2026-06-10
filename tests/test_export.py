"""Tests for ExportService — markdown/JSON export, state management, content cleaning.

Uses temporary directories and mock Email objects.
"""

import json
import pytest
from datetime import datetime
from pathlib import Path

from outlook_cli.services.export import ExportService, clean_surrogates
from outlook_cli.core.models import Email


def make_test_email(
    subject="Test Subject",
    sender="alice@example.com",
    sender_clean="Alice",
    sender_smtp="alice@example.com",
    sender_domain="example.com",
    to="bob@example.com",
    to_names=None,
    to_emails=None,
    cc=None,
    cc_emails=None,
    date=None,
    html_body="<p>Test content</p>",
    text_body="Test content",
    is_sent=False,
    is_read=True,
    message_id="test-id-123",
):
    """Create a test Email object."""
    return Email(
        subject=subject,
        sender=sender,
        sender_clean=sender_clean,
        sender_smtp=sender_smtp,
        sender_domain=sender_domain,
        to=to,
        to_names=to_names or ["Bob"],
        to_emails=to_emails or ["bob@example.com"],
        cc=cc,
        cc_emails=cc_emails or [],
        date=date or datetime(2024, 3, 15, 10, 30),
        html_body=html_body,
        text_body=text_body,
        is_sent=is_sent,
        is_read=is_read,
        message_id=message_id,
    )


class TestCleanSurrogates:
    """Tests for clean_surrogates helper."""

    def test_clean_normal_text(self):
        """Normal text passes through unchanged."""
        text = "Hello, world!"
        assert clean_surrogates(text) == text

    def test_clean_unicode(self):
        """Unicode text passes through."""
        text = "Hello, 世界! 🌍"
        result = clean_surrogates(text)
        assert "Hello" in result
        assert "世界" in result

    def test_clean_empty(self):
        """Empty string returns empty."""
        assert clean_surrogates("") == ""

    def test_clean_none(self):
        """None returns None."""
        assert clean_surrogates(None) is None


class TestExportServiceInit:
    """Tests for ExportService initialization."""

    def test_creates_output_directory(self, tmp_path):
        """Creates output directory if it doesn't exist."""
        output_dir = tmp_path / "export_output"
        assert not output_dir.exists()

        ExportService(output_dir)

        assert output_dir.exists()

    def test_uses_existing_directory(self, tmp_path):
        """Uses existing output directory."""
        output_dir = tmp_path / "export_output"
        output_dir.mkdir()

        svc = ExportService(output_dir)

        assert svc.output_dir == output_dir


class TestExportServiceState:
    """Tests for get_last_run and save_state."""

    def test_get_last_run_no_file(self, tmp_path):
        """Returns None when no state file exists."""
        svc = ExportService(tmp_path / "out", state_dir=tmp_path)

        result = svc.get_last_run()

        assert result is None

    def test_save_and_get_last_run(self, tmp_path):
        """Saves and retrieves last run timestamp."""
        svc = ExportService(tmp_path / "out", state_dir=tmp_path)

        svc.save_state()
        result = svc.get_last_run()

        assert result is not None
        assert isinstance(result, datetime)

    def test_get_last_run_invalid_json(self, tmp_path):
        """Returns None for invalid JSON state file."""
        state_file = tmp_path / "extraction_state.json"
        state_file.write_text("not valid json")
        svc = ExportService(tmp_path / "out", state_dir=tmp_path)

        result = svc.get_last_run()

        assert result is None

    def test_get_last_run_backwards_compat(self, tmp_path):
        """Handles legacy 'last_extraction' key."""
        state_file = tmp_path / "extraction_state.json"
        state_file.write_text('{"last_extraction": "2024-03-15T10:30:00"}')
        svc = ExportService(tmp_path / "out", state_dir=tmp_path)

        result = svc.get_last_run()

        assert result == datetime(2024, 3, 15, 10, 30)


class TestExportServiceExportEmails:
    """Tests for export_emails method."""

    def test_export_empty_list(self, tmp_path):
        """Exporting empty list returns zero counts."""
        svc = ExportService(tmp_path)

        result = svc.export_emails([])

        assert result["emails_processed"] == 0
        assert result["files_created"] == 0

    def test_export_single_email_markdown(self, tmp_path):
        """Exports single email to markdown."""
        svc = ExportService(tmp_path)
        email = make_test_email()

        result = svc.export_emails([email], group_threads=False)

        assert result["files_created"] == 1
        assert result["emails_processed"] == 1
        assert len(result["files"]) == 1
        assert Path(result["files"][0]).suffix == ".md"

    def test_export_thread_groups_by_subject(self, tmp_path):
        """Groups emails with same subject into thread."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Meeting", date=datetime(2024, 3, 15, 10, 0)),
            make_test_email(subject="RE: Meeting", date=datetime(2024, 3, 15, 11, 0)),
        ]

        result = svc.export_emails(emails, group_threads=True)

        assert result["files_created"] == 1  # One thread file
        assert result["emails_processed"] == 2

    def test_export_no_overwrite_skips_existing(self, tmp_path):
        """no_overwrite skips existing files."""
        svc = ExportService(tmp_path)
        email = make_test_email()

        # First export
        svc.export_emails([email], group_threads=False)
        # Second export with no_overwrite
        result = svc.export_emails([email], group_threads=False, no_overwrite=True)

        assert result["files_skipped"] == 1
        assert result["files_created"] == 0

    def test_export_saves_state_when_requested(self, tmp_path):
        """save_state=True saves timestamp."""
        svc = ExportService(tmp_path, state_dir=tmp_path)
        email = make_test_email()

        svc.export_emails([email], save_state=True)

        assert svc.get_last_run() is not None


class TestExportServiceJsonExport:
    """Tests for JSON format export."""

    def test_export_single_json(self, tmp_path):
        """Exports single email to JSON."""
        svc = ExportService(tmp_path)
        email = make_test_email()

        result = svc.export_emails([email], group_threads=False, output_format='json')

        assert result["files_created"] == 1
        assert Path(result["files"][0]).suffix == ".json"

    def test_export_json_batch(self, tmp_path):
        """Exports all emails to single JSON file."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Email 1", message_id="id-1"),
            make_test_email(subject="Email 2", message_id="id-2"),
        ]

        result = svc.export_emails(emails, output_format='json', batch=True)

        assert result["files_created"] == 1

        # Verify JSON content
        json_file = Path(result["files"][0])
        data = json.loads(json_file.read_text())
        assert data["email_count"] == 2

    def test_export_json_threads(self, tmp_path):
        """Exports threads to JSON."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Meeting", date=datetime(2024, 3, 15, 10, 0)),
            make_test_email(subject="RE: Meeting", date=datetime(2024, 3, 15, 11, 0)),
        ]

        result = svc.export_emails(emails, output_format='json', group_threads=True)

        assert result["files_created"] == 1


class TestExportServiceToJsonData:
    """Tests for to_json_data method (no file write)."""

    def test_to_json_data_empty(self, tmp_path):
        """Empty list returns empty structure."""
        svc = ExportService(tmp_path)

        result = svc.to_json_data([])

        assert result["email_count"] == 0
        assert "export_date" in result

    def test_to_json_data_with_threads(self, tmp_path):
        """Groups emails into threads."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Meeting", date=datetime(2024, 3, 15, 10, 0)),
            make_test_email(subject="RE: Meeting", date=datetime(2024, 3, 15, 11, 0)),
        ]

        result = svc.to_json_data(emails, group_threads=True)

        assert result["thread_count"] == 1
        assert result["email_count"] == 2
        assert "threads" in result

    def test_to_json_data_without_threads(self, tmp_path):
        """Returns flat email list without grouping."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Email 1"),
            make_test_email(subject="Email 2"),
        ]

        result = svc.to_json_data(emails, group_threads=False)

        assert result["email_count"] == 2
        assert "emails" in result
        assert len(result["emails"]) == 2


class TestExportServiceGroupIntoThreads:
    """Tests for thread grouping logic."""

    def test_groups_by_normalized_subject(self, tmp_path):
        """Groups RE:/FW: variants together."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Meeting"),
            make_test_email(subject="RE: Meeting"),
            make_test_email(subject="FW: Meeting"),
            make_test_email(subject="Re: RE: Meeting"),
        ]

        threads = svc._group_into_threads(emails)

        assert len(threads) == 1
        assert "Meeting" in threads
        assert len(threads["Meeting"]) == 4

    def test_different_subjects_separate_threads(self, tmp_path):
        """Different subjects create separate threads."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Meeting"),
            make_test_email(subject="Budget Report"),
        ]

        threads = svc._group_into_threads(emails)

        assert len(threads) == 2


class TestExportServiceContentExtraction:
    """Tests for content extraction and cleaning."""

    def test_extract_html_content(self, tmp_path):
        """Extracts content from HTML body."""
        svc = ExportService(tmp_path)
        email = make_test_email(
            html_body="<html><body><p>Hello, this is a test.</p></body></html>",
            text_body="Hello, this is a test."
        )

        content = svc._extract_content(email)

        assert "Hello" in content
        assert "test" in content

    def test_extract_text_fallback(self, tmp_path):
        """Falls back to text body when no HTML."""
        svc = ExportService(tmp_path)
        email = make_test_email(
            html_body=None,
            text_body="Plain text content here"
        )

        content = svc._extract_content(email)

        assert "Plain text content" in content

    def test_clean_content_removes_images(self, tmp_path):
        """Removes image markdown from content."""
        svc = ExportService(tmp_path)
        content = "Hello ![image](http://example.com/img.png) world"

        cleaned = svc._clean_content(content)

        assert "![image]" not in cleaned
        assert "Hello" in cleaned
        assert "world" in cleaned

    def test_clean_content_removes_external_warning(self, tmp_path):
        """Removes external email warning."""
        svc = ExportService(tmp_path)
        content = "CAUTION: This email originated from outside the organization.\n\nActual content here"

        cleaned = svc._clean_content(content)

        assert "CAUTION" not in cleaned
        assert "Actual content" in cleaned


class TestExportServiceMarkdownGeneration:
    """Tests for markdown file generation."""

    def test_markdown_has_frontmatter(self, tmp_path):
        """Generated markdown includes YAML frontmatter."""
        svc = ExportService(tmp_path)
        email = make_test_email()

        result = svc.export_emails([email], group_threads=False)

        md_file = Path(result["files"][0])
        content = md_file.read_text(encoding='utf-8')

        assert content.startswith("---")
        assert "title:" in content
        assert "date:" in content
        assert "tags:" in content

    def test_thread_markdown_has_message_sections(self, tmp_path):
        """Thread markdown has numbered message sections."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(subject="Meeting", date=datetime(2024, 3, 15, 10, 0)),
            make_test_email(subject="RE: Meeting", date=datetime(2024, 3, 15, 11, 0)),
        ]

        result = svc.export_emails(emails, group_threads=True)

        md_file = Path(result["files"][0])
        content = md_file.read_text(encoding='utf-8')

        assert "## Message 1" in content
        assert "## Message 2" in content


class TestExportServiceParticipants:
    """Tests for participant collection."""

    def test_collect_participants(self, tmp_path):
        """Collects unique participants from emails."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(sender_clean="Alice", to_names=["Bob"]),
            make_test_email(sender_clean="Bob", to_names=["Alice"]),
        ]

        participants = svc._collect_participants(emails)

        assert "Alice" in participants
        assert "Bob" in participants
        assert len(participants) == 2  # No duplicates


class TestExportServiceTags:
    """Tests for tag determination."""

    def test_tags_for_sent_only(self, tmp_path):
        """Sent-only thread gets sent tag."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(is_sent=True),
            make_test_email(is_sent=True),
        ]

        tags = svc._determine_tags(emails)

        assert "email/sent" in tags
        assert "email/conversation" not in tags

    def test_tags_for_received_only(self, tmp_path):
        """Received-only thread gets received tag."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(is_sent=False),
            make_test_email(is_sent=False),
        ]

        tags = svc._determine_tags(emails)

        assert "email/received" in tags
        assert "email/conversation" not in tags

    def test_tags_for_conversation(self, tmp_path):
        """Mixed sent/received gets conversation tag."""
        svc = ExportService(tmp_path)
        emails = [
            make_test_email(is_sent=False),
            make_test_email(is_sent=True),
        ]

        tags = svc._determine_tags(emails)

        assert "email/conversation" in tags


class TestExportServiceDateFormatting:
    """Tests for date formatting."""

    def test_format_date_wikilink(self, tmp_path):
        """Formats date as Obsidian wikilink."""
        svc = ExportService(tmp_path)
        date = datetime(2024, 3, 15, 10, 30)

        result = svc._format_date_wikilink(date)

        assert result == "[[2024-03-15]]"

    def test_format_date_wikilink_none(self, tmp_path):
        """Returns 'Unknown' for None date."""
        svc = ExportService(tmp_path)

        result = svc._format_date_wikilink(None)

        assert result == "Unknown"
