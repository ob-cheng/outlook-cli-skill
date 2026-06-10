"""Tests for Email model — dataclass, from_mail_item factory, to_dict serialization.

Uses mock COM objects to avoid win32com dependency.
"""

import pytest
from datetime import datetime
from unittest.mock import MagicMock, PropertyMock

from outlook_cli.core.models import Email


def make_mock_mail_item(
    subject="Test Subject",
    sender_name="Alice Smith",
    sender_email="alice@example.com",
    recipients=None,
    cc="",
    sent_on=None,
    received_time=None,
    html_body="<html><body>Test</body></html>",
    text_body="Test body",
    unread=False,
    importance="1",
    entry_id="test-id-123",
    attachments_count=0,
):
    """Create a mock Outlook MailItem COM object."""
    mail = MagicMock()
    mail.Subject = subject
    mail.SenderName = sender_name
    mail.SenderEmailAddress = sender_email
    mail.CC = cc
    mail.SentOn = sent_on
    mail.ReceivedTime = received_time
    mail.HTMLBody = html_body
    mail.Body = text_body
    mail.UnRead = unread
    mail.Importance = importance
    mail.EntryID = entry_id

    # Mock Sender for Exchange handling
    mail.Sender = MagicMock()
    mail.Sender.Type = "SMTP"

    # Recipients mock
    if recipients is None:
        recipients = [("Bob Jones", "bob@example.com", 1)]  # Type 1 = To
    mail.Recipients = MagicMock()
    mail.Recipients.Count = len(recipients)

    def get_recipient(i):
        name, email, recip_type = recipients[i - 1]
        recip = MagicMock()
        recip.Name = name
        recip.Address = email
        recip.Type = recip_type
        recip.AddressEntry = MagicMock()
        recip.AddressEntry.Type = "SMTP"
        return recip

    mail.Recipients.Item = get_recipient

    # Attachments mock
    mail.Attachments = MagicMock()
    mail.Attachments.Count = attachments_count

    return mail


class TestEmailDataclass:
    """Tests for Email dataclass basics."""

    def test_email_creation(self):
        """Email can be created with all fields."""
        email = Email(
            subject="Test",
            sender="Alice <alice@example.com>",
            sender_clean="Alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="Bob <bob@example.com>",
            to_names=["Bob"],
            to_emails=["bob@example.com"],
            cc=None,
            cc_emails=[],
            date=datetime(2024, 3, 15, 10, 30),
            html_body="<p>Hello</p>",
            text_body="Hello",
        )

        assert email.subject == "Test"
        assert email.sender_smtp == "alice@example.com"
        assert email.to_emails == ["bob@example.com"]

    def test_email_defaults(self):
        """Email has sensible defaults."""
        email = Email(
            subject="Test",
            sender="alice@example.com",
            sender_clean="alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="bob@example.com",
            to_names=[],
            to_emails=[],
            cc=None,
            cc_emails=[],
            date=None,
            html_body=None,
            text_body=None,
        )

        assert email.is_sent is False
        assert email.is_read is True
        assert email.importance == "normal"
        assert email.has_attachments is False
        assert email.message_id is None


class TestEmailFromMailItem:
    """Tests for Email.from_mail_item factory method."""

    def test_basic_conversion(self):
        """Converts basic mail item correctly."""
        mail = make_mock_mail_item()

        email = Email.from_mail_item(mail)

        assert email.subject == "Test Subject"
        assert "alice@example.com" in email.sender.lower()
        assert email.message_id == "test-id-123"

    def test_handles_missing_subject(self):
        """Missing subject becomes 'No Subject'."""
        mail = make_mock_mail_item(subject=None)

        email = Email.from_mail_item(mail)

        assert email.subject == "No Subject"

    def test_handles_exception_in_subject(self):
        """Exception reading subject falls back to 'No Subject'."""
        mail = make_mock_mail_item()
        type(mail).Subject = PropertyMock(side_effect=Exception("COM error"))

        email = Email.from_mail_item(mail)

        assert email.subject == "No Subject"

    def test_is_sent_flag(self):
        """is_sent flag is passed through."""
        mail = make_mock_mail_item()

        email = Email.from_mail_item(mail, is_sent=True)

        assert email.is_sent is True

    def test_unread_to_is_read(self):
        """UnRead=True becomes is_read=False."""
        mail = make_mock_mail_item(unread=True)

        email = Email.from_mail_item(mail)

        assert email.is_read is False

    def test_unread_false_to_is_read(self):
        """UnRead=False becomes is_read=True."""
        mail = make_mock_mail_item(unread=False)

        email = Email.from_mail_item(mail)

        assert email.is_read is True

    def test_importance_high(self):
        """Importance 2 becomes 'high'."""
        mail = make_mock_mail_item(importance="2")

        email = Email.from_mail_item(mail)

        assert email.importance == "high"

    def test_importance_low(self):
        """Importance 0 becomes 'low'."""
        mail = make_mock_mail_item(importance="0")

        email = Email.from_mail_item(mail)

        assert email.importance == "low"

    def test_importance_normal(self):
        """Importance 1 becomes 'normal'."""
        mail = make_mock_mail_item(importance="1")

        email = Email.from_mail_item(mail)

        assert email.importance == "normal"

    def test_has_attachments_true(self):
        """Attachments.Count > 0 sets has_attachments=True."""
        mail = make_mock_mail_item(attachments_count=2)

        email = Email.from_mail_item(mail)

        assert email.has_attachments is True

    def test_has_attachments_false(self):
        """Attachments.Count == 0 sets has_attachments=False."""
        mail = make_mock_mail_item(attachments_count=0)

        email = Email.from_mail_item(mail)

        assert email.has_attachments is False

    def test_extracts_body_content(self):
        """Extracts HTML and text body."""
        mail = make_mock_mail_item(
            html_body="<p>HTML content</p>",
            text_body="Plain text content"
        )

        email = Email.from_mail_item(mail)

        assert email.html_body == "<p>HTML content</p>"
        assert email.text_body == "Plain text content"

    def test_handles_body_exception(self):
        """Exception reading body sets None."""
        mail = make_mock_mail_item()
        type(mail).HTMLBody = PropertyMock(side_effect=Exception("COM error"))

        email = Email.from_mail_item(mail)

        assert email.html_body is None

    def test_uses_received_time_for_inbox(self):
        """Uses ReceivedTime when is_sent=False."""
        received = datetime(2024, 3, 15, 10, 30)
        mail = make_mock_mail_item(received_time=received)

        email = Email.from_mail_item(mail, is_sent=False)

        assert email.date == received

    def test_uses_sent_on_for_sent(self):
        """Uses SentOn when is_sent=True."""
        sent = datetime(2024, 3, 15, 10, 30)
        mail = make_mock_mail_item(sent_on=sent)

        email = Email.from_mail_item(mail, is_sent=True)

        assert email.date == sent


class TestEmailToDict:
    """Tests for Email.to_dict serialization."""

    def test_basic_serialization(self):
        """Serializes all fields to dict."""
        email = Email(
            subject="Test Subject",
            sender="Alice <alice@example.com>",
            sender_clean="Alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="Bob <bob@example.com>",
            to_names=["Bob"],
            to_emails=["bob@example.com"],
            cc="Charlie <charlie@example.com>",
            cc_emails=["charlie@example.com"],
            date=datetime(2024, 3, 15, 10, 30),
            html_body="<p>Hello</p>",
            text_body="Hello",
            is_sent=False,
            is_read=True,
            importance="normal",
            has_attachments=False,
            message_id="test-id",
        )

        result = email.to_dict()

        assert result["subject"] == "Test Subject"
        assert result["sender_smtp"] == "alice@example.com"
        assert result["to_emails"] == ["bob@example.com"]
        assert result["cc_emails"] == ["charlie@example.com"]
        assert result["message_id"] == "test-id"
        assert result["date"] == "2024-03-15T10:30:00"

    def test_include_body_true(self):
        """include_body=True includes both bodies."""
        email = Email(
            subject="Test",
            sender="alice@example.com",
            sender_clean="alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="bob@example.com",
            to_names=[],
            to_emails=[],
            cc=None,
            cc_emails=[],
            date=None,
            html_body="<p>HTML</p>",
            text_body="Text",
        )

        result = email.to_dict(include_body=True)

        assert "text_body" in result
        assert "html_body" in result
        assert result["text_body"] == "Text"
        assert result["html_body"] == "<p>HTML</p>"

    def test_include_body_false(self):
        """include_body=False omits bodies."""
        email = Email(
            subject="Test",
            sender="alice@example.com",
            sender_clean="alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="bob@example.com",
            to_names=[],
            to_emails=[],
            cc=None,
            cc_emails=[],
            date=None,
            html_body="<p>HTML</p>",
            text_body="Text",
        )

        result = email.to_dict(include_body=False)

        assert "text_body" not in result
        assert "html_body" not in result

    def test_text_only_true(self):
        """text_only=True omits html_body."""
        email = Email(
            subject="Test",
            sender="alice@example.com",
            sender_clean="alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="bob@example.com",
            to_names=[],
            to_emails=[],
            cc=None,
            cc_emails=[],
            date=None,
            html_body="<p>HTML</p>",
            text_body="Text",
        )

        result = email.to_dict(text_only=True)

        assert "text_body" in result
        assert "html_body" not in result

    def test_max_body_lines_truncates(self):
        """max_body_lines truncates text_body."""
        email = Email(
            subject="Test",
            sender="alice@example.com",
            sender_clean="alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="bob@example.com",
            to_names=[],
            to_emails=[],
            cc=None,
            cc_emails=[],
            date=None,
            html_body=None,
            text_body="Line 1\nLine 2\nLine 3\nLine 4\nLine 5",
        )

        result = email.to_dict(max_body_lines=2)

        assert "Line 1" in result["text_body"]
        assert "Line 2" in result["text_body"]
        assert "truncated" in result["text_body"]
        assert "3 more lines" in result["text_body"]

    def test_max_body_lines_no_truncation_needed(self):
        """max_body_lines doesn't truncate short body."""
        email = Email(
            subject="Test",
            sender="alice@example.com",
            sender_clean="alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="bob@example.com",
            to_names=[],
            to_emails=[],
            cc=None,
            cc_emails=[],
            date=None,
            html_body=None,
            text_body="Line 1\nLine 2",
        )

        result = email.to_dict(max_body_lines=5)

        assert result["text_body"] == "Line 1\nLine 2"
        assert "truncated" not in result["text_body"]

    def test_date_none_serializes_to_none(self):
        """None date serializes to None."""
        email = Email(
            subject="Test",
            sender="alice@example.com",
            sender_clean="alice",
            sender_domain="example.com",
            sender_smtp="alice@example.com",
            to="bob@example.com",
            to_names=[],
            to_emails=[],
            cc=None,
            cc_emails=[],
            date=None,
            html_body=None,
            text_body=None,
        )

        result = email.to_dict()

        assert result["date"] is None
