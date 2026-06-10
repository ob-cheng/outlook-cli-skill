"""Shared test fixtures and mocks for outlook-cli tests."""

import pytest
from unittest.mock import MagicMock, PropertyMock


@pytest.fixture
def mock_mail_item():
    """Create a mock COM MailItem with all properties from_mail_item needs."""
    mail = MagicMock()
    mail.Class = 43  # olMail
    mail.Subject = "Test Subject"
    mail.SenderName = "Alice"
    mail.SenderEmailAddress = "alice@example.com"
    mail.CC = ""
    mail.SentOn = None
    mail.ReceivedTime = None
    mail.HTMLBody = "<html><body>Test HTML</body></html>"
    mail.Body = "Test body text"
    mail.UnRead = True
    mail.Importance = "1"  # normal
    mail.EntryID = "test-entry-id-123"

    # Recipients mock
    mock_recip = MagicMock()
    mock_recip.Address = "bob@example.com"
    mock_recip.Name = "Bob"
    mail.Recipients.Count = 1
    mail.Recipients.Item.return_value = mock_recip

    # Attachments mock
    mail.Attachments.Count = 0

    return mail


@pytest.fixture
def mock_items_collection(mock_mail_item):
    """Create a mock Outlook Items collection that returns mock_mail_item."""
    items = MagicMock()
    items.Count = 1
    items.Item.return_value = mock_mail_item
    items.Restrict.return_value = items  # Restrict returns same collection for simplicity
    items.Sort = MagicMock()
    return items


@pytest.fixture
def mock_folder(mock_items_collection):
    """Create a mock Outlook folder."""
    folder = MagicMock()
    folder.Items = mock_items_collection
    folder.Name = "Inbox"
    return folder


def build_empty_items():
    """Build an empty Items collection (helper, not a fixture)."""
    items = MagicMock()
    items.Count = 0
    items.Restrict.return_value = items
    items.Sort = MagicMock()
    return items


@pytest.fixture
def mock_namespace(mock_items_collection):
    """Create a mock Outlook MAPI namespace with separate Inbox and Sent folders."""
    ns = MagicMock()

    # Build two separate folders so default search (Inbox + Sent) doesn't
    # return duplicate results from the same mock folder.
    inbox_folder = MagicMock()
    inbox_folder.Items = mock_items_collection
    inbox_folder.Name = "Inbox"

    # Empty sent items by default — tests that want sent items
    # can override this fixture or add sent items separately.
    sent_items = MagicMock()
    sent_items.Items = build_empty_items()
    sent_items.Name = "Sent Items"

    def get_default_folder(folder_id):
        if folder_id == 6:  # olFolderInbox
            return inbox_folder
        elif folder_id == 5:  # olFolderSentMail
            return sent_items
        return MagicMock()

    ns.GetDefaultFolder.side_effect = get_default_folder
    # Expose the inbox folder so tests can inject items into it
    ns._inbox_folder = inbox_folder
    return ns
