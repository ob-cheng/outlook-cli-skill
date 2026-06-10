"""Tests for folder traversal — list_all_folders and the count removal."""

import pytest
from unittest.mock import MagicMock, PropertyMock

from outlook_cli.core.folders import list_all_folders


def make_mock_folder(name, subfolders=None, item_count=0):
    """Build a mock Outlook folder with optional subfolders."""
    folder = MagicMock()
    folder.Name = name
    folder.Items.Count = item_count
    if subfolders:
        folder.Folders.Count = len(subfolders)

        def item_side_effect(i):
            return subfolders[i - 1]

        folder.Folders.Item.side_effect = item_side_effect
    else:
        folder.Folders.Count = 0
    return folder


def make_mock_namespace(stores):
    """Build a mock namespace with stores."""

    namespace = MagicMock()
    namespace.Folders.Count = len(stores)

    def store_side_effect(i):
        return stores[i - 1]

    namespace.Folders.Item.side_effect = store_side_effect
    return namespace


class TestListAllFolders:
    def test_flat_structure(self):
        """Single store with no subfolders."""
        store = make_mock_folder("account@domain.com (Account)")
        ns = make_mock_namespace([store])

        result = list_all_folders(ns)

        assert len(result) == 1
        assert result[0]["name"] == "account@domain.com (Account)"
        assert result[0]["is_store"] is True
        assert result[0]["level"] == 0

    def test_nested_folders(self):
        """Store with subfolders."""
        inbox = make_mock_folder("Inbox")
        drafts = make_mock_folder("Drafts")
        store = make_mock_folder("account@domain.com", subfolders=[inbox, drafts])
        ns = make_mock_namespace([store])

        result = list_all_folders(ns)

        assert len(result) == 3
        names = [f["name"] for f in result]
        assert "account@domain.com" in names
        assert "Inbox" in names
        assert "Drafts" in names

    def test_deeply_nested(self):
        """Three levels of nesting."""
        deep = make_mock_folder("Deep")
        mid = make_mock_folder("Mid", subfolders=[deep])
        top = make_mock_folder("Top", subfolders=[mid])
        store = make_mock_folder("store", subfolders=[top])
        ns = make_mock_namespace([store])

        result = list_all_folders(ns)

        names = [f["name"] for f in result]
        assert "Deep" in names
        assert "Mid" in names
        assert "Top" in names

        # Check levels
        deep_entry = next(f for f in result if f["name"] == "Deep")
        assert deep_entry["level"] == 3

    def test_no_count_field(self):
        """Verify 'count' field has been removed (the fix from issue #3)."""
        store = make_mock_folder("test@domain.com (Account)", item_count=42)
        ns = make_mock_namespace([store])

        result = list_all_folders(ns)

        for entry in result:
            assert "count" not in entry, (
                f"'count' field should not be present (was removed to fix folders --json perf). "
                f"Found in: {entry}"
            )

    def test_path_built_correctly(self):
        """Folder paths should be slash-delimited."""
        inbox = make_mock_folder("Inbox")
        store = make_mock_folder("store", subfolders=[inbox])
        ns = make_mock_namespace([store])

        result = list_all_folders(ns)

        inbox_entry = next(f for f in result if f["name"] == "Inbox")
        assert inbox_entry["path"] == "store/Inbox"

    def test_multiple_stores(self):
        """Multiple accounts (stores) should all be traversed."""
        store1 = make_mock_folder("personal@domain.com (Account)")
        store2 = make_mock_folder("work@domain.com (Account)")
        ns = make_mock_namespace([store1, store2])

        result = list_all_folders(ns)

        stores = [f for f in result if f.get("is_store")]
        assert len(stores) == 2

    def test_exception_in_traversal_is_swallowed(self):
        """A bad folder should be skipped, not crash the traversal."""
        good = make_mock_folder("Good")
        bad = MagicMock()
        # Make Name access blow up so the folder can't be added at all
        type(bad).Name = PropertyMock(side_effect=Exception("COM error"))
        bad.Folders = MagicMock()
        bad.Folders.Count = 0

        store = make_mock_folder("store", subfolders=[good, bad])
        ns = make_mock_namespace([store])

        result = list_all_folders(ns)

        names = [f["name"] for f in result]
        assert "store" in names
        assert "Good" in names
        assert "Bad" not in names  # skipped due to exception
