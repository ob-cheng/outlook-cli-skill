"""Tests for NotesService and Note dataclass.

Uses mock COM objects to avoid win32com dependency.
"""

import pytest
from datetime import datetime
from unittest.mock import MagicMock, patch

from outlook_cli.services.notes import Note, NotesService, NOTE_COLORS, COLOR_TO_OUTLOOK


def make_mock_note_item(
    subject="Shopping list",
    body="Milk, eggs, bread",
    color=3,  # 0=Blue, 1=Green, 2=Pink, 3=Yellow, 4=White
    created=None,
    modified=None,
    categories="",
    entry_id="note-id-123",
    item_class=44,  # olNote
):
    """Create a mock Outlook NoteItem."""
    note = MagicMock()
    note.Subject = subject
    note.Body = body
    note.Color = color
    note.CreationTime = created or datetime(2024, 3, 15, 10, 0)
    note.LastModificationTime = modified or datetime(2024, 3, 15, 10, 30)
    note.Categories = categories
    note.EntryID = entry_id
    note.Class = item_class

    return note


class TestNoteColorConstants:
    """Tests for note color constants."""

    def test_note_colors_mapping(self):
        """NOTE_COLORS maps Outlook constants to names."""
        assert NOTE_COLORS[0] == "blue"
        assert NOTE_COLORS[1] == "green"
        assert NOTE_COLORS[2] == "pink"
        assert NOTE_COLORS[3] == "yellow"
        assert NOTE_COLORS[4] == "white"

    def test_color_to_outlook_reverse_mapping(self):
        """COLOR_TO_OUTLOOK is reverse of NOTE_COLORS."""
        for num, name in NOTE_COLORS.items():
            assert COLOR_TO_OUTLOOK[name] == num


class TestNoteDataclass:
    """Tests for Note dataclass basics."""

    def test_note_creation(self):
        """Note can be created with all fields."""
        note = Note(
            subject="Shopping list",
            body="Buy groceries",
            color="yellow",
        )

        assert note.subject == "Shopping list"
        assert note.body == "Buy groceries"
        assert note.color == "yellow"

    def test_note_defaults(self):
        """Note has sensible defaults."""
        note = Note(subject="Note")

        assert note.body is None
        assert note.color == "yellow"
        assert note.created is None
        assert note.modified is None
        assert note.categories == []
        assert note.entry_id is None


class TestNoteFromNoteItem:
    """Tests for Note.from_note_item factory."""

    def test_basic_conversion(self):
        """Converts basic note item correctly."""
        note_item = make_mock_note_item()

        note = Note.from_note_item(note_item)

        assert note.subject == "Shopping list"
        assert note.body == "Milk, eggs, bread"
        assert note.entry_id == "note-id-123"

    def test_handles_missing_subject(self):
        """Missing subject becomes '(No Subject)'."""
        note_item = make_mock_note_item(subject=None)

        note = Note.from_note_item(note_item)

        assert note.subject == "(No Subject)"

    def test_handles_missing_body(self):
        """Missing body becomes None."""
        note_item = make_mock_note_item(body=None)

        note = Note.from_note_item(note_item)

        assert note.body is None

    def test_color_blue(self):
        """Color 0 becomes 'blue'."""
        note_item = make_mock_note_item(color=0)

        note = Note.from_note_item(note_item)

        assert note.color == "blue"

    def test_color_green(self):
        """Color 1 becomes 'green'."""
        note_item = make_mock_note_item(color=1)

        note = Note.from_note_item(note_item)

        assert note.color == "green"

    def test_color_pink(self):
        """Color 2 becomes 'pink'."""
        note_item = make_mock_note_item(color=2)

        note = Note.from_note_item(note_item)

        assert note.color == "pink"

    def test_color_yellow(self):
        """Color 3 becomes 'yellow'."""
        note_item = make_mock_note_item(color=3)

        note = Note.from_note_item(note_item)

        assert note.color == "yellow"

    def test_color_white(self):
        """Color 4 becomes 'white'."""
        note_item = make_mock_note_item(color=4)

        note = Note.from_note_item(note_item)

        assert note.color == "white"

    def test_extracts_categories(self):
        """Extracts categories from comma-separated string."""
        note_item = make_mock_note_item(categories="Personal, Ideas")

        note = Note.from_note_item(note_item)

        assert "Personal" in note.categories
        assert "Ideas" in note.categories

    def test_extracts_timestamps(self):
        """Extracts created and modified timestamps."""
        created = datetime(2024, 3, 15, 9, 0)
        modified = datetime(2024, 3, 15, 12, 0)
        note_item = make_mock_note_item(created=created, modified=modified)

        note = Note.from_note_item(note_item)

        assert note.created == created
        assert note.modified == modified


class TestNoteToDict:
    """Tests for Note.to_dict serialization."""

    def test_serializes_all_fields(self):
        """Serializes all fields correctly."""
        note = Note(
            subject="Shopping list",
            body="Buy groceries",
            color="green",
            created=datetime(2024, 3, 15, 10, 0),
            modified=datetime(2024, 3, 15, 12, 0),
            categories=["Personal"],
            entry_id="test-id",
        )

        result = note.to_dict()

        assert result["subject"] == "Shopping list"
        assert result["body"] == "Buy groceries"
        assert result["color"] == "green"
        assert result["created"] == "2024-03-15T10:00:00"
        assert result["modified"] == "2024-03-15T12:00:00"
        assert result["categories"] == ["Personal"]
        assert result["entry_id"] == "test-id"

    def test_handles_none_timestamps(self):
        """Handles None timestamps."""
        note = Note(
            subject="Note",
            created=None,
            modified=None,
        )

        result = note.to_dict()

        assert result["created"] is None
        assert result["modified"] is None


class TestNotesServiceListNotes:
    """Tests for NotesService.list_notes."""

    @pytest.fixture
    def mock_namespace(self):
        """Create mock namespace with notes folder."""
        ns = MagicMock()

        notes_folder = MagicMock()
        items = MagicMock()
        notes_folder.Items = items

        ns.GetDefaultFolder.return_value = notes_folder

        return ns

    def test_list_returns_notes(self, mock_namespace):
        """Returns list of Note objects."""
        note_item = make_mock_note_item()
        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 1
        items.Item.return_value = note_item

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            notes = svc.list_notes()

        assert len(notes) == 1
        assert notes[0].subject == "Shopping list"

    def test_filters_by_color(self, mock_namespace):
        """Filters notes by color."""
        note1 = make_mock_note_item(color=3)  # yellow
        note2 = make_mock_note_item(color=0)  # blue

        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 2
        items.Item.side_effect = [note1, note2]

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            notes = svc.list_notes(color="blue")

        assert len(notes) == 1
        assert notes[0].color == "blue"

    def test_filters_by_keyword(self, mock_namespace):
        """Filters notes by keyword in subject/body."""
        note1 = make_mock_note_item(subject="Shopping", body="Buy milk")
        note2 = make_mock_note_item(subject="Ideas", body="Project plan")

        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 2
        items.Item.side_effect = [note1, note2]

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            notes = svc.list_notes(keyword="milk")

        assert len(notes) == 1
        assert notes[0].subject == "Shopping"

    def test_respects_limit(self, mock_namespace):
        """Respects limit parameter."""
        notes_list = [make_mock_note_item(subject=f"Note {i}") for i in range(5)]
        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 5
        items.Item.side_effect = notes_list

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            notes = svc.list_notes(limit=2)

        assert len(notes) == 2

    def test_skips_non_note_items(self, mock_namespace):
        """Skips items that aren't notes (Class != 44)."""
        note_item = make_mock_note_item()
        non_note = make_mock_note_item(item_class=43)  # olMail

        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 2
        items.Item.side_effect = [note_item, non_note]

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            notes = svc.list_notes()

        assert len(notes) == 1


class TestNotesServiceCreateNote:
    """Tests for NotesService.create_note."""

    @pytest.fixture
    def mock_namespace(self):
        """Create mock namespace with notes folder."""
        ns = MagicMock()
        notes_folder = MagicMock()
        ns.GetDefaultFolder.return_value = notes_folder

        # Mock Application.CreateItem
        mock_note = MagicMock()
        mock_note.EntryID = "new-note-id"
        ns.Application.CreateItem.return_value = mock_note

        return ns

    def test_create_basic_note(self, mock_namespace):
        """Creates note with required fields."""
        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            success, entry_id = svc.create_note(
                body="This is a note",
            )

        assert success is True
        assert entry_id == "new-note-id"

    def test_create_note_with_color(self, mock_namespace):
        """Creates note with specified color."""
        mock_note = mock_namespace.Application.CreateItem.return_value

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            svc.create_note(
                body="Note",
                color="green",
            )

        assert mock_note.Color == COLOR_TO_OUTLOOK["green"]

    def test_create_note_with_categories(self, mock_namespace):
        """Creates note with categories."""
        mock_note = mock_namespace.Application.CreateItem.return_value

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(mock_namespace)
            svc.create_note(
                body="Note",
                categories=["Personal", "Ideas"],
            )

        assert "Personal" in mock_note.Categories
        assert "Ideas" in mock_note.Categories


class TestNotesServiceDeleteNote:
    """Tests for NotesService.delete_note."""

    def test_delete_success(self):
        """Deletes note successfully."""
        ns = MagicMock()
        mock_note = MagicMock()
        mock_note.Class = 44  # olNote
        ns.GetItemFromID.return_value = mock_note
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(ns)
            success, message = svc.delete_note("test-id")

        assert success is True
        mock_note.Delete.assert_called_once()

    def test_delete_non_note_fails(self):
        """Returns error when item is not a note."""
        ns = MagicMock()
        mock_item = MagicMock()
        mock_item.Class = 43  # olMail, not a note
        ns.GetItemFromID.return_value = mock_item
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(ns)
            success, message = svc.delete_note("test-id")

        assert success is False
        assert "not a note" in message

    def test_delete_not_found(self):
        """Returns error when note not found."""
        ns = MagicMock()
        ns.GetItemFromID.side_effect = Exception("Item not found")
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(ns)
            success, message = svc.delete_note("nonexistent-id")

        assert success is False
        assert "Failed" in message


class TestNotesServiceGetNote:
    """Tests for NotesService.get_note."""

    def test_get_note_success(self):
        """Gets note by entry ID."""
        ns = MagicMock()
        mock_note = make_mock_note_item()
        ns.GetItemFromID.return_value = mock_note
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(ns)
            note = svc.get_note("test-id")

        assert note is not None
        assert note.subject == "Shopping list"

    def test_get_non_note_returns_none(self):
        """Returns None when item is not a note."""
        ns = MagicMock()
        mock_item = MagicMock()
        mock_item.Class = 43  # olMail, not a note
        ns.GetItemFromID.return_value = mock_item
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(ns)
            note = svc.get_note("test-id")

        assert note is None

    def test_get_not_found_returns_none(self):
        """Returns None when note not found."""
        ns = MagicMock()
        ns.GetItemFromID.side_effect = Exception("Item not found")
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = NotesService(ns)
            note = svc.get_note("nonexistent-id")

        assert note is None
