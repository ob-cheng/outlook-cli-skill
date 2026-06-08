"""Notes service for managing Outlook notes."""

from datetime import datetime
from dataclasses import dataclass
from typing import Optional


# Note color constants (Outlook OlNoteColor)
NOTE_COLORS = {
    0: "blue",
    1: "green",
    2: "pink",
    3: "yellow",
    4: "white",
}

COLOR_TO_OUTLOOK = {v: k for k, v in NOTE_COLORS.items()}


@dataclass
class Note:
    """Represents an Outlook note."""
    subject: str
    body: Optional[str] = None
    color: str = "yellow"
    created: Optional[datetime] = None
    modified: Optional[datetime] = None
    categories: list[str] = None
    entry_id: Optional[str] = None

    def __post_init__(self):
        if self.categories is None:
            self.categories = []

    @classmethod
    def from_note_item(cls, note_item) -> 'Note':
        """Create Note from Outlook NoteItem."""
        try:
            subject = note_item.Subject or "(No Subject)"
        except Exception:
            subject = "(No Subject)"

        try:
            body = note_item.Body or None
        except Exception:
            body = None

        try:
            color = NOTE_COLORS.get(note_item.Color, "yellow")
        except Exception:
            color = "yellow"

        try:
            created = note_item.CreationTime
        except Exception:
            created = None

        try:
            modified = note_item.LastModificationTime
        except Exception:
            modified = None

        try:
            categories = [c.strip() for c in note_item.Categories.split(',')] if note_item.Categories else []
        except Exception:
            categories = []

        try:
            entry_id = note_item.EntryID
        except Exception:
            entry_id = None

        return cls(
            subject=subject,
            body=body,
            color=color,
            created=created,
            modified=modified,
            categories=categories,
            entry_id=entry_id,
        )

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            'entry_id': self.entry_id,
            'subject': self.subject,
            'body': self.body,
            'color': self.color,
            'created': self.created.isoformat() if self.created else None,
            'modified': self.modified.isoformat() if self.modified else None,
            'categories': self.categories,
        }


class NotesService:
    """Service for managing Outlook notes."""

    def __init__(self, namespace):
        """Initialize with Outlook MAPI namespace."""
        self.namespace = namespace
        from ..core.folders import find_folder_across_stores
        self.notes_folder = find_folder_across_stores(namespace, "Notes")
        if self.notes_folder is None:
            self.notes_folder = namespace.GetDefaultFolder(12)  # fallback: 12 = olFolderNotes

    def list_notes(
        self,
        color: str | None = None,
        category: str | None = None,
        keyword: str | None = None,
        limit: int | None = None,
    ) -> list[Note]:
        """List notes with filters.

        Args:
            color: Filter by color (blue, green, pink, yellow, white)
            category: Filter by category name
            keyword: Search keyword in subject/body
            limit: Maximum number of notes to return

        Returns:
            List of Note objects
        """
        notes = []

        try:
            items = self.notes_folder.Items
            items.Sort("[LastModificationTime]", True)  # Most recent first

            count = 0
            for i in range(1, items.Count + 1):
                if limit and count >= limit:
                    break

                try:
                    note_item = items.Item(i)

                    # Check if it's a note (Class 44 = olNote)
                    if note_item.Class != 44:
                        continue

                    # Color filter
                    if color:
                        note_color = NOTE_COLORS.get(note_item.Color, "yellow")
                        if note_color != color.lower():
                            continue

                    # Category filter
                    if category:
                        note_categories = note_item.Categories or ""
                        if category.lower() not in note_categories.lower():
                            continue

                    # Keyword filter
                    if keyword:
                        keyword_lower = keyword.lower()
                        subject = (note_item.Subject or "").lower()
                        body = (note_item.Body or "").lower()
                        if keyword_lower not in subject and keyword_lower not in body:
                            continue

                    note = Note.from_note_item(note_item)
                    notes.append(note)
                    count += 1

                except Exception:
                    continue

        except Exception:
            pass

        return notes

    def get_note(self, entry_id: str) -> Note | None:
        """Get a single note by EntryID.

        Args:
            entry_id: The note's EntryID

        Returns:
            Note or None if not found
        """
        try:
            note_item = self.namespace.GetItemFromID(entry_id)
            if note_item.Class == 44:  # olNote
                return Note.from_note_item(note_item)
        except Exception:
            pass
        return None

    def create_note(
        self,
        body: str,
        color: str = "yellow",
        categories: list[str] | None = None,
    ) -> tuple[bool, str]:
        """Create a new note.

        Args:
            body: Note content
            color: Note color (blue, green, pink, yellow, white)
            categories: List of category names

        Returns:
            tuple: (success: bool, entry_id_or_error: str)
        """
        try:
            outlook = self.namespace.Application
            note_item = outlook.CreateItem(5)  # 5 = olNoteItem

            note_item.Body = body

            if color and color.lower() in COLOR_TO_OUTLOOK:
                note_item.Color = COLOR_TO_OUTLOOK[color.lower()]

            if categories:
                note_item.Categories = ", ".join(categories)

            note_item.Save()
            return True, note_item.EntryID

        except Exception as e:
            return False, f"Failed to create note: {e}"

    def update_note(
        self,
        entry_id: str,
        body: str | None = None,
        color: str | None = None,
        categories: list[str] | None = None,
    ) -> tuple[bool, str]:
        """Update an existing note.

        Args:
            entry_id: Note EntryID
            body: New body (None = no change)
            color: New color (None = no change)
            categories: New categories (None = no change)

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            note_item = self.namespace.GetItemFromID(entry_id)

            if note_item.Class != 44:
                return False, "Item is not a note"

            if body is not None:
                note_item.Body = body

            if color is not None and color.lower() in COLOR_TO_OUTLOOK:
                note_item.Color = COLOR_TO_OUTLOOK[color.lower()]

            if categories is not None:
                note_item.Categories = ", ".join(categories)

            note_item.Save()
            return True, "Note updated successfully"

        except Exception as e:
            return False, f"Failed to update note: {e}"

    def delete_note(self, entry_id: str) -> tuple[bool, str]:
        """Delete a note.

        Args:
            entry_id: Note EntryID

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            note_item = self.namespace.GetItemFromID(entry_id)
            if note_item.Class != 44:
                return False, "Item is not a note"
            note_item.Delete()
            return True, "Note deleted successfully"
        except Exception as e:
            return False, f"Failed to delete note: {e}"
