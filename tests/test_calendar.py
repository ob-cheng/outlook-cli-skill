"""Tests for CalendarService and CalendarEvent.

Uses mock COM objects to avoid win32com dependency.
"""

import pytest
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch

from outlook_cli.services.calendar import CalendarEvent, CalendarService


def make_mock_appointment(
    subject="Team Meeting",
    start=None,
    end=None,
    location="Conference Room A",
    body="Meeting agenda here",
    organizer="alice@example.com",
    recipients=None,
    is_all_day=False,
    is_recurring=False,
    entry_id="appt-id-123",
    categories="",
    importance="1",
    reminder_set=True,
    reminder_minutes=15,
):
    """Create a mock Outlook AppointmentItem."""
    appt = MagicMock()
    appt.Subject = subject
    appt.Start = start or datetime(2024, 3, 15, 10, 0)
    appt.End = end or datetime(2024, 3, 15, 11, 0)
    appt.Location = location
    appt.Body = body
    appt.Organizer = organizer
    appt.AllDayEvent = is_all_day
    appt.IsRecurring = is_recurring
    appt.EntryID = entry_id
    appt.Categories = categories
    appt.Importance = importance
    appt.ReminderSet = reminder_set
    appt.ReminderMinutesBeforeStart = reminder_minutes

    # Recipients mock (attendees)
    if recipients is None:
        recipients = [("bob@example.com", 1), ("charlie@example.com", 2)]  # 1=Required, 2=Optional
    appt.Recipients = MagicMock()
    appt.Recipients.Count = len(recipients)

    def get_recipient(i):
        email, recip_type = recipients[i - 1]
        recip = MagicMock()
        recip.Address = email
        recip.Type = recip_type
        return recip

    appt.Recipients.Item = get_recipient

    return appt


class TestCalendarEventDataclass:
    """Tests for CalendarEvent dataclass basics."""

    def test_event_creation(self):
        """CalendarEvent can be created with all fields."""
        event = CalendarEvent(
            subject="Meeting",
            start=datetime(2024, 3, 15, 10, 0),
            end=datetime(2024, 3, 15, 11, 0),
            location="Room A",
        )

        assert event.subject == "Meeting"
        assert event.location == "Room A"

    def test_event_defaults(self):
        """CalendarEvent has sensible defaults."""
        event = CalendarEvent(
            subject="Meeting",
            start=datetime(2024, 3, 15, 10, 0),
            end=datetime(2024, 3, 15, 11, 0),
        )

        assert event.location is None
        assert event.body is None
        assert event.is_all_day is False
        assert event.is_recurring is False
        assert event.required_attendees == []
        assert event.optional_attendees == []
        assert event.categories == []
        assert event.importance == "normal"


class TestCalendarEventFromAppointment:
    """Tests for CalendarEvent.from_appointment factory."""

    def test_basic_conversion(self):
        """Converts basic appointment correctly."""
        appt = make_mock_appointment()

        event = CalendarEvent.from_appointment(appt)

        assert event.subject == "Team Meeting"
        assert event.location == "Conference Room A"
        assert event.organizer == "alice@example.com"
        assert event.entry_id == "appt-id-123"

    def test_handles_missing_subject(self):
        """Missing subject becomes '(No Subject)'."""
        appt = make_mock_appointment(subject=None)

        event = CalendarEvent.from_appointment(appt)

        assert event.subject == "(No Subject)"

    def test_extracts_required_attendees(self):
        """Extracts required attendees (Type 1)."""
        appt = make_mock_appointment(
            recipients=[("bob@example.com", 1), ("charlie@example.com", 1)]
        )

        event = CalendarEvent.from_appointment(appt)

        assert "bob@example.com" in event.required_attendees
        assert "charlie@example.com" in event.required_attendees
        assert event.optional_attendees == []

    def test_extracts_optional_attendees(self):
        """Extracts optional attendees (Type 2)."""
        appt = make_mock_appointment(
            recipients=[("bob@example.com", 2)]
        )

        event = CalendarEvent.from_appointment(appt)

        assert event.required_attendees == []
        assert "bob@example.com" in event.optional_attendees

    def test_all_day_event(self):
        """Handles all-day events."""
        appt = make_mock_appointment(is_all_day=True)

        event = CalendarEvent.from_appointment(appt)

        assert event.is_all_day is True

    def test_recurring_event(self):
        """Handles recurring events."""
        appt = make_mock_appointment(is_recurring=True)
        # Mock the recurrence pattern
        rec_pattern = MagicMock()
        rec_pattern.RecurrenceType = 1  # Weekly
        rec_pattern.Interval = 1
        appt.GetRecurrencePattern.return_value = rec_pattern

        event = CalendarEvent.from_appointment(appt)

        assert event.is_recurring is True
        assert event.recurrence_pattern is not None

    def test_importance_high(self):
        """Importance 2 becomes 'high'."""
        appt = make_mock_appointment(importance="2")

        event = CalendarEvent.from_appointment(appt)

        assert event.importance == "high"

    def test_importance_low(self):
        """Importance 0 becomes 'low'."""
        appt = make_mock_appointment(importance="0")

        event = CalendarEvent.from_appointment(appt)

        assert event.importance == "low"

    def test_extracts_categories(self):
        """Extracts categories from comma-separated string."""
        appt = make_mock_appointment(categories="Work, Important")

        event = CalendarEvent.from_appointment(appt)

        assert "Work" in event.categories
        assert "Important" in event.categories

    def test_reminder_minutes(self):
        """Extracts reminder minutes when set."""
        appt = make_mock_appointment(reminder_set=True, reminder_minutes=30)

        event = CalendarEvent.from_appointment(appt)

        assert event.reminder_minutes == 30

    def test_no_reminder(self):
        """Returns None when reminder not set."""
        appt = make_mock_appointment(reminder_set=False)

        event = CalendarEvent.from_appointment(appt)

        assert event.reminder_minutes is None


class TestCalendarEventToDict:
    """Tests for CalendarEvent.to_dict serialization."""

    def test_serializes_all_fields(self):
        """Serializes all fields correctly."""
        event = CalendarEvent(
            subject="Meeting",
            start=datetime(2024, 3, 15, 10, 0),
            end=datetime(2024, 3, 15, 11, 0),
            location="Room A",
            organizer="alice@example.com",
            required_attendees=["bob@example.com"],
            optional_attendees=["charlie@example.com"],
            is_all_day=False,
            is_recurring=False,
            entry_id="test-id",
        )

        result = event.to_dict()

        assert result["subject"] == "Meeting"
        assert result["start"] == "2024-03-15T10:00:00"
        assert result["end"] == "2024-03-15T11:00:00"
        assert result["location"] == "Room A"
        assert result["entry_id"] == "test-id"

    def test_handles_none_dates(self):
        """Handles None start/end dates."""
        event = CalendarEvent(
            subject="Meeting",
            start=None,
            end=None,
        )

        result = event.to_dict()

        assert result["start"] is None
        assert result["end"] is None


class TestCalendarEventFormatRecurrence:
    """Tests for recurrence pattern formatting."""

    def test_daily_recurrence(self):
        """Formats daily recurrence."""
        rec_pattern = MagicMock()
        rec_pattern.RecurrenceType = 0  # Daily
        rec_pattern.Interval = 1

        result = CalendarEvent._format_recurrence(rec_pattern)

        assert result == "Daily"

    def test_weekly_recurrence(self):
        """Formats weekly recurrence."""
        rec_pattern = MagicMock()
        rec_pattern.RecurrenceType = 1  # Weekly
        rec_pattern.Interval = 1

        result = CalendarEvent._format_recurrence(rec_pattern)

        assert result == "Weekly"

    def test_interval_greater_than_one(self):
        """Formats interval > 1."""
        rec_pattern = MagicMock()
        rec_pattern.RecurrenceType = 0  # Daily
        rec_pattern.Interval = 3

        result = CalendarEvent._format_recurrence(rec_pattern)

        assert "every 3" in result


class TestCalendarServiceListEvents:
    """Tests for CalendarService.list_events."""

    @pytest.fixture
    def mock_namespace(self):
        """Create mock namespace with calendar folder."""
        ns = MagicMock()

        calendar_folder = MagicMock()
        items = MagicMock()
        calendar_folder.Items = items

        ns.GetDefaultFolder.return_value = calendar_folder

        return ns

    def test_list_returns_events(self, mock_namespace):
        """Returns list of CalendarEvent objects."""
        appt = make_mock_appointment()
        items = mock_namespace.GetDefaultFolder.return_value.Items
        restricted = MagicMock()
        restricted.Count = 1
        restricted.Item.return_value = appt
        items.Restrict.return_value = restricted

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = CalendarService(mock_namespace)
            events = svc.list_events()

        assert len(events) == 1
        assert events[0].subject == "Team Meeting"

    def test_applies_subject_filter(self, mock_namespace):
        """Filters events by subject."""
        appt1 = make_mock_appointment(subject="Team Meeting")
        appt2 = make_mock_appointment(subject="Lunch")

        items = mock_namespace.GetDefaultFolder.return_value.Items
        restricted = MagicMock()
        restricted.Count = 2
        restricted.Item.side_effect = [appt1, appt2]
        items.Restrict.return_value = restricted

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = CalendarService(mock_namespace)
            events = svc.list_events(subject_filter="Team")

        assert len(events) == 1
        assert events[0].subject == "Team Meeting"

    def test_applies_all_day_filter(self, mock_namespace):
        """Filters for all-day events only."""
        appt1 = make_mock_appointment(is_all_day=True)
        appt2 = make_mock_appointment(is_all_day=False)

        items = mock_namespace.GetDefaultFolder.return_value.Items
        restricted = MagicMock()
        restricted.Count = 2
        restricted.Item.side_effect = [appt1, appt2]
        items.Restrict.return_value = restricted

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = CalendarService(mock_namespace)
            events = svc.list_events(all_day_only=True)

        assert len(events) == 1
        assert events[0].is_all_day is True


class TestCalendarServiceCreateEvent:
    """Tests for CalendarService.create_event."""

    @pytest.fixture
    def mock_namespace(self):
        """Create mock namespace with calendar folder."""
        ns = MagicMock()
        calendar_folder = MagicMock()
        ns.GetDefaultFolder.return_value = calendar_folder

        # Mock Application.CreateItem
        mock_appt = MagicMock()
        mock_appt.EntryID = "new-event-id"
        mock_appt.Recipients = MagicMock()
        ns.Application.CreateItem.return_value = mock_appt

        return ns

    def test_create_basic_event(self, mock_namespace):
        """Creates event with required fields."""
        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = CalendarService(mock_namespace)
            success, entry_id = svc.create_event(
                subject="New Meeting",
                start=datetime(2024, 3, 20, 14, 0),
                end=datetime(2024, 3, 20, 15, 0),
            )

        assert success is True
        assert entry_id == "new-event-id"

    def test_create_event_with_attendees(self, mock_namespace):
        """Creates event with required and optional attendees."""
        mock_appt = mock_namespace.Application.CreateItem.return_value

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = CalendarService(mock_namespace)
            svc.create_event(
                subject="Meeting",
                start=datetime(2024, 3, 20, 14, 0),
                end=datetime(2024, 3, 20, 15, 0),
                required_attendees=["bob@example.com"],
                optional_attendees=["charlie@example.com"],
            )

        # Verify recipients were added
        assert mock_appt.Recipients.Add.call_count == 2


class TestCalendarServiceDeleteEvent:
    """Tests for CalendarService.delete_event."""

    def test_delete_success(self):
        """Deletes event successfully."""
        ns = MagicMock()
        mock_appt = MagicMock()
        ns.GetItemFromID.return_value = mock_appt
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = CalendarService(ns)
            success, message = svc.delete_event("test-id")

        assert success is True
        mock_appt.Delete.assert_called_once()

    def test_delete_not_found(self):
        """Returns error when event not found."""
        ns = MagicMock()
        ns.GetItemFromID.side_effect = Exception("Item not found")
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = CalendarService(ns)
            success, message = svc.delete_event("nonexistent-id")

        assert success is False
        assert "Failed" in message
