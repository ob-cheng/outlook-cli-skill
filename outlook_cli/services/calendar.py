"""Calendar service for managing Outlook events."""

from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Optional


@dataclass
class CalendarEvent:
    """Represents a calendar event."""
    subject: str
    start: datetime
    end: datetime
    location: Optional[str] = None
    body: Optional[str] = None
    organizer: Optional[str] = None
    required_attendees: list[str] = None
    optional_attendees: list[str] = None
    is_all_day: bool = False
    is_recurring: bool = False
    recurrence_pattern: Optional[str] = None
    entry_id: Optional[str] = None
    categories: list[str] = None
    importance: str = "normal"
    reminder_minutes: Optional[int] = None

    def __post_init__(self):
        if self.required_attendees is None:
            self.required_attendees = []
        if self.optional_attendees is None:
            self.optional_attendees = []
        if self.categories is None:
            self.categories = []

    @classmethod
    def from_appointment(cls, appt) -> 'CalendarEvent':
        """Create CalendarEvent from Outlook AppointmentItem."""
        try:
            subject = appt.Subject or "(No Subject)"
        except Exception:
            subject = "(No Subject)"

        try:
            start = appt.Start
        except Exception:
            start = None

        try:
            end = appt.End
        except Exception:
            end = None

        try:
            location = appt.Location or None
        except Exception:
            location = None

        try:
            body = appt.Body or None
        except Exception:
            body = None

        try:
            organizer = appt.Organizer or None
        except Exception:
            organizer = None

        # Attendees
        required = []
        optional = []
        try:
            for i in range(1, appt.Recipients.Count + 1):
                recip = appt.Recipients.Item(i)
                email = recip.Address
                if recip.Type == 1:  # Required
                    required.append(email)
                elif recip.Type == 2:  # Optional
                    optional.append(email)
        except Exception:
            pass

        try:
            is_all_day = appt.AllDayEvent
        except Exception:
            is_all_day = False

        try:
            is_recurring = appt.IsRecurring
        except Exception:
            is_recurring = False

        recurrence_pattern = None
        if is_recurring:
            try:
                rec_pattern = appt.GetRecurrencePattern()
                recurrence_pattern = cls._format_recurrence(rec_pattern)
            except Exception:
                pass

        try:
            entry_id = appt.EntryID
        except Exception:
            entry_id = None

        try:
            categories = [c.strip() for c in appt.Categories.split(',')] if appt.Categories else []
        except Exception:
            categories = []

        try:
            importance = str(appt.Importance)
            if importance == "2":
                importance = "high"
            elif importance == "0":
                importance = "low"
            else:
                importance = "normal"
        except Exception:
            importance = "normal"

        try:
            reminder_minutes = appt.ReminderMinutesBeforeStart if appt.ReminderSet else None
        except Exception:
            reminder_minutes = None

        return cls(
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
            organizer=organizer,
            required_attendees=required,
            optional_attendees=optional,
            is_all_day=is_all_day,
            is_recurring=is_recurring,
            recurrence_pattern=recurrence_pattern,
            entry_id=entry_id,
            categories=categories,
            importance=importance,
            reminder_minutes=reminder_minutes,
        )

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            'subject': self.subject,
            'start': self.start.isoformat() if self.start else None,
            'end': self.end.isoformat() if self.end else None,
            'location': self.location,
            'body': self.body,
            'organizer': self.organizer,
            'required_attendees': self.required_attendees,
            'optional_attendees': self.optional_attendees,
            'is_all_day': self.is_all_day,
            'is_recurring': self.is_recurring,
            'recurrence_pattern': self.recurrence_pattern,
            'entry_id': self.entry_id,
            'categories': self.categories,
            'importance': self.importance,
            'reminder_minutes': self.reminder_minutes,
        }

    @staticmethod
    def _format_recurrence(rec_pattern) -> str:
        """Format recurrence pattern to human-readable string."""
        try:
            rec_type = rec_pattern.RecurrenceType
            interval = rec_pattern.Interval

            type_names = {
                0: "Daily",
                1: "Weekly",
                2: "Monthly",
                3: "Monthly (relative)",
                5: "Yearly",
                6: "Yearly (relative)",
            }

            pattern_type = type_names.get(rec_type, "Unknown")

            if interval > 1:
                return f"{pattern_type} (every {interval})"
            return pattern_type
        except Exception:
            return "Custom recurrence"


class CalendarService:
    """Service for managing calendar events."""

    def __init__(self, namespace):
        """Initialize with Outlook MAPI namespace."""
        self.namespace = namespace
        self.calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

    def list_events(
        self,
        start_date: datetime | None = None,
        end_date: datetime | None = None,
        subject_filter: str | None = None,
        location_filter: str | None = None,
        organizer_filter: str | None = None,
        all_day_only: bool = False,
        recurring_only: bool = False,
    ) -> list[CalendarEvent]:
        """List calendar events with filters.

        Args:
            start_date: Start date (default: today)
            end_date: End date (default: 7 days from start)
            subject_filter: Filter by subject (case-insensitive substring)
            location_filter: Filter by location (case-insensitive substring)
            organizer_filter: Filter by organizer email
            all_day_only: Only return all-day events
            recurring_only: Only return recurring events

        Returns:
            List of CalendarEvent objects
        """
        if not start_date:
            start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        if not end_date:
            end_date = start_date + timedelta(days=7)

        events = []

        try:
            items = self.calendar.Items
            items.Sort("[Start]")
            # Note: IncludeRecurrences can be slow on large calendars
            # items.IncludeRecurrences = True

            # Build date filter
            filter_str = (
                f"[Start] >= '{self._format_date(start_date)}' AND "
                f"[Start] <= '{self._format_date(end_date)}'"
            )

            filtered = items.Restrict(filter_str)

            for i in range(1, filtered.Count + 1):
                try:
                    appt = filtered.Item(i)

                    # Apply additional filters
                    if subject_filter and subject_filter.lower() not in (appt.Subject or "").lower():
                        continue
                    if location_filter and location_filter.lower() not in (appt.Location or "").lower():
                        continue
                    if organizer_filter and organizer_filter.lower() not in (appt.Organizer or "").lower():
                        continue
                    if all_day_only and not appt.AllDayEvent:
                        continue
                    if recurring_only and not appt.IsRecurring:
                        continue

                    event = CalendarEvent.from_appointment(appt)
                    events.append(event)

                except Exception:
                    continue

        except Exception:
            pass

        return events

    def get_event(self, entry_id: str) -> CalendarEvent | None:
        """Get a single event by EntryID.

        Args:
            entry_id: The event's EntryID

        Returns:
            CalendarEvent or None if not found
        """
        try:
            appt = self.namespace.GetItemFromID(entry_id)
            return CalendarEvent.from_appointment(appt)
        except Exception:
            return None

    def create_event(
        self,
        subject: str,
        start: datetime,
        end: datetime,
        location: str | None = None,
        body: str | None = None,
        required_attendees: list[str] | None = None,
        optional_attendees: list[str] | None = None,
        reminder_minutes: int | None = 15,
        categories: list[str] | None = None,
    ) -> tuple[bool, str]:
        """Create a new calendar event.

        Args:
            subject: Event subject
            start: Start date/time
            end: End date/time
            location: Event location
            body: Event description
            required_attendees: List of required attendee emails
            optional_attendees: List of optional attendee emails
            reminder_minutes: Reminder before event (None = no reminder)
            categories: List of category names

        Returns:
            tuple: (success: bool, entry_id_or_error: str)
        """
        try:
            outlook = self.namespace.Application
            appt = outlook.CreateItem(1)  # 1 = olAppointmentItem

            appt.Subject = subject
            appt.Start = start
            appt.End = end

            if location:
                appt.Location = location

            if body:
                appt.Body = body

            # Attendees
            if required_attendees:
                for email in required_attendees:
                    recip = appt.Recipients.Add(email)
                    recip.Type = 1  # Required

            if optional_attendees:
                for email in optional_attendees:
                    recip = appt.Recipients.Add(email)
                    recip.Type = 2  # Optional

            # Reminder
            if reminder_minutes is not None:
                appt.ReminderSet = True
                appt.ReminderMinutesBeforeStart = reminder_minutes
            else:
                appt.ReminderSet = False

            # Categories
            if categories:
                appt.Categories = ", ".join(categories)

            appt.Save()
            return True, appt.EntryID

        except Exception as e:
            return False, f"Failed to create event: {e}"

    def update_event(
        self,
        entry_id: str,
        subject: str | None = None,
        start: datetime | None = None,
        end: datetime | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> tuple[bool, str]:
        """Update an existing event.

        Args:
            entry_id: Event EntryID
            subject: New subject (None = no change)
            start: New start time (None = no change)
            end: New end time (None = no change)
            location: New location (None = no change)
            body: New body (None = no change)

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            appt = self.namespace.GetItemFromID(entry_id)

            if subject is not None:
                appt.Subject = subject
            if start is not None:
                appt.Start = start
            if end is not None:
                appt.End = end
            if location is not None:
                appt.Location = location
            if body is not None:
                appt.Body = body

            appt.Save()
            return True, "Event updated successfully"

        except Exception as e:
            return False, f"Failed to update event: {e}"

    def delete_event(self, entry_id: str) -> tuple[bool, str]:
        """Delete a calendar event.

        Args:
            entry_id: Event EntryID

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            appt = self.namespace.GetItemFromID(entry_id)
            appt.Delete()
            return True, "Event deleted successfully"
        except Exception as e:
            return False, f"Failed to delete event: {e}"

    def _format_date(self, dt: datetime) -> str:
        """Format datetime for Outlook filter."""
        return dt.strftime("%m/%d/%Y %H:%M %p")
