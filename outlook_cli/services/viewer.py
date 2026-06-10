"""Terminal viewer service for emails."""

import sys

from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich import box

from ..core.models import Email
from ..core.folders import list_all_folders
from .calendar import CalendarEvent
from .tasks import Task
from .notes import Note


# Force UTF-8 encoding on Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

console = Console(force_terminal=True)


class ViewerService:
    """Service for displaying emails in the terminal."""

    def __init__(self):
        self.console = console

    def print_folders(self, namespace) -> None:
        """Print all available folders as a tree."""
        folders = list_all_folders(namespace)

        self.console.print("\n[bold]Available Folders[/bold]")
        self.console.print("-" * 60)

        for f in folders:
            indent = "  " * f['level']
            name = f['name'].encode('ascii', 'replace').decode()
            count = f.get('count', 0)

            if f.get('is_store'):
                self.console.print(f"{indent}[bold cyan]{name}[/bold cyan] [dim](Account)[/dim]")
            elif count > 0:
                self.console.print(f"{indent}{name} [dim]({count} items)[/dim]")
            else:
                self.console.print(f"{indent}[dim]{name}[/dim]")

        self.console.print("-" * 60)

    def print_email_table(self, emails: list[Email], title: str = "Messages") -> None:
        """Print emails as a formatted table."""
        if not emails:
            self.console.print("[yellow]No emails found.[/yellow]")
            return

        table = Table(
            title=title,
            box=box.ROUNDED,
            show_lines=False,
            header_style="bold cyan",
        )

        table.add_column("", width=1)  # Unread indicator
        table.add_column("!", width=1)  # Importance
        table.add_column("@", width=1)  # Attachments
        table.add_column("From", style="cyan", max_width=25, overflow="ellipsis")
        table.add_column("Subject", style="white", max_width=40, overflow="ellipsis")
        table.add_column("Date", style="green", width=16)
        table.add_column("ID", style="dim", max_width=20, overflow="ellipsis")

        for email in emails:
            # Status indicators (use ASCII-safe characters for Windows console)
            unread = "[bold blue]*[/]" if not email.is_read else ""
            importance = "[bold red]![/]" if email.importance == "high" else ""
            attachments = "[yellow]@[/]" if email.has_attachments else ""

            # Sender
            sender = email.sender_clean or email.sender or "Unknown"
            if len(sender) > 25:
                sender = sender[:22] + "..."

            # Subject with direction indicator (ASCII-safe)
            direction = "[green]>[/]" if email.is_sent else "[cyan]<[/]"
            subject = f"{direction} {email.subject or '(no subject)'}"

            # Date
            date_str = email.date.strftime("%Y-%m-%d %H:%M") if email.date else ""

            # Short ID for display
            short_id = email.message_id[:20] + "..." if email.message_id and len(email.message_id) > 20 else (email.message_id or "")

            table.add_row(
                unread,
                importance,
                attachments,
                sender,
                subject,
                date_str,
                short_id,
            )

        self.console.print(table)
        self.console.print(f"\n[dim]Total: {len(emails)} email(s)[/dim]")

    def print_email_detail(self, email: Email, text_only: bool = False, max_body_lines: int | None = None) -> None:
        """Print a single email in detail.

        Args:
            email: The email to display.
            text_only: Only show plain text body.
            max_body_lines: Truncate body to first N lines.
        """
        # Header info
        header_lines = []
        header_lines.append(f"[bold]From:[/bold] {email.sender_clean or email.sender}")
        header_lines.append(f"[bold]To:[/bold] {email.to}")
        if email.cc:
            header_lines.append(f"[bold]CC:[/bold] {email.cc}")
        if email.date:
            header_lines.append(f"[bold]Date:[/bold] {email.date.strftime('%Y-%m-%d %H:%M:%S')}")

        # Status line
        status_parts = []
        if not email.is_read:
            status_parts.append("[blue]UNREAD[/blue]")
        if email.importance == "high":
            status_parts.append("[red]HIGH IMPORTANCE[/red]")
        if email.has_attachments:
            status_parts.append("[yellow]HAS ATTACHMENTS[/yellow]")
        if email.is_sent:
            status_parts.append("[green]SENT[/green]")
        else:
            status_parts.append("[cyan]RECEIVED[/cyan]")

        if status_parts:
            header_lines.append(" | ".join(status_parts))

        header = "\n".join(header_lines)

        # Print header panel
        title = email.subject or "(no subject)"
        self.console.print(Panel(
            header,
            title=f"[bold]{title}[/bold]",
            border_style="blue",
            padding=(1, 2),
        ))

        # Print body
        body = self._clean_body(email, text_only=text_only)
        if max_body_lines and body:
            lines = body.split('\n')
            if len(lines) > max_body_lines:
                body = '\n'.join(lines[:max_body_lines]) + f'\n... (truncated, {len(lines) - max_body_lines} more lines)'
        if body:
            self.console.print()
            self.console.print(body)
        else:
            self.console.print("\n[dim](empty body)[/dim]")

        # Print message ID for reference
        if email.message_id:
            self.console.print(f"\n[dim]Message ID: {email.message_id}[/dim]")

    def _clean_body(self, email: Email, text_only: bool = False) -> str:
        """Extract and clean email body.

        Args:
            email: The email to extract body from.
            text_only: If True, only use text_body, skip HTML parsing.
        """
        import re
        from bs4 import BeautifulSoup

        if text_only and email.text_body:
            return email.text_body.strip()

        if email.html_body and not text_only:
            soup = BeautifulSoup(email.html_body, 'html.parser')

            # Remove scripts, styles
            for tag in soup(['script', 'style', 'meta', 'head']):
                tag.decompose()

            # Get text
            text = soup.get_text(separator='\n')

            # Clean up
            text = re.sub(r'\n{3,}', '\n\n', text)
            text = text.strip()

            return text

        elif email.text_body:
            return email.text_body.strip()

        return ""

    def print_summary(self, emails: list[Email]) -> None:
        """Print a summary of the email search results."""
        if not emails:
            return

        total = len(emails)
        unread = sum(1 for e in emails if not e.is_read)
        sent = sum(1 for e in emails if e.is_sent)
        received = total - sent
        with_attachments = sum(1 for e in emails if e.has_attachments)
        high_importance = sum(1 for e in emails if e.importance == "high")

        self.console.print(Panel(
            f"[bold]Total:[/bold] {total} | "
            f"[cyan]Received:[/cyan] {received} | "
            f"[green]Sent:[/green] {sent} | "
            f"[blue]Unread:[/blue] {unread} | "
            f"[yellow]With Attachments:[/yellow] {with_attachments} | "
            f"[red]High Importance:[/red] {high_importance}",
            title="Summary",
            border_style="dim",
        ))

    def print_events_table(self, events: list[CalendarEvent], title: str = "Calendar Events") -> None:
        """Print calendar events as a formatted table."""
        if not events:
            self.console.print("[yellow]No events found.[/yellow]")
            return

        table = Table(
            title=title,
            box=box.ROUNDED,
            show_lines=False,
            header_style="bold cyan",
        )

        table.add_column("R", width=1)  # Recurring indicator
        table.add_column("A", width=1)  # All-day indicator
        table.add_column("Subject", style="white", max_width=30, overflow="ellipsis")
        table.add_column("Start", style="green", width=16)
        table.add_column("End", style="green", width=16)
        table.add_column("Location", style="cyan", max_width=20, overflow="ellipsis")
        table.add_column("Organizer", style="yellow", max_width=20, overflow="ellipsis")

        for event in events:
            recurring = "[bold magenta]R[/]" if event.is_recurring else ""
            all_day = "[bold blue]A[/]" if event.is_all_day else ""

            subject = event.subject or "(no subject)"
            start_str = event.start.strftime("%Y-%m-%d %H:%M") if event.start else ""
            end_str = event.end.strftime("%Y-%m-%d %H:%M") if event.end else ""
            location = event.location or ""
            organizer = event.organizer or ""

            table.add_row(
                recurring,
                all_day,
                subject,
                start_str,
                end_str,
                location,
                organizer,
            )

        self.console.print(table)
        self.console.print(f"\n[dim]Total: {len(events)} event(s)[/dim]")

    def print_event_detail(self, event: CalendarEvent) -> None:
        """Print a single calendar event in detail."""
        # Header info
        header_lines = []
        header_lines.append(f"[bold]Organizer:[/bold] {event.organizer or 'N/A'}")

        if event.start:
            header_lines.append(f"[bold]Start:[/bold] {event.start.strftime('%Y-%m-%d %H:%M:%S')}")
        if event.end:
            header_lines.append(f"[bold]End:[/bold] {event.end.strftime('%Y-%m-%d %H:%M:%S')}")

        if event.location:
            header_lines.append(f"[bold]Location:[/bold] {event.location}")

        # Status line
        status_parts = []
        if event.is_all_day:
            status_parts.append("[blue]ALL-DAY[/blue]")
        if event.is_recurring:
            status_parts.append(f"[magenta]RECURRING[/magenta] ({event.recurrence_pattern})")
        if event.importance == "high":
            status_parts.append("[red]HIGH IMPORTANCE[/red]")
        if event.reminder_minutes is not None:
            status_parts.append(f"[yellow]REMINDER: {event.reminder_minutes} min before[/yellow]")

        if status_parts:
            header_lines.append(" | ".join(status_parts))

        # Attendees
        if event.required_attendees:
            attendees_str = ", ".join(event.required_attendees[:5])
            if len(event.required_attendees) > 5:
                attendees_str += f" +{len(event.required_attendees) - 5} more"
            header_lines.append(f"[bold]Required:[/bold] {attendees_str}")

        if event.optional_attendees:
            attendees_str = ", ".join(event.optional_attendees[:5])
            if len(event.optional_attendees) > 5:
                attendees_str += f" +{len(event.optional_attendees) - 5} more"
            header_lines.append(f"[bold]Optional:[/bold] {attendees_str}")

        if event.categories:
            header_lines.append(f"[bold]Categories:[/bold] {', '.join(event.categories)}")

        header = "\n".join(header_lines)

        # Print header panel
        title = event.subject or "(no subject)"
        self.console.print(Panel(
            header,
            title=f"[bold]{title}[/bold]",
            border_style="blue",
            padding=(1, 2),
        ))

        # Print body
        if event.body:
            self.console.print()
            self.console.print(event.body)
        else:
            self.console.print("\n[dim](no description)[/dim]")

        # Print event ID
        if event.entry_id:
            self.console.print(f"\n[dim]Event ID: {event.entry_id}[/dim]")

    def print_tasks_table(self, tasks: list[Task], title: str = "Tasks") -> None:
        """Print tasks as a formatted table."""
        if not tasks:
            self.console.print("[yellow]No tasks found.[/yellow]")
            return

        table = Table(
            title=title,
            box=box.ROUNDED,
            show_lines=False,
            header_style="bold cyan",
        )

        table.add_column("!", width=1)  # Priority indicator
        table.add_column("Status", style="white", width=12)
        table.add_column("Subject", style="white", max_width=35, overflow="ellipsis")
        table.add_column("Due Date", style="green", width=12)
        table.add_column("%", style="cyan", width=4)
        table.add_column("ID", style="dim", max_width=20, overflow="ellipsis")

        status_styles = {
            "not_started": "[white]Not Started[/]",
            "in_progress": "[blue]In Progress[/]",
            "completed": "[green]Completed[/]",
            "waiting": "[yellow]Waiting[/]",
            "deferred": "[dim]Deferred[/]",
        }

        for task in tasks:
            priority = "[bold red]![/]" if task.priority == "high" else ("[dim]v[/]" if task.priority == "low" else "")
            status = status_styles.get(task.status, task.status)
            subject = task.subject or "(no subject)"
            due_str = task.due_date.strftime("%Y-%m-%d") if task.due_date else ""
            percent = f"{task.percent_complete}%" if task.percent_complete > 0 else ""
            short_id = task.entry_id[:20] + "..." if task.entry_id and len(task.entry_id) > 20 else (task.entry_id or "")

            table.add_row(
                priority,
                status,
                subject,
                due_str,
                percent,
                short_id,
            )

        self.console.print(table)
        self.console.print(f"\n[dim]Total: {len(tasks)} task(s)[/dim]")

    def print_task_detail(self, task: Task) -> None:
        """Print a single task in detail."""
        # Header info
        header_lines = []

        # Status with color
        status_display = {
            "not_started": "[white]Not Started[/]",
            "in_progress": "[blue]In Progress[/]",
            "completed": "[green]Completed[/]",
            "waiting": "[yellow]Waiting[/]",
            "deferred": "[dim]Deferred[/]",
        }
        header_lines.append(f"[bold]Status:[/bold] {status_display.get(task.status, task.status)}")

        # Priority
        priority_display = {
            "high": "[red]High[/red]",
            "normal": "Normal",
            "low": "[dim]Low[/dim]",
        }
        header_lines.append(f"[bold]Priority:[/bold] {priority_display.get(task.priority, task.priority)}")

        # Progress
        if task.percent_complete > 0:
            header_lines.append(f"[bold]Progress:[/bold] {task.percent_complete}%")

        # Dates
        if task.start_date:
            header_lines.append(f"[bold]Start Date:[/bold] {task.start_date.strftime('%Y-%m-%d')}")
        if task.due_date:
            header_lines.append(f"[bold]Due Date:[/bold] {task.due_date.strftime('%Y-%m-%d')}")
        if task.completed_date:
            header_lines.append(f"[bold]Completed:[/bold] {task.completed_date.strftime('%Y-%m-%d %H:%M')}")

        # Reminder
        if task.reminder_date:
            header_lines.append(f"[bold]Reminder:[/bold] {task.reminder_date.strftime('%Y-%m-%d %H:%M')}")

        # Categories
        if task.categories:
            header_lines.append(f"[bold]Categories:[/bold] {', '.join(task.categories)}")

        header = "\n".join(header_lines)

        # Print header panel
        title = task.subject or "(no subject)"
        self.console.print(Panel(
            header,
            title=f"[bold]{title}[/bold]",
            border_style="blue",
            padding=(1, 2),
        ))

        # Print body
        if task.body:
            self.console.print()
            self.console.print(task.body)
        else:
            self.console.print("\n[dim](no description)[/dim]")

        # Print task ID
        if task.entry_id:
            self.console.print(f"\n[dim]Task ID: {task.entry_id}[/dim]")

    def print_notes_table(self, notes: list[Note], title: str = "Notes") -> None:
        """Print notes as a formatted table."""
        if not notes:
            self.console.print("[yellow]No notes found.[/yellow]")
            return

        table = Table(
            title=title,
            box=box.ROUNDED,
            show_lines=False,
            header_style="bold cyan",
        )

        table.add_column("Color", width=8)
        table.add_column("Subject", style="white", max_width=40, overflow="ellipsis")
        table.add_column("Modified", style="green", width=16)
        table.add_column("ID", style="dim", max_width=20, overflow="ellipsis")

        color_styles = {
            "blue": "[bold blue]Blue[/]",
            "green": "[bold green]Green[/]",
            "pink": "[bold magenta]Pink[/]",
            "yellow": "[bold yellow]Yellow[/]",
            "white": "[white]White[/]",
        }

        for note in notes:
            color = color_styles.get(note.color, note.color)
            subject = note.subject or "(no subject)"
            modified_str = note.modified.strftime("%Y-%m-%d %H:%M") if note.modified else ""
            short_id = note.entry_id[:20] + "..." if note.entry_id and len(note.entry_id) > 20 else (note.entry_id or "")

            table.add_row(
                color,
                subject,
                modified_str,
                short_id,
            )

        self.console.print(table)
        self.console.print(f"\n[dim]Total: {len(notes)} note(s)[/dim]")

    def print_note_detail(self, note: Note) -> None:
        """Print a single note in detail."""
        # Header info
        header_lines = []

        # Color
        color_styles = {
            "blue": "[bold blue]Blue[/]",
            "green": "[bold green]Green[/]",
            "pink": "[bold magenta]Pink[/]",
            "yellow": "[bold yellow]Yellow[/]",
            "white": "[white]White[/]",
        }
        header_lines.append(f"[bold]Color:[/bold] {color_styles.get(note.color, note.color)}")

        # Dates
        if note.created:
            header_lines.append(f"[bold]Created:[/bold] {note.created.strftime('%Y-%m-%d %H:%M')}")
        if note.modified:
            header_lines.append(f"[bold]Modified:[/bold] {note.modified.strftime('%Y-%m-%d %H:%M')}")

        # Categories
        if note.categories:
            header_lines.append(f"[bold]Categories:[/bold] {', '.join(note.categories)}")

        header = "\n".join(header_lines)

        # Print header panel
        title = note.subject or "(no subject)"
        self.console.print(Panel(
            header,
            title=f"[bold]{title}[/bold]",
            border_style="yellow",
            padding=(1, 2),
        ))

        # Print body
        if note.body:
            self.console.print()
            self.console.print(note.body)
        else:
            self.console.print("\n[dim](empty note)[/dim]")

        # Print note ID
        if note.entry_id:
            self.console.print(f"\n[dim]Note ID: {note.entry_id}[/dim]")
