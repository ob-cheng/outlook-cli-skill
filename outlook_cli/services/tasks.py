"""Task service for managing Outlook tasks/todos."""

from datetime import datetime
from dataclasses import dataclass
from typing import Optional


@dataclass
class Task:
    """Represents an Outlook task."""
    subject: str
    status: str  # not_started, in_progress, completed, waiting, deferred
    due_date: Optional[datetime] = None
    start_date: Optional[datetime] = None
    completed_date: Optional[datetime] = None
    percent_complete: int = 0
    priority: str = "normal"  # low, normal, high
    body: Optional[str] = None
    categories: list[str] = None
    reminder_date: Optional[datetime] = None
    entry_id: Optional[str] = None

    def __post_init__(self):
        if self.categories is None:
            self.categories = []

    @classmethod
    def from_task_item(cls, task_item) -> 'Task':
        """Create Task from Outlook TaskItem."""
        try:
            subject = task_item.Subject or "(No Subject)"
        except Exception:
            subject = "(No Subject)"

        # Status mapping: 0=NotStarted, 1=InProgress, 2=Complete, 3=Waiting, 4=Deferred
        status_map = {
            0: "not_started",
            1: "in_progress",
            2: "completed",
            3: "waiting",
            4: "deferred",
        }
        try:
            status = status_map.get(task_item.Status, "not_started")
        except Exception:
            status = "not_started"

        try:
            due_date = task_item.DueDate if task_item.DueDate and 1900 < task_item.DueDate.year < 4500 else None
        except Exception:
            due_date = None

        try:
            start_date = task_item.StartDate if task_item.StartDate and 1900 < task_item.StartDate.year < 4500 else None
        except Exception:
            start_date = None

        try:
            completed_date = task_item.DateCompleted if task_item.DateCompleted and 1900 < task_item.DateCompleted.year < 4500 else None
        except Exception:
            completed_date = None

        try:
            percent_complete = task_item.PercentComplete or 0
        except Exception:
            percent_complete = 0

        # Priority mapping: 0=Low, 1=Normal, 2=High
        priority_map = {0: "low", 1: "normal", 2: "high"}
        try:
            priority = priority_map.get(task_item.Importance, "normal")
        except Exception:
            priority = "normal"

        try:
            body = task_item.Body or None
        except Exception:
            body = None

        try:
            categories = [c.strip() for c in task_item.Categories.split(',')] if task_item.Categories else []
        except Exception:
            categories = []

        try:
            reminder_date = task_item.ReminderTime if task_item.ReminderSet else None
        except Exception:
            reminder_date = None

        try:
            entry_id = task_item.EntryID
        except Exception:
            entry_id = None

        return cls(
            subject=subject,
            status=status,
            due_date=due_date,
            start_date=start_date,
            completed_date=completed_date,
            percent_complete=percent_complete,
            priority=priority,
            body=body,
            categories=categories,
            reminder_date=reminder_date,
            entry_id=entry_id,
        )

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            'entry_id': self.entry_id,
            'subject': self.subject,
            'status': self.status,
            'due_date': self.due_date.isoformat() if self.due_date else None,
            'start_date': self.start_date.isoformat() if self.start_date else None,
            'completed_date': self.completed_date.isoformat() if self.completed_date else None,
            'percent_complete': self.percent_complete,
            'priority': self.priority,
            'body': self.body,
            'categories': self.categories,
            'reminder_date': self.reminder_date.isoformat() if self.reminder_date else None,
        }


class TaskService:
    """Service for managing Outlook tasks."""

    # Outlook constants
    OL_TASK_NOT_STARTED = 0
    OL_TASK_IN_PROGRESS = 1
    OL_TASK_COMPLETE = 2
    OL_TASK_WAITING = 3
    OL_TASK_DEFERRED = 4

    OL_IMPORTANCE_LOW = 0
    OL_IMPORTANCE_NORMAL = 1
    OL_IMPORTANCE_HIGH = 2

    def __init__(self, namespace):
        """Initialize with Outlook MAPI namespace."""
        self.namespace = namespace
        self.tasks_folder = namespace.GetDefaultFolder(13)  # 13 = olFolderTasks

    def list_tasks(
        self,
        status: str | None = None,
        include_completed: bool = False,
        due_before: datetime | None = None,
        due_after: datetime | None = None,
        priority: str | None = None,
        category: str | None = None,
    ) -> list[Task]:
        """List tasks with filters.

        Args:
            status: Filter by status (not_started, in_progress, completed, waiting, deferred)
            include_completed: Include completed tasks (default: False)
            due_before: Filter tasks due before this date
            due_after: Filter tasks due after this date
            priority: Filter by priority (low, normal, high)
            category: Filter by category name

        Returns:
            List of Task objects
        """
        tasks = []

        status_to_outlook = {
            "not_started": 0,
            "in_progress": 1,
            "completed": 2,
            "waiting": 3,
            "deferred": 4,
        }

        priority_to_outlook = {
            "low": 0,
            "normal": 1,
            "high": 2,
        }

        try:
            items = self.tasks_folder.Items
            items.Sort("[DueDate]")

            for i in range(1, items.Count + 1):
                try:
                    task_item = items.Item(i)

                    # Skip completed unless explicitly included or filtering for completed
                    if not include_completed and status != "completed":
                        if task_item.Status == self.OL_TASK_COMPLETE:
                            continue

                    # Status filter
                    if status and task_item.Status != status_to_outlook.get(status, -1):
                        continue

                    # Due date filters
                    if due_before or due_after:
                        try:
                            due = task_item.DueDate
                            if due and 1900 < due.year < 4500:
                                if due_before and due > due_before:
                                    continue
                                if due_after and due < due_after:
                                    continue
                            elif due_before or due_after:
                                continue  # Skip tasks without due date when filtering by date
                        except Exception:
                            continue

                    # Priority filter
                    if priority and task_item.Importance != priority_to_outlook.get(priority, -1):
                        continue

                    # Category filter
                    if category:
                        task_categories = task_item.Categories or ""
                        if category.lower() not in task_categories.lower():
                            continue

                    task = Task.from_task_item(task_item)
                    tasks.append(task)

                except Exception:
                    continue

        except Exception:
            pass

        return tasks

    def get_task(self, entry_id: str) -> Task | None:
        """Get a single task by EntryID.

        Args:
            entry_id: The task's EntryID

        Returns:
            Task or None if not found
        """
        try:
            task_item = self.namespace.GetItemFromID(entry_id)
            return Task.from_task_item(task_item)
        except Exception:
            return None

    def create_task(
        self,
        subject: str,
        due_date: datetime | None = None,
        start_date: datetime | None = None,
        priority: str = "normal",
        body: str | None = None,
        categories: list[str] | None = None,
        reminder_date: datetime | None = None,
    ) -> tuple[bool, str]:
        """Create a new task.

        Args:
            subject: Task subject
            due_date: Due date
            start_date: Start date
            priority: Priority (low, normal, high)
            body: Task description
            categories: List of category names
            reminder_date: Reminder date/time

        Returns:
            tuple: (success: bool, entry_id_or_error: str)
        """
        priority_map = {
            "low": self.OL_IMPORTANCE_LOW,
            "normal": self.OL_IMPORTANCE_NORMAL,
            "high": self.OL_IMPORTANCE_HIGH,
        }

        try:
            outlook = self.namespace.Application
            task_item = outlook.CreateItem(3)  # 3 = olTaskItem

            task_item.Subject = subject

            if due_date:
                # Set to noon to avoid timezone boundary issues
                task_item.DueDate = due_date.replace(hour=12, minute=0, second=0)

            if start_date:
                task_item.StartDate = start_date.replace(hour=12, minute=0, second=0)

            task_item.Importance = priority_map.get(priority, self.OL_IMPORTANCE_NORMAL)

            if body:
                task_item.Body = body

            if categories:
                task_item.Categories = ", ".join(categories)

            if reminder_date:
                task_item.ReminderSet = True
                task_item.ReminderTime = reminder_date
            else:
                task_item.ReminderSet = False

            task_item.Save()
            return True, task_item.EntryID

        except Exception as e:
            return False, f"Failed to create task: {e}"

    def update_task(
        self,
        entry_id: str,
        subject: str | None = None,
        due_date: datetime | None = None,
        start_date: datetime | None = None,
        priority: str | None = None,
        body: str | None = None,
        percent_complete: int | None = None,
    ) -> tuple[bool, str]:
        """Update an existing task.

        Args:
            entry_id: Task EntryID
            subject: New subject (None = no change)
            due_date: New due date (None = no change)
            start_date: New start date (None = no change)
            priority: New priority (None = no change)
            body: New body (None = no change)
            percent_complete: New percent complete (None = no change)

        Returns:
            tuple: (success: bool, message: str)
        """
        priority_map = {
            "low": self.OL_IMPORTANCE_LOW,
            "normal": self.OL_IMPORTANCE_NORMAL,
            "high": self.OL_IMPORTANCE_HIGH,
        }

        try:
            task_item = self.namespace.GetItemFromID(entry_id)

            if subject is not None:
                task_item.Subject = subject
            if due_date is not None:
                task_item.DueDate = due_date
            if start_date is not None:
                task_item.StartDate = start_date
            if priority is not None:
                task_item.Importance = priority_map.get(priority, self.OL_IMPORTANCE_NORMAL)
            if body is not None:
                task_item.Body = body
            if percent_complete is not None:
                task_item.PercentComplete = percent_complete

            task_item.Save()
            return True, "Task updated successfully"

        except Exception as e:
            return False, f"Failed to update task: {e}"

    def complete_task(self, entry_id: str) -> tuple[bool, str]:
        """Mark a task as complete.

        Args:
            entry_id: Task EntryID

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            task_item = self.namespace.GetItemFromID(entry_id)
            task_item.Status = self.OL_TASK_COMPLETE
            task_item.PercentComplete = 100
            task_item.DateCompleted = datetime.now()
            task_item.Save()
            return True, "Task marked as complete"
        except Exception as e:
            return False, f"Failed to complete task: {e}"

    def delete_task(self, entry_id: str) -> tuple[bool, str]:
        """Delete a task.

        Args:
            entry_id: Task EntryID

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            task_item = self.namespace.GetItemFromID(entry_id)
            task_item.Delete()
            return True, "Task deleted successfully"
        except Exception as e:
            return False, f"Failed to delete task: {e}"
