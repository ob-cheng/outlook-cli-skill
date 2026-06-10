"""Tests for TaskService and Task dataclass.

Uses mock COM objects to avoid win32com dependency.
"""

import pytest
from datetime import datetime
from unittest.mock import MagicMock, patch

from outlook_cli.services.tasks import Task, TaskService


def make_mock_task_item(
    subject="Complete report",
    status=0,  # 0=NotStarted, 1=InProgress, 2=Complete, 3=Waiting, 4=Deferred
    due_date=None,
    start_date=None,
    completed_date=None,
    percent_complete=0,
    importance=1,  # 0=Low, 1=Normal, 2=High
    body="Task description here",
    categories="",
    reminder_set=False,
    reminder_time=None,
    entry_id="task-id-123",
):
    """Create a mock Outlook TaskItem."""
    task = MagicMock()
    task.Subject = subject
    task.Status = status
    task.DueDate = due_date or datetime(2024, 3, 20)
    task.StartDate = start_date
    task.DateCompleted = completed_date
    task.PercentComplete = percent_complete
    task.Importance = importance
    task.Body = body
    task.Categories = categories
    task.ReminderSet = reminder_set
    task.ReminderTime = reminder_time
    task.EntryID = entry_id

    return task


class TestTaskDataclass:
    """Tests for Task dataclass basics."""

    def test_task_creation(self):
        """Task can be created with all fields."""
        task = Task(
            subject="Complete report",
            status="not_started",
            due_date=datetime(2024, 3, 20),
            priority="high",
        )

        assert task.subject == "Complete report"
        assert task.status == "not_started"
        assert task.priority == "high"

    def test_task_defaults(self):
        """Task has sensible defaults."""
        task = Task(
            subject="Task",
            status="not_started",
        )

        assert task.due_date is None
        assert task.start_date is None
        assert task.completed_date is None
        assert task.percent_complete == 0
        assert task.priority == "normal"
        assert task.body is None
        assert task.categories == []
        assert task.reminder_date is None


class TestTaskFromTaskItem:
    """Tests for Task.from_task_item factory."""

    def test_basic_conversion(self):
        """Converts basic task item correctly."""
        task_item = make_mock_task_item()

        task = Task.from_task_item(task_item)

        assert task.subject == "Complete report"
        assert task.entry_id == "task-id-123"

    def test_handles_missing_subject(self):
        """Missing subject becomes '(No Subject)'."""
        task_item = make_mock_task_item(subject=None)

        task = Task.from_task_item(task_item)

        assert task.subject == "(No Subject)"

    def test_status_not_started(self):
        """Status 0 becomes 'not_started'."""
        task_item = make_mock_task_item(status=0)

        task = Task.from_task_item(task_item)

        assert task.status == "not_started"

    def test_status_in_progress(self):
        """Status 1 becomes 'in_progress'."""
        task_item = make_mock_task_item(status=1)

        task = Task.from_task_item(task_item)

        assert task.status == "in_progress"

    def test_status_completed(self):
        """Status 2 becomes 'completed'."""
        task_item = make_mock_task_item(status=2)

        task = Task.from_task_item(task_item)

        assert task.status == "completed"

    def test_status_waiting(self):
        """Status 3 becomes 'waiting'."""
        task_item = make_mock_task_item(status=3)

        task = Task.from_task_item(task_item)

        assert task.status == "waiting"

    def test_status_deferred(self):
        """Status 4 becomes 'deferred'."""
        task_item = make_mock_task_item(status=4)

        task = Task.from_task_item(task_item)

        assert task.status == "deferred"

    def test_priority_low(self):
        """Importance 0 becomes 'low'."""
        task_item = make_mock_task_item(importance=0)

        task = Task.from_task_item(task_item)

        assert task.priority == "low"

    def test_priority_normal(self):
        """Importance 1 becomes 'normal'."""
        task_item = make_mock_task_item(importance=1)

        task = Task.from_task_item(task_item)

        assert task.priority == "normal"

    def test_priority_high(self):
        """Importance 2 becomes 'high'."""
        task_item = make_mock_task_item(importance=2)

        task = Task.from_task_item(task_item)

        assert task.priority == "high"

    def test_extracts_categories(self):
        """Extracts categories from comma-separated string."""
        task_item = make_mock_task_item(categories="Work, Urgent")

        task = Task.from_task_item(task_item)

        assert "Work" in task.categories
        assert "Urgent" in task.categories

    def test_extracts_percent_complete(self):
        """Extracts percent complete."""
        task_item = make_mock_task_item(percent_complete=50)

        task = Task.from_task_item(task_item)

        assert task.percent_complete == 50

    def test_handles_invalid_dates(self):
        """Handles dates outside valid range."""
        task_item = make_mock_task_item()
        # Simulate invalid date year by setting DueDate to a mock
        mock_date = MagicMock()
        mock_date.year = 4501  # Invalid year
        task_item.DueDate = mock_date

        task = Task.from_task_item(task_item)

        assert task.due_date is None


class TestTaskToDict:
    """Tests for Task.to_dict serialization."""

    def test_serializes_all_fields(self):
        """Serializes all fields correctly."""
        task = Task(
            subject="Complete report",
            status="in_progress",
            due_date=datetime(2024, 3, 20),
            start_date=datetime(2024, 3, 15),
            percent_complete=50,
            priority="high",
            body="Work on report",
            entry_id="test-id",
        )

        result = task.to_dict()

        assert result["subject"] == "Complete report"
        assert result["status"] == "in_progress"
        assert result["due_date"] == "2024-03-20T00:00:00"
        assert result["percent_complete"] == 50
        assert result["priority"] == "high"
        assert result["entry_id"] == "test-id"

    def test_handles_none_dates(self):
        """Handles None dates."""
        task = Task(
            subject="Task",
            status="not_started",
            due_date=None,
            start_date=None,
        )

        result = task.to_dict()

        assert result["due_date"] is None
        assert result["start_date"] is None


class TestTaskServiceListTasks:
    """Tests for TaskService.list_tasks."""

    @pytest.fixture
    def mock_namespace(self):
        """Create mock namespace with tasks folder."""
        ns = MagicMock()

        tasks_folder = MagicMock()
        items = MagicMock()
        tasks_folder.Items = items

        ns.GetDefaultFolder.return_value = tasks_folder

        return ns

    def test_list_returns_tasks(self, mock_namespace):
        """Returns list of Task objects."""
        task_item = make_mock_task_item()
        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 1
        items.Item.return_value = task_item

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            tasks = svc.list_tasks()

        assert len(tasks) == 1
        assert tasks[0].subject == "Complete report"

    def test_excludes_completed_by_default(self, mock_namespace):
        """Excludes completed tasks by default."""
        completed_task = make_mock_task_item(status=2)  # Completed
        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 1
        items.Item.return_value = completed_task

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            tasks = svc.list_tasks()

        assert len(tasks) == 0

    def test_includes_completed_when_requested(self, mock_namespace):
        """Includes completed tasks when requested."""
        completed_task = make_mock_task_item(status=2)
        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 1
        items.Item.return_value = completed_task

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            tasks = svc.list_tasks(include_completed=True)

        assert len(tasks) == 1

    def test_filters_by_status(self, mock_namespace):
        """Filters tasks by status."""
        task1 = make_mock_task_item(status=0)  # not_started
        task2 = make_mock_task_item(status=1)  # in_progress

        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 2
        items.Item.side_effect = [task1, task2]

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            tasks = svc.list_tasks(status="in_progress")

        assert len(tasks) == 1
        assert tasks[0].status == "in_progress"

    def test_filters_by_priority(self, mock_namespace):
        """Filters tasks by priority."""
        task1 = make_mock_task_item(importance=1)  # normal
        task2 = make_mock_task_item(importance=2)  # high

        items = mock_namespace.GetDefaultFolder.return_value.Items
        items.Count = 2
        items.Item.side_effect = [task1, task2]

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            tasks = svc.list_tasks(priority="high")

        assert len(tasks) == 1
        assert tasks[0].priority == "high"


class TestTaskServiceCreateTask:
    """Tests for TaskService.create_task."""

    @pytest.fixture
    def mock_namespace(self):
        """Create mock namespace with tasks folder."""
        ns = MagicMock()
        tasks_folder = MagicMock()
        ns.GetDefaultFolder.return_value = tasks_folder

        # Mock Application.CreateItem
        mock_task = MagicMock()
        mock_task.EntryID = "new-task-id"
        ns.Application.CreateItem.return_value = mock_task

        return ns

    def test_create_basic_task(self, mock_namespace):
        """Creates task with required fields."""
        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            success, entry_id = svc.create_task(
                subject="New Task",
            )

        assert success is True
        assert entry_id == "new-task-id"

    def test_create_task_with_due_date(self, mock_namespace):
        """Creates task with due date."""
        mock_task = mock_namespace.Application.CreateItem.return_value

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            svc.create_task(
                subject="Task",
                due_date=datetime(2024, 3, 25),
            )

        assert mock_task.DueDate is not None

    def test_create_task_with_priority(self, mock_namespace):
        """Creates task with specified priority."""
        mock_task = mock_namespace.Application.CreateItem.return_value

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(mock_namespace)
            svc.create_task(
                subject="Task",
                priority="high",
            )

        assert mock_task.Importance == TaskService.OL_IMPORTANCE_HIGH


class TestTaskServiceCompleteTask:
    """Tests for TaskService.complete_task."""

    def test_complete_task_success(self):
        """Marks task as complete."""
        ns = MagicMock()
        mock_task = MagicMock()
        ns.GetItemFromID.return_value = mock_task
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(ns)
            success, message = svc.complete_task("test-id")

        assert success is True
        assert mock_task.Status == TaskService.OL_TASK_COMPLETE
        assert mock_task.PercentComplete == 100
        mock_task.Save.assert_called_once()

    def test_complete_task_not_found(self):
        """Returns error when task not found."""
        ns = MagicMock()
        ns.GetItemFromID.side_effect = Exception("Item not found")
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(ns)
            success, message = svc.complete_task("nonexistent-id")

        assert success is False
        assert "Failed" in message


class TestTaskServiceDeleteTask:
    """Tests for TaskService.delete_task."""

    def test_delete_success(self):
        """Deletes task successfully."""
        ns = MagicMock()
        mock_task = MagicMock()
        ns.GetItemFromID.return_value = mock_task
        ns.GetDefaultFolder.return_value = MagicMock()

        with patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None):
            svc = TaskService(ns)
            success, message = svc.delete_task("test-id")

        assert success is True
        mock_task.Delete.assert_called_once()
