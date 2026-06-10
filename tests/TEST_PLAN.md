# Outlook CLI Test Suite Design

## Overview

This test suite provides comprehensive coverage for the outlook-cli skill. All tests are designed to run **without COM dependencies** using monkeypatching and mocks, making them runnable on WSL/Linux/CI environments.

## Test Architecture

### Design Principles

1. **COM-Free**: All tests use mocks/monkeypatching - no win32com dependency
2. **Isolated**: Each test is self-contained with its own fixtures
3. **Fast**: No network, no file I/O (except temp files), no external dependencies
4. **Comprehensive**: Cover happy paths, edge cases, and error conditions

### Test Categories

| Category | Files | Coverage Target |
|----------|-------|-----------------|
| **Unit Tests** | `test_*.py` | Individual functions/classes |
| **Integration Tests** | `test_integration_*.py` | Cross-module workflows |
| **SKILL Coverage** | `test_skill_*.py` | SKILL.md documented behaviors |

## Module Coverage Matrix

### Core Modules (`outlook_cli/core/`)

| Module | Test File | Classes/Functions | Status |
|--------|-----------|-------------------|--------|
| `config.py` | `test_config.py` | `ConfigManager` (load, save, get, set, clear, show, status_tags) | NEW |
| `people.py` | `test_people.py` | `PeopleManager` (list, lookup, add, extract_and_add) | NEW |
| `models.py` | `test_models.py` | `Email` dataclass, `from_mail_item()`, `to_dict()` | NEW |
| `folders.py` | `test_folders.py` | `list_all_folders`, `find_folder_*` | EXISTS (expand) |
| `connection.py` | `test_connection.py` | `connect_to_outlook()` | NEW |
| `wsl.py` | `test_wsl.py` | `is_wsl()`, `find_windows_python()`, `check_wsl_environment()` | NEW |
| `last_search.py` | `test_last_search.py` | Cache management functions | NEW |

### Service Modules (`outlook_cli/services/`)

| Module | Test File | Classes/Functions | Status |
|--------|-----------|-------------------|--------|
| `search.py` | `test_search.py` | `SearchService` | EXISTS (expand) |
| `compose.py` | `test_compose.py` | `ComposeService` (send_email, reply, forward) | EXISTS (expand) |
| `calendar.py` | `test_calendar.py` | `CalendarEvent`, `CalendarService` | NEW |
| `tasks.py` | `test_tasks.py` | `Task`, `TaskService` | NEW |
| `notes.py` | `test_notes.py` | `Note`, `NotesService` | NEW |
| `export.py` | `test_export.py` | `ExportService` (export_emails, state management) | NEW |
| `viewer.py` | `test_viewer.py` | `ViewerService` (print_* methods) | NEW |

### Utility Modules (`outlook_cli/utils/`)

| Module | Test File | Functions | Status |
|--------|-----------|-----------|--------|
| `formatting.py` | `test_formatting.py` | 10 parsing/formatting functions | NEW |

### CLI Layer

| Module | Test File | Coverage | Status |
|--------|-----------|----------|--------|
| `cli.py` | `test_argparse.py` | Argument parsing | EXISTS |
| `cli.py` | `test_cli_commands.py` | Command dispatch logic | NEW |

## SKILL.md Coverage Tests

These tests verify that documented SKILL behaviors work as specified:

| SKILL Feature | Test | Verification |
|---------------|------|--------------|
| Search → Read → Act workflow | `test_skill_search_read_act` | Search returns IDs, read accepts ID |
| Draft-only default | `test_skill_draft_default` | compose creates draft without --send |
| Direct send requires both | `test_skill_direct_send` | Requires send_mode=send AND --send flag |
| People auto-discovery | `test_skill_people_auto` | read command adds unknown people |
| Multi-account folder paths | `test_skill_multi_account` | `AccountName/Inbox` format works |
| Date filtering | `test_skill_date_filtering` | --days, --from-date, --to-date |
| Export incremental | `test_skill_export_incremental` | --incremental tracks state |
| JSON output | `test_skill_json_output` | All commands support --json |

## Test File Structure

```
tests/
├── conftest.py                 # Shared fixtures (EXISTS - expand)
├── TEST_PLAN.md               # This document
│
├── # Core module tests
├── test_config.py             # ConfigManager tests
├── test_people.py             # PeopleManager tests
├── test_models.py             # Email/Event/Task/Note dataclasses
├── test_folders.py            # Folder traversal (EXISTS - expand)
├── test_connection.py         # Connection error handling
├── test_wsl.py                # WSL detection
├── test_last_search.py        # Search cache
│
├── # Service tests
├── test_search.py             # SearchService (EXISTS - expand)
├── test_compose.py            # ComposeService (EXISTS - expand)
├── test_calendar.py           # CalendarService
├── test_tasks.py              # TaskService
├── test_notes.py              # NotesService
├── test_export.py             # ExportService
├── test_viewer.py             # ViewerService
│
├── # Utility tests
├── test_formatting.py         # Parsing/formatting utilities
├── test_progress.py           # Progress dots (EXISTS)
│
├── # CLI tests
├── test_argparse.py           # Argument parsing (EXISTS)
├── test_cli_commands.py       # Command handlers
│
├── # SKILL compliance tests
└── test_skill_compliance.py   # SKILL.md documented behaviors
```

## Running Tests

```bash
# All tests (from project root)
pytest tests/ -v

# Specific module
pytest tests/test_config.py -v

# With coverage
pytest tests/ --cov=outlook_cli --cov-report=html

# Fast smoke test
pytest tests/ -x -q
```

## Mock Patterns

### Pattern 1: Monkeypatch Email.from_mail_item
```python
with patch.object(Email, "from_mail_item") as mock_from:
    mock_from.side_effect = lambda mail, is_sent=False: mail._test_email
    result = svc.search()
```

### Pattern 2: Temp file for config/people
```python
@pytest.fixture
def temp_config(tmp_path):
    config_file = tmp_path / "config.json"
    return ConfigManager(config_path=config_file)
```

### Pattern 3: Mock COM namespace
```python
@pytest.fixture
def mock_namespace():
    ns = MagicMock()
    ns.GetDefaultFolder.side_effect = get_default_folder
    return ns
```

## Coverage Goals

- **Line Coverage**: ≥85%
- **Branch Coverage**: ≥75%
- **Critical Paths**: 100% (compose, search, export)

## Test Suite Summary (2026-06-10)

**Total Tests: 344 (all passing)**

| Test File | Test Count | Coverage |
|-----------|------------|----------|
| test_config.py | 24 | ConfigManager CRUD, status tags |
| test_people.py | 19 | PeopleManager CRUD, lookup, bulk import |
| test_formatting.py | 33 | Email parsing, date handling, sanitization |
| test_models.py | 22 | Email dataclass, from_mail_item, to_dict |
| test_calendar.py | 23 | CalendarEvent, CalendarService |
| test_tasks.py | 26 | Task, TaskService |
| test_notes.py | 25 | Note, NotesService |
| test_export.py | 28 | ExportService, markdown/JSON export |
| test_last_search.py | 14 | Search result caching |
| test_wsl.py | 12 | WSL detection, Windows Python discovery |
| test_skill_compliance.py | 20 | SKILL.md documented behaviors |
| test_argparse.py | 20 | CLI argument parsing |
| test_search.py | 9 | SearchService filtering, limits |
| test_compose.py | 6 | CC/BCC parsing |
| test_folders.py | 7 | Folder traversal |
| test_progress.py | 7 | Progress indicator |

### Key Testing Patterns

1. **COM-Free**: All tests use mocks/monkeypatching
2. **Temp Files**: Config/people tests use `tmp_path` fixture
3. **Service Mocks**: `patch("outlook_cli.core.folders.find_folder_across_stores", return_value=None)` for services
4. **Factory Functions**: `make_test_email()`, `make_mock_appointment()`, etc.
