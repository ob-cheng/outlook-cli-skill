"""Core modules for Outlook connection and data models."""

from .models import Email
from .folders import list_all_folders, find_folder_by_name, find_folder_by_path, DEFAULT_FOLDERS

# Lazy import: avoid importing win32com at module level so that
# the data models and utilities are importable without pywin32.
def connect_to_outlook():
    """Establish connection to Outlook via COM.

    Lazy-imported to avoid ImportError when pywin32 is not installed.
    """
    from .connection import connect_to_outlook as _connect
    return _connect()


__all__ = [
    'connect_to_outlook',
    'Email',
    'list_all_folders',
    'find_folder_by_name',
    'find_folder_by_path',
    'DEFAULT_FOLDERS',
]
