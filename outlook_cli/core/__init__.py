"""Core modules for Outlook connection and data models."""

from .connection import connect_to_outlook
from .models import Email
from .folders import list_all_folders, find_folder_by_name, find_folder_by_path, DEFAULT_FOLDERS

__all__ = [
    'connect_to_outlook',
    'Email',
    'list_all_folders',
    'find_folder_by_name',
    'find_folder_by_path',
    'DEFAULT_FOLDERS',
]
