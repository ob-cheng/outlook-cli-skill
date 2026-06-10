"""Services for searching, viewing, exporting, composing emails and managing calendar."""

from .search import SearchService
from .viewer import ViewerService
from .calendar import CalendarService, CalendarEvent

# Lazy imports for heavy dependencies (win32com, beautifulsoup4, markdownify)
def get_export_service(output_dir, state_dir=None):
    """Lazy-load ExportService to avoid importing bs4/markdownify at module level."""
    from .export import ExportService
    return ExportService(output_dir, state_dir)


def get_compose_service(namespace):
    """Lazy-load ComposeService to avoid importing win32com at module level."""
    from .compose import ComposeService
    return ComposeService(namespace)


__all__ = ['SearchService', 'ViewerService', 'CalendarService', 'CalendarEvent',
           'get_export_service', 'get_compose_service']
