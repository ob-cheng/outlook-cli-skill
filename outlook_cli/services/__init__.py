"""Services for searching, viewing, exporting, composing emails and managing calendar."""

from .search import SearchService
from .viewer import ViewerService
from .compose import ComposeService
from .calendar import CalendarService, CalendarEvent

# Lazy imports for heavy dependencies (beautifulsoup4, markdownify)
def get_export_service(output_dir, state_dir=None):
    """Lazy-load ExportService to avoid importing bs4/markdownify at module level."""
    from .export import ExportService
    return ExportService(output_dir, state_dir)

__all__ = ['SearchService', 'ViewerService', 'ComposeService', 'CalendarService', 'CalendarEvent', 'get_export_service']
