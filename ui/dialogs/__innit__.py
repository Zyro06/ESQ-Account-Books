# ui/dialogs/__init__.py
# Convenience re-exports so callers can do:
#   from ui.dialogs import LineDialog, ViewDetailsDialog, ViewEntryDialog

from ui.dialogs.line_dialog         import LineDialog
from ui.dialogs.view_details_dialog import ViewDetailsDialog, ViewEntryDialog

__all__ = [
    'LineDialog',
    'ViewDetailsDialog',
    'ViewEntryDialog',
]