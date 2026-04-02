import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtCore    import Qt
from database.db_manager  import DatabaseManager
from ui.main_window       import MainWindow
from ui.startup_dialog    import StartupDialog
from resources.file_paths import get_resource

from qt_material import apply_stylesheet

THEMES = {
    'light': 'light_blue.xml',
    'dark':  'dark_teal.xml',
}

_EXTRA_QSS = """
    QLabel[class="title"] {
        font-size: 18pt;
        font-weight: bold;
        margin: 8px;
    }
    QLabel[class="total"] {
        font-weight: bold;
        font-size: 12pt;
    }
    QPushButton[class="danger"] {
        background-color: #e53935;
        color: white;
    }
    QPushButton[class="danger"]:hover {
        background-color: #c62828;
    }
    QGroupBox {
        font-weight: bold;
        margin-top: 8px;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 4px;
    }
"""

# Keep a reference to the active window so apply_theme can reach the sidebar
_active_window: MainWindow | None = None


def apply_theme(app: QApplication, mode: str = 'light'):
    """Apply Qt Material theme + custom overrides, then update all themed widgets."""
    global _active_window
    theme = THEMES.get(mode, THEMES['light'])
    apply_stylesheet(app, theme=theme, invert_secondary=(mode == 'light'))
    app.setStyleSheet(app.styleSheet() + _EXTRA_QSS)

    if _active_window is not None:
        # Sidebar — hardcoded colours need manual update
        if hasattr(_active_window, 'sidebar'):
            _active_window.sidebar.set_theme(mode)
        # Dashboard — hardcoded card/label colours need manual update
        if hasattr(_active_window, '_dashboard_widget'):
            _active_window._dashboard_widget.set_theme(mode)


def main():
    global _active_window

    app = QApplication(sys.argv)
    app.setApplicationName('ESQ Company Accounting System')

    apply_theme(app, mode='light')

    startup = StartupDialog()
    if startup.exec() != StartupDialog.Accepted or not startup.db_path:
        sys.exit(0)

    db_manager = DatabaseManager(startup.db_path)
    db_manager.initialize_database(
        coa_xlsx_path=startup.coa_xlsx_path,
        use_default_coa=startup._use_default_coa,
    )

    window = MainWindow(db_manager, apply_theme_fn=apply_theme)
    _active_window = window
    window.showMaximized()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()