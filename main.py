import sys
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt
from database.db_manager import DatabaseManager
from ui.main_window import MainWindow
from ui.startup_dialog import StartupDialog
from resources.style_loader import load_stylesheet
from resources.file_paths import get_resource


def main():
    # High DPI support for PyQt5
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setApplicationName("ESQ Company Accounting System")

    # Load and apply stylesheet from resources/
    stylesheet = load_stylesheet(get_resource('style.qss'))
    app.setStyleSheet(stylesheet)

    # ── Startup dialog ──────────────────────────────────────────────────────
    startup = StartupDialog()
    if startup.exec_() != StartupDialog.Accepted or not startup.db_path:
        sys.exit(0)

    db_path       = startup.db_path
    coa_xlsx_path = startup.coa_xlsx_path   # None → use built-in default

    # ── Initialise database ─────────────────────────────────────────────────
    db_manager = DatabaseManager(db_path)
    db_manager.initialize_database(
        coa_xlsx_path=coa_xlsx_path,
        use_default_coa=startup._use_default_coa
    )

    # ── Main window ─────────────────────────────────────────────────────────
    window = MainWindow(db_manager)
    window.showMaximized()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()