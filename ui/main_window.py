import sys
import os
import shutil
from datetime import datetime

from PyQt5.QtWidgets import (
    QMainWindow, QTabWidget, QStatusBar, QWidget, QVBoxLayout, QLabel,
    QMenuBar, QMenu, QAction, QMessageBox, QFileDialog, QProgressDialog,
    QInputDialog, QLineEdit, QDialog, QScrollArea, QTextBrowser,
    QDialogButtonBox, QHBoxLayout, QApplication, QFormLayout,
    QSpinBox, QComboBox, QGroupBox, QCheckBox,
)
from PyQt5.QtCore import Qt, QObject, QEvent
from PyQt5.QtGui import QIcon, QKeySequence, QFont
from PyQt5.QtWidgets import QShortcut

from database.db_manager import DatabaseManager
from resources.file_paths import get_resource
from ui.coa_widget import COAWidget
from ui.alphalist_widget import AlphalistWidget
from ui.sales_journal_widget import SalesJournalWidget
from ui.purchase_journal_widget import PurchaseJournalWidget
from ui.cash_disbursement_widget import CashDisbursementWidget
from ui.cash_receipts_widget import CashReceiptsWidget
from ui.general_journal_widget import GeneralJournalWidget
from ui.general_ledger_widget import GeneralLedgerWidget
from ui.trial_balance_widget import TrialBalanceWidget
from ui.financial_statements_widget import FinancialStatementsWidget


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _saves_dir() -> str:
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(base, 'data', 'saves')
    os.makedirs(path, exist_ok=True)
    return path


def _backups_dir() -> str:
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(base, 'data', 'backups')
    os.makedirs(path, exist_ok=True)
    return path


def _recent_file() -> str:
    """Path to the JSON file that tracks recently opened accounts."""
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(base, 'data')
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, 'recent.json')


def _load_recent() -> list:
    import json
    try:
        with open(_recent_file(), 'r') as f:
            items = json.load(f)
        return [p for p in items if os.path.exists(p)]
    except Exception:
        return []


def _save_recent(path: str):
    import json
    items = _load_recent()
    if path in items:
        items.remove(path)
    items.insert(0, path)
    items = items[:8]           # keep last 8
    try:
        with open(_recent_file(), 'w') as f:
            json.dump(items, f)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Tab key filter
# ---------------------------------------------------------------------------

class TabKeyFilter(QObject):
    def __init__(self, main_window: 'MainWindow'):
        super().__init__(main_window)
        self._main = main_window

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if event.type() == QEvent.KeyPress:
            key  = event.key()
            mods = event.modifiers()
            ctrl       = mods == Qt.ControlModifier
            ctrl_shift = mods == (Qt.ControlModifier | Qt.ShiftModifier)
            if ctrl and key in (Qt.Key_Tab, Qt.Key_Backtab):
                self._main._switch_tab_forward()
                return True
            if ctrl_shift and key in (Qt.Key_Tab, Qt.Key_Backtab):
                self._main._switch_tab_backward()
                return True
        return False


# ---------------------------------------------------------------------------
# Preferences dialog
# ---------------------------------------------------------------------------

class PreferencesDialog(QDialog):
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setWindowTitle("Preferences")
        self.setModal(True)
        self.setMinimumWidth(380)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # ── Fiscal Year ───────────────────────────────────────────────────
        fy_group = QGroupBox("Fiscal Year")
        fy_layout = QFormLayout(fy_group)

        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_spin.setValue(db_manager.get_current_year())
        fy_layout.addRow("Current Year:", self.year_spin)

        self.fy_month_combo = QComboBox()
        months = ["January", "February", "March", "April", "May", "June",
                  "July", "August", "September", "October", "November", "December"]
        self.fy_month_combo.addItems(months)
        fy_layout.addRow("Fiscal Year Start:", self.fy_month_combo)
        layout.addWidget(fy_group)

        # ── Company ───────────────────────────────────────────────────────
        co_group = QGroupBox("Company")
        co_layout = QFormLayout(co_group)

        self.company_input = QLineEdit()
        self.company_input.setPlaceholderText("e.g. ESQ Company")
        co_layout.addRow("Company Name:", self.company_input)

        self.tin_input = QLineEdit()
        self.tin_input.setPlaceholderText("e.g. 123-456-789-000")
        co_layout.addRow("Company TIN:", self.tin_input)

        layout.addWidget(co_group)

        # Load saved prefs
        self._load_prefs()

        # ── Buttons ───────────────────────────────────────────────────────
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._save_and_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    # ------------------------------------------------------------------ prefs I/O

    def _prefs_path(self) -> str:
        db_dir = os.path.dirname(self.db_manager.db_path)
        name   = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]
        return os.path.join(db_dir, f"{name}_prefs.json")

    def _load_prefs(self):
        import json
        try:
            with open(self._prefs_path(), 'r') as f:
                prefs = json.load(f)
            self.year_spin.setValue(prefs.get('year', self.db_manager.get_current_year()))
            month_idx = prefs.get('fy_start_month', 0)
            self.fy_month_combo.setCurrentIndex(month_idx)
            self.company_input.setText(prefs.get('company_name', ''))
            self.tin_input.setText(prefs.get('company_tin', ''))
        except Exception:
            pass

    def _save_and_accept(self):
        import json
        prefs = {
            'year':           self.year_spin.value(),
            'fy_start_month': self.fy_month_combo.currentIndex(),
            'company_name':   self.company_input.text().strip(),
            'company_tin':    self.tin_input.text().strip(),
        }
        try:
            with open(self._prefs_path(), 'w') as f:
                json.dump(prefs, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Warning", f"Could not save preferences:\n{e}")
        self.db_manager.set_current_year(prefs['year'])
        self.accept()

    def get_company_name(self) -> str:
        return self.company_input.text().strip()


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.setWindowTitle(self._make_title())
        self.setGeometry(100, 100, 1400, 800)

        # Record this file as recently opened
        _save_recent(db_manager.db_path)

        self._create_menu_bar()
        self._create_tabs()
        self._create_status_bar()
        self._setup_shortcuts()

        # Apply saved prefs (year, etc.) silently on load
        self._apply_saved_prefs()

    # ------------------------------------------------------------------ title

    def _make_title(self) -> str:
        name = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]
        return f"ESQ Accounting System — {name}"

    # ------------------------------------------------------------------ prefs

    def _prefs_path(self) -> str:
        db_dir = os.path.dirname(self.db_manager.db_path)
        name   = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]
        return os.path.join(db_dir, f"{name}_prefs.json")

    def _apply_saved_prefs(self):
        import json
        try:
            with open(self._prefs_path(), 'r') as f:
                prefs = json.load(f)
            year = prefs.get('year', datetime.now().year)
            self.db_manager.set_current_year(year)
        except Exception:
            pass

    # ------------------------------------------------------------------ menu bar

    def _create_menu_bar(self):
        menubar = self.menuBar()

        # ── File ──────────────────────────────────────────────────────────
        file_menu = menubar.addMenu("&File")

        # Load Account
        load_action = QAction("&Load Account…", self)
        load_action.setShortcut("Ctrl+O")
        load_action.setStatusTip("Open an existing account file")
        load_action.triggered.connect(self._load_account)
        file_menu.addAction(load_action)

        # Recent Accounts submenu
        self._recent_menu = QMenu("&Recent Accounts", self)
        file_menu.addMenu(self._recent_menu)
        self._rebuild_recent_menu()

        file_menu.addSeparator()

        # Close Account
        close_action = QAction("&Close Account", self)
        close_action.setShortcut("Ctrl+W")
        close_action.setStatusTip("Close this account and return to the startup screen")
        close_action.triggered.connect(self._close_account)
        file_menu.addAction(close_action)

        # Rename Account
        rename_action = QAction("Re&name Account…", self)
        rename_action.setStatusTip("Rename the current account file")
        rename_action.triggered.connect(self._rename_account)
        file_menu.addAction(rename_action)

        file_menu.addSeparator()

        # Backup
        backup_action = QAction("&Backup Account", self)
        backup_action.setShortcut("Ctrl+B")
        backup_action.setStatusTip("Save a timestamped backup of this account")
        backup_action.triggered.connect(self._backup_account)
        file_menu.addAction(backup_action)

        # Restore
        restore_action = QAction("Res&tore from Backup…", self)
        restore_action.setStatusTip("Replace this account with a backup copy")
        restore_action.triggered.connect(self._restore_from_backup)
        file_menu.addAction(restore_action)

        file_menu.addSeparator()

        # Import / Export Full Book
        file_menu.addSeparator()

        import_book_action = QAction("&Import Full Book…", self)
        import_book_action.setShortcut("Ctrl+Shift+I")
        import_book_action.setStatusTip(
            "Import COA, Alphalist, SJ, PJ, CDJ, CRJ from a single workbook"
        )
        import_book_action.triggered.connect(self._import_full_book)
        file_menu.addAction(import_book_action)

        export_book_action = QAction("E&xport Full Book…", self)
        export_book_action.setShortcut("Ctrl+Shift+X")
        export_book_action.setStatusTip(
            "Export all journals to a single multi-sheet workbook"
        )
        export_book_action.triggered.connect(self._export_full_book)
        file_menu.addAction(export_book_action)

        file_menu.addSeparator()

        # Refresh
        refresh_action = QAction("&Refresh All", self)
        refresh_action.setShortcut("F5")
        refresh_action.triggered.connect(self._refresh_all_tabs)
        file_menu.addAction(refresh_action)

        file_menu.addSeparator()

        # Exit
        exit_action = QAction("E&xit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # ── View ──────────────────────────────────────────────────────────
        view_menu = menubar.addMenu("&View")

        self._dark_mode_action = QAction("&Dark Mode", self)
        self._dark_mode_action.setShortcut("Ctrl+Shift+D")
        self._dark_mode_action.setCheckable(True)
        self._dark_mode_action.triggered.connect(self._toggle_dark_mode)
        view_menu.addAction(self._dark_mode_action)

        # ── Settings ──────────────────────────────────────────────────────
        settings_menu = menubar.addMenu("&Settings")

        prefs_action = QAction("&Preferences…", self)
        prefs_action.setShortcut("Ctrl+,")
        prefs_action.setStatusTip("Change year, company name, and other settings")
        prefs_action.triggered.connect(self._show_preferences)
        settings_menu.addAction(prefs_action)

        year_action = QAction("Change &Year…", self)
        year_action.setShortcut("Ctrl+Y")
        year_action.setStatusTip("Quickly switch the active fiscal year")
        year_action.triggered.connect(self._change_year)
        settings_menu.addAction(year_action)

        # ── Help ──────────────────────────────────────────────────────────
        help_menu = menubar.addMenu("&Help")

        manual_action = QAction("&User Manual", self)
        manual_action.setShortcut("F1")
        manual_action.triggered.connect(self._show_user_manual)
        help_menu.addAction(manual_action)

        help_menu.addSeparator()

        about_action = QAction("&About", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    # ------------------------------------------------------------------ recent menu

    def _rebuild_recent_menu(self):
        self._recent_menu.clear()
        items = _load_recent()
        current = self.db_manager.db_path if self.db_manager else ''

        # Exclude the currently open file from the list
        items = [p for p in items if p != current]

        if not items:
            empty = QAction("(No recent accounts)", self)
            empty.setEnabled(False)
            self._recent_menu.addAction(empty)
            return

        for path in items:
            name   = os.path.splitext(os.path.basename(path))[0]
            action = QAction(name, self)
            action.setStatusTip(path)
            action.setData(path)
            action.triggered.connect(lambda checked, p=path: self._open_recent(p))
            self._recent_menu.addAction(action)

        self._recent_menu.addSeparator()
        clear_action = QAction("Clear Recent List", self)
        clear_action.triggered.connect(self._clear_recent)
        self._recent_menu.addAction(clear_action)

    def _open_recent(self, path: str):
        if not os.path.exists(path):
            QMessageBox.warning(self, "File Not Found",
                                f"Could not find:\n{path}\n\nIt may have been moved or deleted.")
            self._rebuild_recent_menu()
            return
        self._switch_to_db(path)

    def _clear_recent(self):
        import json
        try:
            with open(_recent_file(), 'w') as f:
                json.dump([], f)
        except Exception:
            pass
        self._rebuild_recent_menu()

    # ------------------------------------------------------------------ File actions

    def _load_account(self):
        """Open a .db file from data/saves/ (or anywhere)."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Account",
            _saves_dir(),
            "Database Files (*.db);;All Files (*)"
        )
        if path:
            self._switch_to_db(path)

    def _switch_to_db(self, path: str):
        """Close the current db and reopen the window with a new one."""
        from database.db_manager import DatabaseManager
        self.db_manager.close()
        new_db = DatabaseManager(path)
        new_db.initialize_database()
        self.db_manager = new_db
        _save_recent(path)

        # Rebuild tabs with the new db
        old_widget = self.centralWidget()
        self._create_tabs()
        if old_widget:
            old_widget.deleteLater()

        self.setWindowTitle(self._make_title())
        self._apply_saved_prefs()
        self._rebuild_recent_menu()
        self.status_bar.showMessage(
            f"Loaded: {os.path.basename(path)}", 4000
        )

    def _close_account(self):
        """Close this account and return to the startup dialog."""
        reply = QMessageBox.question(
            self, "Close Account",
            "Close the current account and return to the startup screen?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.db_manager.close()

        # Re-show the startup dialog
        from ui.startup_dialog import StartupDialog
        from database.db_manager import DatabaseManager
        from resources.style_loader import load_stylesheet

        startup = StartupDialog()
        if startup.exec_() != StartupDialog.Accepted or not startup.db_path:
            # User cancelled startup — just exit
            QApplication.quit()
            return

        new_db = DatabaseManager(startup.db_path)
        new_db.initialize_database(
            coa_xlsx_path=startup.coa_xlsx_path,
            use_default_coa=(startup.coa_xlsx_path is None)
        )
        self.db_manager = new_db
        _save_recent(startup.db_path)

        old_widget = self.centralWidget()
        self._create_tabs()
        if old_widget:
            old_widget.deleteLater()

        self.setWindowTitle(self._make_title())
        self._apply_saved_prefs()
        self._rebuild_recent_menu()
        self.status_bar.showMessage("Account loaded.", 3000)

    def _rename_account(self):
        """Rename the current .db file (and its prefs sidecar if present)."""
        current_path = self.db_manager.db_path
        current_name = os.path.splitext(os.path.basename(current_path))[0]

        new_name, ok = QInputDialog.getText(
            self, "Rename Account",
            "Enter a new name for this account:",
            QLineEdit.Normal, current_name
        )
        if not ok or not new_name.strip():
            return

        new_name = new_name.strip()
        safe = "".join(c for c in new_name if c.isalnum() or c in (' ', '-', '_')).strip()
        if not safe:
            QMessageBox.warning(self, "Invalid Name", "Please enter a valid name.")
            return

        folder   = os.path.dirname(current_path)
        new_path = os.path.join(folder, safe + '.db')

        if os.path.exists(new_path) and new_path != current_path:
            QMessageBox.warning(self, "Name Taken",
                                f'"{safe}.db" already exists. Choose a different name.')
            return

        try:
            self.db_manager.close()
            os.rename(current_path, new_path)

            # Rename prefs sidecar too if it exists
            old_prefs = os.path.join(folder, f"{current_name}_prefs.json")
            new_prefs = os.path.join(folder, f"{safe}_prefs.json")
            if os.path.exists(old_prefs):
                os.rename(old_prefs, new_prefs)

        except OSError as e:
            QMessageBox.critical(self, "Rename Failed", str(e))
            return

        from database.db_manager import DatabaseManager
        self.db_manager = DatabaseManager(new_path)
        self.db_manager.get_connection()   # re-open connection
        self._apply_saved_prefs()

        _save_recent(new_path)
        self._rebuild_recent_menu()
        self.setWindowTitle(self._make_title())
        self.status_bar.showMessage(f'Renamed to "{safe}".', 3000)

    # ------------------------------------------------------------------ Backup / Restore

    def _backup_account(self):
        """Copy the current .db to data/backups/ with a timestamp."""
        src  = self.db_manager.db_path
        name = os.path.splitext(os.path.basename(src))[0]
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst  = os.path.join(_backups_dir(), f"{name}_backup_{ts}.db")

        try:
            # Flush any pending writes by closing and reopening
            self.db_manager.close()
            shutil.copy2(src, dst)
            self.db_manager.get_connection()   # reconnect
        except Exception as e:
            QMessageBox.critical(self, "Backup Failed", str(e))
            return

        self.status_bar.showMessage(f"Backup saved: {os.path.basename(dst)}", 5000)
        QMessageBox.information(
            self, "Backup Successful",
            f"Backup saved to:\n{dst}"
        )

    def _restore_from_backup(self):
        """Replace the current db with a chosen backup."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Restore from Backup",
            _backups_dir(),
            "Database Files (*.db);;All Files (*)"
        )
        if not path:
            return

        reply = QMessageBox.warning(
            self, "Confirm Restore",
            f"This will REPLACE the current account with:\n{os.path.basename(path)}\n\n"
            "All unsaved changes will be lost. Continue?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        dst = self.db_manager.db_path
        try:
            self.db_manager.close()
            shutil.copy2(path, dst)
        except Exception as e:
            QMessageBox.critical(self, "Restore Failed", str(e))
            return

        from database.db_manager import DatabaseManager
        self.db_manager = DatabaseManager(dst)
        self.db_manager.initialize_database()
        self._apply_saved_prefs()

        old_widget = self.centralWidget()
        self._create_tabs()
        if old_widget:
            old_widget.deleteLater()

        self.status_bar.showMessage("Account restored from backup.", 4000)
        QMessageBox.information(self, "Restore Successful",
                                "Account has been restored from the backup.")

    # ------------------------------------------------------------------ Settings

    def _show_preferences(self):
        dlg = PreferencesDialog(self.db_manager, self)
        if dlg.exec_():
            self.status_bar.showMessage("Preferences saved.", 3000)
            self._refresh_all_tabs()

    def _change_year(self):
        year, ok = QInputDialog.getInt(
            self, "Change Year",
            "Enter the fiscal year to work with:",
            self.db_manager.get_current_year(), 2000, 2100
        )
        if ok:
            self.db_manager.set_current_year(year)
            # Persist to prefs file
            import json
            try:
                path = self._prefs_path()
                try:
                    with open(path, 'r') as f:
                        prefs = json.load(f)
                except Exception:
                    prefs = {}
                prefs['year'] = year
                with open(path, 'w') as f:
                    json.dump(prefs, f, indent=2)
            except Exception:
                pass
            self.status_bar.showMessage(f"Year set to {year}.", 3000)
            self._refresh_all_tabs()

    # ------------------------------------------------------------------ Tabs

    def _create_tabs(self):
        self.tab_widget = QTabWidget()
        self.tab_widget.tabBar().setFocusPolicy(Qt.NoFocus)
        self.setCentralWidget(self.tab_widget)

        self._tab_key_filter = TabKeyFilter(self)
        QApplication.instance().installEventFilter(self._tab_key_filter)

        self.tab_widget.addTab(COAWidget(self.db_manager),                "Chart of Accounts")
        self.tab_widget.addTab(AlphalistWidget(self.db_manager),          "Alphalist")
        self.tab_widget.addTab(SalesJournalWidget(self.db_manager),       "Sales Journal")
        self.tab_widget.addTab(PurchaseJournalWidget(self.db_manager),    "Purchase Journal")
        self.tab_widget.addTab(CashDisbursementWidget(self.db_manager),   "Cash Disbursement")
        self.tab_widget.addTab(CashReceiptsWidget(self.db_manager),       "Cash Receipts")
        self.tab_widget.addTab(GeneralJournalWidget(self.db_manager),     "General Journal")
        self.tab_widget.addTab(GeneralLedgerWidget(self.db_manager),      "General Ledger")
        self.tab_widget.addTab(TrialBalanceWidget(self.db_manager),       "Trial Balance")
        self.tab_widget.addTab(FinancialStatementsWidget(self.db_manager),"Financial Statements")

    # ------------------------------------------------------------------ Shortcuts

    def _setup_shortcuts(self):
        for i in range(min(10, self.tab_widget.count())):
            key = str(i + 1) if i < 9 else '0'
            QShortcut(QKeySequence(f'Ctrl+{key}'), self).activated.connect(
                lambda idx=i: self.tab_widget.setCurrentIndex(idx)
            )
        QShortcut(QKeySequence('F5'), self).activated.connect(self._refresh_all_tabs)
        QShortcut(QKeySequence('F1'), self).activated.connect(self._show_user_manual)

    def _switch_tab_forward(self):
        n = self.tab_widget.count()
        self.tab_widget.setCurrentIndex((self.tab_widget.currentIndex() + 1) % n)

    def _switch_tab_backward(self):
        n = self.tab_widget.count()
        self.tab_widget.setCurrentIndex((self.tab_widget.currentIndex() - 1) % n)

    # ------------------------------------------------------------------ Status Bar

    def _create_status_bar(self):
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        db_name = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]
        self.status_bar.showMessage(f"Ready  |  {db_name}  |  Year: {self.db_manager.get_current_year()}")

    # ------------------------------------------------------------------ Helpers

    def _refresh_all_tabs(self):
        for i in range(self.tab_widget.count()):
            widget = self.tab_widget.widget(i)
            if hasattr(widget, 'load_data'):
                widget.load_data()
        year    = self.db_manager.get_current_year()
        db_name = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]
        self.status_bar.showMessage(
            f"All data refreshed  |  {db_name}  |  Year: {year}", 4000
        )

    # ------------------------------------------------------------------ Full Book Import / Export

    def _import_full_book(self):
        """Import COA + Alphalist + all journals from a single workbook."""
        from resources.file_paths import get_import_dir
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Full Book",
            get_import_dir(""),
            "Excel Files (*.xlsx)"
        )
        if not path:
            return

        reply = QMessageBox.question(
            self, "Import Full Book",
            f"Import all supported sheets from:\n{os.path.basename(path)}\n\n"
            "Existing records will NOT be deleted — duplicates are skipped.\n\n"
            "Continue?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply != QMessageBox.Yes:
            return

        # Progress dialog
        progress = QProgressDialog("Importing…", None, 0, 0, self)
        progress.setWindowTitle("Import Full Book")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        progress.show()
        QApplication.processEvents()

        try:
            from ui.fullbook_importer import import_full_book, build_summary_message

            def _on_progress(sheet_name: str):
                progress.setLabelText(f"Importing: {sheet_name}…")
                QApplication.processEvents()

            results = import_full_book(path, self.db_manager, _on_progress)
        except Exception as exc:
            progress.close()
            QMessageBox.critical(self, "Import Failed", str(exc))
            return

        progress.close()
        self._refresh_all_tabs()

        summary = build_summary_message(results)
        QMessageBox.information(self, "Import Complete", summary)

    def _export_full_book(self):
        """Export COA + Alphalist + all journals to a single multi-sheet workbook."""
        from resources.file_paths import get_io_dir
        from datetime import datetime as _dt

        db_name = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]
        year    = self.db_manager.get_current_year()
        default_name = f"{db_name}_{year}_books.xlsx"

        path, _ = QFileDialog.getSaveFileName(
            self, "Export Full Book",
            os.path.join(get_io_dir("Full Book"), default_name),
            "Excel Files (*.xlsx)"
        )
        if not path:
            return

        progress = QProgressDialog("Exporting…", None, 0, 0, self)
        progress.setWindowTitle("Export Full Book")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        progress.show()
        QApplication.processEvents()

        try:
            count, err = self._do_export_full_book(path, progress)
        except Exception as exc:
            progress.close()
            QMessageBox.critical(self, "Export Failed", str(exc))
            return

        progress.close()

        if err:
            QMessageBox.critical(self, "Export Failed", err)
        else:
            QMessageBox.information(
                self, "Export Complete",
                f"Full book exported to:\n{path}\n\n{count} total rows written."
            )

    def _do_export_full_book(self, path: str, progress) -> tuple:
        """Build the multi-sheet workbook and save it. Returns (total_rows, error_str)."""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from datetime import datetime as _dt
        except ImportError:
            return 0, "openpyxl is not installed.\nInstall with: pip install openpyxl"

        year    = self.db_manager.get_current_year()
        db_name = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]

        wb = Workbook()
        wb.remove(wb.active)   # remove default empty sheet

        # ── Styles ────────────────────────────────────────────────────────
        hdr_font  = Font(name='Arial', bold=True, color='FFFFFF', size=11)
        hdr_fill  = PatternFill('solid', start_color='2F5496')
        hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell_font = Font(name='Arial', size=10)
        alt_fill  = PatternFill('solid', start_color='DCE6F1')
        thin      = Side(style='thin', color='B0B0B0')
        border    = Border(left=thin, right=thin, top=thin, bottom=thin)
        title_font = Font(name='Arial', bold=True, size=14)
        sub_font   = Font(name='Arial', italic=True, size=11)
        t_align    = Alignment(horizontal='left', vertical='center')

        def _write_sheet(ws, title, headers, rows_data, col_widths):
            """Write a formatted sheet: title rows + header + data."""
            ws.merge_cells(f'A2:{get_column_letter(len(headers))}2')
            ws['A2'].value = title.upper()
            ws['A2'].font  = title_font
            ws['A2'].alignment = t_align
            ws.row_dimensions[2].height = 22

            ws.merge_cells(f'A3:{get_column_letter(len(headers))}3')
            ws['A3'].value = f'For the Year {year}'
            ws['A3'].font  = sub_font
            ws['A3'].alignment = t_align

            HR = 5
            ws.row_dimensions[HR].height = 28
            for ci, hdr in enumerate(headers, 1):
                c = ws.cell(row=HR, column=ci, value=hdr)
                c.font = hdr_font; c.fill = hdr_fill
                c.alignment = hdr_align; c.border = border

            for ri, row in enumerate(rows_data):
                rn   = 6 + ri
                fill = alt_fill if ri % 2 == 0 else None
                ws.row_dimensions[rn].height = 18
                for ci, val in enumerate(row, 1):
                    c = ws.cell(row=rn, column=ci, value=val)
                    c.font = cell_font; c.border = border
                    c.alignment = Alignment(
                        horizontal='right' if isinstance(val, (int, float)) else 'left',
                        vertical='center'
                    )
                    if fill:
                        c.fill = fill

            for ci, w in col_widths.items():
                ws.column_dimensions[get_column_letter(ci)].width = w

            last_col = get_column_letter(len(headers))
            ws.freeze_panes = 'A6'
            ws.auto_filter.ref = f'A{HR}:{last_col}{HR}'
            return len(rows_data)

        total_rows = 0

        # ── COA ───────────────────────────────────────────────────────────
        progress.setLabelText("Exporting: Chart of Accounts…")
        QApplication.processEvents()
        ws = wb.create_sheet("COA")
        accounts = self.db_manager.get_all_accounts()
        rows = [(a['account_code'], a['account_description'],
                 a.get('normal_balance', 'Debit')) for a in accounts]
        total_rows += _write_sheet(
            ws, "Chart of Accounts",
            ['Account Code', 'Account Description', 'DEBIT/CREDIT'],
            rows, {1: 18, 2: 45, 3: 16}
        )

        # ── Alphalist ─────────────────────────────────────────────────────
        progress.setLabelText("Exporting: Alphalist…")
        QApplication.processEvents()
        ws = wb.create_sheet("Alphalist")
        entries = self.db_manager.get_all_alphalist()
        rows = [(
            e.get('tin', ''), e.get('entry_type', ''), e.get('company_name', ''),
            e.get('first_name', ''), e.get('middle_name', ''), e.get('last_name', ''),
            e.get('address1', ''), e.get('address2', ''), e.get('vat', '')
        ) for e in entries]
        total_rows += _write_sheet(
            ws, "Alphalist",
            ['TIN', 'Entry Type', 'Company Name', 'First Name',
             'Middle Name', 'Last Name', 'Address 1', 'Address 2', 'VAT Type'],
            rows, {1: 16, 2: 16, 3: 30, 4: 18, 5: 18, 6: 18, 7: 28, 8: 28, 9: 16}
        )

        # ── Sales Journal ─────────────────────────────────────────────────
        progress.setLabelText("Exporting: Sales Journal…")
        QApplication.processEvents()
        ws = wb.create_sheet(f"SJ_{str(year)[2:]}")
        sj = self.db_manager.get_sales_journal()
        rows = [(
            e.get('date', ''), e.get('customer_name', ''), e.get('reference_no', ''),
            e.get('tin', ''), e.get('net_amount', 0), e.get('output_vat', 0),
            e.get('gross_amount', 0), e.get('goods', 0), e.get('services', 0),
            e.get('particulars', '')
        ) for e in sj]
        total_rows += _write_sheet(
            ws, "Sales Journal",
            ['Date', 'Customer Name', 'Reference No', 'TIN',
             'Net Amount', 'Output VAT', 'Gross Amount', 'Goods', 'Services', 'Particulars'],
            rows, {1: 14, 2: 30, 3: 16, 4: 16, 5: 14, 6: 14, 7: 14, 8: 14, 9: 14, 10: 30}
        )

        # ── Purchase Journal ──────────────────────────────────────────────
        progress.setLabelText("Exporting: Purchase Journal…")
        QApplication.processEvents()
        ws = wb.create_sheet(f"PJ_{str(year)[2:]}")
        pj = self.db_manager.get_purchase_journal()
        rows = [(
            e.get('date', ''), e.get('payee_name', ''), e.get('reference_no', ''),
            e.get('tin', ''), e.get('branch_code', ''), e.get('net_amount', 0),
            e.get('input_vat', 0), e.get('gross_amount', 0),
            e.get('account_code', ''), e.get('account_description', ''),
            e.get('debit', 0), e.get('credit', 0), e.get('particulars', '')
        ) for e in pj]
        total_rows += _write_sheet(
            ws, "Purchase Journal",
            ['Date', 'Payee Name', 'Reference No', 'TIN', 'Branch Code',
             'Net Amount', 'Input VAT', 'Gross Amount',
             'Account Code', 'Account Description', 'Debit', 'Credit', 'Particulars'],
            rows, {1: 14, 2: 28, 3: 16, 4: 16, 5: 12, 6: 14, 7: 14, 8: 14,
                   9: 14, 10: 28, 11: 14, 12: 14, 13: 30}
        )

        # ── Cash Disbursement Journal ─────────────────────────────────────
        progress.setLabelText("Exporting: Cash Disbursement Journal…")
        QApplication.processEvents()
        ws = wb.create_sheet(f"CDJ_{str(year)[2:]}")
        cdj = self.db_manager.get_cash_disbursement_journal()
        rows = [(
            e.get('date', ''), e.get('reference_no', ''), e.get('particulars', ''),
            e.get('account_code', ''), e.get('account_description', ''),
            e.get('debit', 0), e.get('credit', 0)
        ) for e in cdj]
        total_rows += _write_sheet(
            ws, "Cash Disbursement Journal",
            ['Date', 'Reference No', 'Particulars',
             'Account Code', 'Account Description', 'Debit', 'Credit'],
            rows, {1: 14, 2: 16, 3: 36, 4: 14, 5: 30, 6: 14, 7: 14}
        )

        # ── Cash Receipts Journal ─────────────────────────────────────────
        progress.setLabelText("Exporting: Cash Receipts Journal…")
        QApplication.processEvents()
        ws = wb.create_sheet(f"CRJ_{str(year)[2:]}")
        crj = self.db_manager.get_cash_receipts_journal()
        rows = [(
            e.get('date', ''), e.get('reference_no', ''), e.get('particulars', ''),
            e.get('account_code', ''), e.get('account_description', ''),
            e.get('debit', 0), e.get('credit', 0)
        ) for e in crj]
        total_rows += _write_sheet(
            ws, "Cash Receipts Journal",
            ['Date', 'Reference No', 'Particulars',
             'Account Code', 'Account Description', 'Debit', 'Credit'],
            rows, {1: 14, 2: 16, 3: 36, 4: 14, 5: 30, 6: 14, 7: 14}
        )

        progress.setLabelText("Saving file…")
        QApplication.processEvents()
        wb.save(path)
        return total_rows, ''

    def _show_user_manual(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("ESQ Accounting System — User Manual")
        dlg.resize(820, 680)
        dlg.setMinimumSize(640, 480)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(0, 0, 0, 8)

        header = QLabel(" 📖  ESQ Accounting System — User Manual")
        header.setStyleSheet(
            "background:#1a3a5c; color:#ffffff; font-size:16px; "
            "font-weight:bold; padding:10px 14px;"
        )
        layout.addWidget(header)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(False)
        browser.setFont(QFont("Segoe UI", 10))
        browser.setStyleSheet("background:#fafafa; border:none; padding:4px;")

        manual_path = get_resource('user_manual.html')
        try:
            with open(manual_path, 'r', encoding='utf-8') as f:
                html = f.read()
        except FileNotFoundError:
            html = "<p><b>User manual not found.</b><br>Expected at: " + manual_path + "</p>"
        browser.setHtml(html)
        layout.addWidget(browser)

        btn_box = QDialogButtonBox(QDialogButtonBox.Close)
        btn_box.rejected.connect(dlg.accept)
        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(8, 0, 8, 0)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_box)
        layout.addLayout(btn_layout)

        dlg.exec_()

    def _show_about(self):
        QMessageBox.about(
            self, "About ESQ Company",
            "<h2>ESQ Company Accounting System</h2>"
            "<p>Version 1.0</p>"
            "<p>A complete accounting journal system built with Python and PyQt5.</p>"
            "<p>© 2025 ESQ Company</p>"
        )

    def _toggle_dark_mode(self, checked: bool):
        from resources.style_loader import load_stylesheet
        QApplication.instance().setStyleSheet(
            load_stylesheet(get_resource('style.qss'), dark=checked)
        )
        self.status_bar.showMessage(
            "Dark mode enabled" if checked else "Light mode enabled", 2000
        )

    def _prefs_path(self) -> str:
        db_dir = os.path.dirname(self.db_manager.db_path)
        name   = os.path.splitext(os.path.basename(self.db_manager.db_path))[0]
        return os.path.join(db_dir, f"{name}_prefs.json")