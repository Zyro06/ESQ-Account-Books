"""
startup_dialog.py
-----------------
First window shown on launch.  Lets the user either:
  • Create a new named .db file  (optionally importing a COA from .xlsx)
  • Load an existing .db file from data/saves/
  • Delete an existing .db file
"""

import os
import sys

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QMessageBox, QInputDialog,
    QLineEdit, QFileDialog, QFrame, QSizePolicy
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


def _saves_dir() -> str:
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(base, 'data', 'saves')
    os.makedirs(path, exist_ok=True)
    return path


class StartupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db_path: str | None = None
        self.coa_xlsx_path: str | None = None
        self._use_default_coa = True

        self.setWindowTitle("ESQ Accounting System")
        self.setMinimumWidth(520)
        self.setModal(True)
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(16)
        root.setContentsMargins(24, 24, 24, 24)

        title = QLabel("ESQ Company Accounting System")
        title.setAlignment(Qt.AlignCenter)
        f = QFont()
        f.setPointSize(16)
        f.setBold(True)
        title.setFont(f)
        root.addWidget(title)

        sub = QLabel("Select an option to continue")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet("color: #666;")
        root.addWidget(sub)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        root.addWidget(line)

        # ── Top buttons ───────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        self.create_btn = QPushButton("➕  Create New Account")
        self.create_btn.setFixedHeight(44)
        self.create_btn.setFont(QFont("Segoe UI", 11))
        self.create_btn.clicked.connect(self._on_create)
        btn_row.addWidget(self.create_btn)

        self.load_btn = QPushButton("📂  Load Existing Account")
        self.load_btn.setFixedHeight(44)
        self.load_btn.setFont(QFont("Segoe UI", 11))
        self.load_btn.clicked.connect(self._on_load_selected)
        btn_row.addWidget(self.load_btn)

        root.addLayout(btn_row)

        # ── Saved accounts list ───────────────────────────────────────────
        list_label = QLabel("Saved Accounts:")
        list_label.setFont(QFont("Segoe UI", 10, QFont.Bold))
        root.addWidget(list_label)

        self.saves_list = QListWidget()
        self.saves_list.setAlternatingRowColors(True)
        self.saves_list.setMinimumHeight(180)
        self.saves_list.itemDoubleClicked.connect(self._on_load_selected)
        root.addWidget(self.saves_list)

        self._refresh_saves_list()

        # ── Bottom row: Delete on left, Exit on right ─────────────────────
        bottom = QHBoxLayout()

        self.delete_btn = QPushButton("🗑  Delete Account")
        self.delete_btn.setFixedHeight(34)
        self.delete_btn.setStyleSheet("color: #c0392b;")
        self.delete_btn.setToolTip("Permanently delete the selected account file")
        self.delete_btn.clicked.connect(self._on_delete)
        bottom.addWidget(self.delete_btn)

        bottom.addStretch()

        exit_btn = QPushButton("Exit")
        exit_btn.clicked.connect(self.reject)
        bottom.addWidget(exit_btn)

        root.addLayout(bottom)

    # ------------------------------------------------------------------ helpers

    def _refresh_saves_list(self):
        self.saves_list.clear()
        saves = _saves_dir()
        files = sorted(f for f in os.listdir(saves) if f.endswith('.db'))
        if files:
            for fname in files:
                item = QListWidgetItem(fname.replace('.db', ''))
                item.setData(Qt.UserRole, os.path.join(saves, fname))
                self.saves_list.addItem(item)
        else:
            placeholder = QListWidgetItem("No saved accounts found")
            placeholder.setFlags(Qt.NoItemFlags)
            placeholder.setForeground(Qt.gray)
            self.saves_list.addItem(placeholder)

    # ------------------------------------------------------------------ slots

    def _on_create(self):
        name, ok = QInputDialog.getText(
            self, "New Account", "Enter a name for the new account file:",
            QLineEdit.Normal, "")
        if not ok or not name.strip():
            return
        name = name.strip()
        safe = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).strip()
        if not safe:
            QMessageBox.warning(self, "Invalid Name", "Please enter a valid account name.")
            return
        db_file = os.path.join(_saves_dir(), safe + '.db')
        if os.path.exists(db_file):
            reply = QMessageBox.question(
                self, "File Exists",
                f'"{safe}.db" already exists. Overwrite?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        self.db_path = db_file
        self._ask_coa_source()

    def _ask_coa_source(self):
        dlg = _COASourceDialog(self)
        result = dlg.exec_()

        if result == _COASourceDialog.USE_DEFAULT:
            self.coa_xlsx_path = None
            self._use_default_coa = True
            self.accept()

        elif result == _COASourceDialog.IMPORT_FILE:
            path, _ = QFileDialog.getOpenFileName(
                self, "Import Chart of Accounts",
                "", "Excel Files (*.xlsx *.xls)")
            if not path:
                self.db_path = None
                return
            self.coa_xlsx_path = path
            self._use_default_coa = False
            self.accept()

        elif result == _COASourceDialog.START_EMPTY:
            self.coa_xlsx_path = None
            self._use_default_coa = False
            self.accept()

    def _on_load_selected(self):
        item = self.saves_list.currentItem()
        if item is None or not item.data(Qt.UserRole):
            QMessageBox.information(self, "No Selection",
                                    "Please select an account to load.")
            return
        self.db_path = item.data(Qt.UserRole)
        self.coa_xlsx_path = None
        self._use_default_coa = True
        self.accept()

    def _on_delete(self):
        """Permanently delete the selected .db file from disk."""
        item = self.saves_list.currentItem()
        if item is None or not item.data(Qt.UserRole):
            QMessageBox.information(self, "No Selection",
                                    "Please select an account to delete.")
            return

        db_path = item.data(Qt.UserRole)
        name    = item.text()

        reply = QMessageBox.warning(
            self, "Confirm Delete",
            f'Permanently delete "{name}"?\n\n'
            f'This cannot be undone. All data in this account will be lost.',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        try:
            os.remove(db_path)

            # Also remove the prefs sidecar if it exists
            prefs_path = os.path.splitext(db_path)[0] + '_prefs.json'
            if os.path.exists(prefs_path):
                os.remove(prefs_path)

        except OSError as e:
            QMessageBox.critical(self, "Delete Failed",
                                 f"Could not delete the file:\n{e}")
            return

        self._refresh_saves_list()


# ---------------------------------------------------------------------------

class _COASourceDialog(QDialog):
    """Choose between default COA, import from xlsx, or start empty."""

    USE_DEFAULT  = 1
    IMPORT_FILE  = 2
    START_EMPTY  = 3

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chart of Accounts")
        self.setModal(True)
        self.setMinimumWidth(380)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(20, 20, 20, 20)

        lbl = QLabel("How would you like to set up the\nChart of Accounts?")
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setFont(QFont("Segoe UI", 11))
        layout.addWidget(lbl)

        default_btn = QPushButton("Use Default Chart of Accounts")
        default_btn.setFixedHeight(40)
        default_btn.clicked.connect(self._use_default)
        layout.addWidget(default_btn)

        import_btn = QPushButton("Import from Excel File (.xlsx)")
        import_btn.setFixedHeight(40)
        import_btn.clicked.connect(self._import_file)
        layout.addWidget(import_btn)

        empty_btn = QPushButton("Start with Empty Chart of Accounts")
        empty_btn.setFixedHeight(40)
        empty_btn.clicked.connect(self._start_empty)
        layout.addWidget(empty_btn)

        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        layout.addWidget(cancel_btn)

    def _use_default(self):
        self.done(self.USE_DEFAULT)

    def _import_file(self):
        self.done(self.IMPORT_FILE)

    def _start_empty(self):
        self.done(self.START_EMPTY)