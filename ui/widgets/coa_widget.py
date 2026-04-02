import os
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableWidget,
                             QTableWidgetItem, QPushButton, QLabel, QLineEdit,
                             QDialog, QFormLayout, QDialogButtonBox, QMessageBox,
                             QHeaderView, QGroupBox, QComboBox, QFileDialog)
from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence, QShortcut
from database.db_manager import DatabaseManager
from resources.file_paths import get_io_dir, get_import_dir
from ui.utils.search_utils import SearchFilter

class COADialog(QDialog):
    def __init__(self, parent=None, account_data=None):
        super().__init__(parent)
        self.account_data = account_data
        is_edit = account_data is not None
        self.setWindowTitle("Add Account" if not is_edit else "Edit Account")
        self.setModal(True)
        self.resize(450, 220)
        layout = QFormLayout()
        self.code_input = QLineEdit()
        if account_data:
            self.code_input.setText(account_data['account_code'])
            self.code_input.setReadOnly(True)
            self.code_input.setStyleSheet(
                "background: #f0f0f0; color: #666; border: 1px solid #ccc;")
        layout.addRow("Account Code:", self.code_input)
        self.desc_input = QLineEdit()
        if account_data:
            self.desc_input.setText(account_data['account_description'])
        layout.addRow("Account Description:", self.desc_input)
        self.normal_balance_combo = QComboBox()
        self.normal_balance_combo.addItems(["Debit", "Credit"])
        if account_data and account_data.get('normal_balance'):
            idx = self.normal_balance_combo.findText(account_data['normal_balance'])
            if idx >= 0:
                self.normal_balance_combo.setCurrentIndex(idx)
        layout.addRow("Normal Balance:", self.normal_balance_combo)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        self.setLayout(layout)

    def get_data(self):
        return {
            'account_code':        self.code_input.text().strip(),
            'account_description': self.desc_input.text().strip(),
            'normal_balance':      self.normal_balance_combo.currentText(),
        }


class COAWidget(QWidget):
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.all_accounts = []
        self._setup_ui()
        self._setup_shortcuts()
        self.load_data()

    def _setup_ui(self):
        layout = QVBoxLayout()

        title = QLabel("CHART OF ACCOUNTS")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        search_group = QGroupBox("Search && Filter")
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Type to search by code or description...")
        self.search_input.setClearButtonEnabled(True)
        search_layout.addWidget(self.search_input)
        self.clear_search_btn = QPushButton("Clear")
        self.clear_search_btn.clicked.connect(self._clear_search)
        search_layout.addWidget(self.clear_search_btn)
        self.results_label = QLabel("Showing: 0 of 0")
        search_layout.addWidget(self.results_label)
        search_group.setLayout(search_layout)
        layout.addWidget(search_group)

        button_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Account")
        self.add_btn.clicked.connect(self._add_account)
        button_layout.addWidget(self.add_btn)
        self.edit_btn = QPushButton("Edit Account")
        self.edit_btn.clicked.connect(self._edit_account)
        button_layout.addWidget(self.edit_btn)
        self.import_btn = QPushButton("Import COA")
        self.import_btn.clicked.connect(self._import_coa)
        button_layout.addWidget(self.import_btn)
        self.export_btn = QPushButton("Export COA")
        self.export_btn.clicked.connect(self._export_coa)
        button_layout.addWidget(self.export_btn)
        button_layout.addStretch()
        layout.addLayout(button_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setHorizontalHeaderLabels(
            ["Account Code", "Account Description", "Normal Balance"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table)
        self.setLayout(layout)

        # COA has no date column — SearchFilter used for text only
        self._search = SearchFilter(
            table         = self.table,
            search_input  = self.search_input,
            results_label = self.results_label,
        )

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+N"), self).activated.connect(self._add_account)
        QShortcut(QKeySequence("Ctrl+E"), self).activated.connect(self._edit_account)
        QShortcut(QKeySequence("Ctrl+F"), self).activated.connect(self.search_input.setFocus)
        QShortcut(QKeySequence("Ctrl+I"), self).activated.connect(self._import_coa)
        QShortcut(QKeySequence("Ctrl+Shift+E"), self).activated.connect(self._export_coa)

    def load_data(self):
        self.all_accounts = self.db_manager.get_all_accounts()
        self._populate_table(self.all_accounts)
        self._search.refresh()

    def _populate_table(self, accounts):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(accounts))
        for row, account in enumerate(accounts):
            code_item = QTableWidgetItem(account['account_code'])
            desc_item = QTableWidgetItem(account['account_description'])
            nb_item   = QTableWidgetItem(account.get('normal_balance', 'Debit'))
            code_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            desc_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            nb_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            nb_item.setTextAlignment(Qt.AlignCenter)
            code_item.setData(Qt.UserRole, account['id'])
            self.table.setItem(row, 0, code_item)
            self.table.setItem(row, 1, desc_item)
            self.table.setItem(row, 2, nb_item)
        self.table.setSortingEnabled(True)

    def _clear_search(self):
        self.search_input.clear()

    def _add_account(self):
        dialog = COADialog(self)
        if dialog.exec():
            data = dialog.get_data()
            if data['account_code'] and data['account_description']:
                if self.db_manager.add_account(data):
                    self.load_data()
                    self.search_input.clear()
                    QMessageBox.information(self, "Success", "Account added successfully!")
                else:
                    QMessageBox.warning(self, "Error", "Account code already exists!")

    def _edit_account(self):
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "Warning", "Please select an account to edit")
            return
        account_id = self.table.item(current_row, 0).data(Qt.UserRole)
        nb_item    = self.table.item(current_row, 2)
        account_data = {
            'id':                  account_id,
            'account_code':        self.table.item(current_row, 0).text(),
            'account_description': self.table.item(current_row, 1).text(),
            'normal_balance':      nb_item.text() if nb_item else 'Debit',
        }
        dialog = COADialog(self, account_data)
        if dialog.exec():
            data = dialog.get_data()
            if data['account_description']:
                if self.db_manager.update_account(account_id, data):
                    self.load_data()
                    QMessageBox.information(self, "Success", "Account updated successfully!")
                else:
                    QMessageBox.warning(self, "Error", "Failed to update account!")

    def _import_coa(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Chart of Accounts", get_import_dir(""),
            "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        if self.all_accounts:
            reply = QMessageBox.question(
                self, "Confirm Import",
                f"The current COA has {len(self.all_accounts)} account(s).\n\n"
                "Importing will ADD new accounts (existing codes are skipped).\nContinue?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply != QMessageBox.Yes:
                return
        imported, errors = self.db_manager.import_coa_from_xlsx(path)
        self.load_data()
        self.search_input.clear()
        msg = f"Import complete.\n  Imported: {imported} account(s)"
        if errors:
            msg += f"\n  Skipped / errors: {len(errors)}"
            if len(errors) <= 10:
                msg += "\n\nDetails:\n" + "\n".join(errors)
        QMessageBox.information(self, "Import Summary", msg)

    def _export_coa(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Chart of Accounts",
            os.path.join(get_io_dir("COA"), "chart_of_accounts.xlsx"),
            "Excel Files (*.xlsx)")
        if not path:
            return
        count, err = self.db_manager.export_coa_to_xlsx(path)
        if err:
            QMessageBox.critical(self, "Export Failed", err)
        else:
            QMessageBox.information(self, "Export Successful",
                                    f"{count} account(s) exported to:\n{path}")