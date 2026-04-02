"""
ui/dialogs/line_dialog.py
-------------------------
Shared journal-line add/edit dialog used by:
    - CashDisbursementWidget
    - CashReceiptsWidget
    - GeneralJournalWidget
    - SalesJournalWidget       (via subclass SJLineDialog)
    - PurchaseJournalWidget    (via subclass PJLineDialog)

The base LineDialog covers the CDJ / CRJ / GJ case exactly.
SJ and PJ import it directly — their _LineDialog was already identical
to the base so no subclassing is needed there either.

Usage
-----
    from ui.dialogs.line_dialog import LineDialog

    # open to add a new line
    dlg = LineDialog(self.db_manager, parent=self)
    if dlg.exec_():
        self.lines.append(dlg.get_data())

    # open to edit an existing line
    dlg = LineDialog(self.db_manager, parent=self, line_data=self.lines[row])
    if dlg.exec_():
        self.lines[row] = dlg.get_data()

Migration note (PySide6)
------------------------
Change the two PyQt5 import lines at the top to PySide6 equivalents,
and rename exec_() → exec() at all call sites.
"""

from __future__ import annotations

from PySide6.QtWidgets import (   # ← swap to PySide6.QtWidgets when migrating
    QDialog, QFormLayout, QDialogButtonBox,
    QComboBox, QLineEdit, QDoubleSpinBox, QMessageBox,
)
from PySide6.QtCore import Qt     # ← swap to PySide6.QtCore when migrating


class LineDialog(QDialog):
    """
    Single-line dialog for adding or editing one journal line
    (Account Description, Account Code, Debit, Credit).

    Enforces mutual exclusion between Debit and Credit:
    entering a non-zero value in one field automatically clears the other.

    Parameters
    ----------
    db_manager  : DatabaseManager — used to populate the account combo
    parent      : parent QWidget
    line_data   : dict with keys 'account_description', 'account_code',
                  'debit', 'credit' — pre-fills the form when editing
    title       : optional window title override
    """

    def __init__(
        self,
        db_manager,
        parent=None,
        line_data: dict | None = None,
        title: str = 'Journal Line',
    ):
        super().__init__(parent)
        self.db_manager  = db_manager
        self.account_map: dict[str, str] = {}   # description → code

        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(480)

        layout = QFormLayout()
        layout.setLabelAlignment(Qt.AlignRight)

        # ── Account combo ─────────────────────────────────────────────────
        self.account_combo = QComboBox()
        self.account_combo.setEditable(True)
        self._load_accounts()
        self.account_combo.currentTextChanged.connect(self._on_account_changed)
        layout.addRow('Account Description:', self.account_combo)

        # ── Account code (read-only, auto-filled) ─────────────────────────
        self.account_code_input = QLineEdit()
        self.account_code_input.setReadOnly(True)
        layout.addRow('Account Code:', self.account_code_input)

        # ── Debit ─────────────────────────────────────────────────────────
        self.debit_input = QDoubleSpinBox()
        self.debit_input.setMaximum(99_999_999.99)
        self.debit_input.setDecimals(2)
        self.debit_input.setGroupSeparatorShown(True)
        self.debit_input.valueChanged.connect(self._debit_changed)
        layout.addRow('Debit:', self.debit_input)

        # ── Credit ────────────────────────────────────────────────────────
        self.credit_input = QDoubleSpinBox()
        self.credit_input.setMaximum(99_999_999.99)
        self.credit_input.setDecimals(2)
        self.credit_input.setGroupSeparatorShown(True)
        self.credit_input.valueChanged.connect(self._credit_changed)
        layout.addRow('Credit:', self.credit_input)

        # ── Pre-fill when editing ─────────────────────────────────────────
        if line_data:
            self.account_combo.setCurrentText(
                line_data.get('account_description', ''))
            self.debit_input.setValue(line_data.get('debit', 0))
            self.credit_input.setValue(line_data.get('credit', 0))

        # ── Buttons ───────────────────────────────────────────────────────
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

        self.setLayout(layout)

        # Trigger code auto-fill for the initial selection
        self._on_account_changed(self.account_combo.currentText())

    # ------------------------------------------------------------------
    # Private
    # ------------------------------------------------------------------

    def _load_accounts(self):
        self.account_combo.addItem('')
        for acct in self.db_manager.get_all_accounts():
            desc = acct['account_description']
            self.account_combo.addItem(desc)
            self.account_map[desc] = acct['account_code']

    def _on_account_changed(self, text: str):
        self.account_code_input.setText(self.account_map.get(text, ''))

    def _debit_changed(self, value: float):
        if value > 0:
            self.credit_input.blockSignals(True)
            self.credit_input.setValue(0)
            self.credit_input.blockSignals(False)

    def _credit_changed(self, value: float):
        if value > 0:
            self.debit_input.blockSignals(True)
            self.debit_input.setValue(0)
            self.debit_input.blockSignals(False)

    def _validate_and_accept(self):
        if not self.account_combo.currentText().strip():
            QMessageBox.warning(self, 'Validation', 'Please select an account.')
            return
        if self.debit_input.value() == 0 and self.credit_input.value() == 0:
            QMessageBox.warning(
                self, 'Validation', 'Enter a debit or credit amount.')
            return
        self.accept()

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def get_data(self) -> dict:
        """Return the entered line as a dict ready to append to a lines list."""
        return {
            'account_description': self.account_combo.currentText().strip(),
            'account_code':        self.account_code_input.text().strip(),
            'debit':               self.debit_input.value(),
            'credit':              self.credit_input.value(),
        }