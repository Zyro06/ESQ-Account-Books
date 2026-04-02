"""
ui/widgets/cash_disbursement_widget.py
--------------------------------------
Cash Disbursement Journal widget.

All logic lives in BaseJournalWidget.
This file only declares the widget-specific constants and the entry
dialog that opens when adding / editing an entry.
"""

from __future__ import annotations

from PySide6.QtWidgets import (   # ← swap to PySide6.QtWidgets when migrating
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QDialogButtonBox,
    QLabel, QLineEdit, QDateEdit, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox, QFrame,
)
from PySide6.QtCore import Qt, QDate          # ← swap to PySide6.QtCore
from PySide6.QtGui  import QKeySequence, QFont  # ← swap to PySide6.QtGui

from database.db_manager           import DatabaseManager
from ui.base.base_journal_widget   import BaseJournalWidget
from ui.dialogs.line_dialog        import LineDialog
from ui.utils.search_utils import SearchFilter, add_month_combo


# ---------------------------------------------------------------------------
# XLS column layout  (shared across export + import)
# ---------------------------------------------------------------------------

_XLS_COLUMNS = [
    ('Date',                'date'),
    ('Reference No',        'reference_no'),
    ('Particulars',         'particulars'),
    ('Account Description', 'account_description'),
    ('Account Code',        'account_code'),
    ('Debit',               'debit'),
    ('Credit',              'credit'),
]


# ---------------------------------------------------------------------------
# Entry dialog
# ---------------------------------------------------------------------------

class CashDisbursementDialog(QDialog):
    """Add / edit / copy a Cash Disbursement journal entry."""

    def __init__(self, db_manager, parent=None, entry_data=None, is_copy=False):
        super().__init__(parent)
        self.db_manager = db_manager
        self.lines: list = []

        if is_copy:
            self.setWindowTitle('Copy Entry — Cash Disbursement Journal')
        elif entry_data is None:
            self.setWindowTitle('New Entry — Cash Disbursement Journal')
        else:
            self.setWindowTitle('Edit Entry — Cash Disbursement Journal')

        self.setModal(True)
        self.setMinimumWidth(820)
        self.setMinimumHeight(560)

        root = QVBoxLayout()
        root.setSpacing(10)

        # ── Header fields ─────────────────────────────────────────────
        hdr = QFormLayout()
        hdr.setLabelAlignment(Qt.AlignRight)

        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDisplayFormat('MM/dd/yyyy')
        self.date_input.setDate(
            QDate.fromString(entry_data['date'], 'MM/dd/yyyy')
            if entry_data else QDate.currentDate())
        hdr.addRow('Date:', self.date_input)

        self.reference_input = QLineEdit()
        if entry_data:
            self.reference_input.setText(entry_data.get('reference_no', ''))
        hdr.addRow('Reference No.:', self.reference_input)

        self.particulars_input = QLineEdit()
        if entry_data:
            self.particulars_input.setText(entry_data.get('particulars', '') or '')
        hdr.addRow('Particulars:', self.particulars_input)

        root.addLayout(hdr)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        root.addWidget(sep)

        lbl = QLabel('Journal Lines:')
        lbl.setStyleSheet('font-weight: bold;')
        root.addWidget(lbl)

        # ── Lines table ───────────────────────────────────────────────
        self.lines_table = QTableWidget()
        self.lines_table.setColumnCount(4)
        self.lines_table.setHorizontalHeaderLabels(
            ['Account Description', 'Account Code', 'Debit', 'Credit'])
        self.lines_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.lines_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.lines_table.setSelectionMode(QTableWidget.SingleSelection)
        self.lines_table.setAlternatingRowColors(True)
        self.lines_table.verticalHeader().setVisible(False)
        lhdr = self.lines_table.horizontalHeader()
        lhdr.setSectionResizeMode(0, QHeaderView.Stretch)
        lhdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        lhdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        lhdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        root.addWidget(self.lines_table)

        # ── Line buttons ──────────────────────────────────────────────
        lbtn = QHBoxLayout()
        add_btn    = QPushButton('Add Line');    add_btn.setShortcut(QKeySequence('Ctrl+N'))
        edit_btn   = QPushButton('Edit Line');   edit_btn.setShortcut(QKeySequence('Ctrl+E'))
        remove_btn = QPushButton('Remove Line'); remove_btn.setShortcut(QKeySequence('Ctrl+D'))
        add_btn.clicked.connect(self._add_line)
        edit_btn.clicked.connect(self._edit_line)
        remove_btn.clicked.connect(self._remove_line)
        lbtn.addWidget(add_btn); lbtn.addWidget(edit_btn)
        lbtn.addWidget(remove_btn); lbtn.addStretch()
        root.addLayout(lbtn)

        # ── Totals row ────────────────────────────────────────────────
        bf = QFont(); bf.setBold(True)
        self.total_debit_lbl  = QLabel('Total Debit:  0.00')
        self.total_credit_lbl = QLabel('Total Credit: 0.00')
        self.balance_lbl      = QLabel('')
        self.total_debit_lbl.setFont(bf)
        self.total_credit_lbl.setFont(bf)
        tot = QHBoxLayout()
        tot.addStretch()
        tot.addWidget(self.total_debit_lbl)
        tot.addWidget(QLabel('  |  '))
        tot.addWidget(self.total_credit_lbl)
        tot.addWidget(QLabel('  '))
        tot.addWidget(self.balance_lbl)
        root.addLayout(tot)

        # ── Dialog buttons ────────────────────────────────────────────
        dlg_btns = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        dlg_btns.accepted.connect(self._save)
        dlg_btns.rejected.connect(self.reject)
        root.addWidget(dlg_btns)

        self.setLayout(root)

        # Pre-fill lines when editing / copying
        if entry_data and entry_data.get('lines'):
            for ld in entry_data['lines']:
                self.lines.append(dict(ld))
            self._refresh_lines_table()

    # ── Line management ───────────────────────────────────────────────

    def _add_line(self):
        dlg = LineDialog(self.db_manager, self)
        if dlg.exec():
            self.lines.append(dlg.get_data())
            self._refresh_lines_table()

    def _edit_line(self):
        row = self.lines_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, 'No Selection', 'Select a line to edit.')
            return
        dlg = LineDialog(self.db_manager, self, line_data=self.lines[row])
        if dlg.exec():
            self.lines[row] = dlg.get_data()
            self._refresh_lines_table()

    def _remove_line(self):
        row = self.lines_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, 'No Selection', 'Select a line to remove.')
            return
        del self.lines[row]
        self._refresh_lines_table()

    def _refresh_lines_table(self):
        self.lines_table.setRowCount(len(self.lines))
        td = tc = 0.0
        for r, ld in enumerate(self.lines):
            self.lines_table.setItem(
                r, 0, QTableWidgetItem(ld.get('account_description', '')))
            self.lines_table.setItem(
                r, 1, QTableWidgetItem(ld.get('account_code', '')))
            d = QTableWidgetItem(f"{ld.get('debit',  0):,.2f}")
            c = QTableWidgetItem(f"{ld.get('credit', 0):,.2f}")
            d.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            c.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.lines_table.setItem(r, 2, d)
            self.lines_table.setItem(r, 3, c)
            td += ld.get('debit',  0)
            tc += ld.get('credit', 0)

        self.total_debit_lbl.setText(f'Total Debit:  {td:,.2f}')
        self.total_credit_lbl.setText(f'Total Credit: {tc:,.2f}')
        balanced = abs(td - tc) < 0.005 and td > 0
        if balanced:
            self.balance_lbl.setText('Balanced')
            self.balance_lbl.setStyleSheet('color: green; font-weight: bold;')
        elif self.lines:
            self.balance_lbl.setText('Unbalanced')
            self.balance_lbl.setStyleSheet('color: red; font-weight: bold;')
        else:
            self.balance_lbl.setText('')

    def _save(self):
        if not self.reference_input.text().strip():
            QMessageBox.warning(self, 'Validation', 'Reference No. is required.')
            return
        if not self.lines:
            QMessageBox.warning(self, 'Validation', 'Add at least one journal line.')
            return
        td = sum(l.get('debit',  0) for l in self.lines)
        tc = sum(l.get('credit', 0) for l in self.lines)
        if abs(td - tc) >= 0.005:
            reply = QMessageBox.question(
                self, 'Unbalanced Entry',
                f'Debit ({td:,.2f}) ≠ Credit ({tc:,.2f}).\nSave anyway?',
                QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
        self.accept()

    def get_data(self) -> list:
        date    = self.date_input.date().toString('MM/dd/yyyy')
        ref     = self.reference_input.text().strip()
        particu = self.particulars_input.text().strip()
        return [{
            'date':                date,
            'reference_no':        ref,
            'particulars':         particu,
            'account_description': ld.get('account_description', ''),
            'account_code':        ld.get('account_code', ''),
            'debit':               ld.get('debit',  0),
            'credit':              ld.get('credit', 0),
        } for ld in self.lines]


# ---------------------------------------------------------------------------
# Widget  (all logic inherited from BaseJournalWidget)
# ---------------------------------------------------------------------------

class CashDisbursementWidget(BaseJournalWidget):
    TITLE         = 'CASH DISBURSEMENT JOURNAL'
    DIALOG_CLASS  = CashDisbursementDialog
    GET_METHOD    = 'get_cash_disbursement_journal'
    ADD_METHOD    = 'add_cash_disbursement_entry'
    DELETE_METHOD = 'delete_cash_disbursement_entry'
    EXPORT_FOLDER = 'Cash Disbursement'
    EXPORT_FILE   = 'cash_disbursement_journal_report.xlsx'
    EXPORT_TITLE  = 'Cash Disbursement Journal'
    IMPORT_TITLE  = 'Import Cash Disbursement Journal'
    XLS_COLUMNS   = _XLS_COLUMNS