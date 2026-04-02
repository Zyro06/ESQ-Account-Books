"""
ui/dialogs/view_details_dialog.py
----------------------------------
Shared read-only detail viewer for journal entries.

Covers three layout variants used across five widgets:

    ViewDetailsDialog        — CDJ / CRJ / GJ
        Signature: (parent, date, reference_no, particulars, lines)
        Shows: Date, Reference No., Particulars header + lines table

    ViewEntryDialog          — SJ / PJ (entry-dict variant)
        Signature: (parent, entry, name_label='Customer')
        Shows: Date, name field, Reference, TIN, Particulars header + lines table
        Pass name_label='Payee' for Purchase Journal.

Both classes share the same internal _build_lines_table() helper so the
table layout is guaranteed to be identical everywhere.

Usage — CDJ / CRJ / GJ
-----------------------
    from ui.dialogs.view_details_dialog import ViewDetailsDialog

    # inside _view_details():
    ViewDetailsDialog(
        self,
        group['date'],
        group['reference_no'],
        group['particulars'],
        group['lines'],
    ).exec_()

Usage — SJ
----------
    from ui.dialogs.view_details_dialog import ViewEntryDialog

    ViewEntryDialog(self, entry, name_label='Customer').exec_()

Usage — PJ
----------
    ViewEntryDialog(self, entry, name_label='Payee').exec_()

Migration note (PySide6)
------------------------
Change the two PyQt5 import lines at the top to PySide6 equivalents,
and rename exec_() → exec() at all call sites.
"""

from __future__ import annotations

from PySide6.QtWidgets import (   # ← swap to PySide6.QtWidgets when migrating
    QDialog, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QFrame,
)
from PySide6.QtCore import Qt     # ← swap to PySide6.QtCore when migrating
from PySide6.QtGui  import QFont  # ← swap to PySide6.QtGui  when migrating


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _build_lines_table(lines: list[dict]) -> tuple[QTableWidget, float, float]:
    """
    Build and populate a read-only 4-column lines table.

    Returns
    -------
    (table_widget, total_debit, total_credit)
    """
    tbl = QTableWidget()
    tbl.setColumnCount(4)
    tbl.setHorizontalHeaderLabels(
        ['Account Description', 'Account Code', 'Debit', 'Credit'])
    tbl.setEditTriggers(QTableWidget.NoEditTriggers)
    tbl.setAlternatingRowColors(True)
    tbl.verticalHeader().setVisible(False)

    hdr = tbl.horizontalHeader()
    hdr.setSectionResizeMode(0, QHeaderView.Stretch)
    hdr.setSectionResizeMode(1, QHeaderView.ResizeToContents)
    hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
    hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)

    tbl.setRowCount(len(lines))
    td = tc = 0.0

    for r, ld in enumerate(lines):
        d = float(ld.get('debit',  0) or 0)
        c = float(ld.get('credit', 0) or 0)

        tbl.setItem(r, 0, QTableWidgetItem(ld.get('account_description', '')))
        tbl.setItem(r, 1, QTableWidgetItem(ld.get('account_code', '')))

        d_item = QTableWidgetItem(f'{d:,.2f}' if d else '')
        c_item = QTableWidgetItem(f'{c:,.2f}' if c else '')
        d_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        c_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        d_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        c_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

        tbl.setItem(r, 2, d_item)
        tbl.setItem(r, 3, c_item)
        td += d
        tc += c

    return tbl, td, tc


def _build_totals_row(td: float, tc: float) -> QHBoxLayout:
    """Build the "Total Debit | Total Credit" footer layout."""
    bf = QFont()
    bf.setBold(True)
    tdl = QLabel(f'Total Debit:  {td:,.2f}')
    tcl = QLabel(f'Total Credit: {tc:,.2f}')
    tdl.setFont(bf)
    tcl.setFont(bf)
    row = QHBoxLayout()
    row.addStretch()
    row.addWidget(tdl)
    row.addWidget(QLabel('  |  '))
    row.addWidget(tcl)
    return row


def _build_separator() -> QFrame:
    sep = QFrame()
    sep.setFrameShape(QFrame.HLine)
    sep.setFrameShadow(QFrame.Sunken)
    return sep


# ---------------------------------------------------------------------------
# ViewDetailsDialog  (CDJ / CRJ / GJ)
# ---------------------------------------------------------------------------

class ViewDetailsDialog(QDialog):
    """
    Read-only detail viewer for CDJ, CRJ, and GJ entries.

    Parameters
    ----------
    parent       : parent QWidget
    date         : entry date string e.g. '03/25/2026'
    reference_no : reference number string
    particulars  : particulars / description string (may be empty)
    lines        : list of line dicts with keys
                   'account_description', 'account_code', 'debit', 'credit'
    """

    def __init__(
        self,
        parent,
        date:         str,
        reference_no: str,
        particulars:  str,
        lines:        list[dict],
    ):
        super().__init__(parent)
        self.setWindowTitle(f'View Details — {reference_no}')
        self.setModal(True)
        self.setMinimumWidth(720)
        self.setMinimumHeight(520)

        root = QVBoxLayout()

        # ── Header info ───────────────────────────────────────────────────
        info = QLabel(
            f'<b>Date:</b> {date}'
            f'&nbsp;&nbsp;&nbsp;&nbsp;'
            f'<b>Reference No.:</b> {reference_no}<br>'
            f'<b>Particulars:</b> {particulars or ""}'
        )
        info.setStyleSheet('font-size: 12px; margin-bottom: 6px;')
        root.addWidget(info)
        root.addWidget(_build_separator())

        # ── Lines table ───────────────────────────────────────────────────
        tbl, td, tc = _build_lines_table(lines)
        root.addWidget(tbl)

        # ── Totals ────────────────────────────────────────────────────────
        root.addLayout(_build_totals_row(td, tc))

        # ── Close button ──────────────────────────────────────────────────
        close_btn = QPushButton('Close')
        close_btn.clicked.connect(self.accept)
        root.addWidget(close_btn, alignment=Qt.AlignRight)

        self.setLayout(root)


# ---------------------------------------------------------------------------
# ViewEntryDialog  (SJ / PJ — entry-dict variant)
# ---------------------------------------------------------------------------

class ViewEntryDialog(QDialog):
    """
    Read-only detail viewer for Sales Journal and Purchase Journal entries.

    Parameters
    ----------
    parent      : parent QWidget
    entry       : full entry dict as returned by db_manager.get_sales_journal()
                  or get_purchase_journal() — must have keys:
                  'date', 'reference_no', 'tin', 'particulars', 'lines'
                  and either 'customer_name' (SJ) or 'payee_name' (PJ)
    name_label  : label for the name field — 'Customer' (SJ) or 'Payee' (PJ)
    name_key    : dict key for the name — defaults to lower(name_label)+'_name'
                  e.g. name_label='Customer' → name_key='customer_name'
    """

    def __init__(
        self,
        parent,
        entry:       dict,
        name_label:  str = 'Customer',
        name_key:    str | None = None,
    ):
        ref = entry.get('reference_no', '')
        super().__init__(parent)
        self.setWindowTitle(f'View — {ref}')
        self.setModal(True)
        self.setMinimumWidth(700)
        self.setMinimumHeight(460)

        _name_key = name_key or f'{name_label.lower()}_name'
        name_val  = entry.get(_name_key, '')

        root = QVBoxLayout(self)

        # ── Header info ───────────────────────────────────────────────────
        info = QLabel(
            f'<b>Date:</b> {entry.get("date", "")}'
            f'&nbsp;&nbsp;'
            f'<b>{name_label}:</b> {name_val}'
            f'&nbsp;&nbsp;'
            f'<b>Reference:</b> {ref}'
            f'&nbsp;&nbsp;'
            f'<b>TIN:</b> {entry.get("tin", "") or ""}'
            f'&nbsp;&nbsp;'
            f'<b>Particulars:</b> {entry.get("particulars", "") or ""}'
        )
        info.setWordWrap(True)
        root.addWidget(info)
        root.addWidget(_build_separator())

        # ── Lines table ───────────────────────────────────────────────────
        lines = entry.get('lines', [])
        tbl, td, tc = _build_lines_table(lines)
        root.addWidget(tbl)

        # ── Totals ────────────────────────────────────────────────────────
        root.addLayout(_build_totals_row(td, tc))

        # ── Close button ──────────────────────────────────────────────────
        close_btn = QPushButton('Close')
        close_btn.clicked.connect(self.accept)
        root.addWidget(close_btn, alignment=Qt.AlignRight)


# ---------------------------------------------------------------------------
# __init__.py convenience — so callers can do either:
#   from ui.dialogs.view_details_dialog import ViewDetailsDialog
#   from ui.dialogs.view_details_dialog import ViewEntryDialog
# ---------------------------------------------------------------------------

__all__ = [
    'ViewDetailsDialog',
    'ViewEntryDialog',
]