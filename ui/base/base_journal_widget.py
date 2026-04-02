"""
ui/base/base_journal_widget.py
"""

from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLabel, QLineEdit, QHeaderView, QDateEdit,
    QGroupBox, QFileDialog, QMessageBox,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui  import QKeySequence, QShortcut

from database.db_manager  import DatabaseManager
from ui.utils.search_utils import SearchFilter, add_month_combo
from ui.dialogs.view_details_dialog import ViewDetailsDialog
from utils.date_utils   import DateItem
from utils.export_utils import export_to_xls
from utils.import_utils import import_from_xls
from resources.file_paths import get_import_dir, get_io_dir

try:
    from openpyxl import Workbook  # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False


class BaseJournalWidget(QWidget):

    TITLE:          str  = ''
    DIALOG_CLASS          = None
    GET_METHOD:     str  = ''
    ADD_METHOD:     str  = ''
    DELETE_METHOD:  str  = ''
    EXPORT_FOLDER:  str  = ''
    EXPORT_FILE:    str  = 'journal_report.xlsx'
    EXPORT_TITLE:   str  = 'Journal'
    IMPORT_TITLE:   str  = 'Import Journal'
    XLS_COLUMNS:    list = []
    XLS_COL_WIDTHS: dict = {1: 12, 2: 16, 3: 28, 4: 28, 5: 14, 6: 14, 7: 14}

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager  = db_manager
        self.all_entries: list = []
        self._setup_ui()
        self._setup_shortcuts()
        self.load_data()

    def _setup_ui(self):
        layout = QVBoxLayout()

        # ── Title ─────────────────────────────────────────────────────
        title = QLabel(self.TITLE)
        title.setProperty('class', 'title')
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # ── Search & Filter ───────────────────────────────────────────
        sg = QGroupBox('Search & Filter')
        sl = QHBoxLayout()
        self.month_combo = add_month_combo(sl)
        sl.addWidget(QLabel('Search:'))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(
            'Search by reference, account, or particulars')
        self.search_input.setClearButtonEnabled(True)
        sl.addWidget(self.search_input)
        sl.addWidget(QLabel('From:'))
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat('MM/dd/yyyy')
        self.date_from.setDate(QDate(2000, 1, 1))
        sl.addWidget(self.date_from)
        sl.addWidget(QLabel('To:'))
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat('MM/dd/yyyy')
        self.date_to.setDate(QDate.currentDate())
        sl.addWidget(self.date_to)
        self.clear_filter_btn = QPushButton('Clear Filter')
        self.clear_filter_btn.clicked.connect(self._clear_filters)
        sl.addWidget(self.clear_filter_btn)
        self.results_label = QLabel('Showing: 0 of 0')
        sl.addWidget(self.results_label)
        sg.setLayout(sl)
        layout.addWidget(sg)

        # ── Action buttons row ────────────────────────────────────────
        br = QHBoxLayout()
        self.add_btn    = QPushButton('Add Entry')
        self.edit_btn   = QPushButton('Edit Entry')
        self.copy_btn   = QPushButton('Copy Entry')
        self.view_btn   = QPushButton('View Details')
        self.delete_btn = QPushButton('Delete Entry')
        self.import_btn = QPushButton('Import')
        self.export_btn = QPushButton('Export')

        self.add_btn.clicked.connect(self._add_entry)
        self.edit_btn.clicked.connect(self._edit_entry)
        self.copy_btn.clicked.connect(self._copy_entry)
        self.view_btn.clicked.connect(self._view_details)
        self.delete_btn.clicked.connect(self._delete_entry)
        self.import_btn.clicked.connect(self._import_xls)
        self.export_btn.clicked.connect(self._export_xls)

        self.delete_btn.setProperty('class', 'danger')

        for btn in (self.add_btn, self.edit_btn, self.copy_btn,
                    self.view_btn, self.delete_btn,
                    self.import_btn, self.export_btn):
            br.addWidget(btn)
        br.addStretch()
        layout.addLayout(br)

        # ── Table ──────────────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            'Date', 'Reference No.', 'Particulars',
            'Lines', 'Total Debit', 'Total Credit'])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)   # ← multi-select
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self._view_details)
        self.table.selectionModel().selectionChanged.connect(self._on_selection_changed)
        layout.addWidget(self.table)

        # ── Totals — below the table ───────────────────────────────────
        self.totals_label = QLabel('Totals Debit: 0.00 | Credit: 0.00')
        self.totals_label.setProperty('class', 'total')
        self.totals_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        layout.addWidget(self.totals_label)

        # ── SearchFilter with totals hook ─────────────────────────────
        self._search = SearchFilter(
            table         = self.table,
            search_input  = self.search_input,
            results_label = self.results_label,
            date_from     = self.date_from,
            date_to       = self.date_to,
            month_combo   = self.month_combo,
            date_col      = 0,
        )
        orig_run = self._search._run

        def _run_with_totals():
            orig_run()
            self._update_totals_from_visible()

        self._search._timer.timeout.disconnect()
        self._search._timer.timeout.connect(_run_with_totals)
        self.date_from.dateChanged.connect(_run_with_totals)
        self.date_to.dateChanged.connect(_run_with_totals)
        self.month_combo.currentIndexChanged.connect(_run_with_totals)
        self._run_with_totals = _run_with_totals

        self.setLayout(layout)

    def _on_selection_changed(self):
        """Update delete button label to show count of selected rows."""
        count = len(self._get_selected_groups())
        if count > 1:
            self.delete_btn.setText(f'Delete ({count})')
        else:
            self.delete_btn.setText('Delete Entry')

    def _setup_shortcuts(self):
        QShortcut(QKeySequence('Ctrl+N'),       self).activated.connect(self._add_entry)
        QShortcut(QKeySequence('Ctrl+E'),       self).activated.connect(self._edit_entry)
        QShortcut(QKeySequence('Ctrl+Shift+C'), self).activated.connect(self._copy_entry)
        QShortcut(QKeySequence('Ctrl+V'),       self).activated.connect(self._view_details)
        QShortcut(QKeySequence('Ctrl+D'),       self).activated.connect(self._delete_entry)
        QShortcut(QKeySequence('Ctrl+F'),       self).activated.connect(self.search_input.setFocus)
        QShortcut(QKeySequence('Ctrl+I'),       self).activated.connect(self._import_xls)
        QShortcut(QKeySequence('Ctrl+Shift+E'), self).activated.connect(self._export_xls)

    def load_data(self):
        fetch = getattr(self.db_manager, self.GET_METHOD)
        self.all_entries = fetch()
        self._populate_table(self._group_entries(self.all_entries))
        self._search.refresh()
        self._update_totals_from_visible()

    def _group_entries(self, entries: list) -> list:
        groups: dict = {}
        order:  list = []
        for e in entries:
            key = (e.get('date', ''), e.get('reference_no', ''))
            if key not in groups:
                groups[key] = {
                    'date':         e.get('date', ''),
                    'reference_no': e.get('reference_no', ''),
                    'particulars':  e.get('particulars', '') or '',
                    'lines':        [],
                    'ids':          [],
                }
                order.append(key)
            groups[key]['lines'].append({
                'account_description': e.get('account_description', ''),
                'account_code':        e.get('account_code', '') or '',
                'debit':               e.get('debit',  0),
                'credit':              e.get('credit', 0),
            })
            groups[key]['ids'].append(e['id'])
        return [groups[k] for k in order]

    def _populate_table(self, groups: list):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(groups))
        for r, g in enumerate(groups):
            td = sum(l['debit']  for l in g['lines'])
            tc = sum(l['credit'] for l in g['lines'])

            date_item = DateItem(g['date'])
            date_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            date_item.setData(Qt.UserRole, g)
            self.table.setItem(r, 0, date_item)

            for c, text in enumerate([
                g['reference_no'],
                g['particulars'],
                str(len(g['lines'])),
                f'{td:,.2f}',
                f'{tc:,.2f}',
            ], start=1):
                item = QTableWidgetItem(text)
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                if c >= 3:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(r, c, item)

        self.table.setSortingEnabled(True)

    def _update_totals_from_visible(self):
        td = tc = 0.0
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            try:
                td += float(self.table.item(row, 4).text().replace(',', ''))
                tc += float(self.table.item(row, 5).text().replace(',', ''))
            except (AttributeError, ValueError):
                pass
        self.totals_label.setText(
            f'Totals  |  Debit: {td:,.2f}  |  Credit: {tc:,.2f}')

    def _clear_filters(self):
        self.search_input.clear()
        self.date_from.setDate(QDate(2000, 1, 1))
        self.date_to.setDate(QDate.currentDate())
        self.month_combo.setCurrentIndex(0)

    def _get_selected_group(self) -> dict | None:
        """Return the single currently selected group (for edit/copy/view)."""
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.UserRole) if item else None

    def _get_selected_groups(self) -> list[dict]:
        """Return all selected groups (for multi-delete)."""
        seen = set()
        groups = []
        for index in self.table.selectionModel().selectedRows():
            row  = index.row()
            item = self.table.item(row, 0)
            if item is None:
                continue
            g = item.data(Qt.UserRole)
            if g is None:
                continue
            key = (g.get('date', ''), g.get('reference_no', ''))
            if key not in seen:
                seen.add(key)
                groups.append(g)
        return groups

    def _add_entry(self):
        dialog = self.DIALOG_CLASS(self.db_manager, self)
        if dialog.exec():
            rows = dialog.get_data()
            add_fn = getattr(self.db_manager, self.ADD_METHOD)
            ok = all(add_fn(r) for r in rows)
            self.load_data()
            (QMessageBox.information if ok else QMessageBox.warning)(
                self,
                'Success' if ok else 'Error',
                'Entry added successfully!' if ok else 'Some lines failed to save.')

    def _edit_entry(self):
        group = self._get_selected_group()
        if not group:
            QMessageBox.warning(self, 'Warning', 'Please select an entry to edit.')
            return
        dialog = self.DIALOG_CLASS(self.db_manager, self, group)
        if not dialog.exec():
            return
        new_rows = dialog.get_data()
        del_fn = getattr(self.db_manager, self.DELETE_METHOD)
        add_fn = getattr(self.db_manager, self.ADD_METHOD)
        for old_id in group['ids']:
            del_fn(old_id)
        ok = all(add_fn(r) for r in new_rows)
        self.load_data()
        (QMessageBox.information if ok else QMessageBox.warning)(
            self,
            'Success' if ok else 'Error',
            'Entry updated successfully!' if ok else 'Some lines failed to save.')

    def _copy_entry(self):
        group = self._get_selected_group()
        if not group:
            QMessageBox.warning(self, 'Warning', 'Please select an entry to copy.')
            return
        copy = {k: v for k, v in group.items() if k != 'ids'}
        dialog = self.DIALOG_CLASS(self.db_manager, self, copy, is_copy=True)
        if not dialog.exec():
            return
        rows = dialog.get_data()
        add_fn = getattr(self.db_manager, self.ADD_METHOD)
        ok = all(add_fn(r) for r in rows)
        self.load_data()
        (QMessageBox.information if ok else QMessageBox.warning)(
            self,
            'Success' if ok else 'Error',
            'Entry copied successfully!' if ok else 'Some lines failed to save.')

    def _view_details(self):
        group = self._get_selected_group()
        if not group:
            QMessageBox.warning(self, 'Warning', 'Please select an entry to view.')
            return
        ViewDetailsDialog(
            self,
            group['date'],
            group['reference_no'],
            group['particulars'],
            group['lines'],
        ).exec()

    def _delete_entry(self):
        groups = self._get_selected_groups()
        if not groups:
            QMessageBox.warning(self, 'Warning', 'Please select one or more entries to delete.')
            return

        count   = len(groups)
        del_fn  = getattr(self.db_manager, self.DELETE_METHOD)

        if count == 1:
            g = groups[0]
            msg = (f"Delete all {len(g['ids'])} line(s) for "
                   f"'{g['reference_no']}' ({g['date']})?")
        else:
            refs = ', '.join(g['reference_no'] for g in groups[:5])
            if count > 5:
                refs += f' … and {count - 5} more'
            msg = f"Delete {count} selected entries?\n\n{refs}"

        reply = QMessageBox.question(
            self, 'Confirm Delete', msg,
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            for g in groups:
                for old_id in g['ids']:
                    del_fn(old_id)
            self.load_data()
            QMessageBox.information(
                self, 'Success',
                f'{count} entr{"y" if count == 1 else "ies"} deleted successfully!')

    def _export_xls(self):
        if not _OPENPYXL_OK:
            QMessageBox.critical(
                self, 'Missing Library',
                'openpyxl is required.\nInstall with: pip install openpyxl')
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            f'Export {self.EXPORT_TITLE}',
            os.path.join(get_io_dir(self.EXPORT_FOLDER), self.EXPORT_FILE),
            'Excel Files (*.xlsx)',
        )
        if not path:
            return

        rows = []
        for r in range(self.table.rowCount()):
            if self.table.isRowHidden(r):
                continue
            g = self.table.item(r, 0).data(Qt.UserRole)
            for ld in g['lines']:
                rows.append({
                    'date':                g['date'],
                    'reference_no':        g['reference_no'],
                    'particulars':         g['particulars'],
                    'account_description': ld['account_description'],
                    'account_code':        ld['account_code'],
                    'debit':  f"{ld['debit']:,.2f}",
                    'credit': f"{ld['credit']:,.2f}",
                })

        n, err = export_to_xls(
            rows        = rows,
            path        = path,
            sheet_title = self.EXPORT_TITLE,
            columns     = self.XLS_COLUMNS,
            col_widths  = self.XLS_COL_WIDTHS,
        )
        if err:
            QMessageBox.critical(self, 'Export Failed', err)
        else:
            QMessageBox.information(
                self, 'Export Successful',
                f'{n} line(s) exported to:\n{path}')

    def _import_xls(self):
        if not _OPENPYXL_OK:
            QMessageBox.critical(
                self, 'Missing Library',
                'openpyxl is required.\nInstall with: pip install openpyxl')
            return

        path, _ = QFileDialog.getOpenFileName(
            self, self.IMPORT_TITLE,
            get_import_dir(''),
            'Excel Files (*.xlsx *.xls)',
        )
        if not path:
            return

        try:
            imported, skipped, errors = import_from_xls(
                path            = path,
                db_manager      = self.db_manager,
                add_method_name = self.ADD_METHOD,
                columns         = self.XLS_COLUMNS,
            )
        except Exception as exc:
            QMessageBox.critical(self, 'Import Failed', str(exc))
            return

        self.load_data()
        msg = f'Import complete.\n  Imported: {imported}\n  Skipped: {skipped}'
        if errors:
            msg += '\n\nDetails:\n' + '\n'.join(errors[:20])
        QMessageBox.information(self, 'Import Summary', msg)