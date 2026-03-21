import os
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTableWidget,
                             QTableWidgetItem, QPushButton, QLabel, QLineEdit,
                             QDialog, QFormLayout, QDialogButtonBox, QMessageBox,
                             QHeaderView, QDateEdit, QDoubleSpinBox,
                             QComboBox, QGroupBox, QShortcut, QFileDialog)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QKeySequence
from database.db_manager import DatabaseManager
from resources.file_paths import get_import_dir, get_io_dir
from ui.search_utils import SearchFilter, add_month_combo

try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False

_XLS_COLUMNS = [
    ('Date',         'date'),
    ('Customer Name','customer_name'),
    ('Reference No', 'reference_no'),
    ('TIN',          'tin'),
    ('Net Amount',   'net_amount'),
    ('Output VAT',   'output_vat'),
    ('Gross Amount', 'gross_amount'),
    ('Goods',        'goods'),
    ('Services',     'services'),
    ('Particulars',  'particulars'),
]
_XLS_HEADERS = [h for h, _ in _XLS_COLUMNS]
_XLS_KEYS    = [k for _, k in _XLS_COLUMNS]

def _parse_date(date_str: str):
    for fmt in ("MM/dd/yyyy", "M/d/yyyy", "M/dd/yyyy", "MM/d/yyyy", "yyyy-MM-dd"):
        d = QDate.fromString(date_str, fmt)
        if d.isValid():
            return d
    return QDate()

class _DateItem(QTableWidgetItem):
    def __init__(self, display_text: str):
        super().__init__(display_text)
        qdate = _parse_date(display_text)
        self._sort_key = qdate.toString("yyyy-MM-dd") if qdate.isValid() else display_text
    def __lt__(self, other):
        if isinstance(other, _DateItem):
            return self._sort_key < other._sort_key
        return super().__lt__(other)


class SalesJournalDialog(QDialog):
    def __init__(self, db_manager, parent=None, entry_data=None, is_copy=False):
        super().__init__(parent)
        self.db_manager = db_manager
        self.entry_data = entry_data
        self.is_copy = is_copy
        self.alphalist_map = {}
        if is_copy:
            self.setWindowTitle("Copy Entry (Create New)")
        elif entry_data is None:
            self.setWindowTitle("Add Sales Entry")
        else:
            self.setWindowTitle("Edit Sales Entry")
        self.setModal(True)
        self.resize(500, 520)
        layout = QFormLayout()
        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDisplayFormat("MM/dd/yyyy")
        self.date_input.setDate(
            QDate.fromString(entry_data['date'], "MM/dd/yyyy") if entry_data else QDate.currentDate())
        layout.addRow("Date:", self.date_input)
        self.customer_input = QComboBox()
        self.customer_input.setEditable(True)
        self._load_customers()
        layout.addRow("Customer Name:", self.customer_input)
        self.tin_input = QLineEdit()
        self.tin_input.setReadOnly(True)
        if entry_data:
            self.tin_input.setText(str(entry_data.get('tin', '') or ''))
        layout.addRow("TIN:", self.tin_input)
        self.reference_input = QLineEdit()
        if entry_data:
            self.reference_input.setText(entry_data.get('reference_no', ''))
        layout.addRow("Reference No:", self.reference_input)
        self.goods_input = QDoubleSpinBox()
        self.goods_input.setMaximum(99999999.99)
        self.goods_input.setDecimals(2)
        self.goods_input.setGroupSeparatorShown(True)
        if entry_data:
            self.goods_input.setValue(entry_data.get('goods', 0))
        layout.addRow("Goods:", self.goods_input)
        self.services_input = QDoubleSpinBox()
        self.services_input.setMaximum(99999999.99)
        self.services_input.setDecimals(2)
        self.services_input.setGroupSeparatorShown(True)
        if entry_data:
            self.services_input.setValue(entry_data.get('services', 0))
        layout.addRow("Services:", self.services_input)
        self.net_amount_input = QDoubleSpinBox()
        self.net_amount_input.setMaximum(99999999.99)
        self.net_amount_input.setDecimals(2)
        self.net_amount_input.setGroupSeparatorShown(True)
        self.net_amount_input.setReadOnly(True)
        layout.addRow("Net Amount:", self.net_amount_input)
        self.output_vat_input = QDoubleSpinBox()
        self.output_vat_input.setMaximum(99999999.99)
        self.output_vat_input.setDecimals(2)
        self.output_vat_input.setGroupSeparatorShown(True)
        if entry_data:
            self.output_vat_input.setValue(entry_data.get('output_vat', 0))
        layout.addRow("Output VAT:", self.output_vat_input)
        self.gross_amount_input = QDoubleSpinBox()
        self.gross_amount_input.setMaximum(99999999.99)
        self.gross_amount_input.setDecimals(2)
        self.gross_amount_input.setGroupSeparatorShown(True)
        self.gross_amount_input.setReadOnly(True)
        if entry_data:
            self.gross_amount_input.setValue(entry_data.get('gross_amount', 0))
        layout.addRow("Gross Amount:", self.gross_amount_input)
        self.particulars_input = QLineEdit()
        if entry_data:
            self.particulars_input.setText(entry_data.get('particulars', '') or '')
        layout.addRow("Particulars:", self.particulars_input)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        self.setLayout(layout)
        self.customer_input.currentTextChanged.connect(self._on_customer_changed)
        self.goods_input.valueChanged.connect(self._calculate_totals)
        self.services_input.valueChanged.connect(self._calculate_totals)
        self.output_vat_input.valueChanged.connect(self._calculate_gross)
        if entry_data:
            self.customer_input.setCurrentText(entry_data.get('customer_name', '') or '')
        self._calculate_totals()

    def _load_customers(self):
        self.customer_input.addItem("")
        alphalist = self.db_manager.get_all_alphalist()
        for entry in alphalist:
            if entry.get('entry_type', 'Customer') not in ('Customer', 'Customer&Vendor'):
                continue
            name = entry.get('company_name') or \
                   f"{entry.get('first_name','')} {entry.get('last_name','')}".strip()
            if name:
                self.customer_input.addItem(name)
                self.alphalist_map[name] = {
                    'tin': entry.get('tin', ''),
                    'vat': entry.get('vat', 'VAT Regular'),
                }

    def _on_customer_changed(self, customer_name):
        info = self.alphalist_map.get(customer_name, {})
        self.tin_input.setText(str(info.get('tin', '')) if info else '')
        self._calculate_totals()

    def _get_vat_rate(self) -> float:
        from ui.alphalist_widget import VAT_TYPES
        customer = self.customer_input.currentText()
        vat_type = self.alphalist_map.get(customer, {}).get('vat', 'VAT Regular')
        return VAT_TYPES.get(vat_type, 0.12)

    def _calculate_totals(self):
        goods = self.goods_input.value()
        services = self.services_input.value()
        net = goods + services
        self.net_amount_input.blockSignals(True)
        self.net_amount_input.setValue(net)
        self.net_amount_input.blockSignals(False)
        vat_rate = self._get_vat_rate()
        self.output_vat_input.blockSignals(True)
        self.output_vat_input.setValue(round(net * vat_rate, 2))
        self.output_vat_input.blockSignals(False)
        self._calculate_gross()

    def _calculate_gross(self):
        self.gross_amount_input.setValue(
            round(self.net_amount_input.value() + self.output_vat_input.value(), 2))

    def get_data(self):
        return {
            'date':          self.date_input.date().toString("MM/dd/yyyy"),
            'customer_name': self.customer_input.currentText().strip(),
            'reference_no':  self.reference_input.text().strip(),
            'tin':           self.tin_input.text().strip(),
            'goods':         self.goods_input.value(),
            'services':      self.services_input.value(),
            'net_amount':    self.net_amount_input.value(),
            'output_vat':    self.output_vat_input.value(),
            'gross_amount':  self.gross_amount_input.value(),
            'particulars':   self.particulars_input.text().strip(),
        }


class SalesJournalWidget(QWidget):
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.all_entries = []
        self._setup_ui()
        self._setup_shortcuts()
        self.load_data()

    def _setup_ui(self):
        layout = QVBoxLayout()
        title = QLabel("SALES JOURNAL")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        search_group = QGroupBox("Search & Filter")
        search_layout = QHBoxLayout()
        self.month_combo = add_month_combo(search_layout)
        search_layout.addWidget(QLabel("Search:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by customer, reference, TIN, or particulars...")
        self.search_input.setClearButtonEnabled(True)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(QLabel("From:"))
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat("MM/dd/yyyy")
        self.date_from.setDate(QDate(2000, 1, 1))
        search_layout.addWidget(self.date_from)
        search_layout.addWidget(QLabel("To:"))
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat("MM/dd/yyyy")
        self.date_to.setDate(QDate.currentDate())
        search_layout.addWidget(self.date_to)
        self.clear_filter_btn = QPushButton("Clear Filter")
        self.clear_filter_btn.clicked.connect(self._clear_filters)
        search_layout.addWidget(self.clear_filter_btn)
        self.results_label = QLabel("Showing: 0 of 0")
        search_layout.addWidget(self.results_label)
        search_group.setLayout(search_layout)
        layout.addWidget(search_group)

        button_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Entry")
        self.add_btn.clicked.connect(self._add_entry)
        button_layout.addWidget(self.add_btn)
        self.edit_btn = QPushButton("Edit Entry")
        self.edit_btn.clicked.connect(self._edit_entry)
        button_layout.addWidget(self.edit_btn)
        self.copy_btn = QPushButton("Copy Entry")
        self.copy_btn.clicked.connect(self._copy_entry)
        button_layout.addWidget(self.copy_btn)
        self.delete_btn = QPushButton("Delete Entry")
        self.delete_btn.setProperty("class", "danger")
        self.delete_btn.clicked.connect(self._delete_entry)
        button_layout.addWidget(self.delete_btn)
        self.import_btn = QPushButton("Import")
        self.import_btn.clicked.connect(self._import_xls)
        button_layout.addWidget(self.import_btn)
        self.export_btn = QPushButton("Export")
        self.export_btn.clicked.connect(self._export_xls)
        button_layout.addWidget(self.export_btn)
        button_layout.addStretch()
        self.totals_label = QLabel("Totals: Net: 0.00 | VAT: 0.00 | Gross: 0.00")
        self.totals_label.setProperty("class", "total")
        button_layout.addWidget(self.totals_label)
        layout.addLayout(button_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "Date", "Customer Name", "Reference No.", "TIN",
            "Net Amount", "Output VAT", "Gross Amount", "Goods", "Services", "Particulars"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table)
        self.setLayout(layout)

        self._search = SearchFilter(
            table         = self.table,
            search_input  = self.search_input,
            results_label = self.results_label,
            date_from     = self.date_from,
            date_to       = self.date_to,
            month_combo   = self.month_combo,
            date_col      = 0,
        )
        self._search._run_orig = self._search._run
        # Hook totals update into filter
        orig_run = self._search._run
        def _run_with_totals():
            orig_run()
            self._update_totals_from_visible()
        self._search._timer.timeout.disconnect()
        self._search._timer.timeout.connect(_run_with_totals)
        self.date_from.dateChanged.disconnect()
        self.date_to.dateChanged.disconnect()
        self.month_combo.currentIndexChanged.disconnect()
        self.date_from.dateChanged.connect(_run_with_totals)
        self.date_to.dateChanged.connect(_run_with_totals)
        self.month_combo.currentIndexChanged.connect(_run_with_totals)
        self._run_with_totals = _run_with_totals

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+N"),       self).activated.connect(self._add_entry)
        QShortcut(QKeySequence("Ctrl+E"),       self).activated.connect(self._edit_entry)
        QShortcut(QKeySequence("Ctrl+Shift+C"), self).activated.connect(self._copy_entry)
        QShortcut(QKeySequence("Ctrl+D"),       self).activated.connect(self._delete_entry)
        QShortcut(QKeySequence("Ctrl+F"),       self).activated.connect(self.search_input.setFocus)
        QShortcut(QKeySequence("Ctrl+I"),       self).activated.connect(self._import_xls)
        QShortcut(QKeySequence("Ctrl+Shift+E"), self).activated.connect(self._export_xls)

    def load_data(self):
        self.all_entries = self.db_manager.get_sales_journal()
        self._populate_table(self.all_entries)
        self._search.refresh()
        self._update_totals_from_visible()

    def _populate_table(self, entries):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(entries))
        num_cols = ['net_amount', 'output_vat', 'gross_amount', 'goods', 'services']
        for row, entry in enumerate(entries):
            date_item = _DateItem(entry.get('date', ''))
            date_item.setData(Qt.UserRole, entry['id'])
            self.table.setItem(row, 0, date_item)
            for text, col in [
                (entry.get('customer_name', ''), 1),
                (entry.get('reference_no', ''),  2),
                (str(entry.get('tin', '')),       3),
            ]:
                self.table.setItem(row, col, QTableWidgetItem(text))
            for offset, field in enumerate(num_cols):
                val  = entry.get(field, 0)
                item = QTableWidgetItem(f"{val:,.2f}")
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.table.setItem(row, 4 + offset, item)
            self.table.setItem(row, 9, QTableWidgetItem(entry.get('particulars', '') or ''))
        self.table.setSortingEnabled(True)

    def _update_totals_from_visible(self):
        net = vat = gross = 0.0
        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue
            try:
                net   += float(self.table.item(row, 4).text().replace(',', ''))
                vat   += float(self.table.item(row, 5).text().replace(',', ''))
                gross += float(self.table.item(row, 6).text().replace(',', ''))
            except (AttributeError, ValueError):
                pass
        self.totals_label.setText(
            f"Totals: Net: {net:,.2f} | VAT: {vat:,.2f} | Gross: {gross:,.2f}")

    def _clear_filters(self):
        self.search_input.clear()
        self.date_from.setDate(QDate(2000, 1, 1))
        self.date_to.setDate(QDate.currentDate())
        self.month_combo.setCurrentIndex(0)

    def _get_selected_entry_data(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        return {
            'id':            self.table.item(row, 0).data(Qt.UserRole),
            'date':          self.table.item(row, 0).text(),
            'customer_name': self.table.item(row, 1).text(),
            'reference_no':  self.table.item(row, 2).text(),
            'tin':           self.table.item(row, 3).text(),
            'net_amount':    float(self.table.item(row, 4).text().replace(',', '')),
            'output_vat':    float(self.table.item(row, 5).text().replace(',', '')),
            'gross_amount':  float(self.table.item(row, 6).text().replace(',', '')),
            'goods':         float(self.table.item(row, 7).text().replace(',', '')),
            'services':      float(self.table.item(row, 8).text().replace(',', '')),
            'particulars':   self.table.item(row, 9).text(),
        }

    def _add_entry(self):
        dialog = SalesJournalDialog(self.db_manager, self)
        if dialog.exec_():
            data = dialog.get_data()
            if data['customer_name'] and data['reference_no']:
                if self.db_manager.add_sales_entry(data):
                    self.load_data()
                    QMessageBox.information(self, "Success", "Sales entry added successfully!")
                else:
                    QMessageBox.warning(self, "Error", "Failed to add entry!")

    def _edit_entry(self):
        entry_data = self._get_selected_entry_data()
        if not entry_data:
            QMessageBox.warning(self, "Warning", "Please select an entry to edit")
            return
        dialog = SalesJournalDialog(self.db_manager, self, entry_data)
        if dialog.exec_():
            if self.db_manager.update_sales_entry(entry_data['id'], dialog.get_data()):
                self.load_data()
                QMessageBox.information(self, "Success", "Entry updated successfully!")
            else:
                QMessageBox.warning(self, "Error", "Failed to update entry!")

    def _copy_entry(self):
        entry_data = self._get_selected_entry_data()
        if not entry_data:
            QMessageBox.warning(self, "Warning", "Please select an entry to copy")
            return
        entry_data.pop('id', None)
        dialog = SalesJournalDialog(self.db_manager, self, entry_data, is_copy=True)
        if dialog.exec_():
            data = dialog.get_data()
            if data['customer_name'] and data['reference_no']:
                if self.db_manager.add_sales_entry(data):
                    self.load_data()
                    QMessageBox.information(self, "Success", "Entry copied successfully!")
                else:
                    QMessageBox.warning(self, "Error", "Failed to add copied entry!")

    def _delete_entry(self):
        entry_data = self._get_selected_entry_data()
        if not entry_data:
            QMessageBox.warning(self, "Warning", "Please select an entry to delete")
            return
        reply = QMessageBox.question(
            self, "Confirm Delete",
            f"Delete entry '{entry_data['reference_no']}'?",
            QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            if self.db_manager.delete_sales_entry(entry_data['id']):
                self.load_data()
                QMessageBox.information(self, "Success", "Entry deleted successfully!")
            else:
                QMessageBox.warning(self, "Error", "Failed to delete entry!")

    def _export_xls(self):
        if not _OPENPYXL_OK:
            QMessageBox.critical(self, "Missing Library", "openpyxl is required.\nInstall with: pip install openpyxl")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Sales Journal",
            os.path.join(get_io_dir("Sales Journal"), "sales_journal_report.xlsx"),
            "Excel Files (*.xlsx)")
        if not path:
            return
        rows = []
        for r in range(self.table.rowCount()):
            if not self.table.isRowHidden(r):
                rows.append({k: self.table.item(r, c).text()
                              for c, (_, k) in enumerate(_XLS_COLUMNS)})
        n, err = _export_to_xls(rows, path, "Sales Journal")
        if err:
            QMessageBox.critical(self, "Export Failed", err)
        else:
            QMessageBox.information(self, "Export Successful", f"{n} record(s) exported to:\n{path}")

    def _import_xls(self):
        if not _OPENPYXL_OK:
            QMessageBox.critical(self, "Missing Library", "openpyxl is required.\nInstall with: pip install openpyxl")
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Import Sales Journal", get_import_dir(""), "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        try:
            imported, skipped, errors = _import_from_xls(path, self.db_manager)
        except Exception as exc:
            QMessageBox.critical(self, "Import Failed", str(exc))
            return
        self.load_data()
        msg = f"Import complete.\n  Imported: {imported}\n  Skipped: {skipped}"
        if errors:
            msg += "\n\nDetails:\n" + "\n".join(errors[:20])
        QMessageBox.information(self, "Import Summary", msg)


def _export_to_xls(rows: list, path: str, sheet_title: str) -> tuple:
    from datetime import datetime as _dt
    try:
        wb  = Workbook()
        ws  = wb.active
        ws.title = sheet_title[:31]
        hdr_font  = Font(name='Arial', bold=True, color='FFFFFF', size=11)
        hdr_fill  = PatternFill('solid', start_color='2F5496')
        hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell_font = Font(name='Arial', size=10)
        alt_fill  = PatternFill('solid', start_color='DCE6F1')
        thin      = Side(style='thin', color='B0B0B0')
        border    = Border(left=thin, right=thin, top=thin, bottom=thin)
        t_font    = Font(name='Arial', bold=True, size=14)
        t_align   = Alignment(horizontal='left', vertical='center')
        HEADER_ROW = 5
        ws.row_dimensions[2].height = 22
        ws.merge_cells(f'A2:{get_column_letter(len(_XLS_HEADERS))}2')
        tc = ws['A2']
        tc.value = sheet_title.upper(); tc.font = t_font; tc.alignment = t_align
        ws.row_dimensions[3].height = 18
        ws.merge_cells(f'A3:{get_column_letter(len(_XLS_HEADERS))}3')
        sc = ws['A3']
        sc.value = f'For the Year {_dt.now().year}'
        sc.font = Font(name='Arial', italic=True, size=11); sc.alignment = t_align
        ws.row_dimensions[HEADER_ROW].height = 28
        for ci, hdr in enumerate(_XLS_HEADERS, 1):
            cell = ws.cell(row=HEADER_ROW, column=ci, value=hdr)
            cell.font = hdr_font; cell.fill = hdr_fill
            cell.alignment = hdr_align; cell.border = border
        for ri, entry in enumerate(rows):
            row_idx = 6 + ri
            ws.row_dimensions[row_idx].height = 18
            fill = alt_fill if ri % 2 == 0 else None
            for ci, key in enumerate(_XLS_KEYS, 1):
                val  = entry.get(key, '') or ''
                cell = ws.cell(row=row_idx, column=ci, value=val)
                cell.font = cell_font; cell.border = border
                cell.alignment = Alignment(
                    horizontal='right' if ci >= 5 else 'left', vertical='center')
                if fill:
                    cell.fill = fill
        widths = {1:12, 2:28, 3:16, 4:16, 5:14, 6:14, 7:14, 8:14, 9:14, 10:28}
        for ci, w in widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A6'
        last_col = get_column_letter(len(_XLS_HEADERS))
        ws.auto_filter.ref = f'A{HEADER_ROW}:{last_col}{HEADER_ROW}'
        wb.save(path)
        return len(rows), ''
    except Exception as exc:
        return 0, str(exc)


def _import_from_xls(path: str, db_manager) -> tuple:
    HEADER_MAP = {h.lower(): k for h, k in _XLS_COLUMNS}
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        raise RuntimeError(
            f'Cannot open "{os.path.basename(path)}".\n'
            'If this is a legacy .xls file, re-save as .xlsx first.')
    ws = wb.active
    col_index  = {}
    data_start = None
    for r_idx, row_vals in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
        if not any(v and str(v).strip().lower() in HEADER_MAP for v in row_vals):
            continue
        for ci, cv in enumerate(row_vals):
            if cv is None:
                continue
            key = HEADER_MAP.get(str(cv).strip().lower())
            if key:
                col_index[key] = ci
        data_start = r_idx + 1
        break
    if not col_index:
        raise ValueError("Could not find a matching header row in the first 10 rows.")
    def _val(row_vals, key):
        idx = col_index.get(key)
        if idx is None or idx >= len(row_vals):
            return ''
        v = row_vals[idx]
        return str(v).strip() if v is not None else ''
    imported = skipped = 0
    errors   = []
    for rn, row_vals in enumerate(ws.iter_rows(min_row=data_start, values_only=True), data_start):
        if all(v is None for v in row_vals):
            continue
        try:
            def _f(key): return float(_val(row_vals, key).replace(',', '') or 0)
            data = {k: _val(row_vals, k) for k in _XLS_KEYS}
            data['net_amount']   = _f('net_amount')
            data['output_vat']   = _f('output_vat')
            data['gross_amount'] = _f('gross_amount')
            data['goods']        = _f('goods')
            data['services']     = _f('services')
            if not data.get('reference_no'):
                skipped += 1
                errors.append(f"Row {rn}: missing reference_no skipped")
                continue
            if db_manager.add_sales_entry(data):
                imported += 1
            else:
                skipped += 1
                errors.append(f"Row {rn}: failed to insert skipped")
        except Exception as e:
            skipped += 1
            errors.append(f"Row {rn}: {e}")
    wb.close()
    return imported, skipped, errors