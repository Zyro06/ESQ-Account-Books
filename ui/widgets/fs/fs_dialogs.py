"""
ui/widgets/fs/fs_dialogs.py
----------------------------
Dialogs used by FinancialStatementsWidget:
    GenerateDialog      — statement type / date picker
    PrintPreviewDialog  — monospace preview + Print / PDF buttons
"""

from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, QDialogButtonBox,
    QLabel, QComboBox, QDateEdit, QTextEdit, QPushButton,
    QFrame, QFileDialog, QMessageBox,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui  import QFont, QPageLayout, QPageSize
from PySide6.QtPrintSupport import QPrinter, QPrintDialog
from PySide6.QtCore import QMarginsF

from resources.file_paths import get_io_dir
from ui.widgets.fs.fs_builders import MONTHS, month_last_day
from ui.widgets.fs.fs_pdf      import save_pdf


# ---------------------------------------------------------------------------
# GenerateDialog
# ---------------------------------------------------------------------------

class GenerateDialog(QDialog):
    _POSITION_NAMES    = ["Statement of Financial Position", "Balance Sheet"]
    _PERFORMANCE_NAMES = ["Statement of Financial Performance", "Income Statement"]

    def __init__(self, company_name: str, parent=None):
        super().__init__(parent)
        self.company_name = company_name
        self.setWindowTitle("Generate Financial Statement")
        self.setModal(True)
        self.setMinimumWidth(420)

        root = QVBoxLayout()
        root.setSpacing(14)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        self.type_combo = QComboBox()
        self.type_combo.addItems(["Financial Position", "Financial Performance"])
        self.type_combo.currentIndexChanged.connect(self._on_type_changed)
        form.addRow("Statement:", self.type_combo)

        self.name_combo = QComboBox()
        form.addRow("Name:", self.name_combo)

        self.biz_type_combo = QComboBox()
        self.biz_type_combo.addItems(["Sole Proprietorship", "Partnership", "Corporation"])
        form.addRow("Business Type:", self.biz_type_combo)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)

        # Position date
        self._pos_widget = _make_qwidget()
        pos_form = QFormLayout(self._pos_widget)
        pos_form.setContentsMargins(0, 0, 0, 0)
        pos_form.setLabelAlignment(Qt.AlignRight)
        self.as_of_date = QDateEdit()
        self.as_of_date.setCalendarPopup(True)
        self.as_of_date.setDisplayFormat("MM/dd/yyyy")
        self.as_of_date.setDate(QDate.currentDate())
        pos_form.addRow("As of:", self.as_of_date)

        # Performance date range
        self._perf_widget = _make_qwidget()
        perf_form = QFormLayout(self._perf_widget)
        perf_form.setContentsMargins(0, 0, 0, 0)
        perf_form.setLabelAlignment(Qt.AlignRight)

        current_year = QDate.currentDate().year()
        years = [str(y) for y in range(current_year - 5, current_year + 3)]

        from_row = QHBoxLayout()
        self.from_year_combo = QComboBox()
        self.from_year_combo.addItems(years)
        self.from_year_combo.setCurrentText(str(current_year))
        self.from_year_combo.currentIndexChanged.connect(self._clamp_to_date)
        self.from_month_combo = QComboBox()
        self.from_month_combo.addItems(MONTHS)
        self.from_month_combo.setCurrentIndex(0)
        self.from_month_combo.currentIndexChanged.connect(self._clamp_to_date)
        from_row.addWidget(self.from_year_combo)
        from_row.addWidget(self.from_month_combo)
        perf_form.addRow("From:", from_row)

        to_row = QHBoxLayout()
        self.to_year_combo = QComboBox()
        self.to_year_combo.addItems(years)
        self.to_year_combo.setCurrentText(str(current_year))
        self.to_year_combo.currentIndexChanged.connect(self._clamp_to_date)
        self.to_month_combo = QComboBox()
        self.to_month_combo.addItems(MONTHS)
        self.to_month_combo.setCurrentIndex(QDate.currentDate().month() - 1)
        self.to_month_combo.currentIndexChanged.connect(self._clamp_to_date)
        to_row.addWidget(self.to_year_combo)
        to_row.addWidget(self.to_month_combo)
        perf_form.addRow("To:", to_row)

        root.addLayout(form)
        root.addWidget(sep)
        root.addWidget(self._pos_widget)
        root.addWidget(self._perf_widget)
        root.addStretch()

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("Generate")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

        self.setLayout(root)
        self._on_type_changed(0)

    def _on_type_changed(self, idx):
        is_pos = (idx == 0)
        self.name_combo.blockSignals(True)
        self.name_combo.clear()
        self.name_combo.addItems(
            self._POSITION_NAMES if is_pos else self._PERFORMANCE_NAMES)
        self.name_combo.blockSignals(False)
        self._pos_widget.setVisible(is_pos)
        self._perf_widget.setVisible(not is_pos)

    def _clamp_to_date(self):
        from_year  = int(self.from_year_combo.currentText())
        from_month = self.from_month_combo.currentIndex() + 1
        to_year    = int(self.to_year_combo.currentText())
        to_month   = self.to_month_combo.currentIndex() + 1
        if to_year * 12 + to_month < from_year * 12 + from_month:
            self.to_year_combo.blockSignals(True)
            self.to_month_combo.blockSignals(True)
            self.to_year_combo.setCurrentText(str(from_year))
            self.to_month_combo.setCurrentIndex(from_month - 1)
            self.to_year_combo.blockSignals(False)
            self.to_month_combo.blockSignals(False)

    def get_params(self) -> dict:
        is_pos = self.type_combo.currentIndex() == 0
        params = {
            'type':          'position' if is_pos else 'performance',
            'name':          self.name_combo.currentText(),
            'company_name':  self.company_name,
            'business_type': self.biz_type_combo.currentText(),
        }
        if is_pos:
            params['as_of_date'] = self.as_of_date.date()
        else:
            from_year  = int(self.from_year_combo.currentText())
            from_month = self.from_month_combo.currentIndex() + 1
            to_year    = int(self.to_year_combo.currentText())
            to_month   = self.to_month_combo.currentIndex() + 1
            params['from_date'] = QDate(from_year, from_month, 1)
            params['to_date']   = QDate(to_year, to_month,
                                        month_last_day(to_year, to_month))
        return params


def _make_qwidget():
    """Tiny helper — avoids importing QWidget at module level for the dummy."""
    from PySide6.QtWidgets import QWidget
    return QWidget()


# ---------------------------------------------------------------------------
# PrintPreviewDialog
# ---------------------------------------------------------------------------

class PrintPreviewDialog(QDialog):
    def __init__(self, statement_text: str, title: str,
                 structured: dict = None, parent=None):
        super().__init__(parent)
        self._structured = structured or {}
        self.setWindowTitle(title)
        self.setModal(True)
        self.resize(860, 720)

        root = QVBoxLayout()

        self.display = QTextEdit()
        self.display.setReadOnly(True)
        self.display.setStyleSheet(
            "QTextEdit { font-family: 'Courier New'; font-size: 10pt; }")
        self.display.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.display.document().setDefaultStyleSheet("")
        self.display.setPlainText(statement_text)
        root.addWidget(self.display)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.print_btn = QPushButton("🖨  Print")
        self.print_btn.setFixedHeight(34)
        self.print_btn.clicked.connect(self._print)
        self.pdf_btn = QPushButton("📄  Print as PDF")
        self.pdf_btn.setFixedHeight(34)
        self.pdf_btn.clicked.connect(self._print_pdf)
        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(34)
        close_btn.clicked.connect(self.reject)
        btn_row.addWidget(self.print_btn)
        btn_row.addWidget(self.pdf_btn)
        btn_row.addWidget(close_btn)
        root.addLayout(btn_row)

        self.setLayout(root)

    def _print(self):
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageLayout(QPageLayout(
            QPageSize(QPageSize.PageSizeId.Letter),
            QPageLayout.Orientation.Portrait,
            QMarginsF(10, 10, 10, 10),
        ))
        dlg = QPrintDialog(printer, self)
        if dlg.exec() == QPrintDialog.DialogCode.Accepted:
            self.display.print_(printer)

    def _print_pdf(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save as PDF",
            os.path.join(get_io_dir("Financial Statements"), "financial_statement.pdf"),
            "PDF Files (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            if self._structured:
                save_pdf(self._structured, path, self.windowTitle())
            else:
                raise RuntimeError("No structured data available for PDF export.")
            QMessageBox.information(self, "PDF Saved", f"Saved to:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "PDF Error", str(exc))