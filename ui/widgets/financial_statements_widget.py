"""
ui/widgets/financial_statements_widget.py
-----------------------------------------
Financial Statements tab — orchestration only.

Heavy logic lives in the fs/ sub-package:
    fs_pdf.py      — ReportLab PDF renderer
    fs_db.py       — saved_statements SQLite helpers
    fs_builders.py — position / performance builders (pure logic, no Qt)
    fs_dialogs.py  — GenerateDialog, PrintPreviewDialog
    fs_panels.py   — MetricCard, AnalysisPanel, ValidationPanel,
                     SavedStatementsPanel
"""

from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit,
    QGroupBox, QLineEdit, QDialog, QSplitter, QFileDialog, QMessageBox,
    QInputDialog,
)
from PySide6.QtCore import Qt
from PySide6.QtGui  import QFont, QKeySequence, QShortcut

from database.db_manager import DatabaseManager
from resources.file_paths import get_io_dir

from ui.widgets.fs.fs_pdf      import save_pdf
from ui.widgets.fs.fs_db       import ensure_table
from ui.widgets.fs.fs_builders import build_position, build_performance
from ui.widgets.fs.fs_dialogs  import GenerateDialog, PrintPreviewDialog
from ui.widgets.fs.fs_panels   import (
    AnalysisPanel, ValidationPanel, SavedStatementsPanel,
)


class FinancialStatementsWidget(QWidget):

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager          = db_manager
        self._current_text       = ""
        self._current_title      = ""
        self._current_structured: dict = {}
        self._current_tb:         list = []
        self._current_params:     dict = {}

        ensure_table(db_manager)

        self._setup_ui()
        self._setup_shortcuts()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        title = QLabel("FINANCIAL STATEMENTS")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignCenter)
        root.addWidget(title)

        # ── Company bar ───────────────────────────────────────────────
        cfg_group  = QGroupBox("Company")
        cfg_layout = QHBoxLayout()
        cfg_layout.addWidget(QLabel("Company Name:"))
        self.company_name_input = QLineEdit("ABC COMPANY")
        self.company_name_input.setMaximumWidth(260)
        cfg_layout.addWidget(self.company_name_input)
        cfg_layout.addStretch()

        self._save_btn = QPushButton("💾  Save Statement")
        self._save_btn.setFixedHeight(34)
        self._save_btn.setEnabled(False)
        self._save_btn.clicked.connect(self._save_current)
        cfg_layout.addWidget(self._save_btn)

        self.generate_btn = QPushButton("⚙  Generate Statement…")
        self.generate_btn.setFixedHeight(34)
        self.generate_btn.clicked.connect(self._open_generate_dialog)
        cfg_layout.addWidget(self.generate_btn)
        cfg_group.setLayout(cfg_layout)
        root.addWidget(cfg_group)

        # ── Horizontal splitter: saved | display | analysis ───────────
        h_splitter = QSplitter(Qt.Horizontal)
        h_splitter.setChildrenCollapsible(False)

        self._saved_panel = SavedStatementsPanel(self.db_manager)
        self._saved_panel.on_load_callback = self._load_saved
        h_splitter.addWidget(self._saved_panel)

        centre_widget = QWidget()
        centre_layout = QVBoxLayout(centre_widget)
        centre_layout.setContentsMargins(0, 0, 0, 0)
        centre_layout.setSpacing(4)

        self.display = QTextEdit()
        self.display.setReadOnly(True)
        self.display.setStyleSheet(
            "QTextEdit { font-family: 'Courier New'; font-size: 10pt; }")
        self.display.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self.display.document().setDefaultStyleSheet("")
        self.display.setPlaceholderText(
            "Click  ⚙  Generate Statement…  to produce a financial statement.")
        centre_layout.addWidget(self.display)

        bottom_row = QHBoxLayout()
        bottom_row.addStretch()
        self.print_btn = QPushButton("🖨  Print")
        self.print_btn.setFixedHeight(32)
        self.print_btn.setEnabled(False)
        self.print_btn.clicked.connect(self._print_current)
        self.pdf_btn = QPushButton("📄  Print as PDF")
        self.pdf_btn.setFixedHeight(32)
        self.pdf_btn.setEnabled(False)
        self.pdf_btn.clicked.connect(self._pdf_current)
        bottom_row.addWidget(self.print_btn)
        bottom_row.addWidget(self.pdf_btn)
        centre_layout.addLayout(bottom_row)

        h_splitter.addWidget(centre_widget)

        self._analysis_panel = AnalysisPanel()
        h_splitter.addWidget(self._analysis_panel)

        h_splitter.setStretchFactor(0, 1)
        h_splitter.setStretchFactor(1, 4)
        h_splitter.setStretchFactor(2, 2)

        # ── Vertical splitter: main | validation ──────────────────────
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setChildrenCollapsible(False)
        v_splitter.addWidget(h_splitter)

        self._validation_panel = ValidationPanel()
        v_splitter.addWidget(self._validation_panel)

        v_splitter.setStretchFactor(0, 5)
        v_splitter.setStretchFactor(1, 1)

        root.addWidget(v_splitter)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+G"), self).activated.connect(
            self._open_generate_dialog)
        QShortcut(QKeySequence("Ctrl+P"), self).activated.connect(
            self._print_current)
        QShortcut(QKeySequence("Ctrl+Shift+P"), self).activated.connect(
            self._pdf_current)
        QShortcut(QKeySequence("Ctrl+S"), self).activated.connect(
            self._save_current)

    # ------------------------------------------------------------------
    # Generation flow
    # ------------------------------------------------------------------

    def _open_generate_dialog(self):
        company = self.company_name_input.text().strip().upper() or "ABC COMPANY"
        dlg = GenerateDialog(company, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._generate(dlg.get_params())

    def _generate(self, params: dict):
        ptype = params['type']

        if ptype == 'position':
            as_of = params['as_of_date']
            tb    = self.db_manager.get_trial_balance(
                        date_to=as_of.toString("MM/dd/yyyy"))
            analysis_data, text, structured = build_position(
                tb, params['name'], params['company_name'],
                as_of, params.get('business_type', 'Sole Proprietorship'))
            self._analysis_panel.analyze_position(analysis_data)
        else:
            from_d = params['from_date']
            to_d   = params['to_date']
            tb     = self.db_manager.get_trial_balance(
                         date_from=from_d.toString("MM/dd/yyyy"),
                         date_to=to_d.toString("MM/dd/yyyy"))
            analysis_data, text, structured = build_performance(
                tb, params['name'], params['company_name'], from_d, to_d)
            self._analysis_panel.analyze_performance(analysis_data)

        self.display.setPlainText(text)
        self.print_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)
        self._save_btn.setEnabled(True)
        self._current_text       = text
        self._current_title      = params['name']
        self._current_structured = structured
        self._current_tb         = tb
        self._current_params     = params

        if ptype == 'position':
            self._validation_panel.run_validation(
                tb, 'position',
                total_assets      = analysis_data.get('total_assets', 0),
                total_liabilities = analysis_data.get('total_liabilities', 0),
                total_equity      = analysis_data.get('total_equity', 0),
                net_income        = analysis_data.get('net_income', 0),
            )
        else:
            self._validation_panel.run_validation(
                tb, 'performance',
                total_revenue  = analysis_data.get('total_revenue', 0),
                total_expenses = analysis_data.get('total_expenses', 0),
                net_income     = analysis_data.get('net_income', 0),
            )

    # ------------------------------------------------------------------
    # Save / load
    # ------------------------------------------------------------------

    def _save_current(self):
        if not self._current_text:
            return
        label, ok = QInputDialog.getText(
            self, "Save Statement", "Label for this statement:",
            text=self._current_title)
        if ok and label.strip():
            self._saved_panel.save_statement(
                label.strip(), self._current_text, self._current_params)

    def _load_saved(self, entry: dict):
        self.display.setPlainText(entry['text'])
        self._current_text   = entry['text']
        self._current_title  = entry['label']
        self._current_params = entry['params']
        self.print_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)
        self._save_btn.setEnabled(True)
        self._generate(entry['params'])

    # ------------------------------------------------------------------
    # Print / PDF
    # ------------------------------------------------------------------

    def _print_current(self):
        if not self.print_btn.isEnabled():
            return
        text  = self._current_text  or self.display.toPlainText()
        title = self._current_title or 'Financial Statement'
        PrintPreviewDialog(text, title, self._current_structured, self).exec()

    def _pdf_current(self):
        if not self.pdf_btn.isEnabled():
            return
        if not self._current_structured:
            QMessageBox.warning(self, "No Statement",
                                "Generate a statement first.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save as PDF",
            os.path.join(get_io_dir("Financial Statements"),
                         "financial_statement.pdf"),
            "PDF Files (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            save_pdf(self._current_structured, path,
                     self._current_title or 'Financial Statement')
            QMessageBox.information(self, "PDF Saved", f"Saved to:\n{path}")
        except Exception as exc:
            QMessageBox.critical(self, "PDF Error", str(exc))

    # ------------------------------------------------------------------
    # Required by main_window refresh loop
    # ------------------------------------------------------------------

    def load_data(self):
        pass