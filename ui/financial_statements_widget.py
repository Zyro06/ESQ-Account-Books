from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit,
    QDateEdit, QGroupBox, QLineEdit, QShortcut, QDialog, QFormLayout,
    QComboBox, QDialogButtonBox, QFrame, QSizePolicy, QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QKeySequence, QFont, QTextCursor
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from database.db_manager import DatabaseManager
import calendar

from resources.file_paths import get_io_dir


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

def _fmt_date(qdate: QDate) -> str:
    return f"{MONTHS[qdate.month() - 1]} {qdate.day()}, {qdate.year()}"

def _month_last_day(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]

def _centre(text: str, width: int) -> str:
    return text.center(width)

def _period_label(from_d: QDate, to_d: QDate) -> str:
    num_months = (to_d.year() - from_d.year()) * 12 + \
                 (to_d.month() - from_d.month()) + 1

    SMALL_WORDS = {
        1: "One",   2: "Two",   3: "Three", 4: "Four",
        5: "Five",  6: "Six",   7: "Seven", 8: "Eight",
        9: "Nine", 10: "Ten",  11: "Eleven",
    }
    YEAR_WORDS = {
        1: "the Year", 2: "Two Years", 3: "Three Years",
        4: "Four Years", 5: "Five Years",
    }

    end_date = _fmt_date(to_d)

    if num_months == 1:
        span = "the Month"
    elif num_months % 12 == 0:
        years = num_months // 12
        span  = YEAR_WORDS.get(years, f"{years} Years")
    elif num_months > 12:
        span = f"{num_months} Months"
    else:
        span = f"{SMALL_WORDS[num_months]} Months"

    return f"For {span} Ending {end_date}"


# ---------------------------------------------------------------------------
# Generate dialog
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
        self._pos_widget = QWidget()
        pos_form = QFormLayout(self._pos_widget)
        pos_form.setContentsMargins(0, 0, 0, 0)
        pos_form.setLabelAlignment(Qt.AlignRight)
        self.as_of_date = QDateEdit()
        self.as_of_date.setCalendarPopup(True)
        self.as_of_date.setDisplayFormat("MM/dd/yyyy")
        self.as_of_date.setDate(QDate.currentDate())
        pos_form.addRow("As of:", self.as_of_date)

        # Performance date
        self._perf_widget = QWidget()
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

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Generate")
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
                                        _month_last_day(to_year, to_month))
        return params


# ---------------------------------------------------------------------------
# Print preview dialog
# ---------------------------------------------------------------------------

class PrintPreviewDialog(QDialog):
    def __init__(self, statement_text: str, title: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.resize(860, 720)

        root = QVBoxLayout()

        self.display = QTextEdit()
        self.display.setReadOnly(True)
        self.display.setFont(QFont("Courier New", 10))
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
        printer = QPrinter(QPrinter.HighResolution)
        dlg = QPrintDialog(printer, self)
        if dlg.exec_() == QPrintDialog.Accepted:
            self.display.print_(printer)

    def _print_pdf(self):
        import os
        path, _ = QFileDialog.getSaveFileName(
            self, "Save as PDF",
            os.path.join(get_io_dir("Financial Statements"), "financial_statement.pdf"),
            "PDF Files (*.pdf)"
        )
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(path)
        printer.setPageSize(QPrinter.Letter)
        printer.setPageMargins(4, 0, 0, 0, QPrinter.Millimeter)
        self.display.print_(printer)
        QMessageBox.information(self, "PDF Saved", f"Saved to:\n{path}")


# ---------------------------------------------------------------------------
# Main widget
# ---------------------------------------------------------------------------

class FinancialStatementsWidget(QWidget):

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self._setup_ui()
        self._setup_shortcuts()

    def _setup_ui(self):
        layout = QVBoxLayout()

        title = QLabel("FINANCIAL STATEMENTS")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        cfg_group  = QGroupBox("Company")
        cfg_layout = QHBoxLayout()
        cfg_layout.addWidget(QLabel("Company Name:"))
        self.company_name_input = QLineEdit("ABC COMPANY")
        self.company_name_input.setMaximumWidth(260)
        cfg_layout.addWidget(self.company_name_input)
        cfg_layout.addStretch()
        self.generate_btn = QPushButton("⚙  Generate Statement…")
        self.generate_btn.setFixedHeight(34)
        self.generate_btn.clicked.connect(self._open_generate_dialog)
        cfg_layout.addWidget(self.generate_btn)
        cfg_group.setLayout(cfg_layout)
        layout.addWidget(cfg_group)

        self.display = QTextEdit()
        self.display.setReadOnly(True)
        self.display.setFont(QFont("Courier New", 10))
        self.display.setPlaceholderText(
            "Click  ⚙  Generate Statement…  to produce a financial statement.")
        layout.addWidget(self.display)

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
        layout.addLayout(bottom_row)

        self.setLayout(layout)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+G"), self).activated.connect(
            self._open_generate_dialog)
        QShortcut(QKeySequence("Ctrl+P"), self).activated.connect(
            self._print_current)
        QShortcut(QKeySequence("Ctrl+Shift+P"), self).activated.connect(
            self._pdf_current)

    # ------------------------------------------------------------------ flow

    def _open_generate_dialog(self):
        company = self.company_name_input.text().strip().upper() or "ABC COMPANY"
        dlg = GenerateDialog(company, self)
        if dlg.exec_() != QDialog.Accepted:
            return
        self._generate(dlg.get_params())

    def _generate(self, params: dict):
        ptype = params['type']
        if ptype == 'position':
            as_of   = params['as_of_date']
            tb      = self.db_manager.get_trial_balance(
                          date_to=as_of.toString("MM/dd/yyyy"))
            text    = self._build_position(
                          tb, params['name'], params['company_name'],
                          as_of, params.get('business_type', 'Sole Proprietorship'))
        else:
            from_d  = params['from_date']
            to_d    = params['to_date']
            tb      = self.db_manager.get_trial_balance(
                          date_from=from_d.toString("MM/dd/yyyy"),
                          date_to=to_d.toString("MM/dd/yyyy"))
            text    = self._build_performance(
                          tb, params['name'], params['company_name'],
                          from_d, to_d)

        self.display.setPlainText(text)
        self.print_btn.setEnabled(True)
        self.pdf_btn.setEnabled(True)
        self._current_text  = text
        self._current_title = params['name']

    # ------------------------------------------------------------------ builders

    def _build_position(self, trial_balance, stmt_name: str,
                        company: str, as_of: QDate,
                        business_type: str = 'Sole Proprietorship') -> str:
        """
        Build Statement of Financial Position / Balance Sheet.

        FIX 1: Uses numeric suffix (strips prefix like COA-, ABC-) so any
                account code format works correctly.
        FIX 2: Liabilities use -amount (signed) instead of abs(amount) so
                accounts with abnormal debit balances correctly reduce total
                liabilities instead of inflating them.
        """
        BIZ_EQUITY_LABELS = {
            'Sole Proprietorship': "Owner's Equity",
            'Partnership':        "Partners' Equity",
            'Corporation':        "Stockholders' Equity",
        }
        BIZ_NI_LABELS = {
            'Sole Proprietorship': "Net Income (Loss)",
            'Partnership':        "Net Income (Loss) - Current Period",
            'Corporation':        "Net Income (Loss) for the Period",
        }
        equity_label = BIZ_EQUITY_LABELS.get(business_type, "Equity")
        ni_label     = BIZ_NI_LABELS.get(business_type, "Net Income (Loss)")

        assets      = []   # (code, desc, signed_amount)
        liabilities = []   # (code, desc, signed_amount)
        equity_adds = []   # capital accounts
        equity_deds = []   # drawing accounts
        rev_total   = 0.0
        exp_total   = 0.0

        for e in trial_balance:
            code   = e['account_code']
            desc   = e['account_description']
            amount = e['amount']          # signed: total_debit - total_credit
            nb     = e.get('normal_balance', 'Debit')

            # Strip prefix (COA-, ABC-, or anything else) to get numeric part
            suffix = code.rsplit('-', 1)[-1] if '-' in code else code
            first  = suffix[0] if suffix else ''

            if first == '1':
                # Asset (debit-normal): positive amount = normal debit balance
                # Contra-asset (credit-normal, e.g. Accum Dep): negative amount = normal
                assets.append((code, desc, amount))

            elif first == '2':
                # Liability (credit-normal)
                # Normal credit balance  → amount < 0 → -amount > 0  (positive, shown as liability)
                # Abnormal debit balance → amount > 0 → -amount < 0  (negative, reduces total liab)
                liabilities.append((code, desc, -amount))

            elif first == '3':
                if nb == 'Debit':
                    equity_deds.append((code, desc, amount))    # drawings
                else:
                    equity_adds.append((code, desc, -amount))   # capital (credit-normal)

            elif first == '4':
                # Revenue (credit-normal): amount is negative in TB → use abs
                rev_total += abs(amount)

            elif first in ('5', '6', '7', '8', '9'):
                # Expenses / COGS / Tax (debit-normal): amount is positive
                exp_total += amount

        net_income        = rev_total - exp_total
        total_assets      = sum(a[2] for a in assets)
        total_liabilities = sum(a[2] for a in liabilities)
        total_equity_adds = sum(a[2] for a in equity_adds)
        total_equity_deds = sum(a[2] for a in equity_deds)
        total_equity      = total_equity_adds - total_equity_deds + net_income
        total_l_and_e     = total_liabilities + total_equity

        W = 72
        lines = [
            "=" * W,
            _centre(company, W),
            _centre(stmt_name.upper(), W),
            _centre(f"As of {_fmt_date(as_of)}", W),
            "=" * W, "",
            "ASSETS", "-" * W,
        ]
        for code, desc, amt in assets:
            lines.append(f"  {code:<14} {desc:<38} {amt:>14,.2f}")
        lines += [
            "-" * W,
            f"  {'TOTAL ASSETS':<52} {total_assets:>14,.2f}",
            "",
            f"LIABILITIES AND {equity_label.upper()}",
            "-" * W,
            "  Liabilities:",
        ]
        for code, desc, amt in liabilities:
            lines.append(f"    {code:<12} {desc:<38} {amt:>14,.2f}")
        lines += [
            "-" * W,
            f"  {'TOTAL LIABILITIES':<52} {total_liabilities:>14,.2f}",
            "",
            f"  {equity_label}:",
        ]
        for code, desc, amt in equity_adds:
            lines.append(f"    {code:<12} {desc:<38} {amt:>14,.2f}")
        for code, desc, amt in equity_deds:
            lines.append(f"    {code:<12} {desc:<38} {-amt:>14,.2f}")
        lines += [
            f"    {ni_label:<50} {net_income:>14,.2f}",
            "-" * W,
            f"  {'TOTAL ' + equity_label.upper():<52} {total_equity:>14,.2f}",
            "",
            f"  {'TOTAL LIABILITIES AND ' + equity_label.upper():<52} {total_l_and_e:>14,.2f}",
            "",
        ]

        diff = total_assets - total_l_and_e
        if abs(diff) > 0.005:
            lines += [
                "  " + "!" * (W - 2),
                f"  WARNING: Statement is OUT OF BALANCE by {diff:,.2f}",
                "  Check for missing journal entries or mis-coded accounts.",
                "  " + "!" * (W - 2), "",
            ]

        lines.append("=" * W)
        return "\n".join(lines)

    def _build_performance(self, trial_balance, stmt_name: str,
                           company: str, from_d: QDate, to_d: QDate) -> str:
        """
        Build Statement of Financial Performance / Income Statement.

        FIX: Uses numeric suffix (strips prefix like COA-, ABC-) so any
             account code format works correctly.
        """
        revenue  = []
        cogs     = []
        expenses = []

        for e in trial_balance:
            code   = e['account_code']
            desc   = e['account_description']
            amount = e['amount']

            # Strip prefix to get numeric part
            suffix = code.rsplit('-', 1)[-1] if '-' in code else code
            first  = suffix[0] if suffix else ''

            if first == '4':
                revenue.append((code, desc, abs(amount)))    # credit-normal
            elif first == '5':
                cogs.append((code, desc, amount))            # debit-normal
            elif first in ('6', '7', '8', '9'):
                expenses.append((code, desc, amount))        # debit-normal

        total_revenue  = sum(r[2] for r in revenue)
        total_cogs     = sum(c[2] for c in cogs)
        gross_profit   = total_revenue - total_cogs
        total_expenses = sum(ex[2] for ex in expenses)
        net_income     = gross_profit - total_expenses
        period_label   = _period_label(from_d, to_d)

        W = 72
        lines = [
            "=" * W,
            _centre(company, W),
            _centre(stmt_name.upper(), W),
            _centre(period_label, W),
            "=" * W, "",
            "REVENUE", "-" * W,
        ]
        for code, desc, amt in revenue:
            lines.append(f"  {code:<14} {desc:<38} {amt:>14,.2f}")
        lines += [
            "-" * W,
            f"  {'TOTAL REVENUE':<52} {total_revenue:>14,.2f}", "",
        ]

        if cogs:
            lines += ["COST OF GOODS SOLD", "-" * W]
            for code, desc, amt in cogs:
                lines.append(f"  {code:<14} {desc:<38} {amt:>14,.2f}")
            lines += [
                "-" * W,
                f"  {'TOTAL COST OF GOODS SOLD':<52} {total_cogs:>14,.2f}",
                "",
                f"  {'GROSS PROFIT':<52} {gross_profit:>14,.2f}", "",
            ]

        if expenses:
            lines += ["OPERATING EXPENSES", "-" * W]
            for code, desc, amt in expenses:
                lines.append(f"  {code:<14} {desc:<38} {amt:>14,.2f}")
            lines += [
                "-" * W,
                f"  {'TOTAL OPERATING EXPENSES':<52} {total_expenses:>14,.2f}", "",
            ]

        lines += [
            f"  {'NET INCOME (LOSS)':<52} {net_income:>14,.2f}",
            "", "=" * W,
        ]
        return "\n".join(lines)

    # ------------------------------------------------------------------ print

    def _print_current(self):
        if not self.print_btn.isEnabled():
            return
        text  = getattr(self, '_current_text',  self.display.toPlainText())
        title = getattr(self, '_current_title', 'Financial Statement')
        PrintPreviewDialog(text, title, self).exec_()

    def _pdf_current(self):
        if not self.pdf_btn.isEnabled():
            return
        import os
        text  = getattr(self, '_current_text',  self.display.toPlainText())
        title = getattr(self, '_current_title', 'Financial Statement')
        path, _ = QFileDialog.getSaveFileName(
            self, "Save as PDF",
            os.path.join(get_io_dir("Financial Statements"), "financial_statement.pdf"),
            "PDF Files (*.pdf)"
        )
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        tmp = QTextEdit()
        tmp.setFont(QFont("Courier New", 10))
        tmp.setPlainText(text)
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(path)
        printer.setPageSize(QPrinter.Letter)
        printer.setPageMargins(4, 0, 0, 0, QPrinter.Millimeter)
        tmp.print_(printer)
        QMessageBox.information(self, "PDF Saved", f"Saved to:\n{path}")

    def load_data(self):
        pass