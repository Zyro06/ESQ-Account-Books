"""
ui/widgets/dashboard_widget.py
-------------------------------
Dashboard — fully wired to real database queries.

Cards shown:
    Total Sales       (current month — 4xxx credits from SJ + CRJ + GJ)
    Total Expenses    (current month — 5–9xxx debits from PJ + CDJ + GJ)
    Cash Balance      (all-time trial balance for 1010 + 1020)
    Accounts Receivable (all-time trial balance for 1110)
    Accounts Payable  (all-time trial balance for 2010)
    Net Income        (current month — Sales minus Expenses)
    Recent Transactions (last 10 across all journals)
"""

from __future__ import annotations

from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QFrame, QTableWidget, QTableWidgetItem,
    QHeaderView, QSizePolicy, QPushButton,
)
from PySide6.QtCore import Qt
from PySide6.QtGui  import QColor, QPalette
import qtawesome as qta

from database.db_manager import DatabaseManager, _numeric_suffix


# ---------------------------------------------------------------------------
# Theme-aware colour helpers
# ---------------------------------------------------------------------------

def _theme_color(widget: QWidget, role: QPalette.ColorRole) -> str:
    """Return a hex colour string from the widget's current palette."""
    c = widget.palette().color(role)
    return c.name()


# ---------------------------------------------------------------------------
# Stat card widget
# ---------------------------------------------------------------------------

class _StatCard(QWidget):
    """A single summary card showing an icon, label, and value."""

    def __init__(
        self,
        icon_name:  str,
        label:      str,
        value:      str = '0.00',
        icon_color: str = '#1565C0',
        parent=None,
    ):
        super().__init__(parent)
        self._icon_name  = icon_name
        self._icon_color = icon_color
        self._label_text = label
        self._value_label: QLabel

        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(110)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(14)

        # Icon circle
        self._icon_wrap = QWidget()
        self._icon_wrap.setFixedSize(48, 48)
        self._icon_wrap.setStyleSheet(
            f'background-color: {icon_color}33; border-radius: 24px;')
        icon_lay = QVBoxLayout(self._icon_wrap)
        icon_lay.setContentsMargins(0, 0, 0, 0)
        self._icon_lbl = QLabel()
        self._icon_lbl.setAlignment(Qt.AlignCenter)
        pix = qta.icon(icon_name, color=icon_color).pixmap(24, 24)
        self._icon_lbl.setPixmap(pix)
        icon_lay.addWidget(self._icon_lbl)
        layout.addWidget(self._icon_wrap)

        # Text
        text_col = QVBoxLayout()
        text_col.setSpacing(4)

        self._lbl = QLabel(label)
        text_col.addWidget(self._lbl)

        self._value_label = QLabel(value)
        text_col.addWidget(self._value_label)

        layout.addLayout(text_col, stretch=1)

        # Apply light theme styles immediately so bold/size are correct before
        # set_theme() is called — qt_material would override a plain setFont()
        self.apply_theme(is_dark=False)

    def set_value(self, value: str):
        self._value_label.setText(value)

    def apply_theme(self, is_dark: bool):
        """Call whenever the app theme changes to repaint card colours."""
        if is_dark:
            card_bg      = '#1E1E2E'
            card_border  = '#2E2E3E'
            label_color  = '#AAAAAA'
            value_color  = '#EEEEEE'
        else:
            card_bg      = '#FFFFFF'
            card_border  = '#E0E0E0'
            label_color  = '#888888'
            value_color  = '#212121'

        self.setStyleSheet(
            f'_StatCard {{ background-color: {card_bg}; '
            f'border-radius: 10px; '
            f'border: 1px solid {card_border}; }}'
        )
        self._lbl.setStyleSheet(f'color: {label_color}; font-size: 9pt;')
        self._value_label.setStyleSheet(
            f'font-size: 14pt; font-weight: bold; color: {value_color};')
        self._icon_wrap.setStyleSheet(
            f'background-color: {self._icon_color}33; border-radius: 24px;')


# ---------------------------------------------------------------------------
# Dashboard widget
# ---------------------------------------------------------------------------

class DashboardWidget(QWidget):

    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self._is_dark   = False
        self._setup_ui()
        self._apply_theme()   # ← apply light styles before first paint
        self.load_data()

    # ------------------------------------------------------------------
    # Theme support
    # ------------------------------------------------------------------

    def set_theme(self, mode: str):
        """Called by main_window whenever the app theme toggles."""
        self._is_dark = (mode == 'dark')
        self._apply_theme()

    def _apply_theme(self):
        dark = self._is_dark

        # Page background
        bg   = '#16162A' if dark else '#F5F7FA'
        self.setStyleSheet(f'DashboardWidget {{ background-color: {bg}; }}')

        # Title / subtitle
        title_color    = '#EEEEEE' if dark else '#212121'
        subtitle_color = '#AAAAAA' if dark else '#888888'
        period_color   = '#AAAAAA' if dark else '#888888'
        self._title_lbl.setStyleSheet(
            f'font-size: 25pt; font-weight: bold; color: {title_color};')
        self._subtitle_lbl.setStyleSheet(
            f'color: {subtitle_color}; font-size: 10pt;')
        self._period_label.setStyleSheet(
            f'color: {period_color}; font-size: 11pt;')

        # Separator
        sep_color = '#2E2E3E' if dark else '#E0E0E0'
        self._sep.setStyleSheet(f'background-color: {sep_color};')

        # Recent label + count
        recent_color = '#EEEEEE' if dark else '#212121'
        count_color  = '#AAAAAA' if dark else '#888888'
        self._recent_lbl.setStyleSheet(
            f'font-size: 12pt; font-weight: bold; color: {recent_color};')
        self._recent_count_lbl.setStyleSheet(
            f'color: {count_color}; font-size: 9pt;')

        # Table
        if dark:
            tbl_style = (
                'QTableWidget { border: 1px solid #2E2E3E; border-radius: 8px; '
                '  background-color: #1E1E2E; color: #CCCCCC; }'
                'QHeaderView::section { background-color: #2A2A3E; color: #CCCCCC; '
                '  border: none; padding: 4px; }'
            )
        else:
            tbl_style = (
                'QTableWidget { border: 1px solid #E0E0E0; border-radius: 8px; }'
            )
        self._recent_table.setStyleSheet(tbl_style)

        # Refresh button
        if dark:
            btn_style = (
                'QPushButton { border: 1px solid #2E2E3E; border-radius: 4px; '
                '  background: #1E1E2E; }'
                'QPushButton:hover { background: #2A2A3E; }'
            )
        else:
            btn_style = (
                'QPushButton { border: 1px solid #ddd; border-radius: 4px; '
                '  background: white; }'
                'QPushButton:hover { background: #f0f0f0; }'
            )
        self._refresh_btn.setStyleSheet(btn_style)
        icon_color = '#AAAAAA' if dark else '#555555'
        self._refresh_btn.setIcon(qta.icon('fa5s.sync-alt', color=icon_color))

        # Cards
        for card in self._all_cards:
            card.apply_theme(dark)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(20)

        # ── Page title row ────────────────────────────────────────────
        title_row = QHBoxLayout()

        self._title_lbl = QLabel('Dashboard')
        self._title_lbl.setStyleSheet('font-size: 25pt; font-weight: bold; color: #212121;')
        title_row.addWidget(self._title_lbl)
        title_row.addStretch()

        now = datetime.now()
        self._period_label = QLabel(now.strftime('%B %Y'))
        self._period_label.setStyleSheet('color: #888888; font-size: 11pt;')
        title_row.addWidget(self._period_label)

        self._refresh_btn = QPushButton()
        self._refresh_btn.setIcon(qta.icon('fa5s.sync-alt', color='#555'))
        self._refresh_btn.setToolTip('Refresh dashboard')
        self._refresh_btn.setFixedSize(32, 32)
        self._refresh_btn.setStyleSheet(
            'QPushButton { border: 1px solid #ddd; border-radius: 4px; background: white; }'
            'QPushButton:hover { background: #f0f0f0; }')
        self._refresh_btn.clicked.connect(self.load_data)
        title_row.addWidget(self._refresh_btn)

        root.addLayout(title_row)

        self._subtitle_lbl = QLabel('Overview of your accounting data')
        self._subtitle_lbl.setStyleSheet('color: #888888; font-size: 10pt;')
        root.addWidget(self._subtitle_lbl)

        self._sep = QFrame()
        self._sep.setFrameShape(QFrame.HLine)
        self._sep.setFixedHeight(1)
        self._sep.setStyleSheet('background-color: #E0E0E0;')
        root.addWidget(self._sep)

        # ── Stat cards grid (2 rows × 3 cols) ─────────────────────────
        grid = QGridLayout()
        grid.setSpacing(16)

        self._card_sales    = _StatCard('fa5s.shopping-cart',      'Total Sales (This Month)',    icon_color='#1565C0')
        self._card_expenses = _StatCard('fa5s.receipt',            'Total Expenses (This Month)', icon_color='#C62828')
        self._card_cash     = _StatCard('fa5s.coins',              'Cash Balance',                icon_color='#2E7D32')
        self._card_ar       = _StatCard('fa5s.file-invoice-dollar','Accounts Receivable',         icon_color='#E65100')
        self._card_ap       = _StatCard('fa5s.hand-holding-usd',   'Accounts Payable',            icon_color='#6A1B9A')
        self._card_ni       = _StatCard('fa5s.chart-line',         'Net Income (This Month)',     icon_color='#00838F')

        self._all_cards = [
            self._card_sales, self._card_expenses, self._card_cash,
            self._card_ar,    self._card_ap,        self._card_ni,
        ]

        grid.addWidget(self._card_sales,    0, 0)
        grid.addWidget(self._card_expenses, 0, 1)
        grid.addWidget(self._card_cash,     0, 2)
        grid.addWidget(self._card_ar,       1, 0)
        grid.addWidget(self._card_ap,       1, 1)
        grid.addWidget(self._card_ni,       1, 2)
        root.addLayout(grid)

        # ── Recent transactions ────────────────────────────────────────
        recent_row = QHBoxLayout()
        self._recent_lbl = QLabel('Recent Transactions')
        self._recent_lbl.setStyleSheet(
            'font-size: 12pt; font-weight: bold; color: #212121; margin-top: 8px;')
        recent_row.addWidget(self._recent_lbl)
        recent_row.addStretch()
        self._recent_count_lbl = QLabel('')
        self._recent_count_lbl.setStyleSheet('color: #888; font-size: 9pt;')
        recent_row.addWidget(self._recent_count_lbl)
        root.addLayout(recent_row)

        self._recent_table = QTableWidget()
        self._recent_table.setColumnCount(5)
        self._recent_table.setHorizontalHeaderLabels(
            ['Date', 'Reference', 'Description', 'Journal', 'Amount'])
        self._recent_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._recent_table.setAlternatingRowColors(True)
        self._recent_table.verticalHeader().setVisible(False)
        self._recent_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._recent_table.setMaximumHeight(260)

        hdr = self._recent_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.Fixed)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.Fixed)
        hdr.setSectionResizeMode(4, QHeaderView.Fixed)
        self._recent_table.setColumnWidth(0, 100)
        self._recent_table.setColumnWidth(1, 140)
        self._recent_table.setColumnWidth(3, 120)
        self._recent_table.setColumnWidth(4, 120)

        self._recent_table.setStyleSheet(
            'QTableWidget { border: 1px solid #E0E0E0; border-radius: 8px; }')
        root.addWidget(self._recent_table)
        root.addStretch()

    # ------------------------------------------------------------------
    # Data loading
    # ------------------------------------------------------------------

    def load_data(self):
        now = datetime.now()
        self._period_label.setText(now.strftime('%B %Y'))

        sales    = self._query_sales(now)
        expenses = self._query_expenses(now)

        self._card_sales.set_value(   sales)
        self._card_expenses.set_value(expenses)
        self._card_cash.set_value(    self._query_cash())
        self._card_ar.set_value(      self._query_account_balance('1110'))
        self._card_ap.set_value(      self._query_account_balance('2010'))
        self._card_ni.set_value(      self._query_net_income(now))
        self._populate_recent(        self._query_recent(10))

    # ------------------------------------------------------------------
    # Query helpers
    # ------------------------------------------------------------------

    def _ym(self, now: datetime) -> str:
        return now.strftime('%Y-%m')

    # ── Total Sales ────────────────────────────────────────────────────
    # 4xxx account CREDITS this month from:
    #   • sales_journal_lines  (SJ)
    #   • cash_receipts_journal where account is 4xxx  (CRJ cash sales)
    #   • general_journal where account is 4xxx  (GJ adjustments)
    def _query_sales(self, now: datetime) -> str:
        try:
            conn   = self.db_manager.get_connection()
            cursor = conn.cursor()
            ym     = self._ym(now)
            total  = 0.0

            # SJ — sum credits on 4xxx lines
            cursor.execute("""
                SELECT COALESCE(SUM(sjl.credit), 0)
                FROM sales_journal sj
                JOIN sales_journal_lines sjl ON sjl.journal_id = sj.id
                WHERE substr(sj.date,7,4)||'-'||substr(sj.date,1,2) = ?
                  AND sjl.account_code IN (
                      SELECT account_code FROM chart_of_accounts
                  )
            """, (ym,))
            # Filter to 4xxx in Python (avoids complex SQL suffix logic)
            cursor.execute("""
                SELECT sjl.account_code, COALESCE(SUM(sjl.credit), 0)
                FROM sales_journal sj
                JOIN sales_journal_lines sjl ON sjl.journal_id = sj.id
                WHERE substr(sj.date,7,4)||'-'||substr(sj.date,1,2) = ?
                GROUP BY sjl.account_code
            """, (ym,))
            for code, amt in cursor.fetchall():
                if _numeric_suffix(code) and _numeric_suffix(code)[0] == '4':
                    total += amt or 0.0

            # CRJ — credits on 4xxx accounts
            cursor.execute("""
                SELECT account_code, COALESCE(SUM(credit), 0)
                FROM cash_receipts_journal
                WHERE substr(date,7,4)||'-'||substr(date,1,2) = ?
                GROUP BY account_code
            """, (ym,))
            for code, amt in cursor.fetchall():
                if _numeric_suffix(code) and _numeric_suffix(code)[0] == '4':
                    total += amt or 0.0

            # GJ — credits on 4xxx accounts
            cursor.execute("""
                SELECT account_code, COALESCE(SUM(credit), 0)
                FROM general_journal
                WHERE substr(date,7,4)||'-'||substr(date,1,2) = ?
                GROUP BY account_code
            """, (ym,))
            for code, amt in cursor.fetchall():
                if _numeric_suffix(code) and _numeric_suffix(code)[0] == '4':
                    total += amt or 0.0

            return f'{total:,.2f}'
        except Exception:
            return '0.00'

    # ── Total Expenses ─────────────────────────────────────────────────
    # 5–9xxx account DEBITS this month from:
    #   • purchase_journal_lines  (PJ)
    #   • cash_disbursement_journal where account is 5–9xxx  (CDJ)
    #   • general_journal where account is 5–9xxx  (GJ)
    def _query_expenses(self, now: datetime) -> str:
        try:
            conn   = self.db_manager.get_connection()
            cursor = conn.cursor()
            ym     = self._ym(now)
            total  = 0.0
            _exp   = set('56789')

            # PJ — debits on 5–9xxx lines
            cursor.execute("""
                SELECT pjl.account_code, COALESCE(SUM(pjl.debit), 0)
                FROM purchase_journal pj
                JOIN purchase_journal_lines pjl ON pjl.journal_id = pj.id
                WHERE substr(pj.date,7,4)||'-'||substr(pj.date,1,2) = ?
                GROUP BY pjl.account_code
            """, (ym,))
            for code, amt in cursor.fetchall():
                s = _numeric_suffix(code)
                if s and s[0] in _exp:
                    total += amt or 0.0

            # CDJ — debits on 5–9xxx accounts
            cursor.execute("""
                SELECT account_code, COALESCE(SUM(debit), 0)
                FROM cash_disbursement_journal
                WHERE substr(date,7,4)||'-'||substr(date,1,2) = ?
                GROUP BY account_code
            """, (ym,))
            for code, amt in cursor.fetchall():
                s = _numeric_suffix(code)
                if s and s[0] in _exp:
                    total += amt or 0.0

            # GJ — debits on 5–9xxx accounts
            cursor.execute("""
                SELECT account_code, COALESCE(SUM(debit), 0)
                FROM general_journal
                WHERE substr(date,7,4)||'-'||substr(date,1,2) = ?
                GROUP BY account_code
            """, (ym,))
            for code, amt in cursor.fetchall():
                s = _numeric_suffix(code)
                if s and s[0] in _exp:
                    total += amt or 0.0

            return f'{total:,.2f}'
        except Exception:
            return '0.00'

    # ── Cash Balance ───────────────────────────────────────────────────
    def _query_cash(self) -> str:
        try:
            conn   = self.db_manager.get_connection()
            cursor = conn.cursor()
            total  = 0.0
            for suffix in ('1010', '1020'):
                cursor.execute('SELECT account_code FROM chart_of_accounts')
                for (code,) in cursor.fetchall():
                    if _numeric_suffix(code) == suffix:
                        ledger = self.db_manager.get_general_ledger(code)
                        for e in ledger['entries']:
                            total += float(e.get('debit',  0) or 0)
                            total -= float(e.get('credit', 0) or 0)
                        break
            return f'{total:,.2f}'
        except Exception:
            return '0.00'

    # ── AR / AP balance ────────────────────────────────────────────────
    def _query_account_balance(self, suffix: str) -> str:
        try:
            conn   = self.db_manager.get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT account_code, normal_balance FROM chart_of_accounts')
            for code, nb in cursor.fetchall():
                if _numeric_suffix(code) == suffix:
                    ledger  = self.db_manager.get_general_ledger(code)
                    entries = ledger['entries']
                    td = sum(float(e.get('debit',  0) or 0) for e in entries)
                    tc = sum(float(e.get('credit', 0) or 0) for e in entries)
                    balance = (tc - td) if nb == 'Credit' else (td - tc)
                    return f'{abs(balance):,.2f}'
            return '0.00'
        except Exception:
            return '0.00'

    # ── Net Income ─────────────────────────────────────────────────────
    def _query_net_income(self, now: datetime) -> str:
        try:
            # Reuse the same definitions as Sales and Expenses cards
            sales_str    = self._query_sales(now)
            expenses_str = self._query_expenses(now)
            sales    = float(sales_str.replace(',', ''))
            expenses = float(expenses_str.replace(',', ''))
            return f'{sales - expenses:,.2f}'
        except Exception:
            return '0.00'

    # ── Recent transactions ────────────────────────────────────────────
    def _query_recent(self, limit: int = 10) -> list[dict]:
        try:
            conn   = self.db_manager.get_connection()
            cursor = conn.cursor()
            rows   = []

            cursor.execute("""
                SELECT sj.date, sj.reference_no, sj.customer_name, 'SJ',
                       COALESCE(SUM(CASE WHEN sjl.debit > 0 THEN sjl.debit ELSE 0 END), 0)
                FROM sales_journal sj
                JOIN sales_journal_lines sjl ON sjl.journal_id = sj.id
                GROUP BY sj.id
                ORDER BY sj.date DESC, sj.id DESC LIMIT ?
            """, (limit,))
            for r in cursor.fetchall():
                rows.append({'date': r[0], 'reference': r[1],
                             'description': r[2] or '', 'journal': r[3],
                             'amount': f'{r[4]:,.2f}'})

            cursor.execute("""
                SELECT pj.date, pj.reference_no, pj.payee_name, 'PJ',
                       COALESCE(SUM(CASE WHEN pjl.credit > 0 THEN pjl.credit ELSE 0 END), 0)
                FROM purchase_journal pj
                JOIN purchase_journal_lines pjl ON pjl.journal_id = pj.id
                GROUP BY pj.id
                ORDER BY pj.date DESC, pj.id DESC LIMIT ?
            """, (limit,))
            for r in cursor.fetchall():
                rows.append({'date': r[0], 'reference': r[1],
                             'description': r[2] or '', 'journal': r[3],
                             'amount': f'{r[4]:,.2f}'})

            for tbl, tag in [
                ('cash_disbursement_journal', 'CDJ'),
                ('cash_receipts_journal',     'CRJ'),
                ('general_journal',           'GJ'),
            ]:
                cursor.execute(f"""
                    SELECT date, reference_no, particulars, '{tag}',
                           COALESCE(SUM(debit), 0)
                    FROM {tbl}
                    GROUP BY date, reference_no
                    ORDER BY date DESC LIMIT ?
                """, (limit,))
                for r in cursor.fetchall():
                    rows.append({'date': r[0], 'reference': r[1],
                                 'description': r[2] or '', 'journal': r[3],
                                 'amount': f'{r[4]:,.2f}'})

            def _sort_key(r):
                try:
                    from datetime import datetime as _dt
                    return _dt.strptime(r.get('date', ''), '%m/%d/%Y').strftime('%Y-%m-%d')
                except Exception:
                    return r.get('date', '')

            rows.sort(key=_sort_key, reverse=True)
            return rows[:limit]

        except Exception:
            return []

    # ------------------------------------------------------------------
    # Table population
    # ------------------------------------------------------------------

    def _populate_recent(self, rows: list[dict]):
        # Light-mode tag colours — dark mode handled by table stylesheet
        TAG_COLORS_LIGHT = {
            'SJ':  '#E3F2FD',
            'PJ':  '#FFF3E0',
            'CDJ': '#FCE4EC',
            'CRJ': '#E8F5E9',
            'GJ':  '#F3E5F5',
        }
        TAG_COLORS_DARK = {
            'SJ':  '#1A3A5C',
            'PJ':  '#3E2A00',
            'CDJ': '#3E1A1A',
            'CRJ': '#1A3E1A',
            'GJ':  '#2A1A3E',
        }
        tag_colors = TAG_COLORS_DARK if self._is_dark else TAG_COLORS_LIGHT

        if not rows:
            self._recent_table.setRowCount(1)
            placeholder = QTableWidgetItem('No transactions found')
            placeholder.setTextAlignment(Qt.AlignCenter)
            placeholder.setFlags(Qt.ItemIsEnabled)
            placeholder.setForeground(QColor('#AAAAAA'))
            self._recent_table.setSpan(0, 0, 1, 5)
            self._recent_table.setItem(0, 0, placeholder)
            self._recent_count_lbl.setText('')
            return

        self._recent_table.clearSpans()
        self._recent_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            tag = row.get('journal', '')
            bg  = QColor(tag_colors.get(tag, '#FFFFFF' if not self._is_dark else '#1E1E2E'))

            for c, (key, right) in enumerate([
                ('date',        False),
                ('reference',   False),
                ('description', False),
                ('journal',     True),
                ('amount',      True),
            ]):
                item = QTableWidgetItem(str(row.get(key, '')))
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                if right:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if key == 'journal':
                    item.setBackground(bg)
                self._recent_table.setItem(r, c, item)

        self._recent_count_lbl.setText(f'Showing {len(rows)} most recent')