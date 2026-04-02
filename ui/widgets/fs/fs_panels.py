"""
ui/widgets/fs/fs_panels.py
---------------------------
Side-panel widgets used inside FinancialStatementsWidget:
    MetricCard           — coloured KPI card
    AnalysisPanel        — right-hand ratio / metric display
    ValidationPanel      — bottom error / warning strip
    SavedStatementsPanel — left-hand DB-backed statement list
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame,
    QLabel, QPushButton, QListWidget, QListWidgetItem,
    QScrollArea, QMessageBox, QInputDialog, QSizePolicy,
)
from PySide6.QtCore import Qt, QSize

try:
    import qtawesome as qta
    _QTA_OK = True
except ImportError:
    _QTA_OK = False

from database.db_manager import DatabaseManager
import ui.widgets.fs.fs_db as _db


def _qta_icon(name: str, color: str = "#555555"):
    from PySide6.QtGui import QIcon
    if _QTA_OK:
        try:
            return qta.icon(name, color=color)
        except Exception:
            pass
    return QIcon()


# ---------------------------------------------------------------------------
# MetricCard
# ---------------------------------------------------------------------------

class MetricCard(QFrame):
    STATUS_STYLES = {
        'good':    ("background:#e6f9f0; border:1.5px solid #34c47c; border-radius:6px;",
                    "#1a7a45", "#34c47c"),
        'warn':    ("background:#fff8e6; border:1.5px solid #f5a623; border-radius:6px;",
                    "#8a5a00", "#f5a623"),
        'bad':     ("background:#fdecea; border:1.5px solid #e53935; border-radius:6px;",
                    "#b71c1c", "#e53935"),
        'neutral': ("background:#f0f4ff; border:1.5px solid #5c7cfa; border-radius:6px;",
                    "#1a3a8f", "#5c7cfa"),
    }
    ICONS = {'good': '✅', 'warn': '⚠️', 'bad': '🔴', 'neutral': 'ℹ️'}

    def __init__(self, label: str, value: str, note: str,
                 status: str = 'neutral', parent=None):
        super().__init__(parent)
        frame_style, text_color, _ = self.STATUS_STYLES.get(
            status, self.STATUS_STYLES['neutral'])
        self.setStyleSheet(f"MetricCard {{ {frame_style} }}")
        self.setContentsMargins(8, 6, 8, 6)

        lay = QVBoxLayout(self)
        lay.setSpacing(2)
        lay.setContentsMargins(6, 6, 6, 6)

        lbl_row = QHBoxLayout()
        icon_lbl = QLabel(self.ICONS.get(status, 'ℹ️'))
        icon_lbl.setFixedWidth(20)
        name_lbl = QLabel(label)
        name_lbl.setStyleSheet(
            f"font-size:10px; font-weight:600; color:{text_color};")
        lbl_row.addWidget(icon_lbl)
        lbl_row.addWidget(name_lbl)
        lbl_row.addStretch()
        lay.addLayout(lbl_row)

        val_lbl = QLabel(value)
        val_lbl.setStyleSheet(
            f"font-size:16px; font-weight:800; color:{text_color};")
        lay.addWidget(val_lbl)

        note_lbl = QLabel(note)
        note_lbl.setStyleSheet(
            f"font-size:9px; color:{text_color}; opacity:0.8;")
        note_lbl.setWordWrap(True)
        lay.addWidget(note_lbl)


# ---------------------------------------------------------------------------
# AnalysisPanel
# ---------------------------------------------------------------------------

class AnalysisPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumWidth(220)
        self.setMaximumWidth(300)

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 0, 4, 0)
        root.setSpacing(6)

        header = QLabel("📊  Smart Analysis")
        header.setStyleSheet(
            "font-size:12px; font-weight:700; color:#2d3a5e;"
            "padding:6px 4px; border-bottom:2px solid #5c7cfa;")
        root.addWidget(header)

        self._placeholder = QLabel(
            "Generate a statement to\nsee financial analysis here.")
        self._placeholder.setAlignment(Qt.AlignCenter)
        self._placeholder.setStyleSheet(
            "color:#aaa; font-size:11px; padding:20px 0;")
        root.addWidget(self._placeholder)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._cards_widget = QWidget()
        self._cards_layout = QVBoxLayout(self._cards_widget)
        self._cards_layout.setSpacing(6)
        self._cards_layout.setContentsMargins(0, 0, 0, 0)
        self._cards_layout.addStretch()
        self._scroll.setWidget(self._cards_widget)
        self._scroll.hide()
        root.addWidget(self._scroll)
        root.addStretch()

    def clear(self):
        while self._cards_layout.count() > 1:
            item = self._cards_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._scroll.hide()
        self._placeholder.show()

    def show_analysis(self, metrics: list):
        self.clear()
        self._placeholder.hide()
        for i, (label, value, note, status) in enumerate(metrics):
            card = MetricCard(label, value, note, status)
            self._cards_layout.insertWidget(i, card)
        self._scroll.show()

    def analyze_position(self, data: dict):
        total_assets      = data.get('total_assets', 0)
        total_liabilities = data.get('total_liabilities', 0)
        total_equity      = data.get('total_equity', 0)
        current_assets    = data.get('current_assets', total_assets)
        current_liab      = data.get('current_liabilities', total_liabilities)
        net_income        = data.get('net_income', 0)
        metrics = []

        if current_liab != 0:
            cr = current_assets / current_liab
            cr_status = 'good' if cr >= 2 else ('warn' if cr >= 1 else 'bad')
            cr_note   = ("Strong liquidity" if cr >= 2 else
                         "Adequate liquidity" if cr >= 1 else
                         "Liquidity risk — liabilities exceed assets")
            metrics.append(("Current Ratio", f"{cr:.2f}x", cr_note, cr_status))
        else:
            metrics.append(("Current Ratio", "N/A", "No current liabilities", "neutral"))

        if total_equity != 0:
            de = total_liabilities / total_equity
            de_status = 'good' if de < 1 else ('warn' if de < 2 else 'bad')
            de_note   = ("Low leverage — conservative financing" if de < 1 else
                         "Moderate leverage" if de < 2 else
                         "High leverage — debt-heavy structure")
            metrics.append(("Debt-to-Equity", f"{de:.2f}x", de_note, de_status))
        else:
            metrics.append(("Debt-to-Equity", "N/A", "Equity is zero", "warn"))

        if total_assets != 0:
            er = (total_equity / total_assets) * 100
            er_status = 'good' if er >= 50 else ('warn' if er >= 30 else 'bad')
            er_note   = ("Majority equity-financed" if er >= 50 else
                         "Mixed financing" if er >= 30 else "Majority debt-financed")
            metrics.append(("Equity Ratio", f"{er:.1f}%", er_note, er_status))

        wc = current_assets - current_liab
        wc_status = 'good' if wc > 0 else ('warn' if wc == 0 else 'bad')
        wc_note   = (f"Positive buffer of {wc:,.2f}" if wc > 0 else
                     "Break-even" if wc == 0 else f"Deficit of {abs(wc):,.2f}")
        metrics.append(("Working Capital", f"{wc:,.0f}", wc_note, wc_status))

        if total_equity != 0:
            roe = (net_income / total_equity) * 100
            roe_status = 'good' if roe >= 10 else ('warn' if roe >= 0 else 'bad')
            roe_note   = ("Strong returns" if roe >= 15 else
                          "Moderate returns" if roe >= 5 else
                          "Negative returns" if roe < 0 else "Low returns")
            metrics.append(("Return on Equity", f"{roe:.1f}%", roe_note, roe_status))

        diff       = total_assets - (total_liabilities + total_equity)
        bal_status = 'good' if abs(diff) < 0.01 else 'bad'
        bal_note   = ("Statement is balanced ✓" if abs(diff) < 0.01
                      else f"Out of balance by {diff:,.2f}")
        metrics.append(("Balance Check",
                        "OK" if abs(diff) < 0.01 else "FAIL",
                        bal_note, bal_status))
        self.show_analysis(metrics)

    def analyze_performance(self, data: dict):
        total_revenue  = data.get('total_revenue', 0)
        total_cogs     = data.get('total_cogs', 0)
        gross_profit   = data.get('gross_profit', 0)
        total_expenses = data.get('total_expenses', 0)
        net_income     = data.get('net_income', 0)
        metrics = []

        if total_revenue != 0:
            gm = (gross_profit / total_revenue) * 100
            gm_status = 'good' if gm >= 40 else ('warn' if gm >= 20 else 'bad')
            gm_note   = ("High-margin business" if gm >= 40 else
                         "Average margins" if gm >= 20 else "Thin margins — review COGS")
            metrics.append(("Gross Margin", f"{gm:.1f}%", gm_note, gm_status))
        else:
            metrics.append(("Gross Margin", "N/A", "No revenue recorded", "bad"))

        if total_revenue != 0:
            nm = (net_income / total_revenue) * 100
            nm_status = 'good' if nm >= 10 else ('warn' if nm >= 0 else 'bad')
            nm_note   = ("Healthy profitability" if nm >= 10 else
                         "Slim net profit" if nm >= 0 else
                         "Net loss — expenses exceed revenue")
            metrics.append(("Net Margin", f"{nm:.1f}%", nm_note, nm_status))

        if total_revenue != 0:
            expr = (total_expenses / total_revenue) * 100
            expr_status = 'good' if expr < 40 else ('warn' if expr < 70 else 'bad')
            expr_note   = ("Well-controlled expenses" if expr < 40 else
                           "Moderate expense load" if expr < 70 else "High expense burden")
            metrics.append(("Expense Ratio", f"{expr:.1f}%", expr_note, expr_status))

        if total_revenue != 0 and total_cogs > 0:
            cogsr = (total_cogs / total_revenue) * 100
            cogsr_status = 'good' if cogsr < 50 else ('warn' if cogsr < 70 else 'bad')
            metrics.append(("COGS Ratio", f"{cogsr:.1f}%",
                            "Portion of revenue consumed by COGS", cogsr_status))

        ni_status = 'good' if net_income > 0 else ('warn' if net_income == 0 else 'bad')
        ni_note   = ("Profitable period" if net_income > 0 else
                     "Break-even" if net_income == 0 else "Loss recorded this period")
        metrics.append(("Net Income", f"{net_income:,.0f}", ni_note, ni_status))

        rev_status = 'good' if total_revenue > 0 else 'bad'
        metrics.append(("Revenue Status",
                        "Recorded" if total_revenue > 0 else "None",
                        f"Total: {total_revenue:,.2f}", rev_status))
        self.show_analysis(metrics)


# ---------------------------------------------------------------------------
# ValidationPanel
# ---------------------------------------------------------------------------

class ValidationPanel(QWidget):
    ROW_STYLE = {
        'error':   ("🔴", "#fdecea", "#c62828"),
        'warning': ("⚠️",  "#fff8e6", "#e65100"),
        'ok':      ("✅", "#e6f9f0", "#1b5e20"),
        'info':    ("ℹ️",  "#e8eaf6", "#283593"),
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMaximumHeight(160)
        self.setMinimumHeight(100)

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 4, 4, 4)
        root.setSpacing(4)

        header_row = QHBoxLayout()
        header = QLabel("🔍  Validation & Error Detection")
        header.setStyleSheet(
            "font-size:11px; font-weight:700; color:#2d3a5e; padding:2px 4px;")
        header_row.addWidget(header)
        self._summary_lbl = QLabel("")
        self._summary_lbl.setStyleSheet(
            "font-size:10px; color:#666; padding:2px 4px;")
        header_row.addStretch()
        header_row.addWidget(self._summary_lbl)
        root.addLayout(header_row)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFrameShadow(QFrame.Sunken)
        root.addWidget(sep)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._inner = QWidget()
        self._inner_layout = QVBoxLayout(self._inner)
        self._inner_layout.setSpacing(2)
        self._inner_layout.setContentsMargins(0, 0, 0, 0)
        self._inner_layout.addStretch()
        self._scroll.setWidget(self._inner)
        root.addWidget(self._scroll)
        self._show_idle()

    def _show_idle(self):
        lbl = QLabel("  No statement generated yet. Validation will appear here.")
        lbl.setStyleSheet("color:#aaa; font-size:10px; padding:4px;")
        self._inner_layout.insertWidget(0, lbl)

    def _clear(self):
        while self._inner_layout.count() > 1:
            item = self._inner_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def _add_row(self, level: str, message: str):
        icon, bg, fg = self.ROW_STYLE.get(level, self.ROW_STYLE['info'])
        row = QLabel(f"  {icon}  {message}")
        row.setStyleSheet(
            f"background:{bg}; color:{fg}; font-size:10px; font-weight:500;"
            f"border-radius:4px; padding:3px 6px;")
        self._inner_layout.insertWidget(self._inner_layout.count() - 1, row)

    def run_validation(self, trial_balance: list, stmt_type: str,
                       total_assets=0, total_liabilities=0, total_equity=0,
                       net_income=0, total_revenue=0, total_expenses=0):
        self._clear()
        issues     = {'error': 0, 'warning': 0, 'ok': 0}
        codes_seen = {}
        asset_codes = liab_codes = eq_codes = rev_codes = exp_codes = []
        asset_codes, liab_codes, eq_codes, rev_codes, exp_codes = [], [], [], [], []

        for e in trial_balance:
            code   = e['account_code']
            desc   = e['account_description']
            amount = e['amount']
            nb     = e.get('normal_balance', 'Debit')
            suffix = code.rsplit('-', 1)[-1] if '-' in code else code
            first  = suffix[0] if suffix else ''

            if code in codes_seen:
                self._add_row('error', f"Duplicate account code: {code} ({desc})")
                issues['error'] += 1
            codes_seen[code] = True

            if first == '1':
                asset_codes.append((code, desc, amount, nb))
            elif first == '2':
                liab_codes.append((code, desc, amount, nb))
            elif first == '3':
                eq_codes.append((code, desc, amount, nb))
            elif first == '4':
                rev_codes.append((code, desc, amount, nb))
            elif first in ('5', '6', '7', '8', '9'):
                exp_codes.append((code, desc, amount, nb))
            else:
                self._add_row('warning',
                              f"Unrecognised account range: {code} ({desc})")
                issues['warning'] += 1

            if first == '1' and nb == 'Debit' and amount < -0.01:
                self._add_row('warning',
                              f"Abnormal credit balance on asset: {code} = {amount:,.2f}")
                issues['warning'] += 1
            if first == '2' and nb == 'Credit' and amount > 0.01:
                self._add_row('warning',
                              f"Abnormal debit balance on liability: {code} = {amount:,.2f}")
                issues['warning'] += 1
            if first == '4' and nb == 'Credit' and amount > 0.01:
                self._add_row('warning',
                              f"Abnormal debit balance on revenue: {code} = {amount:,.2f}")
                issues['warning'] += 1

        if stmt_type == 'position':
            if not asset_codes:
                self._add_row('error', "No asset accounts (1xx) found in trial balance")
                issues['error'] += 1
            if not liab_codes:
                self._add_row('warning',
                              "No liability accounts (2xx) — is this intentional?")
                issues['warning'] += 1
            if not eq_codes:
                self._add_row('warning', "No equity accounts (3xx) found")
                issues['warning'] += 1
            diff = total_assets - (total_liabilities + total_equity)
            if abs(diff) > 0.005:
                self._add_row('error',
                              f"Statement OUT OF BALANCE by {diff:,.2f}.")
                issues['error'] += 1
            else:
                self._add_row('ok',
                              "Statement is balanced (Assets = Liabilities + Equity)")
                issues['ok'] += 1
            if total_equity < 0:
                self._add_row('warning',
                              f"Negative equity ({total_equity:,.2f}) — possible insolvency risk")
                issues['warning'] += 1

        elif stmt_type == 'performance':
            if not rev_codes:
                self._add_row('error', "No revenue accounts (4xx) found")
                issues['error'] += 1
            if not exp_codes:
                self._add_row('warning', "No expense accounts (5xx–9xx) found")
                issues['warning'] += 1
            if total_revenue == 0:
                self._add_row('warning', "Zero total revenue recorded for the period")
                issues['warning'] += 1
            else:
                self._add_row('ok', f"Revenue recorded: {total_revenue:,.2f}")
                issues['ok'] += 1
            if net_income < 0:
                self._add_row('warning',
                              f"Net loss of {abs(net_income):,.2f} — expenses exceed revenue")
                issues['warning'] += 1
            elif net_income > 0:
                self._add_row('ok', f"Net income of {net_income:,.2f} recorded")
                issues['ok'] += 1

        parts = []
        if issues['error']:
            parts.append(f"🔴 {issues['error']} error{'s' if issues['error'] > 1 else ''}")
        if issues['warning']:
            parts.append(f"⚠️ {issues['warning']} warning{'s' if issues['warning'] > 1 else ''}")
        if issues['ok']:
            parts.append(f"✅ {issues['ok']} passed")
        self._summary_lbl.setText("  |  ".join(parts))


# ---------------------------------------------------------------------------
# SavedStatementsPanel
# ---------------------------------------------------------------------------

class SavedStatementsPanel(QWidget):
    """Left sidebar listing saved statements (DB-backed)."""

    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setMinimumWidth(190)
        self.setMaximumWidth(260)
        self._entries: list[dict] = []
        self.on_load_callback = None

        root = QVBoxLayout(self)
        root.setContentsMargins(4, 0, 4, 0)
        root.setSpacing(4)

        header = QLabel("📁  Saved Statements")
        header.setStyleSheet(
            "font-size:12px; font-weight:700; color:#2d3a5e;"
            "padding:6px 4px; border-bottom:2px solid #5c7cfa;")
        root.addWidget(header)

        self._list = QListWidget()
        self._list.setAlternatingRowColors(True)
        self._list.setStyleSheet(
            "QListWidget { border:1px solid #dde; border-radius:4px; font-size:10px; }"
            "QListWidget::item { padding:6px 4px; }"
            "QListWidget::item:selected { background:#5c7cfa; color:white; }")
        self._list.itemDoubleClicked.connect(self._load_selected)
        root.addWidget(self._list)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(4)

        self._load_btn = QPushButton("Load")
        self._load_btn.setFixedHeight(28)
        self._load_btn.setEnabled(False)
        self._load_btn.clicked.connect(self._load_selected)

        self._rename_btn = QPushButton()
        self._rename_btn.setFixedSize(28, 28)
        self._rename_btn.setEnabled(False)
        self._rename_btn.setToolTip("Rename")
        rename_icon = _qta_icon('mdi.pencil', color='#555555')
        if not rename_icon.isNull():
            self._rename_btn.setIcon(rename_icon)
            self._rename_btn.setIconSize(QSize(16, 16))
        else:
            self._rename_btn.setText("✏")
        self._rename_btn.clicked.connect(self._rename_selected)

        self._delete_btn = QPushButton()
        self._delete_btn.setFixedSize(28, 28)
        self._delete_btn.setEnabled(False)
        self._delete_btn.setToolTip("Delete")
        delete_icon = _qta_icon('mdi.delete', color='#e53935')
        if not delete_icon.isNull():
            self._delete_btn.setIcon(delete_icon)
            self._delete_btn.setIconSize(QSize(16, 16))
        else:
            self._delete_btn.setText("🗑")
        self._delete_btn.clicked.connect(self._delete_selected)

        btn_row.addWidget(self._load_btn)
        btn_row.addStretch()
        btn_row.addWidget(self._rename_btn)
        btn_row.addWidget(self._delete_btn)
        root.addLayout(btn_row)

        self._list.currentItemChanged.connect(self._on_selection_change)

        self._count_lbl = QLabel("No saved statements")
        self._count_lbl.setStyleSheet(
            "font-size:9px; color:#aaa; padding:2px 4px;")
        root.addWidget(self._count_lbl)

        self._reload_from_db()

    # ---------------------------------------------------------------- internal

    def _reload_from_db(self):
        self._list.clear()
        self._entries = _db.load_all_statements(self.db_manager)
        for entry in self._entries:
            self._add_list_item(entry)
        self._update_count()

    def _add_list_item(self, entry: dict):
        item = QListWidgetItem()
        stmt_type = entry.get('stmt_type',
                              entry.get('params', {}).get('type', ''))
        type_icon = "📊" if stmt_type == 'position' else "📈"
        ts = entry.get('timestamp', '')[:16]
        item.setText(f"{type_icon} {entry['label']}\n  {ts}")
        item.setData(Qt.UserRole, entry['id'])
        self._list.addItem(item)

    def _current_entry(self) -> dict | None:
        item = self._list.currentItem()
        if not item:
            return None
        stmt_id = item.data(Qt.UserRole)
        return next((e for e in self._entries if e['id'] == stmt_id), None)

    # ---------------------------------------------------------------- public

    def save_statement(self, label: str, text: str, params: dict):
        stmt_type = params.get('type', '')
        new_id = _db.save_statement(
            self.db_manager, label, stmt_type, text, params)
        self._reload_from_db()
        for i in range(self._list.count()):
            if self._list.item(i).data(Qt.UserRole) == new_id:
                self._list.setCurrentRow(i)
                break

    # ---------------------------------------------------------------- slots

    def _on_selection_change(self, current, previous):
        has = current is not None
        self._load_btn.setEnabled(has)
        self._rename_btn.setEnabled(has)
        self._delete_btn.setEnabled(has)

    def _load_selected(self):
        entry = self._current_entry()
        if entry and self.on_load_callback:
            self.on_load_callback(entry)

    def _rename_selected(self):
        entry = self._current_entry()
        if not entry:
            return
        new_label, ok = QInputDialog.getText(
            self, "Rename Statement", "New name:", text=entry['label'])
        if ok and new_label.strip():
            _db.rename_statement(self.db_manager, entry['id'], new_label.strip())
            self._reload_from_db()

    def _delete_selected(self):
        entry = self._current_entry()
        if not entry:
            return
        reply = QMessageBox.question(
            self, "Delete Statement", f"Delete '{entry['label']}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            _db.delete_statement(self.db_manager, entry['id'])
            self._reload_from_db()

    def _update_count(self):
        n = len(self._entries)
        self._count_lbl.setText(
            f"{n} statement{'s' if n != 1 else ''} saved" if n
            else "No saved statements")