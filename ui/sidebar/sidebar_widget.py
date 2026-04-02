"""
ui/sidebar/sidebar_widget.py
-----------------------------
The main sidebar widget with full light/dark theme support.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QScrollArea, QSizePolicy, QFrame,
)
from PySide6.QtCore import Qt, Signal, QPropertyAnimation, QEasingCurve, QSize
from PySide6.QtGui  import QFont
import qtawesome as qta

from ui.sidebar.nav_item  import NavItem
from ui.sidebar.nav_group import NavGroup


# ── Page index constants — keep in sync with main_window stack order ──
PAGE_DASHBOARD   = 0
PAGE_COA         = 1
PAGE_ALPHALIST   = 2
PAGE_SJ          = 3
PAGE_PJ          = 4
PAGE_CDJ         = 5
PAGE_CRJ         = 6
PAGE_GJ          = 7
PAGE_GL          = 8
PAGE_TB          = 9
PAGE_FS          = 10
PAGE_SETTINGS    = 11

_EXPANDED_WIDTH  = 220
_COLLAPSED_WIDTH = 54

# ── Theme colour palettes ──────────────────────────────────────────────
_THEMES = {
    'light': {
        'sidebar_bg':    '#FAFAFA',
        'sidebar_border':'#E0E0E0',
        'header_bg':     '#1565C0',
        'header_text':   'white',
        'text':          '#444444',
        'accent':        '#1565C0',
        'hover_bg':      '#E3F2FD',
        'active_bg':     '#BBDEFB',
        'footer_border': '#E0E0E0',
        'sep_color':     '#E0E0E0',
        'icon_color':    '#555555',
    },
    'dark': {
        'sidebar_bg':    '#1E1E2E',
        'sidebar_border':'#2E2E3E',
        'header_bg':     '#0D47A1',
        'header_text':   'white',
        'text':          '#CCCCCC',
        'accent':        '#42A5F5',
        'hover_bg':      '#2A2A3E',
        'active_bg':     '#1A3A5C',
        'footer_border': '#2E2E3E',
        'sep_color':     '#2E2E3E',
        'icon_color':    '#AAAAAA',
    },
}


class SidebarWidget(QWidget):

    page_requested       = Signal(int)
    theme_toggle_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._expanded       = False
        self._active_page    = PAGE_DASHBOARD
        self._all_items:  list[NavItem]  = []
        self._all_groups: list[NavGroup] = []
        self._current_theme  = 'light'

        self.setFixedWidth(_COLLAPSED_WIDTH)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Expanding)
        self.setAttribute(Qt.WA_StyledBackground, True)

        self._anim  = QPropertyAnimation(self, b'minimumWidth')
        self._anim2 = QPropertyAnimation(self, b'maximumWidth')
        for a in (self._anim, self._anim2):
            a.setDuration(220)
            a.setEasingCurve(QEasingCurve.InOutQuad)

        self._separators: list[QFrame] = []   # must exist before _build_ui
        self._build_ui()
        self._apply_theme_colours('light')

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Header ────────────────────────────────────────────────────
        self._header_widget = QWidget()
        self._header_widget.setFixedHeight(56)
        hlay = QHBoxLayout(self._header_widget)
        hlay.setContentsMargins(12, 0, 12, 0)
        hlay.setSpacing(10)

        self._toggle_btn = QPushButton()
        self._toggle_btn.setFixedSize(30, 30)
        self._toggle_btn.setCursor(Qt.PointingHandCursor)
        self._toggle_btn.setStyleSheet(
            'QPushButton { background: transparent; border: none; }'
            'QPushButton:hover { background: rgba(255,255,255,0.15); border-radius: 4px; }')
        self._refresh_toggle_icon()
        self._toggle_btn.clicked.connect(self._toggle_expand)
        hlay.addWidget(self._toggle_btn)

        self._app_label = QLabel('ESQ Accounting')
        font = QFont(); font.setPointSize(11); font.setBold(True)
        self._app_label.setFont(font)
        self._app_label.setVisible(False)
        hlay.addWidget(self._app_label, stretch=1)

        root.addWidget(self._header_widget)

        # ── Scrollable nav area ───────────────────────────────────────
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._scroll.setFrameShape(QFrame.NoFrame)

        self._nav_container = QWidget()
        self._nav_layout = QVBoxLayout(self._nav_container)
        self._nav_layout.setContentsMargins(4, 8, 4, 8)
        self._nav_layout.setSpacing(2)

        self._build_nav()

        self._nav_layout.addStretch()
        self._scroll.setWidget(self._nav_container)
        root.addWidget(self._scroll, stretch=1)

        # ── Footer ────────────────────────────────────────────────────
        self._footer_widget = QWidget()
        self._footer_widget.setFixedHeight(48)
        flay = QHBoxLayout(self._footer_widget)
        flay.setContentsMargins(12, 0, 12, 0)
        flay.setSpacing(8)

        self._theme_btn = QPushButton()
        self._theme_btn.setFixedSize(28, 28)
        self._theme_btn.setCursor(Qt.PointingHandCursor)
        self._theme_btn.setToolTip('Toggle Light / Dark theme')
        self._theme_btn.setStyleSheet(
            'QPushButton { background: transparent; border: none; border-radius: 4px; }'
            'QPushButton:hover { background: #E3F2FD; }')
        self._theme_btn.clicked.connect(self.theme_toggle_clicked)
        flay.addWidget(self._theme_btn)

        self._theme_label = QLabel('Toggle Theme')
        self._theme_label.setVisible(False)
        flay.addWidget(self._theme_label, stretch=1)

        root.addWidget(self._footer_widget)

    def _build_nav(self):
        self._add_item(NavItem('fa5s.chart-line', '    Dashboard', PAGE_DASHBOARD))

        self._separators.append(self._make_separator())
        self._nav_layout.addWidget(self._separators[-1])

        journals = NavGroup('fa5s.book', '    Journals')
        journals.add_child(NavItem('fa5s.shopping-cart',    '    Sales Journal',     PAGE_SJ,  indent=True))
        journals.add_child(NavItem('fa5s.shopping-bag',     '    Purchase Journal',  PAGE_PJ,  indent=True))
        journals.add_child(NavItem('fa5s.money-bill-wave',  '    Cash Disbursement', PAGE_CDJ, indent=True))
        journals.add_child(NavItem('fa5s.hand-holding-usd', '    Cash Receipts',     PAGE_CRJ, indent=True))
        journals.add_child(NavItem('fa5s.pen-alt',          '    General Journal',   PAGE_GJ,  indent=True))
        self._add_group(journals)

        masterlist = NavGroup('fa5s.list', '    Masterlist')
        masterlist.add_child(NavItem('fa5s.sitemap', '    Chart of Accounts', PAGE_COA,       indent=True))
        masterlist.add_child(NavItem('fa5s.users',   '    Alphalist',         PAGE_ALPHALIST, indent=True))
        self._add_group(masterlist)

        self._separators.append(self._make_separator())
        self._nav_layout.addWidget(self._separators[-1])

        reports = NavGroup('fa5s.chart-bar', '    Reports')
        reports.add_child(NavItem('fa5s.book-open',     '    General Ledger',       PAGE_GL, indent=True))
        reports.add_child(NavItem('fa5s.balance-scale', '    Trial Balance',        PAGE_TB, indent=True))
        reports.add_child(NavItem('fa5s.file-invoice',  '    Financial Statements', PAGE_FS, indent=True))
        self._add_group(reports)

        self._separators.append(self._make_separator())
        self._nav_layout.addWidget(self._separators[-1])

        self._add_item(NavItem('fa5s.cog', '    Settings', PAGE_SETTINGS))

    def _add_item(self, item: NavItem):
        item.clicked.connect(self._on_page_requested)
        self._all_items.append(item)
        self._nav_layout.addWidget(item)

    def _add_group(self, group: NavGroup):
        group.page_requested.connect(self._on_page_requested)
        self._all_groups.append(group)
        self._nav_layout.addWidget(group)

    @staticmethod
    def _make_separator() -> QFrame:
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFixedHeight(1)
        sep.setStyleSheet('background-color: #E0E0E0; border: none; margin: 4px 8px;')
        return sep

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def set_active_page(self, page_index: int):
        self._active_page = page_index
        for item in self._all_items:
            item.set_active(item.get_page_index() == page_index)
        for group in self._all_groups:
            if group.has_page(page_index):
                group.set_child_active(page_index)
                group.open()
            else:
                group.set_child_active(-1)

    def set_theme(self, mode: str):
        """Call this whenever the app theme changes. mode = 'light' | 'dark'."""
        self._current_theme = mode
        self._apply_theme_colours(mode)

    # ------------------------------------------------------------------
    # Theme application
    # ------------------------------------------------------------------

    def _apply_theme_colours(self, mode: str):
        c = _THEMES.get(mode, _THEMES['light'])

        # Sidebar background + border
        self.setStyleSheet(
            f'SidebarWidget {{ background-color: {c["sidebar_bg"]}; '
            f'border-right: 1px solid {c["sidebar_border"]}; }}')

        # Header
        self._header_widget.setStyleSheet(
            f'background-color: {c["header_bg"]};')
        self._app_label.setStyleSheet(f'color: {c["header_text"]};')

        # Nav container background
        self._nav_container.setStyleSheet(
            f'background-color: {c["sidebar_bg"]};')
        self._scroll.setStyleSheet(
            f'QScrollArea {{ border: none; background: {c["sidebar_bg"]}; }}')

        # Footer
        self._footer_widget.setStyleSheet(
            f'border-top: 1px solid {c["footer_border"]}; '
            f'background-color: {c["sidebar_bg"]};')
        self._theme_label.setStyleSheet(
            f'color: {c["icon_color"]}; font-size: 9pt;')

        # Theme button icon
        icon = qta.icon('fa5s.adjust', color=c['icon_color'])
        self._theme_btn.setIcon(icon)
        self._theme_btn.setIconSize(QSize(18, 18))
        self._theme_btn.setStyleSheet(
            f'QPushButton {{ background: transparent; border: none; border-radius: 4px; }}'
            f'QPushButton:hover {{ background: {c["hover_bg"]}; }}')

        # Separators
        for sep in self._separators:
            sep.setStyleSheet(
                f'background-color: {c["sep_color"]}; border: none; margin: 4px 8px;')

        # Propagate to all nav items and groups
        for item in self._all_items:
            item.set_theme(c)
        for group in self._all_groups:
            group.set_theme(c)

        self._refresh_toggle_icon()

    # ------------------------------------------------------------------
    # Collapse / Expand
    # ------------------------------------------------------------------

    def _toggle_expand(self):
        self._set_expanded(not self._expanded)

    def _set_expanded(self, expanded: bool):
        self._expanded = expanded
        target = _EXPANDED_WIDTH if expanded else _COLLAPSED_WIDTH

        for anim, prop in ((self._anim, b'minimumWidth'),
                            (self._anim2, b'maximumWidth')):
            anim.stop()
            anim.setPropertyName(prop)
            anim.setTargetObject(self)
            anim.setStartValue(self.width())
            anim.setEndValue(target)
            anim.start()

        self._app_label.setVisible(expanded)
        self._theme_label.setVisible(expanded)
        self._refresh_toggle_icon()

        for item in self._all_items:
            item.set_expanded(expanded)
        for group in self._all_groups:
            group.set_expanded(expanded)

    def _refresh_toggle_icon(self):
        c = _THEMES.get(self._current_theme, _THEMES['light'])
        icon_name = 'fa5s.times' if self._expanded else 'fa5s.bars'
        icon = qta.icon(icon_name, color=c['header_text'])
        self._toggle_btn.setIcon(icon)
        self._toggle_btn.setIconSize(QSize(16, 16))

    # ------------------------------------------------------------------
    # Slot
    # ------------------------------------------------------------------

    def _on_page_requested(self, page_index: int):
        self.set_active_page(page_index)
        self.page_requested.emit(page_index)