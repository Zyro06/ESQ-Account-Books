"""
ui/sidebar/nav_group.py
-----------------------
A collapsible navigation group header with child NavItems.
"""

from __future__ import annotations

from PySide6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel
from PySide6.QtCore    import Qt, Signal, QPropertyAnimation, QEasingCurve, QSize
from PySide6.QtGui     import QFont
import qtawesome as qta

from ui.sidebar.nav_item import NavItem


class NavGroup(QWidget):
    page_requested = Signal(int)

    _COLOUR_HEADER   = '#333333'
    _COLOUR_HOVER_BG = '#E3F2FD'
    _ARROW_OPEN      = 'fa5s.chevron-down'
    _ARROW_CLOSED    = 'fa5s.chevron-right'

    def __init__(self, icon_name: str, label: str, parent=None):
        super().__init__(parent)
        self._icon_name   = icon_name
        self._label_text  = label
        self._open        = False
        self._was_open    = False
        self._children:   list[NavItem] = []
        self._expanded    = True

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ── Header row ────────────────────────────────────────────────
        self._header = QWidget()
        self._header.setFixedHeight(44)
        self._header.setCursor(Qt.PointingHandCursor)
        self._header.setAttribute(Qt.WA_StyledBackground, True)
        self._header.setStyleSheet(
            'QWidget { border-left: 3px solid transparent; border-radius: 4px; }')

        hlay = QHBoxLayout(self._header)
        hlay.setContentsMargins(12, 0, 12, 0)
        hlay.setSpacing(10)

        self._icon_lbl = QLabel()
        self._icon_lbl.setFixedSize(22, 22)
        self._icon_lbl.setAlignment(Qt.AlignCenter)
        self._refresh_icon(self._COLOUR_HEADER)
        hlay.addWidget(self._icon_lbl)

        self._label_lbl = QLabel(label)
        self._label_lbl.setMaximumWidth(0)   # ← hidden initially (sidebar starts collapsed)
        self._label_lbl.setVisible(False)
        font = QFont(); font.setPointSize(10); font.setBold(True)
        self._label_lbl.setFont(font)
        self._label_lbl.setStyleSheet(f'color: {self._COLOUR_HEADER};')
        hlay.addWidget(self._label_lbl, stretch=1)

        self._arrow_lbl = QLabel()
        self._arrow_lbl.setFixedSize(16, 16)
        self._arrow_lbl.setAlignment(Qt.AlignCenter)
        self._arrow_lbl.setMaximumWidth(0)   # ← hidden initially
        self._arrow_lbl.setVisible(False)
        self._refresh_arrow()
        hlay.addWidget(self._arrow_lbl)

        outer.addWidget(self._header)

        # ── Children container ────────────────────────────────────────
        self._children_container = QWidget()
        self._children_layout    = QVBoxLayout(self._children_container)
        self._children_layout.setContentsMargins(0, 0, 0, 0)
        self._children_layout.setSpacing(0)
        self._children_container.setMaximumHeight(0)
        self._children_container.setVisible(False)
        outer.addWidget(self._children_container)

        # ── Animation ─────────────────────────────────────────────────
        self._anim = QPropertyAnimation(
            self._children_container, b'maximumHeight')
        self._anim.setDuration(180)
        self._anim.setEasingCurve(QEasingCurve.InOutQuad)
        self._anim_close_conn = None

        self._header.mousePressEvent = self._on_header_click

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def add_child(self, item: NavItem):
        self._children.append(item)
        self._children_layout.addWidget(item)
        item.clicked.connect(self.page_requested)

    def set_expanded(self, expanded: bool):
        self._expanded = expanded
        self._label_lbl.setVisible(expanded)
        self._label_lbl.setMaximumWidth(16777215 if expanded else 0)
        self._arrow_lbl.setVisible(expanded)
        self._arrow_lbl.setMaximumWidth(16777215 if expanded else 0)
        for child in self._children:
            child.set_expanded(expanded)
        if not expanded and self._open:
            self._set_open(False, animate=False)

    def set_child_active(self, page_index: int):
        for child in self._children:
            child.set_active(child.get_page_index() == page_index)

    def has_page(self, page_index: int) -> bool:
        return any(c.get_page_index() == page_index for c in self._children)

    def open(self):
        if not self._open:
            self._set_open(True)

    def set_theme(self, colours: dict):
        """Update colours when the app theme changes."""
        self._COLOUR_HEADER   = colours['text']
        self._COLOUR_HOVER_BG = colours['hover_bg']

        self._refresh_icon(self._COLOUR_HEADER)
        self._refresh_arrow()
        self._label_lbl.setStyleSheet(f'color: {self._COLOUR_HEADER};')
        self._header.setStyleSheet(
            'QWidget { border-left: 3px solid transparent; border-radius: 4px; }')

        for child in self._children:
            child.set_theme(colours)

    # ------------------------------------------------------------------
    # Private
    # ------------------------------------------------------------------

    def _on_header_click(self, event):
        if event.button() == Qt.LeftButton:
            self._set_open(not self._open)

    def _set_open(self, open_: bool, animate: bool = True):
        self._open = open_
        self._refresh_arrow()

        target_height = self._full_height() if open_ else 0

        if not animate:
            self._children_container.setMaximumHeight(target_height)
            self._children_container.setVisible(open_)
            return

        if open_:
            self._children_container.setVisible(True)

        self._anim.stop()

        if self._anim_close_conn is not None:
            self._anim.finished.disconnect(self._anim_close_conn)
            self._anim_close_conn = None

        self._anim.setStartValue(self._children_container.maximumHeight())
        self._anim.setEndValue(target_height)

        if not open_:
            self._anim_close_conn = self._anim.finished.connect(
                lambda: self._children_container.setVisible(False))

        self._anim.start()

    def _full_height(self) -> int:
        return len(self._children) * 44

    def _refresh_icon(self, colour: str):
        icon = qta.icon(self._icon_name, color=colour)
        self._icon_lbl.setPixmap(icon.pixmap(QSize(20, 20)))

    def _refresh_arrow(self):
        name = self._ARROW_OPEN if self._open else self._ARROW_CLOSED
        icon = qta.icon(name, color=self._COLOUR_HEADER)
        self._arrow_lbl.setPixmap(icon.pixmap(QSize(12, 12)))

    def enterEvent(self, event):
        self._header.setStyleSheet(
            f'QWidget {{ background-color: {self._COLOUR_HOVER_BG}; '
            f'border-left: 3px solid transparent; border-radius: 4px; }}')
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._header.setStyleSheet(
            'QWidget { border-left: 3px solid transparent; border-radius: 4px; }')
        super().leaveEvent(event)