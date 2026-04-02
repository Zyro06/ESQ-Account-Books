"""
ui/sidebar/nav_item.py
----------------------
A single clickable navigation item in the sidebar.
"""

from __future__ import annotations

from PySide6.QtWidgets import QWidget, QHBoxLayout, QLabel
from PySide6.QtCore    import Qt, Signal, QSize
from PySide6.QtGui     import QFont
import qtawesome as qta


class NavItem(QWidget):

    clicked = Signal(int)

    _COLOUR_NORMAL   = '#444444'
    _COLOUR_ACTIVE   = '#1565C0'
    _COLOUR_HOVER_BG = '#E3F2FD'
    _COLOUR_ACTIVE_BG= '#BBDEFB'

    def __init__(self, icon_name: str, label: str, page_index: int,
                 indent: bool = False, parent=None):
        super().__init__(parent)
        self._icon_name  = icon_name
        self._label_text = label
        self._page_index = page_index
        self._indent     = indent
        self._active     = False
        self._sidebar_expanded = True

        self.setFixedHeight(44)
        self.setCursor(Qt.PointingHandCursor)
        self.setAttribute(Qt.WA_StyledBackground, True)

        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(24 if indent else 12, 0, 12, 0)
        self._layout.setSpacing(10)

        self._icon_lbl = QLabel()
        self._icon_lbl.setFixedSize(22, 22)
        self._icon_lbl.setAlignment(Qt.AlignCenter)
        self._set_icon_colour(self._COLOUR_NORMAL)
        self._layout.addWidget(self._icon_lbl)

        self._label_lbl = QLabel(label)
        self._label_lbl.setMaximumWidth(0)   # ← hidden initially (sidebar starts collapsed)
        self._label_lbl.setVisible(False)
        font = QFont()
        font.setPointSize(9 if indent else 10)
        self._label_lbl.setFont(font)
        self._label_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self._layout.addWidget(self._label_lbl, stretch=1)

        self._apply_style(False)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def set_active(self, active: bool):
        self._active = active
        colour = self._COLOUR_ACTIVE if active else self._COLOUR_NORMAL
        self._set_icon_colour(colour)
        self._label_lbl.setStyleSheet(
            f'color: {colour}; font-weight: {"bold" if active else "normal"};')
        self._apply_style(active)

    def set_expanded(self, expanded: bool):
        """Show/hide the text label based on sidebar expanded state."""
        self._sidebar_expanded = expanded
        self._label_lbl.setVisible(expanded)
        self._label_lbl.setMaximumWidth(16777215 if expanded else 0)
        # When collapsed, remove indent so icon stays centred
        self._layout.setContentsMargins(
            (24 if self._indent else 12) if expanded else 12,
            0, 12, 0
        )

    def get_page_index(self) -> int:
        return self._page_index

    def set_theme(self, colours: dict):
        """Update colours when the app theme changes."""
        self._COLOUR_NORMAL    = colours['text']
        self._COLOUR_ACTIVE    = colours['accent']
        self._COLOUR_HOVER_BG  = colours['hover_bg']
        self._COLOUR_ACTIVE_BG = colours['active_bg']
        self.set_active(self._active)

    # ------------------------------------------------------------------
    # Events
    # ------------------------------------------------------------------

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self._page_index)
        super().mousePressEvent(event)

    def enterEvent(self, event):
        if not self._active:
            self._apply_style(False, hovered=True)
        super().enterEvent(event)

    def leaveEvent(self, event):
        if not self._active:
            self._apply_style(False, hovered=False)
        super().leaveEvent(event)

    # ------------------------------------------------------------------
    # Private
    # ------------------------------------------------------------------

    def _set_icon_colour(self, colour: str):
        icon = qta.icon(self._icon_name, color=colour)
        self._icon_lbl.setPixmap(icon.pixmap(QSize(20, 20)))

    def _apply_style(self, active: bool, hovered: bool = False):
        if active:
            bg     = self._COLOUR_ACTIVE_BG
            border = f'border-left: 3px solid {self._COLOUR_ACTIVE};'
        elif hovered:
            bg     = self._COLOUR_HOVER_BG
            border = 'border-left: 3px solid transparent;'
        else:
            bg     = 'transparent'
            border = 'border-left: 3px solid transparent;'
        self.setStyleSheet(
            f'NavItem {{ background-color: {bg}; {border} border-radius: 4px; }}')