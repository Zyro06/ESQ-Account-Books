"""
search_utils.py
---------------
Centralised, optimised search / filter logic for all journal widgets.

Usage
-----
1.  Create a SearchFilter after the table and controls exist:

        self._search = SearchFilter(
            table         = self.table,
            search_input  = self.search_input,
            results_label = self.results_label,
            date_from     = self.date_from,   # optional QDateEdit
            date_to       = self.date_to,     # optional QDateEdit
            month_combo   = self.month_combo, # optional QComboBox
            date_col      = 0,                # column index that holds the date string
            debounce_ms   = 300,              # typing delay before filter fires
        )

2.  Call  self._search.refresh()  any time the underlying data changes
    (e.g. after load_data rebuilds the table rows).

3.  To update the results label externally:
        self._search.update_label()

The filter uses setRowHidden — it never rebuilds QTableWidgetItems, so it
is O(n) in the number of rows rather than O(n * columns).
"""

from PyQt5.QtCore import QDate, QTimer
from PyQt5.QtWidgets import QTableWidget, QLineEdit, QLabel, QDateEdit, QComboBox

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def add_month_combo(layout, label_text: str = "Month:") -> QComboBox:
    """
    Convenience helper — creates and adds a labelled month QComboBox to a
    QHBoxLayout.  Returns the combo so the caller can store a reference.

        from PyQt5.QtWidgets import QLabel
        self.month_combo = add_month_combo(search_layout)
    """
    from PyQt5.QtWidgets import QLabel
    lbl = QLabel(label_text)
    combo = QComboBox()
    combo.addItem("All Months")
    combo.addItems(MONTHS)
    combo.setFixedWidth(150)
    layout.addWidget(lbl)
    layout.addWidget(combo)
    return combo


class SearchFilter:
    """
    Attaches debounced, row-hiding search+filter behaviour to a QTableWidget.

    Parameters
    ----------
    table           QTableWidget to filter
    search_input    QLineEdit  — text search box
    results_label   QLabel     — shows "Showing X of Y"
    date_from       QDateEdit  — optional lower date bound
    date_to         QDateEdit  — optional upper date bound
    month_combo     QComboBox  — optional month picker ("All Months" + 12 months)
    date_col        int        — column index holding the date string (default 0)
    debounce_ms     int        — ms to wait after last keystroke (default 300)
    date_fmt        str        — QDate format for parsing the date column
    """

    def __init__(
        self,
        table:          QTableWidget,
        search_input:   QLineEdit,
        results_label:  QLabel,
        date_from:      QDateEdit  = None,
        date_to:        QDateEdit  = None,
        month_combo:    QComboBox  = None,
        date_col:       int        = 0,
        debounce_ms:    int        = 300,
        date_fmt:       str        = "MM/dd/yyyy",
    ):
        self._table         = table
        self._search_input  = search_input
        self._results_label = results_label
        self._date_from     = date_from
        self._date_to       = date_to
        self._month_combo   = month_combo
        self._date_col      = date_col
        self._date_fmt      = date_fmt

        # Debounce timer — only fires after user pauses typing
        self._timer = QTimer()
        self._timer.setSingleShot(True)
        self._timer.setInterval(debounce_ms)
        self._timer.timeout.connect(self._run)

        # Connect signals
        # Text search — debounced
        search_input.textChanged.connect(self._timer.start)

        # Date pickers — instant (no debounce needed, single interaction)
        if date_from is not None:
            date_from.dateChanged.connect(self._run)
        if date_to is not None:
            date_to.dateChanged.connect(self._run)

        # Month combo — instant
        if month_combo is not None:
            month_combo.currentIndexChanged.connect(self._run)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def refresh(self):
        """Call after load_data rebuilds the table to reapply current filters."""
        self._run()

    def update_label(self):
        """Force-update the results label without re-filtering."""
        shown = sum(
            1 for r in range(self._table.rowCount())
            if not self._table.isRowHidden(r)
        )
        total = self._table.rowCount()
        self._results_label.setText(f"Showing: {shown} of {total}")

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _run(self):
        """Apply all active filters using setRowHidden (no row rebuilding)."""
        search_text = self._search_input.text().strip().lower()

        # Date range
        has_date_range = self._date_from is not None and self._date_to is not None
        if has_date_range:
            qdate_from = self._date_from.date()
            qdate_to   = self._date_to.date()

        # Month filter  (0 = "All Months", 1-12 = specific month)
        month_filter = 0
        if self._month_combo is not None:
            month_filter = self._month_combo.currentIndex()  # 0=All, 1=Jan … 12=Dec

        shown = 0
        total = self._table.rowCount()

        for row in range(total):
            hide = False

            # ── Date column parsing ────────────────────────────────────
            date_item = self._table.item(row, self._date_col)
            if date_item is None:
                self._table.setRowHidden(row, True)
                continue

            date_str  = date_item.text()
            entry_date = QDate.fromString(date_str, self._date_fmt)
            if not entry_date.isValid():
                # Try fallback format
                entry_date = QDate.fromString(date_str, "yyyy-MM-dd")

            # ── Date range filter ──────────────────────────────────────
            if has_date_range and entry_date.isValid():
                if entry_date < qdate_from or entry_date > qdate_to:
                    hide = True

            # ── Month filter ───────────────────────────────────────────
            if not hide and month_filter > 0 and entry_date.isValid():
                if entry_date.month() != month_filter:
                    hide = True

            # ── Text search ────────────────────────────────────────────
            if not hide and search_text:
                haystack = " ".join(
                    self._table.item(row, col).text()
                    for col in range(self._table.columnCount())
                    if self._table.item(row, col)
                ).lower()
                if search_text not in haystack:
                    hide = True

            self._table.setRowHidden(row, hide)
            if not hide:
                shown += 1

        self._results_label.setText(f"Showing: {shown} of {total}")