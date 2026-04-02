"""
utils/date_utils.py
-------------------
Shared date parsing utilities used across all journal widgets
and the fullbook importer.

No Qt imports here on purpose — _norm_date, _norm_float, _norm_str
are pure Python and used by the importer with no Qt dependency.

The Qt-dependent DateItem class is kept here too since it is
copy-pasted identically across CDJ, CRJ, GJ, SJ, and PJ widgets.
"""

from __future__ import annotations

from datetime import datetime, date as _date
from typing import Any


# ---------------------------------------------------------------------------
# Pure-Python helpers (no Qt) — used by fullbook_importer and widgets
# ---------------------------------------------------------------------------

# Date formats tried in order when parsing a free-form date string.
_READ_FORMATS = (
    '%m/%d/%Y',   # MM/dd/yyyy  ← app's native format
    '%m/%d/%y',   # MM/dd/yy
    '%Y-%m-%d %H:%M:%S',
    '%Y-%m-%d',
    '%m-%d-%Y',
    '%d/%m/%Y',
)

# Qt display / parse formats tried in order (used by _parse_date_qdate).
_QT_READ_FORMATS = (
    'MM/dd/yyyy',
    'M/d/yyyy',
    'M/dd/yyyy',
    'MM/d/yyyy',
    'yyyy-MM-dd',
)


def norm_date(value: Any) -> str:
    """
    Convert any date-like value to MM/dd/yyyy string.

    Accepts:
        - datetime or date objects
        - 'YYYY-MM-DD HH:MM:SS', 'YYYY-MM-DD', 'MM/dd/yyyy', 'MM/dd/yy', etc.
        - openpyxl serial numbers are handled automatically because openpyxl
          already converts them to datetime objects before we see them.

    Returns '' on failure.
    """
    if value is None:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%m/%d/%Y')
    if isinstance(value, _date):
        return value.strftime('%m/%d/%Y')
    s = str(value).strip()
    if not s or s.lower() == 'nan':
        return ''
    for fmt in _READ_FORMATS:
        try:
            return datetime.strptime(s, fmt).strftime('%m/%d/%Y')
        except ValueError:
            continue
    return ''


def norm_float(value: Any) -> float:
    """Safely convert any value to float, stripping commas. Returns 0.0 on failure."""
    if value is None:
        return 0.0
    try:
        return float(str(value).replace(',', '').strip() or 0)
    except (ValueError, TypeError):
        return 0.0


def norm_str(value: Any) -> str:
    """Safely convert any value to stripped string. Returns '' for None/nan."""
    if value is None:
        return ''
    s = str(value).strip()
    return '' if s.lower() == 'nan' else s


# ---------------------------------------------------------------------------
# Qt-dependent helpers
# ---------------------------------------------------------------------------

def parse_date_qdate(date_str: str):
    """
    Parse a date string into a QDate using multiple format attempts.
    Returns an invalid QDate() if no format matches.

    Import example:
        from utils.date_utils import parse_date_qdate
        qdate = parse_date_qdate('03/25/2026')
    """
    # Lazy import so the module can be used without Qt installed
    # (e.g. in unit tests that only test norm_date / norm_float).
    from PySide6.QtCore import QDate  # swap to PySide6 when migrating
    for fmt in _QT_READ_FORMATS:
        d = QDate.fromString(date_str, fmt)
        if d.isValid():
            return d
    return QDate()


class DateItem:
    """
    A QTableWidgetItem subclass that sorts by date value rather than
    alphabetically.  Used in every journal widget's date column.

    Usage:
        from utils.date_utils import DateItem

        item = DateItem('03/25/2026')
        item.setData(Qt.UserRole, some_data)
        table.setItem(row, 0, item)
    """

    # Defined as a factory/class body here; the actual class is created
    # on first access so Qt is only imported when needed.
    _cls = None

    @classmethod
    def _build(cls):
        if cls._cls is not None:
            return cls._cls

        from PySide6.QtWidgets import QTableWidgetItem  # swap to PySide6 when migrating

        class _DateItem(QTableWidgetItem):
            def __init__(self, display_text: str):
                super().__init__(display_text)
                qdate = parse_date_qdate(display_text)
                self._sort_key = (
                    qdate.toString('yyyy-MM-dd') if qdate.isValid() else display_text
                )

            def __lt__(self, other):
                if isinstance(other, _DateItem):
                    return self._sort_key < other._sort_key
                return super().__lt__(other)

        cls._cls = _DateItem
        return _DateItem

    def __class_getitem__(cls, item):
        return cls._build()

    def __new__(cls, *args, **kwargs):
        real_cls = cls._build()
        return real_cls(*args, **kwargs)


# ---------------------------------------------------------------------------
# Convenience re-export so callers can do:
#   from utils.date_utils import DateItem, norm_date, norm_float, norm_str
# ---------------------------------------------------------------------------

__all__ = [
    'norm_date',
    'norm_float',
    'norm_str',
    'parse_date_qdate',
    'DateItem',
]