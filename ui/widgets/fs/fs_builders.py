"""
ui/widgets/fs/fs_builders.py
-----------------------------
Pure-logic statement builders — no Qt widgets, no DB access.

Public API:
    fmt_amt(v) -> str
    build_position(trial_balance, stmt_name, company, as_of, business_type)
        -> (analysis_data, screen_text, structured)
    build_performance(trial_balance, stmt_name, company, from_d, to_d)
        -> (analysis_data, screen_text, structured)
"""

from __future__ import annotations

import calendar as _calendar
from PySide6.QtCore import QDate

# ---------------------------------------------------------------------------
# Date / period helpers
# ---------------------------------------------------------------------------

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def fmt_date(qdate: QDate) -> str:
    return f"{MONTHS[qdate.month() - 1]} {qdate.day()}, {qdate.year()}"


def month_last_day(year: int, month: int) -> int:
    return _calendar.monthrange(year, month)[1]


def _centre(text: str, width: int) -> str:
    return text.center(width)


def period_label(from_d: QDate, to_d: QDate) -> str:
    num_months = (to_d.year() - from_d.year()) * 12 + \
                 (to_d.month() - from_d.month()) + 1
    SMALL_WORDS = {1: "One", 2: "Two", 3: "Three", 4: "Four", 5: "Five",
                   6: "Six", 7: "Seven", 8: "Eight", 9: "Nine",
                   10: "Ten", 11: "Eleven"}
    YEAR_WORDS  = {1: "the Year", 2: "Two Years", 3: "Three Years",
                   4: "Four Years", 5: "Five Years"}
    end_date = fmt_date(to_d)
    if num_months == 1:
        span = "the Month"
    elif num_months % 12 == 0:
        span = YEAR_WORDS.get(num_months // 12, f"{num_months // 12} Years")
    elif num_months > 12:
        span = f"{num_months} Months"
    else:
        span = f"{SMALL_WORDS.get(num_months, str(num_months))} Months"
    return f"For {span} Ending {end_date}"


# ---------------------------------------------------------------------------
# Amount formatter
# ---------------------------------------------------------------------------

def fmt_amt(v: float) -> str:
    return f"{v:,.2f}"


# ---------------------------------------------------------------------------
# Internal text-layout helpers
# ---------------------------------------------------------------------------

def _make_line_fn(W: int, AM: int):
    def _line(left_text: str, amt_str: str = "") -> str:
        if not amt_str:
            return left_text
        label_w = W - AM - 1
        label = (left_text + " " * label_w)[:label_w]
        return f"{label} {amt_str:>{AM}}"
    return _line


def _row_fn(line_fn):
    def row(code, desc, amt_str, indent=0):
        return line_fn(" " * indent + f"{code}  {desc}", amt_str)
    return row


def _total_fn(line_fn):
    def total_row(label, amt_str, indent=0):
        return line_fn(" " * indent + label, amt_str)
    return total_row


def _r(fa, code, desc, amt, is_total=False, is_sub=False):
    return (code, desc, fa(amt), is_total, is_sub)


# ---------------------------------------------------------------------------
# Financial Position builder
# ---------------------------------------------------------------------------

def build_position(
    trial_balance: list,
    stmt_name: str,
    company: str,
    as_of: QDate,
    business_type: str = 'Sole Proprietorship',
) -> tuple[dict, str, dict]:
    """
    Returns (analysis_data, screen_text, structured).

    analysis_data keys: total_assets, total_liabilities, total_equity,
                        net_income, current_assets, current_liabilities
    structured keys:    company, stmt_name, subtitle, sections, warning
    """
    BIZ_EQUITY_LABELS = {
        'Sole Proprietorship': "Owner's Equity",
        'Partnership':         "Partners' Equity",
        'Corporation':         "Stockholders' Equity",
    }
    BIZ_NI_LABELS = {
        'Sole Proprietorship': "Net Income (Loss)",
        'Partnership':         "Net Income (Loss) - Current Period",
        'Corporation':         "Net Income (Loss) for the Period",
    }
    equity_label = BIZ_EQUITY_LABELS.get(business_type, "Owner's Equity")
    ni_label     = BIZ_NI_LABELS.get(business_type, "Net Income (Loss)")

    assets      = []
    liabilities = []
    equity_adds = []
    equity_deds = []
    rev_total   = 0.0
    exp_total   = 0.0

    for e in trial_balance:
        code   = e['account_code']
        desc   = e['account_description']
        amount = e['amount']
        nb     = e.get('normal_balance', 'Debit')
        suffix = code.rsplit('-', 1)[-1] if '-' in code else code
        first  = suffix[0] if suffix else ''

        if first == '1':
            assets.append((code, desc, amount))
        elif first == '2':
            liabilities.append((code, desc, -amount))
        elif first == '3':
            if nb == 'Debit':
                equity_deds.append((code, desc, amount))
            else:
                equity_adds.append((code, desc, -amount))
        elif first == '4':
            rev_total += abs(amount)
        elif first in ('5', '6', '7', '8', '9'):
            exp_total += amount

    net_income        = rev_total - exp_total
    total_assets      = sum(a[2] for a in assets)
    total_liabilities = sum(a[2] for a in liabilities)
    total_equity_adds = sum(a[2] for a in equity_adds)
    total_equity_deds = sum(a[2] for a in equity_deds)
    total_equity      = total_equity_adds - total_equity_deds + net_income
    total_l_and_e     = total_liabilities + total_equity
    diff              = total_assets - total_l_and_e

    fa       = fmt_amt
    W, AM    = 85, 15
    _line    = _make_line_fn(W, AM)
    row      = _row_fn(_line)
    tot_row  = _total_fn(_line)

    lines = [
        _centre(company, W),
        _centre(stmt_name.upper(), W),
        _centre(f"As of {fmt_date(as_of)}", W),
        "=" * W, "",
        "ASSETS", "-" * W,
    ]
    for code, desc, amt in assets:
        lines.append(row(code, desc, fa(amt)))
    lines += ["-" * W, tot_row("TOTAL ASSETS", fa(total_assets)), "",
              f"LIABILITIES AND {equity_label.upper()}", "-" * W, "  Liabilities:"]
    for code, desc, amt in liabilities:
        lines.append(row(code, desc, fa(amt), indent=4))
    lines += ["-" * W, tot_row("TOTAL LIABILITIES", fa(total_liabilities)), "",
              f"  {equity_label}:"]
    for code, desc, amt in equity_adds:
        lines.append(row(code, desc, fa(amt), indent=4))
    for code, desc, amt in equity_deds:
        lines.append(row(code, desc, fa(-amt), indent=4))
    lines += [
        _line("    " + ni_label, fa(net_income)),
        "-" * W,
        tot_row(f"TOTAL {equity_label.upper()}", fa(total_equity)),
        "",
        tot_row(f"TOTAL LIABILITIES AND {equity_label.upper()}", fa(total_l_and_e)),
        "",
    ]
    if abs(diff) > 0.005:
        lines += ["!" * W, f"  WARNING: OUT OF BALANCE by {diff:,.2f}", "!" * W, ""]
    lines.append("=" * W)
    screen_text = "\n".join(lines)

    # --- Structured (for PDF) ---
    def _rv(code, desc, amt, is_total=False, is_sub=False):
        return _r(fa, code, desc, amt, is_total, is_sub)

    sections = []

    asset_rows = [_rv(c, d, a) for c, d, a in assets]
    asset_rows.append(_rv("", "TOTAL ASSETS", total_assets, is_total=True))
    sections.append({'title': 'ASSETS', 'rows': asset_rows})

    liab_rows = [_rv(c, d, a) for c, d, a in liabilities]
    liab_rows.append(_rv("", "TOTAL LIABILITIES", total_liabilities, is_total=True))
    sections.append({
        'title':      f"LIABILITIES AND {equity_label.upper()}",
        'sub_header': "Liabilities:",
        'rows':       liab_rows,
    })

    eq_rows  = ([_rv(c, d, a)  for c, d, a in equity_adds] +
                [_rv(c, d, -a) for c, d, a in equity_deds] +
                [_rv("", ni_label, net_income)])
    eq_rows.append(_rv("", f"TOTAL {equity_label.upper()}", total_equity, is_total=True))
    sections.append({'sub_header': f"{equity_label}:", 'rows': eq_rows})

    sections.append({'rows': [_rv("", f"TOTAL LIABILITIES AND {equity_label.upper()}",
                                  total_l_and_e, is_total=True)], 'spacer': False})

    structured = {
        'company':   company,
        'stmt_name': stmt_name.upper(),
        'subtitle':  f"As of {fmt_date(as_of)}",
        'sections':  sections,
        'warning':   (f"OUT OF BALANCE by {diff:,.2f} — check entries"
                      if abs(diff) > 0.005 else None),
    }
    analysis_data = {
        'total_assets':        total_assets,
        'total_liabilities':   total_liabilities,
        'total_equity':        total_equity,
        'net_income':          net_income,
        'current_assets':      total_assets,
        'current_liabilities': total_liabilities,
    }
    return analysis_data, screen_text, structured


# ---------------------------------------------------------------------------
# Financial Performance builder
# ---------------------------------------------------------------------------

def build_performance(
    trial_balance: list,
    stmt_name: str,
    company: str,
    from_d: QDate,
    to_d: QDate,
) -> tuple[dict, str, dict]:
    """
    Returns (analysis_data, screen_text, structured).

    analysis_data keys: total_revenue, total_cogs, gross_profit,
                        total_expenses, net_income
    """
    revenue  = []
    cogs     = []
    expenses = []

    for e in trial_balance:
        code   = e['account_code']
        desc   = e['account_description']
        amount = e['amount']
        suffix = code.rsplit('-', 1)[-1] if '-' in code else code
        first  = suffix[0] if suffix else ''
        if first == '4':
            revenue.append((code, desc, abs(amount)))
        elif first == '5':
            cogs.append((code, desc, amount))
        elif first in ('6', '7', '8', '9'):
            expenses.append((code, desc, amount))

    total_revenue  = sum(r[2] for r in revenue)
    total_cogs     = sum(c[2] for c in cogs)
    gross_profit   = total_revenue - total_cogs
    total_expenses = sum(ex[2] for ex in expenses)
    net_income     = gross_profit - total_expenses
    plabel         = period_label(from_d, to_d)

    fa      = fmt_amt
    W, AM   = 85, 15
    _line   = _make_line_fn(W, AM)
    row     = _row_fn(_line)
    tot_row = _total_fn(_line)

    lines = [
        _centre(company, W),
        _centre(stmt_name.upper(), W),
        _centre(plabel, W),
        "=" * W, "",
        "REVENUE", "-" * W,
    ]
    for code, desc, amt in revenue:
        lines.append(row(code, desc, fa(amt)))
    lines += ["-" * W, tot_row("TOTAL REVENUE", fa(total_revenue)), ""]

    if cogs:
        lines += ["COST OF GOODS SOLD", "-" * W]
        for code, desc, amt in cogs:
            lines.append(row(code, desc, fa(amt)))
        lines += ["-" * W, tot_row("TOTAL COST OF GOODS SOLD", fa(total_cogs)),
                  "", tot_row("GROSS PROFIT", fa(gross_profit)), ""]

    if expenses:
        lines += ["OPERATING EXPENSES", "-" * W]
        for code, desc, amt in expenses:
            lines.append(row(code, desc, fa(amt)))
        lines += ["-" * W, tot_row("TOTAL OPERATING EXPENSES", fa(total_expenses)), ""]

    lines += [tot_row("NET INCOME (LOSS)", fa(net_income)), "", "=" * W]
    screen_text = "\n".join(lines)

    # --- Structured (for PDF) ---
    def _rv(code, desc, amt, is_total=False, is_sub=False):
        return _r(fa, code, desc, amt, is_total, is_sub)

    sections = []

    rev_rows = [_rv(c, d, a) for c, d, a in revenue]
    rev_rows.append(_rv("", "TOTAL REVENUE", total_revenue, is_total=True))
    sections.append({'title': 'REVENUE', 'rows': rev_rows})

    if cogs:
        cogs_rows = [_rv(c, d, a) for c, d, a in cogs]
        cogs_rows.append(_rv("", "TOTAL COST OF GOODS SOLD", total_cogs, is_total=True))
        cogs_rows.append(_rv("", "GROSS PROFIT", gross_profit, is_sub=True))
        sections.append({'title': 'COST OF GOODS SOLD', 'rows': cogs_rows})

    if expenses:
        exp_rows = [_rv(c, d, a) for c, d, a in expenses]
        exp_rows.append(_rv("", "TOTAL OPERATING EXPENSES", total_expenses, is_total=True))
        sections.append({'title': 'OPERATING EXPENSES', 'rows': exp_rows})

    sections.append({'rows': [_rv("", "NET INCOME (LOSS)", net_income, is_total=True)],
                     'spacer': False})

    structured = {
        'company':   company,
        'stmt_name': stmt_name.upper(),
        'subtitle':  plabel,
        'sections':  sections,
        'warning':   None,
    }
    analysis_data = {
        'total_revenue':  total_revenue,
        'total_cogs':     total_cogs,
        'gross_profit':   gross_profit,
        'total_expenses': total_expenses,
        'net_income':     net_income,
    }
    return analysis_data, screen_text, structured