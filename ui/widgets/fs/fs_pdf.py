"""
ui/widgets/fs/fs_pdf.py
-----------------------
ReportLab PDF renderer for financial statements.

Public API:
    save_pdf(structured, path, title)
"""

from __future__ import annotations

import os

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.platypus.tables import Table, TableStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    _REPORTLAB_OK = True
except ImportError:
    _REPORTLAB_OK = False


def save_pdf(structured: dict, path: str, title: str = "Financial Statement") -> None:
    """
    Render *structured* as a formatted PDF at *path*.

    Raises RuntimeError if reportlab is not installed.
    Raises RuntimeError if rendering fails.

    structured keys:
        company   str
        stmt_name str
        subtitle  str
        sections  list[dict]   — see _build_position / _build_performance
        warning   str | None
    """
    if not _REPORTLAB_OK:
        raise RuntimeError("reportlab is not installed.\nRun:  pip install reportlab")

    doc = SimpleDocTemplate(
        path, pagesize=letter,
        leftMargin=0.65 * inch, rightMargin=0.65 * inch,
        topMargin=0.70 * inch,  bottomMargin=0.70 * inch,
        title=title,
    )

    C_NAVY  = colors.HexColor("#1a2a4a")
    C_RULE  = colors.HexColor("#1a2a4a")
    C_LIGHT = colors.HexColor("#f4f6fb")
    C_WARN  = colors.HexColor("#c62828")
    C_BODY  = colors.HexColor("#1c1c1c")
    C_MUTED = colors.HexColor("#555555")

    SANS   = "Helvetica"
    SANS_B = "Helvetica-Bold"

    def _ps(name, **kw):
        base = dict(fontName=SANS, fontSize=9, leading=13,
                    textColor=C_BODY, spaceAfter=0, spaceBefore=0)
        base.update(kw)
        return ParagraphStyle(name, **base)

    sty_company  = _ps("co",   fontName=SANS_B, fontSize=13, leading=17,
                               alignment=TA_CENTER, textColor=C_NAVY, spaceBefore=4)
    sty_stmtname = _ps("sn",   fontName=SANS_B, fontSize=10, leading=14,
                               alignment=TA_CENTER, textColor=C_NAVY)
    sty_subtitle = _ps("sub",  fontSize=9, leading=12,
                               alignment=TA_CENTER, textColor=C_MUTED, spaceAfter=6)
    sty_section  = _ps("sec",  fontName=SANS_B, fontSize=9, leading=13,
                               textColor=C_NAVY, spaceBefore=10, spaceAfter=2)
    sty_sub_hdr  = _ps("shdr", fontName=SANS_B, fontSize=8.5, leading=12,
                               textColor=C_MUTED, spaceBefore=4, spaceAfter=1)
    sty_warn     = _ps("warn", fontName=SANS_B, fontSize=8.5, leading=12,
                               textColor=C_WARN, spaceBefore=8)

    PW       = letter[0] - 1.30 * inch
    COL_CODE = 0.80 * inch
    COL_AMT  = 1.10 * inch
    COL_DESC = PW - COL_CODE - COL_AMT

    def _p(txt, style):
        return Paragraph(str(txt), style)

    def _hr(thick=0.5, clr=C_RULE):
        return HRFlowable(width="100%", thickness=thick, color=clr,
                          spaceAfter=2, spaceBefore=2)

    def _account_table(rows, indent=0):
        tbl_data, row_styles = [], []
        for i, row in enumerate(rows):
            code, desc, amt_str, is_total, is_sub = row
            if is_total:
                fn, fc, fs = SANS_B, C_NAVY, 9
            elif is_sub:
                fn, fc, fs = SANS_B, C_MUTED, 8.5
            else:
                fn, fc, fs = SANS, C_BODY, 9

            def _cp(t, align=TA_LEFT, **kw):
                return Paragraph(t, _ps(f"r{i}", fontName=fn, fontSize=fs,
                                        leading=fs * 1.4, textColor=fc,
                                        alignment=align, **kw))

            tbl_data.append([_cp(code or ""), _cp(desc or ""),
                              _cp(amt_str or "", align=TA_RIGHT)])
            if is_total:
                row_styles += [
                    ('LINEABOVE',   (0, i), (-1, i), 0.5, C_RULE),
                    ('LINEBELOW',   (0, i), (-1, i), 0.5, C_RULE),
                    ('TOPPADDING',  (0, i), (-1, i), 3),
                    ('BACKGROUND',  (0, i), (-1, i), C_LIGHT),
                ]
            elif is_sub:
                row_styles += [('TOPPADDING', (0, i), (-1, i), 4)]

        col_indent = indent * 0.18 * inch
        cw = [COL_CODE, COL_DESC - col_indent, COL_AMT]
        t = Table(tbl_data, colWidths=cw, repeatRows=0)
        t.setStyle(TableStyle([
            ('LEFTPADDING',   (0, 0), (-1, -1), 0),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ('TOPPADDING',    (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ] + row_styles))
        return t

    story = [
        _p(structured['company'],   sty_company),
        _p(structured['stmt_name'], sty_stmtname),
        _p(structured['subtitle'],  sty_subtitle),
        _hr(thick=1.5), Spacer(1, 4),
    ]
    for sec in structured['sections']:
        if sec.get('title'):
            story.append(_p(sec['title'], sty_section))
        if sec.get('sub_header'):
            story.append(_p(sec['sub_header'], sty_sub_hdr))
        if sec.get('rows'):
            story.append(_account_table(sec['rows'], indent=sec.get('indent', 0)))
        if sec.get('total_rows'):
            story.append(_account_table(sec['total_rows'], indent=0))
        if sec.get('spacer', True):
            story.append(Spacer(1, 4))

    if structured.get('warning'):
        story += [
            _hr(thick=1, clr=C_WARN),
            _p(structured['warning'], sty_warn),
            _hr(thick=1, clr=C_WARN),
        ]
    story += [Spacer(1, 6), _hr(thick=1.5)]
    doc.build(story)