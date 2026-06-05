"""給与明細印刷用 PDF（A4縦・1人1ページ）を生成する。"""

from io import BytesIO
from typing import Optional, Set

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from excel_export import _deduction_total, _get_amount, _pay_items

FONT = "HeiseiKakuGo-W5"
pdfmetrics.registerFont(UnicodeCIDFont(FONT))

PAGE_W, PAGE_H = A4
MARGIN = 15 * mm

TITLE_STYLE = ParagraphStyle("Title", fontName=FONT, fontSize=16, alignment=TA_CENTER, leading=20)
HEADER_STYLE = ParagraphStyle("Header", fontName=FONT, fontSize=11, alignment=TA_LEFT, leading=14)
NORMAL_STYLE = ParagraphStyle("Normal", fontName=FONT, fontSize=10, alignment=TA_LEFT, leading=13)
NET_STYLE = ParagraphStyle("Net", fontName=FONT, fontSize=12, alignment=TA_LEFT, leading=15)

TABLE_W = PAGE_W - 2 * MARGIN
LABEL_W = TABLE_W * 0.62
AMOUNT_W = TABLE_W * 0.38

def _grid_style(extra=None):
    style = [
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
    ]
    if extra:
        style.extend(extra)
    return TableStyle(style)


def _money(value) -> str:
    return f"{int(value):,}"


def _info_table(record: dict) -> Table:
    rows = [
        ["月分", str(record.get("月分", ""))],
        ["氏名", str(record.get("氏名", ""))],
        ["作成日", str(record.get("作成日", ""))],
        ["支給日", str(record.get("支給日", ""))],
    ]
    table = Table(rows, colWidths=[LABEL_W, AMOUNT_W])
    table.setStyle(_grid_style())
    return table


def _amount_table(
    rows: list,
    *,
    net_row: Optional[int] = None,
    bold_rows: Optional[Set[int]] = None,
) -> Table:
    bold_rows = bold_rows or set()
    data = [[label, _money(amount)] for label, amount in rows]
    table = Table(data, colWidths=[LABEL_W, AMOUNT_W])
    extra = []
    for idx in bold_rows:
        extra.append(("FONTSIZE", (0, idx), (-1, idx), 11))
    if net_row is not None:
        extra.append(("FONTSIZE", (0, net_row), (-1, net_row), 12))
    table.setStyle(_grid_style(extra))
    return table


def _block_header(title: str) -> Table:
    table = Table([[title, ""]], colWidths=[LABEL_W, AMOUNT_W])
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), FONT),
                ("FONTSIZE", (0, 0), (-1, -1), 11),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("SPAN", (0, 0), (1, 0)),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def _remark_table(record: dict) -> list:
    remark = str(record.get("備考", "") or "")
    header = _block_header("【備考】")
    box = Table([[remark]], colWidths=[TABLE_W], rowHeights=[18 * mm])
    box.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), FONT),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return [header, box]


def _build_story_for_record(record: dict) -> list:
    story = []
    story.append(Paragraph("給与明細書", TITLE_STYLE))
    story.append(Spacer(1, 6 * mm))
    story.append(_info_table(record))
    story.append(Spacer(1, 4 * mm))

    story.append(_block_header("【支給】"))
    pay_rows = _pay_items(record)
    pay_rows.append(("支給合計 A", _get_amount(record, "A_給与総額")))
    pay_table = _amount_table(pay_rows, bold_rows={len(pay_rows) - 1})
    story.append(pay_table)
    story.append(Spacer(1, 4 * mm))

    story.append(_block_header("【控除】"))
    deduct_rows = [
        ("健康保険", _get_amount(record, "健康保険")),
        ("厚生年金", _get_amount(record, "厚生年金")),
        ("雇用保険", _get_amount(record, "雇用保険")),
        ("子ども・子育て支援金", _get_amount(record, "子ども・子育て支援金")),
        ("その他社会保険", _get_amount(record, "その他社会保険")),
        ("源泉所得税", _get_amount(record, "源泉所得税", "所得税")),
        ("市町村民税", _get_amount(record, "市町村民税")),
        ("その他控除", _get_amount(record, "その他控除")),
        ("控除合計", _deduction_total(record)),
    ]
    deduct_table = _amount_table(deduct_rows, bold_rows={len(deduct_rows) - 1})
    story.append(deduct_table)
    story.append(Spacer(1, 4 * mm))

    story.append(_block_header("【差引・計算内訳】"))
    taxable_c = _get_amount(record, "C_課税対象給与")
    if taxable_c == 0:
        taxable_c = _get_amount(record, "A_給与総額")
    summary_rows = [
        ("課税対象給与 C", taxable_c),
        ("社会保険料計 D", _get_amount(record, "D_社会保険料計")),
        ("差引控除後給与額 E", _get_amount(record, "E_差引控除後給与額")),
        ("控除計 F", _get_amount(record, "F_控除計")),
        ("差引支給額 G", _get_amount(record, "G_差引支給額")),
    ]
    summary_table = _amount_table(summary_rows, net_row=len(summary_rows) - 1)
    story.append(summary_table)
    story.append(Spacer(1, 4 * mm))

    story.extend(_remark_table(record))
    return story


def generate_payslip_pdf(records: list) -> BytesIO:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=MARGIN,
        title="給与明細",
    )

    story = []
    for index, record in enumerate(records):
        if index > 0:
            story.append(PageBreak())
        story.extend(_build_story_for_record(record))

    doc.build(story)
    buffer.seek(0)
    return buffer
