"""Generate invoice PDFs matching the Natyanjani invoice template."""

import io
import os
from datetime import date
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER


# ---- Business details (edit here) ----
BUSINESS_NAME = "Natyanjani"
BUSINESS_ADDRESS = "Veenendaal, 3901 HB"
BUSINESS_KVK = "Kvk – 95637494"
BUSINESS_EMAIL = "angee.8493@gmail.com"
BUSINESS_BANK = "NL 74 INGB 011 2999 808"
BUSINESS_BANK_NAME = "Natyanjani"
PAYMENT_DAYS = 30

MONTH_NAMES_NL = {
    1: "Januari", 2: "Februari", 3: "Maart", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Augustus",
    9: "September", 10: "Oktober", 11: "November", 12: "December",
}


def generate_invoice_pdf(billing_entry):
    """Generate a PDF invoice for a single MonthlyBilling entry.

    Returns a BytesIO buffer containing the PDF.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=25 * mm, rightMargin=25 * mm,
        topMargin=20 * mm, bottomMargin=20 * mm,
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        "InvTitle", parent=styles["Heading1"],
        fontSize=22, textColor=HexColor("#1a1a2e"),
        spaceAfter=2 * mm,
    )
    heading_style = ParagraphStyle(
        "InvHeading", parent=styles["Normal"],
        fontSize=8, textColor=HexColor("#888888"),
        spaceAfter=1 * mm,
    )
    value_style = ParagraphStyle(
        "InvValue", parent=styles["Normal"],
        fontSize=10, textColor=HexColor("#1a1a2e"),
        spaceBefore=0, spaceAfter=0,
    )
    addr_style = ParagraphStyle(
        "InvAddr", parent=styles["Normal"],
        fontSize=10, textColor=HexColor("#333333"),
        leading=14,
    )
    footer_style = ParagraphStyle(
        "InvFooter", parent=styles["Normal"],
        fontSize=9, textColor=HexColor("#555555"),
        alignment=TA_CENTER,
        spaceBefore=8 * mm,
    )
    right_style = ParagraphStyle(
        "InvRight", parent=styles["Normal"],
        fontSize=10, alignment=TA_RIGHT,
    )

    elements = []
    b = billing_entry
    client = b.client
    btw_amt = round(b.amount * b.btw_rate / 100, 2)
    total = round(b.amount + btw_amt, 2)
    inv_date = b.paid_date or date.today()
    month_name = MONTH_NAMES_NL.get(b.month, str(b.month))

    # ---- Header: Logo + Business info ----
    # Get logo path
    logo_path = os.path.join(os.path.dirname(__file__), 'static', 'logo.png')
    logo_img = None
    if os.path.exists(logo_path):
        logo_img = Image(logo_path, width=50 * mm, height=50 * mm, kind='proportional')
    
    biz_info = Table(
        [
            [
                logo_img if logo_img else "",
                Paragraph(f"<b>{BUSINESS_NAME}</b>", ParagraphStyle(
                    "biz", parent=styles["Normal"], fontSize=14,
                    textColor=HexColor("#1a1a2e"), alignment=TA_RIGHT,
                )),
            ],
            [
                "",
                Paragraph(
                    f"{BUSINESS_ADDRESS}<br/>{BUSINESS_KVK}<br/>{BUSINESS_EMAIL}",
                    ParagraphStyle("bizd", parent=styles["Normal"], fontSize=9,
                                   textColor=HexColor("#666666"), alignment=TA_RIGHT, leading=13),
                ),
            ],
        ],
        colWidths=["50%", "50%"],
    )
    biz_info.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(biz_info)
    elements.append(Spacer(1, 8 * mm))

    # ---- Invoice date + number row ----
    meta_data = [
        [
            Paragraph("<font color='#888888' size='8'>Factuurdatum</font>", styles["Normal"]),
            Paragraph("<font color='#888888' size='8'>Factuur Nummer</font>", styles["Normal"]),
        ],
        [
            Paragraph(f"<b>{inv_date.strftime('%d-%b-%Y')}</b>", value_style),
            Paragraph(f"<b>{b.invoice_number}</b>", value_style),
        ],
    ]
    meta_table = Table(meta_data, colWidths=["50%", "50%"])
    meta_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 1),
        ("TOPPADDING", (0, 1), (-1, 1), 1),
    ]))
    elements.append(meta_table)
    elements.append(Spacer(1, 6 * mm))

    # ---- Client address block ----
    elements.append(Paragraph("<font color='#888888' size='8'>Factuur Adres,</font>", styles["Normal"]))
    elements.append(Spacer(1, 1 * mm))

    client_addr = client.name
    if client.address:
        client_addr += f"<br/>{client.address}"
    if client.email:
        client_addr += f"<br/>{client.email}"
    if client.phone:
        client_addr += f"<br/>{client.phone}"
    elements.append(Paragraph(client_addr, addr_style))
    elements.append(Spacer(1, 10 * mm))

    # ---- Line items table ----
    accent = HexColor("#1a1a2e")
    header_bg = HexColor("#f0f0f5")

    category_map = {
        "adult": "Dance Classes - Adults",
        "child": "Dance Classes - Kids",
        "child_below_6_5": "Dance Classes - Kids Below 6.5 Years",
        "individual": "Dance Classes - Individual",
        "online": "Dance Classes - Online",
    }
    service_desc = category_map.get(client.student_type, "Dance Classes")

    table_data = [
        [
            Paragraph("<b>SERVICE</b>", ParagraphStyle("th", fontSize=9, textColor=accent)),
            Paragraph("<b>DATUM</b>", ParagraphStyle("th", fontSize=9, textColor=accent)),
            Paragraph("<b>BEDRAG</b>", ParagraphStyle("th", fontSize=9, textColor=accent, alignment=TA_RIGHT)),
            Paragraph("<b>TOTAAL</b>", ParagraphStyle("th", fontSize=9, textColor=accent, alignment=TA_RIGHT)),
        ],
        [
            Paragraph(f"{service_desc}<br/><font size='8' color='#666666'>{month_name} {b.year}</font>", styles["Normal"]),
            Paragraph(inv_date.strftime("%d-%b-%Y"), styles["Normal"]),
            Paragraph(f"€ {b.amount:.2f}", right_style),
            Paragraph(f"€ {b.amount:.2f}", right_style),
        ],
    ]

    if btw_amt > 0:
        table_data.append([
            Paragraph(f"BTW ({b.btw_rate:.0f}%)", ParagraphStyle("btw", fontSize=9, textColor=HexColor("#666666"))),
            "",
            "",
            Paragraph(f"€ {btw_amt:.2f}", right_style),
        ])

    # Totals row
    table_data.append([
        "",
        "",
        Paragraph("<b>TOTAAL</b>", ParagraphStyle("tot", fontSize=11, textColor=accent, alignment=TA_RIGHT)),
        Paragraph(f"<b>€ {total:.2f}</b>", ParagraphStyle("totv", fontSize=11, textColor=accent, alignment=TA_RIGHT)),
    ])

    col_widths = ["35%", "20%", "20%", "25%"]
    page_w = A4[0] - 50 * mm
    col_widths_abs = [page_w * float(w.strip("%")) / 100 for w in col_widths]

    items_table = Table(table_data, colWidths=col_widths_abs)
    items_table.setStyle(TableStyle([
        # Header row
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
        # Grid lines
        ("LINEBELOW", (0, 0), (-1, 0), 1, accent),
        ("LINEABOVE", (0, -1), (-1, -1), 1, accent),
        # General
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 1), (-1, -2), 8),
        ("BOTTOMPADDING", (0, 1), (-1, -2), 8),
        ("TOPPADDING", (0, -1), (-1, -1), 10),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 10),
    ]))
    elements.append(items_table)

    # ---- Payment footer ----
    elements.append(Spacer(1, 12 * mm))
    elements.append(Paragraph(
        f"Ik verzoek je vriendelijk het totaalbedrag binnen {PAYMENT_DAYS} dagen over te maken op: "
        f"<b>{BUSINESS_BANK}</b> ten name van: <b>{BUSINESS_BANK_NAME}</b>",
        footer_style,
    ))

    doc.build(elements)
    buf.seek(0)
    return buf
