"""Generate a professional PDF invoice using reportlab."""
import datetime
from io import BytesIO
from pathlib import Path
from typing import List, Optional

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable,
)
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

from config.settings import (
    INVOICE_COMPANY_NAME,
    INVOICE_COMPANY_REG,
    INVOICE_BILL_TO_NAME,
    INVOICE_BILL_TO_ADDRESS_1,
    INVOICE_BILL_TO_ADDRESS_2,
    INVOICE_BILL_TO_VAT,
    INVOICE_BANK_DETAILS,
    INVOICE_PAYMENT_DAYS,
    normalize_month,
)
from models.project import Project
from services.invoice_service import get_invoice_preview_data, _determine_maintenance_year


# ── Colour palette ─────────────────────────────────────────────────────────────
BLUE_DARK  = colors.HexColor("#1B3A6B")
BLUE_MID   = colors.HexColor("#2D6A9F")
BLUE_LIGHT = colors.HexColor("#EBF5FB")
GREY_LINE  = colors.HexColor("#BDC3C7")
TEXT_DARK  = colors.HexColor("#2C3E50")
RED        = colors.HexColor("#E74C3C")
GREEN      = colors.HexColor("#27AE60")


def generate_invoice_pdf(
    projects: List[Project],
    month_name: str,
    year: int,
    invoice_number: Optional[int] = None,
) -> bytes:
    """
    Build a PDF invoice and return it as bytes.

    Parameters
    ----------
    projects : list of Project
        Projects for this month (already filtered).
    month_name : str
        Full month name, e.g. 'March'.
    year : int
        Invoice year.
    invoice_number : int, optional
        Invoice number to display; shown as '—' if None.

    Returns
    -------
    bytes
        Raw PDF bytes ready for st.download_button or email attachment.
    """
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
    )

    styles = getSampleStyleSheet()

    def _style(name, **kwargs):
        s = ParagraphStyle(name, parent=styles["Normal"], **kwargs)
        return s

    h1 = _style("h1", fontSize=22, textColor=BLUE_DARK, leading=26, spaceAfter=2)
    h2 = _style("h2", fontSize=11, textColor=BLUE_MID, leading=14, spaceBefore=6)
    normal = _style("body", fontSize=9, textColor=TEXT_DARK, leading=12)
    small  = _style("small", fontSize=8, textColor=colors.grey, leading=10)
    bold   = _style("bold", fontSize=9, textColor=TEXT_DARK, leading=12,
                    fontName="Helvetica-Bold")
    right  = _style("right", fontSize=9, textColor=TEXT_DARK, leading=12,
                    alignment=TA_RIGHT)
    center_h1 = _style("ch1", fontSize=22, textColor=BLUE_DARK, leading=26,
                        alignment=TA_CENTER)

    invoice_date = datetime.date.today()
    due_date     = invoice_date + datetime.timedelta(days=INVOICE_PAYMENT_DAYS)
    inv_no_str   = str(invoice_number) if invoice_number else "—"
    month_short  = month_name[:3]

    elems = []

    # ── Header: company + INVOICE title ───────────────────────────────────────
    header_data = [[
        Paragraph(f"<b>{INVOICE_COMPANY_NAME}</b>", h1),
        Paragraph("INVOICE", center_h1),
    ]]
    header_tbl = Table(header_data, colWidths=["50%", "50%"])
    header_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elems.append(header_tbl)
    elems.append(Paragraph(INVOICE_COMPANY_REG, small))
    elems.append(Spacer(1, 4 * mm))
    elems.append(HRFlowable(width="100%", thickness=2, color=BLUE_DARK))
    elems.append(Spacer(1, 4 * mm))

    # ── Meta block: Bill To | Invoice details ─────────────────────────────────
    meta_data = [[
        Paragraph("<b>Bill To:</b>", bold),
        "",
        Paragraph("<b>Invoice Details</b>", bold),
    ], [
        Paragraph(INVOICE_BILL_TO_NAME, normal),
        "",
        Paragraph(f"Invoice No.: <b>{inv_no_str}</b>", normal),
    ], [
        Paragraph(INVOICE_BILL_TO_ADDRESS_1, normal),
        "",
        Paragraph(f"Date: <b>{invoice_date.strftime('%d %B %Y')}</b>", normal),
    ], [
        Paragraph(INVOICE_BILL_TO_ADDRESS_2, normal),
        "",
        Paragraph(f"Due: <b>{due_date.strftime('%d %B %Y')}</b>", normal),
    ], [
        Paragraph(INVOICE_BILL_TO_VAT, normal),
        "",
        Paragraph(
            f"Description: <b>Iretailcheck – Maintenance – {month_short} {year}</b>",
            normal,
        ),
    ]]
    meta_tbl = Table(meta_data, colWidths=["45%", "5%", "50%"])
    meta_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    elems.append(meta_tbl)
    elems.append(Spacer(1, 6 * mm))

    # ── Line items table ───────────────────────────────────────────────────────
    preview = get_invoice_preview_data(projects, month_name, year)

    col_hdrs = ["Supermarket / Project", "Units", "Year", "Price (€)", "Line Total (€)"]
    tbl_data = [col_hdrs]
    subtotal = 0.0
    for row in preview:
        if row["rate"] == "TOTAL":
            subtotal = row["line_total"]
            continue
        tbl_data.append([
            row["project_name"],
            str(row["num_cams"]),
            row["maintenance_year"],
            f"€{row['rate']:,.0f}",
            f"€{row['line_total']:,.0f}",
        ])

    # Totals rows
    tbl_data.append(["", "", "", "", ""])           # spacer row
    tbl_data.append(["", "", "", "Subtotal", f"€{subtotal:,.0f}"])
    tbl_data.append(["", "", "", "VAT (0%)", "€0"])
    tbl_data.append(["", "", "", "TOTAL DUE", f"€{subtotal:,.0f}"])

    n_data   = len(tbl_data) - 1   # rows after header
    n_items  = len(preview) - 1 if preview else 0   # real project rows

    items_tbl = Table(
        tbl_data,
        colWidths=["42%", "10%", "12%", "18%", "18%"],
        repeatRows=1,
    )
    spacer_row = 1 + n_items + 1   # 0-based index of blank spacer row
    sub_row    = spacer_row + 1
    vat_row    = spacer_row + 2
    tot_row    = spacer_row + 3

    items_tbl.setStyle(TableStyle([
        # Header
        ("BACKGROUND",   (0, 0), (-1, 0), BLUE_DARK),
        ("TEXTCOLOR",    (0, 0), (-1, 0), colors.white),
        ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, 0), 9),
        ("ALIGN",        (1, 0), (-1, 0), "CENTER"),
        ("BOTTOMPADDING",(0, 0), (-1, 0), 6),
        ("TOPPADDING",   (0, 0), (-1, 0), 6),
        # Data rows
        ("FONTSIZE",     (0, 1), (-1, -1), 9),
        ("TEXTCOLOR",    (0, 1), (-1, -1), TEXT_DARK),
        ("ALIGN",        (1, 1), (-1, -1), "CENTER"),
        ("ALIGN",        (3, 1), (-1, -1), "RIGHT"),
        ("ROWBACKGROUNDS",(0, 1), (-1, n_items), [colors.white, BLUE_LIGHT]),
        ("GRID",         (0, 0), (-1, n_items), 0.5, GREY_LINE),
        ("BOTTOMPADDING",(0, 1), (-1, n_items), 4),
        ("TOPPADDING",   (0, 1), (-1, n_items), 4),
        # Totals
        ("FONTNAME",     (3, sub_row), (-1, sub_row), "Helvetica"),
        ("FONTNAME",     (3, tot_row), (-1, tot_row), "Helvetica-Bold"),
        ("FONTSIZE",     (3, tot_row), (-1, tot_row), 10),
        ("BACKGROUND",   (3, tot_row), (-1, tot_row), BLUE_DARK),
        ("TEXTCOLOR",    (3, tot_row), (-1, tot_row), colors.white),
        ("TOPPADDING",   (3, tot_row), (-1, tot_row), 5),
        ("BOTTOMPADDING",(3, tot_row), (-1, tot_row), 5),
        ("LINEABOVE",    (3, sub_row), (-1, sub_row), 1, BLUE_MID),
    ]))
    elems.append(items_tbl)
    elems.append(Spacer(1, 8 * mm))

    # ── Bank details ──────────────────────────────────────────────────────────
    elems.append(HRFlowable(width="100%", thickness=1, color=GREY_LINE))
    elems.append(Spacer(1, 3 * mm))
    elems.append(Paragraph("<b>Payment Details</b>", h2))

    bd = INVOICE_BANK_DETAILS
    bank_rows = [
        ["Account Name:", bd["account_name"],  "Bank:", bd["bank_name"]],
        ["Account No.:",  bd["account_no"],     "Branch:", bd["branch"]],
        ["SWIFT:",        bd["swift"],           "IBAN:", bd["iban"]],
        ["Address:",      bd["address"],         "", ""],
        ["Contact:",      bd["contact"],         "", ""],
        ["Website:",      bd["website"],         "", ""],
    ]
    bank_tbl = Table(bank_rows, colWidths=["18%", "32%", "15%", "35%"])
    bank_tbl.setStyle(TableStyle([
        ("FONTSIZE",     (0, 0), (-1, -1), 8),
        ("TEXTCOLOR",    (0, 0), (-1, -1), TEXT_DARK),
        ("FONTNAME",     (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",     (2, 0), (2, -1), "Helvetica-Bold"),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 2),
        ("TOPPADDING",   (0, 0), (-1, -1), 2),
        ("SPAN",         (1, 3), (-1, 3)),
        ("SPAN",         (1, 4), (-1, 4)),
        ("SPAN",         (1, 5), (-1, 5)),
    ]))
    elems.append(bank_tbl)
    elems.append(Spacer(1, 6 * mm))

    # ── Signature ─────────────────────────────────────────────────────────────
    sig_data = [[
        "",
        Paragraph(
            f"______________________<br/><b>{bd['signature_name']}</b><br/>"
            f"<font size='8'>{INVOICE_COMPANY_NAME}</font>",
            _style("sig", fontSize=9, alignment=TA_CENTER, textColor=TEXT_DARK),
        ),
    ]]
    sig_tbl = Table(sig_data, colWidths=["60%", "40%"])
    elems.append(sig_tbl)

    doc.build(elems)
    return buf.getvalue()
