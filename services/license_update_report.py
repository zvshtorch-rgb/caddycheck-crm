"""Next-month license update report and PDF."""
import calendar
import datetime
from functools import lru_cache
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from services.ask_data.executor import _annual_payment_month, _latest_annual_invoice_year
from services.ask_data.helpers import _license_status, _project_license_date, _safe_int, _safe_str


REPORT_COLUMNS = (
    "Project", "Country", "Cameras", "License EOP", "Status", "Annual Payment Status",
)


def _report_payment_status(project, invoices, today: datetime.date) -> str:
    target_year = _latest_annual_invoice_year(_annual_payment_month(project), today)
    if target_year is None:
        return "No Annual Invoice Found"
    if project.installation_year and target_year < project.installation_year:
        return "Not Yet Due"

    project_key = _safe_str(project.project_name).lower().strip()
    states = [
        _safe_str(getattr(inv, "paid", None)).strip().lower()
        for inv in invoices
        if _safe_str(inv.project_name).lower().strip() == project_key
        and inv.year == target_year
    ]
    states = [state for state in states if state != "cancelled"]
    if not states:
        return "No Annual Invoice Found"
    if any(state not in {"yes", "no"} for state in states):
        return "Unknown"
    if "no" in states:
        return "Unpaid"
    return "Paid"


@lru_cache(maxsize=1)
def _report_font() -> TTFont:
    font = TTFont("LicenseUpdateDejaVuSans", str(Path(__file__).parent / "report_fonts" / "DejaVuSans.ttf"))
    pdfmetrics.registerFont(font)
    return font


def build_license_update_rows(projects, invoices, today: datetime.date) -> list[dict]:
    rows = []
    for project in projects:
        if _license_status(project, today) != "Update Next Month":
            continue
        eop = _project_license_date(project)
        rows.append({
            "Project": _safe_str(project.project_name),
            "Country": _safe_str(project.country),
            "Cameras": _safe_int(project.num_cams),
            "License EOP": eop.isoformat(),
            "Status": _safe_str(project.status),
            "Annual Payment Status": _report_payment_status(project, invoices, today),
        })
    return rows


def build_license_update_pdf(rows: list[dict], year: int, month: int) -> bytes:
    font = _report_font()
    for row in rows:
        for column in REPORT_COLUMNS:
            for char in str(row[column]):
                if char not in "\t\r\n" and (ord(char) > 0xFFFF or ord(char) not in font.face.charWidths):
                    raise ValueError(f"License update PDF font does not support character U+{ord(char):04X}")

    pdf_buf = BytesIO()
    doc = SimpleDocTemplate(
        pdf_buf, pagesize=landscape(A4),
        leftMargin=1.2 * cm, rightMargin=1.2 * cm,
        topMargin=1.2 * cm, bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    title = f"Projects Needing Update in {calendar.month_name[month]} {year} (cancelled excluded)"
    cell_style = styles["BodyText"].clone("LicenseUpdateCell", fontSize=8, leading=10, fontName=font.fontName)
    header_style = cell_style.clone("LicenseUpdateHeader", textColor=colors.white)
    data = [[Paragraph(escape(column), header_style) for column in REPORT_COLUMNS]]
    data.extend([
        [Paragraph(escape(str(row[column])), cell_style) for column in REPORT_COLUMNS]
        for row in rows
    ])
    widths = [doc.width * fraction for fraction in (0.28, 0.11, 0.09, 0.14, 0.12, 0.26)]
    table = Table(data, repeatRows=1, colWidths=widths)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1B3A6B")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#EBF5FB")]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDC3C7")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
    ]))
    doc.build([
        Paragraph(escape(title), styles["Title"]),
        Paragraph(
            f"Projects: {len(rows)} | Payment status: latest due annual invoice",
            styles["Normal"],
        ),
        Spacer(1, 0.4 * cm),
        table,
    ])
    return pdf_buf.getvalue()
