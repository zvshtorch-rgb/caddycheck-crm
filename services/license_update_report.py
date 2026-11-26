"""Next-month license update report and PDF."""
import calendar
import datetime
from io import BytesIO
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from services.ask_data.executor import _annual_license_payment_row
from services.ask_data.helpers import _license_status, _project_license_date, _safe_int, _safe_str


REPORT_COLUMNS = (
    "Project", "Country", "Cameras", "License EOP", "Status", "Annual Payment Status",
)


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
            "Annual Payment Status": _annual_license_payment_row(project, eop, invoices, today)[
                "Annual Payment Status"
            ],
        })
    return rows


def build_license_update_pdf(rows: list[dict], year: int, month: int) -> bytes:
    pdf_buf = BytesIO()
    doc = SimpleDocTemplate(
        pdf_buf, pagesize=landscape(A4),
        leftMargin=1.2 * cm, rightMargin=1.2 * cm,
        topMargin=1.2 * cm, bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    title = f"Projects Needing Update in {calendar.month_name[month]} {year} (cancelled excluded)"
    cell_style = styles["BodyText"].clone("LicenseUpdateCell", fontSize=8, leading=10)
    header_style = cell_style.clone("LicenseUpdateHeader", textColor=colors.white, fontName="Helvetica-Bold")
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
