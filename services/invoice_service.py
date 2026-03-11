"""Service for generating monthly invoice Excel files from the template."""
import logging
import shutil
import datetime
from pathlib import Path
from typing import List, Optional

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers

from config.settings import (
    INVOICE_TEMPLATE,
    OUTPUT_DIR,
    INVOICE_BILL_TO_NAME,
    INVOICE_BILL_TO_ADDRESS_1,
    INVOICE_BILL_TO_ADDRESS_2,
    INVOICE_BILL_TO_VAT,
    INVOICE_BANK_DETAILS,
    INVOICE_COMPANY_NAME,
    INVOICE_COMPANY_REG,
    INVOICE_PAYMENT_DAYS,
    RATE_Y1_PER_CAM,
    RATE_Y2_PLUS_PER_CAM,
    normalize_month,
)
from models.project import Project

logger = logging.getLogger(__name__)

# ── Column indices in the invoice template (1-based) ──────────────────────────
COL_SUPERMARKET = 1   # A  – project/store name
COL_UNITS = 7         # G  – number of cameras/units
COL_YEAR = 8          # H  – maintenance year label (Y1, Y2 …)
COL_PRICE = 9         # I  – price per unit
COL_LINE_TOTAL = 10   # J  – line total

# ── Fixed row numbers in the template header/footer ───────────────────────────
ROW_COMPANY_NAME = 3
ROW_DATE_LABEL = 4
ROW_INVOICE_NO = 5
ROW_PAYMENT_DUE = 8
ROW_COMPANY_REG = 9
ROW_BILL_TO_LABEL = 11
ROW_BILL_TO_NAME = 12
ROW_BILL_TO_ADDR1 = 13
ROW_BILL_TO_ADDR2 = 14
ROW_BILL_TO_VAT = 15
ROW_PROJECT_NAME_LABEL = 11   # column E
ROW_PROJECT_TITLE = 12        # column E – "Iretailcheck - Maintenance – Dec 2025"
ROW_HEADER = 19               # Supermarket | … | Units | Year | Price | Line Total
ROW_FIRST_DATA = 20           # First project data row in template

# Offset from last data row for footer sections
FOOTER_SUBTOTAL_OFFSET = 1
FOOTER_DISCOUNT_OFFSET = 2
FOOTER_LICENSE_OFFSET = 3
FOOTER_TAX_RATE_OFFSET = 3
FOOTER_TAX_OFFSET = 4
FOOTER_TOTAL_OFFSET = 5
FOOTER_BLANK_1 = 6
FOOTER_BLANK_2 = 7
FOOTER_BLANK_3 = 8
FOOTER_BLANK_4 = 9
FOOTER_BLANK_5 = 10
FOOTER_BANK_OFFSET = 11


def _determine_maintenance_year(
    project: Project,
    invoice_year: int,
) -> tuple:
    """
    Return (maintenance_year_int, rate, label) for a project in the given invoice year.

    Isolated helper – easy to adjust business rules here.
    """
    my = project.get_maintenance_year(invoice_year)
    label = f"Y{my}"
    rate = RATE_Y1_PER_CAM if my == 1 else RATE_Y2_PLUS_PER_CAM
    return my, rate, label


def generate_monthly_invoice(
    projects: List[Project],
    month_name: str,
    year: int,
    invoice_number: Optional[int] = None,
    output_dir: Path = None,
    template_path: Path = None,
) -> Path:
    """
    Generate a monthly invoice Excel file for the given month/year.

    Parameters
    ----------
    projects : list of Project
        Projects whose payment_month matches the selected month.
    month_name : str
        Full month name, e.g. 'December'.
    year : int
        Invoice year, e.g. 2025.
    invoice_number : int, optional
        Invoice number to use; auto-incremented from last template if None.
    output_dir : Path
        Directory to write the generated file.
    template_path : Path
        Path to the template .xlsx file.

    Returns
    -------
    Path
        Path to the generated Excel file.
    """
    from config.settings import get_data_paths
    paths = get_data_paths()
    if output_dir is None:
        output_dir = paths["output_dir"]
    if template_path is None:
        template_path = paths["invoice_template"]
    output_dir.mkdir(parents=True, exist_ok=True)
    month_abbr = month_name[:3]
    filename = f"CC_M_inv_{invoice_number or 'auto'}_{month_abbr}_{year}.xlsx"
    output_path = output_dir / filename

    # Copy template
    shutil.copy2(template_path, output_path)

    wb = openpyxl.load_workbook(output_path)
    # Use first sheet (template has one sheet named "July")
    ws = wb.worksheets[0]

    # ── Invoice date and due date ──────────────────────────────────────────────
    invoice_date = datetime.datetime.now().date()
    due_date = invoice_date + datetime.timedelta(days=INVOICE_PAYMENT_DAYS)

    # ── Header fields ──────────────────────────────────────────────────────────
    _set_cell(ws, ROW_DATE_LABEL, COL_LINE_TOTAL, invoice_date)
    if invoice_number:
        _set_cell(ws, ROW_INVOICE_NO, COL_LINE_TOTAL, invoice_number)
    _set_cell(ws, ROW_PAYMENT_DUE, COL_LINE_TOTAL, due_date)

    # Bill-to section
    _set_cell(ws, ROW_BILL_TO_NAME, 1, INVOICE_BILL_TO_NAME)
    _set_cell(ws, ROW_BILL_TO_ADDR1, 1, INVOICE_BILL_TO_ADDRESS_1)
    _set_cell(ws, ROW_BILL_TO_ADDR2, 1, INVOICE_BILL_TO_ADDRESS_2)
    _set_cell(ws, ROW_BILL_TO_VAT, 1, INVOICE_BILL_TO_VAT)

    # Invoice title (E12)
    title = f"Iretailcheck - Maintenance - {month_name[:3]} {year}"
    _set_cell(ws, ROW_PROJECT_TITLE, 5, title)

    # ── Clear existing data rows (rows 20 to max) ─────────────────────────────
    # We'll rebuild all project rows from row 20 onwards
    from openpyxl.cell.cell import MergedCell
    max_template_row = ws.max_row
    for r in range(ROW_FIRST_DATA, max_template_row + 1):
        for c in range(1, 11):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

    # ── Write project rows ────────────────────────────────────────────────────
    row = ROW_FIRST_DATA
    subtotal = 0.0

    # Sort projects by name
    sorted_projects = sorted(projects, key=lambda p: p.project_name)

    for proj in sorted_projects:
        if proj.num_cams <= 0:
            logger.warning("Skipping %s – 0 cameras", proj.project_name)
            continue

        _, rate, label = _determine_maintenance_year(proj, year)
        line_total = proj.num_cams * rate
        subtotal += line_total

        _set_cell(ws, row, COL_SUPERMARKET, proj.project_name)
        _set_cell(ws, row, COL_UNITS, proj.num_cams)
        _set_cell(ws, row, COL_YEAR, label)
        _set_cell(ws, row, COL_PRICE, rate)
        _set_cell(ws, row, COL_LINE_TOTAL, line_total)

        row += 1

    # ── Footer rows ───────────────────────────────────────────────────────────
    footer_start = row  # first row after last project

    # Subtotal
    _set_cell(ws, footer_start + 0, COL_PRICE, "Subtotal")
    _set_cell(ws, footer_start + 0, COL_LINE_TOTAL, subtotal)

    # Discount
    _set_cell(ws, footer_start + 1, COL_PRICE, "Discount")
    _set_cell(ws, footer_start + 1, COL_LINE_TOTAL, None)

    # License Period / Tax Rate
    license_date = datetime.datetime(year, 12, 1)
    _set_cell(ws, footer_start + 2, COL_SUPERMARKET, "License Period")
    _set_cell(ws, footer_start + 3, COL_SUPERMARKET, license_date)
    _set_cell(ws, footer_start + 2, COL_PRICE, "Tax/VAT Rate")
    _set_cell(ws, footer_start + 2, COL_LINE_TOTAL, 0)

    # Tax/VAT
    _set_cell(ws, footer_start + 3, COL_PRICE, "Tax/VAT ")
    _set_cell(ws, footer_start + 3, COL_LINE_TOTAL, 0)

    # Total
    _set_cell(ws, footer_start + 4, COL_PRICE, "Total")
    _set_cell(ws, footer_start + 4, COL_LINE_TOTAL, subtotal)

    # ── Bank details ──────────────────────────────────────────────────────────
    bank_row = footer_start + 11
    bd = INVOICE_BANK_DETAILS
    _set_cell(ws, bank_row, 1, "Bank details:")
    _set_cell(ws, bank_row, 2, f"Account name: {bd['account_name']}")
    _set_cell(ws, bank_row, COL_PRICE, "Signature")

    _set_cell(ws, bank_row + 1, 2, f"Account No:{bd['account_no']}")
    _set_cell(ws, bank_row + 1, COL_PRICE, bd["signature_name"])

    _set_cell(ws, bank_row + 2, 1, bd["website"])
    _set_cell(ws, bank_row + 2, 2, f"Bank name: {bd['bank_name']}")
    _set_cell(ws, bank_row + 2, COL_PRICE, invoice_date)

    _set_cell(ws, bank_row + 3, 2, f"SWIFT Code: {bd['swift']}")
    _set_cell(ws, bank_row + 4, 2, f"Branch: {bd['branch']}")
    _set_cell(ws, bank_row + 5, 2, f"IBAN Number: {bd['iban']}")

    _set_cell(ws, bank_row + 7, 1, "If you have any questions concerning this Invoice contact us")
    _set_cell(ws, bank_row + 9, 1, "Thank you for your business!")
    _set_cell(ws, bank_row + 10, 1, f",     {bd['website']}")
    _set_cell(ws, bank_row + 11, 1, bd["address"])
    _set_cell(ws, bank_row + 12, 1, bd["contact"])

    wb.save(output_path)
    logger.info("Invoice saved to %s", output_path)
    return output_path


def _set_cell(ws, row: int, col: int, value):
    """Set a cell value, skipping read-only merged cells."""
    from openpyxl.cell.cell import MergedCell
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value


def get_invoice_preview_data(
    projects: List[Project],
    month_name: str,
    year: int,
) -> list:
    """
    Return a list of dicts representing the invoice line items for preview.

    Each dict has: project_name, num_cams, maintenance_year, rate, line_total.
    """
    rows = []
    subtotal = 0.0
    for proj in sorted(projects, key=lambda p: p.project_name):
        if proj.num_cams <= 0:
            continue
        _, rate, label = _determine_maintenance_year(proj, year)
        line_total = proj.num_cams * rate
        subtotal += line_total
        rows.append({
            "project_name": proj.project_name,
            "num_cams": proj.num_cams,
            "maintenance_year": label,
            "rate": rate,
            "line_total": line_total,
        })
    rows.append({
        "project_name": "",
        "num_cams": "",
        "maintenance_year": "",
        "rate": "TOTAL",
        "line_total": subtotal,
    })
    return rows
