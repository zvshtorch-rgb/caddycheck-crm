"""Invoice data model."""
from collections import defaultdict
from dataclasses import dataclass
from typing import Optional
import datetime


@dataclass
class Invoice:
    """Represents a single invoice from 'Invoice details'."""
    invoice_number: Optional[float]
    project_name: str
    maintenance_year: str           # e.g. 'Y1', 'Y2', 'Paid Trial-0.5Y'
    payment_amount: float
    cameras_number: Optional[float]
    payment_date: Optional[datetime.datetime]
    paid: str                       # 'Yes', 'No', 'cancelled'
    year: Optional[int]

    def is_paid(self) -> bool:
        return str(self.paid).strip().lower() == "yes"

    def is_cancelled(self) -> bool:
        return str(self.paid).strip().lower() == "cancelled"

    def is_unpaid(self) -> bool:
        return str(self.paid).strip().lower() == "no"

    def maintenance_year_number(self) -> int:
        """Extract integer from 'Y1' -> 1, 'Y3' -> 3; returns 0 for trial/other."""
        label = str(self.maintenance_year).strip()
        if label.startswith("Y") and label[1:].isdigit():
            return int(label[1:])
        return 0

    def is_paid_trial_category(self) -> bool:
        """Return True for paid-trial invoice rows."""
        label = str(self.maintenance_year).strip().lower()
        return "paid trial" in label

    def is_new_installation_category(self) -> bool:
        """Return True for first-year debt classification."""
        label = str(self.maintenance_year).strip().lower()
        return label == "y1"

    def is_maintenance_category(self) -> bool:
        """Return True for maintenance debt classification (everything not Y1/trial)."""
        return not self.is_new_installation_category() and not self.is_paid_trial_category()


@dataclass
class DebtSummary:
    """Aggregated debt information for a single project."""
    project_name: str
    country: str
    num_cams: int
    payment_month: str
    installation_year: Optional[int]
    status: str

    # Financials
    total_expected: float = 0.0
    total_paid: float = 0.0
    total_cancelled: float = 0.0

    @property
    def total_unpaid(self) -> float:
        return max(0.0, self.total_expected - self.total_paid - self.total_cancelled)

    @property
    def debt(self) -> float:
        """Alias for unpaid."""
        return self.total_unpaid


@dataclass
class MonthlyInvoiceSummary:
    """Aggregated view of a combined monthly invoice spanning multiple projects."""

    invoice_number: int
    year: Optional[int]
    total_amount: float
    project_count: int
    invoice_rows: int
    paid_rows: int
    unpaid_rows: int
    cancelled_rows: int
    last_payment_date: Optional[datetime.datetime]
    project_names: list[str]

    @property
    def status(self) -> str:
        if self.invoice_rows <= 0:
            return "Unknown"
        if self.paid_rows == self.invoice_rows:
            return "Paid"
        if self.unpaid_rows == self.invoice_rows:
            return "Unpaid"
        if self.cancelled_rows == self.invoice_rows:
            return "Cancelled"
        if self.paid_rows > 0 and self.unpaid_rows > 0 and self.cancelled_rows == 0:
            return "Partial"
        if self.paid_rows > 0 and self.unpaid_rows == 0 and self.cancelled_rows > 0:
            return "Paid / Cancelled"
        if self.paid_rows == 0 and self.unpaid_rows > 0 and self.cancelled_rows > 0:
            return "Unpaid / Cancelled"
        return "Mixed"


def group_monthly_invoices(invoices: list[Invoice]) -> list[MonthlyInvoiceSummary]:
    """Group invoice rows that share a combined monthly invoice number."""

    grouped_invoices: dict[int, list[Invoice]] = defaultdict(list)
    for inv in invoices:
        if inv.invoice_number in (None, ""):
            continue
        try:
            invoice_number = int(float(inv.invoice_number))
        except (TypeError, ValueError):
            continue
        grouped_invoices[invoice_number].append(inv)

    summaries: list[MonthlyInvoiceSummary] = []
    for invoice_number, grouped_rows in grouped_invoices.items():
        project_names = sorted({
            str(inv.project_name).strip()
            for inv in grouped_rows
            if str(inv.project_name).strip()
        })
        if len(project_names) < 2:
            continue

        years = [inv.year for inv in grouped_rows if inv.year is not None]
        payment_dates = [inv.payment_date for inv in grouped_rows if inv.payment_date is not None]

        summaries.append(MonthlyInvoiceSummary(
            invoice_number=invoice_number,
            year=max(years) if years else None,
            total_amount=sum(inv.payment_amount or 0.0 for inv in grouped_rows),
            project_count=len(project_names),
            invoice_rows=len(grouped_rows),
            paid_rows=sum(1 for inv in grouped_rows if inv.is_paid()),
            unpaid_rows=sum(1 for inv in grouped_rows if inv.is_unpaid()),
            cancelled_rows=sum(1 for inv in grouped_rows if inv.is_cancelled()),
            last_payment_date=max(payment_dates) if payment_dates else None,
            project_names=project_names,
        ))

    summaries.sort(key=lambda summary: (summary.year or 0, summary.invoice_number), reverse=True)
    return summaries
