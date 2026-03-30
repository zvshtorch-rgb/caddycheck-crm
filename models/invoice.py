"""Invoice data model."""
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
