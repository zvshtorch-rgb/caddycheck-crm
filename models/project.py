"""Project data model."""
from dataclasses import dataclass, field
from typing import Optional
import datetime


@dataclass
class Project:
    """Represents a single project from 'Projects overview'."""
    project_name: str
    country: str = ""
    num_cams: int = 0
    payment_month: str = ""         # Normalized full month name, e.g. "December"
    installation_year: Optional[int] = None
    project_approval: str = ""
    activation_date: Optional[datetime.datetime] = None
    detection_type: str = ""
    cart_type: str = ""
    vim_version: str = ""
    status: str = ""
    license_eop: Optional[datetime.datetime] = None
    caddy_back: str = ""

    # Raw M-1Y..M-9Y values (invoice numbers assigned per maintenance year)
    maintenance_invoice_numbers: dict = field(default_factory=dict)

    # Per-project rate overrides (None = use global default)
    rate_y1_override: Optional[float] = None
    rate_y2_override: Optional[float] = None

    def get_maintenance_year(self, invoice_year: int) -> int:
        """Return maintenance year (1=Y1, 2=Y2, …) for a given invoice year."""
        if self.installation_year is None:
            return 1
        year = invoice_year - self.installation_year + 1
        return max(1, year)

    def get_rate(self, invoice_year: int) -> float:
        """Return price-per-camera for the given invoice year, respecting overrides."""
        from config.settings import RATE_Y1_PER_CAM, RATE_Y2_PLUS_PER_CAM
        my = self.get_maintenance_year(invoice_year)
        if my == 1:
            return self.rate_y1_override if self.rate_y1_override is not None else RATE_Y1_PER_CAM
        return self.rate_y2_override if self.rate_y2_override is not None else RATE_Y2_PLUS_PER_CAM

    def get_expected_amount(self, invoice_year: int) -> float:
        """Return expected invoice amount for the given year."""
        return self.num_cams * self.get_rate(invoice_year)

    def get_maintenance_year_label(self, invoice_year: int) -> str:
        """Return 'Y1', 'Y2', ... label."""
        return f"Y{self.get_maintenance_year(invoice_year)}"

    def is_active(self) -> bool:
        return str(self.status).strip().lower() == "active"
