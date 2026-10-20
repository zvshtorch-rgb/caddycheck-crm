"""Country breakdown of ledger amounts that always reconciles to the total.

Every row lands in exactly one bucket: a project country, "(country not set)" for a matched project without a
country, or the separate "Unassigned (no matching project)" bucket for ledger names the resolver cannot attach to a
project. The unassigned bucket is never merged away or dropped, so country totals + unassigned == total, to the cent.
Read-only; sums use Decimal so the reconciliation is exact.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Callable, Optional

import pandas as pd

from services.ask_data.resolver import ProjectIndex

UNASSIGNED_LABEL = "Unassigned (no matching project)"
NO_COUNTRY_LABEL = "(country not set)"


def _dec(value) -> Decimal:
    return Decimal(str(value if value is not None else 0))


@dataclass
class Bucket:
    label: str
    amount: Decimal = Decimal("0")
    rows: list = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return len(self.rows)


@dataclass
class DebtBreakdown:
    total: Decimal
    row_count: int
    countries: list  # list[Bucket], one per country, sorted by label
    no_country: Bucket
    unassigned: Bucket

    @property
    def buckets(self) -> list:
        extra = [b for b in (self.no_country, self.unassigned) if b.rows]
        return list(self.countries) + extra

    @property
    def reconciles(self) -> bool:
        return sum((b.amount for b in self.buckets), Decimal("0")) == self.total and \
            sum(b.row_count for b in self.buckets) == self.row_count

    def unassigned_names(self) -> list[str]:
        return sorted({getattr(r, "project_name", "") for r in self.unassigned.rows})

    def to_dataframe(self) -> pd.DataFrame:
        records = [{"Country": b.label, "Rows": b.row_count, "Amount (€)": float(b.amount)} for b in self.buckets]
        records.append({"Country": "Total", "Rows": self.row_count, "Amount (€)": float(self.total)})
        return pd.DataFrame(records)


def debt_by_country(
    invoices,
    projects,
    *,
    row_filter: Optional[Callable] = None,
    year: Optional[int] = None,
    aliases: Optional[dict] = None,
) -> DebtBreakdown:
    """Outstanding (unpaid) amount by country. `row_filter` overrides the default `inv.is_unpaid()`."""
    index = ProjectIndex(projects, aliases)
    predicate = row_filter or (lambda inv: inv.is_unpaid())
    rows = [inv for inv in invoices if predicate(inv) and (year is None or inv.year == year)]

    by_country: dict[str, Bucket] = {}
    no_country, unassigned = Bucket(NO_COUNTRY_LABEL), Bucket(UNASSIGNED_LABEL)
    for inv in rows:
        project = index.project(inv.project_name)
        if project is None:
            bucket = unassigned
        elif not project.country:
            bucket = no_country
        else:
            bucket = by_country.setdefault(project.country, Bucket(project.country))
        bucket.amount += _dec(inv.payment_amount)
        bucket.rows.append(inv)
    total = sum((_dec(inv.payment_amount) for inv in rows), Decimal("0"))
    return DebtBreakdown(total, len(rows), [by_country[k] for k in sorted(by_country)], no_country, unassigned)
