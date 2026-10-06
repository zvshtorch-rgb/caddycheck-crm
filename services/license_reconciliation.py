"""Read-only license <-> payment reconciliation (no Streamlit dependency, unit-testable).

For every project this compares the CURRENT License EOP with the EOP that the actual invoice/payment
history justifies, so finance can review discrepancies before any date is changed. Nothing here writes.

Business rule (agreed wording: "if Year 1 is paid the license runs through the end of Year 1; if Year 2 is
also paid, through the end of Year 2 ..."):

  * License start date  = the project's activation date (fallback: 1 January of its installation year,
    flagged as estimated in the Notes column).
  * "Year k" invoices   = invoice rows labelled Yk in the `invoices` table (maintenance_year), excluding
    cancelled rows, credit rows, trial rows and `Complementary` adjustment invoices (extra-camera top-ups
    must not block or fake a license year).
  * Year k is PAID      = it has at least one such invoice and ALL of them are paid ("Yes"). Some paid and
    some unpaid is reported as "Partial" and does NOT count as paid.
  * Highest paid year N = the largest k such that Y1 ... Yk are ALL paid (consecutive from Y1). A gap (e.g.
    Y1 unpaid, Y2 paid) stops the chain; the later paid years are listed in Notes instead of being trusted.
  * Expected License EOP = license start date + N whole years (end of license year N, same day-of-month as
    the start date). No paid year -> no expected EOP.
  * Needs EOP update    = the current EOP is missing or earlier than the 1st day of the Expected EOP's month
    (the CRM routinely sets EOPs to the 1st of a month, so "1 Nov" covers an expected "19 Nov"). A current
    EOP that is LATER than expected (extension / grace period) is never flagged or proposed for lowering.
  * Confidence          = "High" when the billing years of the Y1..YN invoices fit the start date (each Yk
    billed within +/-1 year of start year + k - 1). "Low" means the invoice history does not line up with the
    activation date (legacy history, re-installation, batch invoices with reused labels): the suggested
    Expected EOP must be verified manually before anything is updated.
"""
from __future__ import annotations

import calendar
import datetime
from typing import Any, Callable, Iterable, Optional

from config.settings import canonical_project_name

MAX_YEAR = 12
_EXCLUDED_TYPES = {"complementary", "complimentary"}


def add_years(base: datetime.date, years: int) -> datetime.date:
    year = base.year + years
    day = min(base.day, calendar.monthrange(year, base.month)[1])  # 29 Feb -> 28 Feb in non-leap years
    return datetime.date(year, base.month, day)


def _covers(current: Optional[datetime.date], expected: datetime.date) -> bool:
    """True when ``current`` reaches the first day of the expected EOP's month."""
    return current is not None and current >= expected.replace(day=1)


def _as_date(value: Any) -> Optional[datetime.date]:
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    return None


def _year_label_number(label: Any) -> Optional[int]:
    text = str(label or "").strip().upper()
    if text.startswith("Y") and text[1:].isdigit():
        return int(text[1:])
    return None


def _status(invoice) -> str:
    return str(getattr(invoice, "paid", "") or "").strip().lower()


def _invoice_number(invoice) -> Optional[int]:
    try:
        return int(float(invoice.invoice_number))
    except (TypeError, ValueError):
        return None


def year_status(rows: list) -> str:
    """'Yes' | 'Partial' | 'No' | 'No invoice' for the invoice rows of one maintenance year."""
    if not rows:
        return "No invoice"
    paid = sum(1 for r in rows if _status(r) == "yes")
    if paid == len(rows):
        return "Yes"
    return "Partial" if paid else "No"


def _license_year_rows(invoices: Iterable, k: int) -> list:
    return [
        i for i in invoices
        if _year_label_number(getattr(i, "maintenance_year", "")) == k
        and _status(i) != "cancelled"
        and str(getattr(i, "invoice_type", "") or "").strip().lower() not in _EXCLUDED_TYPES
    ]


def _numbers(rows: list) -> str:
    return ", ".join(str(n) for n in sorted({n for n in map(_invoice_number, rows) if n is not None}))


def _last_payment_date(rows: list) -> str:
    dates = [d for d in (_as_date(getattr(r, "payment_date", None)) for r in rows if _status(r) == "yes") if d]
    return max(dates).isoformat() if dates else ""


def reconcile_project(project, invoices: list, normalize_status: Callable[[str], str] = lambda s: s or "") -> dict:
    notes: list[str] = []
    start = _as_date(getattr(project, "activation_date", None))
    start_estimated = False
    if start is None and getattr(project, "installation_year", None):
        start, start_estimated = datetime.date(int(project.installation_year), 1, 1), True
        notes.append("License start date estimated from installation year (no activation date)")

    by_year = {k: _license_year_rows(invoices, k) for k in range(1, MAX_YEAR + 1)}
    statuses = {k: year_status(rows) for k, rows in by_year.items()}

    highest = 0
    while highest < MAX_YEAR and statuses[highest + 1] == "Yes":
        highest += 1
    other_paid = [k for k in range(highest + 1, MAX_YEAR + 1) if statuses[k] == "Yes"]
    if other_paid:
        broken = f" (Y{highest + 1} is {statuses[highest + 1].lower()})" if highest + 1 not in other_paid else ""
        notes.append(f"Paid but not consecutive from Y1: {', '.join(f'Y{k}' for k in other_paid)}{broken}")
    for k in (1, 2):
        if statuses[k] == "Partial":
            notes.append(f"Y{k} is only partly paid")
    if not any(by_year.values()):
        has_other = any(_year_label_number(getattr(i, "maintenance_year", "")) and _status(i) != "cancelled" for i in invoices)
        notes.append("Only complementary/adjustment invoices on file" if has_other else "No Y-labelled invoices on file")

    # Does the billing history of the paid chain line up with the activation date?
    inconsistent = []
    if start:
        for k in range(1, highest + 1):
            for row in by_year[k]:
                billed = getattr(row, "year", None)
                if billed and abs(int(billed) - (start.year + k - 1)) > 1:
                    inconsistent.append(f"Y{k} billed {int(billed)}")
    confidence = "High" if not inconsistent and start else "Low"
    if inconsistent:
        shown = ", ".join(sorted(set(inconsistent))[:4])
        notes.append(f"Invoice history does not fit start date {start.isoformat()} ({shown}) - verify manually")

    current = _as_date(getattr(project, "license_eop", None))
    expected = add_years(start, highest) if (start and highest) else None
    if expected is None:
        needs_update, relation, gap_days = "No", "n/a (no paid year)", None
    elif current is None:
        needs_update, relation, gap_days = "Yes", "Missing EOP", None
    elif not _covers(current, expected):
        needs_update, relation, gap_days = "Yes", "Earlier than payments justify", (expected - current).days
    elif current <= expected.replace(day=calendar.monthrange(expected.year, expected.month)[1]):
        needs_update, relation, gap_days = "No", "Matches (same month)", max(0, (expected - current).days)
    else:
        needs_update, relation, gap_days = "No", "Later than payments justify (no action)", -(current - expected).days

    status_text = normalize_status(getattr(project, "status", ""))
    if needs_update == "Yes" and status_text.lower() == "cancelled":
        notes.append("Project is Cancelled - confirm an EOP update is wanted")

    def iso(d):
        return d.isoformat() if d else ""

    return {
        "Project": project.project_name,
        "Project Status": status_text,
        "Country": getattr(project, "country", "") or "",
        "License Start Date": iso(start) + (" (est.)" if start_estimated else ""),
        "Current License EOP": iso(current),
        "Y1 Invoice Number": _numbers(by_year[1]),
        "Y1 Paid?": statuses[1],
        "Y1 Payment Date": _last_payment_date(by_year[1]),
        "Y2 Invoice Number": _numbers(by_year[2]),
        "Y2 Paid?": statuses[2],
        "Y2 Payment Date": _last_payment_date(by_year[2]),
        "Highest Paid Maintenance Year": f"Y{highest}" if highest else "None",
        "Expected License EOP": iso(expected),
        "Needs EOP Update?": needs_update,
        "EOP vs Expected": relation,
        "Days Behind": gap_days,
        "Confidence": confidence if highest else "n/a",
        "Later Years (Y3+)": "; ".join(
            f"Y{k}: {statuses[k]}" + (f" ({_numbers(by_year[k])})" if by_year[k] else "")
            for k in range(3, MAX_YEAR + 1) if by_year[k]
        ),
        "Notes": "; ".join(notes),
        "_highest": highest,
        "_start": start,
        "_current": current,
    }


def build_license_reconciliation(projects: Iterable, invoices: Iterable, normalize_status: Callable[[str], str] = lambda s: s or "") -> list[dict]:
    """One reconciliation row per project (complete invoice history, matched by canonical project name)."""
    grouped: dict[str, list] = {}
    for invoice in invoices:
        key = canonical_project_name(getattr(invoice, "project_name", "")).lower().strip()
        grouped.setdefault(key, []).append(invoice)
    rows = [
        reconcile_project(p, grouped.get(canonical_project_name(p.project_name).lower().strip(), []), normalize_status)
        for p in projects
    ]
    return sorted(rows, key=lambda r: (r["Needs EOP Update?"] != "Yes", r["Confidence"] != "High", r["Project"].lower()))


def filter_reconciliation(
    rows: list[dict],
    *,
    project: Optional[str] = None,
    country: Optional[str] = None,
    project_status: Optional[str] = None,
    paid_year: Optional[int] = None,
    needs_update: Optional[str] = None,
) -> list[dict]:
    """Apply the Ask Data filters.

    ``paid_year=k``: Year k is paid (part of the consecutive paid chain Y1..Yk) AND the current EOP does not
    yet reach the month of start date + k years, i.e. the EOP was not updated for that paid year.
    """
    out = []
    for r in rows:
        if project and r["Project"] != project:
            continue
        if country and r["Country"] != country:
            continue
        if project_status and project_status != "All" and r["Project Status"] != project_status:
            continue
        if needs_update and r["Needs EOP Update?"].lower() != needs_update.lower():
            continue
        if paid_year is not None:
            start = r["_start"]
            if r["_highest"] < paid_year or start is None:
                continue
            if _covers(r["_current"], add_years(start, paid_year)):
                continue
        out.append(r)
    return out


PUBLIC_COLUMNS = [
    "Project", "Project Status", "Country", "License Start Date", "Current License EOP",
    "Y1 Invoice Number", "Y1 Paid?", "Y1 Payment Date", "Y2 Invoice Number", "Y2 Paid?", "Y2 Payment Date",
    "Highest Paid Maintenance Year", "Expected License EOP", "Needs EOP Update?", "EOP vs Expected",
    "Days Behind", "Confidence", "Later Years (Y3+)", "Notes",
]
