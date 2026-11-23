"""Read-only previews for appending one full invoice row to a saved payment."""
from __future__ import annotations

import copy
import datetime as dt
from decimal import Decimal, InvalidOperation
import logging
import os
import re

from services.bank_payment_save import is_row_payable

MANUAL_ALLOCATION_GATE = "CADDYCHECK_MANUAL_ALLOCATION_ENABLED"
logger = logging.getLogger(__name__)


def manual_allocation_gate_state() -> dict[str, str]:
    value = os.environ.get(MANUAL_ALLOCATION_GATE)
    if value is None:
        import streamlit as st
        from streamlit.errors import StreamlitSecretNotFoundError

        try:
            value = st.secrets.get(MANUAL_ALLOCATION_GATE)
        except StreamlitSecretNotFoundError:
            logger.warning("Manual allocation release gate Secrets unavailable; disabled.")
    state = "UNSET" if value is None else (
        {"on": "ON", "off": "OFF"}.get(value.strip().lower(), "INVALID")
        if isinstance(value, str) else "INVALID"
    )
    if state == "INVALID":
        logger.warning("Manual allocation release gate has an invalid value; disabled.")
    return {"state": state, "effective_state": "ON" if state == "ON" else "OFF"}


class AllocationBlocked(ValueError):
    """An explicit eligibility or reconciliation failure."""


def require_admin(role: str) -> None:
    if role != "admin":
        raise AllocationBlocked("Only an admin can add an invoice to an existing transaction.")


def require_manual_allocation_enabled() -> None:
    if manual_allocation_gate_state()["effective_state"] != "ON":
        raise AllocationBlocked("Manual invoice allocation is disabled by its release gate.")


def money(value: object, label: str) -> Decimal:
    if value is None or isinstance(value, bool):
        raise AllocationBlocked(f"{label} is missing or invalid.")
    try:
        amount = Decimal(str(value))
    except InvalidOperation as exc:
        raise AllocationBlocked(f"{label} is invalid.") from exc
    if (not amount.is_finite() or abs(amount) >= Decimal("10000000000")
            or amount != amount.quantize(Decimal("0.01"))):
        raise AllocationBlocked(f"{label} must be a finite amount with at most two decimal places.")
    return amount


def invoice_number(value: object) -> int:
    text = str(value).strip()
    if not re.fullmatch(r"[0-9]+", text) or int(text) <= 0:
        raise AllocationBlocked("Enter one positive existing invoice number.")
    return int(text)


def summary_numbers(value: object) -> set[int]:
    if value in (None, "", 0, "0"):
        return set()
    tokens = str(value).split(",")
    numbers = [invoice_number(token) for token in tokens]
    if len(numbers) != len(set(numbers)):
        raise AllocationBlocked("Payment invoice summary contains duplicate numbers.")
    return set(numbers)


def payment_date(value: object) -> str:
    try:
        text = str(value)
        parsed = dt.date.fromisoformat(text)
    except ValueError as exc:
        raise AllocationBlocked("The saved payment date is missing or invalid.") from exc
    if parsed.isoformat() != text:
        raise AllocationBlocked("The saved payment date must use YYYY-MM-DD.")
    return text


def reconciliation(payment: dict, allocations: list[dict]) -> dict:
    instructed = money(payment.get("instructed_amount"), "Instructed amount")
    if instructed <= 0:
        raise AllocationBlocked("Instructed amount must be positive.")
    total = Decimal("0.00")
    numbers = set()
    row_ids = set()
    for allocation in allocations:
        if allocation.get("payment_id") != payment.get("id"):
            raise AllocationBlocked("An allocation belongs to a different payment.")
        amount = money(allocation.get("amount_applied"), "Existing allocation amount")
        if amount <= 0:
            raise AllocationBlocked("Existing allocation amounts must be positive.")
        row_id = allocation.get("invoice_row_id")
        if row_id is not None:
            if row_id in row_ids:
                raise AllocationBlocked("The saved payment already contains duplicate invoice-row allocations.")
            row_ids.add(row_id)
        numbers.add(invoice_number(allocation.get("invoice_number")))
        total += amount
    if money(payment.get("applied_amount"), "Saved allocated total") != total:
        raise AllocationBlocked("Saved allocated total disagrees with the actual allocation rows.")
    if summary_numbers(payment.get("invoice_number")) != numbers:
        raise AllocationBlocked("Saved invoice summary disagrees with the actual allocation rows.")
    if total > instructed:
        raise AllocationBlocked("Existing allocations exceed the instructed amount.")
    return {"instructed": instructed, "allocated": total, "remaining": instructed - total}


def choose_invoice(rows: list[dict], row_id: int | None = None) -> dict:
    if not rows:
        raise AllocationBlocked("No existing invoice rows found for this number.")
    if row_id is None:
        if len(rows) != 1:
            raise AllocationBlocked("This invoice number has multiple rows. Select one row explicitly.")
        return rows[0]
    matches = [row for row in rows if row.get("id") == row_id]
    if len(matches) != 1:
        raise AllocationBlocked("The selected invoice row is missing or ambiguous.")
    return matches[0]


def preview_allocation(payment: dict, allocations: list[dict], invoice: dict,
                       invoice_allocations: list[dict]) -> dict:
    balance = reconciliation(payment, allocations)
    date = payment_date(payment.get("payment_date"))
    if not isinstance(invoice.get("id"), int) or invoice["id"] <= 0:
        raise AllocationBlocked("The existing invoice must have a valid database row ID.")
    number = invoice_number(invoice.get("invoice_number"))
    for allocation in allocations + invoice_allocations:
        matches = allocation.get("invoice_row_id") == invoice["id"]
        legacy_match = (
            allocation.get("invoice_row_id") is None
            and str(allocation.get("invoice_number")) == str(number)
            and allocation.get("project_name") == invoice.get("project_name")
        )
        if matches or legacy_match:
            if allocation.get("payment_id") == payment["id"]:
                raise AllocationBlocked("This invoice row is already allocated to this payment.")
            raise AllocationBlocked("This invoice row is already allocated elsewhere; duplicate settlement is blocked.")
    if not is_row_payable(invoice):
        raise AllocationBlocked("This invoice row is already paid, cancelled, or otherwise non-payable.")
    amount = money(invoice.get("payment_amount"), "Invoice amount")
    if amount <= 0:
        raise AllocationBlocked("Only a positive full invoice amount can be allocated.")
    # The legacy invoices schema has no currency column and is an EUR-only ledger.
    currency = invoice["currency"] if "currency" in invoice else "EUR"
    if payment.get("currency") != "EUR" or currency != payment.get("currency"):
        raise AllocationBlocked("Payment and invoice currencies must both be EUR and match.")
    if amount > balance["remaining"]:
        raise AllocationBlocked("The full invoice amount exceeds the remaining instructed payment balance.")
    return {
        "payment_id": payment["id"], "invoice_row_id": invoice["id"],
        "invoice_number": number, "project": invoice.get("project_name"),
        "maintenance_year": invoice.get("maintenance_year"), "paid": invoice.get("paid"),
        "amount": str(amount), "payment_date": date,
        "remaining_before": str(balance["remaining"]),
        "remaining_after": str(balance["remaining"] - amount),
        "expected": copy.deepcopy({
            "payment": payment,
            "allocations": sorted(allocations, key=lambda row: row["id"]),
            "invoice": invoice,
        }),
    }
