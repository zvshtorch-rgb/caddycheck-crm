"""Safe save flow for bank payments (no Streamlit dependency, so it is unit-testable).

Why this exists: the old "Save All Parsed Payments" flow marked invoice rows paid FIRST and then wrapped
the bank_payments insert in a broad ``except Exception`` that silently fell back to a local JSON file.
On Streamlit Cloud that file is wiped on every restart, so a failed insert meant the invoices showed as
paid while the payment never appeared under Saved Bank Payments -- with no error shown anywhere.

Rules enforced here:
  * The bank_payments record (+ planned allocations) is saved to Supabase BEFORE any invoice row is
    touched. If that save fails nothing else is modified (no partial state).
  * Every failure is logged with its traceback and returned as an explicit error message.
  * A local-JSON copy is only ever a best-effort backup; it never counts as a successful save.
  * If some invoice rows cannot be marked paid afterwards, the Supabase payment record is re-saved so its
    allocations list exactly the rows that really were marked paid, and the failure is reported.
"""
from __future__ import annotations

import datetime
import logging
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Optional

logger = logging.getLogger(__name__)


def is_row_payable(row: dict) -> bool:
    """The single rule deciding whether an invoice row may be marked paid by a bank payment.

    Only rows whose status is "No" (a missing, blank or whitespace-only status counts as "No") are
    payable. "Yes" (already paid), "cancelled" and any unrecognised status are NOT payable, so they are
    never marked paid and never appear in allocations or applied_amount. Used by every payment flow.
    """
    return _status_of(row) in ("", "no")


def _status_of(row: dict) -> str:
    return str(row.get("paid") or "").strip().lower()


def _num(value: Any) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if number != number else number  # NaN -> 0.0


def filter_payable_rows(rows: Iterable[dict], *, positive_amount_only: bool = False) -> list[dict]:
    """Rows that pass is_row_payable (optionally also with a payment_amount > 0).

    Every flow that lists, selects, credits or pays invoice rows must go through this so the
    eligibility rule can never differ between the Save All, Manual Lookup, Partial and Apply Credit flows.
    """
    return [
        r for r in rows
        if is_row_payable(r) and (not positive_amount_only or _num(r.get("payment_amount")) > 0)
    ]


def plan_partial_payments(edited_rows: Iterable[dict], payable_ids: set[int]) -> list[dict]:
    """Turn the Partial tab's editor records into payment plans.

    ``payable_ids`` are the ids of rows that passed is_row_payable when the tab was rendered; any
    editor record with another id (cancelled, already paid, unknown status, tampered) is ignored.
    """
    plans = []
    for rec in edited_rows:
        try:
            row_id = int(rec["id"])
        except (KeyError, TypeError, ValueError):
            continue
        if row_id not in payable_ids:
            continue
        original = _num(rec.get("Original (€)"))
        paid_now = max(0.0, _num(rec.get("Paid Now (€)")))
        if original <= 0 or paid_now <= 0:
            continue
        paid_now = min(original, paid_now)
        plans.append({
            "id": row_id,
            "project": str(rec.get("Project") or ""),
            "maintenance_year": str(rec.get("Maint. Year") or ""),
            "original": original,
            "paid_now": paid_now,
            "remaining": original - paid_now,
        })
    return plans


def apply_partial_payment_plans(
    plans: list[dict],
    *,
    invoice_number: int,
    payment_date: datetime.date,
    year: int,
    update_row: Callable[..., Any],
    insert_remainder_row: Callable[..., Any],
) -> list[dict]:
    """Write the partial payments; returns the allocation rows. Only rows named in ``plans`` are touched."""
    allocations = []
    for plan in plans:
        update_row(
            plan["id"],
            payment_amount=plan["paid_now"],
            paid="Yes",
            payment_date=payment_date,
            description=f"Partial settled: paid €{plan['paid_now']:,.2f} of €{plan['original']:,.2f}",
        )
        if plan["remaining"] > 0.005:
            insert_remainder_row(
                invoice_number=int(invoice_number),
                project_name=plan["project"],
                maintenance_year=plan["maintenance_year"],
                payment_amount=plan["remaining"],
                year=year,
                invoice_type="Complementary",
                description=f"Remaining debt after partial payment INV#{int(invoice_number)}",
            )
        allocations.append({
            "invoice_row_id": plan["id"],
            "invoice_number": int(invoice_number),
            "project_name": plan["project"],
            "maintenance_year": plan["maintenance_year"],
            "year": year,
            "amount_applied": plan["paid_now"],
        })
    return allocations


def plan_credit_application(target_rows: Iterable[dict], available_credit: float) -> list[dict]:
    """Spread ``available_credit`` over payable, positive-amount rows in order.

    Rows are re-checked with filter_payable_rows here, so a non-payable row can never be funded even if a
    caller passes it in. Returns [{row, use_amt, new_amt, fully_paid}] for rows that actually get credit.
    """
    plan = []
    remaining = _num(available_credit)
    for row in filter_payable_rows(target_rows, positive_amount_only=True):
        if remaining <= 0:
            break
        row_amount = _num(row.get("payment_amount"))
        use_amount = min(row_amount, remaining)
        new_amount = row_amount - use_amount
        plan.append({"row": row, "use_amt": use_amount, "new_amt": new_amount, "fully_paid": new_amount <= 0.005})
        remaining -= use_amount
    return plan


def apply_credit_plan(
    plan: list[dict],
    *,
    invoice_number: int,
    apply_date: datetime.date,
    update_row: Callable[..., Any],
) -> tuple[list[dict], float]:
    """Write the credit application; returns (allocation rows, total credit consumed)."""
    allocations = []
    consumed = 0.0
    for item in plan:
        row = item["row"]
        update_row(
            int(row["id"]),
            payment_amount=item["new_amt"],
            paid="Yes" if item["fully_paid"] else "No",
            payment_date=apply_date if item["fully_paid"] else None,
            description=f"Credit applied €{item['use_amt']:,.2f}",
        )
        allocations.append({
            "invoice_row_id": int(row["id"]),
            "invoice_number": int(invoice_number),
            "project_name": str(row.get("project_name") or ""),
            "maintenance_year": str(row.get("maintenance_year") or ""),
            "year": int(_num(row.get("year"))) or None,
            "amount_applied": item["use_amt"],
        })
        consumed += item["use_amt"]
    return allocations, consumed


@dataclass
class PaymentSaveOutcome:
    payment_saved: bool = False  # bank_payments record durably saved in Supabase
    errors: list[str] = field(default_factory=list)  # visible problems; any entry means NOT fully successful
    notes: list[str] = field(default_factory=list)  # informational, non-failure messages
    marked_rows: int = 0
    confirmed_invoice_numbers: set[int] = field(default_factory=set)

    @property
    def fully_successful(self) -> bool:
        return self.payment_saved and not self.errors


def persist_bank_payment(
    entry: dict,
    allocations: list[dict],
    *,
    save_remote: Callable[[dict, list[dict]], Any],
    save_local_backup: Optional[Callable[[dict], Any]] = None,
) -> Optional[str]:
    """Save to Supabase. Returns None on success, otherwise a user-presentable error message."""
    try:
        save_remote(entry, allocations)
        return None
    except Exception as exc:
        label = entry.get("source_name") or entry.get("invoice_number") or "payment"
        logger.exception("Could not save bank payment %r to Supabase", label)
        backup_note = ""
        if save_local_backup is not None:
            try:
                save_local_backup({**entry, "allocations": allocations})
                backup_note = (
                    " A local backup copy was written, but it is NOT a durable save on Streamlit Cloud "
                    "(the file is erased on every restart)."
                )
            except Exception:
                logger.exception("Could not write local backup for bank payment %r", label)
        return f"{label}: the bank payment record was NOT saved to Supabase ({type(exc).__name__}: {exc}).{backup_note}"


def save_parsed_bank_payment(
    *,
    payment_entry: dict,
    invoice_numbers: Iterable[int],
    payment_date: datetime.date,
    get_invoice_rows: Callable[[int], list[dict]],
    mark_row_paid: Callable[..., Any],
    save_remote: Callable[[dict, list[dict]], Any],
    save_local_backup: Optional[Callable[[dict], Any]] = None,
    on_row_paid: Optional[Callable[[dict], Optional[str]]] = None,
) -> PaymentSaveOutcome:
    """
    Save one parsed bank payment and settle its invoice rows.

    ``payment_entry`` carries the common fields (payment_date, source_*, fingerprint, amounts, raw text,
    storage metadata). ``invoice_number``, ``applied_amount`` and ``notes`` are set here.
    ``on_row_paid(row)`` runs after each row is marked paid (renewal links etc.); it may return a note and
    may raise -- a raise is reported as an error but does not undo the row.
    """
    out = PaymentSaveOutcome()
    label = payment_entry.get("source_name") or "payment"
    inv_nos = sorted({int(n) for n in invoice_numbers if n and int(n) > 0})
    base = {**payment_entry, "invoice_number": ",".join(str(n) for n in inv_nos) or None}

    def save(entry: dict, allocations: list[dict]) -> bool:
        error = persist_bank_payment(
            entry, allocations, save_remote=save_remote, save_local_backup=save_local_backup
        )
        if error:
            out.errors.append(error)
            return False
        out.payment_saved = True
        return True

    if not inv_nos:
        if save({**base, "applied_amount": None, "notes": "Auto-saved from batch upload without invoice number."}, []):
            out.notes.append(f"{label}: saved without invoice match")
        return out

    rows: list[dict] = []
    try:
        for n in inv_nos:
            for row in get_invoice_rows(n):
                rows.append({**row, "_invoice_number": n})
    except Exception as exc:
        logger.exception("Could not look up invoice rows for %r", label)
        out.errors.append(f"{label}: could not look up invoice rows ({type(exc).__name__}: {exc}). Nothing was saved.")
        return out

    if not rows:
        if save({**base, "applied_amount": None, "notes": "Auto-saved from batch upload but no invoice rows were found."}, []):
            out.notes.append(f"{label}: no invoice rows found")
        return out

    unpaid = [r for r in rows if is_row_payable(r)]
    cancelled_count = sum(1 for r in rows if _status_of(r) == "cancelled")
    if cancelled_count:
        out.notes.append(f"{label}: {cancelled_count} cancelled invoice row(s) were excluded from this payment")
    if not unpaid:
        if cancelled_count:
            note, detail = "Auto-saved from batch upload; no payable rows (already paid or cancelled).", "no payable rows"
        else:
            note, detail = "Auto-saved from batch upload; invoice rows were already paid.", "rows already paid"
        if save({**base, "applied_amount": None, "notes": note}, []):
            out.notes.append(f"{label}: {detail}")
        return out

    def allocation_for(row: dict) -> dict:
        return {
            "invoice_row_id": int(row["id"]),
            "invoice_number": row.get("_invoice_number"),
            "project_name": row["project_name"],
            "maintenance_year": str(row.get("maintenance_year") or ""),
            "year": int(row["year"]) if row.get("year") not in (None, "", 0) else None,
            "amount_applied": float(row.get("payment_amount") or 0.0),
        }

    # Step 1: durable payment record first. If this fails no invoice row has been touched yet.
    planned = [allocation_for(r) for r in unpaid]
    planned_total = sum(a["amount_applied"] for a in planned)
    if not save({**base, "applied_amount": planned_total}, planned):
        out.errors.append(f"{label}: no invoice rows were marked paid, so nothing is half-saved. Please retry.")
        return out

    # Step 2: mark rows paid.
    done: list[tuple[dict, dict]] = []
    failed: list[str] = []
    for row in unpaid:
        project = row["project_name"]
        try:
            mark_row_paid(db_id=int(row["id"]), payment_date=payment_date, payment_amount=row.get("payment_amount"))
        except Exception as exc:
            logger.exception("Could not mark invoice row id=%s (%s) paid", row.get("id"), project)
            failed.append(f"{project} (invoice {row.get('_invoice_number')}): {type(exc).__name__}: {exc}")
            continue
        done.append((row, allocation_for(row)))
        out.marked_rows += 1
        if row.get("_invoice_number"):
            out.confirmed_invoice_numbers.add(int(row["_invoice_number"]))
        if on_row_paid is not None:
            try:
                note = on_row_paid(row)
                if note:
                    out.notes.append(note)
            except Exception as exc:
                logger.exception("Post-payment step failed for %s", project)
                out.errors.append(f"{project}: marked paid, but a follow-up step failed ({type(exc).__name__}: {exc}).")

    # Step 3: if any row could not be marked, make the saved payment record match reality.
    if failed:
        out.errors.append(f"{label}: these invoice rows could NOT be marked paid: " + "; ".join(failed))
        actual = [alloc for _, alloc in done]
        actual_total = sum(a["amount_applied"] for a in actual)
        note = (
            "Auto-saved from batch upload; only part of the invoice rows could be marked paid "
            f"({len(done)} of {len(unpaid)}) - see error log."
        )
        if not save({**base, "applied_amount": actual_total, "notes": note}, actual):
            out.errors.append(
                f"{label}: the saved payment record still lists allocations for rows that were not marked paid - "
                "please review it under Saved Bank Payments."
            )
    return out
