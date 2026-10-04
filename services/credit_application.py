"""Atomic "Apply Credit": a pure plan + two executors (no Streamlit dependency, unit-testable).

The whole operation -- consuming the credit source rows, writing the usage rows, updating/splitting the
target invoice rows and saving the credit-application payment record + allocations -- is described as
one ordered list of declarative operations (``CreditPlan.ops``). It is then executed either

  1. as ONE database transaction via the Postgres function ``apply_credit_atomic`` (see
     migrations/create_apply_credit_atomic.sql): every row is committed together or none is; or
  2. (only if that function is not installed yet) by a compensating executor that applies the operations
     one by one and, on any failure, undoes the ones already applied in reverse order.

Both executors enforce the same preconditions: a credit/target row is only touched while it still has the
amount it had when loaded and is still open, so stale data can never spend the same credit twice.
"""
from __future__ import annotations

import datetime
import hashlib
import logging
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Optional

from services.bank_payment_save import (
    CREDIT_EPSILON,
    _int_or_none,
    _num,
    available_credit_total,
    plan_credit_application,
    plan_credit_consumption,
)

logger = logging.getLogger(__name__)

_BEFORE_FIELDS = ("payment_amount", "paid", "payment_date", "description")


class CreditConflict(Exception):
    """A credit or target row changed (or is no longer open) since it was loaded."""


@dataclass
class CreditPlan:
    ops: list[dict] = field(default_factory=list)
    allocations: list[dict] = field(default_factory=list)  # one per TARGET row that receives credit
    credit_usage: list[dict] = field(default_factory=list)  # one per credit SOURCE row that is consumed
    consumed: float = 0.0
    applied: float = 0.0
    payment_entry: dict = field(default_factory=dict)

    def rpc_payload(self) -> dict:
        return {"ops": self.ops}


@dataclass
class CreditExecutionResult:
    ok: bool = False
    mode: str = ""  # "transaction" | "compensation"
    consumed: float = 0.0  # only non-zero when ok
    applied: float = 0.0  # only non-zero when ok
    allocations: list[dict] = field(default_factory=list)
    credit_usage: list[dict] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    rolled_back: bool = False  # every change was undone (or never committed)
    rollback_errors: list[str] = field(default_factory=list)  # something could NOT be undone -> inconsistent
    outcome_unknown: bool = False  # transport failure: the transaction may or may not have committed


def _before(row: dict) -> dict:
    return {name: row.get(name) for name in _BEFORE_FIELDS}


def plan_credit_application_ops(
    *,
    target_rows: Iterable[dict],
    credit_rows: Iterable[dict],
    invoice_number: int,
    apply_date: datetime.date,
    year: int,
) -> CreditPlan:
    """Build the full, deterministic operation list. Returns an empty plan when there is nothing to apply."""
    targets, credits = list(target_rows), list(credit_rows)
    available = available_credit_total(credits)
    target_plan = plan_credit_application(targets, available)
    need = round(sum(item["use_amt"] for item in target_plan), 2)
    plan = CreditPlan()
    if need <= CREDIT_EPSILON:
        return plan

    apply_iso = apply_date.isoformat()
    # 1. credit source rows (+ usage rows for partly used ones)
    for item in plan_credit_consumption(credits, need):
        row, use, remaining = item["row"], item["use_amt"], item["remaining_after"]
        old_description = str(row.get("description") or "").strip()
        note = f"Credit applied €{use:,.2f} to INV#{int(invoice_number)} on {apply_iso}"
        mark_paid = remaining <= CREDIT_EPSILON
        plan.ops.append({
            "op": "credit_consume",
            "id": int(row["id"]),
            "expected_amount": _num(row.get("payment_amount")),
            "use_amount": use,
            "new_amount": None if mark_paid else -remaining,
            "mark_paid": mark_paid,
            "payment_date": apply_iso,
            "description": f"{old_description} | {note}" if old_description else note,
            "before": _before(row),
        })
        if not mark_paid:
            plan.ops.append({"op": "insert_row", "row": {
                "invoice_number": _int_or_none(row.get("invoice_number")),
                "project_name": str(row.get("project_name") or ""),
                "maintenance_year": "Credit",
                "payment_amount": -use,
                "year": _int_or_none(row.get("year")) or year,
                "invoice_type": "Complementary",
                "description": f"Credit used: €{use:,.2f} applied to INV#{int(invoice_number)} (source credit row id={row['id']})",
                "paid": "Yes",
                "payment_date": apply_iso,
            }})
        plan.credit_usage.append({
            "credit_row_id": int(row["id"]), "used": use, "remaining_after": remaining, "usage_row_created": not mark_paid,
        })
        plan.consumed = round(plan.consumed + use, 2)

    # 2. target rows (+ remainder rows for partly funded ones)
    for item in target_plan:
        row, use = item["row"], item["use_amt"]
        row_amount = _num(row.get("payment_amount"))
        if item["fully_paid"]:
            fields = {"paid": "Yes", "payment_date": apply_iso, "description": f"Paid with credit €{use:,.2f}"}
        else:
            fields = {
                "payment_amount": use, "paid": "Yes", "payment_date": apply_iso,
                "description": f"Partially paid with credit: €{use:,.2f} of €{row_amount:,.2f}",
            }
        plan.ops.append({
            "op": "target_update", "id": int(row["id"]), "expected_amount": row_amount,
            "fields": fields, "before": _before(row),
        })
        if not item["fully_paid"]:
            plan.ops.append({"op": "insert_row", "row": {
                "invoice_number": int(invoice_number),
                "project_name": str(row.get("project_name") or ""),
                "maintenance_year": str(row.get("maintenance_year") or ""),
                "payment_amount": round(item["new_amt"], 2),
                "year": _int_or_none(row.get("year")) or year,
                "invoice_type": "Complementary",
                "description": f"Remaining debt after credit applied INV#{int(invoice_number)}",
                "paid": "No",
                "payment_date": None,
            }})
        plan.allocations.append({
            "invoice_row_id": int(row["id"]),
            "invoice_number": int(invoice_number),
            "project_name": str(row.get("project_name") or ""),
            "maintenance_year": str(row.get("maintenance_year") or ""),
            "year": _int_or_none(row.get("year")),
            "amount_applied": use,
        })
        plan.applied = round(plan.applied + use, 2)

    # 3. the audit/payment record -- last, so a failure anywhere earlier never leaves a "successful" record
    fingerprint = hashlib.sha256(
        (
            f"credit|{int(invoice_number)}|{apply_iso}|{plan.applied:.2f}|"
            + ",".join(str(a["invoice_row_id"]) for a in plan.allocations)
            + "|"
            + ",".join(str(u["credit_row_id"]) for u in plan.credit_usage)
        ).encode("utf-8")
    ).hexdigest()
    plan.payment_entry = {
        "payment_date": apply_iso,
        "invoice_number": str(int(invoice_number)),
        "source_name": f"credit-apply-{int(invoice_number)}",
        "source_kind": "credit-apply",
        "payment_fingerprint": fingerprint,
        "currency": "EUR",
        "applied_amount": plan.applied,
        "parsed_payload": {"credit_sources": plan.credit_usage},
        "notes": "Credit applied to unpaid rows (credit source rows consumed in place)",
    }
    plan.ops.append({"op": "save_payment", "entry": plan.payment_entry, "allocations": plan.allocations})
    return plan


def _success(plan: CreditPlan, mode: str) -> CreditExecutionResult:
    return CreditExecutionResult(
        ok=True, mode=mode, consumed=plan.consumed, applied=plan.applied,
        allocations=list(plan.allocations), credit_usage=list(plan.credit_usage),
    )


def execute_with_compensation(
    plan: CreditPlan,
    *,
    consume_credit_row: Callable[..., bool],
    update_target_row: Callable[[int, float, dict], bool],
    insert_row: Callable[..., Any],
    save_payment: Callable[[dict, list[dict]], Any],
    restore_row: Callable[[int, dict], Any],
    delete_row: Callable[[int], Any],
    delete_payment: Callable[[int], Any],
) -> CreditExecutionResult:
    """Apply the plan step by step; on any failure undo the applied steps in reverse order.

    Not a real transaction: it narrows the window of inconsistency but cannot close it (a process crash
    between steps, or an undo that itself fails, leaves partial state -- reported in ``rollback_errors``).
    """
    undo_stack: list[tuple[str, Callable[[], Any]]] = []
    try:
        for op in plan.ops:
            kind = op["op"]
            if kind == "credit_consume":
                done = consume_credit_row(
                    op["id"], op["expected_amount"], op["use_amount"],
                    datetime.date.fromisoformat(op["payment_date"]), op["description"],
                )
                if not done:
                    raise CreditConflict(f"credit row id={op['id']} changed since it was loaded or is no longer open")
                undo_stack.append((f"restore credit row id={op['id']}", lambda o=op: restore_row(o["id"], o["before"])))
            elif kind == "insert_row":
                row = dict(op["row"])
                if row.get("payment_date"):
                    row["payment_date"] = datetime.date.fromisoformat(row["payment_date"])
                inserted = insert_row(**row)
                new_id = int(inserted["id"])
                undo_stack.append((f"delete inserted row id={new_id}", lambda i=new_id: delete_row(i)))
            elif kind == "target_update":
                done = update_target_row(op["id"], op["expected_amount"], op["fields"])
                if not done:
                    raise CreditConflict(f"invoice row id={op['id']} changed since it was loaded or is no longer open")
                undo_stack.append((f"restore invoice row id={op['id']}", lambda o=op: restore_row(o["id"], o["before"])))
            elif kind == "save_payment":
                saved = save_payment(op["entry"], op["allocations"])
                payment_id = int(saved["id"])
                undo_stack.append((f"delete payment record id={payment_id}", lambda p=payment_id: delete_payment(p)))
            else:  # pragma: no cover - defensive
                raise ValueError(f"unknown operation {kind!r}")
    except Exception as exc:
        logger.exception("Apply Credit failed; rolling back %d applied step(s)", len(undo_stack))
        result = CreditExecutionResult(mode="compensation")
        result.errors.append(
            f"Apply Credit failed and was rolled back ({type(exc).__name__}: {exc}). No credit was used."
        )
        for description, undo in reversed(undo_stack):
            try:
                undo()
            except Exception as undo_exc:
                logger.exception("Rollback step failed: %s", description)
                result.rollback_errors.append(f"{description}: {type(undo_exc).__name__}: {undo_exc}")
        result.rolled_back = not result.rollback_errors
        if result.rollback_errors:
            result.errors.append(
                "ROLLBACK INCOMPLETE - the database may be inconsistent. Please review: "
                + "; ".join(result.rollback_errors)
            )
        return result
    return _success(plan, "compensation")


def _is_function_missing(exc: Exception) -> bool:
    return getattr(exc, "code", None) == "PGRST202" or "Could not find the function" in str(exc)


def apply_credit_atomically(
    plan: CreditPlan,
    *,
    rpc: Optional[Callable[[dict], Any]],
    fallback: Callable[[CreditPlan], CreditExecutionResult],
) -> CreditExecutionResult:
    """Run the plan in one DB transaction (``rpc``); use ``fallback`` only if the function is not installed."""
    if not plan.ops:
        return CreditExecutionResult(ok=True, mode="transaction")
    if rpc is not None:
        try:
            rpc(plan.rpc_payload())
            return _success(plan, "transaction")
        except Exception as exc:
            if _is_function_missing(exc):
                logger.warning("apply_credit_atomic() is not installed in the database; using compensating rollback")
            else:
                logger.exception("apply_credit_atomic() failed")
                definite = getattr(exc, "code", None) is not None  # the database answered: the transaction rolled back
                result = CreditExecutionResult(mode="transaction", rolled_back=definite, outcome_unknown=not definite)
                if definite:
                    result.errors.append(
                        f"Apply Credit was rejected by the database and nothing was changed ({exc}). No credit was used."
                    )
                else:
                    result.errors.append(
                        f"Apply Credit lost contact with the database ({type(exc).__name__}: {exc}). The outcome is "
                        "UNKNOWN - check the credit rows and invoice rows before retrying."
                    )
                return result
    return fallback(plan)
