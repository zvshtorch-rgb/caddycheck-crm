"""Minimal safety net for the Invoice Details save: an existing invoice's number must never change through it.

Why: the "Invoice #" cell of the Invoice Details grid was editable and the save rewrote every row in batches of 50,
so a changed number could be written by one batch while another failed (partial save), and a stale browser tab
could silently write an old number back. Renumbering needs its own recorded, deliberate operation.

Two independent checks (both run BEFORE anything is written):
  * `find_invoice_number_changes`      - compares the grid with the rows as originally loaded (UI check)
  * `assert_no_invoice_number_changes` - compares the rows about to be saved with the CURRENT database values,
                                         so it also catches stale or manipulated clients
No Streamlit dependency; needs no new tables.
"""
from __future__ import annotations

from typing import Any, Iterable, Optional

PAGE_SIZE = 1000


class InvoiceNumberChangeBlocked(RuntimeError):
    """An invoice number would change. Nothing was written."""


def normalize_invoice_number(value: Any) -> Optional[str]:
    """'8713', 8713, 8713.0, ' 8713 ' -> '8713'; None / '' / NaN -> None."""
    if value is None:
        return None
    if isinstance(value, float) and value != value:  # NaN
        return None
    text = str(value).strip()
    if not text or text.lower() in ("nan", "none"):
        return None
    try:
        return str(int(float(text)))
    except (TypeError, ValueError):
        return text


def find_invoice_number_changes(edited_rows: Iterable[dict], original_invoices: list) -> list[dict]:
    """Rows of the edited grid whose 'Invoice #' differs from the originally loaded invoice they belong to.

    `edited_rows` are dict-like rows carrying '_invoice_index' (position in `original_invoices`, blank/negative for
    new rows) and 'Invoice #'. New rows and indexes that do not exist are ignored."""
    changes = []
    for row in edited_rows:
        raw = row.get("_invoice_index")
        try:
            index = int(float(raw))
        except (TypeError, ValueError):
            continue
        if not 0 <= index < len(original_invoices):
            continue
        original = original_invoices[index]
        old = normalize_invoice_number(getattr(original, "invoice_number", None))
        new = normalize_invoice_number(row.get("Invoice #"))
        if old != new:
            changes.append({
                "index": index,
                "project": str(getattr(original, "project_name", "") or row.get("Project") or ""),
                "old": old,
                "new": new,
            })
    return changes


def describe_changes(changes: list[dict], limit: int = 5) -> str:
    shown = [f"{c['project'] or 'row ' + str(c['index'])}: #{c['old'] or '(blank)'} -> #{c['new'] or '(blank)'}"
             for c in changes[:limit]]
    more = f" (+{len(changes) - limit} more)" if len(changes) > limit else ""
    return "; ".join(shown) + more


def assert_no_invoice_number_changes(client, update_rows: Iterable[dict]) -> None:
    """Raise InvoiceNumberChangeBlocked if any row about to be updated carries a different invoice number than the
    database currently holds. Reads only; called before the first write so a violation can never cause a partial save."""
    wanted = {r["id"]: r.get("invoice_number") for r in update_rows if r.get("id") is not None}
    if not wanted:
        return
    current: dict = {}
    offset = 0
    while True:  # paginated: Supabase caps a single request at 1000 rows
        batch = (
            client.table("invoices").select("id,invoice_number").order("id")
            .range(offset, offset + PAGE_SIZE - 1).execute().data or []
        )
        for row in batch:
            current[row["id"]] = row.get("invoice_number")
        if len(batch) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    changed = [
        (rid, current[rid], wanted[rid]) for rid in wanted
        if rid in current and normalize_invoice_number(current[rid]) != normalize_invoice_number(wanted[rid])
    ]
    if changed:
        listing = "; ".join(f"row {rid}: #{old} -> #{new}" for rid, old, new in changed[:5])
        more = f" (+{len(changed) - 5} more)" if len(changed) > 5 else ""
        raise InvoiceNumberChangeBlocked(
            f"Changing an invoice number is not allowed here ({listing}{more}). "
            "Nothing was saved. Reload the page; renumbering needs a recorded, deliberate operation."
        )
