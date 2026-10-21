"""
Order approval service.

Self-contained data + storage + token layer for the customer purchase-order
approval workflow (MVP). Kept separate from ``services.supabase_service`` so the
standalone polling script can run outside the Streamlit runtime and so the new
feature cannot break the existing CRM.

Security notes:
- Approval tokens are random (``secrets.token_urlsafe``).
- Only the SHA-256 *hash* of the token is stored in the database.
- Tokens expire after 7 days and are single-use.
- Purchase-order PDFs live in a private bucket and are exposed only via
  short-lived signed URLs.
"""
from __future__ import annotations

import datetime
import hashlib
import logging
import os
import re
import secrets
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

PURCHASE_ORDER_BUCKET = "purchase-order-pdfs"
TOKEN_EXPIRY_DAYS = 7

# Order-intake reliability settings (see migrations/harden_order_intake_reliability.sql).
MAX_INTAKE_ATTEMPTS = 5
INTAKE_LEASE_MINUTES = 20
REQUIRED_INTAKE_SCHEMA_VERSION = 20261015

_client = None

_UNRELIABLE_PROJECT_NAME_WORDS = {
    "document", "doc", "file", "scan", "copy", "invoice", "order", "po", "purchase", "attachment", "untitled",
}


def _project_name_looks_unreliable(name: Optional[str]) -> bool:
    """True when a filename-derived project name still looks like raw filename noise (a generic word
    and/or an unstripped date/reference number) rather than an actual project/store name.

    Production incident this guards against: a PO's PDF was named
    'document_iretailche_20260923_2026-0047-2.pdf' (not the usual 'YYYY-MM-NNN_ProjectName.pdf'
    convention), so _project_name_from_filename() (order_intake_poll.py) only stripped the extension and
    swapped underscores for spaces -- producing the literal garbage name
    'document iretailche 20260923 2026-0047-2', which then flowed untouched into the auto-created
    invoice, project, and order record on approval. Used by both the Order Approvals page (to warn +
    offer an editable field before approving) and the order-confirmation validator (as a hard blocker --
    see services/order_confirmation_service.py)."""
    text = (name or "").strip()
    if not text:
        return True
    first_word = text.split(" ", 1)[0].strip().lower()
    if first_word in _UNRELIABLE_PROJECT_NAME_WORDS:
        return True
    if re.search(r"\d{6,}", text):  # an unstripped date/reference number, e.g. 20260923, 2026-0047-2
        return True
    return False


# ── Supabase client (works in Streamlit *and* standalone scripts) ─────────────
def _get_client():
    """Return a cached Supabase client using env vars first, then st.secrets."""
    global _client
    if _client is not None:
        return _client

    url = os.environ.get("SUPABASE_URL", "").strip()
    key = (
        os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
        or os.environ.get("SUPABASE_KEY", "").strip()
    )

    if not url or not key:
        try:
            import streamlit as st  # noqa: WPS433 (optional dependency at runtime)

            cfg = st.secrets.get("supabase", {})
            url = url or str(cfg.get("url", "")).strip()
            key = key or str(cfg.get("service_role_key", cfg.get("anon_key", ""))).strip()
        except Exception:
            pass

    if not url or not key:
        raise RuntimeError(
            "Supabase credentials not configured. Set SUPABASE_URL and "
            "SUPABASE_SERVICE_ROLE_KEY env vars, or a [supabase] section in "
            "Streamlit secrets."
        )

    from supabase import create_client

    _client = create_client(url, key)
    return _client


def _ensure_bucket(client, bucket_name: str) -> None:
    try:
        buckets = client.storage.list_buckets()
        for bucket in buckets:
            name = bucket.get("name") if isinstance(bucket, dict) else getattr(bucket, "name", None)
            if name == bucket_name:
                return
    except Exception:
        pass
    try:
        client.storage.create_bucket(bucket_name, options={"public": False})
    except TypeError:
        try:
            client.storage.create_bucket(bucket_name)
        except Exception as exc:
            if "already exists" not in str(exc).lower():
                raise
    except Exception as exc:
        if "already exists" not in str(exc).lower():
            raise


# ── Token helpers ─────────────────────────────────────────────────────────────
def generate_token() -> str:
    """Return a new random URL-safe token (the raw secret, never stored)."""
    return secrets.token_urlsafe(32)


def hash_token(token: str) -> str:
    """Return the SHA-256 hex digest used as the stored token identifier."""
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def is_valid_token_format(token: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9_-]{20,128}", token or ""))


# ── PDF storage ───────────────────────────────────────────────────────────────
def _pdf_storage_path(filename: str) -> str:
    base = Path(filename).name or "purchase_order.pdf"
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_") or "purchase_order.pdf"
    stamp = datetime.datetime.utcnow().strftime("%Y/%m/%Y%m%d-%H%M%S")
    return f"{stamp}-{safe_name}"


def upload_purchase_order_pdf(
    file_bytes: bytes,
    filename: str,
    storage_path: Optional[str] = None,
) -> Dict[str, str]:
    """Upload a PO PDF to the private bucket. Returns bucket + path metadata."""
    client = _get_client()
    _ensure_bucket(client, PURCHASE_ORDER_BUCKET)
    target_path = storage_path or _pdf_storage_path(filename)

    try:
        client.storage.from_(PURCHASE_ORDER_BUCKET).remove([target_path])
    except Exception:
        pass

    suffix = Path(filename).suffix.lower() or ".pdf"
    content_type = "application/pdf" if suffix == ".pdf" else "application/octet-stream"

    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp_file:
        tmp_file.write(file_bytes)
        tmp_path = Path(tmp_file.name)

    try:
        client.storage.from_(PURCHASE_ORDER_BUCKET).upload(
            path=target_path,
            file=str(tmp_path),
            file_options={"content-type": content_type, "upsert": "true"},
        )
    finally:
        try:
            tmp_path.unlink()
        except Exception:
            pass

    return {
        "pdf_storage_bucket": PURCHASE_ORDER_BUCKET,
        "pdf_storage_path": target_path,
    }


def create_purchase_order_signed_url(
    bucket_name: str,
    storage_path: str,
    expires_in: int = 3600,
) -> Optional[str]:
    """Return a short-lived signed URL for a private PO PDF (or None)."""
    if not bucket_name or not storage_path:
        return None
    client = _get_client()
    try:
        resp = client.storage.from_(bucket_name).create_signed_url(storage_path, expires_in)
    except Exception as exc:
        logger.warning("Could not create signed URL for purchase order PDF: %s", exc)
        return None
    if isinstance(resp, dict):
        return resp.get("signedURL") or resp.get("signed_url") or resp.get("signedUrl")
    return None


# ── Incoming emails ───────────────────────────────────────────────────────────
def incoming_email_exists(message_id: str) -> bool:
    if not message_id:
        return False
    client = _get_client()
    resp = (
        client.table("incoming_order_emails")
        .select("id, processing_status")
        .eq("message_id", message_id)
        .limit(1)
        .execute()
    )
    if not resp.data:
        return False
    # Error records are retried, not skipped
    return resp.data[0].get("processing_status") != "error"


def get_incoming_email_status(message_id: str) -> Optional[str]:
    """Return the processing_status of an existing record, or None if absent."""
    if not message_id:
        return None
    client = _get_client()
    resp = (
        client.table("incoming_order_emails")
        .select("processing_status")
        .eq("message_id", message_id)
        .limit(1)
        .execute()
    )
    if not resp.data:
        return None
    return resp.data[0].get("processing_status")


def delete_incoming_email_by_message_id(message_id: str) -> None:
    """Delete a previously failed (error) incoming-email record so it can be retried."""
    client = _get_client()
    client.table("incoming_order_emails").delete().eq("message_id", message_id).execute()


def record_incoming_email(**fields) -> dict:
    """Insert a raw incoming-email row. ``message_id`` must be unique."""
    client = _get_client()
    resp = client.table("incoming_order_emails").insert(fields).execute()
    return resp.data[0] if resp.data else {}


def update_incoming_email(email_id: str, **fields) -> dict:
    client = _get_client()
    resp = (
        client.table("incoming_order_emails")
        .update(fields)
        .eq("id", email_id)
        .execute()
    )
    return resp.data[0] if resp.data else {}


def list_incoming_emails(limit: int = 200) -> List[dict]:
    client = _get_client()
    resp = (
        client.table("incoming_order_emails")
        .select("*")
        .order("created_at", desc=True)
        .limit(limit)
        .execute()
    )
    return resp.data or []


def get_incoming_email(email_id: str) -> Optional[dict]:
    """Fetch a single incoming-email row by id (used to recover its extracted PDF text)."""
    if not email_id:
        return None
    client = _get_client()
    resp = (
        client.table("incoming_order_emails")
        .select("*")
        .eq("id", email_id)
        .limit(1)
        .execute()
    )
    return resp.data[0] if resp.data else None


def download_purchase_order_pdf_bytes(bucket_name: str, storage_path: str) -> Optional[bytes]:
    """Best-effort download of a private PO PDF's raw bytes (or None)."""
    if not bucket_name or not storage_path:
        return None
    client = _get_client()
    try:
        return client.storage.from_(bucket_name).download(storage_path)
    except Exception as exc:
        logger.warning("Could not download purchase order PDF: %s", exc)
        return None


# ── Purchase orders ───────────────────────────────────────────────────────────
def create_purchase_order(**fields) -> dict:
    client = _get_client()
    resp = client.table("purchase_orders").insert(fields).execute()
    return resp.data[0] if resp.data else {}


def update_purchase_order(order_id: str, **fields) -> dict:
    client = _get_client()
    fields["updated_at"] = datetime.datetime.utcnow().isoformat()
    resp = (
        client.table("purchase_orders")
        .update(fields)
        .eq("id", order_id)
        .execute()
    )
    return resp.data[0] if resp.data else {}


def get_purchase_order(order_id: str) -> Optional[dict]:
    client = _get_client()
    resp = (
        client.table("purchase_orders")
        .select("*")
        .eq("id", order_id)
        .limit(1)
        .execute()
    )
    return resp.data[0] if resp.data else None


def list_purchase_orders(limit: int = 200) -> List[dict]:
    client = _get_client()
    resp = (
        client.table("purchase_orders")
        .select("*")
        .order("created_at", desc=True)
        .limit(limit)
        .execute()
    )
    return resp.data or []


def delete_purchase_order(order_id: str) -> None:
    """Delete a purchase order (and its PDF). Related approvals/notifications cascade."""
    client = _get_client()
    order = get_purchase_order(order_id)
    if order and order.get("pdf_storage_bucket") and order.get("pdf_storage_path"):
        try:
            client.storage.from_(order["pdf_storage_bucket"]).remove([order["pdf_storage_path"]])
        except Exception:
            pass
    client.table("purchase_orders").delete().eq("id", order_id).execute()


def delete_purchase_orders_by_status(status: str) -> int:
    """Delete all purchase orders with the given status. Returns count deleted."""
    client = _get_client()
    resp = client.table("purchase_orders").select("id").eq("status", status).execute()
    order_ids = [row["id"] for row in (resp.data or [])]
    for order_id in order_ids:
        delete_purchase_order(order_id)
    return len(order_ids)


# ── Approval tokens ───────────────────────────────────────────────────────────
def create_approval(purchase_order_id: str, expires_days: int = TOKEN_EXPIRY_DAYS) -> str:
    """
    Create an approval token for a purchase order.

    Returns the *raw* token (caller embeds it in the email link). Only the hash
    is persisted.
    """
    client = _get_client()
    token = generate_token()
    expires_at = datetime.datetime.utcnow() + datetime.timedelta(days=expires_days)

    client.table("order_approvals").insert(
        {
            "purchase_order_id": purchase_order_id,
            "token_hash": hash_token(token),
            "status": "pending",
            "expires_at": expires_at.isoformat(),
        }
    ).execute()
    return token


def get_approval_by_token(token: str) -> Optional[dict]:
    """Look up an approval row by raw token (matched against the stored hash)."""
    if not is_valid_token_format(token):
        return None
    client = _get_client()
    resp = (
        client.table("order_approvals")
        .select("*")
        .eq("token_hash", hash_token(token))
        .limit(1)
        .execute()
    )
    return resp.data[0] if resp.data else None


_DECISION_TO_STATUS = {
    "approve": "approved",
    "reject": "rejected",
    "request_correction": "needs_correction",
}


def apply_decision(
    token: str,
    decision: str,
    comment: str = "",
    decided_by: str = "",
    approved_by_ip: str = "",
    user_agent: str = "",
) -> dict:
    """
    Validate the token and apply an approval decision (single-use, 7-day expiry).

    ``decision`` is one of: approve | reject | request_correction.
    ``ip_address`` / ``user_agent`` are stored as an audit trail.
    Returns ``{"success": bool, "message": str, "order"?: dict}``.
    """
    decision = (decision or "").strip().lower()
    if decision not in _DECISION_TO_STATUS:
        return {"success": False, "message": "Unknown decision."}

    approval = get_approval_by_token(token)
    if not approval:
        return {"success": False, "message": "Invalid or unknown approval link."}

    if approval["status"] != "pending":
        return {
            "success": False,
            "message": "This approval link has already been used.",
        }

    # Expiry check.
    client = _get_client()
    try:
        expires_at = datetime.datetime.fromisoformat(
            str(approval["expires_at"]).replace("Z", "+00:00")
        )
        if datetime.datetime.now(datetime.timezone.utc) > expires_at:
            client.table("order_approvals").update({"status": "expired"}).eq(
                "id", approval["id"]
            ).execute()
            return {"success": False, "message": "This approval link has expired."}
    except Exception:
        pass

    new_status = _DECISION_TO_STATUS[decision]
    now_iso = datetime.datetime.utcnow().isoformat()

    # Mark token used (single-use) and record the audit trail.
    client.table("order_approvals").update(
        {
            "status": new_status,
            "decision": decision,
            "decision_comment": comment or None,
            "decided_by": decided_by or None,
            "approved_by_ip": approved_by_ip or None,
            "user_agent": (user_agent[:1000] if user_agent else None),
            "action_timestamp": now_iso,
            "used_at": now_iso,
        }
    ).eq("id", approval["id"]).execute()

    # Update the related purchase order.
    order = update_purchase_order(
        approval["purchase_order_id"],
        status=new_status,
        decided_at=now_iso,
        decision_comment=comment or None,
    )

    label = {
        "approved": "approved",
        "rejected": "rejected",
        "needs_correction": "marked as needing correction",
    }[new_status]

    # Raise a CRM notification so the change surfaces on the dashboard.
    _notify_status_change(order or {}, new_status, comment=comment, source="approval link")

    # On approval, automatically create a draft CRM invoice (idempotent).
    if new_status == "approved":
        auto_create_invoice_for_approved_order(approval["purchase_order_id"])

    return {
        "success": True,
        "message": f"Purchase order {label}.",
        "order": order,
    }


# ── Auto invoice creation from approved purchase orders ──────────────────


def get_invoice_for_purchase_order(purchase_order_id: str) -> Optional[dict]:
    """Return the invoice already created from this purchase order, or None."""
    if not purchase_order_id:
        return None
    client = _get_client()
    resp = (
        client.table("invoices")
        .select("*")
        .eq("source_purchase_order_id", purchase_order_id)
        .limit(1)
        .execute()
    )
    return resp.data[0] if resp.data else None


def _next_invoice_number(client) -> int:
    resp = client.table("invoices").select("invoice_number").execute()
    max_no = 0
    for row in resp.data or []:
        try:
            n = int(row["invoice_number"])
            if n > max_no:
                max_no = n
        except Exception:
            pass
    return max_no + 1


def create_invoice_from_purchase_order(purchase_order_id: str) -> dict:
    """
    Idempotently create a draft ("ready_to_send") invoice from an approved PO.

    The invoice is NOT emailed to the customer; it is stored in the existing
    ``invoices`` table with ``send_status='ready_to_send'`` for manual review.

    Returns ``{"created": bool, "invoice"?: dict, "message": str}``.
    """
    order = get_purchase_order(purchase_order_id)
    if not order:
        return {"created": False, "message": "Purchase order not found."}

    existing = get_invoice_for_purchase_order(purchase_order_id)
    if existing:
        return {
            "created": False,
            "invoice": existing,
            "message": "Invoice already exists for this purchase order.",
        }

    client = _get_client()
    invoice_number = _next_invoice_number(client)

    created_at = str(order.get("created_at") or "")
    year = datetime.datetime.utcnow().year
    if len(created_at) >= 4 and created_at[:4].isdigit():
        year = int(created_at[:4])

    customer = order.get("customer_name") or order.get("customer_email") or "Unknown customer"
    currency = order.get("currency") or "EUR"
    amount = order.get("amount")
    summary = (order.get("summary") or "").strip()

    pdf_ref = ""
    if order.get("pdf_storage_bucket") and order.get("pdf_storage_path"):
        pdf_ref = f"{order['pdf_storage_bucket']}/{order['pdf_storage_path']}"

    description_lines = [
        "Auto-created from approved purchase order.",
        f"Customer: {customer}",
        f"PO reference: {order.get('order_reference') or 'n/a'}",
        f"Order date: {created_at[:10] or 'n/a'}",
        f"Amount: {amount if amount is not None else 'n/a'} {currency}",
    ]
    if summary:
        description_lines.append(f"Summary: {summary[:500]}")
    if pdf_ref:
        description_lines.append(f"PDF: {pdf_ref}")
    description = "\n".join(description_lines)

    now_iso = datetime.datetime.utcnow().isoformat()
    row = {
        "invoice_number": str(invoice_number),
        "project_name": order.get("project_name") or customer,
        "maintenance_year": "Y1",
        "payment_amount": float(amount) if amount is not None else 0.0,
        "cameras_number": None,
        "payment_date": None,
        "paid": "No",
        "year": year,
        "invoice_type": "Purchase Order",
        "description": description,
        "source_type": "purchase_order",
        "source_purchase_order_id": purchase_order_id,
        "auto_created": True,
        "auto_created_at": now_iso,
        "send_status": "ready_to_send",
    }
    resp = client.table("invoices").insert(row).execute()
    invoice = resp.data[0] if resp.data else {}
    logger.info(
        "Auto-created invoice #%s from purchase order %s",
        invoice_number,
        purchase_order_id,
    )
    return {"created": True, "invoice": invoice, "message": "Invoice created (ready to send)."}


def auto_create_invoice_for_approved_order(purchase_order_id: str) -> dict:
    """
    Best-effort wrapper used after a PO is approved.

    Never raises: on success it posts a ``success`` CRM notification, on failure
    it leaves the PO approved and posts an ``error`` notification instead.
    """
    try:
        result = create_invoice_from_purchase_order(purchase_order_id)
    except Exception as exc:  # noqa: BLE001 - approval must not be rolled back
        logger.error(
            "Auto invoice creation failed for purchase order %s: %s",
            purchase_order_id,
            exc,
        )
        try:
            create_notification(
                title="Auto invoice creation failed",
                message=(
                    "The purchase order was approved, but an invoice could not be "
                    f"created automatically: {exc}"
                ),
                severity="error",
                purchase_order_id=purchase_order_id,
            )
        except Exception:
            pass
        return {"created": False, "message": str(exc)}

    if result.get("created"):
        invoice = result.get("invoice") or {}
        try:
            create_notification(
                title="Invoice created automatically from approved purchase order",
                message=(
                    f"Invoice #{invoice.get('invoice_number')} was created and marked "
                    "ready to send (not emailed to the customer yet)."
                ),
                severity="success",
                purchase_order_id=purchase_order_id,
            )
        except Exception:
            pass
    return result


# ── CRM notifications ────────────────────────────────────────────────────
_STATUS_SEVERITY = {
    "approved": "success",
    "rejected": "error",
    "needs_correction": "warning",
    "superseded": "info",
}


def create_notification(
    title: str,
    message: str = "",
    severity: str = "info",
    category: str = "order_approval",
    purchase_order_id: Optional[str] = None,
) -> dict:
    """Insert a CRM notification row (best effort; never raises)."""
    try:
        client = _get_client()
        resp = client.table("crm_notifications").insert(
            {
                "title": title,
                "message": message or None,
                "severity": severity if severity in ("info", "success", "warning", "error") else "info",
                "category": category,
                "purchase_order_id": purchase_order_id,
            }
        ).execute()
        return resp.data[0] if resp.data else {}
    except Exception as exc:
        logger.warning("Could not create CRM notification: %s", exc)
        return {}


def _notify_status_change(order: dict, new_status: str, comment: str = "", source: str = "") -> None:
    who = order.get("project_name") or order.get("customer_name") or order.get("customer_email") or "a purchase order"
    label = {
        "approved": "approved",
        "rejected": "rejected",
        "needs_correction": "flagged for correction",
        "superseded": "superseded by a revision",
    }.get(new_status, new_status)
    suffix = f" via {source}" if source else ""
    msg = f"Purchase order for {who} was {label}{suffix}."
    if comment:
        msg += f" Comment: {comment}"
    create_notification(
        title=f"Order {label}: {who}",
        message=msg,
        severity=_STATUS_SEVERITY.get(new_status, "info"),
        purchase_order_id=order.get("id"),
    )


def list_notifications(limit: int = 50, unread_only: bool = False) -> List[dict]:
    client = _get_client()
    query = client.table("crm_notifications").select("*")
    if unread_only:
        query = query.eq("is_read", False)
    resp = query.order("created_at", desc=True).limit(limit).execute()
    return resp.data or []


def count_unread_notifications() -> int:
    try:
        client = _get_client()
        resp = client.table("crm_notifications").select("id").eq("is_read", False).execute()
        return len(resp.data or [])
    except Exception:
        return 0


def mark_notification_read(notification_id: str) -> None:
    client = _get_client()
    client.table("crm_notifications").update({"is_read": True}).eq("id", notification_id).execute()


def mark_all_notifications_read() -> None:
    client = _get_client()
    client.table("crm_notifications").update({"is_read": True}).eq("is_read", False).execute()


# ── Revisions (customer re-sends a revised PDF) ───────────────────────────────
def find_current_open_order_by_reference(
    order_reference: str, exclude_id: Optional[str] = None
) -> Optional[dict]:
    """
    Return the current (non-superseded) purchase order for a given customer
    reference that has NOT yet been decided, or None.

    A reference is required — without one we cannot reliably correlate emails,
    so each email is treated as a distinct order. ``exclude_id`` skips one order
    (used when resuming, so a new revision is not matched against itself).
    """
    ref = (order_reference or "").strip()
    if not ref:
        return None
    client = _get_client()
    query = (
        client.table("purchase_orders")
        .select("*")
        .eq("order_reference", ref)
        .eq("is_current", True)
        .eq("status", "pending_approval")
    )
    if exclude_id:
        query = query.neq("id", exclude_id)
    resp = query.order("revision", desc=True).limit(1).execute()
    return resp.data[0] if resp.data else None


def supersede_purchase_order(old_order_id: str, new_order_id: str) -> None:
    """Mark an older purchase order as superseded by a newer revision and
    expire any pending approval tokens that pointed at it."""
    client = _get_client()
    now_iso = datetime.datetime.utcnow().isoformat()
    client.table("purchase_orders").update(
        {
            "status": "superseded",
            "is_current": False,
            "superseded_by": new_order_id,
            "updated_at": now_iso,
        }
    ).eq("id", old_order_id).execute()
    # Invalidate outstanding tokens for the superseded order.
    client.table("order_approvals").update({"status": "expired"}).eq(
        "purchase_order_id", old_order_id
    ).eq("status", "pending").execute()


# ── Order intake: idempotent claims and resumable steps ──────────────────────
# Requires migrations/harden_order_intake_reliability.sql. Every step below is
# safe to repeat: rows are keyed by message id, a lease (claimed_by/claimed_at)
# gives one worker exclusive processing, and progress is derived from the data
# (PO exists? approval email sent?) rather than from in-memory state.
def _utcnow() -> datetime.datetime:
    return datetime.datetime.now(datetime.timezone.utc)


def _parse_dt(value: Any) -> Optional[datetime.datetime]:
    if not value:
        return None
    try:
        parsed = datetime.datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=datetime.timezone.utc)
    return parsed


def _is_duplicate_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return "23505" in text or "duplicate key" in text


def schema_ready() -> tuple[bool, str]:
    """True when the intake-hardening migration is applied (checked via an RPC marker)."""
    try:
        resp = _get_client().rpc("order_intake_schema_version").execute()
    except Exception as exc:
        return False, f"order_intake_schema_version() not available ({type(exc).__name__})"
    data = resp.data
    if isinstance(data, list):
        data = data[0] if data else None
    if isinstance(data, dict):
        data = next(iter(data.values()), None)
    try:
        version = int(data)
    except (TypeError, ValueError):
        return False, "order_intake_schema_version() returned no version"
    if version < REQUIRED_INTAKE_SCHEMA_VERSION:
        return False, f"schema version {version} is older than {REQUIRED_INTAKE_SCHEMA_VERSION}"
    return True, f"schema version {version}"


def deterministic_pdf_storage_path(
    filename: str, sha256: str, received_at: Optional[datetime.datetime] = None
) -> str:
    """Storage path derived from the content hash, so retries overwrite instead of orphaning files."""
    base = Path(filename).name or "purchase_order.pdf"
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_") or "purchase_order.pdf"
    stamp = (received_at or _utcnow()).strftime("%Y/%m")
    return f"{stamp}/{sha256[:16]}-{safe_name}"


def find_incoming_email(message_id: str, graph_message_id: str = "") -> Optional[dict]:
    """Return the intake row for a message, matched by RFC message id or immutable Graph id."""
    client = _get_client()
    if message_id:
        resp = (
            client.table("incoming_order_emails")
            .select("*")
            .eq("message_id", message_id)
            .limit(1)
            .execute()
        )
        if resp.data:
            return resp.data[0]
    if graph_message_id:
        resp = (
            client.table("incoming_order_emails")
            .select("*")
            .eq("graph_message_id", graph_message_id)
            .limit(1)
            .execute()
        )
        if resp.data:
            return resp.data[0]
    return None


def find_incoming_by_pdf_sha256(sha256: str, exclude_message_id: str = "") -> Optional[dict]:
    """Earlier intake row carrying an identical PDF (used only to flag possible duplicates)."""
    if not sha256:
        return None
    client = _get_client()
    query = client.table("incoming_order_emails").select("id,message_id,subject,created_at").eq(
        "pdf_sha256", sha256
    )
    if exclude_message_id:
        query = query.neq("message_id", exclude_message_id)
    resp = query.order("created_at").limit(1).execute()
    return resp.data[0] if resp.data else None


def insert_incoming_email_claimed(fields: dict, run_id: str) -> Optional[dict]:
    """Insert a new intake row already leased to ``run_id``; None if the message already exists."""
    row = {
        **fields,
        "processing_status": "received",
        "claimed_by": run_id,
        "claimed_at": _utcnow().isoformat(),
    }
    try:
        resp = _get_client().table("incoming_order_emails").insert(row).execute()
    except Exception as exc:
        if _is_duplicate_error(exc):
            return None
        raise
    return resp.data[0] if resp.data else None


def list_unfinished_incoming(
    created_after: datetime.datetime,
    max_attempts: int = MAX_INTAKE_ATTEMPTS,
    limit: int = 100,
) -> List[dict]:
    """Intake rows that still need work, newer than ``created_after`` (go-live floor)."""
    resp = (
        _get_client()
        .table("incoming_order_emails")
        .select("*")
        .in_("processing_status", ["received", "error"])
        .lt("attempt_count", max_attempts)
        .gte("created_at", created_after.isoformat())
        .order("created_at")
        .limit(limit)
        .execute()
    )
    return resp.data or []


def list_abandoned_incoming(
    created_after: datetime.datetime, max_attempts: int = MAX_INTAKE_ATTEMPTS
) -> List[dict]:
    resp = (
        _get_client()
        .table("incoming_order_emails")
        .select("id,message_id,subject,error_message,attempt_count")
        .eq("processing_status", "error")
        .gte("attempt_count", max_attempts)
        .gte("created_at", created_after.isoformat())
        .execute()
    )
    return resp.data or []


def claim_incoming_email(
    email_id: str, run_id: str, lease_minutes: int = INTAKE_LEASE_MINUTES
) -> Optional[dict]:
    """Atomically lease an unfinished row. Returns the row, or None if someone else holds it."""
    client = _get_client()
    now = _utcnow()
    claim = {"claimed_by": run_id, "claimed_at": now.isoformat()}

    def _update(extra_filter):
        query = (
            client.table("incoming_order_emails")
            .update(claim)
            .eq("id", email_id)
            .in_("processing_status", ["received", "error"])
        )
        return extra_filter(query).execute()

    resp = _update(lambda q: q.is_("claimed_at", "null"))
    if not resp.data:
        stale = (now - datetime.timedelta(minutes=lease_minutes)).isoformat()
        resp = _update(lambda q: q.lt("claimed_at", stale))
    return resp.data[0] if resp.data else None


def finish_incoming_email(email_id: str, run_id: str) -> None:
    (
        _get_client()
        .table("incoming_order_emails")
        .update(
            {
                "processing_status": "processed",
                "error_message": None,
                "claimed_by": None,
                "claimed_at": None,
            }
        )
        .eq("id", email_id)
        .eq("claimed_by", run_id)
        .execute()
    )


def fail_incoming_email(email_id: str, run_id: str, error: str, attempt_count: int) -> None:
    """Record a failed attempt and release the lease. Never deletes the row."""
    (
        _get_client()
        .table("incoming_order_emails")
        .update(
            {
                "processing_status": "error",
                "error_message": (error or "")[:500],
                "attempt_count": attempt_count,
                "claimed_by": None,
                "claimed_at": None,
            }
        )
        .eq("id", email_id)
        .eq("claimed_by", run_id)
        .execute()
    )
    if attempt_count >= MAX_INTAKE_ATTEMPTS:
        create_notification(
            title="Order intake gave up on an email",
            message=(
                f"An incoming purchase-order email failed {attempt_count} times and will not be "
                "retried automatically. See the Order Approvals page and the workflow log."
            ),
            severity="error",
        )


def get_purchase_order_for_incoming(email_id: str) -> Optional[dict]:
    if not email_id:
        return None
    resp = (
        _get_client()
        .table("purchase_orders")
        .select("*")
        .eq("incoming_email_id", email_id)
        .limit(1)
        .execute()
    )
    return resp.data[0] if resp.data else None


def get_or_create_purchase_order_for_incoming(email_id: str, fields: dict) -> tuple[dict, bool]:
    """Return (order, created). At most one order per intake row, enforced by a unique index."""
    existing = get_purchase_order_for_incoming(email_id)
    if existing:
        return existing, False
    try:
        return create_purchase_order(incoming_email_id=email_id, **fields), True
    except Exception as exc:
        if _is_duplicate_error(exc):
            existing = get_purchase_order_for_incoming(email_id)
            if existing:
                return existing, False
        raise


def prepare_approval_email(
    purchase_order_id: str, lease_minutes: int = INTAKE_LEASE_MINUTES
) -> dict:
    """
    Decide whether the approval email must be (re)sent and hand out a token.

    Returns ``{"state": ...}`` where state is:
    - ``send``: caller owns the send; ``token`` (raw), ``approval`` and ``attempt`` are included.
    - ``already_sent`` / ``not_needed`` / ``in_progress``: nothing to do.

    One approval row exists per purchase order. When a resend is needed the same row's
    token is rotated in place (the raw token is never stored), so a resend never creates a
    second active token and the previous link stops working.
    """
    client = _get_client()
    order = get_purchase_order(purchase_order_id)
    if not order or order.get("status") != "pending_approval" or order.get("is_current") is False:
        return {"state": "not_needed"}

    resp = (
        client.table("order_approvals")
        .select("*")
        .eq("purchase_order_id", purchase_order_id)
        .order("created_at", desc=True)
        .limit(1)
        .execute()
    )
    approval = resp.data[0] if resp.data else None

    now = _utcnow()
    token = generate_token()
    token_fields = {
        "token_hash": hash_token(token),
        "expires_at": (now + datetime.timedelta(days=TOKEN_EXPIRY_DAYS)).isoformat(),
        "email_status": "sending",
        "email_claimed_at": now.isoformat(),
    }

    if approval is None:
        try:
            ins = (
                client.table("order_approvals")
                .insert(
                    {
                        "purchase_order_id": purchase_order_id,
                        "status": "pending",
                        "email_attempts": 1,
                        **token_fields,
                    }
                )
                .execute()
            )
        except Exception as exc:
            if _is_duplicate_error(exc):
                return {"state": "in_progress"}
            raise
        return {"state": "send", "token": token, "approval": ins.data[0], "attempt": 1}

    if approval.get("status") != "pending":
        return {"state": "not_needed"}
    email_status = approval.get("email_status") or "pending"
    if email_status == "sent":
        return {"state": "already_sent", "approval": approval}
    if email_status == "sending":
        claimed = _parse_dt(approval.get("email_claimed_at"))
        if claimed and now - claimed < datetime.timedelta(minutes=lease_minutes):
            return {"state": "in_progress"}

    attempts = int(approval.get("email_attempts") or 0)
    upd = (
        client.table("order_approvals")
        .update({**token_fields, "email_attempts": attempts + 1})
        .eq("id", approval["id"])
        .eq("email_status", email_status)
        .eq("email_attempts", attempts)
        .execute()
    )
    if not upd.data:
        return {"state": "in_progress"}
    return {"state": "send", "token": token, "approval": upd.data[0], "attempt": attempts + 1}


def mark_approval_email_sent(approval_id: str) -> None:
    (
        _get_client()
        .table("order_approvals")
        .update(
            {
                "email_status": "sent",
                "email_sent_at": _utcnow().isoformat(),
                "email_last_error": None,
            }
        )
        .eq("id", approval_id)
        .execute()
    )


def mark_approval_email_failed(approval_id: str, error: str) -> None:
    (
        _get_client()
        .table("order_approvals")
        .update({"email_status": "failed", "email_last_error": (error or "")[:500]})
        .eq("id", approval_id)
        .execute()
    )


def get_order_delivery_state(order: dict) -> Optional[dict]:
    """Intake/approval-email state for the Order Approvals page; None if unavailable."""
    try:
        client = _get_client()
        state: Dict[str, Any] = {}
        if order.get("incoming_email_id"):
            row = get_incoming_email(order["incoming_email_id"]) or {}
            state.update(
                intake_status=row.get("processing_status"),
                attempt_count=int(row.get("attempt_count") or 0),
                error_message=row.get("error_message"),
            )
        resp = (
            client.table("order_approvals")
            .select("email_status,email_attempts,email_last_error")
            .eq("purchase_order_id", order["id"])
            .order("created_at", desc=True)
            .limit(1)
            .execute()
        )
        if resp.data:
            state.update(
                email_status=resp.data[0].get("email_status"),
                email_attempts=resp.data[0].get("email_attempts"),
                email_last_error=resp.data[0].get("email_last_error"),
            )
        return state
    except Exception as exc:
        logger.warning("Could not load delivery state: %s", type(exc).__name__)
        return None


def retry_order_intake(purchase_order_id: str) -> bool:
    """Give a failed intake row a fresh set of attempts; the poller resumes it on its next run."""
    try:
        order = get_purchase_order(purchase_order_id)
        if not order or not order.get("incoming_email_id"):
            return False
        resp = (
            _get_client()
            .table("incoming_order_emails")
            .update({"attempt_count": 0})
            .eq("id", order["incoming_email_id"])
            .eq("processing_status", "error")
            .execute()
        )
        return bool(resp.data)
    except Exception as exc:
        logger.warning("Could not reset intake attempts: %s", type(exc).__name__)
        return False
