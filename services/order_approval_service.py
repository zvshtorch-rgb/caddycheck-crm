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

_client = None


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
        "maintenance_year": "Purchase Order",
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
def find_current_open_order_by_reference(order_reference: str) -> Optional[dict]:
    """
    Return the current (non-superseded) purchase order for a given customer
    reference that has NOT yet been decided, or None.

    A reference is required — without one we cannot reliably correlate emails,
    so each email is treated as a distinct order.
    """
    ref = (order_reference or "").strip()
    if not ref:
        return None
    client = _get_client()
    resp = (
        client.table("purchase_orders")
        .select("*")
        .eq("order_reference", ref)
        .eq("is_current", True)
        .eq("status", "pending_approval")
        .order("revision", desc=True)
        .limit(1)
        .execute()
    )
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
