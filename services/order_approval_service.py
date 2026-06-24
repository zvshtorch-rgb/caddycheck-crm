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
        .select("id")
        .eq("message_id", message_id)
        .limit(1)
        .execute()
    )
    return bool(resp.data)


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

    return {
        "success": True,
        "message": f"Purchase order {label}.",
        "order": order,
    }


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
