"""Global-off admin approval orchestration with one explicitly approved production canary.

Only the fixed order ID/name/reference may start while live_enabled is false. A read-only final
review is required for that canary; any durable attempt blocks further approvals/resends until
separately reviewed. Customer token approvals never call this module.

Database RPCs own claims, permanent registry reservations, insert-only ledger commits and atomic
status/audit updates. The final PDF is rendered only from committed rows. A stored PDF/envelope is
reused byte-for-byte on explicit resend. Graph acceptance has no message ID or exactly-once guarantee:
timeouts and uncertain result persistence block resends pending manual review.
"""
from __future__ import annotations

import base64
import hashlib
import json
import logging
import uuid
from decimal import Decimal, InvalidOperation
from typing import Any

import requests

from config.settings import get_order_confirmation_config
from services import order_confirmation_service as preview
from services.email_service import send_graph_email
from services.order_approval_service import _get_client, get_purchase_order

logger = logging.getLogger(__name__)

# Explicitly approved production canary; global enablement remains configuration-controlled.
CANARY_ORDER_ID = "e6c9a8b7-df7c-4318-bec9-187e6ed97cbf"


def live_approval_enabled() -> bool:
    value = get_order_confirmation_config().get("live_enabled", False)
    return value is True or (isinstance(value, str) and value.strip().lower() == "true")


def order_approval_enabled(order: dict) -> bool:
    if live_approval_enabled():
        return True
    return (
        order.get("id") == CANARY_ORDER_ID
        and order.get("project_name") == "AD DELHAIZE Roeselare (Mandel)"
        and order.get("order_reference") == "2026-01-005"
        and order.get("status") == "pending_approval"
        and order.get("is_current") is not False
        and order.get("confirmation_status") == "none"
        and not any(order.get(key) for key in (
            "confirmation_invoice_number", "confirmation_lock_token", "confirmation_stage",
            "confirmation_error", "confirmation_pdf_base64",
        ))
    )


def build_approval_review(purchase_order_id: str) -> dict:
    """Read-only final review. No reservation, ledger write, approval, or Graph request."""
    order = get_purchase_order(purchase_order_id)
    if not order or not order_approval_enabled(order):
        raise ValueError("Live approval is disabled for this order or its canary attempt has already started.")
    resolved, context = _validated_order(order)
    envelope = _envelope()
    number = preview._next_invoice_number_preview()
    subject, _ = preview.render_confirmation_email(
        **context, cameras=resolved["ordered_cameras"], amount=resolved["amount"],
        invoice_number=number, sender=envelope["sender"], proposed=False,
    )
    review = {
        **context, **envelope, "order_id": purchase_order_id,
        "cameras": resolved["ordered_cameras"], "amount": resolved["amount"],
        "invoice_number_preview": number, "subject": subject,
    }
    review["fingerprint"] = hashlib.sha256(
        json.dumps({"review": review, "order": order}, sort_keys=True, default=str).encode()
    ).hexdigest()
    return review


def _rpc(name: str, **params: Any) -> Any:
    result = _get_client().rpc(name, params).execute().data
    if result is None:
        raise RuntimeError(f"{name} returned no result; refresh and inspect the durable confirmation state.")
    return result


def _envelope() -> dict:
    cfg = get_order_confirmation_config()
    if cfg.get("sender") != preview.TEST_CONFIRMATION_SENDER:
        raise ValueError(f"Confirmation sender must be {preview.TEST_CONFIRMATION_SENDER}.")
    for key in ("to", "cc"):
        values = cfg.get(key, [])
        if not isinstance(values, list) or (key == "to" and not values):
            raise ValueError(f"Configure the order-confirmation {key} recipients.")
        if any(not isinstance(value, str) or "@" not in value or "\n" in value or "\r" in value for value in values):
            raise ValueError(f"Invalid configured {key} recipient.")
    return {"sender": cfg["sender"], "to": list(cfg["to"]), "cc": list(cfg.get("cc", []))}


def _validated_order(order: dict) -> tuple[dict, dict]:
    validation = preview.validate_order_for_year1_invoicing(order)
    errors = list(validation["errors"])
    if not str(order.get("order_reference") or "").strip():
        errors.append("A real PO/order reference is required; a synthetic reference cannot be invoiced.")
    if str(order.get("currency") or "EUR").upper() != "EUR":
        errors.append("Only EUR orders are supported by the Year-1 invoice workflow.")
    for key in ("amount", "payment_amount", "ordered_cameras", "num_cams"):
        if order.get(key) not in (None, ""):
            try:
                value = Decimal(str(order[key]))
                if not value.is_finite() or value < 0 or ("cam" in key and value != value.to_integral_value()):
                    raise InvalidOperation
            except InvalidOperation:
                errors.append(f"Invalid {key}: a finite non-negative {'whole count' if 'cam' in key else 'amount'} is required.")
    if errors:
        raise ValueError(" ".join(errors))
    cameras, amount = preview.compute_year1_amount(order)
    resolved = dict(order, ordered_cameras=cameras, amount=amount)
    context = {
        "project_name": preview.canonical_project_name(str(order["project_name"])),
        "customer_display": preview.resolve_customer_network(order),
        "country": preview.resolve_project_country(str(order["project_name"])),
        "po_reference": str(order["order_reference"]).strip(),
    }
    return resolved, context


def approve_order(purchase_order_id: str, actor: str, *, resend: bool = False,
                  review_fingerprint: str | None = None) -> dict:
    """Admin backend entry point. No legacy approval helper and no destructive ledger replacement."""
    if not actor.strip():
        raise ValueError("An authenticated approving user is required.")
    order = get_purchase_order(purchase_order_id)
    if not order:
        raise ValueError("Purchase order not found.")
    canary = not live_approval_enabled()
    if not order_approval_enabled(order) or (canary and resend):
        raise ValueError("Live admin order confirmation is disabled for this order; canary retries require review.")
    if canary and not review_fingerprint:
        raise ValueError("Review the final confirmation panel before approving the canary.")
    if review_fingerprint:
        current_review = build_approval_review(purchase_order_id)
        if current_review["fingerprint"] != review_fingerprint:
            raise ValueError("Order, recipients, or proposed invoice changed. Reopen the final review before approval.")
    if order.get("confirmation_status") == "sent":
        return {"status": "sent", "already_sent": True,
                "invoice_number": order.get("confirmation_invoice_number"),
                "message": "Confirmation already sent; no invoice or email was created."}
    if order.get("confirmation_lock_token"):
        return {"status": "busy", "message": "Confirmation is in progress or its outcome needs manual review. No retry was sent."}

    # Resend uses the original saved envelope, never a changed order or recipient configuration.
    resolved, context = (None, None) if resend else _validated_order(order)
    envelope = dict(order.get("confirmation_details") or {}) if resend else _envelope()
    if resend and (not envelope or not order.get("confirmation_pdf_base64")):
        raise ValueError("Stored confirmation PDF/envelope missing; resend is blocked.")
    token = str(uuid.uuid4())
    expected = {key: order.get(key) for key in (
        "id", "project_name", "amount", "currency", "order_reference", "status",
        "incoming_email_id", "customer_name", "customer_email", "confirmation_invoice_number",
        "confirmation_status", "confirmation_stage",
    )}
    claim = _rpc("claim_order_confirmation", p_order_id=purchase_order_id, p_token=token,
                 p_expected=expected, p_actor=actor, p_resend=resend)
    if claim["state"] != "claimed":
        return {"status": claim["state"], "message": "Order is already sent or another confirmation attempt owns it. Nothing sent."}
    order = claim["order"]
    stage = "reservation"
    sending = False
    number = order.get("confirmation_invoice_number")
    try:
        number = int(_rpc("reserve_order_confirmation_invoice", p_order_id=purchase_order_id,
                          p_token=token, p_actor=actor))
        stage = "invoice"
        payload = preview.build_year1_invoice_rows(resolved, number) if resolved else []
        saved = _rpc("save_order_confirmation_invoice", p_order_id=purchase_order_id,
                     p_token=token, p_rows=payload, p_actor=actor)
        if len(saved) != 1 or any(
            str(row.get("source_purchase_order_id")) != purchase_order_id
            or str(row.get("invoice_number")) != str(number) for row in saved
        ):
            raise ValueError("Saved invoice ownership/row count verification failed; nothing sent.")
        stage = "pdf"
        if order.get("confirmation_pdf_base64"):
            attachment = base64.b64decode(order["confirmation_pdf_base64"], validate=True)
            envelope = dict(order["confirmation_details"])
        else:
            attachment = preview.generate_year1_invoice_pdf_bytes(number, saved)
            row = saved[0]
            subject, body = preview.render_confirmation_email(
                project_name=str(row["project_name"]), customer_display=context["customer_display"],
                country=context["country"], po_reference=context["po_reference"],
                cameras=int(row.get("cameras_number") or 0), amount=float(row["payment_amount"]),
                invoice_number=number, sender=envelope["sender"], proposed=False,
            )
            envelope.update({"subject": subject, "body": body})
        preview.validate_invoice_pdf_bytes(attachment)
        stage = "prepare_send"
        _rpc("start_order_confirmation_send", p_order_id=purchase_order_id, p_token=token,
             p_saved_rows=saved, p_pdf_base64=base64.b64encode(attachment).decode("ascii"),
             p_details=envelope, p_actor=actor)
        stage = "email"
        sending = True
        response = send_graph_email(
            subject=envelope["subject"], body=envelope["body"], recipients=envelope["to"],
            cc=envelope["cc"], sender_mailbox=envelope["sender"], attachment_bytes=attachment,
            attachment_filename=f"invoice_{number}.pdf", attachment_content_type="application/pdf",
        )
        if response.get("http_status") != 202 or not response.get("accepted"):
            raise RuntimeError("Graph response was not confirmed as 202 Accepted; check Sent Items before any retry.")
    except Exception as exc:
        logger.exception("Admin order confirmation failed at %s for order %s", stage, purchase_order_id)
        # A definite Graph 4xx rejection (except request timeout) permits explicit retry; all other
        # post-send errors are ambiguous. A failure before Graph is called is safe to retry.
        rejected = (isinstance(exc, requests.HTTPError) and exc.response is not None
                    and 400 <= exc.response.status_code < 500 and exc.response.status_code != 408)
        safe = not sending or rejected
        error = f"{stage}: {type(exc).__name__}"
        if isinstance(exc, requests.HTTPError) and exc.response is not None:
            details = preview._graph_error_details(exc.response)
            error += f" (HTTP {exc.response.status_code}): {details['code']}: {details['message']}"
        elif not isinstance(exc, requests.RequestException):
            error += f": {exc}"
        outcome = "failed" if safe else "unknown"
        try:
            failed = _rpc("finish_order_confirmation", p_order_id=purchase_order_id, p_token=token,
                          p_outcome=outcome, p_stage=stage, p_error=error, p_retry_safe=safe,
                          p_details=envelope, p_actor=actor)
        except Exception as persistence_error:
            raise RuntimeError(
                "Confirmation error could not be recorded. Its durable claim remains locked; "
                "do not retry or send manually until reviewed."
            ) from persistence_error
        return {"status": outcome, "stage": stage, "error": error,
                "message": error + ("; recover using the existing invoice." if safe else "; manual review required, no automatic resend."),
                "invoice_number": str(number) if number else None,
                "retry_safe": safe, "approved": failed["status"] == "approved",
                "saved_invoice_row": saved[0] if stage in ("pdf", "prepare_send", "email") else None}

    # Deliberately outside the send try: if persistence fails after acceptance, NEVER retry Graph.
    try:
        final = _rpc("finish_order_confirmation", p_order_id=purchase_order_id, p_token=token,
                     p_outcome="sent", p_stage="complete", p_error=None, p_retry_safe=False,
                     p_details=envelope, p_actor=actor)
    except Exception as exc:
        raise RuntimeError(
            "Graph accepted the email, but status/audit persistence failed. Do not resend; "
            "the durable pending claim blocks duplicate attempts until manually reconciled."
        ) from exc
    return {"status": "sent", "invoice_number": final["confirmation_invoice_number"], "approved": True,
            "saved_invoice_row": saved[0],
            "message": f"Order approved \u00b7 Invoice {number} created \u00b7 "
                       "Graph send accepted (202); delivery not confirmed"}
