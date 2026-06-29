"""
Order intake poller.

Reads the configured mailbox for new purchase-order emails (with a PDF
attachment), stores each PDF privately in Supabase, creates a ``purchase_orders``
record plus a secure single-use approval token, and emails the CEO an approval
link.

Run manually or on a schedule (cron / Task Scheduler / GitHub Actions):

    py order_intake_poll.py
    py order_intake_poll.py --dry-run

Required configuration (env vars or Streamlit secrets) is documented in the
README under "Purchase-order approval workflow".
"""
from __future__ import annotations

import argparse
import imaplib
import logging
import os
import sys

from services import order_approval_service as approvals
from services.email_intake_service import (
    IncomingMessage,
    extract_pdf_text,
    get_email_provider,
)
from services.email_service import send_simple_email

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("order_intake_poll")


def _get_setting(key: str, default: str = "") -> str:
    value = os.environ.get(key, "").strip()
    if value:
        return value
    try:
        import streamlit as st  # noqa: WPS433

        cfg = st.secrets.get("order_intake", {})
        return str(cfg.get(key, cfg.get(key.lower(), default))).strip()
    except Exception:
        return default


def _app_base_url() -> str:
    """Public base URL of the Streamlit app (used to build approval links)."""
    url = _get_setting("APP_BASE_URL", "").rstrip("/")
    return url or "http://localhost:8501"


def _ceo_recipients() -> list[str]:
    raw = _get_setting("ORDER_APPROVAL_CEO_EMAIL", "")
    return [addr.strip() for addr in raw.split(",") if addr.strip()]


def _should_skip_fetch_error(exc: Exception) -> bool:
    """Return True for mailbox access failures that should not fail the schedule."""
    if isinstance(exc, imaplib.IMAP4.error):
        message = str(exc).lower()
        return "login failed" in message or "authenticate" in message

    status_code = getattr(getattr(exc, "response", None), "status_code", None)
    return status_code in {401, 403}


def _build_approval_email(order: dict, token: str) -> tuple[str, str, str]:
    base = _app_base_url()
    approve_link = f"{base}/?approval={token}"
    project = order.get("project_name") or order.get("customer_name") or "(unknown)"
    amount = order.get("amount")
    amount_str = f"{amount:,.2f} {order.get('currency') or 'EUR'}" if amount is not None else "n/a"

    subject = f"Purchase order approval needed: {project}"
    body = (
        "A new customer purchase order requires your approval.\n\n"
        f"Project / Customer: {project}\n"
        f"Reference: {order.get('order_reference') or 'n/a'}\n"
        f"Amount: {amount_str}\n"
        f"From: {order.get('customer_email') or 'n/a'}\n\n"
        f"Review and decide here (link valid for 7 days):\n{approve_link}\n\n"
        "You can approve, reject, or request a correction on that page.\n\n"
        "— CaddyCheck CRM"
    )
    html = (
        f"<p>A new customer purchase order requires your approval.</p>"
        f"<ul>"
        f"<li><b>Project / Customer:</b> {project}</li>"
        f"<li><b>Reference:</b> {order.get('order_reference') or 'n/a'}</li>"
        f"<li><b>Amount:</b> {amount_str}</li>"
        f"<li><b>From:</b> {order.get('customer_email') or 'n/a'}</li>"
        f"</ul>"
        f'<p><a href="{approve_link}">Review &amp; decide</a> (link valid for 7 days).</p>'
        f"<p>— CaddyCheck CRM</p>"
    )
    return subject, body, html


def _extract_order_reference(*texts: str) -> str:
    """Best-effort detection of a customer PO/order reference from text."""
    import re

    patterns = [
        r"\b(?:purchase\s+order|order|po)\s*(?:no\.?|number|#|:)?\s*([A-Z0-9][A-Z0-9\-/]{3,})",
        r"\bPO[-#\s]?([0-9]{3,})\b",
    ]
    for text in texts:
        if not text:
            continue
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                return match.group(1).strip().upper()
    return ""


def _process_message(message: IncomingMessage, dry_run: bool) -> bool:
    """Process a single incoming email. Returns True on success."""
    if approvals.incoming_email_exists(message.message_id):
        logger.info("Skipping already-processed message: %s", message.message_id)
        return False

    logger.info("Processing message from %s | %s", message.from_address, message.subject)

    if dry_run:
        logger.info("[dry-run] Would store PDF '%s' (%d bytes) and notify CEO.",
                    message.pdf_filename, len(message.pdf_bytes))
        return True

    # 1. Store the PDF privately.
    storage = approvals.upload_purchase_order_pdf(message.pdf_bytes, message.pdf_filename)

    # 2. Best-effort text extraction.
    extracted_text = extract_pdf_text(message.pdf_bytes)

    # 3. Record the raw incoming email.
    email_row = approvals.record_incoming_email(
        message_id=message.message_id,
        provider=message.provider,
        from_address=message.from_address,
        subject=message.subject,
        received_at=message.received_at.isoformat() if message.received_at else None,
        body_text=message.body_text or None,
        pdf_filename=message.pdf_filename,
        pdf_storage_bucket=storage["pdf_storage_bucket"],
        pdf_storage_path=storage["pdf_storage_path"],
        extracted_text=extracted_text or None,
        processing_status="received",
    )

    # 4. Detect an order reference and look up any existing open revision first.
    order_reference = _extract_order_reference(message.subject, extracted_text)
    previous = None
    if order_reference:
        try:
            previous = approvals.find_current_open_order_by_reference(order_reference)
        except Exception as exc:
            logger.warning("Revision lookup failed for ref %s: %s", order_reference, exc)

    # 4b. Create the purchase order (a new revision if it supersedes an open one).
    new_revision = (previous.get("revision") or 1) + 1 if previous else 1
    order = approvals.create_purchase_order(
        incoming_email_id=email_row.get("id"),
        customer_email=message.from_address,
        project_name=None,
        order_reference=order_reference or None,
        summary=(message.subject or "")[:500] or None,
        pdf_storage_bucket=storage["pdf_storage_bucket"],
        pdf_storage_path=storage["pdf_storage_path"],
        revision=new_revision,
        status="pending_approval",
    )

    # 4c. Supersede the previous open revision, if any.
    if previous and previous.get("id") and previous["id"] != order.get("id"):
        try:
            approvals.supersede_purchase_order(previous["id"], order["id"])
            logger.info(
                "Order %s supersedes previous revision %s (ref %s).",
                order["id"], previous["id"], order_reference,
            )
        except Exception as exc:
            logger.warning("Revision linking failed for ref %s: %s", order_reference, exc)

    # 5. Create a secure single-use approval token.
    token = approvals.create_approval(order["id"])

    # 6. Notify the CEO.
    recipients = _ceo_recipients()
    if recipients:
        subject, body, html = _build_approval_email(order, token)
        try:
            send_simple_email(subject, body, recipients, html_body=html)
        except Exception as exc:
            logger.error("Failed to send CEO approval email: %s", exc)
            approvals.update_incoming_email(
                email_row["id"], processing_status="error", error_message=str(exc)
            )
            return False
    else:
        logger.warning("ORDER_APPROVAL_CEO_EMAIL not set; skipping CEO notification.")

    approvals.update_incoming_email(email_row["id"], processing_status="processed")
    logger.info("Created purchase order %s and approval token.", order.get("id"))
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Poll inbox for purchase-order emails.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Read the inbox and log what would happen without writing anything.",
    )
    args = parser.parse_args()

    try:
        provider = get_email_provider()
    except Exception as exc:
        logger.error("Could not initialise email provider: %s", exc)
        return 1

    if not provider.is_configured():
        logger.info(
            "Order intake mailbox is not configured (no '%s' credentials set); "
            "nothing to poll. Skipping.",
            provider.name,
        )
        return 0

    logger.info("Polling inbox via '%s' provider...", provider.name)
    try:
        messages = provider.fetch_unread_with_pdf()
    except Exception as exc:
        if _should_skip_fetch_error(exc):
            logger.warning("Mailbox access failed for provider '%s': %s. Skipping.", provider.name, exc)
            return 0
        logger.error("Failed to fetch emails: %s", exc)
        return 1

    if not messages:
        logger.info("No new purchase-order emails found.")
        return 0

    processed = 0
    for message in messages:
        try:
            if _process_message(message, args.dry_run):
                processed += 1
                if not args.dry_run:
                    provider.mark_processed(message)
        except Exception as exc:
            logger.exception("Error processing message %s: %s", message.message_id, exc)

    logger.info("Done. Processed %d of %d message(s).", processed, len(messages))
    return 0


if __name__ == "__main__":
    sys.exit(main())
