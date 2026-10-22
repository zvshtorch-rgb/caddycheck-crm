"""Order confirmation email (Order Approvals extension).

When an admin approves a pending purchase order in the CRM, a Year-1 invoice should be created, a PDF
generated from it, and a confirmation email sent to the customer contacts. Canonical Y1 payloads and the
shared PDF renderer are reused by order_confirmation_workflow; the insert-only ledger RPC uses the
existing registry/audit protections rather than the destructive replace_invoice_rows helper.

This module currently implements the DRY-RUN preview only (`build_dry_run_preview`): it performs every
step up to and including PDF bytes generation without a directory-read prerequisite,
but never writes to the invoice ledger, never updates the purchase order, and never sends an email. It is
the first of two rollout steps (see the Order Approvals page); the live orchestrator that actually
creates the invoice/PDF/email and marks the order approved is separately gated and disabled by default.

No recipient addresses are hardcoded here: Tom/Karen/Yoram's real addresses live in
st.secrets["order_confirmation"] (see config.settings.get_order_confirmation_config), never in source.

Customer/Network uses the canonical project-network mapping or a non-email customer name. Pending-order
camera quantities can be recovered read-only from the stored source PDF text. The displayed invoice
number is only a proposed candidate from a SELECT: it is never reserved or consumed by a preview.

The separate explicit internal send test sends only a TEST ONLY non-invoice PDF to the dedicated
[order_confirmation].test_recipient. It never calls approval, invoice, registry, or audit persistence.
"""
from __future__ import annotations

import datetime
import base64
import binascii
import io
import json
import re
from types import SimpleNamespace
from typing import Any, Dict, List, Optional

from config.settings import canonical_project_name, get_order_confirmation_config
from services.order_approval_service import _project_name_looks_unreliable, get_purchase_order

NEW_INSTALLATION_RATE_PER_CAMERA = 778.0


def _safe_str(value: object) -> str:
    return "" if value is None else str(value).strip()


def _safe_int(value: object, default: int = 0) -> int:
    try:
        if value is None or value == "":
            return default
        return int(float(value))
    except Exception:
        return default


def _safe_float(value: object, default: float = 0.0) -> float:
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default


def _fallback_order_reference(purchase_order_id: str, order_reference: str) -> str:
    """Same fallback already used by streamlit_app.py's `_auto_provision_order_and_project` for the
    Orders-page record, reused here so the two never disagree on what a PO's reference is."""
    return order_reference or f"PO-{str(purchase_order_id)[:8]}"


def resolve_project_country(project_name: str) -> str:
    """Best-effort: the country of an ALREADY-EXISTING project matching this name, read-only. Returns ""
    (never raises, never guesses) when no match is found -- country is a soft check (see
    validate_order_for_year1_invoicing), so an empty result just means the confirmation email omits it."""
    target = canonical_project_name(project_name).strip().lower()
    if not target:
        return ""
    try:
        from services.supabase_service import load_projects

        for project in load_projects():
            if _safe_str(project.project_name).strip().lower() == target:
                return _safe_str(project.country)
    except Exception:
        pass
    return ""


def resolve_ordered_cameras(order: Dict[str, Any]) -> int:
    """Read camera units from the order, its linked Orders record, or its stored source PDF text.

    Pending orders may not have an Orders record yet. Only explicit camera quantities in the source
    text are accepted: checkout counts and amount/rate arithmetic are not evidence of camera units.
    """
    direct = _safe_int(order.get("ordered_cameras") or order.get("num_cams"), default=0)
    if direct > 0:
        return direct
    purchase_order_id = _safe_str(order.get("id"))
    if not purchase_order_id:
        return 0
    from services.supabase_service import load_orders

    for row in load_orders():
        if purchase_order_id in _safe_str(row.get("notes")):
            cameras = _safe_int(row.get("ordered_cameras"), default=0)
            if cameras > 0:
                return cameras

    incoming_email_id = _safe_str(order.get("incoming_email_id"))
    if not incoming_email_id:
        return 0
    from services.order_approval_service import get_incoming_email

    incoming = get_incoming_email(incoming_email_id)
    text = _safe_str((incoming or {}).get("extracted_text"))
    patterns = (
        r"\bCaddyCheck\s+Camera\s*-\s*starting\s+licen[cs]e\s+(\d{1,4})\b",
        r"\b(?:number\s+of\s+cameras|camera\s+count|cameras|cams)\s*:\s*(\d{1,4})\b",
    )
    quantities = {
        int(match.group(1))
        for pattern in patterns
        for match in re.finditer(pattern, text, re.IGNORECASE)
        if int(match.group(1)) > 0
    }
    if len(quantities) > 1:
        raise ValueError("Conflicting camera quantities in the source PDF; verify the order before previewing.")
    return next(iter(quantities), 0)


def _customer_label(value: object) -> str:
    label = _safe_str(value)
    return label if "@" not in label and "\n" not in label and "\r" not in label else ""


def resolve_customer_network(order: Dict[str, Any]) -> str:
    """Use the existing deterministic network mapping, never a contact or sender email."""
    from services.ask_data.resolver import NetworkIndex

    project_name = canonical_project_name(_safe_str(order.get("project_name")))
    network = NetworkIndex([SimpleNamespace(project_name=project_name)]).network_of(project_name)
    return _customer_label(network) or _customer_label(order.get("customer_name"))


def compute_year1_amount(order: Dict[str, Any]) -> tuple[int, float]:
    """(camera_count, final_amount) for a Year-1 invoice from a purchase order's own fields (plus the
    matching Orders-page record for the camera count -- see resolve_ordered_cameras): the order's
    explicit payment_amount when positive, otherwise cameras * the standard new-installation rate -- the
    exact same fallback already used by the Orders-page "Create Invoice From Order" flow
    (streamlit_app.py) for a non-continuation order."""
    cameras = resolve_ordered_cameras(order)
    amount = _safe_float(order.get("amount") or order.get("payment_amount"), default=0.0)
    if amount <= 0 and cameras > 0:
        amount = round(cameras * NEW_INSTALLATION_RATE_PER_CAMERA, 2)
    return cameras, round(amount, 2)



def validate_order_for_year1_invoicing(order: Dict[str, Any]) -> Dict[str, List[str]]:
    """Hard/soft validation for Year-1 invoicing of an approved purchase order.

    Hard errors block invoice creation entirely. Soft warnings are reported but never block -- the
    corresponding confirmation-email field is simply omitted (see render_confirmation_email), per the
    approved design: project/site, PO reference, camera-count-or-amount, and the final computed price are
    hard requirements; country and customer billing details are soft (genuinely often unavailable)."""
    errors: List[str] = []
    warnings: List[str] = []

    project_name = canonical_project_name(_safe_str(order.get("project_name")))
    if not project_name or _project_name_looks_unreliable(project_name):
        errors.append(f"Project/site name is missing or looks unreliable ('{project_name}').")

    cameras, amount = compute_year1_amount(order)
    if cameras <= 0 and _safe_float(order.get("amount") or order.get("payment_amount"), default=0.0) <= 0:
        errors.append("Neither a camera count nor an explicit amount is available on this order.")
    if amount <= 0:
        errors.append("The computed Year-1 amount is not greater than zero.")

    country = resolve_project_country(project_name)
    if not country:
        warnings.append("Country is unknown for this project -- it will be omitted from the confirmation email.")

    customer_display = resolve_customer_network(order)
    if not customer_display:
        warnings.append("No customer/network name is resolved for this order -- 'Customer/Network' will be omitted.")

    return {"errors": errors, "warnings": warnings}


def render_confirmation_email(
    *, project_name: str, customer_display: str, country: str, po_reference: str,
    cameras: int, amount: float, invoice_number: int, sender: str, proposed: bool = True,
) -> tuple[str, str]:
    """(subject, body) for the order-confirmation email. Fields that are genuinely unavailable (empty
    string / zero) are OMITTED from the body entirely, never shown as a blank placeholder."""
    subject = f"Order Confirmation \u2013 CaddyCheck \u2013 {project_name} \u2013 Invoice {invoice_number}"

    detail_lines = [f"- Project/Site: {project_name}"]
    customer_label = _customer_label(customer_display)
    if customer_label:
        detail_lines.append(f"- Customer/Network: {customer_label}")
    if country:
        detail_lines.append(f"- Country: {country}")
    if po_reference:
        detail_lines.append(f"- Order / PO reference: {po_reference}")
    if cameras:
        detail_lines.append(f"- Number of cameras: {cameras}")
    detail_lines.append(f"- Year-1 amount: \u20ac{amount:,.2f}")
    label = "Proposed invoice number" if proposed else "Invoice number"
    detail_lines.append(f"- {label}: {invoice_number}")

    body_lines = [
        "Dear Tom and Karen,",
        "",
        "Thank you for your order.",
        "",
        f"We are pleased to confirm receipt and acceptance of the CaddyCheck order for {project_name}.",
        "",
        "Order details:",
        *detail_lines,
        "",
        "Please find attached the Year-1 invoice for this order.",
        "",
        "Thank you for your continued cooperation.",
        "",
        "Best regards,",
        "",
        "CaddyCheck",
        "Video Inform",
        sender or "",
    ]
    return subject, "\n".join(body_lines)


def build_year1_invoice_rows(order: Dict[str, Any], invoice_number: int) -> List[Dict[str, Any]]:
    """The exact single-project invoice-row payload shape used by the Orders-page 'Create Invoice From
    Order' flow (streamlit_app.py), built here as a reusable, testable function. Does NOT write
    anything -- the caller passes this to services.supabase_service.replace_invoice_rows (live mode only)."""
    project_name = canonical_project_name(_safe_str(order.get("project_name")))
    cameras, amount = compute_year1_amount(order)
    reference = _fallback_order_reference(order.get("id"), _safe_str(order.get("order_reference")))
    year = datetime.datetime.now(datetime.timezone.utc).year
    return [{
        "invoice_number": str(int(invoice_number)),
        "project_name": project_name,
        "maintenance_year": "Y1",
        "payment_amount": amount,
        "cameras_number": cameras or None,
        "payment_date": None,
        "paid": "No",
        "year": year,
        "invoice_type": "New installation",
        "for_month": None,
        "sent_at": None,
        "description": f"Order {reference}",
    }]


def generate_year1_invoice_pdf_bytes(invoice_number: int, rows: List[Dict[str, Any]]) -> bytes:
    """Build the invoice PDF via the SAME shared renderer used everywhere else in the app
    (services.pdf_service.generate_invoice_pdf_from_rows) -- never a separate PDF path. That function
    reads attributes (not dict keys), so the row dicts are wrapped as Invoice objects first."""
    from models.invoice import Invoice
    from services.pdf_service import generate_invoice_pdf_from_rows

    invoice_objects = [
        Invoice(
            invoice_number=row.get("invoice_number"),
            project_name=row.get("project_name"),
            maintenance_year=row.get("maintenance_year"),
            payment_amount=row.get("payment_amount"),
            cameras_number=row.get("cameras_number"),
            payment_date=row.get("payment_date"),
            paid=row.get("paid"),
            year=row.get("year"),
            invoice_type=row.get("invoice_type") or "",
            description=row.get("description"),
        )
        for row in rows
    ]
    return generate_invoice_pdf_from_rows(
        invoice_rows=invoice_objects, invoice_number=int(invoice_number),
        description=rows[0].get("description") if rows else None,
    )

def validate_invoice_pdf_bytes(pdf_bytes: bytes) -> Dict[str, int]:
    """Parse the generated in-memory document without downloading, saving, or writing CRM records."""
    import pdfplumber

    if not pdf_bytes:
        raise ValueError("Generated invoice PDF is empty.")
    if not re.match(rb"%PDF-[12]\.\d", pdf_bytes):
        raise ValueError("Generated invoice document has no valid PDF signature.")
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        if not pdf.pages:
            raise ValueError("Generated invoice PDF has no pages.")
        for page in pdf.pages:
            page.extract_text()
        return {"byte_count": len(pdf_bytes), "page_count": len(pdf.pages)}


def _graph_token_diagnostics(token: str, tenant_id: str, client_id: str) -> Dict[str, Any]:
    """Allowlisted claims for diagnosis only, not token validation or authorization decisions."""
    try:
        payload = token.split(".")[1]
        claims = json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))
        if not isinstance(claims, dict):
            raise ValueError("Token claims are not an object.")
    except (IndexError, ValueError, binascii.Error, UnicodeDecodeError):
        return {"claims_available": False, "note": "Token claims could not be decoded; authorization not inferred."}
    return {
        "claims_available": True,
        "application_roles": claims.get("roles", []),
        "delegated_scopes": claims.get("scp", ""),
        "tenant_matches_config": claims.get("tid", "").lower() == tenant_id.lower(),
        "client_matches_config": str(claims.get("azp") or claims.get("appid") or "").lower() == client_id.lower(),
        "note": "Decoded claims are diagnostic only; Exchange policy/RBAC scope is not tested.",
    }


def _graph_error_details(response, *sensitive_values: str) -> Dict[str, str]:
    try:
        body = response.json()
    except ValueError:
        return {"code": "NonJsonResponse", "message": "Graph returned a non-JSON error response."}
    error = body.get("error", {}) if isinstance(body, dict) else {}
    if isinstance(error, dict):
        details = {"code": str(error.get("code", "")), "message": str(error.get("message", ""))}
    else:
        details = {"code": str(error), "message": str(body.get("error_description", ""))}
    for key, value in details.items():
        for sensitive in sensitive_values:
            if sensitive:
                value = value.replace(sensitive, "[REDACTED]")
        details[key] = value[:1000]
    return details


def check_graph_mailbox_reachable(mailbox: str) -> Dict[str, Any]:
    """Read-only directory lookup, NOT a mailbox or send-authorization test. Never sends mail."""
    diagnostics: Dict[str, Any] = {
        "endpoint": f"https://graph.microsoft.com/v1.0/users/{mailbox}",
        "method": "GET", "permission_model": "application (client_credentials, scope .default)",
        "directory_lookup": "Optional diagnostic only. Do not broaden permissions for sending.",
        "send_authorization": "Not tested. This directory GET does not test Mail.Send or Exchange scope.",
    }
    if not mailbox:
        return {"ok": False, "message": "No sender mailbox configured."}
    try:
        import requests

        from services.email_service import _get_graph_setting

        tenant_id = _get_graph_setting("ORDER_INTAKE_GRAPH_TENANT_ID")
        client_id = _get_graph_setting("ORDER_INTAKE_GRAPH_CLIENT_ID")
        client_secret = _get_graph_setting("ORDER_INTAKE_GRAPH_CLIENT_SECRET")
        if not all([tenant_id, client_id, client_secret]):
            return {"ok": False, "message": "Graph credentials are not configured (ORDER_INTAKE_GRAPH_*)."}

        token_resp = requests.post(
            f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
            data={
                "client_id": client_id, "client_secret": client_secret,
                "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials",
            },
            timeout=15,
        )
        if token_resp.status_code != 200:
            diagnostics.update({
                "stage": "token", "http_status": token_resp.status_code,
                "error": _graph_error_details(token_resp, client_secret),
            })
            return {"ok": False, "message": "Graph token request failed; see safe diagnostics.",
                    "diagnostics": diagnostics}
        access_token = token_resp.json()["access_token"]
        diagnostics["token"] = _graph_token_diagnostics(access_token, tenant_id, client_id)

        resp = requests.get(
            diagnostics["endpoint"],
            headers={"Authorization": f"Bearer {access_token}"}, timeout=15,
        )
        diagnostics.update({"stage": "directory_lookup", "http_status": resp.status_code})
        if resp.status_code == 200:
            return {"ok": True, "message": f"Directory user '{mailbox}' found. Mailbox/send authorization is not verified.",
                    "diagnostics": diagnostics}
        diagnostics["error"] = _graph_error_details(resp, client_secret, access_token)
        status = "not_authorized" if resp.status_code == 403 else "not_checked"
        return {"ok": False, "status": status,
                "message": f"Directory lookup: {status.replace('_', ' ')} (HTTP {resp.status_code}); "
                "optional and not a prerequisite for sending.", "diagnostics": diagnostics}
    except Exception as exc:
        return {"ok": False, "status": "not_checked", "message": f"Directory lookup: not checked: {exc}"}


def _next_invoice_number_preview() -> int:
    """Read-only: the invoice number that WOULD be assigned next. Does not reserve/write anything, so
    calling it repeatedly for preview purposes is always safe."""
    from services.order_approval_service import _next_invoice_number, _get_client

    return _next_invoice_number(_get_client())


def build_dry_run_preview(purchase_order_id: str) -> Dict[str, Any]:
    """Everything the "Preview order confirmation (dry run)" button needs to display: validation
    errors/warnings, resolved sender/To/CC, the exact subject/body that would be sent, a preview invoice
    number + amount, and the generated PDF bytes. Directory lookup is optional and not performed.

    Writes NOTHING: no invoice row, no purchase-order update, no email, no audit log row. Safe to call
    repeatedly, including for an already-approved order."""
    order = get_purchase_order(purchase_order_id)
    if not order:
        return {"ok": False, "errors": ["Purchase order not found."], "warnings": []}

    validation = validate_order_for_year1_invoicing(order)
    cfg = get_order_confirmation_config(include_diagnostics=True)
    config_errors = []
    if not cfg.get("sender"):
        config_errors.append("No sender mailbox configured (st.secrets['order_confirmation']['sender']).")
    if not cfg.get("to"):
        config_errors.append("No 'To' recipients configured (st.secrets['order_confirmation']['to']).")

    errors = validation["errors"] + config_errors
    result: Dict[str, Any] = {
        "ok": not errors, "errors": errors, "warnings": validation["warnings"],
        "sender": cfg.get("sender", ""), "to": cfg.get("to", []), "cc": cfg.get("cc", []),
    }
    if errors:
        return result

    project_name = canonical_project_name(_safe_str(order.get("project_name")))
    cameras, amount = compute_year1_amount(order)
    country = resolve_project_country(project_name)
    customer_display = resolve_customer_network(order)
    po_reference = _fallback_order_reference(purchase_order_id, _safe_str(order.get("order_reference")))
    invoice_number_preview = _next_invoice_number_preview()

    subject, body = render_confirmation_email(
        project_name=project_name, customer_display=customer_display, country=country,
        po_reference=po_reference, cameras=cameras, amount=amount,
        invoice_number=invoice_number_preview, sender=cfg.get("sender", ""),
    )
    rows = build_year1_invoice_rows(order, invoice_number_preview)
    pdf_bytes = generate_year1_invoice_pdf_bytes(invoice_number_preview, rows)
    pdf_validation = validate_invoice_pdf_bytes(pdf_bytes)
    mailbox_check = {
        "status": "not_checked",
        "message": "Directory lookup: not checked (optional; directory-read permission is not requested). "
                   "This is not a prerequisite for sending. Send authorization is verified separately.",
    }

    result.update({
        "project_name": project_name, "country": country, "customer_display": customer_display,
        "po_reference": po_reference, "cameras": cameras, "amount": amount,
        "invoice_number_preview": invoice_number_preview, "subject": subject, "body": body,
        "pdf_bytes": pdf_bytes, "pdf_validation": pdf_validation, "mailbox_check": mailbox_check,
        "config_diagnostics": cfg.get("config_diagnostics", {}),
    })
    return result


TEST_CONFIRMATION_SENDER = "orders.caddycheck@video-inform.com"
TEST_CONFIRMATION_SUBJECT = "TEST ONLY \u2013 CaddyCheck Order Confirmation Email"


def get_confirmation_send_test_config() -> Dict[str, Any]:
    """Dedicated internal test destination; normal distribution lists are never a fallback."""
    cfg = get_order_confirmation_config()
    sender = _safe_str(cfg.get("sender"))
    recipient = _safe_str(cfg.get("test_recipient"))
    errors = []
    if sender.lower() != TEST_CONFIRMATION_SENDER:
        errors.append(f"Send test requires configured sender {TEST_CONFIRMATION_SENDER}.")
    if not re.fullmatch(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@video-inform\.com", recipient, re.IGNORECASE):
        errors.append("Configure one internal [order_confirmation].test_recipient address; no To/CC fallback is allowed.")
    if recipient.lower() in {_safe_str(address).lower() for address in cfg.get("to", [])}:
        errors.append("The test recipient must not be a normal order-confirmation To recipient.")
    return {"ok": not errors, "errors": errors, "sender": sender, "recipient": recipient,
            "subject": TEST_CONFIRMATION_SUBJECT}


def send_order_confirmation_test() -> Dict[str, Any]:
    """Explicit, single internal send using the shared Graph transport. No CRM access or retries."""
    import requests
    from reportlab.pdfgen import canvas
    from services.email_service import send_graph_email

    cfg = get_confirmation_send_test_config()
    if not cfg["ok"]:
        raise ValueError(" ".join(cfg["errors"]))
    output = io.BytesIO()
    pdf = canvas.Canvas(output)
    pdf.drawString(72, 780, "TEST ONLY - CaddyCheck Order Confirmation Email")
    pdf.drawString(72, 750, "Transport test document. NOT AN INVOICE. No order approved.")
    pdf.save()
    attachment = output.getvalue()
    validation = validate_invoice_pdf_bytes(attachment)
    result: Dict[str, Any] = {
        "endpoint": f"https://graph.microsoft.com/v1.0/users/{cfg['sender']}/sendMail",
        "method": "POST", "sender": cfg["sender"], "recipient": cfg["recipient"], "cc": [],
        "subject": cfg["subject"], "attachment_bytes": len(attachment),
        "attachment_pages": validation["page_count"], "http_status": None, "accepted": False,
    }
    try:
        response = send_graph_email(
            subject=cfg["subject"],
            body="TEST ONLY. Internal CaddyCheck Graph transport test. The attachment is not an invoice. "
                 "No purchase order has been approved and no invoice has been created.",
            recipients=[cfg["recipient"]], cc=[], sender_mailbox=cfg["sender"],
            attachment_bytes=attachment, attachment_filename="caddycheck_TEST_ONLY.pdf",
            attachment_content_type="application/pdf",
        )
    except requests.HTTPError as exc:
        http_response = exc.response
        if http_response is None:
            result.update({"status": "unknown", "error": {
                "code": "HttpErrorWithoutResponse",
                "message": "HTTP failure without response metadata. Do not retry automatically.",
            }})
        else:
            result.update({
                "status": "failed", "http_status": http_response.status_code,
                "failed_endpoint": http_response.url,
                "error": _graph_error_details(http_response),
            })
    except requests.RequestException:
        result.update({"status": "unknown", "error": {
            "code": "TransportOutcomeUnknown",
            "message": "Transport failed without a confirmed sendMail outcome. Check Sent Items before any retry.",
        }})
    else:
        result.update(response)
        result["status"] = "accepted" if response["accepted"] else "unexpected_response"
    return result
