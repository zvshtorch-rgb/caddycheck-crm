"""Order confirmation email (Order Approvals extension).

When an admin approves a pending purchase order in the CRM, a Year-1 invoice should be created, a PDF
generated from it, and a confirmation email sent to the customer contacts -- reusing the EXISTING,
already-guarded invoice-ledger write (`services.supabase_service.replace_invoice_rows`) and PDF renderer
(`services.pdf_service.generate_invoice_pdf_from_rows`), never a separate invoice-generation path.

This module currently implements the DRY-RUN preview only (`build_dry_run_preview`): it performs every
step up to and including PDF bytes generation and a read-only Microsoft Graph mailbox-reachability check,
but never writes to the invoice ledger, never updates the purchase order, and never sends an email. It is
the first of two rollout steps (see the Order Approvals page); the live orchestrator that actually
creates the invoice/PDF/email and marks the order approved is a deliberately separate, later change.

No recipient addresses are hardcoded here: Tom/Karen/Yoram's real addresses live in
st.secrets["order_confirmation"] (see config.settings.get_order_confirmation_config), never in source.

Customer/Network uses the canonical project-network mapping or a non-email customer name. Pending-order
camera quantities can be recovered read-only from the stored source PDF text. The displayed invoice
number is only a proposed candidate from a SELECT: it is never reserved or consumed by a preview.
"""
from __future__ import annotations

import datetime
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
    cameras: int, amount: float, invoice_number: int, sender: str,
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
    detail_lines.append(f"- Proposed invoice number: {invoice_number}")

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


def check_graph_mailbox_reachable(mailbox: str) -> Dict[str, Any]:
    """Safe, READ-ONLY check that the existing Graph app credentials can resolve the given mailbox
    (GET /users/{mailbox}) -- never sends mail. Never raises; returns {'ok': bool, 'message': str}."""
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
        token_resp.raise_for_status()
        access_token = token_resp.json()["access_token"]

        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{mailbox}",
            headers={"Authorization": f"Bearer {access_token}"}, timeout=15,
        )
        if resp.status_code == 200:
            return {"ok": True, "message": f"Mailbox '{mailbox}' found in the tenant."}
        return {"ok": False, "message": f"Graph could not resolve mailbox '{mailbox}' (HTTP {resp.status_code})."}
    except Exception as exc:
        return {"ok": False, "message": f"Could not verify mailbox reachability: {exc}"}


def _next_invoice_number_preview() -> int:
    """Read-only: the invoice number that WOULD be assigned next. Does not reserve/write anything, so
    calling it repeatedly for preview purposes is always safe."""
    from services.order_approval_service import _next_invoice_number, _get_client

    return _next_invoice_number(_get_client())


def build_dry_run_preview(purchase_order_id: str) -> Dict[str, Any]:
    """Everything the "Preview order confirmation (dry run)" button needs to display: validation
    errors/warnings, resolved sender/To/CC, the exact subject/body that would be sent, a preview invoice
    number + amount, the generated PDF bytes, and a Graph mailbox-reachability check.

    Writes NOTHING: no invoice row, no purchase-order update, no email, no audit log row. Safe to call
    repeatedly, including for an already-approved order."""
    order = get_purchase_order(purchase_order_id)
    if not order:
        return {"ok": False, "errors": ["Purchase order not found."], "warnings": []}

    validation = validate_order_for_year1_invoicing(order)
    cfg = get_order_confirmation_config()
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
    mailbox_check = check_graph_mailbox_reachable(cfg.get("sender", ""))

    result.update({
        "project_name": project_name, "country": country, "customer_display": customer_display,
        "po_reference": po_reference, "cameras": cameras, "amount": amount,
        "invoice_number_preview": invoice_number_preview, "subject": subject, "body": body,
        "pdf_bytes": pdf_bytes, "mailbox_check": mailbox_check,
    })
    return result
