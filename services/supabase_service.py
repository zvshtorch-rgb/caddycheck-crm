"""Supabase service — cloud storage for projects, invoices, and tickets."""
import datetime
import logging
from typing import List, Optional, Dict, Any

logger = logging.getLogger(__name__)


def _get_client():
    from supabase import create_client
    try:
        import streamlit as st
        cfg = st.secrets.get("supabase", {})
        url = cfg.get("url", "")
        key = cfg.get("service_role_key", cfg.get("anon_key", ""))
    except Exception:
        url = ""
        key = ""
    if not url or not key:
        raise RuntimeError("Supabase credentials not configured in st.secrets[supabase]")
    return create_client(url, key)


def _parse_date(val) -> Optional[datetime.datetime]:
    if not val:
        return None
    try:
        return datetime.datetime.fromisoformat(str(val))
    except Exception:
        return None


# ── Projects ──────────────────────────────────────────────────────────────────

def load_projects() -> list:
    from models.project import Project
    from config.settings import get_project_overrides
    client = _get_client()
    resp = client.table("projects").select("*").order("project_name").execute()
    projects = []
    for row in resp.data:
        inv_numbers: dict = {}
        for i in range(1, 10):
            v = row.get(f"m{i}y")
            if v is not None:
                try:
                    inv_numbers[i] = float(v)
                except Exception:
                    pass
        p = Project(
            project_name=row["project_name"],
            country=row.get("country") or "",
            num_cams=row.get("num_cameras") or 0,
            payment_month=row.get("payment_month") or "",
            installation_year=row.get("installation_year"),
            project_approval=row.get("project_approval") or "",
            activation_date=_parse_date(row.get("activation_date")),
            detection_type=row.get("detection_type") or "",
            cart_type=row.get("cart_type") or "",
            vim_version=row.get("vim_version") or "",
            status=row.get("status") or "",
            license_eop=_parse_date(row.get("license_eop")),
            caddy_back=row.get("caddy_back") or "",
            maintenance_invoice_numbers=inv_numbers,
            rate_y1_override=row.get("rate_y1_override"),
            rate_y2_override=row.get("rate_y2_override"),
        )
        projects.append(p)

    overrides = get_project_overrides()
    for proj in projects:
        key = proj.project_name.lower().strip()
        if key in overrides:
            proj.rate_y1_override = overrides[key].get("y1_rate")
            proj.rate_y2_override = overrides[key].get("y2_rate")

    logger.info("Loaded %d projects from Supabase", len(projects))
    return projects


def upsert_projects(projects: list) -> None:
    client = _get_client()
    rows = []
    for p in projects:
        row: Dict[str, Any] = {
            "project_name": p.project_name,
            "country": p.country or None,
            "num_cameras": p.num_cams or None,
            "payment_month": p.payment_month or None,
            "installation_year": p.installation_year,
            "project_approval": p.project_approval or None,
            "activation_date": p.activation_date.date().isoformat() if p.activation_date else None,
            "detection_type": p.detection_type or None,
            "cart_type": p.cart_type or None,
            "vim_version": p.vim_version or None,
            "status": p.status or None,
            "license_eop": p.license_eop.date().isoformat() if p.license_eop else None,
            "caddy_back": p.caddy_back or None,
            "rate_y1_override": p.rate_y1_override,
            "rate_y2_override": p.rate_y2_override,
        }
        for i in range(1, 10):
            row[f"m{i}y"] = (
                str(int(p.maintenance_invoice_numbers[i]))
                if i in p.maintenance_invoice_numbers else None
            )
        rows.append(row)

    batch_size = 50
    for i in range(0, len(rows), batch_size):
        client.table("projects").upsert(rows[i:i+batch_size], on_conflict="project_name").execute()

    logger.info("Upserted %d projects to Supabase", len(projects))


# ── Invoices ──────────────────────────────────────────────────────────────────

# Cache: (project_name_lower, maintenance_year, year) -> db row id
_invoice_id_map: Dict[tuple, int] = {}


def load_invoices() -> list:
    from models.invoice import Invoice
    global _invoice_id_map
    _invoice_id_map = {}
    client = _get_client()
    resp = client.table("invoices").select("*").execute()
    invoices = []
    for row in resp.data:
        inv = Invoice(
            invoice_number=float(row["invoice_number"]) if row.get("invoice_number") else None,
            project_name=row.get("project_name", ""),
            maintenance_year=row.get("maintenance_year", ""),
            payment_amount=float(row.get("payment_amount") or 0),
            cameras_number=row.get("cameras_number"),
            payment_date=_parse_date(row.get("payment_date")),
            paid=row.get("paid", "No"),
            year=row.get("year"),
        )
        invoices.append(inv)
        key = (
            row.get("project_name", "").lower().strip(),
            str(row.get("maintenance_year", "")),
            row.get("year"),
        )
        _invoice_id_map[key] = row["id"]

    logger.info("Loaded %d invoices from Supabase", len(invoices))
    return invoices


def upsert_invoices(invoices: list) -> None:
    client = _get_client()
    to_update: list = []
    to_insert: list = []

    for inv in invoices:
        key = (inv.project_name.lower().strip(), str(inv.maintenance_year), inv.year)
        db_id = _invoice_id_map.get(key)
        row: Dict[str, Any] = {
            "invoice_number": str(int(inv.invoice_number)) if inv.invoice_number else None,
            "project_name": inv.project_name,
            "maintenance_year": inv.maintenance_year,
            "payment_amount": inv.payment_amount,
            "cameras_number": int(inv.cameras_number) if inv.cameras_number else None,
            "payment_date": inv.payment_date.date().isoformat() if inv.payment_date else None,
            "paid": inv.paid,
            "year": inv.year,
        }
        if db_id is not None:
            row["id"] = db_id
            to_update.append(row)
        else:
            to_insert.append(row)

    if to_update:
        batch_size = 50
        for i in range(0, len(to_update), batch_size):
            client.table("invoices").upsert(to_update[i:i+batch_size]).execute()
    if to_insert:
        batch_size = 50
        for i in range(0, len(to_insert), batch_size):
            client.table("invoices").insert(to_insert[i:i+batch_size]).execute()

    logger.info("Updated %d + inserted %d invoices", len(to_update), len(to_insert))


def append_invoice_rows(invoice_number: int, projects: list, year: int) -> int:
    """Append invoice rows for a monthly batch. Returns count appended."""
    client = _get_client()
    resp = (
        client.table("invoices")
        .select("project_name")
        .eq("invoice_number", str(invoice_number))
        .execute()
    )
    existing = {row["project_name"].lower().strip() for row in resp.data}

    rows_to_insert = []
    for proj in sorted(projects, key=lambda p: p.project_name):
        if proj.num_cams <= 0:
            continue
        if proj.project_name.lower().strip() in existing:
            continue
        rows_to_insert.append({
            "invoice_number": str(invoice_number),
            "project_name": proj.project_name,
            "maintenance_year": proj.get_maintenance_year_label(year),
            "payment_amount": proj.get_expected_amount(year),
            "cameras_number": proj.num_cams,
            "paid": "No",
            "year": year,
        })

    if rows_to_insert:
        client.table("invoices").insert(rows_to_insert).execute()

    logger.info("Appended %d invoice rows for invoice #%d", len(rows_to_insert), invoice_number)
    return len(rows_to_insert)


def get_invoices_by_number(invoice_number: int) -> List[dict]:
    """Return all invoice rows (with id) for a given invoice number."""
    client = _get_client()
    resp = (
        client.table("invoices")
        .select("*")
        .eq("invoice_number", str(invoice_number))
        .order("project_name")
        .execute()
    )
    return resp.data


def mark_invoice_row_paid(
    db_id: int,
    payment_date: datetime.date,
    payment_amount: Optional[float] = None,
) -> None:
    """Mark a single invoice row as paid by its DB id."""
    client = _get_client()
    fields: Dict[str, Any] = {
        "paid": "Yes",
        "payment_date": payment_date.isoformat(),
    }
    if payment_amount is not None:
        fields["payment_amount"] = payment_amount
    client.table("invoices").update(fields).eq("id", db_id).execute()
    logger.info("Marked invoice row id=%d as paid (date=%s)", db_id, payment_date)


def get_next_invoice_number() -> int:
    """Return max invoice_number + 1 from the invoices table."""
    client = _get_client()
    resp = client.table("invoices").select("invoice_number").execute()
    max_no = 0
    for row in resp.data:
        try:
            n = int(row["invoice_number"])
            if n > max_no:
                max_no = n
        except Exception:
            pass
    return max_no + 1


# ── Tickets ───────────────────────────────────────────────────────────────────

def get_tickets(
    project_name: Optional[str] = None,
    status: Optional[str] = None,
) -> List[dict]:
    client = _get_client()
    query = client.table("tickets").select("*").order("created_at", desc=True)
    if project_name:
        query = query.eq("project_name", project_name)
    if status:
        query = query.eq("status", status)
    return query.execute().data


def create_ticket(
    project_name: str,
    title: str,
    description: str = "",
    priority: str = "Medium",
) -> dict:
    client = _get_client()
    resp = client.table("tickets").select("id").order("id", desc=True).limit(1).execute()
    next_seq = (resp.data[0]["id"] + 1) if resp.data else 1
    ticket_number = f"TK-{next_seq:04d}"
    row = {
        "ticket_number": ticket_number,
        "project_name": project_name,
        "title": title,
        "description": description,
        "priority": priority,
        "status": "Open",
    }
    resp = client.table("tickets").insert(row).execute()
    return resp.data[0] if resp.data else {}


def update_ticket(ticket_id: int, **fields) -> dict:
    client = _get_client()
    fields["updated_at"] = datetime.datetime.utcnow().isoformat()
    if fields.get("status") in ("Resolved", "Closed"):
        fields.setdefault("resolved_at", datetime.datetime.utcnow().isoformat())
    resp = client.table("tickets").update(fields).eq("id", ticket_id).execute()
    return resp.data[0] if resp.data else {}


def delete_ticket(ticket_id: int) -> None:
    client = _get_client()
    client.table("tickets").delete().eq("id", ticket_id).execute()
