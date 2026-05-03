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
    rows_by_name: Dict[str, Dict[str, Any]] = {}
    for p in projects:
        project_name = str(p.project_name or "").strip()
        if not project_name:
            continue
        row: Dict[str, Any] = {
            "project_name": project_name,
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
        rows_by_name[project_name.lower()] = row

    rows = list(rows_by_name.values())

    batch_size = 50
    for i in range(0, len(rows), batch_size):
        client.table("projects").upsert(rows[i:i+batch_size], on_conflict="project_name").execute()

    logger.info("Upserted %d projects to Supabase", len(rows))


def delete_projects(project_names: list[str]) -> int:
    """Delete project rows by exact project name and return the number removed."""
    cleaned = [str(name).strip() for name in project_names if str(name).strip()]
    if not cleaned:
        return 0

    client = _get_client()
    deleted = 0
    for name in cleaned:
        client.table("projects").delete().eq("project_name", name).execute()
        deleted += 1

    logger.info("Deleted %d project row(s) from Supabase", deleted)
    return deleted


# ── Invoices ──────────────────────────────────────────────────────────────────

# Cache lookups for matching in-memory invoices back to DB rows.
_invoice_id_map: Dict[tuple, int] = {}
_invoice_number_project_id_map: Dict[tuple, int] = {}


def _invoice_identity(project_name: str, maintenance_year: str, year: Optional[int]) -> tuple:
    return (
        str(project_name or "").strip().lower(),
        str(maintenance_year or "").strip(),
        int(year) if year not in (None, "") else None,
    )


def _invoice_number_project_identity(invoice_number: Any, project_name: str) -> Optional[tuple]:
    project_key = str(project_name or "").strip().lower()
    if not project_key or invoice_number in (None, ""):
        return None
    try:
        return (str(int(float(invoice_number))), project_key)
    except (TypeError, ValueError):
        return None


def _invoice_row(inv) -> Dict[str, Any]:
    return {
        "invoice_number": str(int(inv.invoice_number)) if inv.invoice_number else None,
        "project_name": inv.project_name,
        "maintenance_year": inv.maintenance_year,
        "payment_amount": inv.payment_amount,
        "cameras_number": int(inv.cameras_number) if inv.cameras_number else None,
        "payment_date": inv.payment_date.date().isoformat() if inv.payment_date else None,
        "paid": inv.paid,
        "year": inv.year,
    }


def load_invoices() -> list:
    from models.invoice import Invoice
    global _invoice_id_map, _invoice_number_project_id_map
    _invoice_id_map = {}
    _invoice_number_project_id_map = {}
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
        key = _invoice_identity(
            row.get("project_name", ""),
            row.get("maintenance_year", ""),
            row.get("year"),
        )
        _invoice_id_map[key] = row["id"]
        numbered_key = _invoice_number_project_identity(
            row.get("invoice_number"),
            row.get("project_name", ""),
        )
        if numbered_key is not None:
            _invoice_number_project_id_map[numbered_key] = row["id"]

    logger.info("Loaded %d invoices from Supabase", len(invoices))
    return invoices


def upsert_invoices(invoices: list) -> None:
    client = _get_client()
    to_update: list = []
    to_insert: list = []
    deduped_invoices: Dict[tuple, Any] = {}
    duplicate_count = 0

    for inv in invoices:
        numbered_key = _invoice_number_project_identity(inv.invoice_number, inv.project_name)
        logical_key = _invoice_identity(inv.project_name, inv.maintenance_year, inv.year)
        dedupe_key = numbered_key if numbered_key is not None else logical_key
        if dedupe_key in deduped_invoices:
            duplicate_count += 1
        deduped_invoices[dedupe_key] = inv

    for dedupe_key, inv in deduped_invoices.items():
        logical_key = _invoice_identity(inv.project_name, inv.maintenance_year, inv.year)
        db_id = None
        if len(dedupe_key) == 2:
            db_id = _invoice_number_project_id_map.get(dedupe_key)
        if db_id is None:
            db_id = _invoice_id_map.get(logical_key)
        row = _invoice_row(inv)
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

    if duplicate_count:
        logger.warning(
            "Collapsed %d duplicate invoice row(s) before saving to Supabase",
            duplicate_count,
        )
    logger.info("Updated %d + inserted %d invoices", len(to_update), len(to_insert))


def append_invoice_rows(invoice_number: int, projects: list, year: int) -> int:
    rows_to_insert = []
    for proj in sorted(projects, key=lambda p: p.project_name):
        if proj.num_cams <= 0:
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

    replace_invoice_rows(invoice_number, rows_to_insert)

    logger.info("Replaced invoice #%d with %d row(s)", invoice_number, len(rows_to_insert))
    return len(rows_to_insert)


def replace_invoice_rows(invoice_number: int, rows: list[Dict[str, Any]]) -> int:
    """Replace all rows for a single invoice number with the provided payload."""
    client = _get_client()
    invoice_number_str = str(int(invoice_number))
    client.table("invoices").delete().eq("invoice_number", invoice_number_str).execute()

    if not rows:
        logger.info("Cleared invoice #%s with no replacement rows", invoice_number_str)
        return 0

    batch_size = 50
    normalized_rows = []
    for row in rows:
        normalized = dict(row)
        normalized["invoice_number"] = invoice_number_str
        normalized_rows.append(normalized)

    for i in range(0, len(normalized_rows), batch_size):
        client.table("invoices").insert(normalized_rows[i:i + batch_size]).execute()

    logger.info("Replaced invoice #%s with %d rows", invoice_number_str, len(normalized_rows))
    return len(normalized_rows)


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


# ── Subscriptions ──────────────────────────────────────────────────────────────

def get_subscription(project_name: str) -> Optional[dict]:
    """Return the current subscription row for a project, or None."""
    client = _get_client()
    resp = (
        client.table("subscriptions")
        .select("*")
        .eq("project_name", project_name)
        .execute()
    )
    return resp.data[0] if resp.data else None


def upsert_subscription(
    project_name: str,
    valid_until: datetime.date,
    cameras_allowed: int,
    valid_from: Optional[datetime.date] = None,
    module_name: str = "Video Inform Profiler",
) -> dict:
    """Create or update the subscription record for a project."""
    client = _get_client()
    row: Dict[str, Any] = {
        "project_name":    project_name,
        "valid_until":     valid_until.isoformat(),
        "cameras_allowed": cameras_allowed,
        "valid_from":      (valid_from or datetime.date.today()).isoformat(),
        "module_name":     module_name,
        "status":          "active",
        "updated_at":      datetime.datetime.utcnow().isoformat(),
    }
    resp = (
        client.table("subscriptions")
        .upsert(row, on_conflict="project_name")
        .execute()
    )
    return resp.data[0] if resp.data else {}


def create_renewal_link(
    project_name: str,
    target_valid_until: datetime.date,
    cameras_allowed: int,
    invoice_number: Optional[str] = None,
    payment_amount: Optional[float] = None,
    expires_days: int = 30,
) -> str:
    """
    Generate a secure renewal token, persist it, and return the token string.
    The caller builds the full URL as: <base_url>/?token=<returned_token>
    """
    import secrets as _secrets
    client = _get_client()

    sub = get_subscription(project_name)
    token = _secrets.token_urlsafe(32)
    expires_at = datetime.datetime.utcnow() + datetime.timedelta(days=expires_days)

    row: Dict[str, Any] = {
        "project_name":       project_name,
        "subscription_id":    sub["id"] if sub else None,
        "token":              token,
        "expires_at":         expires_at.isoformat(),
        "target_valid_until": target_valid_until.isoformat(),
        "cameras_allowed":    cameras_allowed,
        "status":             "pending",
        "invoice_number":     str(invoice_number) if invoice_number else None,
        "payment_amount":     payment_amount,
    }
    client.table("renewal_links").insert(row).execute()
    logger.info("Created renewal token for %s → valid_until=%s", project_name, target_valid_until)
    return token


def process_renewal_token(token: str) -> dict:
    """
    Validate and apply a renewal token.
    Returns dict: {success, message, project_name?, valid_until?, cameras_allowed?}
    """
    client = _get_client()
    resp = client.table("renewal_links").select("*").eq("token", token).execute()

    if not resp.data:
        return {"success": False, "message": "Invalid renewal link."}

    link = resp.data[0]

    if link["status"] == "used":
        used_str = (link.get("used_at") or "")[:10]
        return {
            "success": False,
            "message": f"This renewal link was already used on {used_str}.",
        }

    # Check expiry
    try:
        expires_at = datetime.datetime.fromisoformat(
            link["expires_at"].replace("Z", "+00:00")
        )
        if datetime.datetime.now(datetime.timezone.utc) > expires_at:
            client.table("renewal_links").update({"status": "expired"}).eq("id", link["id"]).execute()
            return {"success": False, "message": "This renewal link has expired."}
    except Exception:
        pass

    # Apply renewal
    target_date = datetime.date.fromisoformat(link["target_valid_until"])
    project_name = link["project_name"]
    cameras = link.get("cameras_allowed") or 0

    upsert_subscription(
        project_name=project_name,
        valid_until=target_date,
        cameras_allowed=cameras,
    )

    # Mark token as used
    client.table("renewal_links").update({
        "status": "used",
        "used_at": datetime.datetime.utcnow().isoformat(),
    }).eq("id", link["id"]).execute()

    logger.info("Renewal token used: %s → %s until %s", project_name, cameras, target_date)
    return {
        "success":         True,
        "message":         f"Subscription renewed until {target_date.strftime('%B %d, %Y')}.",
        "project_name":    project_name,
        "valid_until":     target_date,
        "cameras_allowed": cameras,
    }


def get_renewal_links(project_name: Optional[str] = None) -> List[dict]:
    """Return renewal link history, newest first. Optionally filtered by project."""
    client = _get_client()
    query = (
        client.table("renewal_links")
        .select("*")
        .order("created_at", desc=True)
    )
    if project_name:
        query = query.eq("project_name", project_name)
    return query.execute().data
