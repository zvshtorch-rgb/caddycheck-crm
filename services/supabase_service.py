"""Supabase service — cloud storage for projects, invoices, and tickets."""
import datetime
import logging
import re
from pathlib import Path
from typing import List, Optional, Dict, Any

logger = logging.getLogger(__name__)
SENT_INVOICE_BUCKET = "sent-invoices"
ORDER_PDF_BUCKET = "order-pdfs"
BANK_PAYMENT_BUCKET = "bank-payments"
LICENSE_CHANGE_LOG_TABLE = "license_change_log"
BANK_PAYMENTS_TABLE = "bank_payments"
BANK_PAYMENT_ALLOCATIONS_TABLE = "bank_payment_allocations"
SUPABASE_PAGE_SIZE = 1000


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


def _select_all_rows(table_name: str, order_column: str = "id") -> list[dict]:
    client = _get_client()
    all_rows: list[dict] = []
    offset = 0

    while True:
        resp = (
            client.table(table_name)
            .select("*")
            .order(order_column)
            .range(offset, offset + SUPABASE_PAGE_SIZE - 1)
            .execute()
        )
        batch = resp.data or []
        if not batch:
            break
        all_rows.extend(batch)
        if len(batch) < SUPABASE_PAGE_SIZE:
            break
        offset += SUPABASE_PAGE_SIZE

    return all_rows


def _parse_date(val) -> Optional[datetime.datetime]:
    if not val:
        return None
    try:
        return datetime.datetime.fromisoformat(str(val))
    except Exception:
        return None


def _normalize_license_change_log_entry(entry: Dict[str, Any]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    if entry.get("id") not in (None, ""):
        normalized["id"] = int(entry["id"])

    normalized["changed_at"] = str(entry.get("changed_at") or datetime.datetime.utcnow().isoformat())
    normalized["project_name"] = str(entry.get("project_name") or "").strip() or None
    normalized["country"] = str(entry.get("country") or "").strip() or None
    normalized["old_license_eop"] = str(entry.get("old_license_eop") or "").strip() or None
    normalized["new_license_eop"] = str(entry.get("new_license_eop") or "").strip() or None
    normalized["action"] = str(entry.get("action") or "").strip() or "Updated"
    normalized["updated_by"] = str(entry.get("updated_by") or "").strip() or None
    normalized["source_name"] = str(entry.get("source_name") or "").strip() or None
    normalized["notes"] = str(entry.get("notes") or "").strip() or None
    normalized["updated_at"] = datetime.datetime.utcnow().isoformat()
    return normalized


def _normalize_bank_payment_entry(entry: Dict[str, Any]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    if entry.get("id") not in (None, ""):
        normalized["id"] = int(entry["id"])

    normalized["created_at"] = str(entry.get("created_at") or datetime.datetime.utcnow().isoformat())
    normalized["updated_at"] = datetime.datetime.utcnow().isoformat()
    normalized["payment_date"] = str(entry.get("payment_date") or "").strip() or None
    normalized["invoice_number"] = int(entry["invoice_number"]) if entry.get("invoice_number") not in (None, "") else None
    normalized["source_name"] = str(entry.get("source_name") or "").strip() or None
    normalized["source_kind"] = str(entry.get("source_kind") or "").strip() or None
    normalized["payment_fingerprint"] = str(entry.get("payment_fingerprint") or "").strip() or None
    normalized["instructed_amount"] = float(entry["instructed_amount"]) if entry.get("instructed_amount") not in (None, "") else None
    normalized["received_amount"] = float(entry["received_amount"]) if entry.get("received_amount") not in (None, "") else None
    normalized["applied_amount"] = float(entry["applied_amount"]) if entry.get("applied_amount") not in (None, "") else None
    normalized["fee_amount"] = float(entry["fee_amount"]) if entry.get("fee_amount") not in (None, "") else None
    normalized["currency"] = str(entry.get("currency") or "EUR").strip() or "EUR"
    normalized["raw_text"] = str(entry.get("raw_text") or "").strip() or None
    normalized["parsed_payload"] = entry.get("parsed_payload") if entry.get("parsed_payload") is not None else {}
    normalized["notes"] = str(entry.get("notes") or "").strip() or None
    normalized["pdf_storage_bucket"] = str(entry.get("pdf_storage_bucket") or "").strip() or None
    normalized["pdf_storage_path"] = str(entry.get("pdf_storage_path") or "").strip() or None
    return normalized


def _normalize_bank_payment_allocation_entry(entry: Dict[str, Any]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    if entry.get("id") not in (None, ""):
        normalized["id"] = int(entry["id"])
    normalized["payment_id"] = int(entry["payment_id"])
    normalized["invoice_row_id"] = int(entry["invoice_row_id"]) if entry.get("invoice_row_id") not in (None, "") else None
    normalized["invoice_number"] = int(entry["invoice_number"]) if entry.get("invoice_number") not in (None, "") else None
    normalized["project_name"] = str(entry.get("project_name") or "").strip() or None
    normalized["maintenance_year"] = str(entry.get("maintenance_year") or "").strip() or None
    normalized["year"] = int(entry["year"]) if entry.get("year") not in (None, "") else None
    normalized["amount_applied"] = float(entry["amount_applied"]) if entry.get("amount_applied") not in (None, "") else None
    normalized["created_at"] = str(entry.get("created_at") or datetime.datetime.utcnow().isoformat())
    return normalized


# ── Projects ──────────────────────────────────────────────────────────────────

def load_projects() -> list:
    from models.project import Project
    from config.settings import get_project_overrides, canonical_project_name
    projects = []
    for row in _select_all_rows("projects", order_column="project_name"):
        inv_numbers: dict = {}
        for i in range(1, 10):
            v = row.get(f"m{i}y")
            if v is not None:
                try:
                    inv_numbers[i] = float(v)
                except Exception:
                    pass
        p = Project(
            project_name=canonical_project_name(row["project_name"]),
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
    from config.settings import canonical_project_name
    rows_by_name: Dict[str, Dict[str, Any]] = {}
    for p in projects:
        project_name = canonical_project_name(str(p.project_name or "").strip())
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
    """Delete project rows whose canonical name matches the selected project."""
    from config.settings import canonical_project_name

    cleaned = [canonical_project_name(name) for name in project_names if str(name).strip()]
    if not cleaned:
        return 0

    client = _get_client()
    deleted = 0
    for name in cleaned:
        resp = client.table("projects").select("project_name").execute()
        matching_names = {
            str(row.get("project_name") or "").strip()
            for row in (resp.data or [])
            if canonical_project_name(row.get("project_name")) == name
        }
        for raw_name in matching_names:
            client.table("projects").delete().eq("project_name", raw_name).execute()
            deleted += 1

    logger.info("Deleted %d project row(s) from Supabase", deleted)
    return deleted


def rename_invoice_project_names(rename_map: Dict[str, str]) -> int:
    """Rename invoice project names in Supabase and return the number of updated rows."""
    cleaned = {
        str(old_name).strip(): str(new_name).strip()
        for old_name, new_name in rename_map.items()
        if str(old_name).strip() and str(new_name).strip() and str(old_name).strip() != str(new_name).strip()
    }
    if not cleaned:
        return 0

    client = _get_client()
    updated = 0
    for old_name, new_name in cleaned.items():
        resp = client.table("invoices").update({"project_name": new_name}).eq("project_name", old_name).execute()
        updated += len(resp.data or [])

    logger.info("Renamed %d invoice row(s) in Supabase", updated)
    return updated


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
    from config.settings import canonical_project_name
    global _invoice_id_map, _invoice_number_project_id_map
    _invoice_id_map = {}
    _invoice_number_project_id_map = {}
    invoices = []
    for row in _select_all_rows("invoices"):
        inv = Invoice(
            invoice_number=float(row["invoice_number"]) if row.get("invoice_number") else None,
            project_name=canonical_project_name(row.get("project_name", "")),
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
    from config.settings import canonical_project_name
    to_update: list = []
    to_insert: list = []
    deduped_invoices: Dict[tuple, Any] = {}
    duplicate_count = 0

    for inv in invoices:
        inv.project_name = canonical_project_name(inv.project_name)
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
    from config.settings import canonical_project_name

    rows_to_insert = []
    for proj in sorted(projects, key=lambda p: p.project_name):
        if proj.num_cams <= 0:
            continue
        rows_to_insert.append({
            "invoice_number": str(invoice_number),
            "project_name": canonical_project_name(proj.project_name),
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
    from config.settings import canonical_project_name

    client = _get_client()
    query = client.table("tickets").select("*").order("created_at", desc=True)
    if project_name:
        query = query.eq("project_name", canonical_project_name(project_name))
    if status:
        query = query.eq("status", status)
    rows = query.execute().data or []
    for row in rows:
        row["project_name"] = canonical_project_name(row.get("project_name"))
    return rows


def create_ticket(
    project_name: str,
    title: str,
    description: str = "",
    priority: str = "Medium",
    subcategory: str = "",
) -> dict:
    from config.settings import canonical_project_name

    client = _get_client()
    resp = client.table("tickets").select("id").order("id", desc=True).limit(1).execute()
    next_seq = (resp.data[0]["id"] + 1) if resp.data else 1
    ticket_number = f"TK-{next_seq:04d}"
    row = {
        "ticket_number": ticket_number,
        "project_name": canonical_project_name(project_name),
        "title": title,
        "description": description,
        "priority": priority,
        "subcategory": subcategory or None,
        "status": "Open",
    }
    resp = client.table("tickets").insert(row).execute()
    return resp.data[0] if resp.data else {}


def update_ticket(ticket_id: int, **fields) -> dict:
    from config.settings import canonical_project_name

    client = _get_client()
    if "project_name" in fields:
        fields["project_name"] = canonical_project_name(fields.get("project_name"))
    fields["updated_at"] = datetime.datetime.utcnow().isoformat()
    if fields.get("status") in ("Resolved", "Closed"):
        fields.setdefault("resolved_at", datetime.datetime.utcnow().isoformat())
    resp = client.table("tickets").update(fields).eq("id", ticket_id).execute()
    return resp.data[0] if resp.data else {}


def delete_ticket(ticket_id: int) -> None:
    client = _get_client()
    client.table("tickets").delete().eq("id", ticket_id).execute()


# ── Orders ───────────────────────────────────────────────────────────────────

def load_orders() -> List[dict]:
    from config.settings import canonical_project_name

    client = _get_client()
    resp = client.table("orders").select("*").order("created_at", desc=True).execute()
    rows = resp.data or []
    for row in rows:
        row["project_name"] = canonical_project_name(row.get("project_name"))
    return rows


def _normalize_order_fields(fields: Dict[str, Any]) -> Dict[str, Any]:
    from config.settings import canonical_project_name

    normalized = dict(fields)
    if "project_name" in normalized:
        normalized["project_name"] = canonical_project_name(normalized.get("project_name"))
    for date_field in ("order_date", "requested_activation_date"):
        value = normalized.get(date_field)
        if isinstance(value, datetime.datetime):
            normalized[date_field] = value.date().isoformat()
        elif isinstance(value, datetime.date):
            normalized[date_field] = value.isoformat()
    return normalized


def create_order(**fields) -> dict:
    client = _get_client()
    row = _normalize_order_fields(fields)
    resp = client.table("orders").insert(row).execute()
    return resp.data[0] if resp.data else {}


def create_orders(rows: List[Dict[str, Any]]) -> int:
    normalized_rows = [_normalize_order_fields(row) for row in rows if row.get("project_name")]
    if not normalized_rows:
        return 0
    client = _get_client()
    client.table("orders").insert(normalized_rows).execute()
    return len(normalized_rows)


def update_order(order_id: int, **fields) -> dict:
    client = _get_client()
    normalized = _normalize_order_fields(fields)
    normalized["updated_at"] = datetime.datetime.utcnow().isoformat()
    resp = client.table("orders").update(normalized).eq("id", order_id).execute()
    return resp.data[0] if resp.data else {}


def delete_order(order_id: int) -> None:
    client = _get_client()
    client.table("orders").delete().eq("id", order_id).execute()


# ── Sent invoices ─────────────────────────────────────────────────────────────

def _normalize_str_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str):
        return [item.strip() for item in value.split(",") if item.strip()]
    return [str(value).strip()] if str(value).strip() else []


def _sent_invoice_storage_path(entry: Dict[str, Any], file_path: Path) -> str:
    year = entry.get("year") or "unknown"
    month = str(entry.get("month") or "unknown").strip().lower().replace(" ", "-")
    invoice_number = entry.get("invoice_number") or "unknown"
    sent_at = str(entry.get("sent_at") or datetime.datetime.utcnow().isoformat()).replace(":", "-")
    return f"{year}/{month}/{invoice_number}_{sent_at}_{file_path.name}"


def _ensure_storage_bucket(client, bucket_name: str) -> None:
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


def upload_sent_invoice_pdf(file_path: str | Path, storage_path: Optional[str] = None) -> Dict[str, str]:
    client = _get_client()
    pdf_path = Path(file_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"Sent invoice PDF was not found: {pdf_path}")

    _ensure_storage_bucket(client, SENT_INVOICE_BUCKET)
    target_path = storage_path or pdf_path.name
    file_bytes = pdf_path.read_bytes()

    try:
        client.storage.from_(SENT_INVOICE_BUCKET).remove([target_path])
    except Exception:
        pass

    client.storage.from_(SENT_INVOICE_BUCKET).upload(
        target_path,
        file_bytes,
        {"content-type": "application/pdf"},
    )
    return {
        "pdf_storage_bucket": SENT_INVOICE_BUCKET,
        "pdf_storage_path": target_path,
    }


def download_sent_invoice_pdf(bucket_name: str, storage_path: str) -> bytes:
    client = _get_client()
    return client.storage.from_(bucket_name).download(storage_path)


def _order_pdf_storage_path(filename: str) -> str:
    base = Path(filename).name or "order.pdf"
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_") or "order.pdf"
    stamp = datetime.datetime.utcnow().strftime("%Y/%m/%Y%m%d-%H%M%S")
    return f"{stamp}-{safe_name}"


def upload_order_pdf(file_bytes: bytes, filename: str, storage_path: Optional[str] = None) -> Dict[str, str]:
    import tempfile

    client = _get_client()
    _ensure_storage_bucket(client, ORDER_PDF_BUCKET)
    target_path = storage_path or _order_pdf_storage_path(filename)

    try:
        client.storage.from_(ORDER_PDF_BUCKET).remove([target_path])
    except Exception:
        pass

    suffix = Path(filename).suffix.lower() or ".pdf"
    content_type = "application/pdf" if suffix == ".pdf" else "application/octet-stream"

    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp_file:
        tmp_file.write(file_bytes)
        tmp_path = Path(tmp_file.name)

    try:
        client.storage.from_(ORDER_PDF_BUCKET).upload(
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
        "pdf_storage_bucket": ORDER_PDF_BUCKET,
        "pdf_storage_path": target_path,
    }


def download_order_pdf(bucket_name: str, storage_path: str) -> bytes:
    client = _get_client()
    return client.storage.from_(bucket_name).download(storage_path)


def create_order_pdf_signed_url(bucket_name: str, storage_path: str, expires_in: int = 3600) -> Optional[str]:
    client = _get_client()
    try:
        resp = client.storage.from_(bucket_name).create_signed_url(storage_path, expires_in)
    except Exception as exc:
        logger.warning("Could not create signed URL for order PDF: %s", exc)
        return None
    if isinstance(resp, dict):
        return resp.get("signedURL") or resp.get("signed_url") or resp.get("signedUrl")
    return None


def _bank_payment_storage_path(filename: str) -> str:
    base = Path(filename).name or "bank_payment.pdf"
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("_") or "bank_payment.pdf"
    stamp = datetime.datetime.utcnow().strftime("%Y/%m/%Y%m%d-%H%M%S")
    return f"{stamp}-{safe_name}"


def upload_bank_payment_pdf(file_bytes: bytes, filename: str, storage_path: Optional[str] = None) -> Dict[str, str]:
    import tempfile

    client = _get_client()
    _ensure_storage_bucket(client, BANK_PAYMENT_BUCKET)
    target_path = storage_path or _bank_payment_storage_path(filename)

    try:
        client.storage.from_(BANK_PAYMENT_BUCKET).remove([target_path])
    except Exception:
        pass

    suffix = Path(filename).suffix.lower() or ".pdf"
    content_type = "application/pdf" if suffix == ".pdf" else "application/octet-stream"

    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp_file:
        tmp_file.write(file_bytes)
        tmp_path = Path(tmp_file.name)

    try:
        client.storage.from_(BANK_PAYMENT_BUCKET).upload(
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
        "pdf_storage_bucket": BANK_PAYMENT_BUCKET,
        "pdf_storage_path": target_path,
    }


def download_bank_payment_pdf(bucket_name: str, storage_path: str) -> bytes:
    client = _get_client()
    return client.storage.from_(bucket_name).download(storage_path)


def create_bank_payment_pdf_signed_url(bucket_name: str, storage_path: str, expires_in: int = 3600) -> Optional[str]:
    client = _get_client()
    try:
        resp = client.storage.from_(bucket_name).create_signed_url(storage_path, expires_in)
    except Exception as exc:
        logger.warning("Could not create signed URL for bank payment PDF: %s", exc)
        return None
    if isinstance(resp, dict):
        return resp.get("signedURL") or resp.get("signed_url") or resp.get("signedUrl")
    return None


def create_sent_invoice_pdf_signed_url(bucket_name: str, storage_path: str, expires_in: int = 3600) -> Optional[str]:
    client = _get_client()
    try:
        resp = client.storage.from_(bucket_name).create_signed_url(storage_path, expires_in)
    except Exception as exc:
        logger.warning("Could not create signed URL for sent invoice PDF: %s", exc)
        return None
    if isinstance(resp, dict):
        return resp.get("signedURL") or resp.get("signed_url") or resp.get("signedUrl")
    return None


def _normalize_sent_invoice_entry(entry: Dict[str, Any]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    if entry.get("id") not in (None, ""):
        normalized["id"] = int(entry["id"])

    normalized["sent_at"] = str(entry.get("sent_at") or datetime.datetime.utcnow().isoformat())
    normalized["invoice_number"] = int(entry["invoice_number"]) if entry.get("invoice_number") not in (None, "") else None
    normalized["month"] = str(entry.get("month") or "").strip() or None
    normalized["year"] = int(entry["year"]) if entry.get("year") not in (None, "") else None
    normalized["pdf_filename"] = str(entry.get("pdf_filename") or "").strip() or None
    normalized["pdf_archive_path"] = str(entry.get("pdf_archive_path") or "").strip() or None
    normalized["pdf_storage_bucket"] = str(entry.get("pdf_storage_bucket") or "").strip() or None
    normalized["pdf_storage_path"] = str(entry.get("pdf_storage_path") or "").strip() or None
    normalized["recipients"] = _normalize_str_list(entry.get("recipients"))
    normalized["cc"] = _normalize_str_list(entry.get("cc"))
    normalized["subject"] = str(entry.get("subject") or "").strip() or None
    normalized["project_count"] = int(entry["project_count"]) if entry.get("project_count") not in (None, "") else None
    normalized["total_amount"] = float(entry["total_amount"]) if entry.get("total_amount") not in (None, "") else None
    normalized["saved_to_ledger"] = bool(entry.get("saved_to_ledger"))
    normalized["ledger_rows_added"] = int(entry["ledger_rows_added"]) if entry.get("ledger_rows_added") not in (None, "") else None
    normalized["source_name"] = str(entry.get("source_name") or "").strip() or None
    normalized["updated_at"] = datetime.datetime.utcnow().isoformat()
    return normalized


def _attach_sent_invoice_storage(entry: Dict[str, Any]) -> Dict[str, Any]:
    normalized = _normalize_sent_invoice_entry(entry)
    archive_path_text = normalized.get("pdf_archive_path")
    if not archive_path_text:
        return normalized

    archive_path = Path(archive_path_text)
    if not archive_path.exists():
        return normalized

    if normalized.get("pdf_storage_bucket") and normalized.get("pdf_storage_path"):
        return normalized

    try:
        storage_meta = upload_sent_invoice_pdf(
            archive_path,
            storage_path=_sent_invoice_storage_path(normalized, archive_path),
        )
        normalized.update(storage_meta)
    except Exception as exc:
        logger.warning("Could not upload sent invoice PDF to Supabase storage: %s", exc)
    return normalized


def load_sent_invoices() -> List[dict]:
    client = _get_client()
    resp = client.table("sent_invoices").select("*").order("sent_at", desc=True).execute()
    rows = resp.data or []
    for row in rows:
        row["recipients"] = _normalize_str_list(row.get("recipients"))
        row["cc"] = _normalize_str_list(row.get("cc"))
    return rows


def append_sent_invoice(entry: Dict[str, Any]) -> dict:
    client = _get_client()
    row = _attach_sent_invoice_storage(entry)
    resp = client.table("sent_invoices").insert(row).execute()
    saved = resp.data[0] if resp.data else {}
    if saved:
        saved["recipients"] = _normalize_str_list(saved.get("recipients"))
        saved["cc"] = _normalize_str_list(saved.get("cc"))
    return saved


def save_sent_invoices(entries: List[Dict[str, Any]]) -> None:
    client = _get_client()
    existing_rows = client.table("sent_invoices").select("id").execute().data or []
    keep_ids: set[int] = set()

    for entry in entries:
        row = _attach_sent_invoice_storage(entry)
        row_id = row.pop("id", None)
        if row_id is not None:
            resp = client.table("sent_invoices").update(row).eq("id", row_id).execute()
            keep_ids.add(row_id)
            if resp.data:
                continue
        resp = client.table("sent_invoices").insert(row).execute()
        if resp.data:
            saved_id = resp.data[0].get("id")
            if saved_id is not None:
                keep_ids.add(int(saved_id))

    for existing in existing_rows:
        existing_id = existing.get("id")
        if existing_id is None or int(existing_id) in keep_ids:
            continue
        client.table("sent_invoices").delete().eq("id", int(existing_id)).execute()


# ── Bank payments ────────────────────────────────────────────────────────────

def load_bank_payments() -> List[dict]:
    client = _get_client()
    resp = client.table(BANK_PAYMENTS_TABLE).select("*").order("payment_date", desc=True).execute()
    return resp.data or []


def load_bank_payment_allocations(payment_id: Optional[int] = None) -> List[dict]:
    client = _get_client()
    query = client.table(BANK_PAYMENT_ALLOCATIONS_TABLE).select("*").order("id", desc=False)
    if payment_id is not None:
        query = query.eq("payment_id", int(payment_id))
    return query.execute().data or []


def append_bank_payment(entry: Dict[str, Any]) -> dict:
    client = _get_client()
    row = _normalize_bank_payment_entry(entry)
    fingerprint = row.get("payment_fingerprint")
    existing_id = None
    if fingerprint:
        existing = client.table(BANK_PAYMENTS_TABLE).select("id").eq("payment_fingerprint", fingerprint).execute().data or []
        if existing:
            existing_id = int(existing[0]["id"])

    if existing_id is not None:
        resp = client.table(BANK_PAYMENTS_TABLE).update(row).eq("id", existing_id).execute()
    else:
        resp = client.table(BANK_PAYMENTS_TABLE).insert(row).execute()

    saved = resp.data[0] if resp.data else {}
    return saved


def save_bank_payments(entries: List[Dict[str, Any]]) -> None:
    client = _get_client()
    existing_rows = client.table(BANK_PAYMENTS_TABLE).select("id").execute().data or []
    keep_ids: set[int] = set()

    for entry in entries:
        row = _normalize_bank_payment_entry(entry)
        row_id = row.pop("id", None)
        if row_id is not None:
            resp = client.table(BANK_PAYMENTS_TABLE).update(row).eq("id", row_id).execute()
            keep_ids.add(row_id)
            if resp.data:
                continue
        resp = client.table(BANK_PAYMENTS_TABLE).insert(row).execute()
        if resp.data:
            saved_id = resp.data[0].get("id")
            if saved_id is not None:
                keep_ids.add(int(saved_id))

    for existing in existing_rows:
        existing_id = existing.get("id")
        if existing_id is None or int(existing_id) in keep_ids:
            continue
        client.table(BANK_PAYMENTS_TABLE).delete().eq("id", int(existing_id)).execute()


def save_bank_payment_bundles(entries: List[Dict[str, Any]]) -> None:
    """Save bank payments together with their allocation rows."""
    client = _get_client()
    existing_rows = client.table(BANK_PAYMENTS_TABLE).select("id").execute().data or []
    keep_ids: set[int] = set()

    for entry in entries:
        allocations = list(entry.get("allocations", []) or [])
        row = _normalize_bank_payment_entry(entry)
        row.pop("id", None)
        saved_payment = append_bank_payment(row)
        payment_id = saved_payment.get("id")
        if payment_id is None:
            continue
        payment_id = int(payment_id)
        keep_ids.add(payment_id)
        client.table(BANK_PAYMENT_ALLOCATIONS_TABLE).delete().eq("payment_id", payment_id).execute()
        normalized_allocations = []
        for allocation in allocations:
            normalized_allocations.append(_normalize_bank_payment_allocation_entry({**allocation, "payment_id": payment_id}))
        if normalized_allocations:
            client.table(BANK_PAYMENT_ALLOCATIONS_TABLE).insert(normalized_allocations).execute()

    for existing in existing_rows:
        existing_id = existing.get("id")
        if existing_id is None or int(existing_id) in keep_ids:
            continue
        client.table(BANK_PAYMENTS_TABLE).delete().eq("id", int(existing_id)).execute()


def append_bank_payment_with_allocations(entry: Dict[str, Any], allocations: List[Dict[str, Any]]) -> dict:
    client = _get_client()
    saved_payment = append_bank_payment(entry)
    payment_id = saved_payment.get("id")
    if payment_id is None:
        raise RuntimeError("Could not save bank payment record")

    payment_id = int(payment_id)
    client.table(BANK_PAYMENT_ALLOCATIONS_TABLE).delete().eq("payment_id", payment_id).execute()
    normalized_allocations = []
    for allocation in allocations:
        normalized = _normalize_bank_payment_allocation_entry({**allocation, "payment_id": payment_id})
        normalized_allocations.append(normalized)
    if normalized_allocations:
        client.table(BANK_PAYMENT_ALLOCATIONS_TABLE).insert(normalized_allocations).execute()
    saved_payment["allocations"] = normalized_allocations
    return saved_payment


# ── License change log ───────────────────────────────────────────────────────

def load_license_change_log() -> List[dict]:
    client = _get_client()
    resp = client.table(LICENSE_CHANGE_LOG_TABLE).select("*").order("changed_at", desc=True).execute()
    return resp.data or []


def append_license_change_log(entry: Dict[str, Any]) -> dict:
    client = _get_client()
    row = _normalize_license_change_log_entry(entry)
    resp = client.table(LICENSE_CHANGE_LOG_TABLE).insert(row).execute()
    return resp.data[0] if resp.data else {}


def save_license_change_log(entries: List[Dict[str, Any]]) -> None:
    client = _get_client()
    existing_rows = client.table(LICENSE_CHANGE_LOG_TABLE).select("id").execute().data or []
    keep_ids: set[int] = set()

    for entry in entries:
        row = _normalize_license_change_log_entry(entry)
        row_id = row.pop("id", None)
        if row_id is not None:
            resp = client.table(LICENSE_CHANGE_LOG_TABLE).update(row).eq("id", row_id).execute()
            keep_ids.add(row_id)
            if resp.data:
                continue
        resp = client.table(LICENSE_CHANGE_LOG_TABLE).insert(row).execute()
        if resp.data:
            saved_id = resp.data[0].get("id")
            if saved_id is not None:
                keep_ids.add(int(saved_id))

    for existing in existing_rows:
        existing_id = existing.get("id")
        if existing_id is None or int(existing_id) in keep_ids:
            continue
        client.table(LICENSE_CHANGE_LOG_TABLE).delete().eq("id", int(existing_id)).execute()


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
