"""CaddyCheck CRM — Streamlit web app (role-based access)."""
import datetime
import calendar
import io
import re
import sys
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import openpyxl

# Make sure project root is on the path
sys.path.insert(0, str(Path(__file__).parent))

import math

def _safe_int(v, default=0):
    """Convert v to int safely, returning default for None/NaN/empty."""
    try:
        if v is None:
            return default
        if isinstance(v, float) and math.isnan(v):
            return default
        return int(v)
    except Exception:
        return default

def _safe_float(v, default=0.0):
    """Convert v to float safely, returning default for None/NaN/empty."""
    try:
        if v is None:
            return default
        f = float(v)
        return default if math.isnan(f) else f
    except Exception:
        return default

def _safe_str(v):
    """Convert v to str, returning '' for None/NaN."""
    if v is None:
        return ""
    try:
        if isinstance(v, float) and math.isnan(v):
            return ""
    except Exception:
        pass
    return str(v)


def _invoice_category_label(invoice) -> str:
    return str(getattr(invoice, "maintenance_year", "")).strip().lower()


def _is_paid_trial_category(invoice) -> bool:
    checker = getattr(invoice, "is_paid_trial_category", None)
    if callable(checker):
        try:
            return bool(checker())
        except Exception:
            pass
    return "paid trial" in _invoice_category_label(invoice)


def _is_new_installation_category(invoice) -> bool:
    checker = getattr(invoice, "is_new_installation_category", None)
    if callable(checker):
        try:
            return bool(checker())
        except Exception:
            pass
    return _invoice_category_label(invoice) == "y1"


def _is_maintenance_category(invoice) -> bool:
    checker = getattr(invoice, "is_maintenance_category", None)
    if callable(checker):
        try:
            return bool(checker())
        except Exception:
            pass
    return not _is_new_installation_category(invoice) and not _is_paid_trial_category(invoice)

from config.settings import (
    MONTH_ORDER,
    get_email_config,
    save_email_config,
    load_sent_invoices_log,
    append_sent_invoice_log,
)
from services.supabase_service import (
    load_projects,
    load_invoices,
    upsert_projects,
    upsert_invoices,
    replace_invoice_rows,
    get_next_invoice_number as _supa_next_inv_no,
    append_invoice_rows as _supa_append_invoice,
)
from services.excel_service import (
    compute_debt_summaries,
    get_yearly_summary,
    get_projects_for_month,
    load_projects as load_projects_excel,
    load_invoices as load_invoices_excel,
    save_projects_to_excel,
    save_invoices_to_excel,
    get_next_invoice_number as _excel_next_inv_no,
    append_monthly_invoice_rows as _excel_append_invoice,
)
from services.invoice_service import generate_monthly_invoice, get_invoice_preview_data
from models.invoice import group_monthly_invoices

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CaddyCheck CRM",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Role-based login ──────────────────────────────────────────────────────────
# Passwords are stored in Streamlit secrets (secrets.toml or Streamlit Cloud secrets).
# Format in .streamlit/secrets.toml:
#   [passwords]
#   admin  = "your_admin_password"
#   viewer = "your_viewer_password"
#
# Roles: "admin" → can edit data   |   "viewer" → read-only

ROLES = {
    "admin":  {"label": "Admin",  "can_edit": True},
    "viewer": {"label": "Viewer", "can_edit": False},
}

def _check_login(username: str, password: str):
    """Return role string if credentials match, else None."""
    try:
        passwords = st.secrets.get("passwords", {})
    except Exception:
        passwords = {}
    # Fall back to hardcoded defaults if secrets not configured
    defaults = {"admin": "admin123", "viewer": "view123"}
    passwords = {**defaults, **passwords}
    if username in passwords and passwords[username] == password:
        return username  # role == username key
    return None

def _login_form():
    st.markdown("## 🔐 CaddyCheck CRM Login")
    with st.form("login_form"):
        username = st.selectbox("Role", list(ROLES.keys()), format_func=lambda k: ROLES[k]["label"])
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
    if submitted:
        role = _check_login(username, password)
        if role:
            st.session_state["role"] = role
            st.rerun()
        else:
            st.error("Incorrect password. Try again.")


def _consume_flash_success(current_page: str) -> str:
    message = st.session_state.get("_flash_success", "")
    if not message:
        return ""
    flash_page = st.session_state.get("_flash_success_page")
    if flash_page not in (None, current_page):
        return ""
    st.session_state.pop("_flash_success", None)
    st.session_state.pop("_flash_success_page", None)
    return message

# ── Public renewal token handler (no login required) ──────────────────────────
_renew_token = st.query_params.get("token", "")
if _renew_token:
    from services.supabase_service import process_renewal_token as _process_token
    st.markdown("## 🔑 Subscription Renewal")
    with st.spinner("Validating renewal link…"):
        try:
            _result = _process_token(_renew_token)
        except Exception as _e:
            _result = {"success": False, "message": str(_e)}
    if _result["success"]:
        st.success(f"✅ {_result['message']}")
        st.markdown(
            f"**Project:** {_result['project_name']}  \n"
            f"**Valid until:** {_result['valid_until'].strftime('%B %d, %Y')}  \n"
            f"**Cameras licensed:** {_result['cameras_allowed']}"
        )
        st.info("Your subscription has been updated. You can close this page.")
    else:
        st.error(f"❌ {_result['message']}")
        st.caption("Contact your account manager if you believe this is an error.")
    st.stop()

# Gate access
if "role" not in st.session_state:
    _login_form()
    st.stop()

ROLE       = st.session_state["role"]
CAN_EDIT   = ROLES[ROLE]["can_edit"]

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 16px 20px;
        border-left: 5px solid #2D6A9F;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        margin-bottom: 4px;
    }
    .metric-title { font-size: 13px; font-weight: 600; color: #666; margin-bottom: 4px; }
    .metric-value { font-size: 26px; font-weight: 800; color: #2C3E50; }
    .card-income  { border-left-color: #27AE60; }
    .card-paid    { border-left-color: #2980B9; }
    .card-debt    { border-left-color: #E74C3C; }
    .card-monthly { border-left-color: #16A085; }
    .card-yearly  { border-left-color: #2C3E50; }
    .card-projects{ border-left-color: #8E44AD; }
    .card-cameras { border-left-color: #F39C12; }
    .stDataFrame thead tr th { background: #2D6A9F !important; color: white !important; }
</style>
""", unsafe_allow_html=True)


# ── Data loading (cached) ─────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    source_name = "Supabase"
    try:
        projects = load_projects()
        invoices = load_invoices()
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            raise
        projects = load_projects_excel()
        invoices = load_invoices_excel()
        source_name = "Excel (local fallback)"
    debt_summaries = compute_debt_summaries(projects, invoices)
    yearly_summary = get_yearly_summary(invoices)
    return projects, invoices, debt_summaries, yearly_summary, source_name


def _is_excel_source(source_name: str) -> bool:
    return source_name.startswith("Excel")


def _save_projects(projects, source_name: str) -> None:
    if _is_excel_source(source_name):
        save_projects_to_excel(projects)
        return
    upsert_projects(projects)


def _save_invoices(invoices, source_name: str) -> None:
    if _is_excel_source(source_name):
        save_invoices_to_excel(invoices)
        return
    upsert_invoices(invoices)


def _get_next_invoice_number(invoices, source_name: str) -> int:
    if _is_excel_source(source_name):
        return _excel_next_inv_no(invoices)
    return _supa_next_inv_no()


def _normalize_query_text(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9+ ]", " ", _safe_str(value).lower())).strip()


def _extract_question_year(question: str) -> Optional[int]:
    match = re.search(r"\b(20\d{2})\b", question)
    return int(match.group(1)) if match else None


def _extract_question_invoice_number(question: str) -> Optional[int]:
    match = re.search(r"(?:invoice\s*#?|inv\s*#?)\s*(\d{3,6})", question, re.IGNORECASE)
    if match:
        return int(match.group(1))
    standalone = re.search(r"\b(8\d{3,5})\b", question)
    return int(standalone.group(1)) if standalone else None


def _project_license_date(project) -> Optional[datetime.date]:
    if not project.license_eop:
        return None
    if isinstance(project.license_eop, datetime.datetime):
        return project.license_eop.date()
    if isinstance(project.license_eop, datetime.date):
        return project.license_eop
    return None


def _add_months(base_date: datetime.date, months: int) -> datetime.date:
    month_index = base_date.month - 1 + months
    year = base_date.year + month_index // 12
    month = month_index % 12 + 1
    day = min(base_date.day, calendar.monthrange(year, month)[1])
    return datetime.date(year, month, day)


def _license_status(project, today: Optional[datetime.date] = None) -> str:
    today = today or datetime.date.today()
    license_date = _project_license_date(project)
    if license_date is None:
        return "Missing"
    if license_date < today:
        return "Expired"
    next_month = today.month % 12 + 1
    next_month_year = today.year if today.month < 12 else today.year + 1
    if license_date.month == next_month and license_date.year == next_month_year:
        return "Update Next Month"
    return "Active"


def _extract_question_month(question: str) -> str:
    q = _normalize_query_text(question)
    for month in MONTH_ORDER:
        full = month.lower()
        short = month[:3].lower()
        if re.search(rf"\b{re.escape(full)}\b", q) or re.search(rf"\b{re.escape(short)}\b", q):
            return month
    return ""


def _match_project_from_question(question: str, projects, invoices) -> str:
    q = _normalize_query_text(question)
    candidates = sorted({p.project_name for p in projects} | {i.project_name for i in invoices}, key=len, reverse=True)
    for name in candidates:
        normalized = _normalize_query_text(name)
        if normalized and normalized in q:
            return name
    return ""


def _match_country_from_question(question: str, projects) -> str:
    q = _normalize_query_text(question)
    countries = sorted({p.country for p in projects if p.country}, key=len, reverse=True)
    for country in countries:
        normalized = _normalize_query_text(country)
        if normalized and normalized in q:
            return country
    return ""


def _build_invoice_answer_df(invoice_rows, projects) -> pd.DataFrame:
    country_map = {p.project_name.lower().strip(): p.country for p in projects}
    return pd.DataFrame([
        {
            "Invoice #": str(int(inv.invoice_number)) if inv.invoice_number else "—",
            "Project": inv.project_name,
            "Country": country_map.get(inv.project_name.lower().strip(), ""),
            "Maint. Year": inv.maintenance_year,
            "Amount (€)": float(inv.payment_amount),
            "Paid": inv.paid,
            "Year": inv.year or "",
        }
        for inv in invoice_rows
    ])


def _build_project_answer_df(project_rows) -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Project": p.project_name,
            "Country": p.country,
            "# Cams": p.num_cams,
            "Payment Month": p.payment_month,
            "Status": p.status,
        }
        for p in project_rows
    ])


def _answer_data_question(question: str, projects, invoices, debt_summaries) -> tuple[str, Optional[pd.DataFrame]]:
    q = _normalize_query_text(question)
    if not q:
        return "Ask a question about projects, invoices, debt, or sent PDFs.", None

    project_name = _match_project_from_question(question, projects, invoices)
    country_name = _match_country_from_question(question, projects)
    year = _extract_question_year(question)
    invoice_number = _extract_question_invoice_number(question)
    month_name = _extract_question_month(question)

    if "next invoice" in q:
        next_inv = _get_next_invoice_number(invoices, _data_path)
        return f"The next invoice number is {next_inv}.", None

    if invoice_number is not None and ("invoice" in q or "inv" in q):
        invoice_rows = [inv for inv in invoices if _safe_int(inv.invoice_number, default=0) == invoice_number]
        if year is not None:
            invoice_rows = [inv for inv in invoice_rows if inv.year == year]

        if ("sent" in q or "emailed" in q or "email" in q):
            sent_rows = [row for row in load_sent_invoices_log() if _safe_int(row.get("invoice_number"), default=0) == invoice_number]
            if not sent_rows:
                return f"Invoice {invoice_number} has no sent PDF log entry.", None
            df = pd.DataFrame([
                {
                    "Sent At": _safe_str(row.get("sent_at", "")).replace("T", " ")[:19],
                    "Invoice #": _safe_int(row.get("invoice_number"), default=0),
                    "Month": _safe_str(row.get("month", "")),
                    "Year": _safe_int(row.get("year"), default=0),
                    "PDF": _safe_str(row.get("pdf_filename", "")),
                    "To": ", ".join(row.get("recipients", [])),
                    "Saved To Ledger": "Yes" if row.get("saved_to_ledger") else "No",
                }
                for row in reversed(sent_rows)
            ])
            return f"Found {len(df)} sent PDF log record(s) for invoice {invoice_number}.", df

        if not invoice_rows:
            return f"Invoice {invoice_number} is not in the invoice ledger.", None

        total_amount = sum(float(inv.payment_amount) for inv in invoice_rows)
        df = _build_invoice_answer_df(invoice_rows, projects)
        return f"Invoice {invoice_number} has {len(df)} row(s) totaling €{total_amount:,.0f}.", df

    if ("sent" in q or "emailed" in q or "email" in q) and ("pdf" in q or "invoice" in q):
        rows = load_sent_invoices_log()
        if invoice_number is not None:
            rows = [row for row in rows if _safe_int(row.get("invoice_number"), default=0) == invoice_number]
        if year is not None:
            rows = [row for row in rows if _safe_int(row.get("year"), default=0) == year]
        if project_name:
            rows = [row for row in rows if project_name.lower() in _safe_str(row.get("subject", "")).lower()]
        if not rows:
            return "No sent PDF invoices match that question.", None
        df = pd.DataFrame([
            {
                "Sent At": _safe_str(row.get("sent_at", "")).replace("T", " ")[:19],
                "Invoice #": _safe_int(row.get("invoice_number"), default=0),
                "Month": _safe_str(row.get("month", "")),
                "Year": _safe_int(row.get("year"), default=0),
                "PDF": _safe_str(row.get("pdf_filename", "")),
                "To": ", ".join(row.get("recipients", [])),
            }
            for row in reversed(rows)
        ])
        return f"Found {len(df)} sent PDF invoice record(s).", df

    if "active" in q and "project" in q and ("how many" in q or "count" in q or "number" in q):
        active_projects = [p for p in projects if p.is_active()]
        if country_name:
            active_projects = [p for p in active_projects if p.country == country_name]
        return f"There are {len(active_projects)} active project(s){' in ' + country_name if country_name else ''}.", None

    if ("project" in q or "bill" in q) and month_name:
        month_projects = get_projects_for_month(projects, month_name)
        if year is not None:
            filtered_projects = month_projects
        else:
            filtered_projects = month_projects
        if country_name:
            filtered_projects = [p for p in filtered_projects if p.country == country_name]
        if not filtered_projects:
            return f"No projects found for {month_name}{' ' + str(year) if year else ''}.", None
        df = _build_project_answer_df(filtered_projects)
        return f"Found {len(df)} project(s) billed in {month_name}{' ' + str(year) if year else ''}.", df

    if project_name and any(token in q for token in ["project", "status", "country", "payment month", "camera", "cams", "details"]):
        project = next((p for p in projects if p.project_name == project_name), None)
        if project is None:
            return f"I could not find project details for {project_name}.", None
        df = pd.DataFrame([{
            "Project": project.project_name,
            "Country": project.country,
            "# Cams": project.num_cams,
            "Payment Month": project.payment_month,
            "Install Year": project.installation_year or "",
            "Status": project.status,
        }])
        return f"Here are the project details for {project_name}.", df

    if "debt" in q or "unpaid" in q:
        debt_rows = [inv for inv in invoices if inv.is_unpaid()]
        if invoice_number is not None:
            debt_rows = [inv for inv in debt_rows if _safe_int(inv.invoice_number, default=0) == invoice_number]
        if year is not None:
            debt_rows = [inv for inv in debt_rows if inv.year == year]
        if country_name:
            project_names_in_country = {p.project_name for p in projects if p.country == country_name}
            debt_rows = [inv for inv in debt_rows if inv.project_name in project_names_in_country]
        if project_name:
            debt_rows = [inv for inv in debt_rows if inv.project_name == project_name]
        if "trial" in q:
            debt_rows = [inv for inv in debt_rows if _is_paid_trial_category(inv)]
        elif "y1" in q or "first year" in q or "new installation" in q:
            debt_rows = [inv for inv in debt_rows if _is_new_installation_category(inv)]
        elif "y2+" in q or "maintenance" in q:
            debt_rows = [inv for inv in debt_rows if _is_maintenance_category(inv)]
        if not debt_rows:
            return "No unpaid invoice rows match that question.", None
        total_amount = sum(float(inv.payment_amount) for inv in debt_rows)
        df = _build_invoice_answer_df(debt_rows, projects)
        return f"I found {len(df)} unpaid invoice row(s) totaling €{total_amount:,.0f}.", df

    if "top" in q and "debt" in q:
        debt_by_project = {}
        for inv in invoices:
            if not inv.is_unpaid():
                continue
            if year is not None and inv.year != year:
                continue
            debt_by_project.setdefault(inv.project_name, 0.0)
            debt_by_project[inv.project_name] += float(inv.payment_amount)
        if not debt_by_project:
            return "No unpaid debt rows match that question.", None
        df = pd.DataFrame([
            {"Project": name, "Debt (€)": amount}
            for name, amount in sorted(debt_by_project.items(), key=lambda item: item[1], reverse=True)[:10]
        ])
        return "Here are the top debt projects.", df

    if "invoice" in q and project_name:
        project_invoices = [inv for inv in invoices if inv.project_name == project_name]
        if year is not None:
            project_invoices = [inv for inv in project_invoices if inv.year == year]
        if not project_invoices:
            return f"No invoice rows found for {project_name}{' in ' + str(year) if year else ''}.", None
        df = _build_invoice_answer_df(project_invoices, projects)
        return f"I found {len(df)} invoice row(s) for {project_name}.", df

    if "country" in q and ("debt" in q or "unpaid" in q) and not country_name:
        return "I could not match the country name in that question.", None

    return (
        "I can answer questions like: 'What is the Y1 debt for 2026?', 'Show unpaid invoices for AD Denderleeuw', "
        "'How many active projects are in Belgium?', 'What invoices exist for Rewe Schorn - Bergheim?', "
        "'Show invoice 8676', 'Was invoice 8676 sent?', 'Which projects are billed in April?', or 'Show sent PDF invoices for 2026'.",
        None,
    )


def _find_invoice_header_row(worksheet) -> Optional[tuple[int, dict[str, int]]]:
    for row_idx in range(1, min(worksheet.max_row, 80) + 1):
        labels = {}
        for col_idx in range(1, min(worksheet.max_column, 20) + 1):
            cell_value = _safe_str(worksheet.cell(row=row_idx, column=col_idx).value).strip().lower()
            if not cell_value:
                continue
            normalized = re.sub(r"\s+", " ", cell_value)
            if "supermarket" in normalized or normalized == "project" or "project name" in normalized:
                labels["project_name"] = col_idx
            elif normalized in ("units", "cameras", "camera", "qty", "quantity"):
                labels["cameras_number"] = col_idx
            elif normalized in ("year", "maint. year", "maintenance year"):
                labels["maintenance_year"] = col_idx
            elif "line total" in normalized or normalized in ("amount", "amount (€)", "total"):
                labels["payment_amount"] = col_idx
        if {"project_name", "cameras_number", "payment_amount"}.issubset(labels):
            return row_idx, labels
    return None


def _infer_maintenance_year_label(title: str) -> str:
    normalized_title = _safe_str(title).strip().lower()
    if "trial" in normalized_title:
        return "Paid Trial-0.5Y"
    return ""


def _extract_invoice_number(worksheet) -> Optional[int]:
    preferred_cells = [(5, 10), (5, 9), (4, 10)]
    for row_idx, col_idx in preferred_cells:
        try:
            return int(worksheet.cell(row=row_idx, column=col_idx).value)
        except (TypeError, ValueError):
            pass

    for row_idx in range(1, min(worksheet.max_row, 20) + 1):
        for col_idx in range(1, min(worksheet.max_column, 12) + 1):
            label = _safe_str(worksheet.cell(row=row_idx, column=col_idx).value).strip().lower()
            if "invoice" in label and "#" in label or label == "invoice #":
                for neighbor_col in range(col_idx + 1, min(col_idx + 4, worksheet.max_column) + 1):
                    try:
                        return int(worksheet.cell(row=row_idx, column=neighbor_col).value)
                    except (TypeError, ValueError):
                        continue
    return None


def _extract_invoice_title_and_year(worksheet) -> tuple[str, Optional[int]]:
    candidate_text = []
    for row_idx in range(1, min(worksheet.max_row, 20) + 1):
        for col_idx in range(1, min(worksheet.max_column, 12) + 1):
            value = _safe_str(worksheet.cell(row=row_idx, column=col_idx).value).strip()
            if value:
                candidate_text.append(value)

    title = ""
    for value in candidate_text:
        if "maintenance" in value.lower() or "trial" in value.lower():
            title = value
            break
    if not title and candidate_text:
        title = candidate_text[0]

    year_match = re.search(r"(20\d{2})", title)
    if not year_match:
        for value in candidate_text:
            year_match = re.search(r"(20\d{2})", value)
            if year_match:
                break
    return title, int(year_match.group(1)) if year_match else None


def _parse_uploaded_invoice_xlsx(file_bytes: bytes) -> tuple[dict, list[dict]]:
    workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    worksheet = workbook.worksheets[0]

    invoice_number = _extract_invoice_number(worksheet)
    title, invoice_year = _extract_invoice_title_and_year(worksheet)
    header_info = _find_invoice_header_row(worksheet)
    if header_info is None:
        raise ValueError("Could not find invoice table headers in the uploaded XLSX")

    header_row_idx, columns = header_info

    rows: list[dict] = []
    for row_idx in range(header_row_idx + 1, worksheet.max_row + 1):
        project_name = _safe_str(worksheet.cell(row=row_idx, column=columns["project_name"]).value).strip()
        if not project_name:
            continue
        if project_name.lower() == "license period":
            break

        cameras_number = _safe_int(worksheet.cell(row=row_idx, column=columns["cameras_number"]).value, default=0)
        payment_amount = _safe_float(worksheet.cell(row=row_idx, column=columns["payment_amount"]).value, default=0.0)
        maint_year_col = columns.get("maintenance_year")
        maint_year = ""
        if maint_year_col is not None:
            maint_year = _safe_str(worksheet.cell(row=row_idx, column=maint_year_col).value).strip()
        if not maint_year:
            maint_year = _infer_maintenance_year_label(title)
        if not maint_year and payment_amount <= 0:
            break

        if cameras_number <= 0 and payment_amount <= 0:
            continue

        rows.append({
            "invoice_number": str(invoice_number) if invoice_number is not None else None,
            "project_name": project_name,
            "maintenance_year": maint_year,
            "payment_amount": payment_amount,
            "cameras_number": cameras_number or None,
            "payment_date": None,
            "paid": "No",
            "year": invoice_year,
        })

    metadata = {
        "invoice_number": invoice_number,
        "title": title,
        "year": invoice_year,
        "row_count": len(rows),
        "total_amount": sum(float(row["payment_amount"] or 0) for row in rows),
    }
    return metadata, rows


def _is_missing_supabase_table_error(exc: Exception, table_name: str) -> bool:
    message = str(exc).lower()
    return (
        "could not find the table" in message
        and table_name.lower() in message
    )


def _append_invoice_rows(invoice_number: int, projects, year: int, source_name: str) -> int:
    if _is_excel_source(source_name):
        return _excel_append_invoice(invoice_number=invoice_number, projects=projects, year=year)
    return _supa_append_invoice(invoice_number=invoice_number, projects=projects, year=year)


def _push_excel_to_github() -> tuple[bool, str]:
    """Commit the updated Excel file back to GitHub so changes survive app restarts."""
    try:
        gh_cfg = st.secrets.get("github", {})
        token    = gh_cfg.get("token", "")
        repo_name = gh_cfg.get("repo", "zvshtorch-rgb/caddycheck-crm")
        if not token:
            return False, "GitHub token not set in secrets — changes saved locally only."
        try:
            from github import Github
        except ImportError:
            return False, "PyGithub is not installed — changes saved locally only."
        from config.settings import get_data_paths
        data_file = get_data_paths()["projects_file"]
        with open(data_file, "rb") as f:
            content = f.read()
        g = Github(token)
        repo = g.get_repo(repo_name)
        gh_file = repo.get_contents("data/CaddyCheckProjectsInfo.xlsx")
        repo.update_file(
            "data/CaddyCheckProjectsInfo.xlsx",
            "CRM: update data via web app",
            content,
            gh_file.sha,
        )
        return True, "Saved and committed to GitHub."
    except Exception as e:
        return False, f"Local save OK but GitHub commit failed: {e}"


def card(title: str, value: str, css_class: str):
    st.markdown(
        f'<div class="metric-card {css_class}">'
        f'<div class="metric-title">{title}</div>'
        f'<div class="metric-value">{value}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ── Sidebar navigation ────────────────────────────────────────────────────────
st.sidebar.title("📊 CaddyCheck CRM")
st.sidebar.markdown("---")
page = st.sidebar.radio(
    "Navigation",
    ["📊 Dashboard", "❓ Ask Data", "🏗️ Projects", "🔐 Licenses", "🧾 Invoice Details", "💸 Debt Report", "📅 Monthly Invoice", "🎫 Tickets", "🏦 Bank Payment", "⚙️ Settings"],
    label_visibility="collapsed",
)
st.sidebar.markdown("---")
role_icon = "✏️" if CAN_EDIT else "👁️"
st.sidebar.caption(f"{role_icon} Logged in as **{ROLES[ROLE]['label']}**")
if st.sidebar.button("Refresh Data"):
    load_data.clear()
    st.cache_data.clear()
    st.rerun()
if st.sidebar.button("Logout"):
    del st.session_state["role"]
    st.rerun()

# Load data
try:
    projects, invoices, debt_summaries, yearly_summary, _data_path = load_data()
    st.sidebar.caption(f"Data: `{_data_path}`")
except Exception as e:
    st.error(f"Failed to load data: {e}")
    st.stop()

page_flash_success = _consume_flash_success(page)
if page_flash_success:
    st.toast(page_flash_success, icon="✅")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "📊 Dashboard":
    st.title("📊 Dashboard")
    if page_flash_success:
        st.success(page_flash_success)

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander("Filters", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        years = sorted({inv.year for inv in invoices if inv.year}, reverse=True)
        year_opts = ["All"] + [str(y) for y in years]
        sel_year   = col1.selectbox("Year",    year_opts)
        sel_month  = col2.selectbox("Month",   ["All"] + MONTH_ORDER)
        countries  = sorted({p.country for p in projects if p.country})
        sel_country= col3.selectbox("Country", ["All"] + countries)
        sel_status = col4.selectbox("Status",  ["All", "Paid", "Unpaid", "Cancelled"])

    # ── Filter invoices ───────────────────────────────────────────────────────
    proj_country = {p.project_name.lower(): p.country for p in projects}
    month_num = MONTH_ORDER.index(sel_month) + 1 if sel_month != "All" else None

    def filter_invoices(invs):
        result = []
        for inv in invs:
            if sel_year != "All" and inv.year != int(sel_year):
                continue
            if month_num and inv.payment_date and inv.payment_date.month != month_num:
                continue
            if sel_country != "All":
                pc = proj_country.get(inv.project_name.lower().strip(), "")
                if pc != sel_country:
                    continue
            if sel_status == "Paid" and not inv.is_paid():
                continue
            if sel_status == "Unpaid" and not inv.is_unpaid():
                continue
            if sel_status == "Cancelled" and not inv.is_cancelled():
                continue
            result.append(inv)
        return result

    def filter_projects(projs):
        if sel_country != "All":
            return [p for p in projs if p.country == sel_country]
        return projs

    f_inv  = filter_invoices(invoices)
    f_proj = filter_projects(projects)

    # ── Summary cards ─────────────────────────────────────────────────────────
    total_paid   = sum(i.payment_amount for i in f_inv if i.is_paid())
    total_unpaid = sum(i.payment_amount for i in f_inv if i.is_unpaid())
    total_income = total_paid + total_unpaid
    active_count = sum(1 for p in f_proj if p.is_active())
    total_cams   = sum(p.num_cams for p in f_proj)

    # Monthly/yearly ref year
    if sel_year != "All":
        ref_year = int(sel_year)
    else:
        paid_years = [i.year for i in f_inv if i.is_paid() and i.year]
        ref_year = max(paid_years) if paid_years else datetime.datetime.now().year

    yearly_val = sum(
        i.payment_amount for i in f_inv
        if i.is_paid() and i.year == ref_year
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: card("Total Income",            f"€{total_income:,.0f}", "card-income")
    with c2: card("Total Paid",              f"€{total_paid:,.0f}",   "card-paid")
    with c3: card("Total Debt",              f"€{total_unpaid:,.0f}", "card-debt")
    with c4: card(f"Yearly Income ({ref_year})",  f"€{yearly_val:,.0f}",  "card-yearly")
    with c5: card("Active Projects",         str(active_count),       "card-projects")
    with c6: card("Total Cameras",           str(total_cams),         "card-cameras")

    st.markdown("---")

    # ── Cameras by Project table ───────────────────────────────────────────────
    st.subheader("Cameras by Project")
    sorted_proj = sorted(f_proj, key=lambda p: (0 if p.is_active() else 1, p.project_name))
    proj_df = pd.DataFrame([{
        "Project Name":    _safe_str(p.project_name),
        "Country":         _safe_str(p.country),
        "# Cams":          _safe_int(p.num_cams),
        "Payment Month":   _safe_str(p.payment_month),
        "Install Year":    _safe_str(p.installation_year),
        "Status":          _safe_str(p.status),
    } for p in sorted_proj])

    def color_status(val):
        if str(val).strip().lower() == "active":
            return "color: #27AE60; font-weight: bold"
        return "color: #E74C3C"

    st.dataframe(
        proj_df.style.map(color_status, subset=["Status"]) if "Status" in proj_df.columns and len(proj_df) > 0 else proj_df,
        use_container_width=True,
        height=300,
    )

    st.markdown("---")

    # ── Trend Chart ───────────────────────────────────────────────────────────
    st.subheader("Trends")
    cc1, cc2, cc3, cc4 = st.columns([2, 2, 1, 1])
    metric     = cc1.selectbox("Show",       ["Income (Paid)", "Income (All)", "Active Projects", "Cameras"], key="ch_metric")
    resolution = cc2.selectbox("Resolution", ["Yearly", "Monthly"], key="ch_res")
    all_years  = sorted({inv.year for inv in invoices if inv.year})
    if not all_years:
        all_years = [datetime.datetime.now().year]
    from_yr = cc3.selectbox("From Year", [int(y) for y in all_years], index=0, key="ch_from")
    to_yr   = cc4.selectbox("To Year",   [int(y) for y in all_years], index=len(all_years)-1, key="ch_to")
    if from_yr > to_yr:
        from_yr, to_yr = to_yr, from_yr

    is_income = metric.startswith("Income")
    y_label   = "EUR (€)" if is_income else "Count"

    if resolution == "Yearly":
        labels, values = [], []
        for yr in range(from_yr, to_yr + 1):
            labels.append(str(yr))
            if metric == "Income (Paid)":
                v = sum(i.payment_amount for i in invoices if i.is_paid() and i.year == yr)
            elif metric == "Income (All)":
                v = sum(i.payment_amount for i in invoices if i.year == yr)
            elif metric == "Active Projects":
                v = sum(1 for p in projects if p.installation_year and p.installation_year <= yr and p.is_active())
            else:
                v = sum(p.num_cams for p in projects if p.installation_year and p.installation_year <= yr and p.is_active())
            values.append(float(v))

        fig = px.bar(
            x=labels, y=values,
            labels={"x": "Year", "y": y_label},
            title=f"{metric} — Yearly ({from_yr}–{to_yr})",
            color_discrete_sequence=["#2980B9"],
        )
        fig.update_traces(hovertemplate="<b>%{x}</b><br>" + y_label + ": %{y:,.0f}<extra></extra>")
        fig.update_layout(showlegend=False, height=380)

    else:  # Monthly
        rows = []
        for yr in range(from_yr, to_yr + 1):
            for mo in range(1, 13):
                if metric == "Income (Paid)":
                    v = sum(i.payment_amount for i in invoices
                            if i.is_paid() and i.payment_date
                            and i.payment_date.year == yr and i.payment_date.month == mo)
                elif metric == "Income (All)":
                    v = sum(i.payment_amount for i in invoices
                            if i.payment_date
                            and i.payment_date.year == yr and i.payment_date.month == mo)
                elif metric == "Active Projects":
                    v = sum(1 for p in projects if p.installation_year and p.installation_year <= yr and p.is_active())
                else:
                    v = sum(p.num_cams for p in projects if p.installation_year and p.installation_year <= yr and p.is_active())
                rows.append({"date": datetime.date(yr, mo, 1), "value": float(v)})

        df_line = pd.DataFrame(rows)
        fig = px.line(
            df_line, x="date", y="value",
            labels={"date": "Month", "value": y_label},
            title=f"{metric} — Monthly ({from_yr}–{to_yr})",
            color_discrete_sequence=["#2980B9"],
            markers=True,
        )
        fig.update_traces(hovertemplate="<b>%{x|%b %Y}</b><br>" + y_label + ": %{y:,.0f}<extra></extra>")
        fig.update_layout(height=380)

    st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: ASK DATA
# ══════════════════════════════════════════════════════════════════════════════
elif page == "❓ Ask Data":
    st.title("❓ Ask Data")
    st.caption("Ask questions about projects, invoices, debt, or sent PDF invoices.")

    quick_question_cols = st.columns(4)
    quick_questions = [
        "What is the Y1 debt for 2026?",
        "Show unpaid invoices for AD Denderleeuw",
        "Show invoice 8669",
        "Show sent PDF invoices for 2026",
    ]
    for idx, quick_question in enumerate(quick_questions):
        if quick_question_cols[idx].button(quick_question, key=f"ask_quick_{idx}", use_container_width=True):
            st.session_state["ask_data_question"] = quick_question
            answer_text, answer_df = _answer_data_question(quick_question, projects, invoices, debt_summaries)
            st.session_state["ask_data_answer_text"] = answer_text
            st.session_state["ask_data_answer_df"] = answer_df.to_dict(orient="records") if answer_df is not None else None
            st.rerun()

    with st.form("ask_data_form"):
        question = st.text_area(
            "Question",
            value=st.session_state.get("ask_data_question", ""),
            height=100,
            placeholder="Examples: What is the Y2+ debt for 2026? Show unpaid invoices for AD Denderleeuw. Show invoice 8676. Which projects are billed in April?",
        )
        submitted = st.form_submit_button("Ask")

    if submitted:
        st.session_state["ask_data_question"] = question
        answer_text, answer_df = _answer_data_question(question, projects, invoices, debt_summaries)
        st.session_state["ask_data_answer_text"] = answer_text
        st.session_state["ask_data_answer_df"] = answer_df.to_dict(orient="records") if answer_df is not None else None

    answer_text = st.session_state.get("ask_data_answer_text")
    answer_rows = st.session_state.get("ask_data_answer_df")

    if answer_text:
        st.success(answer_text)
    else:
        st.info(
            "Try: 'What is the Y1 debt for 2026?', 'Show unpaid invoices for Proxy Muizen', "
            "'How many active projects are in Belgium?', 'Show invoice 8676', or 'Show sent PDF invoices'."
        )

    if answer_rows:
        st.dataframe(pd.DataFrame(answer_rows), use_container_width=True, hide_index=True, height=420)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: PROJECTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🏗️ Projects":
    st.title("🏗️ Projects")
    if page_flash_success:
        st.success(page_flash_success)

    def _parse_project_date(value):
        if value in (None, ""):
            return None
        if isinstance(value, datetime.datetime):
            return value
        if isinstance(value, datetime.date):
            return datetime.datetime.combine(value, datetime.time.min)
        try:
            return datetime.datetime.strptime(str(value), "%Y-%m-%d")
        except Exception:
            return None

    # Filters
    col1, col2, col3 = st.columns(3)
    countries = sorted({p.country for p in projects if p.country})
    project_statuses = sorted({p.status for p in projects if p.status})
    install_year_options = [""] + [str(year) for year in range(datetime.date.today().year + 1, 2014, -1)]
    payment_month_options = [""] + MONTH_ORDER
    sel_country = col1.selectbox("Country", ["All"] + countries, key="proj_country")
    sel_status  = col2.selectbox("Status",  ["All", "Active", "Offline"], key="proj_status")
    search      = col3.text_input("Search project name", key="proj_search")

    filtered = projects
    if sel_country != "All":
        filtered = [p for p in filtered if p.country == sel_country]
    if sel_status == "Active":
        filtered = [p for p in filtered if p.is_active()]
    elif sel_status == "Offline":
        filtered = [p for p in filtered if not p.is_active()]
    if search:
        filtered = [p for p in filtered if search.lower() in p.project_name.lower()]

    filtered = sorted(filtered, key=lambda p: (0 if p.is_active() else 1, p.project_name))

    st.caption(f"Showing {len(filtered)} of {len(projects)} projects")

    df = pd.DataFrame([{
        "Project Name":    _safe_str(p.project_name),
        "Country":         _safe_str(p.country),
        "# Cams":          _safe_int(p.num_cams),
        "Payment Month":   _safe_str(p.payment_month),
        "Install Year":    _safe_str(p.installation_year),
        "Activation Date": p.activation_date.date() if p.activation_date else None,
        "Status":          _safe_str(p.status),
        "License EOP":     p.license_eop.date() if p.license_eop else None,
    } for p in filtered])

    def color_status(val):
        if str(val).strip().lower() == "active":
            return "color: #27AE60; font-weight: bold"
        return "color: #E74C3C"

    if CAN_EDIT:
        st.info("✏️ Admin mode: you can edit cells directly. Click **Save Changes** when done.")

        if st.button("➕ Add New Project", key="btn_add_proj"):
            st.session_state["add_proj_row"] = st.session_state.get("add_proj_row", 0) + 1

        _empty_proj = {"Project Name": "", "Country": "", "# Cams": 0,
                       "Payment Month": "", "Install Year": "",
                       "Activation Date": None, "Status": "Active", "License EOP": None}
        n_new = st.session_state.get("add_proj_row", 0)
        if n_new:
            empty_rows = pd.DataFrame([_empty_proj] * n_new)
            df_edit = pd.concat([empty_rows, df.reset_index(drop=True)], ignore_index=True)
        else:
            df_edit = df.reset_index(drop=True)

        edited_df = st.data_editor(
            df_edit,
            use_container_width=True,
            height=600,
            num_rows="dynamic",
            column_config={
                "Country": st.column_config.SelectboxColumn(
                    "Country",
                    options=[""] + countries,
                ),
                "Install Year": st.column_config.SelectboxColumn(
                    "Install Year",
                    options=install_year_options,
                ),
                "Payment Month": st.column_config.SelectboxColumn(
                    "Payment Month",
                    options=payment_month_options,
                ),
                "Activation Date": st.column_config.DateColumn(
                    "Activation Date",
                    format="YYYY-MM-DD",
                ),
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    options=([""] + project_statuses) if project_statuses else ["", "Active", "Offline"],
                ),
                "License EOP": st.column_config.DateColumn(
                    "License EOP",
                    format="YYYY-MM-DD",
                ),
            },
            key=f"proj_editor_{n_new}",
        )
        if st.button("💾 Save Changes", key="save_projects"):
            from models.project import Project as ProjectModel
            proj_map = {p.project_name: p for p in projects}
            new_count = 0
            for _, row in edited_df.iterrows():
                name = _safe_str(row.get("Project Name", "")).strip()
                if not name:
                    continue
                p = proj_map.get(name)
                if p is None:
                    # New project row
                    p = ProjectModel(project_name=name)
                    projects.append(p)
                    proj_map[name] = p
                    new_count += 1
                p.country           = _safe_str(row.get("Country", ""))
                p.num_cams          = _safe_int(row.get("# Cams", 0))
                p.payment_month     = _safe_str(row.get("Payment Month", ""))
                p.installation_year = _safe_int(row.get("Install Year")) or None
                p.status            = _safe_str(row.get("Status", ""))
                p.activation_date   = _parse_project_date(row.get("Activation Date"))
                p.license_eop       = _parse_project_date(row.get("License EOP"))
            try:
                _save_projects(projects, _data_path)
                load_data.clear()
                st.session_state.pop("add_proj_row", None)
                msg = f"Saved! {new_count} new project(s) added." if new_count else "Projects saved successfully!"
                st.session_state["_flash_success"] = msg
                st.session_state["_flash_success_page"] = "🏗️ Projects"
                st.rerun()
            except Exception as e:
                st.error(f"Save failed: {e}")
    else:
        st.dataframe(
            df.style.map(color_status, subset=["Status"]) if "Status" in df.columns and len(df) > 0 else df,
            use_container_width=True,
            height=600,
        )

    # Revenue breakdown
    st.markdown("---")
    st.subheader("Revenue Breakdown by Year")
    cur_year = datetime.datetime.now().year
    rev_rows = []
    for p in filtered:
        if p.installation_year:
            for yr in range(p.installation_year, cur_year + 1):
                rev_rows.append({
                    "Project":   _safe_str(p.project_name),
                    "Year":      int(yr),
                    "Maint. Year": _safe_str(p.get_maintenance_year_label(yr)),
                    "Rate/Cam":  f"€{p.get_rate(yr):,.0f}",
                    "Expected":  f"€{p.get_expected_amount(yr):,.0f}",
                })
    if rev_rows:
        st.dataframe(pd.DataFrame(rev_rows), use_container_width=True, height=400)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: LICENSES
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔐 Licenses":
    st.title("🔐 Licenses")
    if page_flash_success:
        st.success(page_flash_success)

    today = datetime.date.today()
    next_month = today.month % 12 + 1
    next_month_year = today.year if today.month < 12 else today.year + 1

    license_rows = []
    for project in projects:
        license_date = _project_license_date(project)
        license_rows.append({
            "Project": _safe_str(project.project_name),
            "Country": _safe_str(project.country),
            "Cameras": _safe_int(project.num_cams),
            "Status": _safe_str(project.status),
            "License EOP": license_date.strftime("%Y-%m-%d") if license_date else "",
            "License Status": _license_status(project, today),
        })

    next_month_rows = [
        row for row in license_rows
        if row["License Status"] == "Update Next Month"
    ]
    active_license_count = sum(1 for row in license_rows if row["License Status"] == "Active")
    missing_license_count = sum(1 for row in license_rows if row["License Status"] == "Missing")
    expired_license_count = sum(1 for row in license_rows if row["License Status"] == "Expired")
    update_next_month_count = sum(1 for row in license_rows if row["License Status"] == "Update Next Month")

    lc1, lc2, lc3, lc4 = st.columns(4)
    lc1.metric("Active licenses", active_license_count)
    lc2.metric("Need update next month", update_next_month_count)
    lc3.metric("Expired / missing", expired_license_count + missing_license_count)
    lc4.metric("Total tracked projects", len(license_rows))

    with st.expander("Filters", expanded=True):
        lf1, lf2, lf3 = st.columns(3)
        license_country = lf1.selectbox(
            "Country",
            ["All"] + sorted({row["Country"] for row in license_rows if row["Country"]}),
            key="license_country",
        )
        license_status = lf2.selectbox(
            "License Status",
            ["All", "Active", "Update Next Month", "Expired", "Missing"],
            key="license_status",
        )
        license_search = lf3.text_input("Search project", key="license_search")

    filtered_license_rows = [
        row for row in license_rows
        if (license_country == "All" or row["Country"] == license_country)
        and (license_status == "All" or row["License Status"] == license_status)
        and (not license_search.strip() or license_search.lower() in row["Project"].lower())
    ]

    st.subheader(f"Projects Needing Update in {calendar.month_name[next_month]} {next_month_year}")
    if next_month_rows:
        st.dataframe(
            pd.DataFrame(next_month_rows)[["Project", "Country", "Cameras", "License EOP", "Status"]],
            use_container_width=True,
            hide_index=True,
            height=220,
        )
    else:
        st.info(f"No projects currently expire in {calendar.month_name[next_month]} {next_month_year}.")

    if CAN_EDIT and license_rows:
        st.markdown("---")
        st.subheader("Add or Extend License EOP")
        project_names = sorted(row["Project"] for row in license_rows)
        with st.form("license_update_form"):
            lu1, lu2 = st.columns(2)
            selected_project_name = lu1.selectbox("Project", project_names, key="license_project")
            selected_project = next((project for project in projects if project.project_name == selected_project_name), None)
            current_license_date = _project_license_date(selected_project) if selected_project else None
            extend_action = lu2.selectbox(
                "Action",
                ["Set exact date", "Extend by 1 month", "Extend by 12 months"],
                key="license_action",
            )
            base_license_date = current_license_date or today
            new_license_date = st.date_input(
                "New License EOP",
                value=_add_months(base_license_date, 12),
                key="license_new_date",
            )
            if current_license_date:
                st.caption(f"Current License EOP: {current_license_date.strftime('%Y-%m-%d')}")
            else:
                st.caption("Current License EOP: not set")
            submit_license = st.form_submit_button("Save License Update")

        if submit_license and selected_project is not None:
            if extend_action == "Set exact date":
                target_license_date = new_license_date
            elif extend_action == "Extend by 1 month":
                target_license_date = _add_months(max(base_license_date, today), 1)
            else:
                target_license_date = _add_months(max(base_license_date, today), 12)

            selected_project.license_eop = datetime.datetime.combine(target_license_date, datetime.time.min)
            try:
                _save_projects(projects, _data_path)
                load_data.clear()
                st.session_state["_flash_success"] = (
                    f"License EOP updated for {selected_project.project_name}: {target_license_date.strftime('%Y-%m-%d')}"
                )
                st.session_state["_flash_success_page"] = "🔐 Licenses"
                st.rerun()
            except Exception as exc:
                st.error(f"Failed to save license update: {exc}")

    st.markdown("---")
    st.subheader("All Project Licenses")
    license_table_df = pd.DataFrame([
        {
            "Project": row["Project"],
            "Country": row["Country"],
            "Cameras": row["Cameras"],
            "Project Status": row["Status"],
            "License EOP": row["License EOP"],
            "License Status": row["License Status"],
        }
        for row in filtered_license_rows
    ])

    def color_license_status(value):
        key = str(value).strip().lower()
        if key == "active":
            return "color: #27AE60; font-weight: bold"
        if key == "update next month":
            return "color: #F39C12; font-weight: bold"
        if key == "expired":
            return "color: #E74C3C; font-weight: bold"
        if key == "missing":
            return "color: #7F8C8D; font-weight: bold"
        return ""

    st.dataframe(
        license_table_df.style.map(color_license_status, subset=["License Status"]),
        use_container_width=True,
        hide_index=True,
        height=480,
    )


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: INVOICE DETAILS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🧾 Invoice Details":
    st.title("🧾 Invoice Details")
    if page_flash_success:
        st.success(page_flash_success)

    with st.expander("🔍 Filters", expanded=False):
        col1, col2, col3, col4 = st.columns(4)
        years = sorted({inv.year for inv in invoices if inv.year}, reverse=True)
        sel_year    = col1.selectbox("Year",    ["All"] + [str(y) for y in years], key="inv_year")
        sel_paid    = col2.selectbox("Paid Status", ["All", "Paid", "Unpaid", "Cancelled"], key="inv_paid")
        countries   = sorted({p.country for p in projects if p.country})
        sel_country = col3.selectbox("Country", ["All"] + countries, key="inv_country")
        search      = col4.text_input("Search project name", key="inv_search")

        col5, col6, col7, col8 = st.columns(4)
        maint_years = sorted({inv.maintenance_year for inv in invoices if inv.maintenance_year})
        sel_maint     = col5.selectbox("Maint. Year", ["All"] + maint_years, key="inv_maint")
        inv_no_search = col6.text_input("Invoice #", key="inv_no_search")
        amt_min       = col7.number_input("Min Amount (€)", min_value=0, value=0, step=100, key="inv_amt_min")
        amt_max       = col8.number_input("Max Amount (€)", min_value=0, value=0, step=100, key="inv_amt_max",
                                          help="0 = no limit")

    proj_country = {p.project_name.lower(): p.country for p in projects}

    filtered_inv = invoices
    if sel_year != "All":
        filtered_inv = [i for i in filtered_inv if i.year == int(sel_year)]
    if sel_paid == "Paid":
        filtered_inv = [i for i in filtered_inv if i.is_paid()]
    elif sel_paid == "Unpaid":
        filtered_inv = [i for i in filtered_inv if i.is_unpaid()]
    elif sel_paid == "Cancelled":
        filtered_inv = [i for i in filtered_inv if i.is_cancelled()]
    if sel_country != "All":
        filtered_inv = [i for i in filtered_inv
                        if proj_country.get(i.project_name.lower().strip()) == sel_country]
    if search:
        filtered_inv = [i for i in filtered_inv if search.lower() in i.project_name.lower()]
    if sel_maint != "All":
        filtered_inv = [i for i in filtered_inv if _safe_str(i.maintenance_year) == sel_maint]
    if inv_no_search.strip():
        filtered_inv = [i for i in filtered_inv if inv_no_search.strip() in _safe_str(i.invoice_number)]
    if amt_min > 0:
        filtered_inv = [i for i in filtered_inv if i.payment_amount >= amt_min]
    if amt_max > 0:
        filtered_inv = [i for i in filtered_inv if i.payment_amount <= amt_max]

    # Summary row
    total_all_f    = sum(i.payment_amount for i in filtered_inv)
    total_paid_f   = sum(i.payment_amount for i in filtered_inv if i.is_paid())
    total_unpaid_f = sum(i.payment_amount for i in filtered_inv if i.is_unpaid())
    total_cancel_f = sum(i.payment_amount for i in filtered_inv if i.is_cancelled())
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Invoices shown",    len(filtered_inv))
    c2.metric("Total Amount",      f"€{total_all_f:,.0f}")
    c3.metric("Total Paid",        f"€{total_paid_f:,.0f}")
    c4.metric("Total Unpaid",      f"€{total_unpaid_f:,.0f}")
    c5.metric("Total Cancelled",   f"€{total_cancel_f:,.0f}")

    st.caption(f"Showing {len(filtered_inv)} of {len(invoices)} invoices — use filters above to narrow results")

    _invoice_columns = [
        "Invoice #",
        "Project",
        "Maint. Year",
        "Amount (€)",
        "Cameras",
        "Payment Date",
        "Paid",
        "Year",
    ]

    invoice_index_map = {id(inv): idx for idx, inv in enumerate(invoices)}

    df_inv = pd.DataFrame(
        [
            {
                "_invoice_index": invoice_index_map[id(i)],
                "Invoice #": _safe_int(i.invoice_number) or None,
                "Project": _safe_str(i.project_name),
                "Maint. Year": _safe_str(i.maintenance_year),
                "Amount (€)": _safe_float(i.payment_amount),
                "Cameras": _safe_int(i.cameras_number),
                "Payment Date": i.payment_date.date() if i.payment_date else None,
                "Paid": _safe_str(i.paid),
                "Year": _safe_str(_safe_int(i.year) or ""),
            }
            for i in filtered_inv
        ],
        columns=["_invoice_index", *_invoice_columns],
    ).sort_values(["Invoice #", "Project"], ignore_index=True)

    def color_paid(val):
        v = str(val).strip().lower()
        if v == "yes":     return "color: #27AE60; font-weight: bold"
        if v == "no":      return "color: #E74C3C"
        if v == "cancelled": return "color: #F39C12"
        return ""

    if CAN_EDIT:
        st.info("✏️ Admin mode: you can edit cells directly. Click **Save Changes** when done.")

        def _parse_invoice_date(value):
            if value in (None, ""):
                return None
            if isinstance(value, datetime.datetime):
                return value
            if isinstance(value, datetime.date):
                return datetime.datetime.combine(value, datetime.time.min)
            try:
                return datetime.datetime.strptime(str(value), "%Y-%m-%d")
            except Exception:
                return None

        def _invoice_row_has_values(row) -> bool:
            return any([
                row.get("Invoice #") not in (None, ""),
                _safe_str(row.get("Maint. Year", "")).strip(),
                _safe_float(row.get("Amount (€)", 0)) != 0,
                _safe_int(row.get("Cameras", 0)) != 0,
                row.get("Payment Date") not in (None, ""),
                _safe_str(row.get("Paid", "")).strip() not in ("", "No"),
                _safe_str(row.get("Year", "")).strip(),
            ])

        blank_project_invoices = [inv for inv in invoices if not _safe_str(inv.project_name).strip()]
        if blank_project_invoices:
            st.warning(
                f"Found {len(blank_project_invoices)} invoice row(s) with an empty project name. "
                "Remove them before saving other invoice changes."
            )
            if st.button("🧹 Remove Blank Project Invoice Rows", key="remove_blank_project_invoices"):
                invoices[:] = [inv for inv in invoices if _safe_str(inv.project_name).strip()]
                try:
                    _save_invoices(invoices, _data_path)
                    load_data.clear()
                    st.session_state["_flash_success"] = "Blank-project invoice rows were removed."
                    st.session_state["_flash_success_page"] = "🧾 Invoice Details"
                    st.rerun()
                except Exception as exc:
                    st.error(f"Failed to remove blank-project invoice rows: {exc}")

        with st.expander("📥 Import Invoice XLSX", expanded=False):
            if _is_excel_source(_data_path):
                st.warning("XLSX import is only available when the app is connected to Supabase.")
            else:
                uploaded_invoice_xlsx = st.file_uploader(
                    "Upload invoice XLSX",
                    type=["xlsx"],
                    key="invoice_xlsx_upload",
                    help="Upload a generated invoice workbook and replace that invoice number in Supabase.",
                )
                if uploaded_invoice_xlsx is not None:
                    try:
                        import_meta, import_rows = _parse_uploaded_invoice_xlsx(uploaded_invoice_xlsx.getvalue())
                        inferred_invoice_number = import_meta["invoice_number"] or 0
                        target_invoice_number = st.number_input(
                            "Invoice number to replace",
                            min_value=1,
                            step=1,
                            value=max(1, inferred_invoice_number),
                            key="invoice_import_target_number",
                        )
                        st.caption(
                            f"Parsed {import_meta['row_count']} rows from '{import_meta['title'] or uploaded_invoice_xlsx.name}' "
                            f"for year {import_meta['year'] or 'unknown'} with total €{import_meta['total_amount']:,.0f}."
                        )
                        preview_df = pd.DataFrame([
                            {
                                "Project": row["project_name"],
                                "Maint. Year": row["maintenance_year"],
                                "Amount (€)": row["payment_amount"],
                                "Cameras": row["cameras_number"],
                                "Year": row["year"],
                            }
                            for row in import_rows
                        ])
                        st.dataframe(preview_df, use_container_width=True, height=250)

                        if preview_df.empty:
                            st.warning("No invoice rows were found in the uploaded XLSX.")
                        else:
                            blank_import_rows = [row for row in import_rows if not _safe_str(row.get("project_name", "")).strip()]
                            project_counts = preview_df["Project"].value_counts()
                            duplicate_projects = project_counts[project_counts > 1]
                            if blank_import_rows:
                                st.error("The uploaded file contains invoice rows with an empty project name. Import was blocked.")
                            elif not duplicate_projects.empty:
                                st.error(
                                    "The uploaded file contains duplicate project rows: "
                                    + ", ".join(duplicate_projects.index.tolist())
                                )
                            elif st.button("Replace Invoice From XLSX", key="replace_invoice_from_xlsx"):
                                replace_invoice_rows(target_invoice_number, import_rows)
                                load_data.clear()
                                st.session_state["_flash_success"] = (
                                    f"Invoice {target_invoice_number} replaced from uploaded XLSX with {len(import_rows)} row(s)."
                                )
                                st.session_state["_flash_success_page"] = "🧾 Invoice Details"
                                st.rerun()
                    except Exception as exc:
                        st.error(f"Failed to parse uploaded invoice XLSX: {exc}")

        if st.button("➕ Add New Invoice", key="btn_add_inv"):
            st.session_state["add_inv_row"] = st.session_state.get("add_inv_row", 0) + 1

        invoice_project_options = [""] + sorted(
            {
                _safe_str(p.project_name).strip()
                for p in projects
                if _safe_str(p.project_name).strip()
            }
            | {
                _safe_str(inv.project_name).strip()
                for inv in invoices
                if _safe_str(inv.project_name).strip()
            }
        )
        invoice_paid_options = ["No", "Yes", "cancelled"]
        invoice_year_options = [""] + [str(year) for year in range(datetime.date.today().year + 1, 2012, -1)]
        invoice_maint_options = []
        for label in maint_years + [f"Y{i}" for i in range(1, 11)] + ["Paid Trial-0.5Y"]:
            if label and label not in invoice_maint_options:
                invoice_maint_options.append(label)

        _empty_inv = {"_invoice_index": None, "Invoice #": None, "Project": "", "Maint. Year": "Y1",
                      "Amount (€)": 0.0, "Cameras": 0,
                      "Payment Date": None, "Paid": "No", "Year": str(datetime.date.today().year)}
        n_new_inv = st.session_state.get("add_inv_row", 0)
        if n_new_inv:
            empty_rows = pd.DataFrame([_empty_inv] * n_new_inv)
            df_inv_edit = pd.concat([empty_rows, df_inv.reset_index(drop=True)], ignore_index=True)
        else:
            df_inv_edit = df_inv.reset_index(drop=True)

        edited_inv = st.data_editor(
            df_inv_edit,
            use_container_width=True,
            height=550,
            num_rows="dynamic",
            column_config={
                "_invoice_index": None,
                "Invoice #": st.column_config.NumberColumn(
                    "Invoice #",
                    min_value=0,
                    step=1,
                    format="%d",
                ),
                "Project": st.column_config.SelectboxColumn(
                    "Project",
                    options=invoice_project_options,
                ),
                "Maint. Year": st.column_config.SelectboxColumn(
                    "Maint. Year",
                    options=invoice_maint_options,
                ),
                "Payment Date": st.column_config.DateColumn(
                    "Payment Date",
                    format="YYYY-MM-DD",
                ),
                "Paid": st.column_config.SelectboxColumn(
                    "Paid",
                    options=invoice_paid_options,
                ),
                "Year": st.column_config.SelectboxColumn(
                    "Year",
                    options=invoice_year_options,
                ),
            },
            key=f"inv_editor_{n_new_inv}",
        )
        if st.button("💾 Save Changes", key="save_invoices"):
            invalid_editor_rows = []
            for row_idx, row in edited_inv.iterrows():
                project = _safe_str(row.get("Project", "")).strip()
                if not project and _invoice_row_has_values(row):
                    invalid_editor_rows.append(row_idx + 1)

            if invalid_editor_rows:
                st.error(
                    "Cannot save invoice rows with invoice data but no project name. "
                    f"Fix or remove row(s): {', '.join(str(idx) for idx in invalid_editor_rows[:10])}"
                )
                st.stop()

            from models.invoice import Invoice as InvoiceModel
            inv_map = {(i.invoice_number, i.project_name.strip().lower()): i for i in invoices if i.invoice_number}
            # Secondary lookup for invoices without invoice numbers — match by (project, maint_year, year)
            no_inv_map = {
                (i.project_name.strip().lower(), str(i.maintenance_year).strip(), str(i.year) if i.year else ""): i
                for i in invoices if not i.invoice_number
            }
            new_count = 0
            for _, row in edited_inv.iterrows():
                project = _safe_str(row.get("Project", "")).strip()
                if not project:
                    continue
                raw_invoice_index = row.get("_invoice_index", "")
                existing_invoice_index = _safe_int(raw_invoice_index, default=-1)
                inv = None
                if existing_invoice_index >= 0 and existing_invoice_index < len(invoices):
                    inv = invoices[existing_invoice_index]
                inv_no_str = _safe_str(row.get("Invoice #", "")).strip()
                try:
                    inv_no = float(inv_no_str) if inv_no_str else None
                except Exception:
                    inv_no = None
                maint_year = _safe_str(row.get("Maint. Year", "")).strip()
                inv_year = str(_safe_int(row.get("Year")) or "")
                if inv is None and inv_no:
                    inv = inv_map.get((inv_no, project.lower()))
                elif inv is None:
                    inv = no_inv_map.get((project.lower(), maint_year, inv_year))
                if inv is None:
                    # Truly new invoice row
                    inv = InvoiceModel(
                        invoice_number=inv_no,
                        project_name=project,
                        maintenance_year=maint_year,
                        payment_amount=_safe_float(row.get("Amount (€)", 0)),
                        cameras_number=_safe_int(row.get("Cameras", 0)) or None,
                        payment_date=None,
                        paid=_safe_str(row.get("Paid", "No")),
                        year=_safe_int(row.get("Year")) or None,
                    )
                    invoices.append(inv)
                    if inv_no:
                        inv_map[(inv_no, project.lower())] = inv
                    else:
                        no_inv_map[(project.lower(), maint_year, inv_year)] = inv
                    new_count += 1
                else:
                    inv.invoice_number = inv_no
                    inv.project_name   = project
                    inv.maintenance_year = _safe_str(row.get("Maint. Year", ""))
                    inv.paid           = _safe_str(row.get("Paid", ""))
                    inv.payment_amount = _safe_float(row.get("Amount (€)", 0))
                    inv.cameras_number = _safe_int(row.get("Cameras", 0)) or None
                    inv.year           = _safe_int(row.get("Year")) or None
                inv.payment_date = _parse_invoice_date(row.get("Payment Date"))
            try:
                _save_invoices(invoices, _data_path)
                load_data.clear()
                st.session_state.pop("add_inv_row", None)
                msg = f"Saved! {new_count} new invoice(s) added." if new_count else "Invoices saved successfully!"
                st.session_state["_flash_success"] = msg
                st.session_state["_flash_success_page"] = "🧾 Invoice Details"
                st.rerun()
            except Exception as e:
                st.error(f"Save failed: {e}")
    else:
        st.dataframe(
            df_inv.style.map(color_paid, subset=["Paid"]),
            use_container_width=True,
            height=550,
        )


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DEBT REPORT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "💸 Debt Report":
    st.title("💸 Debt Report")

    debt_type_options = ["All", "New Installation (Y1)", "Paid Trials", "Maintenance (Y2+)"]
    pending_debt_type = st.session_state.pop("_pending_dr_debt_type", None)
    if pending_debt_type in debt_type_options:
        st.session_state["dr_debt_type"] = pending_debt_type
        st.session_state["dr_debt_type_widget"] = pending_debt_type
    active_debt_type = st.session_state.get("dr_debt_type", "All")
    if active_debt_type not in debt_type_options:
        active_debt_type = "All"
        st.session_state["dr_debt_type"] = active_debt_type
    widget_debt_type = st.session_state.get("dr_debt_type_widget", active_debt_type)
    if widget_debt_type not in debt_type_options:
        widget_debt_type = active_debt_type
        st.session_state["dr_debt_type_widget"] = widget_debt_type

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander("🔍 Filters", expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        debt_years    = sorted({inv.year for inv in invoices if inv.year}, reverse=True)
        dsel_year     = fc1.selectbox("Year", ["All"] + [str(y) for y in debt_years], key="dr_year")
        debt_countries = sorted({p.country for p in projects if p.country})
        dsel_country  = fc2.selectbox("Country", ["All"] + debt_countries, key="dr_country")
        dsel_search   = fc3.text_input("Search project", key="dr_search")
        dsel_debt_type = fc4.selectbox(
            "Debt Type",
            debt_type_options,
            index=debt_type_options.index(widget_debt_type),
            key="dr_debt_type_widget",
        )
        st.session_state["dr_debt_type"] = dsel_debt_type

    # Only unpaid invoices
    debt_inv = [i for i in invoices if i.is_unpaid()]
    if dsel_year != "All":
        debt_inv = [i for i in debt_inv if i.year == int(dsel_year)]

    # Country lookup — normalize names aggressively, fallback to debt_summaries
    import re as _re
    import unicodedata as _ud
    def _norm(s):
        s = str(s or "").strip()
        s = _re.sub(r'\s*\([^)]*\)', '', s).strip()          # strip "(coplementary)" etc.
        s = _ud.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')  # strip accents
        return _re.sub(r'\s+', ' ', s).strip().lower()
    proj_country_map = {_norm(p.project_name): p.country for p in projects}
    ds_country_map   = {_norm(ds.project_name): ds.country for ds in debt_summaries}
    def _get_country(name):
        k = _norm(name)
        result = proj_country_map.get(k) or ds_country_map.get(k)
        if result:
            return result
        # Partial match: invoice name starts with a known project name or vice versa
        for proj_k, country in proj_country_map.items():
            if proj_k and k and (k.startswith(proj_k) or proj_k.startswith(k)):
                return country
        return ""

    if dsel_country != "All":
        debt_inv = [i for i in debt_inv if _get_country(i.project_name) == dsel_country]
    if dsel_search.strip():
        debt_inv = [i for i in debt_inv if dsel_search.lower() in i.project_name.lower()]
    if dsel_debt_type == "New Installation (Y1)":
        debt_inv = [i for i in debt_inv if _is_new_installation_category(i)]
    elif dsel_debt_type == "Paid Trials":
        debt_inv = [i for i in debt_inv if _is_paid_trial_category(i)]
    elif dsel_debt_type == "Maintenance (Y2+)":
        debt_inv = [i for i in debt_inv if _is_maintenance_category(i)]

    # ── Summary metrics ───────────────────────────────────────────────────────
    total_debt_amt  = sum(i.payment_amount for i in debt_inv)
    proj_with_debt  = len({i.project_name for i in debt_inv})
    all_unpaid = [i for i in invoices if i.is_unpaid()]
    if dsel_year != "All":
        all_unpaid = [i for i in all_unpaid if i.year == int(dsel_year)]
    if dsel_country != "All":
        all_unpaid = [i for i in all_unpaid if _get_country(i.project_name) == dsel_country]
    if dsel_search.strip():
        all_unpaid = [i for i in all_unpaid if dsel_search.lower() in i.project_name.lower()]

    y1_total_amt = sum(i.payment_amount for i in all_unpaid if _is_new_installation_category(i))
    paid_trial_total_amt = sum(i.payment_amount for i in all_unpaid if _is_paid_trial_category(i))
    y2_total_amt = sum(i.payment_amount for i in all_unpaid if _is_maintenance_category(i))

    mc1, mc2, mc3, mc4, mc5, mc6 = st.columns(6)
    mc1.metric("Unpaid Invoices",  len(debt_inv))
    mc2.metric("Projects with Debt", proj_with_debt)
    mc3.metric("Total Debt",       f"€{total_debt_amt:,.0f}")
    mc4.metric("Y1 Debt",          f"€{y1_total_amt:,.0f}")
    mc5.metric("Paid Trials",      f"€{paid_trial_total_amt:,.0f}")
    mc6.metric("Y2+ Debt",         f"€{y2_total_amt:,.0f}")

    if mc4.button("Show Y1 Invoice List", key="dr_show_y1", use_container_width=True):
        st.session_state["_pending_dr_debt_type"] = "New Installation (Y1)"
        st.rerun()
    if mc5.button("Show Paid Trials", key="dr_show_trials", use_container_width=True):
        st.session_state["_pending_dr_debt_type"] = "Paid Trials"
        st.rerun()
    if mc6.button("Show Y2+ Invoice List", key="dr_show_y2", use_container_width=True):
        st.session_state["_pending_dr_debt_type"] = "Maintenance (Y2+)"
        st.rerun()

    if dsel_debt_type != "All":
        if st.button("Clear Debt Type Filter", key="dr_clear_debt_type"):
            st.session_state["_pending_dr_debt_type"] = "All"
            st.rerun()

    if dsel_debt_type == "New Installation (Y1)":
        st.caption("Showing only first-year debt (Y1).")
    elif dsel_debt_type == "Paid Trials":
        st.caption("Showing only paid-trial debt.")
    elif dsel_debt_type == "Maintenance (Y2+)":
        st.caption("Showing only maintenance debt (Y2+).")
    else:
        st.caption("Use the Debt Type filter to switch between all debt, Y1, paid trials, and Y2+.")

    st.markdown("---")

    # ── Detailed table: one row per unpaid invoice ────────────────────────────
    st.subheader("Unpaid Invoices Detail")
    detail_rows = [{
        "Invoice #":      str(int(i.invoice_number)) if i.invoice_number else "—",
        "Project Name":   _safe_str(i.project_name),
        "Country":        _safe_str(_get_country(i.project_name)),
        "Maint. Year":    _safe_str(i.maintenance_year),
        "Amount (€)":     _safe_float(i.payment_amount),
        "Year":           _safe_str(_safe_int(i.year) or ""),
        "Payment Date":   i.payment_date.strftime("%Y-%m-%d") if i.payment_date else "",
    } for i in sorted(debt_inv, key=lambda x: (x.project_name, x.year or 0))]

    if detail_rows:
        detail_df = pd.DataFrame(detail_rows)
        st.dataframe(detail_df, use_container_width=True, hide_index=True, height=350)

        # Download detail as CSV
        csv_detail = detail_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Detail CSV", csv_detail,
                           file_name="debt_detail.csv", mime="text/csv")
    else:
        st.success("No unpaid invoices match the current filters.")

    st.markdown("---")

    # ── Summary table: grouped by project ────────────────────────────────────
    st.subheader("Debt Summary by Project")
    from collections import defaultdict
    proj_debt: dict = defaultdict(lambda: {"inv_nos": set(), "total": 0.0, "country": ""})
    for i in debt_inv:
        key = i.project_name
        if i.invoice_number:
            proj_debt[key]["inv_nos"].add(str(int(i.invoice_number)))
        proj_debt[key]["total"]   += i.payment_amount
        proj_debt[key]["country"]  = _get_country(i.project_name)

    summary_rows = [{
        "Project Name":    name,
        "Country":         d["country"],
        "Invoice Numbers": ", ".join(sorted(d["inv_nos"])),
        "Total Debt (€)":  int(round(d["total"])),
    } for name, d in sorted(proj_debt.items(), key=lambda x: -x[1]["total"])]

    if summary_rows:
        summary_df = pd.DataFrame(summary_rows)

        def color_debt_amt(val):
            try:
                return "color: #E74C3C; font-weight: bold" if float(val) > 0 else ""
            except Exception:
                return ""

        st.dataframe(
            summary_df.style.map(color_debt_amt, subset=["Total Debt (€)"]),
            use_container_width=True, hide_index=True, height=400,
        )

        # Download summary as CSV
        csv_summary = summary_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Summary CSV", csv_summary,
                           file_name="debt_summary.csv", mime="text/csv")

        # Download summary as PDF (split Y1 vs Y2+)
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import cm
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.enums import TA_CENTER
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib import colors as rl_colors
            from io import BytesIO

            # Split unpaid invoices into Y1, paid trials, and Y2+
            y1_inv  = [i for i in debt_inv if _is_new_installation_category(i)]
            trial_inv = [i for i in debt_inv if _is_paid_trial_category(i)]
            y2_inv  = [i for i in debt_inv if _is_maintenance_category(i)]

            def _proj_debt_rows(inv_list):
                pd_: dict = defaultdict(lambda: {"inv_nos": set(), "total": 0.0, "country": ""})
                for i in inv_list:
                    if i.invoice_number:
                        pd_[i.project_name]["inv_nos"].add(str(int(i.invoice_number)))
                    pd_[i.project_name]["total"]   += i.payment_amount
                    pd_[i.project_name]["country"]  = _get_country(i.project_name)
                return sorted(pd_.items(), key=lambda x: -x[1]["total"])

            def _make_section_table(rows, rl_colors):
                hdr_color = rl_colors.HexColor("#1B3A6B")
                tbl_data = [["Project Name", "Country", "Invoice #s", "Debt (€)"]]
                for name, d in rows:
                    tbl_data.append([
                        name, d["country"],
                        ", ".join(sorted(d["inv_nos"])),
                        f"€{d['total']:,.0f}",
                    ])
                t = Table(tbl_data, repeatRows=1,
                          colWidths=[7.5*cm, 2*cm, 5*cm, 3*cm])
                t.setStyle(TableStyle([
                    ("BACKGROUND",     (0, 0), (-1, 0), hdr_color),
                    ("TEXTCOLOR",      (0, 0), (-1, 0), rl_colors.white),
                    ("FONTNAME",       (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE",       (0, 0), (-1, -1), 8),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                     [rl_colors.white, rl_colors.HexColor("#EBF5FB")]),
                    ("GRID",           (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#BDC3C7")),
                    ("BOTTOMPADDING",  (0, 0), (-1, -1), 4),
                    ("TOPPADDING",     (0, 0), (-1, -1), 4),
                    ("ALIGN",          (3, 0), (3, -1), "RIGHT"),
                ]))
                return t

            pdf_buf = BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=A4,
                                    leftMargin=1.5*cm, rightMargin=1.5*cm,
                                    topMargin=1.5*cm, bottomMargin=1.5*cm)
            styles = getSampleStyleSheet()
            sec_style = ParagraphStyle("sec", parent=styles["Heading2"],
                                       textColor=rl_colors.HexColor("#1B3A6B"))
            elems = []
            elems.append(Paragraph("<b>CaddyCheck CRM — Debt Report</b>", styles["Title"]))
            elems.append(Paragraph(
                f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d')}  |  "
                f"Total Debt: €{total_debt_amt:,.0f}  |  Projects: {proj_with_debt}",
                styles["Normal"]))
            elems.append(Spacer(1, 0.4*cm))

            # Section A: Y1
            y1_total = sum(i.payment_amount for i in y1_inv)
            elems.append(Paragraph(
                f"A. New Installation Debt (Y1)  —  €{y1_total:,.0f}", sec_style))
            elems.append(Spacer(1, 0.2*cm))
            y1_rows = _proj_debt_rows(y1_inv)
            if y1_rows:
                elems.append(_make_section_table(y1_rows, rl_colors))
            else:
                elems.append(Paragraph("No Y1 unpaid invoices.", styles["Normal"]))
            elems.append(Spacer(1, 0.6*cm))

            # Section B: Paid Trials
            trial_total = sum(i.payment_amount for i in trial_inv)
            elems.append(Paragraph(
                f"B. Paid Trials  —  €{trial_total:,.0f}", sec_style))
            elems.append(Spacer(1, 0.2*cm))
            trial_rows = _proj_debt_rows(trial_inv)
            if trial_rows:
                elems.append(_make_section_table(trial_rows, rl_colors))
            else:
                elems.append(Paragraph("No unpaid paid-trial invoices.", styles["Normal"]))
            elems.append(Spacer(1, 0.6*cm))

            # Section C: Y2+
            y2_total = sum(i.payment_amount for i in y2_inv)
            elems.append(Paragraph(
                f"C. Maintenance Debt (Y2+)  —  €{y2_total:,.0f}", sec_style))
            elems.append(Spacer(1, 0.2*cm))
            y2_rows = _proj_debt_rows(y2_inv)
            if y2_rows:
                elems.append(_make_section_table(y2_rows, rl_colors))
            else:
                elems.append(Paragraph("No Y2+ unpaid invoices.", styles["Normal"]))

            doc.build(elems)
            st.download_button("⬇️ Download Debt PDF", pdf_buf.getvalue(),
                               file_name="debt_report.pdf", mime="application/pdf")
        except Exception as e:
            st.warning(f"PDF export unavailable: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: MONTHLY INVOICE
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📅 Monthly Invoice":
    st.title("📅 Monthly Invoice")

    next_inv_no = _get_next_invoice_number(invoices, _data_path)

    col1, col2, col3 = st.columns([2, 1, 2])
    with col1:
        sel_month = st.selectbox("Month", MONTH_ORDER, index=datetime.date.today().month - 1)
    with col2:
        sel_year = st.number_input("Year", min_value=2015, max_value=2035,
                                   value=datetime.date.today().year, step=1)
    with col3:
        invoice_number = st.number_input("Invoice Number",
                                         min_value=1, value=next_inv_no, step=1)

    month_projects = get_projects_for_month(projects, sel_month)
    st.markdown(f"**{len(month_projects)} project(s)** billed in **{sel_month}**")

    if not month_projects:
        st.warning("No projects found for the selected month.")
    else:
        preview_rows = get_invoice_preview_data(month_projects, sel_month, int(sel_year))
        preview_total_amount = sum(
            float(r["line_total"]) for r in preview_rows if isinstance(r.get("line_total"), (int, float))
        )
        preview_df = pd.DataFrame([{
            "Project": _safe_str(r["project_name"]),
            "# Cams": _safe_str(r["num_cams"]),
            "Maint. Year": _safe_str(r["maintenance_year"]),
            "Rate (€)": f"€{r['rate']:,.0f}" if isinstance(r["rate"], (int, float)) else _safe_str(r["rate"]),
            "Line Total (€)": f"€{r['line_total']:,.0f}" if isinstance(r["line_total"], (int, float)) else "",
        } for r in preview_rows])
        st.subheader("Invoice Preview")
        st.dataframe(preview_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("Monthly Invoice Status")
        st.caption(
            "Combined monthly invoices are inferred when the same invoice number appears on multiple projects, which matches the shared invoices introduced from Aug 2025 onward."
        )

        monthly_invoice_summaries = group_monthly_invoices(invoices)
        if monthly_invoice_summaries:
            monthly_years = sorted({summary.year for summary in monthly_invoice_summaries if summary.year}, reverse=True)
            monthly_statuses = sorted({summary.status for summary in monthly_invoice_summaries})

            mi1, mi2 = st.columns(2)
            monthly_year_filter = mi1.selectbox(
                "Monthly Invoice Year",
                ["All"] + [str(year) for year in monthly_years],
                key="monthly_invoice_status_year",
            )
            monthly_status_filter = mi2.selectbox(
                "Monthly Invoice Status",
                ["All"] + monthly_statuses,
                key="monthly_invoice_status_filter",
            )

            filtered_monthly_summaries = [
                summary for summary in monthly_invoice_summaries
                if (monthly_year_filter == "All" or str(summary.year or "") == monthly_year_filter)
                and (monthly_status_filter == "All" or summary.status == monthly_status_filter)
            ]

            mm1, mm2, mm3, mm4 = st.columns(4)
            mm1.metric("Monthly invoices", len(filtered_monthly_summaries))
            mm2.metric(
                "Grouped amount",
                f"€{sum(summary.total_amount for summary in filtered_monthly_summaries):,.0f}",
            )
            mm3.metric(
                "Paid monthly invoices",
                sum(1 for summary in filtered_monthly_summaries if summary.status == "Paid"),
            )
            mm4.metric(
                "Open / mixed",
                sum(1 for summary in filtered_monthly_summaries if summary.status != "Paid"),
            )

            monthly_status_df = pd.DataFrame([
                {
                    "Invoice #": str(summary.invoice_number),
                    "Year": _safe_str(summary.year or ""),
                    "Projects": summary.project_count,
                    "Total (€)": f"{summary.total_amount:,.0f}",
                    "Status": summary.status,
                    "Last Payment": summary.last_payment_date.strftime("%Y-%m-%d") if summary.last_payment_date else "",
                    "Included Projects": ", ".join(summary.project_names),
                }
                for summary in filtered_monthly_summaries
            ])

            def color_monthly_status(val):
                value = str(val).strip().lower()
                if value == "paid":
                    return "color: #27AE60; font-weight: bold"
                if value in {"partial", "mixed", "paid / cancelled"}:
                    return "color: #F39C12; font-weight: bold"
                if value in {"unpaid", "unpaid / cancelled"}:
                    return "color: #E74C3C; font-weight: bold"
                if value == "cancelled":
                    return "color: #7F8C8D; font-weight: bold"
                return ""

            st.dataframe(
                monthly_status_df.style.map(color_monthly_status, subset=["Status"]),
                use_container_width=True,
                hide_index=True,
                height=260,
                column_config={
                    "Included Projects": st.column_config.TextColumn(width="large"),
                },
            )

            with st.expander("Show Full Included Projects", expanded=False):
                for summary in filtered_monthly_summaries:
                    st.markdown(
                        f"**Invoice #{summary.invoice_number}**  |  {summary.project_count} project(s)  |  {summary.status}"
                    )
                    st.write(", ".join(summary.project_names))
        else:
            st.info("No combined monthly invoices were found in the ledger yet.")

        if CAN_EDIT:
            st.markdown("---")
            st.subheader("Generate Invoice")

            from config.settings import get_data_paths
            from services.pdf_service import generate_invoice_pdf
            paths = get_data_paths()
            inv_no = int(invoice_number)
            month_abbr = sel_month[:3]
            pdf_filename = f"CC_M-inv_{inv_no}_{month_abbr}_{int(sel_year)}.pdf"
            xlsx_filename = f"CC_M-inv_{inv_no}_{month_abbr}_{int(sel_year)}.xlsx"

            col_pdf, col_xlsx = st.columns(2)

            with col_pdf:
                try:
                    pdf_bytes = generate_invoice_pdf(
                        projects=month_projects,
                        month_name=sel_month,
                        year=int(sel_year),
                        invoice_number=inv_no,
                    )
                    st.download_button(
                        label="Download PDF Invoice",
                        data=pdf_bytes,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        type="primary",
                    )
                except Exception as e:
                    st.error(f"PDF generation failed: {e}")

            with col_xlsx:
                template_ok = paths["invoice_template"].exists()
                if not template_ok:
                    st.warning("Excel template not found.")
                else:
                    if st.button("Generate Excel Invoice"):
                        import tempfile
                        with tempfile.TemporaryDirectory() as tmp_dir:
                            try:
                                out_path = generate_monthly_invoice(
                                    projects=month_projects,
                                    month_name=sel_month,
                                    year=int(sel_year),
                                    invoice_number=inv_no,
                                    output_dir=Path(tmp_dir),
                                    template_path=paths["invoice_template"],
                                )
                                with open(out_path, "rb") as f:
                                    file_bytes = f.read()
                                st.download_button(
                                    label=f"Download {out_path.name}",
                                    data=file_bytes,
                                    file_name=out_path.name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                )
                            except Exception as e:
                                st.error(f"Excel generation failed: {e}")

            # ── Save to Ledger ─────────────────────────────────────────────
            st.markdown("---")
            st.subheader("Save to Invoice Ledger")
            st.caption(f"Appends invoice #{inv_no} rows for all {len(month_projects)} project(s) to Invoice Details.")
            if st.button("💾 Save Invoice to Ledger", type="primary"):
                try:
                    n = _append_invoice_rows(invoice_number=inv_no, projects=month_projects, year=int(sel_year), source_name=_data_path)
                    if n == 0:
                        st.info(f"Invoice #{inv_no} already fully recorded — no new rows added.")
                    else:
                        st.success(f"Added {n} row(s) for invoice #{inv_no} to Invoice Details.")
                        st.cache_data.clear()
                except Exception as e:
                    st.error(f"Failed to save: {e}")

            # ── Email section ──────────────────────────────────────────────
            st.markdown("---")
            st.subheader("Send Invoice by Email")
            st.caption("Sending the PDF does not add invoice rows to the ledger unless you enable the option below.")
            email_cfg = get_email_config()

            with st.form("email_form"):
                to_addrs = st.text_input("To (comma-separated)",
                                         ", ".join(email_cfg.get("default_recipients", [])))
                cc_addrs = st.text_input("CC (comma-separated)",
                                         ", ".join(email_cfg.get("default_cc", [])))
                subject_tpl = email_cfg.get("default_subject_template", "Monthly Invoice - {month} {year}")
                subject = st.text_input("Subject", subject_tpl.format(month=sel_month, year=sel_year))
                body_tpl = email_cfg.get("default_body_template", "")
                body = st.text_area("Body", body_tpl.format(month=sel_month, year=sel_year), height=150)
                save_to_ledger_on_send = st.checkbox("Also save invoice to ledger before sending", value=True)
                send_btn = st.form_submit_button("Send Email")

            if send_btn:
                if not email_cfg.get("smtp_username"):
                    st.error("SMTP not configured. Go to ⚙️ Settings to set up email.")
                else:
                    from services.email_service import send_invoice_email
                    from services.invoice_service import generate_monthly_invoice_pdf
                    import tempfile
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        try:
                            ledger_rows_added = None
                            if save_to_ledger_on_send:
                                ledger_rows_added = _append_invoice_rows(
                                    invoice_number=inv_no,
                                    projects=month_projects,
                                    year=int(sel_year),
                                    source_name=_data_path,
                                )
                            out_path = generate_monthly_invoice_pdf(
                                projects=month_projects,
                                month_name=sel_month,
                                year=int(sel_year),
                                invoice_number=inv_no,
                                output_dir=Path(tmp_dir),
                            )
                            recipients = [r.strip() for r in to_addrs.split(",") if r.strip()]
                            cc_list = [c.strip() for c in cc_addrs.split(",") if c.strip()]
                            send_invoice_email(
                                attachment_path=out_path,
                                recipients=recipients,
                                cc=cc_list,
                                subject=subject,
                                body=body,
                                config=email_cfg,
                            )
                            append_sent_invoice_log({
                                "sent_at": datetime.datetime.utcnow().isoformat(),
                                "invoice_number": inv_no,
                                "month": sel_month,
                                "year": int(sel_year),
                                "pdf_filename": out_path.name,
                                "recipients": recipients,
                                "cc": cc_list,
                                "subject": subject,
                                "project_count": len(month_projects),
                                "total_amount": preview_total_amount,
                                "saved_to_ledger": bool(save_to_ledger_on_send),
                                "ledger_rows_added": ledger_rows_added,
                            })
                            if save_to_ledger_on_send:
                                st.cache_data.clear()
                            if save_to_ledger_on_send and ledger_rows_added is not None:
                                st.success(
                                    f"Email sent with PDF attachment: {out_path.name}. Ledger rows added: {ledger_rows_added}."
                                )
                            else:
                                st.success(f"Email sent with PDF attachment: {out_path.name}")
                        except Exception as e:
                            st.error(f"Email failed: {e}")

            st.markdown("---")
            st.subheader("Sent PDF Invoices")
            sent_invoice_rows = load_sent_invoices_log()
            if sent_invoice_rows:
                sent_invoice_df = pd.DataFrame([
                    {
                        "Sent At": _safe_str(row.get("sent_at", "")).replace("T", " ")[:19],
                        "Invoice #": _safe_int(row.get("invoice_number"), default=0),
                        "Month": _safe_str(row.get("month", "")),
                        "Year": _safe_int(row.get("year"), default=0),
                        "PDF": _safe_str(row.get("pdf_filename", "")),
                        "Projects": _safe_int(row.get("project_count"), default=0),
                        "Total (€)": f"€{_safe_float(row.get('total_amount', 0.0)):,.0f}",
                        "To": ", ".join(row.get("recipients", [])),
                        "Saved To Ledger": "Yes" if row.get("saved_to_ledger") else "No",
                    }
                    for row in reversed(sent_invoice_rows)
                ])
                st.dataframe(sent_invoice_df, use_container_width=True, hide_index=True, height=260)
            else:
                st.info("No invoice PDFs have been logged as sent yet.")
        else:
            st.info("Invoice generation requires Admin access.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⚙️ Settings":
    st.title("⚙️ Settings")

    if not CAN_EDIT:
        st.info("Settings can only be changed by Admin users.")
    else:
        email_cfg = get_email_config()

        st.subheader("Email / SMTP Configuration")
        with st.form("smtp_form"):
            col1, col2 = st.columns(2)
            with col1:
                smtp_host = st.text_input("SMTP Host", email_cfg.get("smtp_host", "smtp.gmail.com"))
                smtp_user = st.text_input("Username / Email", email_cfg.get("smtp_username", ""))
                sender_name = st.text_input("Sender Name", email_cfg.get("sender_name", "CaddyCheck CRM"))
            with col2:
                smtp_port = st.number_input("SMTP Port", min_value=1, max_value=65535,
                                            value=int(email_cfg.get("smtp_port", 587)))
                smtp_pass = st.text_input("Password", email_cfg.get("smtp_password", ""),
                                          type="password")
                sender_email = st.text_input("Sender Email", email_cfg.get("sender_email", ""))
            use_tls = st.checkbox("Use STARTTLS", value=bool(email_cfg.get("smtp_use_tls", True)))

            st.markdown("---")
            st.subheader("Default Recipients")
            recipients = st.text_input("To (comma-separated)",
                                       ", ".join(email_cfg.get("default_recipients", [])))
            cc = st.text_input("CC (comma-separated)",
                                ", ".join(email_cfg.get("default_cc", [])))
            subject_tpl = st.text_input("Subject Template",
                                         email_cfg.get("default_subject_template",
                                                        "Monthly Invoice - {month} {year}"))
            body_tpl = st.text_area("Body Template",
                                     email_cfg.get("default_body_template", ""), height=120)

            save_btn = st.form_submit_button("Save Settings", type="primary")

        if save_btn:
            new_cfg = {
                "smtp_host": smtp_host.strip(),
                "smtp_port": int(smtp_port),
                "smtp_use_tls": use_tls,
                "smtp_username": smtp_user.strip(),
                "smtp_password": smtp_pass,
                "sender_name": sender_name.strip(),
                "sender_email": sender_email.strip(),
                "default_recipients": [r.strip() for r in recipients.split(",") if r.strip()],
                "default_cc": [c.strip() for c in cc.split(",") if c.strip()],
                "default_subject_template": subject_tpl.strip(),
                "default_body_template": body_tpl,
            }
            try:
                save_email_config(new_cfg)
                st.success("Settings saved!")
            except Exception as e:
                st.error(f"Save failed: {e}")

        st.markdown("---")
        st.subheader("Test SMTP Connection")
        if st.button("Test Connection"):
            from services.email_service import test_smtp_connection
            cfg_to_test = {
                "smtp_host": email_cfg.get("smtp_host", ""),
                "smtp_port": email_cfg.get("smtp_port", 587),
                "smtp_use_tls": email_cfg.get("smtp_use_tls", True),
                "smtp_username": email_cfg.get("smtp_username", ""),
                "smtp_password": email_cfg.get("smtp_password", ""),
            }
            try:
                success, msg = test_smtp_connection(cfg_to_test)
                if success:
                    st.success(msg)
                else:
                    st.error(msg)
            except Exception as e:
                st.error(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: TICKETS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🎫 Tickets":
    from services.supabase_service import (
        get_tickets, create_ticket, update_ticket, delete_ticket,
    )

    st.title("🎫 Support Tickets")

    # ── Summary metrics ───────────────────────────────────────────────────────
    all_tickets = get_tickets()
    open_count     = sum(1 for t in all_tickets if t["status"] == "Open")
    inprog_count   = sum(1 for t in all_tickets if t["status"] == "In Progress")
    resolved_count = sum(1 for t in all_tickets if t["status"] in ("Resolved", "Closed"))

    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("Total Tickets",   len(all_tickets))
    mc2.metric("Open",            open_count)
    mc3.metric("In Progress",     inprog_count)
    mc4.metric("Resolved/Closed", resolved_count)

    st.markdown("---")

    # ── New ticket form ───────────────────────────────────────────────────────
    if CAN_EDIT:
        with st.expander("➕ New Ticket", expanded=False):
            with st.form("new_ticket_form"):
                project_names = sorted({p.project_name for p in projects})
                t_project  = st.selectbox("Project", project_names, key="nt_proj")
                t_title    = st.text_input("Title", key="nt_title")
                t_desc     = st.text_area("Description", height=100, key="nt_desc")
                t_priority = st.selectbox("Priority", ["Low", "Medium", "High", "Critical"], index=1, key="nt_prio")
                submitted  = st.form_submit_button("Create Ticket", type="primary")
            if submitted:
                if not t_title.strip():
                    st.error("Title is required.")
                else:
                    try:
                        ticket = create_ticket(t_project, t_title.strip(), t_desc.strip(), t_priority)
                        st.success(f"Ticket {ticket.get('ticket_number', '')} created!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Failed to create ticket: {e}")

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander("🔍 Filters", expanded=True):
        fc1, fc2, fc3, fc4 = st.columns(4)
        project_names_all = ["All"] + sorted({t["project_name"] for t in all_tickets})
        tf_proj   = fc1.selectbox("Project",  project_names_all, key="tf_proj")
        tf_status = fc2.selectbox("Status",   ["All", "Open", "In Progress", "Resolved", "Closed"], key="tf_status")
        tf_prio   = fc3.selectbox("Priority", ["All", "Critical", "High", "Medium", "Low"], key="tf_prio")
        tf_search = fc4.text_input("Search title", key="tf_search")

    filtered_tickets = all_tickets
    if tf_proj != "All":
        filtered_tickets = [t for t in filtered_tickets if t["project_name"] == tf_proj]
    if tf_status != "All":
        filtered_tickets = [t for t in filtered_tickets if t["status"] == tf_status]
    if tf_prio != "All":
        filtered_tickets = [t for t in filtered_tickets if t["priority"] == tf_prio]
    if tf_search.strip():
        filtered_tickets = [t for t in filtered_tickets if tf_search.lower() in t["title"].lower()]

    st.caption(f"Showing {len(filtered_tickets)} of {len(all_tickets)} tickets")

    # ── Ticket table ──────────────────────────────────────────────────────────
    PRIORITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
    STATUS_ORDER   = {"Open": 0, "In Progress": 1, "Resolved": 2, "Closed": 3}
    filtered_tickets = sorted(
        filtered_tickets,
        key=lambda t: (STATUS_ORDER.get(t["status"], 9), PRIORITY_ORDER.get(t["priority"], 9)),
    )

    def _prio_color(val):
        return {"Critical": "color:#C0392B;font-weight:bold",
                "High":     "color:#E67E22;font-weight:bold",
                "Medium":   "color:#2980B9",
                "Low":      "color:#27AE60"}.get(val, "")

    def _status_color(val):
        return {"Open":        "color:#E74C3C;font-weight:bold",
                "In Progress": "color:#F39C12;font-weight:bold",
                "Resolved":    "color:#27AE60",
                "Closed":      "color:#95A5A6"}.get(val, "")

    if filtered_tickets:
        ticket_df = pd.DataFrame([{
            "Ticket #":    t["ticket_number"],
            "Project":     t["project_name"],
            "Title":       t["title"],
            "Priority":    t["priority"],
            "Status":      t["status"],
            "Created":     t["created_at"][:10] if t.get("created_at") else "",
            "Updated":     t["updated_at"][:10] if t.get("updated_at") else "",
        } for t in filtered_tickets])

        st.dataframe(
            ticket_df.style
                .map(_prio_color,    subset=["Priority"])
                .map(_status_color,  subset=["Status"]),
            use_container_width=True,
            hide_index=True,
            height=350,
        )
    else:
        st.info("No tickets match the current filters.")

    # ── Edit / update a ticket ────────────────────────────────────────────────
    if CAN_EDIT and filtered_tickets:
        st.markdown("---")
        st.subheader("Update Ticket")
        ticket_options = {f"{t['ticket_number']} — {t['title']}": t for t in filtered_tickets}
        selected_label = st.selectbox("Select ticket to update", list(ticket_options.keys()), key="sel_ticket")
        sel_ticket = ticket_options[selected_label]

        with st.form("update_ticket_form"):
            uc1, uc2 = st.columns(2)
            new_status   = uc1.selectbox("Status",   ["Open", "In Progress", "Resolved", "Closed"],
                                          index=["Open", "In Progress", "Resolved", "Closed"].index(sel_ticket["status"])
                                          if sel_ticket["status"] in ["Open", "In Progress", "Resolved", "Closed"] else 0,
                                          key="upd_status")
            new_priority = uc2.selectbox("Priority", ["Low", "Medium", "High", "Critical"],
                                          index=["Low", "Medium", "High", "Critical"].index(sel_ticket["priority"])
                                          if sel_ticket["priority"] in ["Low", "Medium", "High", "Critical"] else 1,
                                          key="upd_prio")
            new_notes = st.text_area("Notes / update comment", value=sel_ticket.get("notes") or "", height=80, key="upd_notes")
            col_upd, col_del = st.columns([3, 1])
            upd_btn = col_upd.form_submit_button("💾 Update Ticket", type="primary")
            del_btn = col_del.form_submit_button("🗑️ Delete", type="secondary")

        if upd_btn:
            try:
                update_ticket(sel_ticket["id"], status=new_status, priority=new_priority, notes=new_notes)
                st.success("Ticket updated!")
                st.rerun()
            except Exception as e:
                st.error(f"Update failed: {e}")

        if del_btn:
            try:
                delete_ticket(sel_ticket["id"])
                st.success(f"Ticket {sel_ticket['ticket_number']} deleted.")
                st.rerun()
            except Exception as e:
                st.error(f"Delete failed: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: BANK PAYMENT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🏦 Bank Payment":
    from services.bank_pdf_service import parse_swift_pdf
    from services.supabase_service import (
        get_invoices_by_number, mark_invoice_row_paid,
        get_subscription, upsert_subscription, create_renewal_link,
    )

    st.title("🏦 Bank Payment Import")

    if not CAN_EDIT:
        st.warning("Read-only mode — contact an admin to record payments.")
        st.stop()

    # ── Show renewal result from previous confirmation ────────────────────────
    if "_renewal_result" in st.session_state:
        _rr = st.session_state.pop("_renewal_result")
        st.success(
            f"✅ Invoice **#{_rr['inv_no']}** — marked **{_rr['count']}** project(s) as paid."
        )
        _base_url = st.secrets.get("app", {}).get("base_url", "").rstrip("/")
        if _rr["links"]:
            st.markdown("### 🔑 Renewal Links")
            st.caption(
                "Send each link to the customer. "
                "When opened, the subscription is activated automatically."
            )
            for _rl in _rr["links"]:
                _url = f"{_base_url}/?token={_rl['token']}" if _base_url else f"?token={_rl['token']}"
                with st.container(border=True):
                    _lc1, _lc2 = st.columns([2, 3])
                    _lc1.markdown(
                        f"**{_rl['project']}**  \n"
                        f"📅 Valid until: **{_rl['valid_until'].strftime('%d %b %Y')}**  \n"
                        f"📷 Cameras: **{_rl['cameras']}**"
                    )
                    _lc2.code(_url, language=None)
        st.markdown("---")

    def _render_invoice_lookup(inv_no: int, pay_date: datetime.date, key_prefix: str):
        """Look up rows for inv_no and render selection + confirm UI."""
        if inv_no <= 0:
            st.info("Enter a valid invoice number to look up matching rows.")
            return

        with st.spinner("Looking up invoice in database…"):
            try:
                rows = get_invoices_by_number(inv_no)
            except Exception as exc:
                st.error(f"Lookup failed: {exc}")
                return

        if not rows:
            st.warning(f"No invoice rows found for invoice **#{inv_no}**.")
            return

        df = pd.DataFrame([{
            "id":           r["id"],
            "Project":      r["project_name"],
            "Maint. Year":  r["maintenance_year"],
            "Amount (€)":   _safe_float(r.get("payment_amount")),
            "Cameras":      _safe_int(r.get("cameras_number")),
            "Paid":         r.get("paid", "No"),
            "Payment Date": (r["payment_date"] or "")[:10] if r.get("payment_date") else "",
        } for r in rows])

        st.dataframe(
            df[["Project", "Maint. Year", "Amount (€)", "Cameras", "Paid", "Payment Date"]],
            use_container_width=True,
            hide_index=True,
        )

        unpaid_projects = list(df[df["Paid"] != "Yes"]["Project"])
        if not unpaid_projects:
            st.success(f"All rows for invoice **#{inv_no}** are already marked as paid.")
            return

        selected = st.multiselect(
            "Select project row(s) to mark as paid",
            options=list(df["Project"]),
            default=unpaid_projects,
            key=f"{key_prefix}_sel",
        )

        if not selected:
            return

        confirm_date = st.date_input(
            "Payment date to record",
            value=pay_date,
            key=f"{key_prefix}_date",
        )

        sel_df = df[df["Project"].isin(selected)][["Project", "Amount (€)", "Cameras"]].copy()
        st.caption("The following rows will be marked paid. If subscription tables exist, renewal links will also be generated.")
        st.dataframe(sel_df, use_container_width=True, hide_index=True)

        if st.button("✅ Confirm — Mark as Paid & Generate Renewal Links", type="primary", key=f"{key_prefix}_confirm"):
            errors = []
            renewal_links = []
            renewal_warnings = []

            for _, row in df[df["Project"].isin(selected)].iterrows():
                proj = row["Project"]
                try:
                    # 1. Mark invoice row as paid
                    mark_invoice_row_paid(
                        db_id=int(row["id"]),
                        payment_date=confirm_date,
                        payment_amount=row["Amount (€)"],
                    )

                    try:
                        # 2. Compute new valid_until (extend from current expiry or today)
                        sub = get_subscription(proj)
                        if sub and sub.get("valid_until"):
                            current_until = datetime.date.fromisoformat(sub["valid_until"][:10])
                            base = max(current_until, confirm_date)
                        else:
                            base = confirm_date
                        try:
                            target_until = base.replace(year=base.year + 1)
                        except ValueError:           # Feb 29 edge case
                            target_until = base.replace(year=base.year + 1, day=28)

                        cameras = _safe_int(row["Cameras"])

                        # 3. Upsert subscription record
                        upsert_subscription(
                            project_name=proj,
                            valid_until=target_until,
                            cameras_allowed=cameras,
                            valid_from=confirm_date,
                        )

                        # 4. Generate renewal token
                        token = create_renewal_link(
                            project_name=proj,
                            target_valid_until=target_until,
                            cameras_allowed=cameras,
                            invoice_number=str(inv_no),
                            payment_amount=row["Amount (€)"],
                        )

                        renewal_links.append({
                            "project":     proj,
                            "valid_until": target_until,
                            "cameras":     cameras,
                            "token":       token,
                        })
                    except Exception as renewal_exc:
                        if (
                            _is_missing_supabase_table_error(renewal_exc, "subscriptions")
                            or _is_missing_supabase_table_error(renewal_exc, "renewal_links")
                        ):
                            renewal_warnings.append(
                                f"{proj}: payment recorded, but renewal tables are not set up in Supabase yet."
                            )
                        else:
                            raise renewal_exc

                except Exception as exc:
                    errors.append(f"{proj}: {exc}")

            if errors:
                st.error("Some updates failed:\n" + "\n".join(errors))
            else:
                if renewal_warnings:
                    st.warning("\n".join(renewal_warnings))
                st.cache_data.clear()
                st.session_state["_renewal_result"] = {
                    "inv_no": inv_no,
                    "count":  len(selected),
                    "links":  renewal_links,
                }
                st.rerun()

    # ── PDF upload section ────────────────────────────────────────────────────
    st.markdown("Upload a SWIFT MT103 bank transfer PDF to auto-extract payment details.")
    uploaded = st.file_uploader("Upload bank transfer PDF", type=["pdf"])

    if uploaded:
        with st.spinner("Parsing PDF…"):
            try:
                parsed = parse_swift_pdf(uploaded.read())
            except Exception as exc:
                st.error(f"Failed to parse PDF: {exc}")
                st.stop()

        st.markdown("### Extracted Payment Data")
        st.caption("All fields are editable — correct any parsing errors before proceeding.")

        ec1, ec2, ec3 = st.columns(3)
        inv_no = ec1.number_input(
            "Invoice #",
            value=int(parsed["invoice_number"]) if parsed["invoice_number"] else 0,
            min_value=0, step=1, key="pdf_inv_no",
        )
        pay_date = ec2.date_input(
            "Payment Date",
            value=parsed["payment_date"] if parsed["payment_date"] else datetime.date.today(),
            key="pdf_pay_date",
        )
        ec3.metric(
            "Instructed Amount",
            f"€{parsed['instructed_amount']:,.2f}" if parsed["instructed_amount"] else "—",
        )

        if parsed["received_amount"] and parsed["instructed_amount"]:
            fee = parsed["instructed_amount"] - parsed["received_amount"]
            if fee > 0:
                st.info(
                    f"Bank fee deducted: €{fee:,.2f}  "
                    f"(instructed €{parsed['instructed_amount']:,.2f} → "
                    f"received €{parsed['received_amount']:,.2f}). "
                    "The invoiced amount will be kept as-is."
                )

        if not parsed["invoice_number"] and not parsed["payment_date"]:
            st.warning(
                "Could not extract payment data from this PDF. "
                "Use the manual lookup below instead."
            )

        st.markdown("---")
        st.markdown("### Matching Invoice Rows")
        _render_invoice_lookup(inv_no, pay_date, key_prefix="pdf")

    # ── Manual lookup (no PDF) ────────────────────────────────────────────────
    st.markdown("---")
    with st.expander("🔍 Manual Lookup (no PDF)", expanded=not bool(uploaded)):
        mc1, mc2 = st.columns(2)
        manual_inv_no = mc1.number_input(
            "Invoice #", min_value=0, step=1, key="manual_inv_no",
        )
        manual_date = mc2.date_input(
            "Payment Date", value=datetime.date.today(), key="manual_date",
        )
        _render_invoice_lookup(manual_inv_no, manual_date, key_prefix="manual")
