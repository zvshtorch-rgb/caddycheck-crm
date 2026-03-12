"""CaddyCheck CRM — Streamlit web app (role-based access)."""
import datetime
import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

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

from config.settings import MONTH_ORDER, get_email_config, save_email_config
from services.excel_service import (
    load_projects,
    load_invoices,
    compute_debt_summaries,
    get_yearly_summary,
    save_projects_to_excel,
    save_invoices_to_excel,
    get_projects_for_month,
)
from services.invoice_service import generate_monthly_invoice, get_invoice_preview_data

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
    from config.settings import get_data_paths
    paths = get_data_paths()
    projects = load_projects(paths["projects_file"])
    invoices = load_invoices(paths["projects_file"])
    debt_summaries = compute_debt_summaries(projects, invoices)
    yearly_summary = get_yearly_summary(invoices)
    return projects, invoices, debt_summaries, yearly_summary, str(paths["projects_file"])


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
    ["📊 Dashboard", "🏗️ Projects", "🧾 Invoice Details", "💸 Debt Report", "📅 Monthly Invoice", "⚙️ Settings"],
    label_visibility="collapsed",
)
st.sidebar.markdown("---")
role_icon = "✏️" if CAN_EDIT else "👁️"
st.sidebar.caption(f"{role_icon} Logged in as **{ROLES[ROLE]['label']}**")
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


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if page == "📊 Dashboard":
    st.title("📊 Dashboard")

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

    cur_month  = datetime.datetime.now().month
    monthly_val = sum(
        i.payment_amount for i in f_inv
        if i.is_paid() and i.payment_date
        and i.payment_date.year == ref_year and i.payment_date.month == cur_month
    )
    yearly_val = sum(
        i.payment_amount for i in f_inv
        if i.is_paid() and i.year == ref_year
    )

    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    with c1: card("Total Income",            f"€{total_income:,.0f}", "card-income")
    with c2: card("Total Paid",              f"€{total_paid:,.0f}",   "card-paid")
    with c3: card("Total Debt",              f"€{total_unpaid:,.0f}", "card-debt")
    with c4: card(f"Monthly Income ({ref_year})", f"€{monthly_val:,.0f}", "card-monthly")
    with c5: card(f"Yearly Income ({ref_year})",  f"€{yearly_val:,.0f}",  "card-yearly")
    with c6: card("Active Projects",         str(active_count),       "card-projects")
    with c7: card("Total Cameras",           str(total_cams),         "card-cameras")

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
# PAGE: PROJECTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🏗️ Projects":
    st.title("🏗️ Projects")

    # Filters
    col1, col2, col3 = st.columns(3)
    countries = sorted({p.country for p in projects if p.country})
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
        "Activation Date": p.activation_date.strftime("%Y-%m-%d") if p.activation_date else "",
        "Status":          _safe_str(p.status),
        "License EOP":     p.license_eop.strftime("%Y-%m-%d") if p.license_eop else "",
    } for p in filtered])

    def color_status(val):
        if str(val).strip().lower() == "active":
            return "color: #27AE60; font-weight: bold"
        return "color: #E74C3C"

    if CAN_EDIT:
        st.info("✏️ Admin mode: you can edit cells directly. Click **Save Changes** when done.")

        if st.button("➕ Add New Project", key="btn_add_proj"):
            st.session_state["add_proj_row"] = True

        _empty_proj = {"Project Name": "", "Country": "", "# Cams": 0,
                       "Payment Month": "", "Install Year": "",
                       "Activation Date": "", "Status": "Active", "License EOP": ""}
        if st.session_state.get("add_proj_row"):
            df_edit = pd.concat([pd.DataFrame([_empty_proj]), df.reset_index(drop=True)], ignore_index=True)
        else:
            df_edit = df.reset_index(drop=True)

        edited_df = st.data_editor(
            df_edit,
            use_container_width=True,
            height=600,
            num_rows="dynamic",
            key="proj_editor",
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
                eop = row.get("License EOP", "")
                if eop:
                    try:
                        p.license_eop = datetime.datetime.strptime(str(eop), "%Y-%m-%d")
                    except Exception:
                        pass
                else:
                    p.license_eop = None
            try:
                save_projects_to_excel(projects)
                load_data.clear()
                st.session_state.pop("add_proj_row", None)
                msg = f"Saved! {new_count} new project(s) added." if new_count else "Projects saved successfully!"
                st.success(msg)
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
# PAGE: INVOICE DETAILS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🧾 Invoice Details":
    st.title("🧾 Invoice Details")

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

    df_inv = pd.DataFrame([{
        "Invoice #":      _safe_str(_safe_int(i.invoice_number) or ""),
        "Project":        _safe_str(i.project_name),
        "Maint. Year":    _safe_str(i.maintenance_year),
        "Amount (€)":     _safe_float(i.payment_amount),
        "Cameras":        _safe_int(i.cameras_number),
        "Payment Date":   i.payment_date.strftime("%Y-%m-%d") if i.payment_date else "",
        "Paid":           _safe_str(i.paid),
        "Year":           _safe_str(_safe_int(i.year) or ""),
    } for i in filtered_inv])

    def color_paid(val):
        v = str(val).strip().lower()
        if v == "yes":     return "color: #27AE60; font-weight: bold"
        if v == "no":      return "color: #E74C3C"
        if v == "cancelled": return "color: #F39C12"
        return ""

    if CAN_EDIT:
        st.info("✏️ Admin mode: you can edit cells directly. Click **Save Changes** when done.")

        if st.button("➕ Add New Invoice", key="btn_add_inv"):
            st.session_state["add_inv_row"] = True

        _empty_inv = {"Invoice #": "", "Project": "", "Maint. Year": "Y1",
                      "Amount (€)": 0.0, "Cameras": 0,
                      "Payment Date": "", "Paid": "No", "Year": str(datetime.date.today().year)}
        if st.session_state.get("add_inv_row"):
            df_inv_edit = pd.concat([pd.DataFrame([_empty_inv]), df_inv.reset_index(drop=True)], ignore_index=True)
        else:
            df_inv_edit = df_inv.reset_index(drop=True)

        edited_inv = st.data_editor(
            df_inv_edit,
            use_container_width=True,
            height=550,
            num_rows="dynamic",
            key="inv_editor",
        )
        if st.button("💾 Save Changes", key="save_invoices"):
            from models.invoice import Invoice as InvoiceModel
            inv_map = {i.invoice_number: i for i in invoices if i.invoice_number}
            new_count = 0
            for _, row in edited_inv.iterrows():
                project = _safe_str(row.get("Project", "")).strip()
                if not project:
                    continue
                inv_no_str = _safe_str(row.get("Invoice #", "")).strip()
                try:
                    inv_no = float(inv_no_str) if inv_no_str else None
                except Exception:
                    inv_no = None
                inv = inv_map.get(inv_no) if inv_no else None
                if inv is None:
                    # New invoice row
                    inv = InvoiceModel(
                        invoice_number=inv_no,
                        project_name=project,
                        maintenance_year=_safe_str(row.get("Maint. Year", "")),
                        payment_amount=_safe_float(row.get("Amount (€)", 0)),
                        cameras_number=_safe_int(row.get("Cameras", 0)) or None,
                        payment_date=None,
                        paid=_safe_str(row.get("Paid", "No")),
                        year=_safe_int(row.get("Year")) or None,
                    )
                    invoices.append(inv)
                    if inv_no:
                        inv_map[inv_no] = inv
                    new_count += 1
                else:
                    inv.project_name   = project
                    inv.maintenance_year = _safe_str(row.get("Maint. Year", ""))
                    inv.paid           = _safe_str(row.get("Paid", ""))
                    inv.payment_amount = _safe_float(row.get("Amount (€)", 0))
                    inv.cameras_number = _safe_int(row.get("Cameras", 0)) or None
                    inv.year           = _safe_int(row.get("Year")) or None
                pd_str = _safe_str(row.get("Payment Date", "")).strip()
                if pd_str:
                    try:
                        inv.payment_date = datetime.datetime.strptime(pd_str, "%Y-%m-%d")
                    except Exception:
                        pass
            try:
                save_invoices_to_excel(invoices)
                load_data.clear()
                st.session_state.pop("add_inv_row", None)
                msg = f"Saved! {new_count} new invoice(s) added." if new_count else "Invoices saved successfully!"
                st.success(msg)
            except Exception as e:
                st.error(f"Save failed: {e}")
    else:
        st.dataframe(
            df_inv.style.map(color_paid, subset=["Paid"]),
            use_container_width=True,
            height=550,
        )

    # ── Debt Summary ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Debt Summary by Project")
    debt_df = pd.DataFrame([{
        "Project":        _safe_str(ds.project_name),
        "Country":        _safe_str(ds.country),
        "# Cams":         _safe_int(ds.num_cams),
        "Status":         _safe_str(ds.status),
        "Expected (€)":   f"€{ds.total_expected:,.0f}",
        "Paid (€)":       f"€{ds.total_paid:,.0f}",
        "Cancelled (€)":  f"€{ds.total_cancelled:,.0f}",
        "Debt (€)":       f"€{ds.total_unpaid:,.0f}",
    } for ds in debt_summaries])

    if sel_country != "All":
        country_proj = {p.project_name for p in projects if p.country == sel_country}
        debt_df = debt_df[debt_df["Project"].isin(country_proj)]

    def color_debt(val):
        try:
            v = float(str(val).replace("€", "").replace(",", ""))
            if v > 0:   return "color: #E74C3C; font-weight: bold"
            return "color: #27AE60"
        except Exception:
            return ""

    st.dataframe(
        debt_df.style.map(color_debt, subset=["Debt (€)"]),
        use_container_width=True,
        height=400,
    )


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DEBT REPORT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "💸 Debt Report":
    st.title("💸 Debt Report")

    # ── Filters ───────────────────────────────────────────────────────────────
    with st.expander("🔍 Filters", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        debt_years    = sorted({inv.year for inv in invoices if inv.year}, reverse=True)
        dsel_year     = fc1.selectbox("Year", ["All"] + [str(y) for y in debt_years], key="dr_year")
        debt_countries = sorted({p.country for p in projects if p.country})
        dsel_country  = fc2.selectbox("Country", ["All"] + debt_countries, key="dr_country")
        dsel_search   = fc3.text_input("Search project", key="dr_search")

    # Only unpaid invoices
    debt_inv = [i for i in invoices if i.is_unpaid()]
    if dsel_year != "All":
        debt_inv = [i for i in debt_inv if i.year == int(dsel_year)]

    # Country lookup — normalize names aggressively, fallback to debt_summaries
    import re as _re
    def _norm(s):
        return _re.sub(r'\s+', ' ', str(s or "").strip().lower())
    proj_country_map = {_norm(p.project_name): p.country for p in projects}
    ds_country_map   = {_norm(ds.project_name): ds.country for ds in debt_summaries}
    def _get_country(name):
        k = _norm(name)
        return proj_country_map.get(k) or ds_country_map.get(k) or ""

    if dsel_country != "All":
        debt_inv = [i for i in debt_inv if _get_country(i.project_name) == dsel_country]
    if dsel_search.strip():
        debt_inv = [i for i in debt_inv if dsel_search.lower() in i.project_name.lower()]

    # ── Summary metrics ───────────────────────────────────────────────────────
    total_debt_amt  = sum(i.payment_amount for i in debt_inv)
    proj_with_debt  = len({i.project_name for i in debt_inv})
    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("Unpaid Invoices",  len(debt_inv))
    mc2.metric("Projects with Debt", proj_with_debt)
    mc3.metric("Total Debt",       f"€{total_debt_amt:,.0f}")

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
    proj_debt: dict = defaultdict(lambda: {"invoices": [], "total": 0.0, "country": ""})
    for i in debt_inv:
        key = i.project_name
        proj_debt[key]["invoices"].append(str(int(i.invoice_number)) if i.invoice_number else "")
        proj_debt[key]["total"]   += i.payment_amount
        proj_debt[key]["country"]  = _get_country(i.project_name)

    summary_rows = [{
        "Project Name":    name,
        "Country":         d["country"],
        "Invoice Numbers": ", ".join(filter(None, d["invoices"])),
        "Total Debt (€)":  d["total"],
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

        # Download summary as PDF
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib import colors as rl_colors
            from reportlab.lib.styles import getSampleStyleSheet
            from io import BytesIO
            pdf_buf = BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=A4,
                                    leftMargin=18*3.54, rightMargin=18*3.54,
                                    topMargin=16*3.54, bottomMargin=16*3.54)
            styles = getSampleStyleSheet()
            elems = []
            elems.append(Paragraph("<b>CaddyCheck CRM — Debt Report</b>",
                                   styles["Title"]))
            elems.append(Paragraph(
                f"Total Debt: €{total_debt_amt:,.0f}  |  "
                f"Projects: {proj_with_debt}  |  Invoices: {len(debt_inv)}",
                styles["Normal"]))
            elems.append(Spacer(1, 12))
            tbl_data = [["Project Name", "Country", "Invoice Numbers", "Total Debt (€)"]]
            for r in summary_rows:
                tbl_data.append([
                    r["Project Name"], r["Country"],
                    r["Invoice Numbers"], f"€{r['Total Debt (€)']:,.0f}",
                ])
            t = Table(tbl_data, repeatRows=1,
                      colWidths=["35%", "12%", "33%", "20%"])
            t.setStyle(TableStyle([
                ("BACKGROUND",    (0, 0), (-1, 0), rl_colors.HexColor("#1B3A6B")),
                ("TEXTCOLOR",     (0, 0), (-1, 0), rl_colors.white),
                ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE",      (0, 0), (-1, -1), 8),
                ("ROWBACKGROUNDS",(0, 1), (-1, -1),
                 [rl_colors.white, rl_colors.HexColor("#EBF5FB")]),
                ("GRID",          (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#BDC3C7")),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ]))
            elems.append(t)
            doc.build(elems)
            st.download_button("⬇️ Download Summary PDF", pdf_buf.getvalue(),
                               file_name="debt_summary.pdf", mime="application/pdf")
        except Exception as e:
            st.warning(f"PDF export unavailable: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE: MONTHLY INVOICE
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📅 Monthly Invoice":
    st.title("📅 Monthly Invoice")

    col1, col2, col3 = st.columns([2, 1, 2])
    with col1:
        sel_month = st.selectbox("Month", MONTH_ORDER, index=datetime.date.today().month - 1)
    with col2:
        sel_year = st.number_input("Year", min_value=2015, max_value=2035,
                                   value=datetime.date.today().year, step=1)
    with col3:
        invoice_number = st.number_input("Invoice Number (optional, 0 = auto)",
                                         min_value=0, value=0, step=1)

    month_projects = get_projects_for_month(projects, sel_month)
    st.markdown(f"**{len(month_projects)} project(s)** billed in **{sel_month}**")

    if not month_projects:
        st.warning("No projects found for the selected month.")
    else:
        preview_rows = get_invoice_preview_data(month_projects, sel_month, int(sel_year))
        preview_df = pd.DataFrame([{
            "Project": _safe_str(r["project_name"]),
            "# Cams": _safe_str(r["num_cams"]),
            "Maint. Year": _safe_str(r["maintenance_year"]),
            "Rate (€)": f"€{r['rate']:,.0f}" if isinstance(r["rate"], (int, float)) else _safe_str(r["rate"]),
            "Line Total (€)": f"€{r['line_total']:,.0f}" if isinstance(r["line_total"], (int, float)) else "",
        } for r in preview_rows])
        st.subheader("Invoice Preview")
        st.dataframe(preview_df, use_container_width=True, hide_index=True)

        if CAN_EDIT:
            st.markdown("---")
            st.subheader("Generate Invoice")

            from config.settings import get_data_paths
            from services.pdf_service import generate_invoice_pdf
            paths = get_data_paths()
            inv_no = invoice_number if invoice_number > 0 else None
            month_abbr = sel_month[:3]
            pdf_filename = f"CC_M_inv_{inv_no or 'auto'}_{month_abbr}_{int(sel_year)}.pdf"
            xlsx_filename = f"CC_M_inv_{inv_no or 'auto'}_{month_abbr}_{int(sel_year)}.xlsx"

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

            # ── Email section ──────────────────────────────────────────────
            st.markdown("---")
            st.subheader("Send Invoice by Email")
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
                send_btn = st.form_submit_button("Send Email")

            if send_btn:
                if not email_cfg.get("smtp_username"):
                    st.error("SMTP not configured. Go to ⚙️ Settings to set up email.")
                else:
                    from services.email_service import send_invoice_email
                    import tempfile
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        try:
                            out_path = generate_monthly_invoice(
                                projects=month_projects,
                                month_name=sel_month,
                                year=int(sel_year),
                                invoice_number=invoice_number if invoice_number > 0 else None,
                                output_dir=Path(tmp_dir),
                                template_path=paths["invoice_template"],
                            )
                            recipients = [r.strip() for r in to_addrs.split(",") if r.strip()]
                            cc_list = [c.strip() for c in cc_addrs.split(",") if c.strip()]
                            success, msg = send_invoice_email(
                                attachment_path=out_path,
                                recipients=recipients,
                                cc=cc_list,
                                subject=subject,
                                body=body,
                                config=email_cfg,
                            )
                            if success:
                                st.success(f"Email sent! {msg}")
                            else:
                                st.error(f"Email failed: {msg}")
                        except Exception as e:
                            st.error(f"Error: {e}")
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
