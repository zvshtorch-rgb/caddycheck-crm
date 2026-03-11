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

from config.settings import MONTH_ORDER
from services.excel_service import (
    load_projects,
    load_invoices,
    compute_debt_summaries,
    get_yearly_summary,
    save_projects_to_excel,
    save_invoices_to_excel,
)

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
    projects = load_projects()
    invoices = load_invoices()
    debt_summaries = compute_debt_summaries(projects, invoices)
    yearly_summary = get_yearly_summary(invoices)
    return projects, invoices, debt_summaries, yearly_summary


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
    ["📊 Dashboard", "🏗️ Projects", "🧾 Invoice Details"],
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
    projects, invoices, debt_summaries, yearly_summary = load_data()
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
        "Project Name":    p.project_name,
        "Country":         p.country,
        "# Cams":          p.num_cams,
        "Payment Month":   p.payment_month,
        "Install Year":    p.installation_year,
        "Status":          p.status,
    } for p in sorted_proj])

    def color_status(val):
        if str(val).strip().lower() == "active":
            return "color: #27AE60; font-weight: bold"
        return "color: #E74C3C"

    st.dataframe(
        proj_df.style.applymap(color_status, subset=["Status"]),
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
        "Project Name":    p.project_name,
        "Country":         p.country,
        "# Cams":          p.num_cams,
        "Payment Month":   p.payment_month,
        "Install Year":    p.installation_year,
        "Activation Date": p.activation_date.strftime("%Y-%m-%d") if p.activation_date else "",
        "Status":          p.status,
        "License EOP":     p.license_eop.strftime("%Y-%m-%d") if p.license_eop else "",
    } for p in filtered])

    def color_status(val):
        if str(val).strip().lower() == "active":
            return "color: #27AE60; font-weight: bold"
        return "color: #E74C3C"

    if CAN_EDIT:
        st.info("✏️ Admin mode: you can edit cells directly. Click **Save Changes** when done.")
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            height=600,
            num_rows="fixed",
            key="proj_editor",
        )
        if st.button("💾 Save Changes", key="save_projects"):
            # Apply edits back to project objects
            proj_map = {p.project_name: p for p in filtered}
            for _, row in edited_df.iterrows():
                p = proj_map.get(row["Project Name"])
                if p is None:
                    continue
                p.country           = row["Country"]
                p.num_cams          = int(row["# Cams"]) if row["# Cams"] else 0
                p.payment_month     = row["Payment Month"]
                p.installation_year = int(row["Install Year"]) if row["Install Year"] else None
                p.status            = row["Status"]
                if row["License EOP"]:
                    try:
                        p.license_eop = datetime.datetime.strptime(str(row["License EOP"]), "%Y-%m-%d")
                    except Exception:
                        pass
                else:
                    p.license_eop = None
            try:
                save_projects_to_excel(projects)
                load_data.clear()
                st.success("Projects saved successfully!")
            except Exception as e:
                st.error(f"Save failed: {e}")
    else:
        st.dataframe(
            df.style.applymap(color_status, subset=["Status"]),
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
                    "Project":   p.project_name,
                    "Year":      yr,
                    "Maint. Year": p.get_maintenance_year_label(yr),
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

    # Filters
    col1, col2, col3, col4 = st.columns(4)
    years = sorted({inv.year for inv in invoices if inv.year}, reverse=True)
    sel_year    = col1.selectbox("Year",    ["All"] + [str(y) for y in years], key="inv_year")
    sel_paid    = col2.selectbox("Status",  ["All", "Paid", "Unpaid", "Cancelled"], key="inv_paid")
    countries   = sorted({p.country for p in projects if p.country})
    sel_country = col3.selectbox("Country", ["All"] + countries, key="inv_country")
    search      = col4.text_input("Search project", key="inv_search")

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

    # Summary row
    total_paid_f   = sum(i.payment_amount for i in filtered_inv if i.is_paid())
    total_unpaid_f = sum(i.payment_amount for i in filtered_inv if i.is_unpaid())
    c1, c2, c3 = st.columns(3)
    c1.metric("Invoices shown",  len(filtered_inv))
    c2.metric("Total Paid",      f"€{total_paid_f:,.0f}")
    c3.metric("Total Unpaid",    f"€{total_unpaid_f:,.0f}")

    st.caption(f"Showing {len(filtered_inv)} of {len(invoices)} invoices")

    df_inv = pd.DataFrame([{
        "Invoice #":      int(i.invoice_number) if i.invoice_number else "",
        "Project":        i.project_name,
        "Maint. Year":    i.maintenance_year,
        "Amount (€)":     i.payment_amount,
        "Cameras":        int(i.cameras_number) if i.cameras_number else "",
        "Payment Date":   i.payment_date.strftime("%Y-%m-%d") if i.payment_date else "",
        "Paid":           i.paid,
        "Year":           i.year or "",
    } for i in filtered_inv])

    def color_paid(val):
        v = str(val).strip().lower()
        if v == "yes":     return "color: #27AE60; font-weight: bold"
        if v == "no":      return "color: #E74C3C"
        if v == "cancelled": return "color: #F39C12"
        return ""

    if CAN_EDIT:
        st.info("✏️ Admin mode: you can edit cells directly. Click **Save Changes** when done.")
        edited_inv = st.data_editor(
            df_inv,
            use_container_width=True,
            height=550,
            num_rows="fixed",
            key="inv_editor",
        )
        if st.button("💾 Save Changes", key="save_invoices"):
            inv_map = {i.invoice_number: i for i in filtered_inv if i.invoice_number}
            for _, row in edited_inv.iterrows():
                inv_no = row["Invoice #"]
                try:
                    inv_no = float(inv_no) if inv_no != "" else None
                except Exception:
                    inv_no = None
                inv = inv_map.get(inv_no)
                if inv is None:
                    continue
                inv.paid           = str(row["Paid"])
                inv.payment_amount = float(str(row["Amount (€)"]).replace(",", "")) if row["Amount (€)"] != "" else 0.0
                if row["Payment Date"]:
                    try:
                        inv.payment_date = datetime.datetime.strptime(str(row["Payment Date"]), "%Y-%m-%d")
                    except Exception:
                        pass
            try:
                save_invoices_to_excel(invoices)
                load_data.clear()
                st.success("Invoices saved successfully!")
            except Exception as e:
                st.error(f"Save failed: {e}")
    else:
        st.dataframe(
            df_inv.style.applymap(color_paid, subset=["Paid"]),
            use_container_width=True,
            height=550,
        )

    # ── Debt Summary ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Debt Summary by Project")
    debt_df = pd.DataFrame([{
        "Project":        ds.project_name,
        "Country":        ds.country,
        "# Cams":         ds.num_cams,
        "Status":         ds.status,
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
        debt_df.style.applymap(color_debt, subset=["Debt (€)"]),
        use_container_width=True,
        height=400,
    )
