"""Ask Data v2 deterministic tool executor (read-only).

Moved verbatim from streamlit_app.py (Ask Data Phase 1). Every number is computed here in plain
Python from the already-loaded CRM data; the LLM never sees or calculates any of it.

Deliberate, minimal deviations from the original (everything else is byte-for-byte the same logic):
  * `get_next_invoice_number` used the streamlit_app globals `_get_next_invoice_number` and `_data_path`;
    it now calls an injected zero-argument provider (`next_invoice_number`), supplied by the app wrapper.
  * KNOWN LATENT BUG, deliberately preserved in Phase 1: `get_projects_for_month` is not imported in
    streamlit_app.py (nor here), so month-filtered project questions raise NameError exactly as before.
    A characterization test pins this; fixing it is a separate, approved change.
"""
import datetime
from typing import Optional

import pandas as pd

from config.settings import MONTH_ORDER, load_sent_invoices_log
from services.ask_data.helpers import (
    _build_invoice_answer_df,
    _build_project_answer_df,
    _is_maintenance_category,
    _is_new_installation_category,
    _is_paid_trial_category,
    _license_status,
    _normalize_project_status,
    _project_detection_counts,
    _project_license_date,
    _safe_float,
    _safe_int,
    _safe_str,
    DETECTION_CAMERA_FIELDS,
)
from services.license_reconciliation import (
    PUBLIC_COLUMNS as LICENSE_RECONCILIATION_COLUMNS,
    build_license_reconciliation,
    filter_reconciliation,
)


def _camera_metric_for_project(project, camera_type: str) -> int:
    if not camera_type or camera_type == "all":
        return _safe_int(project.num_cams)
    counts = _project_detection_counts(project)
    field_map = {label: field_name for label, field_name, _ in DETECTION_CAMERA_FIELDS}
    field_name = field_map.get(camera_type)
    return _safe_int(counts.get(field_name, 0)) if field_name else 0


def _filter_invoices_for_debt_tools(invoices, projects, args: dict, paid_predicate) -> list:
    rows = [inv for inv in invoices if paid_predicate(inv)]
    if args.get("year") is not None:
        rows = [inv for inv in rows if inv.year == args["year"]]
    if args.get("paid_year") is not None:
        rows = [inv for inv in rows if inv.payment_date and inv.payment_date.year == args["paid_year"]]
    if args.get("country"):
        names_in_country = {p.project_name for p in projects if p.country == args["country"]}
        rows = [inv for inv in rows if inv.project_name in names_in_country]
    if args.get("project"):
        rows = [inv for inv in rows if inv.project_name == args["project"]]
    category = args.get("category")
    if category == "trial":
        rows = [inv for inv in rows if _is_paid_trial_category(inv)]
    elif category == "y1":
        rows = [inv for inv in rows if _is_new_installation_category(inv)]
    elif category == "y2plus":
        rows = [inv for inv in rows if _is_maintenance_category(inv)]
    return rows


def _filter_license_entries(projects, args: dict) -> list:
    """Shared by get_licenses and get_license_payment_status: filter/sort (project, eop_date, status) triples."""
    today = datetime.date.today()
    entries = [(p, _project_license_date(p), _license_status(p, today)) for p in projects]
    if args.get("project"):
        entries = [e for e in entries if e[0].project_name == args["project"]]
    if args.get("country"):
        entries = [e for e in entries if e[0].country == args["country"]]
    if args.get("license_status"):
        entries = [e for e in entries if e[2] == args["license_status"]]
    if args.get("project_status") and args["project_status"] != "All":
        # project_status is the project's OWN operational status (Active/Offline/Cancelled/...),
        # deliberately separate from license_status (the computed license classification above) —
        # lets a question distinguish "expired license" from "project still operationally active".
        entries = [e for e in entries if _normalize_project_status(e[0].status) == args["project_status"]]
    if args.get("eop_year") is not None:
        entries = [e for e in entries if e[1] and e[1].year == args["eop_year"]]
    if args.get("eop_month"):
        month_index = MONTH_ORDER.index(args["eop_month"]) + 1
        entries = [e for e in entries if e[1] and e[1].month == month_index]
    if args.get("eop_date") is not None:
        entries = [e for e in entries if e[1] == args["eop_date"]]
    entries.sort(key=lambda e: e[1] or datetime.date.max)
    return entries


def _annual_payment_month(project) -> str:
    """Resolve the project's recurring annual billing month: `payment_month` (the real
    field driving get_projects_for_month/the monthly invoicing job) with a fallback to
    the License EOP's month name if payment_month is blank."""
    month = _safe_str(getattr(project, "payment_month", "")).strip()
    if month in MONTH_ORDER:
        return month
    eop = _project_license_date(project)
    return MONTH_ORDER[eop.month - 1] if eop else ""


def _latest_annual_invoice_year(annual_month: str, today: datetime.date) -> Optional[int]:
    """Latest calendar year whose Annual Payment Month has already started (<= today)."""
    if annual_month not in MONTH_ORDER:
        return None
    month_index = MONTH_ORDER.index(annual_month) + 1
    return today.year if today.month >= month_index else today.year - 1


def _annual_license_payment_row(project, eop, invoices, today: datetime.date) -> dict:
    """
    Determine whether THIS project's latest already-due annual license/maintenance
    invoice (not simply its most recent/highest-numbered invoice) has been paid.
    See `_annual_payment_month`/`_latest_annual_invoice_year` for how the due period
    is anchored. Historical all-time totals are included for reference only — they
    never influence the Annual Payment Status.
    """
    key = _safe_str(project.project_name).lower().strip()
    active_rows = [
        inv for inv in invoices
        if _safe_str(inv.project_name).lower().strip() == key and not inv.is_cancelled()
    ]
    hist_invoiced = sum(float(inv.payment_amount) for inv in active_rows)
    hist_paid = sum(float(inv.payment_amount) for inv in active_rows if inv.is_paid())
    hist_outstanding = sum(float(inv.payment_amount) for inv in active_rows if inv.is_unpaid())

    annual_month = _annual_payment_month(project)
    target_year = _latest_annual_invoice_year(annual_month, today)

    invoice_number, period_text = "", ""
    invoice_amount, outstanding = 0.0, 0.0
    paid_flag = "N/A"

    if target_year is None:
        status = "No Annual Invoice Found"
    elif project.installation_year and target_year < project.installation_year:
        status = "Not Yet Due"
    else:
        maintenance_year_number = project.get_maintenance_year(target_year) if project.installation_year else None
        period_rows = [inv for inv in active_rows if inv.year == target_year]
        label = f"Y{maintenance_year_number}" if maintenance_year_number else (period_rows[0].maintenance_year if period_rows else "")
        period_text = f"{target_year} ({label})" if label else str(target_year)
        if not period_rows:
            status = "No Annual Invoice Found"
        else:
            invoice_amount = sum(float(inv.payment_amount) for inv in period_rows)
            outstanding = sum(float(inv.payment_amount) for inv in period_rows if inv.is_unpaid())
            invoice_number = ", ".join(
                str(n) for n in sorted({_safe_int(inv.invoice_number, default=0) for inv in period_rows if inv.invoice_number})
            )
            paid_flag = "Yes" if outstanding == 0 else "No"
            status = "Paid" if outstanding == 0 else "Unpaid"

    return {
        "Project": project.project_name,
        "Country": project.country,
        "License EOP": eop.isoformat() if eop else "",
        "Annual Payment Month": annual_month or "Unknown",
        "Latest Annual Invoice Number": invoice_number,
        "Latest Annual Invoice Period / Maintenance Year": period_text,
        "Latest Annual Invoice Amount": invoice_amount,
        "Latest Annual Invoice Paid?": paid_flag,
        "Latest Annual Invoice Outstanding Amount": outstanding,
        "Annual Payment Status": status,
        "Total Invoiced (All-Time)": hist_invoiced,
        "Total Paid (All-Time)": hist_paid,
        "Outstanding Balance (All-Time)": hist_outstanding,
    }


def _execute_ask_data_tool(
    tool_name: str, args: dict, projects, invoices, debt_summaries, *, next_invoice_number=None
) -> tuple[str, Optional[pd.DataFrame]]:
    """
    Execute one validated, read-only tool call against the already-loaded
    CRM data. All numbers are computed here in plain Python — the LLM never
    sees or calculates any of this.
    """
    if tool_name == "get_invoice":
        invoice_number = args["invoice_number"]
        rows = [inv for inv in invoices if _safe_int(inv.invoice_number, default=0) == invoice_number]
        if args.get("year") is not None:
            rows = [inv for inv in rows if inv.year == args["year"]]
        if not rows:
            return f"Invoice {invoice_number} is not in the invoice ledger.", None
        total = sum(float(inv.payment_amount) for inv in rows)
        if len(rows) == 1:
            paid_txt = "Cancelled" if rows[0].is_cancelled() else ("Paid" if rows[0].is_paid() else "Unpaid")
            summary = f"Invoice {invoice_number} totals €{total:,.0f} and is {paid_txt}."
        else:
            summary = (
                f"Invoice {invoice_number} has {len(rows)} row(s) totaling €{total:,.0f}; "
                "see the Paid column below for each row's status."
            )
        return summary, _build_invoice_answer_df(rows, projects)

    if tool_name in ("get_invoices", "get_unpaid_invoices"):
        rows = list(invoices)
        if tool_name == "get_unpaid_invoices":
            rows = [inv for inv in rows if inv.is_unpaid()]
        elif args.get("paid") == "yes":
            rows = [inv for inv in rows if inv.is_paid()]
        elif args.get("paid") == "no":
            rows = [inv for inv in rows if inv.is_unpaid()]
        elif args.get("paid") == "cancelled":
            rows = [inv for inv in rows if inv.is_cancelled()]
        if args.get("project"):
            rows = [inv for inv in rows if inv.project_name == args["project"]]
        if args.get("year") is not None:
            rows = [inv for inv in rows if inv.year == args["year"]]
        if args.get("paid_year") is not None:
            rows = [inv for inv in rows if inv.payment_date and inv.payment_date.year == args["paid_year"]]
        if args.get("country"):
            names_in_country = {p.project_name for p in projects if p.country == args["country"]}
            rows = [inv for inv in rows if inv.project_name in names_in_country]
        category = args.get("category")
        if category == "trial":
            rows = [inv for inv in rows if _is_paid_trial_category(inv)]
        elif category == "y1":
            rows = [inv for inv in rows if _is_new_installation_category(inv)]
        elif category == "y2plus":
            rows = [inv for inv in rows if _is_maintenance_category(inv)]
        if not rows:
            return "No invoices match that question.", None
        total = sum(float(inv.payment_amount) for inv in rows)
        label = "unpaid invoice" if tool_name == "get_unpaid_invoices" else "invoice"
        return f"Found {len(rows)} {label} row(s) totaling €{total:,.0f}.", _build_invoice_answer_df(rows, projects)

    if tool_name == "get_project":
        project = next((p for p in projects if p.project_name == args["project"]), None)
        if project is None:
            return f"I could not find project details for {args['project']}.", None
        df = pd.DataFrame([{
            "Project": project.project_name,
            "Country": project.country,
            "# Cams": project.num_cams,
            "Payment Month": project.payment_month,
            "Install Year": project.installation_year or "",
            "Status": project.status,
        }])
        return f"Here are the project details for {project.project_name}.", df

    if tool_name in ("get_projects", "count_projects"):
        rows = list(projects)
        status_arg = args.get("status")
        status_scope = ""
        if status_arg == "All":
            pass
        elif status_arg:
            rows = [p for p in rows if _normalize_project_status(p.status) == status_arg]
        else:
            # Default to active/installed only, matching the Dashboard's "Active Projects" KPI.
            rows = [p for p in rows if p.is_active()]
            status_scope = "active "
        if args.get("country"):
            rows = [p for p in rows if p.country == args["country"]]
        if args.get("month"):
            rows = get_projects_for_month(rows, args["month"])  # noqa: F821 - known latent bug preserved in Phase 1 (see module docstring)
        if args.get("min_cameras") is not None:
            rows = [p for p in rows if _safe_int(p.num_cams) > args["min_cameras"]]
        if tool_name == "count_projects":
            qualifiers = []
            if status_arg and status_arg != "All":
                qualifiers.append(f"with status {status_arg}")
            if args.get("country"):
                qualifiers.append(f"in {args['country']}")
            suffix = (" " + " ".join(qualifiers)) if qualifiers else ""
            return f"There are {len(rows)} {status_scope}project(s){suffix}.", None
        if not rows:
            return "No projects match that question.", None
        return f"Found {len(rows)} {status_scope}project(s).", _build_project_answer_df(rows)

    if tool_name == "get_debt":
        rows = _filter_invoices_for_debt_tools(invoices, projects, args, lambda inv: inv.is_unpaid())
        if not rows:
            return "No unpaid invoice rows match that question.", None
        total = sum(float(inv.payment_amount) for inv in rows)
        return f"I found {len(rows)} unpaid invoice row(s) totaling €{total:,.0f}.", _build_invoice_answer_df(rows, projects)

    if tool_name == "get_outstanding_amount":
        rows = _filter_invoices_for_debt_tools(invoices, projects, args, lambda inv: inv.is_unpaid())
        total = sum(float(inv.payment_amount) for inv in rows)
        subject = args.get("project") or args.get("country") or "the selected filters"
        return f"Outstanding (unpaid) amount for {subject} is €{total:,.0f} across {len(rows)} invoice row(s).", None

    if tool_name == "get_paid_amount":
        rows = _filter_invoices_for_debt_tools(invoices, projects, args, lambda inv: inv.is_paid())
        total = sum(float(inv.payment_amount) for inv in rows)
        subject = args.get("project") or args.get("country") or "the selected filters"
        # Also report the outstanding balance so "is X fully paid?" can never be answered with
        # just a paid-amount figure that silently omits money still owed (paid_year is dropped
        # here since it's a payment-date filter and unpaid invoices have no payment_date).
        outstanding_args = {k: v for k, v in args.items() if k != "paid_year"}
        outstanding_rows = _filter_invoices_for_debt_tools(invoices, projects, outstanding_args, lambda inv: inv.is_unpaid())
        outstanding_total = sum(float(inv.payment_amount) for inv in outstanding_rows)
        if not rows and not outstanding_rows:
            # Zero paid AND zero unpaid rows means no matching invoices exist at all — never
            # claim "fully paid" for a scope with no invoices to be paid in the first place.
            status_txt = " No paid or unpaid invoices were found for this scope, so payment status cannot be determined."
        elif outstanding_total <= 0:
            status_txt = " No outstanding balance remains — fully paid."
        else:
            status_txt = (
                f" €{outstanding_total:,.0f} across {len(outstanding_rows)} invoice row(s) is still "
                "outstanding — NOT fully paid."
            )
        return f"Paid amount for {subject} is €{total:,.0f} across {len(rows)} invoice row(s).{status_txt}", None

    if tool_name == "get_top_debt_projects":
        year = args.get("year")
        limit = max(1, args.get("limit") or 10)
        debt_by_project: dict[str, float] = {}
        for inv in invoices:
            if not inv.is_unpaid():
                continue
            if year is not None and inv.year != year:
                continue
            debt_by_project.setdefault(inv.project_name, 0.0)
            debt_by_project[inv.project_name] += float(inv.payment_amount)
        if not debt_by_project:
            return "No unpaid debt rows match that question.", None
        top = sorted(debt_by_project.items(), key=lambda item: item[1], reverse=True)[:limit]
        df = pd.DataFrame([{"Project": name, "Debt (€)": amount} for name, amount in top])
        return f"Here are the top {len(top)} debt project(s).", df

    if tool_name == "get_sent_invoices":
        rows = load_sent_invoices_log()
        if args.get("invoice_number") is not None:
            rows = [row for row in rows if _safe_int(row.get("invoice_number"), default=0) == args["invoice_number"]]
        if args.get("year") is not None:
            rows = [row for row in rows if _safe_int(row.get("year"), default=0) == args["year"]]
        if args.get("project"):
            rows = [row for row in rows if args["project"].lower() in _safe_str(row.get("subject", "")).lower()]
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

    if tool_name == "get_next_invoice_number":
        next_inv = next_invoice_number()
        return f"The next invoice number is {next_inv}.", None

    if tool_name == "get_last_bank_payment":
        from config.settings import load_bank_payments_log
        try:
            payments = load_bank_payments_log()
        except Exception as exc:
            return f"Could not load bank payments: {exc}", None
        if not payments:
            return "No bank payments have been recorded yet.", None

        try:
            from services.supabase_service import load_bank_payment_allocations
        except Exception:
            load_bank_payment_allocations = None

        def _allocations_for(payment):
            payment_id = payment.get("id")
            if not payment_id or load_bank_payment_allocations is None:
                return []
            try:
                return load_bank_payment_allocations(payment_id) or []
            except Exception:
                return []

        target_project = args.get("project")
        last_payment, allocations = None, []
        if target_project:
            for payment in payments:  # already newest-first
                candidate_allocations = _allocations_for(payment)
                names_in_payment = {a.get("project_name") for a in candidate_allocations if a.get("project_name")}
                if target_project in names_in_payment:
                    last_payment, allocations = payment, candidate_allocations
                    break
            if last_payment is None:
                return (
                    f"No recorded bank payment event was found for {target_project}. This audit log "
                    "may be incomplete and this result does not indicate whether the project is paid "
                    "or unpaid — ask about its payment/paid status separately to check that.",
                    None,
                )
        else:
            last_payment = payments[0]
            allocations = _allocations_for(last_payment)

        date_txt = _safe_str(last_payment.get("payment_date")) or "unknown date"
        if allocations:
            df = pd.DataFrame([
                {
                    "Invoice #": a.get("invoice_number"),
                    "Project": a.get("project_name"),
                    "Maint. Year": a.get("maintenance_year"),
                    "Amount (€)": a.get("amount_applied"),
                    "Year": a.get("year"),
                }
                for a in allocations
            ])
            total = sum(_safe_float(a.get("amount_applied")) for a in allocations)
            return f"Last bank payment ({date_txt}) covered {len(df)} invoice line(s) totaling €{total:,.0f}.", df

        total = _safe_float(last_payment.get("applied_amount")) or _safe_float(last_payment.get("received_amount"))
        inv_no = last_payment.get("invoice_number")
        subject = f"invoice #{inv_no}" if inv_no else (last_payment.get("source_name") or "an invoice")
        return f"Last bank payment ({date_txt}): {subject}, €{total:,.0f}.", None

    if tool_name == "get_licenses":
        entries = _filter_license_entries(projects, args)
        if not entries:
            return "No projects match that license question.", None
        df = pd.DataFrame([
            {
                "Project": p.project_name,
                "Country": p.country,
                "License EOP": eop.isoformat() if eop else "",
                "License Status": status,
                "Project Status": p.status,
            }
            for p, eop, status in entries
        ])
        status_counts: dict[str, int] = {}
        for _, _, status in entries:
            status_counts[status] = status_counts.get(status, 0) + 1
        summary = ", ".join(f"{count} {status}" for status, count in status_counts.items())
        return f"Found {len(entries)} project(s) ({summary}).", df

    if tool_name == "get_license_payment_status":
        entries = _filter_license_entries(projects, args)
        if not entries:
            return "No projects match that license question.", None
        today = datetime.date.today()
        rows = [_annual_license_payment_row(p, eop, invoices, today) for p, eop, _status in entries]
        df = pd.DataFrame(rows)
        status_counts: dict[str, int] = {}
        for row in rows:
            status_counts[row["Annual Payment Status"]] = status_counts.get(row["Annual Payment Status"], 0) + 1
        summary = ", ".join(f"{count} {status}" for status, count in status_counts.items())
        return f"Found {len(rows)} project(s) with latest annual payment status ({summary}).", df

    if tool_name == "get_license_eop_reconciliation":
        all_rows = build_license_reconciliation(projects, invoices, _normalize_project_status)
        rows = filter_reconciliation(
            all_rows,
            project=args.get("project"),
            country=args.get("country"),
            project_status=args.get("project_status"),
            paid_year=args.get("paid_year"),
            needs_update=args.get("needs_update"),
        )
        if not rows:
            return "No projects match that License EOP reconciliation question. (Read-only: no License EOP was changed.)", None
        need_rows = [r for r in rows if r["Needs EOP Update?"] == "Yes"]
        low_count = sum(1 for r in need_rows if r["Confidence"] == "Low")
        year_txt = f" with Year {args['paid_year']} paid" if args.get("paid_year") is not None else ""
        summary = (
            f"Found {len(rows)} project(s){year_txt}; {len(need_rows)} need a License EOP update"
            f" ({len(need_rows) - low_count} high confidence, {low_count} low confidence"
            f" - verify manually). Expected EOP = license start date + the highest consecutively paid "
            "maintenance year (Y1, Y2, ...) from the full invoice history. Read-only: no License EOP was changed."
        )
        df = pd.DataFrame([{col: r.get(col) for col in LICENSE_RECONCILIATION_COLUMNS} for r in rows])
        return summary, df

    if tool_name == "get_camera_statistics":
        rows = list(projects)
        status_arg = args.get("status")
        status_scope = ""
        if status_arg == "All":
            pass
        elif status_arg:
            rows = [p for p in rows if _normalize_project_status(p.status) == status_arg]
        else:
            # Default to active/installed only, matching the Dashboard's "Total Cameras" KPI.
            rows = [p for p in rows if p.is_active()]
            status_scope = "active "
        if args.get("country"):
            rows = [p for p in rows if p.country == args["country"]]
        camera_type = args.get("camera_type") or "all"
        aggregation = args.get("aggregation") or "sum"
        metric_label = "camera" if camera_type == "all" else camera_type

        if aggregation == "group_by_country":
            totals: dict[str, int] = {}
            for p in rows:
                key = p.country or "Unknown"
                totals[key] = totals.get(key, 0) + _camera_metric_for_project(p, camera_type)
            if not totals:
                return "No projects match that question.", None
            df = pd.DataFrame([
                {"Country": country, f"{metric_label.title()} Cameras": total}
                for country, total in sorted(totals.items(), key=lambda item: item[1], reverse=True)
            ])
            top_country, top_total = max(totals.items(), key=lambda item: item[1])
            return f"{top_country} has the most {metric_label} cameras ({top_total}, {status_scope or 'any status '}projects).", df

        if aggregation == "list_projects":
            if not rows:
                return "No projects match that question.", None
            df = pd.DataFrame([
                {
                    "Project": p.project_name,
                    "Country": p.country,
                    f"{metric_label.title()} Cameras": _camera_metric_for_project(p, camera_type),
                    "Status": p.status,
                }
                for p in rows
            ])
            return f"Found {len(rows)} {status_scope}project(s).", df

        total = sum(_camera_metric_for_project(p, camera_type) for p in rows)
        country_txt = f" in {args['country']}" if args.get("country") else ""
        if aggregation == "count_projects":
            return f"There are {len(rows)} {status_scope}project(s){country_txt}.", None
        if aggregation == "average_per_project":
            avg = (total / len(rows)) if rows else 0.0
            return f"Average {metric_label} cameras per {status_scope}project is {avg:.1f} (across {len(rows)} project(s)).", None
        return f"There are {total} {metric_label} camera(s) across {len(rows)} {status_scope}project(s){country_txt}.", None

    return "I could not execute that tool.", None


# Public aliases for later phases.
execute_tool = _execute_ask_data_tool
