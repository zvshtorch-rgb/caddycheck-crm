"""Ask Data v2 deterministic tool executor (read-only).

Phase 1 moved this verbatim from streamlit_app.py. Every number is computed here in plain Python from the
already-loaded CRM data; the LLM never sees or calculates any of it.

Phase 2 changes (everything else is unchanged, including float arithmetic and output formatting):
  * country and project filters match LEDGER rows to projects through the canonical `ProjectIndex`
    (case/accent/punctuation-insensitive + approved aliases) instead of exact project-name equality;
  * `get_projects_for_month` is now imported (it was missing, so month-filtered questions raised NameError).

Phase 1 deviation kept: `get_next_invoice_number` calls an injected zero-argument provider
(`next_invoice_number`) supplied by the app wrapper instead of reading streamlit_app globals.
"""
import datetime
from typing import Optional

import pandas as pd

from config.settings import MONTH_ORDER, load_sent_invoices_log
from services.ask_data import data_clock
from services.ask_data.breakdown import UNASSIGNED_LABEL, debt_by_country
from services.ask_data.money import ZERO, fmt_eur, fmt_percent, percent, sum_money, to_money
from services.ask_data.resolver import NetworkIndex, ProjectIndex
from services.ask_data.tool_result import Component, ToolResult, new_result_id
from services.excel_service import get_projects_for_month
from services.ask_data.helpers import (
    _build_invoice_answer_df,
    _build_project_answer_df,
    _is_maintenance_category,
    _is_new_installation_category,
    _is_paid_trial_category,
    _license_status,
    _normalize_country,
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
    index = ProjectIndex(projects) if (args.get("country") or args.get("project") or args.get("network")) else None
    if args.get("country"):
        rows = [inv for inv in rows if index.country(inv.project_name) == args["country"]]
    if args.get("project"):
        rows = [inv for inv in rows if index.canonical_name(inv.project_name) == args["project"]]
    if args.get("network"):
        network_index = NetworkIndex(projects)
        rows = [inv for inv in rows if _ledger_row_network(inv, index, network_index) == args["network"]]
    category = args.get("category")
    if category == "trial":
        rows = [inv for inv in rows if _is_paid_trial_category(inv)]
    elif category == "y1":
        rows = [inv for inv in rows if _is_new_installation_category(inv)]
    elif category == "y2plus":
        rows = [inv for inv in rows if _is_maintenance_category(inv)]
    return rows


def _ledger_row_network(inv, project_index: ProjectIndex, network_index: NetworkIndex):
    """A ledger row's network, via its already-resolved canonical project (never fuzzy; goal 1 + 2)."""
    canonical = project_index.canonical_name(inv.project_name)
    return network_index.network_of(canonical) if canonical else None


def _filter_license_entries(projects, args: dict) -> list:
    """Shared by get_licenses and get_license_payment_status: filter/sort (project, eop_date, status) triples."""
    today = datetime.date.today()
    entries = [(p, _project_license_date(p), _license_status(p, today)) for p in projects]
    if args.get("project"):
        entries = [e for e in entries if e[0].project_name == args["project"]]
    if args.get("network"):
        network_index = NetworkIndex(projects)
        entries = [e for e in entries if network_index.network_of(e[0].project_name) == args["network"]]
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
    hist_invoiced = sum_money(inv.payment_amount for inv in active_rows)
    hist_paid = sum_money(inv.payment_amount for inv in active_rows if inv.is_paid())
    hist_outstanding = sum_money(inv.payment_amount for inv in active_rows if inv.is_unpaid())

    annual_month = _annual_payment_month(project)
    target_year = _latest_annual_invoice_year(annual_month, today)

    invoice_number, period_text = "", ""
    invoice_amount, outstanding = ZERO, ZERO
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
            invoice_amount = sum_money(inv.payment_amount for inv in period_rows)
            outstanding = sum_money(inv.payment_amount for inv in period_rows if inv.is_unpaid())
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


def _charges_and_credits(rows, noun: str):
    """(calculation text, components, warnings) for a sum of ledger rows; unpaid credits are netted and called out."""
    total = sum_money(inv.payment_amount for inv in rows)
    credits = [inv for inv in rows if to_money(inv.payment_amount) < 0]
    if not credits:
        return (f"sum of {len(rows)} {noun} row(s) = {fmt_eur(total)}",
                [Component("Total", total, len(rows), tuple(rows))], [])
    charges = [inv for inv in rows if to_money(inv.payment_amount) >= 0]
    charges_total = sum_money(inv.payment_amount for inv in charges)
    credits_total = sum_money(inv.payment_amount for inv in credits)
    if not charges:
        return (f"{len(credits)} credit row(s) = {fmt_eur(total)}",
                [Component("Credits", credits_total, len(credits), tuple(credits))], [])
    calc = (f"{fmt_eur(charges_total)} charges ({len(charges)} row(s)) − {fmt_eur(abs(credits_total))} "
            f"credits ({len(credits)} row(s)) = {fmt_eur(total)}")
    return (calc,
            [Component("Charges", charges_total, len(charges), tuple(charges)),
             Component("Credits", credits_total, len(credits), tuple(credits))],
            [f"{len(credits)} credit row(s) totalling {fmt_eur(credits_total)} are netted against the charges."])


def _financial_result(tool_name: str, args: dict, projects, text: str, rows, *, noun: str = "invoice",
                      value=None, calculation=None, components=None, warnings=None, dataframe=None, derived=None) -> ToolResult:
    """Build a financial ToolResult whose value/calculation come from the retained ledger rows (all Decimal)."""
    rows = list(rows)
    if value is None:
        value = sum_money(inv.payment_amount for inv in rows)
    if calculation is None and components is None:
        calculation, components, extra = _charges_and_credits(rows, noun)
        warnings = list(warnings or []) + extra
    if dataframe is None and rows:
        dataframe = _build_invoice_answer_df(rows, projects)
    return ToolResult(
        tool=tool_name, text=text, dataframe=dataframe, financial=True, value=value, row_count=len(rows), rows=rows,
        components=list(components or []), calculation=calculation, derived_amounts=list(derived or []),
        warnings=list(warnings or []), frame_builder=lambda r: _build_invoice_answer_df(r, projects),
    )


def _country_label(code: str) -> str:
    full = _normalize_country(code)
    return code if full == code else f"{full} ({code})"


def _debt_by_country_result(tool_name: str, args: dict, projects, invoices) -> ToolResult:
    """Outstanding debt per country. Country buckets + 'Unassigned' always add up exactly to the total."""
    category = args.get("category")

    def predicate(inv):
        if not inv.is_unpaid():
            return False
        if category == "trial":
            return _is_paid_trial_category(inv)
        if category == "y1":
            return _is_new_installation_category(inv)
        if category == "y2plus":
            return _is_maintenance_category(inv)
        return True

    bd = debt_by_country(invoices, projects, row_filter=predicate, year=args.get("year"))
    total = bd.total
    ranked = sorted(bd.countries, key=lambda b: (-b.amount, b.label))
    buckets = ranked + ([bd.no_country] if bd.no_country.rows else []) + [bd.unassigned]  # Unassigned: always shown, even zero
    all_rows = [r for b in buckets for r in b.rows]
    by_code = {b.label: b for b in ranked}

    def share(bucket):
        return percent(bucket.amount, total)

    warnings = []
    if bd.unassigned.rows:
        names = ", ".join(bd.unassigned_names()[:5])
        warnings.append(f"{bd.unassigned.row_count} unpaid row(s) totalling {fmt_eur(bd.unassigned.amount)} match no project and "
                        f"are shown as '{UNASSIGNED_LABEL}' ({names}).")
    if bd.no_country.rows:
        warnings.append(f"{bd.no_country.row_count} unpaid row(s) totalling {fmt_eur(bd.no_country.amount)} belong to projects "
                        f"with no country set and are shown as '{bd.no_country.label}'.")

    records = [{"Country": _country_label(b.label) if index < len(ranked) else b.label, "Rows": b.row_count,
                "Amount (€)": b.amount, "Share of total (%)": share(b)} for index, b in enumerate(buckets)]
    records.append({"Country": "Total", "Rows": bd.row_count, "Amount (€)": total, "Share of total (%)": percent(total, total)})
    frame = pd.DataFrame(records)
    components = [Component(b.label, b.amount, b.row_count, tuple(b.rows)) for b in buckets]
    parts_calc = " + ".join(f"{b.label} {fmt_eur(b.amount)}" for b in buckets)

    if not all_rows:
        return ToolResult(
            tool=tool_name, financial=True, value=ZERO, text="There is no outstanding debt for the selected filters.",
            calculation=f"no unpaid rows = {fmt_eur(ZERO)}", warnings=warnings,
        )

    country_a, country_b = args.get("country"), args.get("compare_country")
    if country_a and country_b and country_a != country_b:
        a = by_code[country_a].amount if country_a in by_code else ZERO
        b = by_code[country_b].amount if country_b in by_code else ZERO
        other = total - a - b
        other_rows = [r for bucket in buckets if bucket.label not in (country_a, country_b) for r in bucket.rows]
        label_a, label_b = _country_label(country_a), _country_label(country_b)
        diff = abs(a - b)
        if a == b:
            text = f"{label_a} and {label_b} have the same outstanding debt: {fmt_eur(a)} each."
        else:
            (high_label, high), (low_label, low) = sorted(((label_a, a), (label_b, b)), key=lambda t: -t[1])
            pct_txt = fmt_percent(percent(diff, low)) if low > 0 else None
            text = (f"{label_a} has {fmt_eur(a)} and {label_b} has {fmt_eur(b)} of outstanding debt; {high_label} has {fmt_eur(diff)} more"
                    + (f" ({pct_txt} higher)." if pct_txt else f" ({low_label} has none)."))
        components = [
            Component(country_a, a, by_code[country_a].row_count if country_a in by_code else 0,
                      tuple(by_code[country_a].rows) if country_a in by_code else ()),
            Component(country_b, b, by_code[country_b].row_count if country_b in by_code else 0,
                      tuple(by_code[country_b].rows) if country_b in by_code else ()),
            Component("All other countries and unassigned", other, len(other_rows), tuple(other_rows)),
        ]
        calculation = f"{country_a} {fmt_eur(a)} + {country_b} {fmt_eur(b)} + other {fmt_eur(other)} = total {fmt_eur(total)}"
        return _financial_result(tool_name, args, projects, text, all_rows, value=total, calculation=calculation,
                                 components=components, warnings=warnings, dataframe=frame, derived=[diff, other])

    if country_a:
        bucket = by_code.get(country_a)
        amount = bucket.amount if bucket else ZERO
        count = bucket.row_count if bucket else 0
        pct_txt = fmt_percent(percent(amount, total))
        text = (f"{_country_label(country_a)} has {fmt_eur(amount)} of outstanding debt ({pct_txt} of the {fmt_eur(total)} total) "
                f"across {count} row(s).")
    else:
        text = f"Total outstanding debt is {fmt_eur(total)} across {bd.row_count} unpaid row(s)."
        if ranked:
            top = ranked[0]
            text += f" {_country_label(top.label)} has the most: {fmt_eur(top.amount)} ({fmt_percent(share(top))} of the total)."
    return _financial_result(tool_name, args, projects, text, all_rows, value=total,
                             calculation=f"{parts_calc} = {fmt_eur(total)}", components=components, warnings=warnings, dataframe=frame)


def _execute_ask_data_tool(
    tool_name: str, args: dict, projects, invoices, debt_summaries, *, next_invoice_number=None
) -> tuple[str, Optional[pd.DataFrame]]:
    """Backward-compatible tuple form: (display text, table). Prefer `execute_tool`, which returns the ToolResult."""
    result = execute_tool(tool_name, args, projects, invoices, debt_summaries, next_invoice_number=next_invoice_number)
    return result.display_text(), result.dataframe


def execute_tool(
    tool_name: str, args: dict, projects, invoices, debt_summaries, *, next_invoice_number=None
) -> ToolResult:
    """Execute one validated, read-only tool call and return its ToolResult (value, rows, calculation, filters,
    data timestamp, caveats, result_id). Everything is computed here from the loaded CRM data."""
    outcome = _dispatch_tool(tool_name, args, projects, invoices, debt_summaries, next_invoice_number=next_invoice_number)
    result = outcome if isinstance(outcome, ToolResult) else ToolResult(tool=tool_name, text=outcome[0], dataframe=outcome[1])
    result.filters = {k: v for k, v in args.items() if v is not None}
    result.result_id = result.result_id or new_result_id()
    if result.financial:
        result.data_as_of = data_clock.loaded_at()
        if result.data_as_of is None:
            result.warnings.append("The data load time is unknown, so no freshness timestamp can be shown.")
    return result


def _dispatch_tool(
    tool_name: str, args: dict, projects, invoices, debt_summaries, *, next_invoice_number=None
):
    """
    Execute one validated, read-only tool call against the already-loaded
    CRM data. All numbers are computed here in plain Python — the LLM never
    sees or calculates any of this. Returns a ToolResult (financial tools) or a (text, table) pair.
    """
    if tool_name == "get_invoice":
        invoice_number = args["invoice_number"]
        rows = [inv for inv in invoices if _safe_int(inv.invoice_number, default=0) == invoice_number]
        if args.get("year") is not None:
            rows = [inv for inv in rows if inv.year == args["year"]]
        if not rows:
            return f"Invoice {invoice_number} is not in the invoice ledger.", None
        total = sum_money(inv.payment_amount for inv in rows)
        if len(rows) == 1:
            paid_txt = "Cancelled" if rows[0].is_cancelled() else ("Paid" if rows[0].is_paid() else "Unpaid")
            summary = f"Invoice {invoice_number} totals {fmt_eur(total)} and is {paid_txt}."
        else:
            summary = (
                f"Invoice {invoice_number} has {len(rows)} row(s) totaling {fmt_eur(total)}; "
                "see the Paid column below for each row's status."
            )
        return _financial_result(tool_name, args, projects, summary, rows)

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
        index = ProjectIndex(projects) if (args.get("country") or args.get("project") or args.get("network")) else None
        if args.get("project"):
            rows = [inv for inv in rows if index.canonical_name(inv.project_name) == args["project"]]
        if args.get("year") is not None:
            rows = [inv for inv in rows if inv.year == args["year"]]
        if args.get("paid_year") is not None:
            rows = [inv for inv in rows if inv.payment_date and inv.payment_date.year == args["paid_year"]]
        if args.get("country"):
            rows = [inv for inv in rows if index.country(inv.project_name) == args["country"]]
        if args.get("network"):
            network_index = NetworkIndex(projects)
            rows = [inv for inv in rows if _ledger_row_network(inv, index, network_index) == args["network"]]
        category = args.get("category")
        if category == "trial":
            rows = [inv for inv in rows if _is_paid_trial_category(inv)]
        elif category == "y1":
            rows = [inv for inv in rows if _is_new_installation_category(inv)]
        elif category == "y2plus":
            rows = [inv for inv in rows if _is_maintenance_category(inv)]
        if not rows:
            return "No invoices match that question.", None
        total = sum_money(inv.payment_amount for inv in rows)
        label = "unpaid invoice" if tool_name == "get_unpaid_invoices" else "invoice"
        return _financial_result(tool_name, args, projects, f"Found {len(rows)} {label} row(s) totaling {fmt_eur(total)}.", rows, noun=label)

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
        if args.get("network"):
            network_index = NetworkIndex(rows)
            rows = [p for p in rows if network_index.network_of(p.project_name) == args["network"]]
        if args.get("month"):
            rows = get_projects_for_month(rows, args["month"])
        if args.get("min_cameras") is not None:
            rows = [p for p in rows if _safe_int(p.num_cams) > args["min_cameras"]]
        if tool_name == "count_projects":
            qualifiers = []
            if status_arg and status_arg != "All":
                qualifiers.append(f"with status {status_arg}")
            if args.get("network"):
                qualifiers.append(f"in network {args['network']}")
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
        total = sum_money(inv.payment_amount for inv in rows)
        return _financial_result(
            tool_name, args, projects, f"I found {len(rows)} unpaid invoice row(s) totaling {fmt_eur(total)}.", rows, noun="unpaid invoice")

    if tool_name == "get_outstanding_amount":
        rows = _filter_invoices_for_debt_tools(invoices, projects, args, lambda inv: inv.is_unpaid())
        total = sum_money(inv.payment_amount for inv in rows)
        subject = args.get("project") or args.get("network") or args.get("country") or "the selected filters"
        return _financial_result(
            tool_name, args, projects,
            f"Outstanding (unpaid) amount for {subject} is {fmt_eur(total)} across {len(rows)} invoice row(s).", rows, noun="unpaid invoice")

    if tool_name == "get_paid_amount":
        rows = _filter_invoices_for_debt_tools(invoices, projects, args, lambda inv: inv.is_paid())
        total = sum_money(inv.payment_amount for inv in rows)
        subject = args.get("project") or args.get("network") or args.get("country") or "the selected filters"
        # Also report the outstanding balance so "is X fully paid?" can never be answered with
        # just a paid-amount figure that silently omits money still owed (paid_year is dropped
        # here since it's a payment-date filter and unpaid invoices have no payment_date).
        outstanding_args = {k: v for k, v in args.items() if k != "paid_year"}
        outstanding_rows = _filter_invoices_for_debt_tools(invoices, projects, outstanding_args, lambda inv: inv.is_unpaid())
        outstanding_total = sum_money(inv.payment_amount for inv in outstanding_rows)
        if not rows and not outstanding_rows:
            # Zero paid AND zero unpaid rows means no matching invoices exist at all — never
            # claim "fully paid" for a scope with no invoices to be paid in the first place.
            status_txt = " No paid or unpaid invoices were found for this scope, so payment status cannot be determined."
        elif outstanding_total <= 0:
            status_txt = " No outstanding balance remains — fully paid."
        else:
            status_txt = (
                f" {fmt_eur(outstanding_total)} across {len(outstanding_rows)} invoice row(s) is still "
                "outstanding — NOT fully paid."
            )
        components = [
            Component("Paid", total, len(rows), tuple(rows)),
            Component("Still outstanding", outstanding_total, len(outstanding_rows), tuple(outstanding_rows)),
        ]
        calculation = (f"{fmt_eur(total)} paid ({len(rows)} row(s)); {fmt_eur(outstanding_total)} still outstanding "
                       f"({len(outstanding_rows)} row(s))")
        return _financial_result(
            tool_name, args, projects,
            f"Paid amount for {subject} is {fmt_eur(total)} across {len(rows)} invoice row(s).{status_txt}", rows,
            value=total, calculation=calculation, components=components)

    if tool_name == "get_top_debt_projects":
        year = args.get("year")
        limit = max(1, args.get("limit") or 10)
        rows_by_project: dict[str, list] = {}
        for inv in invoices:
            if not inv.is_unpaid():
                continue
            if year is not None and inv.year != year:
                continue
            rows_by_project.setdefault(inv.project_name, []).append(inv)
        if not rows_by_project:
            return "No unpaid debt rows match that question.", None
        debt_by_project = {name: sum_money(r.payment_amount for r in rs) for name, rs in rows_by_project.items()}
        top = sorted(debt_by_project.items(), key=lambda item: item[1], reverse=True)[:limit]
        df = pd.DataFrame([{"Project": name, "Debt (€)": amount} for name, amount in top])
        components = [Component(name, amount, len(rows_by_project[name]), tuple(rows_by_project[name])) for name, amount in top]
        shown_rows = [r for name, _amount in top for r in rows_by_project[name]]
        shown_total = sum_money(amount for _name, amount in top)
        return _financial_result(
            tool_name, args, projects, f"Here are the top {len(top)} debt project(s).", shown_rows, value=shown_total,
            calculation=f"unpaid rows summed per project; the {len(top)} project(s) shown total {fmt_eur(shown_total)}",
            components=components, dataframe=df)

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
        bank_caveat = "The bank-payment log may be incomplete; it is an event log, not the authoritative paid/unpaid status."
        if allocations:
            def _allocation_frame(items):
                return pd.DataFrame([
                    {
                        "Invoice #": a.get("invoice_number"),
                        "Project": a.get("project_name"),
                        "Maint. Year": a.get("maintenance_year"),
                        "Amount (€)": to_money(a.get("amount_applied")),
                        "Year": a.get("year"),
                    }
                    for a in items
                ])

            df = _allocation_frame(allocations)
            total = sum_money(a.get("amount_applied") for a in allocations)
            result = ToolResult(
                tool=tool_name, financial=True, value=total, dataframe=df, row_count=len(allocations), rows=list(allocations),
                text=f"Last bank payment ({date_txt}) covered {len(df)} invoice line(s) totaling {fmt_eur(total)}.",
                calculation=f"sum of {len(allocations)} allocation line(s) = {fmt_eur(total)}",
                components=[Component("Allocations", total, len(allocations), tuple(allocations))],
                warnings=[bank_caveat], frame_builder=_allocation_frame,
            )
            return result

        applied = to_money(last_payment.get("applied_amount"))
        total = applied if applied != 0 else to_money(last_payment.get("received_amount"))
        inv_no = last_payment.get("invoice_number")
        subject = f"invoice #{inv_no}" if inv_no else (last_payment.get("source_name") or "an invoice")
        return ToolResult(
            tool=tool_name, financial=True, value=total,
            text=f"Last bank payment ({date_txt}): {subject}, {fmt_eur(total)}.",
            calculation=f"recorded payment amount (no invoice allocations are logged for it) = {fmt_eur(total)}",
            warnings=[bank_caveat],
        )

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
        return ToolResult(
            tool=tool_name, financial=True, dataframe=df, row_count=len(rows),
            text=f"Found {len(rows)} project(s) with latest annual payment status ({summary}).",
        )

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
        if args.get("network"):
            network_index = NetworkIndex(rows)
            rows = [p for p in rows if network_index.network_of(p.project_name) == args["network"]]
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
        network_txt = f" in network {args['network']}" if args.get("network") else ""
        country_txt = f" in {args['country']}" if args.get("country") else ""
        if aggregation == "count_projects":
            return f"There are {len(rows)} {status_scope}project(s){network_txt}{country_txt}.", None
        if aggregation == "average_per_project":
            avg = (total / len(rows)) if rows else 0.0
            return f"Average {metric_label} cameras per {status_scope}project is {avg:.1f} (across {len(rows)} project(s)).", None
        return f"There are {total} {metric_label} camera(s) across {len(rows)} {status_scope}project(s){network_txt}{country_txt}.", None

    if tool_name == "debt_by_country":
        return _debt_by_country_result(tool_name, args, projects, invoices)

    return "I could not execute that tool.", None
