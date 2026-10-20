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


# ── Ask Data Phase 5: historical/trend tools ─────────────────────────────────────────────────────────
# Only ever reachable when the `ask_data_history` feature flag is on (conversation.py gates this before the
# planner/executor are even invoked for a historical question). All actual arithmetic lives in
# services/ask_data/history.py (pure, no I/O); this module only does the lazy, try/except-wrapped Supabase
# reads, mirroring the existing get_last_bank_payment pattern above -- a missing/not-yet-migrated table
# degrades to a clear message, never a crash.

def _load_snapshot_runs():
    from services.ask_data import history
    try:
        from services.supabase_service import load_invoice_snapshot_runs
    except Exception as exc:
        return None, f"Historical snapshot data is not available: {exc}"
    try:
        raw_runs = load_invoice_snapshot_runs()
    except Exception as exc:
        return None, f"Could not load invoice snapshot history: {exc}"
    return history.parse_snapshot_runs(raw_runs), None


def _snapshot_day_loader():
    from services.supabase_service import load_invoice_snapshot
    return load_invoice_snapshot


def _lookup_debt_at(requested, runs, projects, args):
    """(lookup, filtered_resolved_rows, total, warnings) for one historical date, or (None, ...) with a
    message if history does not reach back that far."""
    from services.ask_data import history
    earliest = history.earliest_available_date(runs)
    if earliest is None:
        return None, None, ZERO, ["No invoice snapshot history has been captured yet."]
    if requested < earliest:
        return None, None, ZERO, [
            f"History is available from {earliest.isoformat()} onward; {requested.isoformat()} predates that, "
            "so I can't reconstruct the debt for that date.",
        ]
    lookup = history.resolve_snapshot(runs, requested, _snapshot_day_loader())
    if not lookup.available:
        return None, None, ZERO, [f"No snapshot history is available on or before {requested.isoformat()}."]
    resolved = history.resolve_snapshot_rows(lookup.rows, projects)
    filtered = history.apply_filters(resolved, projects, args)
    total = history.debt_total(filtered)
    warnings = history.mismatch_warnings(filtered)
    coverage = lookup.coverage_note()
    if coverage:
        warnings = [coverage] + warnings
    return lookup, filtered, total, warnings


def _captured_label(lookup) -> str:
    when = lookup.captured_at.strftime("%Y-%m-%d %H:%M UTC") if lookup.captured_at else None
    return f"snapshot date {lookup.actual_date.isoformat()}" + (f", captured {when}" if when else "")


_MISSING_DATE_TEXT = (
    "I need a date (or date range) to answer that -- e.g. 'today', 'last week', 'this month', 'since "
    "October 1', or an explicit date like '30 September'. Ask again with one of those, or ask the "
    "original question again if you expected this to carry over from an earlier turn."
)


def _get_debt_at_result(tool_name: str, args: dict, projects) -> ToolResult:
    if args.get("at_date") is None:
        return ToolResult(tool=tool_name, text=_MISSING_DATE_TEXT)
    runs, error = _load_snapshot_runs()
    if error:
        return ToolResult(tool=tool_name, text=error)
    requested = args["at_date"]
    lookup, filtered, total, warnings = _lookup_debt_at(requested, runs, projects, args)
    if lookup is None:
        return ToolResult(tool=tool_name, text=warnings[0], warnings=warnings[1:])
    rows = [r.row for r in filtered if _row_is_unpaid(r.row)]
    text = f"Outstanding debt as of {requested.isoformat()} was {fmt_eur(total)} ({_captured_label(lookup)})."
    return ToolResult(
        tool=tool_name, financial=True, value=total, text=text, row_count=len(rows), rows=rows,
        calculation=f"sum of {len(rows)} unpaid snapshot row(s) = {fmt_eur(total)}",
        components=[Component("Outstanding", total, len(rows), tuple(rows))], warnings=warnings,
    )


def _row_is_unpaid(row: dict) -> bool:
    return str(row.get("paid") or "").strip().lower() == "no"


def _compare_debt_dates_result(tool_name: str, args: dict, projects) -> ToolResult:
    if args.get("date_a") is None or args.get("date_b") is None:
        return ToolResult(tool=tool_name, text=_MISSING_DATE_TEXT)
    runs, error = _load_snapshot_runs()
    if error:
        return ToolResult(tool=tool_name, text=error)
    date_a, date_b = args["date_a"], args["date_b"]
    lookup_a, filtered_a, total_a, warn_a = _lookup_debt_at(date_a, runs, projects, args)
    lookup_b, filtered_b, total_b, warn_b = _lookup_debt_at(date_b, runs, projects, args)
    if lookup_a is None or lookup_b is None:
        return ToolResult(tool=tool_name, text=(warn_a + warn_b)[0], warnings=(warn_a + warn_b)[1:])
    rows_a = [r.row for r in filtered_a if _row_is_unpaid(r.row)]
    rows_b = [r.row for r in filtered_b if _row_is_unpaid(r.row)]
    diff = total_b - total_a
    direction = "an increase" if diff > 0 else ("a decrease" if diff < 0 else "no change")
    text = (
        f"Debt on {date_a.isoformat()} was {fmt_eur(total_a)} ({_captured_label(lookup_a)}); "
        f"on {date_b.isoformat()} it was {fmt_eur(total_b)} ({_captured_label(lookup_b)}) -- {direction} of "
        f"{fmt_eur(abs(diff))}."
    )
    frame = pd.DataFrame([
        {"Date": date_a.isoformat(), "Debt (€)": total_a, "Snapshot": _captured_label(lookup_a)},
        {"Date": date_b.isoformat(), "Debt (€)": total_b, "Snapshot": _captured_label(lookup_b)},
    ])
    return ToolResult(
        tool=tool_name, financial=True, value=diff, text=text, dataframe=frame,
        row_count=len(rows_a) + len(rows_b), rows=rows_a + rows_b,
        calculation=f"{fmt_eur(total_b)} - {fmt_eur(total_a)} = {fmt_eur(diff)}",
        components=[Component(date_a.isoformat(), total_a, len(rows_a), tuple(rows_a)),
                   Component(date_b.isoformat(), total_b, len(rows_b), tuple(rows_b))],
        derived_amounts=[diff], warnings=list(warn_a) + list(warn_b),
    )


def _resolve_period_snapshots(period, runs, projects, args, invoices):
    """(opening_lookup, opening_filtered, closing_lookup, closing_filtered, live_closing, warnings)."""
    from services.ask_data import history
    import datetime as _dt
    opening_lookup, opening_filtered, _opening_total, warnings = _lookup_debt_at(period.date_from, runs, projects, args)
    if opening_lookup is None:
        return None, None, None, None, None, warnings
    closing_lookup, closing_filtered, closing_total, closing_warnings = _lookup_debt_at(period.date_to, runs, projects, args)
    if closing_lookup is None:
        return None, None, None, None, None, list(warnings) + list(closing_warnings)
    all_warnings = list(warnings) + list(closing_warnings)
    live_closing = None
    if period.date_to == _dt.date.today():
        # Maximum accuracy for a period ending today: use the LIVE current debt total (same code as
        # get_debt()), not yesterday's snapshot, as the closing figure.
        resolved_live = history.resolve_snapshot_rows(
            [{"invoice_row_id": f"live:{inv.db_id}", "project_name": inv.project_name, "payment_amount": inv.payment_amount,
              "paid": inv.paid, "year": inv.year, "maintenance_year": inv.maintenance_year} for inv in invoices],
            projects,
        )
        live_closing = history.debt_total(history.apply_filters(resolved_live, projects, args))
    return opening_lookup, opening_filtered, closing_lookup, closing_filtered, live_closing, all_warnings


def _get_debt_change_result(tool_name: str, args: dict, projects, invoices) -> ToolResult:
    if args.get("period") is None:
        return ToolResult(tool=tool_name, text=_MISSING_DATE_TEXT)
    from services.ask_data import history
    runs, error = _load_snapshot_runs()
    if error:
        return ToolResult(tool=tool_name, text=error)
    period = args["period"]
    opening_lookup, opening_filtered, closing_lookup, closing_filtered, live_closing, warnings = (
        _resolve_period_snapshots(period, runs, projects, args, invoices))
    if opening_lookup is None:
        return ToolResult(tool=tool_name, text=warnings[0], warnings=warnings[1:])
    bridge = history.compute_bridge(opening_filtered, closing_filtered, live_closing=live_closing)
    change_log, change_log_error = _load_change_log(period.date_from, period.date_to)
    if change_log is not None:
        warnings = warnings + history.classify_with_change_log(bridge, change_log)
    elif change_log_error:
        warnings = warnings + [change_log_error]
    warnings = warnings + bridge.warnings
    closing_label = "live current data" if live_closing is not None else _captured_label(closing_lookup)
    text = (
        f"Debt change for {period.label} ({period.date_from.isoformat()} to {period.date_to.isoformat()}): "
        f"{fmt_eur(bridge.change)}.\n\n"
        f"Opening debt ({period.date_from.isoformat()}, {_captured_label(opening_lookup)}): {fmt_eur(bridge.opening)}\n"
        f"New invoices: {fmt_eur(bridge.new_invoices)}\n"
        f"Payments: -{fmt_eur(bridge.payments)}\n"
        f"Credit notes: -{fmt_eur(bridge.credit_notes)}\n"
        f"Cancellations: -{fmt_eur(bridge.cancellations)}\n"
        + (f"Removed rows: -{fmt_eur(bridge.removed)}\n" if bridge.removed != ZERO else "")
        + (f"Other adjustments: {fmt_eur(bridge.adjustments)}\n" if bridge.adjustments != ZERO else "")
        + f"Closing debt ({period.date_to.isoformat()}, {closing_label}): {fmt_eur(bridge.closing)}"
    )
    all_rows = (bridge.new_invoices_rows + bridge.payments_rows + bridge.credit_notes_rows
                + bridge.cancellations_rows + bridge.adjustments_rows + bridge.removed_rows)
    components = [
        Component("Opening", bridge.opening, 0, ()),
        Component("New invoices", bridge.new_invoices, len(bridge.new_invoices_rows), tuple(bridge.new_invoices_rows)),
        Component("Payments", -bridge.payments, len(bridge.payments_rows), tuple(bridge.payments_rows)),
        Component("Credit notes", bridge.credit_notes, len(bridge.credit_notes_rows), tuple(bridge.credit_notes_rows)),
        Component("Cancellations", -bridge.cancellations, len(bridge.cancellations_rows), tuple(bridge.cancellations_rows)),
        Component("Closing", bridge.closing, 0, ()),
    ]
    return ToolResult(
        tool=tool_name, financial=True, value=bridge.change, text=text, row_count=len(all_rows), rows=all_rows,
        calculation=(f"{fmt_eur(bridge.opening)} + {fmt_eur(bridge.new_invoices)} - {fmt_eur(bridge.payments)} - "
                    f"{fmt_eur(bridge.credit_notes)} - {fmt_eur(bridge.cancellations)} = {fmt_eur(bridge.closing)}"),
        components=components, derived_amounts=[bridge.change, bridge.unexplained], warnings=warnings,
    )


def _load_change_log(date_from, date_to):
    try:
        from services.supabase_service import load_invoice_change_log
    except Exception as exc:
        return None, f"invoice_change_log is not available: {exc}"
    try:
        import datetime as _dt
        return load_invoice_change_log(_dt.datetime.combine(date_from, _dt.time.min),
                                       _dt.datetime.combine(date_to, _dt.time.max)), None
    except Exception as exc:
        return None, f"Could not load invoice_change_log: {exc}"


def _get_invoice_changes_result(tool_name: str, args: dict, projects, invoices) -> ToolResult:
    if args.get("period") is None:
        return ToolResult(tool=tool_name, text=_MISSING_DATE_TEXT)
    from services.ask_data import history
    runs, error = _load_snapshot_runs()
    if error:
        return ToolResult(tool=tool_name, text=error)
    period = args["period"]
    opening_lookup, opening_filtered, closing_lookup, closing_filtered, live_closing, warnings = (
        _resolve_period_snapshots(period, runs, projects, args, invoices))
    if opening_lookup is None:
        return ToolResult(tool=tool_name, text=warnings[0], warnings=warnings[1:])
    bridge = history.compute_bridge(opening_filtered, closing_filtered, live_closing=live_closing)
    warnings = warnings + bridge.warnings
    records = []
    for action, rows in (("New", bridge.new_invoices_rows), ("Payment", bridge.payments_rows),
                        ("Credit note", bridge.credit_notes_rows), ("Cancellation", bridge.cancellations_rows),
                        ("Amount/status adjustment", bridge.adjustments_rows), ("Removed", bridge.removed_rows)):
        for row in rows:
            records.append({"Action": action, "Invoice #": row.get("invoice_number"), "Project": row.get("project_name"),
                           "Amount (€)": to_money(row.get("payment_amount")), "Year": row.get("year")})
    all_rows = (bridge.new_invoices_rows + bridge.payments_rows + bridge.credit_notes_rows
                + bridge.cancellations_rows + bridge.adjustments_rows + bridge.removed_rows)
    if not records:
        return ToolResult(
            tool=tool_name, financial=True, value=ZERO,
            text=f"No invoice changes were found for {period.label} ({period.date_from.isoformat()} to {period.date_to.isoformat()}).",
            warnings=warnings,
        )
    frame = pd.DataFrame(records)
    text = (f"Found {len(records)} invoice change(s) for {period.label} "
            f"({period.date_from.isoformat()} to {period.date_to.isoformat()}, {_captured_label(opening_lookup)} -> "
            f"{'live current data' if live_closing is not None else _captured_label(closing_lookup)}).")
    return ToolResult(
        tool=tool_name, financial=True, value=bridge.new_invoices - bridge.payments - bridge.credit_notes - bridge.cancellations,
        text=text, dataframe=frame, row_count=len(all_rows), rows=all_rows,
        calculation=f"{len(records)} classified row(s) across New/Payment/Credit note/Cancellation/Adjustment/Removed",
        components=[Component("New invoices", bridge.new_invoices, len(bridge.new_invoices_rows), tuple(bridge.new_invoices_rows)),
                   Component("Payments", bridge.payments, len(bridge.payments_rows), tuple(bridge.payments_rows)),
                   Component("Cancellations", bridge.cancellations, len(bridge.cancellations_rows), tuple(bridge.cancellations_rows))],
        warnings=warnings,
    )


def _get_payments_in_period_result(tool_name: str, args: dict, projects, invoices) -> ToolResult:
    if args.get("period") is None:
        return ToolResult(tool=tool_name, text=_MISSING_DATE_TEXT)
    from services.ask_data import history
    period = args["period"]
    try:
        from services.supabase_service import load_bank_payments, load_bank_payment_allocations
        bank_payments = load_bank_payments()
        allocations = load_bank_payment_allocations()
    except Exception as exc:
        return ToolResult(tool=tool_name, text=f"Could not load bank payment history: {exc}")
    project_filter = args.get("project")
    total, matched_allocations = history.payments_from_bank_log(
        bank_payments, allocations, period.date_from, period.date_to, project_filter=project_filter)
    warnings = ["Sourced from the bank_payments/bank_payment_allocations audit log, which may be incomplete "
               "(not every paid invoice has a logged bank-payment event)."]
    runs, run_error = _load_snapshot_runs()
    if runs:
        opening_lookup, opening_filtered, closing_lookup, closing_filtered, live_closing, snap_warnings = (
            _resolve_period_snapshots(period, runs, projects, args, invoices))
        if opening_lookup is not None:
            bridge = history.compute_bridge(opening_filtered, closing_filtered, live_closing=live_closing)
            if to_money(bridge.payments) != to_money(total):
                warnings.append(
                    f"Snapshot-based cross-check shows {fmt_eur(bridge.payments)} in rows that flipped to paid "
                    f"this period, which differs from the {fmt_eur(total)} recorded in the bank-payment log -- "
                    "shown as a discrepancy, not silently resolved.")
    if not matched_allocations and total == ZERO:
        text = f"No recorded bank payments were found for {period.label} ({period.date_from.isoformat()} to {period.date_to.isoformat()})."
        return ToolResult(tool=tool_name, financial=True, value=ZERO, text=text, warnings=warnings)
    frame = pd.DataFrame([
        {"Invoice #": a.get("invoice_number"), "Project": a.get("project_name"),
         "Amount (€)": to_money(a.get("amount_applied")), "Year": a.get("year")}
        for a in matched_allocations
    ]) if matched_allocations else None
    text = (f"Payments received for {period.label} ({period.date_from.isoformat()} to {period.date_to.isoformat()}): "
            f"{fmt_eur(total)} across {len(matched_allocations)} allocation line(s).")
    return ToolResult(
        tool=tool_name, financial=True, value=total, text=text, dataframe=frame,
        row_count=len(matched_allocations), rows=matched_allocations,
        calculation=f"sum of {len(matched_allocations)} bank-payment allocation line(s) = {fmt_eur(total)}",
        components=[Component("Payments", total, len(matched_allocations), tuple(matched_allocations))],
        warnings=warnings,
    )


def _get_new_invoices_in_period_result(tool_name: str, args: dict, projects, invoices) -> ToolResult:
    if args.get("period") is None:
        return ToolResult(tool=tool_name, text=_MISSING_DATE_TEXT)
    from services.ask_data import history
    runs, error = _load_snapshot_runs()
    if error:
        return ToolResult(tool=tool_name, text=error)
    period = args["period"]
    opening_lookup, opening_filtered, closing_lookup, closing_filtered, live_closing, warnings = (
        _resolve_period_snapshots(period, runs, projects, args, invoices))
    if opening_lookup is None:
        return ToolResult(tool=tool_name, text=warnings[0], warnings=warnings[1:])
    bridge = history.compute_bridge(opening_filtered, closing_filtered, live_closing=live_closing)
    new_rows = bridge.new_invoices_rows + bridge.credit_notes_rows
    warnings = warnings + bridge.warnings
    try:
        from config.settings import load_sent_invoices_log
        sent = load_sent_invoices_log() or []
        sent_in_period = [s for s in sent if s.get("sent_at") and period.date_from.isoformat() <= str(s["sent_at"])[:10] <= period.date_to.isoformat()]
        if len(sent_in_period) != len(new_rows):
            warnings.append(
                f"The sent_invoices log shows {len(sent_in_period)} invoice(s) sent in this period, which differs "
                f"from the {len(new_rows)} new ledger row(s) found -- shown as a secondary signal only, not "
                "authoritative.")
    except Exception:
        pass
    if not new_rows:
        text = f"No new invoices were added for {period.label} ({period.date_from.isoformat()} to {period.date_to.isoformat()})."
        return ToolResult(tool=tool_name, financial=True, value=ZERO, text=text, warnings=warnings)
    total = sum_money(to_money(r.get("payment_amount")) for r in new_rows)
    frame = pd.DataFrame([
        {"Invoice #": r.get("invoice_number"), "Project": r.get("project_name"),
         "Amount (€)": to_money(r.get("payment_amount")), "Year": r.get("year")}
        for r in new_rows
    ])
    text = (f"{len(new_rows)} new invoice row(s) were added for {period.label} "
            f"({period.date_from.isoformat()} to {period.date_to.isoformat()}), totalling {fmt_eur(total)}.")
    return ToolResult(
        tool=tool_name, financial=True, value=total, text=text, dataframe=frame, row_count=len(new_rows), rows=new_rows,
        calculation=f"sum of {len(new_rows)} new snapshot row(s) = {fmt_eur(total)}",
        components=[Component("New invoices", total, len(new_rows), tuple(new_rows))], warnings=warnings,
    )


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

    if tool_name == "get_debt_at":
        return _get_debt_at_result(tool_name, args, projects)

    if tool_name == "compare_debt_dates":
        return _compare_debt_dates_result(tool_name, args, projects)

    if tool_name == "get_debt_change":
        return _get_debt_change_result(tool_name, args, projects, invoices)

    if tool_name == "get_invoice_changes":
        return _get_invoice_changes_result(tool_name, args, projects, invoices)

    if tool_name == "get_payments_in_period":
        return _get_payments_in_period_result(tool_name, args, projects, invoices)

    if tool_name == "get_new_invoices_in_period":
        return _get_new_invoices_in_period_result(tool_name, args, projects, invoices)

    return "I could not execute that tool.", None
