"""Ask Data v2 tool catalog: argument specs, per-tool descriptions and the catalog text shown to the planner.

Moved verbatim from streamlit_app.py (Ask Data Phase 1). Data only; every tool is read-only.
"""

# arg_name -> (kind, required). kind is one of:
#   "int", "project", "country", "status", "month", or "enum:a,b,c"
_ASK_DATA_TOOL_SPECS: dict[str, dict[str, tuple[str, bool]]] = {
    "get_invoice": {
        "invoice_number": ("int", True),
        "year": ("int", False),
    },
    "get_invoices": {
        "project": ("project", False),
        "year": ("int", False),
        "paid_year": ("int", False),
        "country": ("country", False),
        "paid": ("enum:yes,no,cancelled", False),
    },
    "get_unpaid_invoices": {
        "project": ("project", False),
        "year": ("int", False),
        "paid_year": ("int", False),
        "country": ("country", False),
        "category": ("enum:trial,y1,y2plus", False),
    },
    "get_project": {
        "project": ("project", True),
    },
    "get_projects": {
        "country": ("country", False),
        "status": ("status", False),
        "month": ("month", False),
        "min_cameras": ("int", False),
    },
    "count_projects": {
        "country": ("country", False),
        "status": ("status", False),
    },
    "get_debt": {
        "project": ("project", False),
        "year": ("int", False),
        "country": ("country", False),
        "category": ("enum:trial,y1,y2plus", False),
    },
    "get_paid_amount": {
        "project": ("project", False),
        "year": ("int", False),
        "paid_year": ("int", False),
        "country": ("country", False),
    },
    "get_outstanding_amount": {
        "project": ("project", False),
        "year": ("int", False),
        "country": ("country", False),
    },
    "get_top_debt_projects": {
        "year": ("int", False),
        "limit": ("int", False),
    },
    "get_sent_invoices": {
        "invoice_number": ("int", False),
        "year": ("int", False),
        "project": ("project", False),
    },
    "get_next_invoice_number": {},
    "get_camera_statistics": {
        "country": ("country", False),
        "status": ("status", False),
        "camera_type": ("enum:all,TopDown,Backtray,Pushout", False),
        "aggregation": ("enum:sum,count_projects,average_per_project,group_by_country,list_projects", False),
    },
    "get_last_bank_payment": {
        "project": ("project", False),
    },
    "get_licenses": {
        "project": ("project", False),
        "country": ("country", False),
        "license_status": ("enum:Active,Expired,Missing,Update Next Month,Cancelled", False),
        "project_status": ("status", False),
        "eop_year": ("int", False),
        "eop_month": ("month", False),
        "eop_date": ("date", False),
    },
    "get_license_payment_status": {
        "project": ("project", False),
        "country": ("country", False),
        "license_status": ("enum:Active,Expired,Missing,Update Next Month,Cancelled", False),
        "project_status": ("status", False),
        "eop_year": ("int", False),
        "eop_month": ("month", False),
        "eop_date": ("date", False),
    },
    "get_license_eop_reconciliation": {
        "project": ("project", False),
        "country": ("country", False),
        "project_status": ("status", False),
        "paid_year": ("int", False),
        "needs_update": ("enum:yes,no", False),
    },
}

# Per-tool description shown to the LLM alongside its argument list. This is the primary
# lever for disambiguating "payment event/history" (bank_payments/bank_payment_allocations —
# an audit log that may be incomplete) from "financial/payment status" (invoices table — the
# authoritative source for paid/unpaid/outstanding/debt). See `_llm_parse_data_question`'s
# prompt for the higher-level routing rules that reference these tools by name.
_ASK_DATA_TOOL_DESCRIPTIONS: dict[str, str] = {
    "get_invoice": (
        "Looks up a single invoice by number from the invoices table, the authoritative source "
        "for invoice-level financial status. Use this for 'is invoice N paid', 'show invoice N', "
        "'was invoice N sent' style questions."
    ),
    "get_invoices": (
        "Lists/filters invoice ledger rows (invoices table). Use this for actual payment status, "
        "paid/unpaid status, and invoice-level or project-level financial status questions. The "
        "invoices table is the authoritative source for these questions — do NOT use "
        "get_last_bank_payment for them."
    ),
    "get_unpaid_invoices": (
        "Lists unpaid invoice ledger rows (invoices table), filterable by project/country/year/"
        "category. Authoritative for 'which invoices/projects are unpaid or not fully paid' "
        "questions — do NOT use get_last_bank_payment for these."
    ),
    "get_project": "Looks up static details (country, cameras, payment month, status) for one project.",
    "get_projects": "Lists/filters projects by country, status, billing month, or minimum camera count.",
    "count_projects": "Counts projects matching a country/status filter.",
    "get_debt": (
        "Computes modeled/expected outstanding debt from the invoices table. Use this for debt, "
        "outstanding balance, and 'how much is owed' questions — the invoices table is the "
        "authoritative source for actual payment status, paid/unpaid status, outstanding "
        "balances, debt, and partial payments, not the bank-payment event log."
    ),
    "get_paid_amount": (
        "Sums amounts actually paid from the invoices table (authoritative paid-status source). "
        "Use this for 'has X paid', 'is X fully paid', 'how much has X paid' financial-status "
        "questions — NOT get_last_bank_payment, which only covers a possibly-incomplete log of "
        "individual transfer events, not overall paid/unpaid status."
    ),
    "get_outstanding_amount": (
        "Sums unpaid invoice amounts from the invoices table. Use this for 'what is still "
        "outstanding/owed for X' and 'how much is outstanding across all projects' questions — "
        "the invoices table is authoritative for outstanding balances, not the bank-payment log."
    ),
    "get_top_debt_projects": "Ranks projects by outstanding debt (invoices table), optionally for one year.",
    "get_sent_invoices": "Looks up the log of invoice PDFs actually emailed out (not payment status).",
    "get_next_invoice_number": "Returns the next unused invoice number.",
    "get_camera_statistics": "Aggregates camera counts (sum/average/by-country/etc.) across projects.",
    "get_last_bank_payment": (
        "Returns recorded bank-payment events and allocations from the bank_payments/"
        "bank_payment_allocations audit log (optionally for one project). Use ONLY for questions "
        "about a specific payment/transfer/remittance EVENT — when it happened, or what invoices/"
        "amounts it covered (e.g. 'when was the last bank payment for X', 'what was included in "
        "the last transfer', 'show the bank payment history'). This audit log may be INCOMPLETE "
        "(not every paid invoice has a logged event here) and must NOT be used to determine "
        "whether a project or invoice is paid, unpaid, fully paid, or has outstanding debt — use "
        "get_invoice/get_invoices/get_unpaid_invoices/get_debt/get_paid_amount/"
        "get_outstanding_amount/get_license_payment_status for those instead."
    ),
    "get_licenses": (
        "Filters projects by license_status (the computed license classification: Active, "
        "Expired, Missing, Update Next Month, Cancelled) and/or expiry date. 'license_status' is "
        "NOT the same as 'project_status' (the project's own operational status: Active, "
        "Offline, Cancelled, New, etc. -- the same values used by get_projects). For questions "
        "combining both concepts, e.g. 'expired licenses on Active/Offline projects', 'licenses "
        "that expired but the project is still active', set license_status='Expired' AND "
        "project_status accordingly -- never drop the 'expired' part of the question just "
        "because a project-status word like 'Active' is also mentioned. Not for payment/debt "
        "questions."
    ),
    "get_license_payment_status": (
        "Same license filtering as get_licenses (including the license_status vs project_status "
        "distinction -- see get_licenses), PLUS the authoritative paid/unpaid status of each "
        "project's latest already-due annual invoice, joined from the invoices table. Use this "
        "instead of get_licenses when a license question also asks about invoices, payments, "
        "debt, outstanding balance, or whether projects are paid/fully paid. It looks ONLY at the "
        "latest annual invoice -- for 'was the License EOP updated for the paid years' questions "
        "use get_license_eop_reconciliation instead."
    ),
    "get_license_eop_reconciliation": (
        "READ-ONLY reconciliation of every project's CURRENT License EOP against the EOP its full "
        "invoice/payment history justifies (Year 1, Year 2, ... all maintenance years from the invoices "
        "table, not just the latest invoice). Returns per project: license start date, current EOP, Y1/Y2 "
        "invoice numbers + paid status + payment dates, highest consecutively paid maintenance year, "
        "expected EOP (start date + that many years), 'Needs EOP Update?' and a confidence flag. Use it "
        "for: 'Year 1 paid but License EOP not updated', 'Year 2 paid but EOP not updated', 'paid licenses "
        "whose EOP is earlier than the payments justify', 'which licenses need their EOP updated', "
        "'expected license EOP'. Set paid_year=1 or 2 (etc.) for 'Year N paid but EOP not updated'; set "
        "needs_update='yes' for 'EOP earlier than justified by payments'. It never changes any date."
    ),
}


def _ask_data_tool_catalog_text() -> str:
    lines = []
    for tool_name, spec in _ASK_DATA_TOOL_SPECS.items():
        if not spec:
            signature = f"- {tool_name}(): no arguments"
        else:
            parts = []
            for arg_name, (kind, required) in spec.items():
                marker = "required" if required else "optional"
                if kind.startswith("enum:"):
                    parts.append(f"{arg_name} ({marker}, one of: {kind.split(':', 1)[1]})")
                else:
                    parts.append(f"{arg_name} ({marker}, {kind})")
            signature = f"- {tool_name}({', '.join(parts)})"
        description = _ASK_DATA_TOOL_DESCRIPTIONS.get(tool_name)
        lines.append(f"{signature}\n  {description}" if description else signature)
    return "\n".join(lines)


# Public aliases for later phases (the underscore names are kept so moved code stays verbatim).
TOOL_SPECS = _ASK_DATA_TOOL_SPECS
TOOL_DESCRIPTIONS = _ASK_DATA_TOOL_DESCRIPTIONS
tool_catalog_text = _ask_data_tool_catalog_text
