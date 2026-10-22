"""Read-only billing evidence and draft validation. No persistence or send operations."""
import datetime as dt
from types import SimpleNamespace

from services.ask_data.resolver import NetworkIndex, ProjectIndex
from services.supabase_service import _select_all_rows


EVIDENCE_GROUPS = {
    "Historically evidenced": {
        8681: (43, 93, 9, 80, 44, 42, 73, 74, 18991, 79),
        8688: (39, 19230, 100, 126),
    },
    "Possible operational candidates": {
        8677: (13, 15, 16, 17, 54, 59, 101, 121, 19316),
        8678: (12, 45, 46, 47, 48, 49, 50, 51, 52, 53, 57, 83, 96),
        8713: (38, 109, 128),
        8714: (5, 7, 41, 85, 88, 89, 18793),
        8717: (6, 36, 87, 102, 105, 111, 113, 122, 129),
    },
    "Disputed historical evidence": {
        8671: (77, 81, 99, 104),
        8672: (69, 70, 108, 123, 125, 132),
    },
}
DECISIONS = ("Needs confirmation", "Approve for future billing", "Historical only", "Reject mapping")


def _number(row):
    value = str(row.get("invoice_number") or "").removesuffix(".0")
    return int(value) if value.isdigit() else 0


def build_review(projects, invoices, accounts, mappings):
    objects = [SimpleNamespace(**p) for p in projects]
    index, networks = ProjectIndex(objects), NetworkIndex(objects)
    by_project = {p["id"]: [] for p in projects}
    for invoice in invoices:
        project = index.project(invoice.get("project_name"))
        if project:
            by_project[project.id].append(invoice)
    projects_by_id = {p["id"]: p for p in projects}
    account_names = {a["id"]: a["account_code"] for a in accounts}
    groups = {}
    for group, evidence in EVIDENCE_GROUPS.items():
        rows = []
        for number, ids in evidence.items():
            for project_id in ids:
                if project_id not in projects_by_id:
                    raise ValueError(f"Evidence project {project_id} is missing; review the evidence roster.")
                p = projects_by_id[project_id]
                ledger = by_project[project_id]
                recurring = [
                    i for i in ledger if i.get("invoice_type") == "Monthly"
                    or str(i.get("description") or "").lower().startswith("maintenance monthly-")
                    or _number(i) in (8671, 8672, 8677, 8678, 8681, 8688)
                ]
                key = lambda i: (int(i.get("year") or 0), _number(i), i["id"])
                latest = max(ledger, key=key) if ledger else {}
                last_recurring = max(recurring, key=key) if recurring else {}
                assigned = [m for m in mappings if m["project_id"] == project_id]
                reference = f"Invoice {number} / " + (
                    "reviewed Iretailcheck-addressed PDF" if group == "Historically evidenced"
                    else "disputed PDF: displays 8669; reconcile identity/amount/year"
                    if group == "Disputed historical evidence"
                    else "monthly ledger only; legal payer unverified"
                )
                rows.append({
                    "project_id": project_id, "project_name": p["project_name"],
                    "network": networks.network_of(p["project_name"]) or "Unknown",
                    "country": p.get("country") or "Unknown",
                    "payment_month": p.get("payment_month") or "Unknown",
                    "status": p.get("status") or "Unknown",
                    "historical_invoice_evidence": reference,
                    "latest_maintenance_year": latest.get("maintenance_year") or "Unknown",
                    "latest_recurring_invoice": (
                        f"{last_recurring.get('invoice_number')} / {last_recurring.get('year')} / "
                        f"{last_recurring.get('for_month') or last_recurring.get('description') or 'period unknown'}"
                    ) if last_recurring else "None identified",
                    "license_expiry": p.get("license_eop") or "Unknown",
                    "mapping_status": "; ".join(
                        f"{account_names.get(m['billing_account_id'], 'Unknown account')} "
                        f"[{m['effective_from']}, {m['effective_to'] or 'open-ended'})"
                        for m in assigned
                    ) or "Unassigned",
                    "proposed_effective_from": "Unconfirmed",
                    "proposed_effective_to": "Unconfirmed (not open-ended authorization)",
                    "source_reference": reference,
                    "recommendation": (
                        "BLOCKED: AD Aalter maintenance-year inconsistency must be reconciled"
                        if project_id == 43 else "NEEDS MANUAL CONFIRMATION"
                    ),
                    "approval_blocked": project_id == 43,
                })
        groups[group] = rows
    return {"groups": groups, "accounts": accounts, "mappings": mappings}


def load_review():
    return build_review(
        _select_all_rows("projects"), _select_all_rows("invoices"),
        _select_all_rows("billing_accounts"), _select_all_rows("project_billing_accounts"),
    )


def validate_mapping_draft(row, account, mappings, *, decision, start, end, open_ended,
                           actor, approved_at, source_reference):
    errors = []
    if decision not in DECISIONS:
        errors.append("Select a supported decision.")
    if not actor.strip() or not source_reference.strip():
        errors.append("Approving identity and source reference / business justification are required.")
    if not isinstance(approved_at, dt.datetime) or approved_at.utcoffset() is None:
        errors.append("A timezone-aware approval timestamp is required.")
    if decision == "Approve for future billing":
        if row["approval_blocked"]:
            errors.append("AD Aalter maintenance-year evidence is inconsistent; reconcile it before approval.")
        if not account:
            errors.append("A billing account must be selected.")
        valid_start = type(start) is dt.date
        valid_end = type(end) is dt.date
        if not valid_start:
            errors.append("Enter an explicit effective_from date.")
        if open_ended and end is not None:
            errors.append("Choose an end date OR open-ended authorization, not both.")
        if not open_ended and not valid_end:
            errors.append("Enter effective_to or explicitly authorize an open-ended period.")
        if valid_start and valid_end and end <= start:
            errors.append("effective_to must be later than effective_from.")
        if valid_start and (valid_end or (open_ended and end is None)):
            for mapping in mappings:
                if mapping["project_id"] != row["project_id"]:
                    continue
                old_start = dt.date.fromisoformat(mapping["effective_from"])
                old_end = dt.date.fromisoformat(mapping["effective_to"]) if mapping["effective_to"] else dt.date.max
                if start < old_end and old_start < (end or dt.date.max):
                    errors.append("Proposed dates overlap an existing billing-account assignment.")
                    break
    return errors


def validate_account_draft(account, confirmations, reference, actor):
    errors = []
    for field in ("legal_name", "address_line_1", "city", "country_code", "vat_number", "recipient_config_key"):
        if not str(account.get(field) or "").strip():
            errors.append(f"Missing account field: {field}.")
    if not all(confirmations.get(k) is True for k in ("legal_name", "address", "vat", "recipients")):
        errors.append("Explicitly confirm legal name, address, VAT, and recipient configuration.")
    if not reference.strip() or not actor.strip():
        errors.append("Verification reference and authenticated verifying identity are required.")
    return errors
