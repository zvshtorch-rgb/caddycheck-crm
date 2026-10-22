"""SELECT-only annual-maintenance candidate preview for project 93 / July 2027."""
import datetime as dt
import json
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from types import SimpleNamespace

from config import settings
from models.project import Project
from services.ask_data.resolver import NetworkIndex, ProjectIndex
from services.supabase_service import _select_all_rows


PROJECT_ID = 93
PERIOD = dt.date(2027, 7, 1)
PERIOD_END = dt.date(2027, 8, 1)
PLANNED_DATE = dt.date(2027, 6, 1)


def _positive_integer(value):
    try:
        number = Decimal(str(value))
        return int(number) if number.is_finite() and number > 0 and number == number.to_integral_value() else None
    except (ValueError, InvalidOperation):
        return None


def _period_claims(row, *, sent=False):
    """Read explicit billing metadata, never payment/issue/send timestamps."""
    claims = set()
    year = _positive_integer(row.get("year"))
    month = row.get("month" if sent else "for_month")
    if month:
        text = str(month).strip()
        iso = re.fullmatch(r"(\d{4})-(\d{2})(?:-01)?", text)
        if iso and 1 <= int(iso[2]) <= 12:
            claims.add((int(iso[1]), int(iso[2])))
        elif year and settings.normalize_month(text) in settings.MONTH_ORDER:
            claims.add((year, settings.MONTH_ORDER.index(settings.normalize_month(text)) + 1))
    for match in re.finditer(r"Maintenance Monthly[- ]([A-Za-z]+)[- ](\d{4})",
                             str(row.get("description") or ""), re.I):
        normalized = settings.normalize_month(match[1])
        if normalized in settings.MONTH_ORDER:
            claims.add((int(match[2]), settings.MONTH_ORDER.index(normalized) + 1))
    return claims


def _history_check(project, projects, invoices, sent_history, maintenance_year):
    errors, checked = [], []
    index = ProjectIndex([SimpleNamespace(**p) for p in projects])
    target = (PERIOD.year, PERIOD.month)
    for row in invoices:
        resolved = index.project(row.get("project_name"))
        if not resolved and row.get("project_id") != PROJECT_ID:
            if target in _period_claims(row) or _positive_integer(row.get("year")) == PERIOD.year:
                errors.append(f"Unassigned ledger identity in the target year/period: invoice {row.get('invoice_number')}.")
            continue
        if row.get("project_id") != PROJECT_ID and (not resolved or resolved.id != PROJECT_ID):
            continue
        claims = _period_claims(row)
        year = _positive_integer(row.get("year"))
        label = str(row.get("maintenance_year") or "").strip().upper()
        kind = str(row.get("invoice_type") or "").strip().lower()
        checked.append(f"ledger:{row.get('id')} / invoice {row.get('invoice_number')}")
        if target in claims:
            errors.append(f"Existing billing-period ledger reference: invoice {row.get('invoice_number')}.")
        elif maintenance_year and label == f"Y{maintenance_year}":
            errors.append(f"Maintenance-year ledger reference needs review: invoice {row.get('invoice_number')}.")
        elif len(claims) > 1 or (claims and year and any(y != year for y, _ in claims)):
            errors.append(f"Conflicting ledger period metadata: invoice {row.get('invoice_number')}.")
        elif not claims and (year == PERIOD.year or (
                not year and (not label or label == f"Y{maintenance_year}"))):
            if kind not in ("complementary", "paid trial", "trial"):
                errors.append(f"Ambiguous recurring coverage: invoice {row.get('invoice_number')}.")
    reference = project.get(f"m{maintenance_year}y") if maintenance_year else None
    if reference not in (None, "", 0, "0"):
        errors.append(f"Maintenance Y{maintenance_year} already references invoice {reference}; review coverage.")
    for row in sent_history:
        claims = _period_claims(row, sent=True)
        year = _positive_integer(row.get("year"))
        # Historical sent rows have no reliable project membership. A batch for
        # this period, or one whose period is unknown, cannot prove exclusion.
        if target in claims or len(claims) > 1 or (not claims and year in (None, PERIOD.year)):
            errors.append(f"Sent-history coverage requires review: invoice {row.get('invoice_number')}.")
        checked.append(f"sent:{row.get('id', 'local')} / invoice {row.get('invoice_number')}")
    return errors, checked


def build_candidate(projects, accounts, mappings, invoices, sent_history, overrides, email_config):
    """Canonical money values are decimal strings; no invoice number is proposed or reserved."""
    errors = []
    candidate = {
        "preview_only": True, "billing_account_id": None, "billing_account": None,
        "project_id": PROJECT_ID, "project_name": None, "network": None,
        "billing_period": "2027-07", "billing_period_start": PERIOD.isoformat(),
        "billing_period_end_exclusive": PERIOD_END.isoformat(),
        "invoice_type": "annual_maintenance", "planned_send_date": PLANNED_DATE.isoformat(),
        "planned_invoice_date": PLANNED_DATE.isoformat(), "maintenance_year": None,
        "camera_quantity": None, "unit_price": None, "pricing_source": None,
        "total_amount": None, "currency": "EUR", "recipient_configuration_source": None,
        "recipients": {"to": [], "cc": []}, "history_checked": [],
    }
    matches = [p for p in projects if p["id"] == PROJECT_ID]
    if len(matches) != 1 or matches[0]["project_name"] != "AD Aartselar":
        errors.append("Canary project identity is missing or changed.")
        return candidate | {"validation": {"state": "blocked", "errors": errors}}
    project = matches[0]
    if ProjectIndex([SimpleNamespace(**p) for p in projects]).project(project["project_name"]) is None:
        errors.append("Project name is ambiguous in legacy financial history.")
    candidate["project_name"] = project["project_name"]
    candidate["network"] = NetworkIndex([SimpleNamespace(**p) for p in projects]).network_of(project["project_name"])
    if str(project.get("status") or "").lower() not in ("active", "installed"):
        errors.append("Project is not active/installed.")
    if settings.normalize_month(project.get("payment_month") or "") != "July":
        errors.append("Project billing month is not July.")
    applicable = []
    for mapping in mappings:
        if mapping["project_id"] != PROJECT_ID:
            continue
        try:
            start = dt.date.fromisoformat(mapping["effective_from"])
            end = dt.date.fromisoformat(mapping["effective_to"]) if mapping["effective_to"] else dt.date.max
        except (ValueError, TypeError):
            errors.append("Invalid billing mapping dates.")
            continue
        if start < PERIOD_END and PERIOD < end:
            applicable.append(mapping)
    if len(applicable) != 1:
        errors.append("Billing mapping missing or ambiguous for the billing month.")
    else:
        mapping = applicable[0]
        if not all(mapping.get(k) for k in ("approved_by", "approved_at", "source_reference")):
            errors.append("Billing mapping approval evidence is incomplete.")
        if (mapping["effective_from"] > PERIOD.isoformat()
                or (mapping["effective_to"] and mapping["effective_to"] < PERIOD_END.isoformat())):
            errors.append("Mapping does not cover the complete billing month.")
        account_matches = [a for a in accounts if a["id"] == mapping["billing_account_id"]]
        if len(account_matches) != 1:
            errors.append("Mapped billing account is missing or ambiguous.")
        else:
            account = account_matches[0]
            candidate["billing_account_id"] = account["id"]
            fields = ("account_code", "legal_name", "address_line_1", "address_line_2",
                      "postal_code", "city", "country_code", "vat_number")
            candidate["billing_account"] = {k: account.get(k) for k in fields}
            if account.get("account_code") != "IRETAILCHECK_BE":
                errors.append("Mapped account is outside this canary.")
            if account.get("active") is not True or not all(account.get(k) for k in (
                    "verified_by", "verified_at", "verification_reference")):
                errors.append("Billing account is inactive or unverified.")
            if not all(account.get(k) for k in ("legal_name", "address_line_1", "city", "country_code", "vat_number")):
                errors.append("Billing account legal identity is incomplete.")
            config_key = account.get("recipient_config_key")
            candidate["recipient_configuration_source"] = f"billing_accounts:{account['id']} / [{config_key}]"
            if config_key != "email":
                errors.append("Unsupported recipient configuration.")
            else:
                for source, target in (("default_recipients", "to"), ("default_cc", "cc")):
                    values = email_config.get(source, [])
                    if not isinstance(values, list) or any(
                        not isinstance(v, str) or not re.fullmatch(r"[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+", v)
                        for v in values
                    ):
                        errors.append(f"Invalid recipient configuration: {source}.")
                    else:
                        candidate["recipients"][target] = list(values)
                if not candidate["recipients"]["to"]:
                    errors.append("Recipient To configuration is empty.")
    quantity = _positive_integer(project.get("num_cameras"))
    installed = _positive_integer(project.get("installation_year"))
    candidate["camera_quantity"] = quantity
    if quantity is None:
        errors.append("Camera quantity is missing or invalid.")
    maintenance_year = None
    if installed is None or installed >= PERIOD.year:
        errors.append("Recurring maintenance year is unresolved (installation year required, Y2+ only).")
    else:
        model = Project(project["project_name"], installation_year=installed)
        maintenance_year = model.get_maintenance_year(PERIOD.year)
        candidate["maintenance_year"] = f"Y{maintenance_year}"
        # Match load_projects precedence: local project override entry overrides DB rates,
        # including an explicit None falling back to the global rate.
        name = settings.canonical_project_name(project["project_name"]).lower().strip()
        local = overrides.get(name)
        rate = project.get("rate_y2_override")
        source = "projects.rate_y2_override"
        if local is not None:
            if not isinstance(local, dict):
                errors.append("Project pricing override configuration is invalid.")
                rate = "invalid"
            else:
                rate = local.get("y2_rate")
                source = "project_overrides.y2_rate"
        if rate is None:
            source = "config.settings.RATE_Y2_PLUS_PER_CAM"
        try:
            if rate is not None:
                parsed = Decimal(str(rate))
                if not parsed.is_finite() or parsed <= 0:
                    raise ValueError("Nonpositive/nonfinite rate")
                model.rate_y2_override = float(parsed)
            unit = Decimal(str(model.get_rate(PERIOD.year)))
            if not unit.is_finite() or unit <= 0 or unit != unit.quantize(Decimal("0.01")):
                raise ValueError("Invalid EUR unit-price precision")
            candidate["unit_price"] = f"{unit:.2f}"
            candidate["pricing_source"] = source
            if quantity:
                candidate["total_amount"] = f"{(unit * quantity).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):.2f}"
        except (ValueError, InvalidOperation, TypeError):
            errors.append("Pricing is unresolved or invalid; no default substituted.")
    history_errors, checked = _history_check(project, projects, invoices, sent_history, maintenance_year)
    errors.extend(history_errors)
    candidate["history_checked"] = checked
    candidate["validation"] = {"state": "blocked" if errors else "eligible_preview", "errors": errors}
    return candidate


def load_candidate():
    """Use raw SELECTs, not legacy history loaders which may migrate/write records."""
    overrides = {}
    if settings.OVERRIDES_FILE.exists():
        overrides = json.loads(settings.OVERRIDES_FILE.read_text(encoding="utf-8"))
        if not isinstance(overrides, dict):
            raise ValueError("Pricing overrides must be an object.")
    sent = _select_all_rows("sent_invoices")
    # Local history can contain sends absent from the ledger/remote history.
    if settings.SENT_INVOICES_LOG_FILE.exists():
        local = json.loads(settings.SENT_INVOICES_LOG_FILE.read_text(encoding="utf-8"))
        if not isinstance(local, list) or any(not isinstance(r, dict) for r in local):
            raise ValueError("Local sent history must be a list of records.")
        sent += local
    return build_candidate(
        _select_all_rows("projects"), _select_all_rows("billing_accounts"),
        _select_all_rows("project_billing_accounts"), _select_all_rows("invoices"),
        sent, overrides, settings.get_email_config(),
    )
