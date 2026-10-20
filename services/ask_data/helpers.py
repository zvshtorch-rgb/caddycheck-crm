"""Pure helpers shared by Ask Data and the rest of the CRM app.

Moved verbatim from streamlit_app.py (Ask Data Phase 1). streamlit_app.py imports every name back
under the same name, so the rest of the app is unaffected. No Streamlit dependency, no I/O.
"""
import datetime
import math
import re
from typing import Optional

import pandas as pd

from services.ask_data.money import to_money
from services.ask_data.resolver import ProjectIndex

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


def _normalize_project_name_key(value: object) -> str:
    text = _safe_str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def _suggest_best_project_match(candidate: object, project_names: list[str]) -> tuple[str, float]:
    candidate_key = _normalize_project_name_key(candidate)
    if not candidate_key or not project_names:
        return "", 0.0

    from difflib import SequenceMatcher

    best_name = ""
    best_score = 0.0
    for project_name in project_names:
        project_key = _normalize_project_name_key(project_name)
        if not project_key:
            continue
        if candidate_key == project_key:
            return project_name, 1.0
        if candidate_key in project_key or project_key in candidate_key:
            score = 0.95
        else:
            score = SequenceMatcher(None, candidate_key, project_key).ratio()
        if score > best_score:
            best_score = score
            best_name = project_name
    return best_name, best_score


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
    if _invoice_category_label(invoice) == "credit":
        return True
    checker = getattr(invoice, "is_new_installation_category", None)
    if callable(checker):
        try:
            return bool(checker())
        except Exception:
            pass
    return _invoice_category_label(invoice) in {"y1", "credit"}


def _is_maintenance_category(invoice) -> bool:
    if _invoice_category_label(invoice) == "credit":
        return False
    checker = getattr(invoice, "is_maintenance_category", None)
    if callable(checker):
        try:
            return bool(checker())
        except Exception:
            pass
    return not _is_new_installation_category(invoice) and not _is_paid_trial_category(invoice)


ORDER_STATUS_OPTIONS = [
    "New",
    "Ordered",
    "In Progress",
    "Installed",
    "Active",
    "Cancelled",
]

PROJECT_STATUS_OPTIONS = ORDER_STATUS_OPTIONS[:1] + ["Trial"] + ORDER_STATUS_OPTIONS[1:-1] + ["Uninstalled", "Offline"] + ORDER_STATUS_OPTIONS[-1:]

DETECTION_CAMERA_FIELDS = [
    ("Backtray", "backtray_cameras", "Backtray Cams"),
    ("TopDown", "topdown_cameras", "TopDown Cams"),
    ("Pushout", "pushout_cameras", "Pushout Cams"),
]


def _normalize_query_text(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9+ ]", " ", _safe_str(value).lower())).strip()


def _normalize_project_status(value: str) -> str:
    cleaned = _safe_str(value).strip()
    if not cleaned:
        return "New"
    if cleaned.lower() == "offline":
        return "Offline"
    for option in PROJECT_STATUS_OPTIONS:
        if option.lower() == cleaned.lower():
            return option
    return cleaned


def _normalize_country(value: str) -> str:
    """Normalize country codes and names to full country names."""
    cleaned = _safe_str(value).strip()
    if not cleaned:
        return ""
    country_mapping = {
        "bel": "Belgium",
        "ger": "Germany",
        "hol": "Netherlands",
        "lux": "Luxembourg",
        "fra": "France",
        "ita": "Italy",
        "esp": "Spain",
        "gbr": "United Kingdom",
        "deu": "Germany",
        "nld": "Netherlands",
    }
    normalized_key = cleaned.lower()
    return country_mapping.get(normalized_key, cleaned)


def _project_license_date(project) -> Optional[datetime.date]:
    if not project.license_eop:
        return None
    if isinstance(project.license_eop, datetime.datetime):
        return project.license_eop.date()
    if isinstance(project.license_eop, datetime.date):
        return project.license_eop
    return None


def _license_status(project, today: Optional[datetime.date] = None) -> str:
    today = today or datetime.date.today()
    if _normalize_project_status(getattr(project, "status", "")).lower() == "cancelled":
        return "Cancelled"
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


def _normalize_detection_key(value: str) -> str:
    text = re.sub(r"[^a-z0-9]+", "", _safe_str(value).lower())
    aliases = {
        "back": "backtray_cameras",
        "backtray": "backtray_cameras",
        "top": "topdown_cameras",
        "topdown": "topdown_cameras",
        "push": "pushout_cameras",
        "pushout": "pushout_cameras",
    }
    return aliases.get(text, "")


def _parse_detection_camera_counts(detection_text: str) -> dict[str, int]:
    counts = {field_name: 0 for _, field_name, _ in DETECTION_CAMERA_FIELDS}
    text = _safe_str(detection_text)
    if not text.strip():
        return counts

    type_pattern = r"backtray|back|topdown|top\s*down|pushout|push\s*out"
    for match in re.finditer(rf"(\d+)\s*(?:cams?|cameras?)?\s*[-:]?\s*({type_pattern})", text, flags=re.IGNORECASE):
        field_name = _normalize_detection_key(match.group(2))
        if field_name:
            counts[field_name] += _safe_int(match.group(1), default=0)
    for match in re.finditer(rf"({type_pattern})\s*[-:]?\s*(\d+)", text, flags=re.IGNORECASE):
        field_name = _normalize_detection_key(match.group(1))
        if field_name and counts[field_name] == 0:
            counts[field_name] += _safe_int(match.group(2), default=0)
    return counts


def _project_detection_counts(project) -> dict[str, int]:
    counts = {
        field_name: _safe_int(getattr(project, field_name, 0), default=0)
        for _, field_name, _ in DETECTION_CAMERA_FIELDS
    }
    if any(counts.values()):
        return counts
    return _parse_detection_camera_counts(getattr(project, "detection_type", ""))


def _match_countries_from_question(question: str, projects, fuzzy_threshold: float = 0.8) -> list[str]:
    """
    Return every distinct project country mentioned in the question.

    Matches both the short DB code (e.g. 'Bel') and its full name (e.g.
    'Belgium'), and tolerates small typos in the full name (e.g. 'blegium')
    via a fuzzy ratio so users don't need exact spelling.
    """
    q = _normalize_query_text(question)
    if not q:
        return []
    words = q.split(" ")
    known_codes = sorted({p.country for p in projects if p.country})
    matched = []
    for code in known_codes:
        code_lower = code.lower()
        full_name = _normalize_country(code).lower()
        if code_lower in q or full_name in q:
            matched.append(code)
            continue
        if len(full_name) >= 4:
            from difflib import SequenceMatcher
            for word in words:
                if len(word) < 4:
                    continue
                if SequenceMatcher(None, word, full_name).ratio() >= fuzzy_threshold:
                    matched.append(code)
                    break
    return matched


def _build_invoice_answer_df(invoice_rows, projects) -> pd.DataFrame:
    project_index = ProjectIndex(projects)
    return pd.DataFrame([
        {
            "Invoice #": str(int(inv.invoice_number)) if inv.invoice_number else "—",
            "Project": inv.project_name,
            "Country": project_index.country(inv.project_name) or "",
            "Maint. Year": inv.maintenance_year,
            "Amount (€)": to_money(inv.payment_amount),
            "Paid": inv.paid,
            "Payment Date": inv.payment_date.date().isoformat() if inv.payment_date else "",
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
