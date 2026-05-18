"""Application settings and constants."""
import os
import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# Base directory of the project
BASE_DIR = Path(__file__).parent.parent

# Data paths
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
CONFIG_DIR = BASE_DIR / "config"

PROJECTS_FILE = DATA_DIR / "CaddyCheckProjectsInfo.xlsx"
INVOICE_TEMPLATE = DATA_DIR / "CC_M_inv_8669_Dec_2025.xlsx"
EMAIL_CONFIG_FILE = CONFIG_DIR / "email_config.json"
OVERRIDES_FILE = CONFIG_DIR / "project_overrides.json"
DATA_PATHS_FILE = CONFIG_DIR / "data_paths.json"
SENT_INVOICES_LOG_FILE = CONFIG_DIR / "sent_invoices_log.json"
LICENSE_CHANGE_LOG_FILE = CONFIG_DIR / "license_change_log.json"
ORDERS_FILE = CONFIG_DIR / "orders.json"

# Sheet names
SHEET_PROJECTS_OVERVIEW = "Projects overview"
SHEET_INVOICE_DETAILS = "Invoice details"
SHEET_PROJECT_PAYMENT_SUMMARY = "Project Payment Summary"
SHEET_YEARLY_PAYMENT_SUMMARY = "Yearly Payment Summary"

# Revenue rules
RATE_Y1_PER_CAM = 778       # Year 1 maintenance rate per camera
RATE_Y2_PLUS_PER_CAM = 228  # Year 2+ maintenance rate per camera

# Invoice constants
INVOICE_COMPANY_NAME = "Video Inform Ltd"
INVOICE_COMPANY_REG = "Company Registration No.: 514046077"
INVOICE_BILL_TO_NAME = "Iretailcheck"
INVOICE_BILL_TO_ADDRESS_1 = "Bijkhoevelaan 11"
INVOICE_BILL_TO_ADDRESS_2 = "2110 WIJNEGEM, Belgium"
INVOICE_BILL_TO_VAT = "VAT BE 0537.905.085"
INVOICE_PAYMENT_DAYS = 30   # Days until payment due

INVOICE_BANK_DETAILS = {
    "account_name": "Video Inform LTD",
    "account_no": "100700/13",
    "bank_name": "Bank Leumi Le-Israel Ltd.",
    "swift": "LUMIILITTLV",
    "branch": "709",
    "iban": "IBAN IL530107090000010070013",
    "website": "http://www.video-inform.com/",
    "address": "100 Jabotinsky st, P.O 3331 Petah Tikva 4959257 Israel",
    "contact": "P +972-3-6448968  F +972-3-9210190, info@video-inform.com",
    "signature_name": "Yoram Sagher",
}

# Logging
LOG_LEVEL = logging.DEBUG
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"

# Month name normalization map
MONTH_ALIASES = {
    "jan": "January", "feb": "February", "mar": "March",
    "apr": "April", "may": "May", "jun": "June",
    "jul": "July", "july": "July", "aug": "August",
    "sep": "September", "oct": "October", "nov": "November", "dec": "December",
}

MONTH_ORDER = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def normalize_month(month_str: str) -> str:
    """Normalize month abbreviation or name to full month name."""
    if not month_str:
        return ""
    key = str(month_str).strip().lower()[:4].rstrip(".")
    # Try 3-char prefix first, then 4-char
    result = MONTH_ALIASES.get(key[:3])
    if result:
        return result
    result = MONTH_ALIASES.get(key)
    if result:
        return result
    # Try full name match
    for full in MONTH_ORDER:
        if full.lower().startswith(key):
            return full
    return month_str.strip().title()


def get_data_paths() -> dict:
    """
    Return resolved file paths for data files and output directory.

    Reads from data_paths.json if it exists (set by the user via Settings),
    otherwise falls back to the default local ./data/ directory.

    Returns a dict with Path values:
      projects_file, invoice_template, output_dir
    """
    defaults = {
        "projects_file":    str(PROJECTS_FILE),
        "invoice_template": str(INVOICE_TEMPLATE),
        "output_dir":       str(OUTPUT_DIR),
    }
    if DATA_PATHS_FILE.exists():
        try:
            with open(DATA_PATHS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                defaults.update(saved)
        except Exception:
            pass
    return {
        "projects_file":    Path(defaults["projects_file"]),
        "invoice_template": Path(defaults["invoice_template"]),
        "output_dir":       Path(defaults["output_dir"]),
    }


def save_data_paths(paths: dict) -> None:
    """Persist file path configuration (projects file, template, output dir)."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(DATA_PATHS_FILE, "w", encoding="utf-8") as f:
        json.dump({k: str(v) for k, v in paths.items()}, f, indent=2)


def get_email_config() -> dict:
    """Load email configuration from st.secrets (cloud) or local file, with defaults."""
    defaults = {
        "smtp_host": "smtp.gmail.com",
        "smtp_port": 587,
        "smtp_use_tls": True,
        "smtp_username": "",
        "smtp_password": "",
        "sender_name": "CaddyCheck CRM",
        "sender_email": "",
        "default_recipients": [],
        "default_cc": [],
        "default_subject_template": "Monthly Invoice - {month} {year}",
        "default_body_template": (
            "Dear Team,\n\n"
            "Please find attached the monthly maintenance invoice for {month} {year}.\n\n"
            "Best regards,\n"
            "Video Inform Ltd"
        ),
    }
    # 1. Load from local file (local dev / Settings page saves here)
    if EMAIL_CONFIG_FILE.exists():
        try:
            with open(EMAIL_CONFIG_FILE, "r", encoding="utf-8") as f:
                file_config = json.load(f)
                file_config.pop("smtp_password", None)
                defaults.update(file_config)
        except Exception:
            pass
    # 2. Override with st.secrets [email] section if available (Streamlit Cloud)
    try:
        import streamlit as st
        email_secrets = st.secrets.get("email", {})
        if email_secrets:
            for k, v in email_secrets.items():
                # Convert comma-separated strings to lists for recipient fields
                if k in ("default_recipients", "default_cc") and isinstance(v, str):
                    v = [x.strip() for x in v.split(",") if x.strip()]
                defaults[k] = v
        session_password = st.session_state.get("_smtp_password_override", "")
        if session_password:
            defaults["smtp_password"] = session_password
    except Exception:
        pass
    return defaults


def get_project_overrides() -> dict:
    """
    Load per-project rate overrides from file.

    Returns dict keyed by project_name (lowercase) with keys:
      y1_rate, y2_rate  (floats, optional)
    """
    if OVERRIDES_FILE.exists():
        try:
            with open(OVERRIDES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_project_overrides(overrides: dict) -> None:
    """Persist per-project rate overrides."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(OVERRIDES_FILE, "w", encoding="utf-8") as f:
        json.dump(overrides, f, indent=2)


def save_email_config(config: dict) -> None:
    """Persist email configuration to file without storing SMTP passwords."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    safe_config = dict(config)
    safe_config.pop("smtp_password", None)
    with open(EMAIL_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(safe_config, f, indent=2)


def _is_missing_supabase_table_error(exc: Exception, table_name: str) -> bool:
    message = str(exc).lower()
    return table_name.lower() in message and (
        "could not find the table" in message
        or "does not exist" in message
        or "42p01" in message
    )


def _load_local_sent_invoices_log() -> list:
    if SENT_INVOICES_LOG_FILE.exists():
        try:
            with open(SENT_INVOICES_LOG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            pass
    return []


def _append_local_sent_invoice_log(entry: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    entries = _load_local_sent_invoices_log()
    entries.append(entry)
    with open(SENT_INVOICES_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)


def _save_local_sent_invoices_log(entries: list) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(SENT_INVOICES_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)


def _load_local_license_change_log() -> list:
    if LICENSE_CHANGE_LOG_FILE.exists():
        try:
            with open(LICENSE_CHANGE_LOG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            pass
    return []


def _append_local_license_change_log(entry: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    entries = _load_local_license_change_log()
    entries.append(entry)
    with open(LICENSE_CHANGE_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)


def _save_local_license_change_log(entries: list) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(LICENSE_CHANGE_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)


def load_sent_invoices_log() -> list:
    """Load sent invoice email history, preferring Supabase when available."""
    local_entries = _load_local_sent_invoices_log()
    try:
        from services.supabase_service import load_sent_invoices, save_sent_invoices

        remote_entries = load_sent_invoices()
        if not remote_entries and local_entries:
            save_sent_invoices(local_entries)
            remote_entries = load_sent_invoices()
        return remote_entries or local_entries
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            logger.warning("Falling back to local sent invoice log: %s", exc)
    except Exception as exc:
        if not _is_missing_supabase_table_error(exc, "sent_invoices"):
            logger.warning("Falling back to local sent invoice log: %s", exc)
    return local_entries


def append_sent_invoice_log(entry: dict) -> None:
    """Append a sent invoice email record, preferring Supabase and keeping a local backup."""
    try:
        from services.supabase_service import append_sent_invoice

        append_sent_invoice(entry)
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            logger.warning("Could not append sent invoice log to Supabase: %s", exc)
    except Exception as exc:
        if not _is_missing_supabase_table_error(exc, "sent_invoices"):
            logger.warning("Could not append sent invoice log to Supabase: %s", exc)
    _append_local_sent_invoice_log(entry)


def save_sent_invoices_log(entries: list) -> None:
    """Replace the sent invoice email history, preferring Supabase and keeping a local backup."""
    try:
        from services.supabase_service import save_sent_invoices

        save_sent_invoices(entries)
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            logger.warning("Could not save sent invoice log to Supabase: %s", exc)
    except Exception as exc:
        if not _is_missing_supabase_table_error(exc, "sent_invoices"):
            logger.warning("Could not save sent invoice log to Supabase: %s", exc)
    _save_local_sent_invoices_log(entries)


def load_license_change_log() -> list:
    """Load license change history, preferring Supabase when available."""
    local_entries = _load_local_license_change_log()
    try:
        from services.supabase_service import load_license_change_log as load_remote_license_change_log
        from services.supabase_service import save_license_change_log as save_remote_license_change_log

        remote_entries = load_remote_license_change_log()
        if not remote_entries and local_entries:
            save_remote_license_change_log(local_entries)
            remote_entries = local_entries
        return remote_entries or local_entries
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            logger.warning("Falling back to local license change log: %s", exc)
    except Exception as exc:
        if not _is_missing_supabase_table_error(exc, "license_change_log"):
            logger.warning("Falling back to local license change log: %s", exc)
    return local_entries


def append_license_change_log(entry: dict) -> None:
    """Append a license update record, preferring Supabase and keeping a local backup."""
    try:
        from services.supabase_service import append_license_change_log as append_remote_license_change_log

        append_remote_license_change_log(entry)
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            logger.warning("Could not append license change log to Supabase: %s", exc)
    except Exception as exc:
        if not _is_missing_supabase_table_error(exc, "license_change_log"):
            logger.warning("Could not append license change log to Supabase: %s", exc)
    _append_local_license_change_log(entry)


def save_license_change_log(entries: list) -> None:
    """Replace the license update history, preferring Supabase and keeping a local backup."""
    try:
        from services.supabase_service import save_license_change_log as save_remote_license_change_log

        save_remote_license_change_log(entries)
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            logger.warning("Could not save license change log to Supabase: %s", exc)
    except Exception as exc:
        if not _is_missing_supabase_table_error(exc, "license_change_log"):
            logger.warning("Could not save license change log to Supabase: %s", exc)
    _save_local_license_change_log(entries)


def load_orders_records() -> list:
    """Load locally stored order records."""
    if ORDERS_FILE.exists():
        try:
            with open(ORDERS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except Exception:
            pass
    return []


def save_orders_records(entries: list) -> None:
    """Replace the local order records file."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(ORDERS_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, indent=2)
