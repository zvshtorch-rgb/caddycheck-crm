"""Application settings and constants."""
import os
import json
import logging
from pathlib import Path

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
                defaults.update(json.load(f))
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
    """Persist email configuration to file."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(EMAIL_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)
