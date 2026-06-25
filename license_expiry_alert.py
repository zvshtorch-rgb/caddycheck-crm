"""Send email alerts for licenses expiring soon.

Standalone scheduled job. It does not depend on Streamlit runtime.

Examples:
    py license_expiry_alert.py --dry-run
    py license_expiry_alert.py
    py license_expiry_alert.py --days-before 3 --to ops@example.com
"""
from __future__ import annotations

import argparse
import datetime as dt
import logging

from config.settings import get_email_config
from services.email_service import send_simple_email
from services.excel_service import load_projects as load_projects_excel
from services.supabase_service import (
    append_license_expiry_alert_log,
    has_license_expiry_alert_sent,
    load_projects as load_projects_supabase,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("license_expiry_alert")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Email configured recipients when licenses expire soon.",
    )
    parser.add_argument(
        "--days-before",
        type=int,
        default=3,
        help="Alert when License EOP is this many days away. Default: 3.",
    )
    parser.add_argument(
        "--source",
        choices=["auto", "supabase", "excel"],
        default="auto",
        help="Data source to use. Default: auto.",
    )
    parser.add_argument("--to", help="Override recipients with a comma-separated list.")
    parser.add_argument("--cc", help="Override CC recipients with a comma-separated list.")
    parser.add_argument("--subject", help="Override email subject.")
    parser.add_argument("--dry-run", action="store_true", help="Preview without sending email or writing alert log.")
    parser.add_argument("--force", action="store_true", help="Send even if this alert was already logged.")
    return parser.parse_args()


def _split_csv(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def _load_projects(source: str) -> tuple[list, str]:
    if source == "excel":
        return load_projects_excel(), "Excel (local fallback)"
    if source == "supabase":
        return load_projects_supabase(), "Supabase"

    try:
        return load_projects_supabase(), "Supabase"
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            raise
        logger.info("Supabase not configured. Falling back to Excel data files.")
        return load_projects_excel(), "Excel (local fallback)"


def _project_license_date(project) -> dt.date | None:
    value = getattr(project, "license_eop", None)
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return None


def _is_cancelled(project) -> bool:
    return str(getattr(project, "status", "") or "").strip().lower() == "cancelled"


def _candidate_projects(projects: list, target_date: dt.date) -> list:
    candidates_by_key = {}
    for project in projects:
        if _is_cancelled(project):
            continue
        license_date = _project_license_date(project)
        if license_date != target_date:
            continue
        key = (str(getattr(project, "project_name", "") or "").strip().lower(), license_date)
        candidates_by_key[key] = project
    return sorted(candidates_by_key.values(), key=lambda p: str(getattr(p, "project_name", "") or "").lower())


def _filter_already_sent(projects: list, days_before: int, force: bool, source_name: str) -> list:
    if force or source_name.startswith("Excel"):
        return projects

    unsent = []
    for project in projects:
        project_name = str(getattr(project, "project_name", "") or "").strip()
        license_date = _project_license_date(project)
        try:
            if license_date and has_license_expiry_alert_sent(project_name, license_date, days_before):
                logger.info("Skipping already-sent alert: %s | %s | %d day(s)", project_name, license_date, days_before)
                continue
        except Exception as exc:
            logger.warning("Could not check license alert log; will allow send: %s", exc)
        unsent.append(project)
    return unsent


def _build_email(projects: list, target_date: dt.date, days_before: int) -> tuple[str, str, str]:
    plural = "licenses" if len(projects) != 1 else "license"
    subject = f"CaddyCheck alert: {len(projects)} {plural} expire in {days_before} days"
    lines = [
        f"The following {plural} expire on {target_date.isoformat()} (in {days_before} days):",
        "",
    ]
    html_rows = []
    for project in projects:
        name = str(getattr(project, "project_name", "") or "").strip()
        country = str(getattr(project, "country", "") or "").strip()
        cameras = getattr(project, "num_cams", 0) or 0
        status = str(getattr(project, "status", "") or "").strip()
        lines.append(f"- {name} | {country or 'n/a'} | {cameras} camera(s) | status: {status or 'n/a'}")
        html_rows.append(
            "<tr>"
            f"<td>{name}</td>"
            f"<td>{country or 'n/a'}</td>"
            f"<td>{cameras}</td>"
            f"<td>{status or 'n/a'}</td>"
            "</tr>"
        )
    lines.extend(["", "Please update the License EOP in CaddyCheck CRM if the license was extended.", "", "CaddyCheck CRM"])
    body = "\n".join(lines)
    html = (
        f"<p>The following {plural} expire on <b>{target_date.isoformat()}</b> "
        f"(in {days_before} days):</p>"
        "<table border=\"1\" cellpadding=\"6\" cellspacing=\"0\">"
        "<thead><tr><th>Project</th><th>Country</th><th>Cameras</th><th>Status</th></tr></thead>"
        f"<tbody>{''.join(html_rows)}</tbody></table>"
        "<p>Please update the License EOP in CaddyCheck CRM if the license was extended.</p>"
        "<p>CaddyCheck CRM</p>"
    )
    return subject, body, html


def _log_sent_alerts(projects: list, days_before: int, recipients: list[str], cc_list: list[str], source_name: str) -> None:
    if source_name.startswith("Excel"):
        return
    for project in projects:
        license_date = _project_license_date(project)
        if not license_date:
            continue
        try:
            append_license_expiry_alert_log({
                "project_name": str(getattr(project, "project_name", "") or "").strip(),
                "license_eop": license_date.isoformat(),
                "days_before": days_before,
                "sent_to": ", ".join(recipients),
                "sent_cc": ", ".join(cc_list),
            })
        except Exception as exc:
            logger.warning("Could not write license alert log for %s: %s", getattr(project, "project_name", ""), exc)


def main() -> int:
    args = _parse_args()
    if args.days_before < 0:
        logger.error("--days-before must be zero or greater.")
        return 1

    today = dt.date.today()
    target_date = today + dt.timedelta(days=args.days_before)
    email_cfg = get_email_config()
    recipients = _split_csv(args.to) or list(email_cfg.get("default_recipients", []))
    cc_list = _split_csv(args.cc) or list(email_cfg.get("default_cc", []))

    if not recipients:
        logger.error("No recipients configured. Set default recipients, EMAIL_DEFAULT_RECIPIENTS, or pass --to.")
        return 1

    if not args.dry_run and not email_cfg.get("smtp_username"):
        logger.error("SMTP is not configured. Set SMTP_* env vars or app email settings.")
        return 1

    projects, source_name = _load_projects(args.source)
    candidates = _candidate_projects(projects, target_date)
    if not candidates:
        logger.info("No licenses expire on %s (%d day(s) from today) using %s.", target_date, args.days_before, source_name)
        return 0

    projects_to_alert = _filter_already_sent(candidates, args.days_before, args.force, source_name)
    if not projects_to_alert:
        logger.info("All matching license-expiry alerts were already sent.")
        return 0

    subject, body, html = _build_email(projects_to_alert, target_date, args.days_before)
    if args.subject:
        subject = args.subject

    logger.info(
        "Prepared license expiry alert using %s: %d project(s), target date %s, recipients=%s",
        source_name,
        len(projects_to_alert),
        target_date,
        recipients,
    )
    if args.dry_run:
        logger.info("Dry run only. No email sent and no alert log written.\n%s", body)
        return 0

    send_simple_email(subject, body, recipients, cc=cc_list, html_body=html, config=email_cfg)
    _log_sent_alerts(projects_to_alert, args.days_before, recipients, cc_list, source_name)
    logger.info("License expiry alert sent successfully.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
