"""Scheduled monthly invoice generation and email sender."""
import argparse
import datetime as dt
import logging
import sys
import tempfile
from pathlib import Path

from config.settings import append_sent_invoice_log, get_email_config, load_sent_invoices_log
from services.email_service import send_invoice_email
from services.excel_service import get_monthly_invoice_projects, get_next_invoice_number as get_excel_next_invoice_number
from services.excel_service import load_invoices as load_invoices_excel
from services.excel_service import load_projects as load_projects_excel
from services.excel_service import append_monthly_invoice_rows as append_excel_invoice_rows
from services.invoice_service import generate_monthly_invoice_pdf
from services.supabase_service import append_invoice_rows as append_supabase_invoice_rows
from services.supabase_service import get_next_invoice_number as get_supabase_next_invoice_number
from services.supabase_service import load_invoices as load_invoices_supabase
from services.supabase_service import load_projects as load_projects_supabase


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("monthly_auto_send")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate and send the monthly invoice on a schedule.",
    )
    parser.add_argument(
        "--month",
        help="Billing month name. Defaults to the previous calendar month.",
    )
    parser.add_argument(
        "--year",
        type=int,
        help="Billing year. Defaults to the year of the previous calendar month.",
    )
    parser.add_argument(
        "--invoice-number",
        type=int,
        help="Use a specific invoice number instead of the next available number.",
    )
    parser.add_argument(
        "--source",
        choices=["auto", "supabase", "excel"],
        default="auto",
        help="Data source to use. Default: auto.",
    )
    parser.add_argument(
        "--to",
        help="Override primary recipients with a comma-separated list.",
    )
    parser.add_argument(
        "--cc",
        help="Override CC recipients with a comma-separated list.",
    )
    parser.add_argument(
        "--subject",
        help="Override the email subject.",
    )
    parser.add_argument(
        "--body",
        help="Override the email body.",
    )
    parser.add_argument(
        "--skip-ledger-save",
        action="store_true",
        help="Do not add invoice rows to the ledger before sending.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Allow sending again even if the month/year already appears in the sent log.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview what would be sent without saving to the ledger or sending the email.",
    )
    return parser.parse_args()


def _previous_month_period(today: dt.date | None = None) -> tuple[str, int]:
    today = today or dt.date.today()
    first_of_month = today.replace(day=1)
    previous_month_last_day = first_of_month - dt.timedelta(days=1)
    return previous_month_last_day.strftime("%B"), previous_month_last_day.year


def _split_csv(value: str | None) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in value.split(",") if item.strip()]


def _load_data(source: str) -> tuple[list, list, str]:
    if source == "excel":
        return load_projects_excel(), load_invoices_excel(), "Excel (local fallback)"
    if source == "supabase":
        return load_projects_supabase(), load_invoices_supabase(), "Supabase"

    try:
        return load_projects_supabase(), load_invoices_supabase(), "Supabase"
    except RuntimeError as exc:
        if "Supabase credentials not configured" not in str(exc):
            raise
        logger.info("Supabase not configured. Falling back to Excel data files.")
        return load_projects_excel(), load_invoices_excel(), "Excel (local fallback)"


def _is_excel_source(source_name: str) -> bool:
    return source_name.startswith("Excel")


def _get_next_invoice_number(invoices: list, source_name: str) -> int:
    if _is_excel_source(source_name):
        return get_excel_next_invoice_number(invoices)
    return get_supabase_next_invoice_number()


def _append_invoice_rows(invoice_number: int, projects: list, year: int, source_name: str) -> int:
    if _is_excel_source(source_name):
        return append_excel_invoice_rows(invoice_number=invoice_number, projects=projects, year=year)
    return append_supabase_invoice_rows(invoice_number=invoice_number, projects=projects, year=year)


def _already_sent(month_name: str, year: int) -> bool:
    for row in load_sent_invoices_log():
        if str(row.get("month", "")).strip().lower() == month_name.strip().lower() and int(row.get("year", 0)) == year:
            return True
    return False


def main() -> int:
    args = _parse_args()
    month_name, year = _previous_month_period()
    if args.month:
        month_name = args.month.strip()
    if args.year:
        year = args.year

    if _already_sent(month_name, year) and not args.force:
        logger.error(
            "Invoice for %s %s already appears in sent_invoices_log.json. Use --force to send again.",
            month_name,
            year,
        )
        return 1

    email_cfg = get_email_config()
    recipients = _split_csv(args.to) or list(email_cfg.get("default_recipients", []))
    cc_list = _split_csv(args.cc) or list(email_cfg.get("default_cc", []))

    if not recipients:
        logger.error("No recipients configured. Set default recipients or pass --to.")
        return 1

    if not args.dry_run and not email_cfg.get("smtp_username"):
        logger.error("SMTP is not configured. Update the email settings before scheduling this script.")
        return 1

    projects, invoices, source_name = _load_data(args.source)
    month_projects = get_monthly_invoice_projects(projects, month_name, year)
    if not month_projects:
        logger.error("No projects found for billing month %s.", month_name)
        return 1

    invoice_number = args.invoice_number or _get_next_invoice_number(invoices, source_name)
    total_amount = sum(project.get_expected_amount(year) for project in month_projects)

    subject_template = email_cfg.get("default_subject_template", "Monthly Invoice - {month} {year}")
    body_template = email_cfg.get(
        "default_body_template",
        "Dear Team,\n\nPlease find attached the monthly maintenance invoice for {month} {year}.\n\nBest regards,\nVideo Inform Ltd",
    )
    subject = args.subject or subject_template.format(month=month_name, year=year)
    body = args.body or body_template.format(month=month_name, year=year)

    logger.info(
        "Prepared invoice #%s for %s %s using %s: %d project(s), total €%,.0f",
        invoice_number,
        month_name,
        year,
        source_name,
        len(month_projects),
        total_amount,
    )
    logger.info("Recipients: %s", recipients)
    if cc_list:
        logger.info("CC: %s", cc_list)

    if args.dry_run:
        logger.info("Dry run only. No ledger rows saved and no email sent.")
        return 0

    with tempfile.TemporaryDirectory() as tmp_dir:
        ledger_rows_added = None
        if not args.skip_ledger_save:
            ledger_rows_added = _append_invoice_rows(
                invoice_number=invoice_number,
                projects=month_projects,
                year=year,
                source_name=source_name,
            )

        out_path = generate_monthly_invoice_pdf(
            projects=month_projects,
            month_name=month_name,
            year=year,
            invoice_number=invoice_number,
            output_dir=Path(tmp_dir),
        )
        send_invoice_email(
            attachment_path=out_path,
            recipients=recipients,
            cc=cc_list,
            subject=subject,
            body=body,
            config=email_cfg,
        )

    append_sent_invoice_log({
        "sent_at": dt.datetime.utcnow().isoformat(),
        "invoice_number": invoice_number,
        "month": month_name,
        "year": year,
        "pdf_filename": f"CC_M-inv_{invoice_number}_{month_name[:3]}_{year}.pdf",
        "recipients": recipients,
        "cc": cc_list,
        "subject": subject,
        "project_count": len(month_projects),
        "total_amount": total_amount,
        "saved_to_ledger": not args.skip_ledger_save,
        "ledger_rows_added": ledger_rows_added,
        "source_name": source_name,
    })

    logger.info("Invoice email sent successfully for %s %s.", month_name, year)
    return 0


if __name__ == "__main__":
    sys.exit(main())