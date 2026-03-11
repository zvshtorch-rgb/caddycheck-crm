"""Service for sending invoice emails via SMTP."""
import logging
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import List, Optional

from config.settings import get_email_config

logger = logging.getLogger(__name__)


def send_invoice_email(
    attachment_path: Path,
    subject: str,
    body: str,
    recipients: List[str],
    cc: Optional[List[str]] = None,
    config: Optional[dict] = None,
) -> None:
    """
    Send an email with the invoice file attached.

    Parameters
    ----------
    attachment_path : Path
        Path to the Excel invoice file.
    subject : str
        Email subject line.
    body : str
        Plain-text body of the email.
    recipients : list of str
        Primary recipient email addresses.
    cc : list of str, optional
        CC recipient addresses.
    config : dict, optional
        Email configuration; loaded from settings if not provided.

    Raises
    ------
    smtplib.SMTPException
        On any SMTP-level error.
    ValueError
        If required configuration fields are missing.
    """
    if config is None:
        config = get_email_config()

    smtp_host = config.get("smtp_host", "")
    smtp_port = int(config.get("smtp_port", 587))
    use_tls = bool(config.get("smtp_use_tls", True))
    username = config.get("smtp_username", "")
    password = config.get("smtp_password", "")
    sender_name = config.get("sender_name", "CaddyCheck CRM")
    sender_email = config.get("sender_email", username)

    if not smtp_host:
        raise ValueError("SMTP host is not configured.")
    if not sender_email:
        raise ValueError("Sender email is not configured.")
    if not recipients:
        raise ValueError("No recipients specified.")

    # Build message
    msg = MIMEMultipart()
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = ", ".join(recipients)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Attach invoice file
    if attachment_path and attachment_path.exists():
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={attachment_path.name}",
        )
        msg.attach(part)

    all_recipients = list(recipients) + (cc or [])

    logger.info(
        "Sending email '%s' to %s via %s:%s",
        subject,
        all_recipients,
        smtp_host,
        smtp_port,
    )

    if use_tls:
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            if username and password:
                server.login(username, password)
            server.sendmail(sender_email, all_recipients, msg.as_string())
    else:
        with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30) as server:
            if username and password:
                server.login(username, password)
            server.sendmail(sender_email, all_recipients, msg.as_string())

    logger.info("Email sent successfully.")


def test_smtp_connection(config: dict) -> tuple:
    """
    Test SMTP connection with provided config.

    Returns
    -------
    (success: bool, message: str)
    """
    try:
        smtp_host = config.get("smtp_host", "")
        smtp_port = int(config.get("smtp_port", 587))
        use_tls = bool(config.get("smtp_use_tls", True))
        username = config.get("smtp_username", "")
        password = config.get("smtp_password", "")

        if use_tls:
            context = ssl.create_default_context()
            with smtplib.SMTP(smtp_host, smtp_port, timeout=10) as server:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
                if username and password:
                    server.login(username, password)
        else:
            with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=10) as server:
                if username and password:
                    server.login(username, password)

        return True, "Connection successful!"
    except Exception as e:
        logger.error("SMTP test failed: %s", e)
        return False, str(e)
