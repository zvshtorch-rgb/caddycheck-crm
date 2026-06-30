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
    sender_email = config.get("sender_email", "") or username

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


def send_simple_email(
    subject: str,
    body: str,
    recipients: List[str],
    cc: Optional[List[str]] = None,
    html_body: Optional[str] = None,
    config: Optional[dict] = None,
) -> None:
    """
    Send a plain-text (optionally HTML) email with no attachment.

    Used for notifications such as the CEO purchase-order approval request.
    Reuses the same SMTP configuration as :func:`send_invoice_email`.
    """
    if config is None:
        config = get_email_config()

    smtp_host = config.get("smtp_host", "")
    smtp_port = int(config.get("smtp_port", 587))
    use_tls = bool(config.get("smtp_use_tls", True))
    username = config.get("smtp_username", "")
    password = config.get("smtp_password", "")
    sender_name = config.get("sender_name", "CaddyCheck CRM")
    sender_email = config.get("sender_email", "") or username

    if not smtp_host:
        raise ValueError("SMTP host is not configured.")
    if not sender_email:
        raise ValueError("Sender email is not configured.")
    if not recipients:
        raise ValueError("No recipients specified.")

    msg = MIMEMultipart("alternative")
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = ", ".join(recipients)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))
    if html_body:
        msg.attach(MIMEText(html_body, "html", "utf-8"))

    all_recipients = list(recipients) + (cc or [])

    logger.info("Sending notification '%s' to %s", subject, all_recipients)

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

    logger.info("Notification email sent successfully.")


def _get_graph_setting(key: str) -> str:
    """Read a Graph credential from env var, then st.secrets [order_intake]."""
    import os
    value = os.environ.get(key, "").strip()
    if value:
        return value
    try:
        import streamlit as st
        cfg = st.secrets.get("order_intake", {})
        value = str(cfg.get(key, "")).strip()
        if value:
            return value
        value = str(st.secrets.get(key, "")).strip()
        return value
    except Exception:
        return ""


def graph_email_available() -> bool:
    """Return True when all required Graph credentials are reachable."""
    return all(
        _get_graph_setting(k)
        for k in (
            "ORDER_INTAKE_GRAPH_TENANT_ID",
            "ORDER_INTAKE_GRAPH_CLIENT_ID",
            "ORDER_INTAKE_GRAPH_CLIENT_SECRET",
            "ORDER_INTAKE_GRAPH_MAILBOX",
        )
    )


def send_graph_invoice_email(
    attachment_path: Path,
    subject: str,
    body: str,
    recipients: List[str],
    cc: Optional[List[str]] = None,
) -> None:
    """Send an invoice email with a file attachment via Microsoft Graph API."""
    if not attachment_path or not attachment_path.exists():
        raise ValueError(f"Attachment not found: {attachment_path}")
    attachment_bytes = attachment_path.read_bytes()
    suffix = attachment_path.suffix.lower()
    content_type = (
        "application/pdf" if suffix == ".pdf"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if suffix in (".xlsx", ".xlsm")
        else "application/octet-stream"
    )
    send_graph_email(
        subject=subject,
        body=body,
        recipients=recipients,
        cc=cc,
        attachment_bytes=attachment_bytes,
        attachment_filename=attachment_path.name,
        attachment_content_type=content_type,
    )


def send_graph_email(
    subject: str,
    body: str,
    recipients: List[str],
    cc: Optional[List[str]] = None,
    html_body: Optional[str] = None,
    sender_mailbox: Optional[str] = None,
    attachment_bytes: Optional[bytes] = None,
    attachment_filename: Optional[str] = None,
    attachment_content_type: str = "application/pdf",
) -> None:
    """
    Send an email via Microsoft Graph API (app-only client credentials).

    Uses the same Graph credentials as the order-intake provider
    (ORDER_INTAKE_GRAPH_TENANT_ID / CLIENT_ID / CLIENT_SECRET / MAILBOX).
    Avoids SMTP entirely — no Basic Auth required.
    """
    import requests

    tenant_id = _get_graph_setting("ORDER_INTAKE_GRAPH_TENANT_ID")
    client_id = _get_graph_setting("ORDER_INTAKE_GRAPH_CLIENT_ID")
    client_secret = _get_graph_setting("ORDER_INTAKE_GRAPH_CLIENT_SECRET")
    mailbox = sender_mailbox or _get_graph_setting("ORDER_INTAKE_GRAPH_MAILBOX")

    if not all([tenant_id, client_id, client_secret, mailbox]):
        raise ValueError(
            "Graph credentials not fully configured. "
            "Set ORDER_INTAKE_GRAPH_TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAILBOX."
        )

    # Acquire app-only token
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_resp = requests.post(
        token_url,
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    token_resp.raise_for_status()
    access_token = token_resp.json()["access_token"]

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    to_recipients = [{"emailAddress": {"address": r}} for r in recipients]
    cc_recipients = [{"emailAddress": {"address": r}} for r in (cc or [])]

    content_type = "html" if html_body else "text"
    content_value = html_body if html_body else body

    message_payload: dict = {
        "subject": subject,
        "body": {"contentType": content_type, "content": content_value},
        "toRecipients": to_recipients,
        "ccRecipients": cc_recipients,
    }

    if attachment_bytes and attachment_filename:
        import base64
        message_payload["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": attachment_filename,
                "contentType": attachment_content_type,
                "contentBytes": base64.b64encode(attachment_bytes).decode("ascii"),
            }
        ]

    payload = {
        "message": message_payload,
        "saveToSentItems": "true",
    }

    send_url = f"https://graph.microsoft.com/v1.0/users/{mailbox}/sendMail"
    resp = requests.post(send_url, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    logger.info("Graph email '%s' sent to %s", subject, recipients)


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
