"""
Email intake service.

Reads an inbox and yields incoming purchase-order emails that carry a PDF
attachment. The concrete transport is pluggable so Gmail / generic IMAP /
Microsoft Graph (Office 365) can be swapped via configuration.

Configuration is read from environment variables (works in standalone scripts)
with a fallback to Streamlit secrets ``[order_intake]`` when running in the app.

Provider selection (``ORDER_INTAKE_PROVIDER``):
- ``imap``   → generic IMAP (also works for Gmail and Office 365 IMAP).
- ``graph``  → Microsoft Graph API (preferred for Office 365 mailboxes).

Required settings per provider are documented in the README.
"""
from __future__ import annotations

import datetime
import email
import imaplib
import logging
import os
from dataclasses import dataclass, field
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


@dataclass
class IncomingMessage:
    """A normalized email carrying at least one PDF attachment."""

    message_id: str
    provider: str
    from_address: str
    subject: str
    received_at: Optional[datetime.datetime]
    body_text: str
    pdf_filename: str
    pdf_bytes: bytes
    raw_id: str = ""  # provider-specific id used to mark as read
    extra: Dict[str, Any] = field(default_factory=dict)


# ── Configuration ─────────────────────────────────────────────────────────────
def _get_setting(key: str, default: str = "") -> str:
    """Read a setting from env first, then Streamlit secrets ``[order_intake]``."""
    value = os.environ.get(key, "").strip()
    if value:
        return value
    try:
        import streamlit as st  # noqa: WPS433

        cfg = st.secrets.get("order_intake", {})
        return str(cfg.get(key, cfg.get(key.lower(), default))).strip()
    except Exception:
        return default


# ── PDF text extraction (best effort, optional) ───────────────────────────────
def extract_pdf_text(pdf_bytes: bytes) -> str:
    try:
        import io

        import pdfplumber

        text_parts: List[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text_parts.append(page.extract_text() or "")
        return "\n".join(part for part in text_parts if part).strip()
    except Exception as exc:
        logger.info("PDF text extraction skipped: %s", exc)
        return ""


def _decode(value: Optional[str]) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


# ── Provider base ─────────────────────────────────────────────────────────────
class EmailProvider:
    name = "base"

    def is_configured(self) -> bool:
        """Return True only when the credentials needed to poll are present."""
        return False

    def fetch_unread_with_pdf(self) -> List[IncomingMessage]:
        raise NotImplementedError

    def mark_processed(self, message: IncomingMessage) -> None:
        """Optionally mark the source message as read/seen. Default: no-op."""
        return None


# ── Generic IMAP provider (Gmail / Office 365 IMAP) ───────────────────────────
class ImapEmailProvider(EmailProvider):
    name = "imap"

    def __init__(self) -> None:
        self.host = _get_setting("ORDER_INTAKE_IMAP_HOST")
        self.port = int(_get_setting("ORDER_INTAKE_IMAP_PORT", "993") or "993")
        self.username = _get_setting("ORDER_INTAKE_IMAP_USERNAME")
        self.password = _get_setting("ORDER_INTAKE_IMAP_PASSWORD")
        self.folder = _get_setting("ORDER_INTAKE_IMAP_FOLDER", "INBOX") or "INBOX"
        self._seen_uids: List[bytes] = []

    def is_configured(self) -> bool:
        return bool(self.host and self.username and self.password)

    def _connect(self) -> imaplib.IMAP4_SSL:
        if not self.host or not self.username or not self.password:
            raise RuntimeError(
                "IMAP settings missing. Set ORDER_INTAKE_IMAP_HOST/USERNAME/PASSWORD."
            )
        conn = imaplib.IMAP4_SSL(self.host, self.port)
        conn.login(self.username, self.password)
        return conn

    def fetch_unread_with_pdf(self) -> List[IncomingMessage]:
        messages: List[IncomingMessage] = []
        self._seen_uids = []
        conn = self._connect()
        try:
            conn.select(self.folder)
            status, data = conn.uid("search", None, "UNSEEN")
            if status != "OK":
                return messages
            uids = (data[0] or b"").split()
            for uid in uids:
                status, msg_data = conn.uid("fetch", uid, "(RFC822)")
                if status != "OK" or not msg_data or not msg_data[0]:
                    continue
                raw_email = msg_data[0][1]
                parsed = self._parse(raw_email, uid.decode())
                if parsed is not None:
                    messages.append(parsed)
                    self._seen_uids.append(uid)
        finally:
            try:
                conn.logout()
            except Exception:
                pass
        return messages

    def _parse(self, raw_email: bytes, uid: str) -> Optional[IncomingMessage]:
        msg = email.message_from_bytes(raw_email)
        pdf_filename = ""
        pdf_bytes = b""
        body_text = ""

        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition") or "")
            filename = _decode(part.get_filename())
            if filename.lower().endswith(".pdf") or content_type == "application/pdf":
                payload = part.get_payload(decode=True)
                if payload:
                    pdf_filename = filename or "purchase_order.pdf"
                    pdf_bytes = payload
            elif content_type == "text/plain" and "attachment" not in disposition.lower():
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    try:
                        body_text += payload.decode(charset, errors="replace")
                    except Exception:
                        body_text += payload.decode("utf-8", errors="replace")

        if not pdf_bytes:
            return None  # no PDF → not a purchase order we care about

        received_at: Optional[datetime.datetime] = None
        try:
            received_at = parsedate_to_datetime(msg.get("Date"))
        except Exception:
            received_at = None

        message_id = _decode(msg.get("Message-ID")) or f"imap-{self.host}-{uid}"
        return IncomingMessage(
            message_id=message_id,
            provider=self.name,
            from_address=_decode(msg.get("From")),
            subject=_decode(msg.get("Subject")),
            received_at=received_at,
            body_text=body_text.strip(),
            pdf_filename=pdf_filename,
            pdf_bytes=pdf_bytes,
            raw_id=uid,
        )

    def mark_processed(self, message: IncomingMessage) -> None:
        if not message.raw_id:
            return
        try:
            conn = self._connect()
            try:
                conn.select(self.folder)
                conn.uid("store", message.raw_id, "+FLAGS", "(\\Seen)")
            finally:
                conn.logout()
        except Exception as exc:
            logger.warning("Could not mark IMAP message as seen: %s", exc)


# ── Microsoft Graph provider (Office 365) ─────────────────────────────────────
class GraphEmailProvider(EmailProvider):
    name = "graph"

    GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

    def __init__(self) -> None:
        self.tenant_id = _get_setting("ORDER_INTAKE_GRAPH_TENANT_ID")
        self.client_id = _get_setting("ORDER_INTAKE_GRAPH_CLIENT_ID")
        self.client_secret = _get_setting("ORDER_INTAKE_GRAPH_CLIENT_SECRET")
        self.mailbox = _get_setting("ORDER_INTAKE_GRAPH_MAILBOX")  # user UPN/objectId
        self.folder = _get_setting("ORDER_INTAKE_GRAPH_FOLDER", "Inbox") or "Inbox"
        self._processed_ids: List[str] = []

    def is_configured(self) -> bool:
        return bool(self.tenant_id and self.client_id and self.client_secret and self.mailbox)

    def _token(self) -> str:
        import requests

        if not all([self.tenant_id, self.client_id, self.client_secret, self.mailbox]):
            raise RuntimeError(
                "Graph settings missing. Set ORDER_INTAKE_GRAPH_TENANT_ID/"
                "CLIENT_ID/CLIENT_SECRET/MAILBOX."
            )
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        resp = requests.post(
            url,
            data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()["access_token"]

    def fetch_unread_with_pdf(self) -> List[IncomingMessage]:
        import requests

        token = self._token()
        headers = {"Authorization": f"Bearer {token}"}
        messages: List[IncomingMessage] = []
        self._processed_ids = []

        list_url = (
            f"{self.GRAPH_ROOT}/users/{self.mailbox}/mailFolders/"
            f"{self.folder}/messages"
            "?$filter=isRead eq false and hasAttachments eq true"
            "&$select=id,subject,from,receivedDateTime,bodyPreview,internetMessageId"
            "&$top=25"
        )
        resp = requests.get(list_url, headers=headers, timeout=30)
        resp.raise_for_status()
        for item in resp.json().get("value", []):
            msg = self._parse_message(item, headers)
            if msg is not None:
                messages.append(msg)
        return messages

    def _parse_message(self, item: Dict[str, Any], headers: Dict[str, str]) -> Optional[IncomingMessage]:
        import requests

        msg_id = item.get("id", "")
        att_url = (
            f"{self.GRAPH_ROOT}/users/{self.mailbox}/messages/{msg_id}/attachments"
        )
        att_resp = requests.get(att_url, headers=headers, timeout=30)
        att_resp.raise_for_status()

        pdf_filename = ""
        pdf_bytes = b""
        for att in att_resp.json().get("value", []):
            name = att.get("name", "")
            ctype = att.get("contentType", "")
            if name.lower().endswith(".pdf") or ctype == "application/pdf":
                content_b64 = att.get("contentBytes")
                if content_b64:
                    import base64

                    pdf_filename = name or "purchase_order.pdf"
                    pdf_bytes = base64.b64decode(content_b64)
                    break

        if not pdf_bytes:
            return None

        received_at: Optional[datetime.datetime] = None
        try:
            received_at = datetime.datetime.fromisoformat(
                str(item.get("receivedDateTime", "")).replace("Z", "+00:00")
            )
        except Exception:
            received_at = None

        from_address = ""
        try:
            from_address = item["from"]["emailAddress"]["address"]
        except Exception:
            from_address = ""

        return IncomingMessage(
            message_id=item.get("internetMessageId") or msg_id,
            provider=self.name,
            from_address=from_address,
            subject=item.get("subject", ""),
            received_at=received_at,
            body_text=item.get("bodyPreview", ""),
            pdf_filename=pdf_filename,
            pdf_bytes=pdf_bytes,
            raw_id=msg_id,
        )

    def mark_processed(self, message: IncomingMessage) -> None:
        import requests

        if not message.raw_id:
            return
        try:
            token = self._token()
            url = f"{self.GRAPH_ROOT}/users/{self.mailbox}/messages/{message.raw_id}"
            requests.patch(
                url,
                headers={"Authorization": f"Bearer {token}"},
                json={"isRead": True},
                timeout=30,
            )
        except Exception as exc:
            logger.warning("Could not mark Graph message as read: %s", exc)


# ── Factory ───────────────────────────────────────────────────────────────────
def get_email_provider() -> EmailProvider:
    provider = (_get_setting("ORDER_INTAKE_PROVIDER", "imap") or "imap").lower()
    if provider == "graph":
        return GraphEmailProvider()
    if provider in ("imap", "gmail"):
        return ImapEmailProvider()
    raise RuntimeError(f"Unknown ORDER_INTAKE_PROVIDER: {provider!r}")
