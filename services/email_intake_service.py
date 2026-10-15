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
import re
import time
from dataclasses import dataclass, field
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime
from typing import Any, Dict, List, Optional
from urllib.parse import quote

logger = logging.getLogger(__name__)

_sleep = time.sleep


class IntakeError(Exception):
    """A mailbox/Graph failure that must fail the run.

    ``kind`` is one of: ``config`` | ``auth`` | ``mailbox`` | ``transient``.
    """

    def __init__(self, kind: str, message: str, status: Optional[int] = None, code: str = "") -> None:
        super().__init__(message)
        self.kind = kind
        self.status = status
        self.code = code


@dataclass
class MessageHeader:
    """Metadata of a mailbox message that has attachments (no attachment bytes)."""

    message_id: str  # RFC internetMessageId; deduplication key shared with legacy rows
    graph_id: str  # immutable Graph id (stable across folder moves)
    from_address: str
    subject: str
    received_at: Optional[datetime.datetime]
    body_preview: str = ""


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
        logger.info("PDF text extraction skipped (%s)", type(exc).__name__)
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
        self._access_token = ""
        self._token_expiry = 0.0

    def is_configured(self) -> bool:
        return bool(self.tenant_id and self.client_id and self.client_secret and self.mailbox)

    def missing_settings(self) -> List[str]:
        """Names (never values) of required settings that are not set."""
        required = {
            "ORDER_INTAKE_GRAPH_TENANT_ID": self.tenant_id,
            "ORDER_INTAKE_GRAPH_CLIENT_ID": self.client_id,
            "ORDER_INTAKE_GRAPH_CLIENT_SECRET": self.client_secret,
            "ORDER_INTAKE_GRAPH_MAILBOX": self.mailbox,
        }
        return [name for name, value in required.items() if not value]

    # -- HTTP helpers ----------------------------------------------------------
    _RETRY_STATUSES = (429, 500, 502, 503, 504)
    _MAX_ATTEMPTS = 3

    @staticmethod
    def _retry_delay(resp: Any, attempt: int) -> float:
        try:
            return max(0.0, min(float(resp.headers.get("Retry-After", 2 * attempt)), 30.0))
        except (TypeError, ValueError):
            return float(2 * attempt)

    def _http(self, method: str, url: str, **kwargs: Any) -> Any:
        import requests

        for attempt in range(1, self._MAX_ATTEMPTS + 1):
            try:
                resp = requests.request(method, url, timeout=30, **kwargs)
            except requests.RequestException as exc:
                if attempt == self._MAX_ATTEMPTS:
                    raise IntakeError("transient", f"network error: {type(exc).__name__}") from None
                _sleep(2 * attempt)
                continue
            if resp.status_code in self._RETRY_STATUSES and attempt < self._MAX_ATTEMPTS:
                _sleep(self._retry_delay(resp, attempt))
                continue
            return resp
        raise IntakeError("transient", "request retries exhausted")  # pragma: no cover

    @staticmethod
    def _error_from_response(resp: Any, what: str, token_endpoint: bool = False) -> IntakeError:
        code, detail = "", ""
        try:
            body = resp.json()
            err = body.get("error")
            if isinstance(err, dict):
                code = str(err.get("code") or "")
                detail = str(err.get("message") or "")
            else:
                code = str(err or "")
                detail = str(body.get("error_description") or "")
        except Exception:
            pass
        # Free-text descriptions can echo tenant/client ids or mailbox addresses; keep only the codes.
        aad_code = re.search(r"AADSTS\d+", detail)
        detail = aad_code.group(0) if aad_code else ""
        status = resp.status_code
        if (token_endpoint and status < 500) or status in (401, 403):
            kind = "auth"
        elif status == 404:
            kind = "mailbox"
        else:
            kind = "transient"
        message = f"{what} failed: HTTP {status} {code} {detail}".strip()
        return IntakeError(kind, message, status=status, code=code)

    def _token(self) -> str:
        if self._access_token and time.time() < self._token_expiry - 60:
            return self._access_token
        missing = self.missing_settings()
        if missing:
            raise IntakeError("config", "Graph settings missing: " + ", ".join(missing))
        url = f"https://login.microsoftonline.com/{quote(self.tenant_id, safe='')}/oauth2/v2.0/token"
        resp = self._http(
            "POST",
            url,
            data={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
        )
        if not resp.ok:
            raise self._error_from_response(resp, "Token request", token_endpoint=True)
        try:
            payload = resp.json()
            self._access_token = payload["access_token"]
            self._token_expiry = time.time() + int(payload.get("expires_in", 3600))
        except (ValueError, KeyError, TypeError):
            raise IntakeError("auth", "Token request returned no access token") from None
        return self._access_token

    def _headers(self) -> Dict[str, str]:
        # ImmutableId keeps message ids stable if a message is moved between folders.
        return {"Authorization": f"Bearer {self._token()}", "Prefer": 'IdType="ImmutableId"'}

    def _mailbox_url(self, suffix: str) -> str:
        return f"{self.GRAPH_ROOT}/users/{quote(self.mailbox, safe='@')}{suffix}"

    def _folder_path(self) -> str:
        return f"/mailFolders/{quote(self.folder, safe='')}"

    # -- Read-only mailbox operations -----------------------------------------
    def preflight(self) -> Dict[str, Any]:
        """Verify token acquisition and read access to the configured folder."""
        missing = self.missing_settings()
        if missing:
            raise IntakeError("config", "Graph settings missing: " + ", ".join(missing))
        resp = self._http(
            "GET",
            self._mailbox_url(self._folder_path()),
            headers=self._headers(),
            params={"$select": "id,displayName,totalItemCount,unreadItemCount"},
        )
        if not resp.ok:
            raise self._error_from_response(resp, "Mailbox folder check")
        data = resp.json()
        return {
            "folder": data.get("displayName") or self.folder,
            "total_items": data.get("totalItemCount"),
            "unread_items": data.get("unreadItemCount"),
        }

    def list_message_headers(self, since: datetime.datetime, max_pages: int = 20) -> List[MessageHeader]:
        """List messages with attachments received at/after ``since``, any read state."""
        since_utc = since.astimezone(datetime.timezone.utc)
        params: Optional[Dict[str, str]] = {
            "$filter": f"receivedDateTime ge {since_utc:%Y-%m-%dT%H:%M:%SZ} and hasAttachments eq true",
            "$orderby": "receivedDateTime asc",
            "$select": "id,internetMessageId,subject,from,receivedDateTime,bodyPreview",
            "$top": "50",
        }
        url: Optional[str] = self._mailbox_url(self._folder_path() + "/messages")
        headers: List[MessageHeader] = []
        pages = 0
        while url:
            pages += 1
            if pages > max_pages:
                raise IntakeError(
                    "mailbox", f"Message listing exceeded {max_pages} pages; narrow the lookback window"
                )
            resp = self._http("GET", url, headers=self._headers(), params=params)
            if not resp.ok:
                raise self._error_from_response(resp, "Message listing")
            payload = resp.json()
            for item in payload.get("value", []):
                headers.append(self._to_header(item))
            url = payload.get("@odata.nextLink")
            params = None  # nextLink already carries the query
        return headers

    @staticmethod
    def _to_header(item: Dict[str, Any]) -> MessageHeader:
        graph_id = str(item.get("id") or "")
        received_at: Optional[datetime.datetime] = None
        try:
            received_at = datetime.datetime.fromisoformat(
                str(item.get("receivedDateTime", "")).replace("Z", "+00:00")
            )
        except ValueError:
            received_at = None
        try:
            from_address = item["from"]["emailAddress"]["address"] or ""
        except (KeyError, TypeError):
            from_address = ""
        return MessageHeader(
            message_id=str(item.get("internetMessageId") or graph_id),
            graph_id=graph_id,
            from_address=from_address,
            subject=str(item.get("subject") or ""),
            received_at=received_at,
            body_preview=str(item.get("bodyPreview") or ""),
        )

    def fetch_pdf_message(self, header: MessageHeader) -> Optional[IncomingMessage]:
        """Download the first PDF attachment of a message, or None when it has none."""
        base = self._mailbox_url(f"/messages/{quote(header.graph_id, safe='')}/attachments")
        resp = self._http(
            "GET",
            base,
            headers=self._headers(),
            params={"$select": "id,name,contentType,size,isInline"},
        )
        if not resp.ok:
            raise self._error_from_response(resp, "Attachment listing")
        pdfs = [
            att
            for att in resp.json().get("value", [])
            if str(att.get("name", "")).lower().endswith(".pdf")
            or str(att.get("contentType", "")).lower() == "application/pdf"
        ]
        if not pdfs:
            return None
        att = pdfs[0]
        raw = self._http(
            "GET",
            f"{base}/{quote(str(att.get('id', '')), safe='')}/$value",
            headers=self._headers(),
        )
        if not raw.ok:
            raise self._error_from_response(raw, "Attachment download")
        if not raw.content:
            return None
        return IncomingMessage(
            message_id=header.message_id,
            provider=self.name,
            from_address=header.from_address,
            subject=header.subject,
            received_at=header.received_at,
            body_text=header.body_preview,
            pdf_filename=str(att.get("name") or "purchase_order.pdf"),
            pdf_bytes=raw.content,
            raw_id=header.graph_id,
            extra={"pdf_count": len(pdfs)},
        )

    # Intentionally no mark_processed(): the Graph provider never changes read state.


# ── Factory ───────────────────────────────────────────────────────────────────
def get_email_provider() -> EmailProvider:
    provider = (_get_setting("ORDER_INTAKE_PROVIDER", "graph") or "graph").lower()
    if provider == "graph":
        return GraphEmailProvider()
    if provider in ("imap", "gmail"):
        return ImapEmailProvider()
    raise RuntimeError(f"Unknown ORDER_INTAKE_PROVIDER: {provider!r}")
