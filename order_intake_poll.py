"""
Order intake poller.

Scans the shared order mailbox through Microsoft Graph (application
authentication) for purchase-order emails with a PDF attachment, whatever their
read/unread state, and never changes that state. Each new email is stored once
(keyed by message id), a ``purchase_orders`` record is created, and the CEO is
emailed a single-use approval link. Every step is resumable and idempotent, so
retries and overlapping runs cannot create duplicate orders or approval emails.

Run manually or on a schedule (GitHub Actions):

    py order_intake_poll.py --check      # connectivity + schema checks only
    py order_intake_poll.py --dry-run    # list what would happen, write nothing
    py order_intake_poll.py              # live processing

Exit codes: 0 = ok, 1 = fatal (config/auth/mailbox/database), 2 = some messages failed.

Required configuration (env vars or Streamlit secrets) is documented in the
README under "Purchase-order approval workflow".
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import logging
import os
import re
import sys
import time
import uuid
from dataclasses import dataclass, field
from typing import Optional

from services import order_approval_service as approvals
from services.email_intake_service import (
    IncomingMessage,
    IntakeError,
    MessageHeader,
    extract_pdf_text,
    get_email_provider,
)
from services.email_service import send_graph_email

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("order_intake_poll")
# Request-level loggers would print URLs (message ids, project names in storage paths).
for _noisy in ("httpx", "httpcore", "hpack", "urllib3", "postgrest", "storage3", "gotrue", "supabase"):
    logging.getLogger(_noisy).setLevel(logging.WARNING)

EXIT_OK = 0
EXIT_FATAL = 1
EXIT_PARTIAL = 2
DEFAULT_LOOKBACK_HOURS = 168  # seven days
DEFAULT_MAX_NEW = 10
MAX_BACKFILL_DAYS = 30
MIN_FILTER_LENGTH = 6


def _get_setting(key: str, default: str = "") -> str:
    value = os.environ.get(key, "").strip()
    if value:
        return value
    try:
        import streamlit as st  # noqa: WPS433

        cfg = st.secrets.get("order_intake", {})
        return str(cfg.get(key, cfg.get(key.lower(), default))).strip()
    except Exception:
        return default


def _app_base_url() -> str:
    """Public base URL of the Streamlit app (used to build approval links)."""
    url = _get_setting("APP_BASE_URL", "").rstrip("/")
    return url or "http://localhost:8501"


def _ceo_recipients() -> list[str]:
    raw = _get_setting("ORDER_APPROVAL_CEO_EMAIL", "")
    return [addr.strip() for addr in raw.split(",") if addr.strip()]


def _build_approval_email(order: dict, token: str, resent: bool = False) -> tuple[str, str, str]:
    base = _app_base_url()
    approve_link = f"{base}/?approval={token}"
    project = order.get("project_name") or order.get("customer_name") or "(unknown)"
    amount = order.get("amount")
    amount_str = f"{amount:,.2f} {order.get('currency') or 'EUR'}" if amount is not None else "n/a"

    subject = f"{'[Resent] ' if resent else ''}Purchase order approval needed: {project}"
    resent_text = "This is a resend; any earlier approval link for this order no longer works.\n\n" if resent else ""
    body = (
        "A new customer purchase order requires your approval.\n\n"
        f"{resent_text}"
        f"Project / Customer: {project}\n"
        f"Reference: {order.get('order_reference') or 'n/a'}\n"
        f"Amount: {amount_str}\n"
        f"From: {order.get('customer_email') or 'n/a'}\n\n"
        f"Review and decide here (link valid for 7 days):\n{approve_link}\n\n"
        "You can approve, reject, or request a correction on that page.\n\n"
        "— CaddyCheck CRM"
    )
    html = (
        f"<p>A new customer purchase order requires your approval.</p>"
        f"{'<p><i>This is a resend; any earlier approval link for this order no longer works.</i></p>' if resent else ''}"
        f"<ul>"
        f"<li><b>Project / Customer:</b> {project}</li>"
        f"<li><b>Reference:</b> {order.get('order_reference') or 'n/a'}</li>"
        f"<li><b>Amount:</b> {amount_str}</li>"
        f"<li><b>From:</b> {order.get('customer_email') or 'n/a'}</li>"
        f"</ul>"
        f'<p><a href="{approve_link}">Review &amp; decide</a> (link valid for 7 days).</p>'
        f"<p>— CaddyCheck CRM</p>"
    )
    return subject, body, html


def _project_name_from_filename(filename: str) -> str:
    """Derive a human-readable project name from a PDF filename.

    E.g. '2025-01-026_AD_Delhaize_Fernelmont.pdf' → 'AD Delhaize Fernelmont'
    """
    import re
    stem = filename.rsplit(".", 1)[0]  # strip extension
    # Remove leading date/reference patterns like '2025-01-026' or '20260630-055950-'
    stem = re.sub(r"^\d{4}-\d{2}-\d{3,}[-_]?", "", stem)
    stem = re.sub(r"^\d{8}-\d{6}-", "", stem)
    return stem.replace("_", " ").strip()


def _extract_amount_from_text(text: str) -> float:
    """Best-effort extraction of the order total amount from PDF text."""
    import re
    if not text:
        return 0.0
    # Patterns ordered from most specific to least:
    # 1. Explicit total/grand total labels
    # 2. EUR/€ amounts
    patterns = [
        r"(?:grand\s+total|order\s+total|total\s+amount|total\s+excl\.?|subtotal|amount\s+due)[^\d]*([\d]{1,6}[.,][\d]{2,3}(?:[.,][\d]{2})?)",
        r"(?:€|EUR)\s*([\d]{1,6}[.,][\d]{2,3}(?:[.,][\d]{2})?)",
        r"([\d]{1,6}[.,][\d]{3}[.,][\d]{2})\s*(?:€|EUR)",
    ]
    candidates: list[float] = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, re.IGNORECASE):
            token = match.group(1).strip()
            # Normalise European vs US decimal separators
            if token.count(',') == 1 and token.count('.') == 1:
                # e.g. 5.834,00 or 5,834.00
                if token.rfind(',') > token.rfind('.'):
                    token = token.replace('.', '').replace(',', '.')
                else:
                    token = token.replace(',', '')
            elif token.count(',') == 1:
                frac = token.split(',')[1]
                token = token.replace(',', '.') if len(frac) == 2 else token.replace(',', '')
            elif token.count('.') > 1:
                token = token.replace('.', '')
            try:
                candidates.append(float(token))
            except ValueError:
                pass
    if not candidates:
        return 0.0
    # Return the largest value — most likely the grand total
    return round(max(candidates), 2)


def _extract_order_reference(*texts: str) -> str:
    """Best-effort detection of a customer PO/order reference from text."""
    import re

    patterns = [
        r"\b(?:purchase\s+order|order|po)\s*(?:no\.?|number|#|:)?\s*([A-Z0-9][A-Z0-9\-/]{3,})",
        r"\bPO[-#\s]?([0-9]{3,})\b",
        r"\b(\d{4}-\d{2}-\d{3,})\b",  # filename-style reference e.g. 2025-01-026
    ]
    for text in texts:
        if not text:
            continue
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                return match.group(1).strip().upper()
    return ""


# ── Logging, annotations and run summary ──────────────────────────────────────
def _utcnow() -> datetime.datetime:
    return datetime.datetime.now(datetime.timezone.utc)


def _short(value: object, limit: int = 200) -> str:
    return " ".join(str(value).split())[:limit]


def _failure_text(exc: BaseException) -> str:
    """Log-safe description: exception type and HTTP status/code only, never its message."""
    parts = [type(exc).__name__]
    for attr in ("kind", "status", "code"):
        value = getattr(exc, attr, None)
        if isinstance(value, (int, str)) and re.fullmatch(r"[\w.-]{1,40}", str(value)):
            parts.append(f"{attr}={value}")
    return " ".join(parts)


def _msg_ref(message_id: object) -> str:
    """Short stable reference for log lines; avoids logging raw message ids."""
    return hashlib.sha256(str(message_id or "").encode("utf-8")).hexdigest()[:10]


def _kv(event: str, **fields: object) -> str:
    parts = [f"event={event}"]
    for key, value in fields.items():
        if value is None:
            continue
        text = str(value)
        if not text or re.search(r'[\s="]', text):
            text = json.dumps(text)
        parts.append(f"{key}={text}")
    return " ".join(parts)


def _log(event: str, level: int = logging.INFO, **fields: object) -> None:
    logger.log(level, _kv(event, **fields))


def _annotate(level: str, message: str) -> None:
    """Emit a GitHub Actions annotation (no-op outside Actions)."""
    if os.environ.get("GITHUB_ACTIONS") != "true":
        return
    escaped = message.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")
    print(f"::{level}::{escaped}", flush=True)


@dataclass
class RunSummary:
    run_id: str
    mode: str
    since: str = ""
    scanned: int = 0
    filtered_out: int = 0
    skipped_existing: int = 0
    skipped_no_pdf: int = 0
    skipped_in_progress: int = 0
    eligible: int = 0
    resumed: int = 0
    created: int = 0
    emailed: int = 0
    failed: int = 0
    possible_duplicates: int = 0
    abandoned: int = 0
    failures: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    planned: list = field(default_factory=list)
    fatal: str = ""

    @property
    def skipped(self) -> int:
        return self.filtered_out + self.skipped_existing + self.skipped_no_pdf + self.skipped_in_progress

    def record_failure(self, ref: str, step: str, exc: BaseException) -> None:
        self.failed += 1
        self.failures.append((ref, step, _failure_text(exc)))

    def counters(self) -> dict:
        return {
            "scanned": self.scanned,
            "eligible": self.eligible,
            "skipped": self.skipped,
            "created": self.created,
            "emailed": self.emailed,
            "failed": self.failed,
            "resumed": self.resumed,
            "possible_duplicates": self.possible_duplicates,
            "abandoned": self.abandoned,
        }

    def markdown(self) -> str:
        status = "FAILED" if self.fatal or self.failed else "OK"
        lines = [
            f"## Order intake: {status}",
            "",
            f"- Mode: `{self.mode}`",
            f"- Run id: `{self.run_id}`",
            f"- Window start (UTC): `{self.since or 'n/a'}`",
            "",
            "| Metric | Count |",
            "| --- | ---: |",
        ]
        lines += [f"| {name} | {value} |" for name, value in self.counters().items()]
        lines += [
            "",
            f"Skipped detail: already processed {self.skipped_existing}, no PDF {self.skipped_no_pdf}, "
            f"filtered out {self.filtered_out}, held by another run {self.skipped_in_progress}.",
        ]
        if self.fatal:
            lines += ["", f"**Fatal:** {self.fatal}"]
        if self.planned:
            lines += ["", "### Eligible messages (nothing was written)", "",
                      "| Ref | Action | Received (UTC) | PDF bytes |", "| --- | --- | --- | ---: |"]
            lines += [f"| `{ref}` | {action} | {received} | {size} |" for ref, action, received, size in self.planned]
        if self.failures:
            lines += ["", "### Failed messages", "", "| Ref | Step | Error |", "| --- | --- | --- |"]
            lines += [f"| `{ref}` | {step} | {err.replace('|', '/')} |" for ref, step, err in self.failures]
        if self.warnings:
            lines += ["", "### Warnings"] + [f"- {w}" for w in self.warnings]
        return "\n".join(lines) + "\n"


def _emit_summary(summary: RunSummary) -> None:
    _log("summary", mode=summary.mode, run_id=summary.run_id, fatal=summary.fatal or None,
         **summary.counters())
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if path:
        try:
            with open(path, "a", encoding="utf-8") as handle:
                handle.write(summary.markdown())
        except OSError as exc:
            logger.warning("Could not write step summary: %s", type(exc).__name__)


@dataclass
class RunContext:
    run_id: str
    summary: RunSummary
    touched: set = field(default_factory=set)


# ── Configuration and window ──────────────────────────────────────────────────
def _presence(name: str) -> str:
    return "set" if _get_setting(name) else "missing"


def _parse_when(value: str, label: str) -> datetime.datetime:
    try:
        parsed = datetime.datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
    except ValueError:
        raise IntakeError("config", f"{label} must be an ISO-8601 timestamp, e.g. 2026-10-15T08:00:00Z") from None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=datetime.timezone.utc)
    return parsed.astimezone(datetime.timezone.utc)


def _int_setting(cli_value: Optional[int], env_name: str, default: int, label: str, low: int, high: int) -> int:
    raw = cli_value if cli_value is not None else _get_setting(env_name)
    if raw in (None, ""):
        return default
    try:
        value = int(raw)
    except (TypeError, ValueError):
        raise IntakeError("config", f"{label} must be an integer") from None
    if not low <= value <= high:
        raise IntakeError("config", f"{label} must be between {low} and {high}")
    return value


def _resolve_window(args: argparse.Namespace, now: datetime.datetime):
    """Return (since, floor, explicit). ``floor`` is the go-live cutoff for rolling runs."""
    floor_raw = _get_setting("ORDER_INTAKE_MIN_RECEIVED")
    floor = _parse_when(floor_raw, "ORDER_INTAKE_MIN_RECEIVED") if floor_raw else None
    if args.since:
        return _parse_when(args.since, "--since"), floor, True
    hours = _int_setting(args.lookback_hours, "ORDER_INTAKE_LOOKBACK_HOURS",
                         DEFAULT_LOOKBACK_HOURS, "lookback hours", 1, 720)
    since = now - datetime.timedelta(hours=hours)
    if floor and floor > since:
        since = floor
    return since, floor, False


def _matches_filters(header: MessageHeader, args: argparse.Namespace) -> bool:
    if args.only_from and not any(s.lower() in header.from_address.lower() for s in args.only_from):
        return False
    if args.subject_contains and args.subject_contains.lower() not in header.subject.lower():
        return False
    return True


def _log_config(provider, mode: str, since: datetime.datetime, floor, explicit: bool) -> None:
    mailbox = str(getattr(provider, "mailbox", "") or "")
    _log(
        "config",
        provider=type(provider).__name__,
        mode=mode,
        folder=getattr(provider, "folder", None),
        mailbox_domain=mailbox.split("@", 1)[1] if "@" in mailbox else ("set" if mailbox else "missing"),
        graph_tenant_id=_presence("ORDER_INTAKE_GRAPH_TENANT_ID"),
        graph_client_id=_presence("ORDER_INTAKE_GRAPH_CLIENT_ID"),
        graph_client_secret=_presence("ORDER_INTAKE_GRAPH_CLIENT_SECRET"),
        supabase_url="set" if os.environ.get("SUPABASE_URL", "").strip() else "missing",
        supabase_service_key="set" if (os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
                                       or os.environ.get("SUPABASE_KEY", "").strip()) else "missing",
        app_base_url=_presence("APP_BASE_URL"),
        ceo_recipients=len(_ceo_recipients()),
        window="explicit" if explicit else "rolling",
        since=since.isoformat(),
        min_received=floor.isoformat() if floor else "unset",
    )


def _validate_live(args: argparse.Namespace, floor, explicit: bool, now: datetime.datetime) -> None:
    if explicit:
        if not (args.only_from or args.subject_contains):
            raise IntakeError(
                "config",
                "--since in live mode requires --only-from and/or --subject-contains, "
                "so a backfill cannot import unrelated historical email",
            )
        filters = list(args.only_from) + ([args.subject_contains] if args.subject_contains else [])
        if any(len(f.strip()) < MIN_FILTER_LENGTH for f in filters) or any(
            "@" not in f and "." not in f for f in args.only_from
        ):
            raise IntakeError(
                "config",
                f"backfill filters are too broad: use a full sender address/domain and at least "
                f"{MIN_FILTER_LENGTH} characters",
            )
        if _parse_when(args.since, "--since") < now - datetime.timedelta(days=MAX_BACKFILL_DAYS):
            raise IntakeError("config", f"--since may not be older than {MAX_BACKFILL_DAYS} days")
    elif floor is None:
        raise IntakeError(
            "config",
            "ORDER_INTAKE_MIN_RECEIVED must be set for live processing (go-live timestamp), "
            "so the lookback cannot import unrelated historical email",
        )
    if not _ceo_recipients():
        raise IntakeError("config", "ORDER_APPROVAL_CEO_EMAIL is not set")
    if not _get_setting("APP_BASE_URL"):
        raise IntakeError("config", "APP_BASE_URL is not set; approval links would be unusable")


# ── Message processing (each step is idempotent and resumable) ───────────────
def _create_order(row: dict, ctx: RunContext, ref: str) -> dict:
    extracted_text = row.get("extracted_text") or ""
    pdf_filename = row.get("pdf_filename") or ""
    order_reference = _extract_order_reference(row.get("subject") or "", extracted_text, pdf_filename)
    previous = None
    if order_reference:
        try:
            previous = approvals.find_current_open_order_by_reference(order_reference)
        except Exception as exc:
            _log("revision_lookup_failed", logging.WARNING, ref=ref, error=type(exc).__name__)
    revision = (previous.get("revision") or 1) + 1 if previous else 1
    order, created = approvals.get_or_create_purchase_order_for_incoming(
        row["id"],
        dict(
            customer_email=row.get("from_address"),
            project_name=_project_name_from_filename(pdf_filename) or None,
            order_reference=order_reference or None,
            summary=(row.get("subject") or "")[:500] or None,
            pdf_storage_bucket=row.get("pdf_storage_bucket"),
            pdf_storage_path=row.get("pdf_storage_path"),
            amount=_extract_amount_from_text(extracted_text) or None,
            revision=revision,
            status="pending_approval",
        ),
    )
    if created:
        ctx.summary.created += 1
        _log("order_created", ref=ref, order_id=order.get("id"), revision=revision)
    return order


def _supersede_older_revisions(order: dict, ref: str) -> None:
    reference = order.get("order_reference")
    if not reference or order.get("status") != "pending_approval":
        return
    previous = approvals.find_current_open_order_by_reference(reference, exclude_id=order["id"])
    if previous and previous.get("id") and (previous.get("revision") or 1) < (order.get("revision") or 1):
        approvals.supersede_purchase_order(previous["id"], order["id"])
        _log("order_superseded", ref=ref, order_id=order["id"], previous_order_id=previous["id"])


def _send_approval(order: dict, row: dict, ctx: RunContext, ref: str, pdf_bytes: Optional[bytes]) -> None:
    prep = approvals.prepare_approval_email(order["id"])
    state = prep["state"]
    if state in ("already_sent", "not_needed"):
        _log("approval_email_skipped", ref=ref, order_id=order["id"], state=state)
        return
    if state == "in_progress":
        raise RuntimeError("approval email is being sent by another worker or an earlier attempt")

    approval = prep["approval"]
    recipients = _ceo_recipients()
    if not recipients:
        approvals.mark_approval_email_failed(approval["id"], "ORDER_APPROVAL_CEO_EMAIL is not set")
        raise RuntimeError("ORDER_APPROVAL_CEO_EMAIL is not set")

    attachment = pdf_bytes or approvals.download_purchase_order_pdf_bytes(
        order.get("pdf_storage_bucket"), order.get("pdf_storage_path")
    )
    if attachment is None:
        _log("approval_attachment_unavailable", logging.WARNING, ref=ref, order_id=order["id"])
    subject, body, html = _build_approval_email(order, prep["token"], resent=prep["attempt"] > 1)
    try:
        send_graph_email(
            subject, body, recipients, html_body=html,
            attachment_bytes=attachment,
            attachment_filename=row.get("pdf_filename") or "purchase_order.pdf",
        )
    except Exception as exc:
        approvals.mark_approval_email_failed(approval["id"], f"{type(exc).__name__}: {_short(exc, 300)}")
        raise
    approvals.mark_approval_email_sent(approval["id"])
    ctx.summary.emailed += 1
    _log("approval_email_sent", ref=ref, order_id=order["id"], attempt=prep["attempt"])


def _advance(row: dict, ctx: RunContext, pdf_bytes: Optional[bytes] = None) -> bool:
    """Run the remaining steps for a leased intake row. Returns True when fully processed."""
    ref = _msg_ref(row.get("message_id"))
    ctx.touched.add(row["id"])
    step = "purchase_order"
    started = time.monotonic()
    try:
        order = approvals.get_purchase_order_for_incoming(row["id"]) or _create_order(row, ctx, ref)
        step = "supersede"
        _supersede_older_revisions(order, ref)
        step = "approval_email"
        _send_approval(order, row, ctx, ref, pdf_bytes)
        step = "finalize"
        approvals.finish_incoming_email(row["id"], ctx.run_id)
    except Exception as exc:
        attempts = int(row.get("attempt_count") or 0) + 1
        try:
            approvals.fail_incoming_email(
                row["id"], ctx.run_id, f"{step}: {type(exc).__name__}: {_short(exc, 300)}", attempts
            )
        except Exception as inner:
            _log("fail_record_error", logging.ERROR, ref=ref, error=type(inner).__name__)
        ctx.summary.record_failure(ref, step, exc)
        _log("message_failed", logging.ERROR, ref=ref, step=step, attempt=attempts, error=_failure_text(exc))
        _annotate("error", f"Order intake failed at step '{step}' for message {ref}: {type(exc).__name__}")
        return False
    _log("message_done", ref=ref, order_id=order.get("id"), duration_s=round(time.monotonic() - started, 2))
    return True


def _ingest_new(message: IncomingMessage, ctx: RunContext) -> None:
    ref = _msg_ref(message.message_id)
    sha256 = hashlib.sha256(message.pdf_bytes).hexdigest()
    try:
        storage = approvals.upload_purchase_order_pdf(
            message.pdf_bytes,
            message.pdf_filename,
            storage_path=approvals.deterministic_pdf_storage_path(
                message.pdf_filename, sha256, message.received_at
            ),
        )
        extracted_text = extract_pdf_text(message.pdf_bytes)
        duplicate = approvals.find_incoming_by_pdf_sha256(sha256, exclude_message_id=message.message_id)
        row = approvals.insert_incoming_email_claimed(
            {
                "message_id": message.message_id,
                "graph_message_id": message.raw_id or None,
                "provider": message.provider,
                "from_address": message.from_address,
                "subject": message.subject,
                "received_at": message.received_at.isoformat() if message.received_at else None,
                "body_text": message.body_text or None,
                "pdf_filename": message.pdf_filename,
                "pdf_storage_bucket": storage["pdf_storage_bucket"],
                "pdf_storage_path": storage["pdf_storage_path"],
                "pdf_sha256": sha256,
                "duplicate_of_email_id": duplicate["id"] if duplicate else None,
                "extracted_text": extracted_text or None,
            },
            ctx.run_id,
        )
    except Exception as exc:
        ctx.summary.record_failure(ref, "ingest", exc)
        _log("message_failed", logging.ERROR, ref=ref, step="ingest", error=_failure_text(exc))
        _annotate("error", f"Order intake failed at step 'ingest' for message {ref}: {type(exc).__name__}")
        return
    if row is None:
        ctx.summary.skipped_in_progress += 1
        _log("skipped_claimed_elsewhere", ref=ref)
        return

    ctx.summary.eligible += 1
    sender_domain = message.from_address.split("@", 1)[-1] if "@" in message.from_address else "unknown"
    _log("message_ingested", ref=ref, received_at=message.received_at, pdf_bytes=len(message.pdf_bytes))
    if duplicate:
        ctx.summary.possible_duplicates += 1
        ctx.summary.warnings.append(f"Message {ref} carries a PDF identical to an earlier intake (flagged, still processed).")
        _log("possible_duplicate_pdf", logging.WARNING, ref=ref, duplicate_of=duplicate["id"])
        approvals.create_notification(
            title="Possible duplicate purchase-order PDF",
            message=(f"An email from {sender_domain} carries a PDF identical to an earlier intake. "
                     "It was processed normally; review it before approving."),
            severity="warning",
        )
    _advance(row, ctx, pdf_bytes=message.pdf_bytes)


def _resume_rows(rows: list, ctx: RunContext) -> None:
    for row in rows:
        if row["id"] in ctx.touched:
            continue
        ctx.touched.add(row["id"])
        claimed = approvals.claim_incoming_email(row["id"], ctx.run_id)
        if not claimed:
            ctx.summary.skipped_in_progress += 1
            continue
        ctx.summary.eligible += 1
        ctx.summary.resumed += 1
        _log("message_resuming", ref=_msg_ref(row.get("message_id")), attempt=int(row.get("attempt_count") or 0) + 1)
        _advance(claimed, ctx)
        ctx.touched.add(row["id"])


# ── Run ───────────────────────────────────────────────────────────────────────
def _run(args: argparse.Namespace, summary: RunSummary, ctx: RunContext) -> int:
    live = summary.mode == "live"
    now = _utcnow()
    try:
        provider = get_email_provider()
    except RuntimeError as exc:
        raise IntakeError("config", str(exc)) from None
    if not hasattr(provider, "list_message_headers"):
        raise IntakeError(
            "config",
            f"Provider '{provider.name}' cannot poll without read/unread state; set ORDER_INTAKE_PROVIDER=graph",
        )

    since, floor, explicit = _resolve_window(args, now)
    summary.since = since.isoformat()
    _log_config(provider, summary.mode, since, floor, explicit)

    if not provider.is_configured():
        raise IntakeError("config", "Mailbox is not configured; missing: " + ", ".join(provider.missing_settings()))
    if live:
        _validate_live(args, floor, explicit, now)

    info = provider.preflight()
    _log("preflight_ok", folder=info.get("folder"), total_items=info.get("total_items"))

    ready, detail = approvals.schema_ready()
    _log("schema_check", level=logging.INFO if ready else logging.WARNING, ready=ready, detail=detail)
    if not ready:
        if live:
            raise IntakeError("config", f"Database migration not applied: {detail}")
        summary.warnings.append(f"Database migration not applied ({detail}); live processing is blocked until it is.")
        _annotate("warning", f"Order intake migration not applied: {detail}")
    if args.check:
        return EXIT_OK

    headers = provider.list_message_headers(since)
    summary.scanned = len(headers)
    max_new = _int_setting(args.max_new, "ORDER_INTAKE_MAX_NEW_PER_RUN", DEFAULT_MAX_NEW, "max new", 1, 1000)

    new_messages: list[IncomingMessage] = []
    resume_rows: dict[str, dict] = {}
    for header in headers:
        ref = _msg_ref(header.message_id)
        if not _matches_filters(header, args):
            summary.filtered_out += 1
            continue
        existing = approvals.find_incoming_email(header.message_id, header.graph_id if ready else "")
        if existing:
            if existing.get("processing_status") == "processed":
                summary.skipped_existing += 1
            elif int(existing.get("attempt_count") or 0) >= approvals.MAX_INTAKE_ATTEMPTS:
                summary.skipped_existing += 1
            else:
                resume_rows[existing["id"]] = existing
            continue
        try:
            message = provider.fetch_pdf_message(header)
        except IntakeError as exc:
            if exc.kind in ("auth", "config"):
                raise
            summary.record_failure(ref, "fetch", exc)
            _log("message_failed", logging.ERROR, ref=ref, step="fetch", error=_failure_text(exc))
            continue
        if message is None:
            summary.skipped_no_pdf += 1
            continue
        new_messages.append(message)

    if live and len(new_messages) > max_new:
        raise IntakeError(
            "config",
            f"{len(new_messages)} new messages exceed the safety limit of {max_new}; "
            "nothing was processed. Check the window/floor, or raise --max-new deliberately",
        )

    if not explicit and ready and floor is not None:
        for row in approvals.list_unfinished_incoming(floor):
            resume_rows.setdefault(row["id"], row)
        summary.abandoned = len(approvals.list_abandoned_incoming(floor))
        if summary.abandoned:
            summary.warnings.append(f"{summary.abandoned} message(s) exceeded the retry limit and need manual attention.")
            _annotate("warning", f"Order intake: {summary.abandoned} message(s) exceeded the retry limit")

    if not live:
        for message in new_messages:
            _log("would_ingest", ref=_msg_ref(message.message_id), received_at=message.received_at,
                 pdf_bytes=len(message.pdf_bytes))
            received = message.received_at.astimezone(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M") \
                if message.received_at else "unknown"
            summary.planned.append((_msg_ref(message.message_id), "new", received, len(message.pdf_bytes)))
        for row in resume_rows.values():
            _log("would_resume", ref=_msg_ref(row.get("message_id")), status=row.get("processing_status"))
            summary.planned.append((_msg_ref(row.get("message_id")), "resume", str(row.get("received_at") or "unknown")[:16], "-"))
        summary.eligible = len(new_messages) + len(resume_rows)
        return EXIT_OK

    for message in new_messages:
        _ingest_new(message, ctx)
    _resume_rows(list(resume_rows.values()), ctx)
    return EXIT_PARTIAL if summary.failed else EXIT_OK


def run(args: argparse.Namespace) -> int:
    mode = "check" if args.check else "dry-run" if args.dry_run else "live"
    run_id = f"{os.environ.get('GITHUB_RUN_ID') or 'local'}-{uuid.uuid4().hex[:8]}"
    summary = RunSummary(run_id=run_id, mode=mode)
    ctx = RunContext(run_id=run_id, summary=summary)
    try:
        code = _run(args, summary, ctx)
        if code == EXIT_OK and summary.failed:
            code = EXIT_PARTIAL
    except IntakeError as exc:
        summary.fatal = f"[{exc.kind}] {_short(exc)}"
        _log("fatal", logging.ERROR, kind=exc.kind, status=exc.status, error=_short(exc))
        _annotate("error", f"Order intake failed ({exc.kind}): {_short(exc)}")
        code = EXIT_FATAL
    except Exception as exc:  # noqa: BLE001 - any unexpected error must fail the run
        summary.fatal = f"unexpected {type(exc).__name__}"
        _log("fatal", logging.ERROR, kind="unexpected", error=summary.fatal)
        _annotate("error", f"Order intake failed unexpectedly: {type(exc).__name__}")
        code = EXIT_FATAL
    _emit_summary(summary)
    return code


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Poll the order mailbox for purchase-order emails.")
    parser.add_argument("--check", action="store_true",
                        help="Verify configuration, mailbox access and database schema, then exit.")
    parser.add_argument("--dry-run", action="store_true",
                        help="List what would be processed without writing anything.")
    parser.add_argument("--since", help="Explicit window start (ISO-8601) for a targeted backfill; "
                                        "in live mode requires --only-from and/or --subject-contains.")
    parser.add_argument("--only-from", action="append", default=[], metavar="TEXT",
                        help="Only messages whose sender address contains TEXT (repeatable).")
    parser.add_argument("--subject-contains", metavar="TEXT", help="Only messages whose subject contains TEXT.")
    parser.add_argument("--lookback-hours", type=int, help=f"Rolling window in hours (default {DEFAULT_LOOKBACK_HOURS}).")
    parser.add_argument("--max-new", type=int, help=f"Abort live runs that would ingest more than N new messages (default {DEFAULT_MAX_NEW}).")
    return run(parser.parse_args(argv))


if __name__ == "__main__":
    sys.exit(main())
