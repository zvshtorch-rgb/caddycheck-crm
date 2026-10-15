"""In-memory fakes for order-intake tests (Supabase client and mailbox provider)."""
import copy
import datetime
import threading
import uuid

from services.email_intake_service import IncomingMessage, IntakeError, MessageHeader


class FakeResponse:
    def __init__(self, data):
        self.data = data


# (table, columns, predicate) -> partial unique indexes mirroring the migration.
_UNIQUE = [
    ("incoming_order_emails", ("message_id",), lambda r: True),
    ("incoming_order_emails", ("graph_message_id",), lambda r: True),
    ("purchase_orders", ("incoming_email_id",), lambda r: True),
    ("order_approvals", ("token_hash",), lambda r: True),
    ("order_approvals", ("purchase_order_id",), lambda r: r.get("status") == "pending"),
]


def _now_iso():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def _defaults(table):
    base = {"id": str(uuid.uuid4()), "created_at": _now_iso()}
    if table == "incoming_order_emails":
        base.update(attempt_count=0, claimed_by=None, claimed_at=None, graph_message_id=None,
                    duplicate_of_email_id=None, error_message=None)
    elif table == "purchase_orders":
        base.update(status="pending_approval", revision=1, is_current=True, incoming_email_id=None)
    elif table == "order_approvals":
        base.update(email_status="pending", email_attempts=0, email_sent_at=None,
                    email_claimed_at=None, email_last_error=None)
    elif table == "crm_notifications":
        base.update(is_read=False)
    return base


class _Query:
    def __init__(self, db, table):
        self._db, self._table = db, table
        self._op, self._payload = "select", None
        self._filters, self._order, self._limit = [], None, None

    def select(self, _cols="*"):
        self._op = "select"
        return self

    def insert(self, row):
        self._op, self._payload = "insert", row
        return self

    def update(self, values):
        self._op, self._payload = "update", values
        return self

    def delete(self):
        self._op = "delete"
        return self

    def eq(self, col, val):
        self._filters.append(lambda r: r.get(col) == val)
        return self

    def neq(self, col, val):
        self._filters.append(lambda r: r.get(col) != val)
        return self

    def in_(self, col, vals):
        self._filters.append(lambda r: r.get(col) in vals)
        return self

    def is_(self, col, val):
        assert val == "null"
        self._filters.append(lambda r: r.get(col) is None)
        return self

    def lt(self, col, val):
        self._filters.append(lambda r: r.get(col) is not None and r[col] < val)
        return self

    def gte(self, col, val):
        self._filters.append(lambda r: r.get(col) is not None and r[col] >= val)
        return self

    def order(self, col, desc=False):
        self._order = (col, desc)
        return self

    def limit(self, n):
        self._limit = n
        return self

    def execute(self):
        with self._db.lock:
            rows = self._db.tables.setdefault(self._table, [])
            matching = [r for r in rows if all(f(r) for f in self._filters)]
            if self._op == "insert":
                row = {**_defaults(self._table), **copy.deepcopy(self._payload)}
                self._db.check_unique(self._table, row, ignore=None)
                rows.append(row)
                return FakeResponse([copy.deepcopy(row)])
            if self._op == "update":
                updated = []
                for r in matching:
                    candidate = {**r, **copy.deepcopy(self._payload)}
                    self._db.check_unique(self._table, candidate, ignore=r)
                    r.update(copy.deepcopy(self._payload))
                    updated.append(copy.deepcopy(r))
                return FakeResponse(updated)
            if self._op == "delete":
                for r in matching:
                    rows.remove(r)
                return FakeResponse([copy.deepcopy(r) for r in matching])
            if self._order:
                col, desc = self._order
                matching.sort(key=lambda r: (r.get(col) is None, r.get(col)), reverse=desc)
            if self._limit is not None:
                matching = matching[: self._limit]
            return FakeResponse(copy.deepcopy(matching))


class _Bucket:
    def __init__(self, storage, name):
        self._storage, self._name = storage, name

    def remove(self, paths):
        for p in paths:
            self._storage.files.pop((self._name, p), None)

    def upload(self, path, file, file_options=None):
        with open(file, "rb") as handle:
            self._storage.files[(self._name, path)] = handle.read()

    def download(self, path):
        return self._storage.files[(self._name, path)]

    def create_signed_url(self, path, _expires):
        return {"signedURL": f"https://signed.example/{path}"}


class _Storage:
    def __init__(self):
        self.files = {}

    def list_buckets(self):
        return [{"name": "purchase-order-pdfs"}]

    def create_bucket(self, *_a, **_k):
        return None

    def from_(self, name):
        return _Bucket(self, name)


class _Rpc:
    def __init__(self, db, name):
        self._db, self._name = db, name

    def execute(self):
        if self._db.schema_version is None:
            raise RuntimeError("function order_intake_schema_version() does not exist")
        return FakeResponse(self._db.schema_version)


class FakeSupabase:
    def __init__(self, schema_version=20261015):
        self.tables = {}
        self.lock = threading.RLock()
        self.schema_version = schema_version
        self.storage = _Storage()

    def table(self, name):
        return _Query(self, name)

    def rpc(self, name):
        return _Rpc(self, name)

    def check_unique(self, table, row, ignore):
        for tbl, cols, pred in _UNIQUE:
            if tbl != table or not pred(row):
                continue
            key = tuple(row.get(c) for c in cols)
            if any(v is None for v in key):
                continue
            for other in self.tables.get(table, []):
                if other is ignore or not pred(other):
                    continue
                if tuple(other.get(c) for c in cols) == key:
                    raise Exception(f'duplicate key value violates unique constraint on {cols} (code 23505)')

    def rows(self, table):
        return self.tables.get(table, [])


NOW = datetime.datetime(2026, 10, 15, 13, 0, tzinfo=datetime.timezone.utc)
FLOOR = "2026-10-15T08:00:00Z"


def make_message(n=1, *, sender="Pat Buyer <pat.buyer@customer-a.example>", subject=None,
                 pdf=None, received=None, has_pdf=True, is_read=False, filename=None,
                 message_id=None, graph_id=None):
    return {
        "n": n,
        "message_id": message_id or f"<msg{n}@mail.example>",
        "graph_id": graph_id or f"AAMk-immutable-{n}",
        "from": sender,
        "subject": subject or f"New order {n}",
        "pdf": pdf if pdf is not None else f"%PDF-1.4 order {n}".encode(),
        "received": received or datetime.datetime(2026, 10, 15, 12, 15, tzinfo=datetime.timezone.utc),
        "has_pdf": has_pdf,
        "is_read": is_read,
        "filename": filename or f"2026-10-{n:03d}_Test_Order_{n}.pdf",
    }


class FakeProvider:
    """Graph-provider stand-in. Read state is stored but must never matter or change."""

    name = "graph"
    mailbox = "orders@example.com"
    folder = "Inbox"

    def __init__(self, messages=(), preflight_error=None, list_barrier=None):
        self.messages = list(messages)
        self.preflight_error = preflight_error
        self.list_barrier = list_barrier
        self.list_calls = []
        self.fetch_calls = 0

    def is_configured(self):
        return True

    def missing_settings(self):
        return []

    def preflight(self):
        if self.preflight_error:
            raise self.preflight_error
        return {"folder": "Inbox", "total_items": len(self.messages), "unread_items": 0}

    def list_message_headers(self, since):
        self.list_calls.append(since)
        if self.list_barrier:
            self.list_barrier.wait(timeout=5)
        return [
            MessageHeader(
                message_id=m["message_id"], graph_id=m["graph_id"], from_address=m["from"].split("<")[-1].rstrip(">"),
                subject=m["subject"], received_at=m["received"], body_preview="PRIVATE BODY TEXT",
            )
            for m in self.messages
            if m["received"] >= since
        ]

    def fetch_pdf_message(self, header):
        self.fetch_calls += 1
        for m in self.messages:
            if m["graph_id"] == header.graph_id:
                if not m["has_pdf"]:
                    return None
                return IncomingMessage(
                    message_id=m["message_id"], provider="graph", from_address=header.from_address,
                    subject=m["subject"], received_at=m["received"], body_text="PRIVATE BODY TEXT",
                    pdf_filename=m["filename"], pdf_bytes=m["pdf"], raw_id=m["graph_id"],
                )
        raise IntakeError("mailbox", "message vanished", status=404)

    def mark_processed(self, _message):
        raise AssertionError("the poller must never change read/unread state")


class FakeMailer:
    """Replaces send_graph_email; can be told to fail the next N sends."""

    def __init__(self, delay=0.0):
        self.sent = []
        self.fail_next = 0
        self.delay = delay
        self.lock = threading.Lock()

    def __call__(self, subject, body, recipients, **kwargs):
        import time

        if self.delay:
            time.sleep(self.delay)
        with self.lock:
            if self.fail_next > 0:
                self.fail_next -= 1
                raise RuntimeError("Graph sendMail failed: HTTP 503")
            self.sent.append({"subject": subject, "body": body, "recipients": recipients, **kwargs})
