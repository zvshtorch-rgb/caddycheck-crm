"""Ask Data Phase 5: pure, deterministic historical-state and debt-bridge logic.

No I/O here -- every function takes already-fetched data (snapshot run rows, one day's snapshot rows,
invoice_change_log rows, bank_payments/bank_payment_allocations rows) and returns plain dataclasses. The
lazy Supabase I/O itself lives in executor.py's dispatch block (mirroring the existing get_last_bank_payment
pattern: a lazy, try/except-wrapped import, never a hard dependency of the executor module). Keeping all
the actual arithmetic here, free of I/O, makes it directly unit-testable with plain fixtures.

Core design decisions (approved):
  * country/network are resolved LIVE (current ProjectIndex/NetworkIndex) against each snapshot row's
    stored `project_name`, falling back to the row's own frozen `country_at_capture`/`network_at_capture`
    only when live resolution fails, with an explicit warning on disagreement.
  * "debt at a date" = sum of `payment_amount` for snapshot rows whose `paid` == 'No', matching
    `Invoice.is_unpaid()` exactly.
  * A snapshot that is a row-for-row EXACT match for the requested date is "exact"; otherwise the nearest
    earlier successful run is used and every answer must say so explicitly, including both dates/timestamps.
  * The bridge never hides a residual: `unexplained = closing - (opening + new - payments - credit_notes
    - cancellations - removed + adjustments)`. A non-zero residual is always shown, never silently folded
    into another bucket.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from typing import Callable, Optional

from services.ask_data.money import ZERO, to_money
from services.ask_data.resolver import NetworkIndex, ProjectIndex


def _row_amount(row: dict) -> Decimal:
    return to_money(row.get("payment_amount"))


def _row_paid(row: dict) -> str:
    return str(row.get("paid") or "").strip().lower()


def _row_key(row: dict):
    return row.get("invoice_row_id")


@dataclass(frozen=True)
class SnapshotRun:
    snapshot_date: date
    captured_at: Optional[datetime]
    status: str
    row_count: Optional[int] = None


def parse_snapshot_runs(raw_runs: list[dict]) -> list[SnapshotRun]:
    runs = []
    for row in raw_runs or []:
        sd = row.get("snapshot_date")
        if not sd:
            continue
        snapshot_date = sd if isinstance(sd, date) else date.fromisoformat(str(sd)[:10])
        captured_raw = row.get("completed_at") or row.get("started_at")
        captured_at = None
        if captured_raw:
            try:
                captured_at = datetime.fromisoformat(str(captured_raw).replace("Z", "+00:00"))
            except Exception:
                captured_at = None
        runs.append(SnapshotRun(snapshot_date, captured_at, str(row.get("status") or ""), row.get("row_count")))
    return sorted(runs, key=lambda r: r.snapshot_date)


def earliest_available_date(runs: list[SnapshotRun]) -> Optional[date]:
    successful = [r.snapshot_date for r in runs if r.status == "success"]
    return min(successful) if successful else None


def nearest_run_on_or_before(runs: list[SnapshotRun], target: date) -> Optional[SnapshotRun]:
    """The latest successful run with snapshot_date <= target, or None if history does not reach back
    that far (never looks ahead -- never extrapolates forward past `target`)."""
    candidates = [r for r in runs if r.status == "success" and r.snapshot_date <= target]
    return max(candidates, key=lambda r: r.snapshot_date) if candidates else None


@dataclass
class SnapshotLookup:
    """One resolved point-in-time snapshot: the date/time actually used, whether it was an exact match for
    the request, and its rows."""
    requested_date: date
    actual_date: Optional[date]
    captured_at: Optional[datetime]
    rows: list = field(default_factory=list)
    available: bool = False

    @property
    def exact(self) -> bool:
        return self.available and self.actual_date == self.requested_date

    def coverage_note(self) -> Optional[str]:
        """Required wording when the exact requested date has no snapshot (goal: never silently
        substitute)."""
        if not self.available:
            return None
        if self.exact:
            return None
        when = self.captured_at.strftime("%Y-%m-%d %H:%M UTC") if self.captured_at else str(self.actual_date)
        return (
            f"No snapshot exists for {self.requested_date.isoformat()}; using the latest available "
            f"snapshot from {when}."
        )


def resolve_snapshot(
    runs: list[SnapshotRun], requested_date: date, load_day: Callable[[date], list],
) -> SnapshotLookup:
    """`load_day(date) -> list[dict]` is the caller's (executor's) already-fetched-or-lazy row loader for
    one calendar day -- never fetches more than one day's rows per lookup."""
    run = nearest_run_on_or_before(runs, requested_date)
    if run is None:
        return SnapshotLookup(requested_date, None, None, [], available=False)
    rows = load_day(run.snapshot_date)
    return SnapshotLookup(requested_date, run.snapshot_date, run.captured_at, rows, available=True)


@dataclass
class ResolvedRow:
    row: dict
    country: Optional[str]
    network: Optional[str]
    country_mismatch: bool = False
    network_mismatch: bool = False


def resolve_snapshot_rows(rows: list[dict], projects) -> list[ResolvedRow]:
    """Live country/network resolution against each row's stored project_name (reusing the CURRENT
    ProjectIndex/NetworkIndex, per the approved design), falling back to the row's own frozen
    country_at_capture/network_at_capture only when live resolution fails, flagging any disagreement."""
    project_index = ProjectIndex(projects)
    network_index = NetworkIndex(projects)
    resolved = []
    for row in rows:
        name = row.get("project_name")
        live_country = project_index.country(name)
        canonical = project_index.canonical_name(name)
        live_network = network_index.network_of(canonical) if canonical else None
        frozen_country = row.get("country_at_capture")
        frozen_network = row.get("network_at_capture")
        country = live_country if live_country is not None else frozen_country
        network = live_network if live_network is not None else frozen_network
        country_mismatch = bool(live_country and frozen_country and live_country != frozen_country)
        network_mismatch = bool(live_network and frozen_network and live_network != frozen_network)
        resolved.append(ResolvedRow(row, country, network, country_mismatch, network_mismatch))
    return resolved


def apply_filters(resolved_rows: list[ResolvedRow], projects, args: dict) -> list[ResolvedRow]:
    project_index = ProjectIndex(projects) if args.get("project") else None
    out = resolved_rows
    if args.get("country"):
        out = [r for r in out if r.country == args["country"]]
    if args.get("network"):
        out = [r for r in out if r.network == args["network"]]
    if args.get("project") and project_index is not None:
        target = args["project"]
        out = [r for r in out if project_index.canonical_name(r.row.get("project_name")) == target]
    if args.get("year") is not None:
        out = [r for r in out if r.row.get("year") == args["year"]]
    category = args.get("category")
    if category == "trial":
        out = [r for r in out if "paid trial" in str(r.row.get("maintenance_year") or "").strip().lower()]
    elif category == "y1":
        out = [r for r in out if str(r.row.get("maintenance_year") or "").strip().lower() == "y1"]
    elif category == "y2plus":
        out = [r for r in out if str(r.row.get("maintenance_year") or "").strip().lower() not in ("y1",)
               and "paid trial" not in str(r.row.get("maintenance_year") or "").strip().lower()]
    return out


def debt_total(resolved_rows: list[ResolvedRow]) -> Decimal:
    return sum((_row_amount(r.row) for r in resolved_rows if _row_paid(r.row) == "no"), ZERO)


def mismatch_warnings(resolved_rows: list[ResolvedRow]) -> list[str]:
    mismatched = [r for r in resolved_rows if r.country_mismatch or r.network_mismatch]
    if not mismatched:
        return []
    names = sorted({r.row.get("project_name") for r in mismatched})[:5]
    return [
        f"{len(mismatched)} snapshot row(s) resolve to a different country/network today than what was "
        f"captured at snapshot time ({', '.join(names)}); the current resolver's classification was used."
    ]


@dataclass
class BridgeResult:
    opening: Decimal
    closing: Decimal
    new_invoices: Decimal = ZERO
    new_invoices_rows: list = field(default_factory=list)
    payments: Decimal = ZERO
    payments_rows: list = field(default_factory=list)
    credit_notes: Decimal = ZERO
    credit_notes_rows: list = field(default_factory=list)
    cancellations: Decimal = ZERO
    cancellations_rows: list = field(default_factory=list)
    adjustments: Decimal = ZERO
    adjustments_rows: list = field(default_factory=list)
    removed: Decimal = ZERO
    removed_rows: list = field(default_factory=list)
    unexplained: Decimal = ZERO
    warnings: list = field(default_factory=list)

    @property
    def reconciles(self) -> bool:
        return self.unexplained == ZERO

    @property
    def change(self) -> Decimal:
        return self.closing - self.opening


def compute_bridge(
    opening_rows: list[ResolvedRow], closing_rows: list[ResolvedRow], *, live_closing: Optional[Decimal] = None,
) -> BridgeResult:
    """Opening/closing debt plus the deterministic opening->closing bridge:
    opening + new_invoices - payments - credit_notes - cancellations - removed + adjustments == closing
    (any residual is reported as `unexplained`, never hidden)."""
    opening_by_id = {_row_key(r.row): r for r in opening_rows if _row_key(r.row) is not None}
    closing_by_id = {_row_key(r.row): r for r in closing_rows if _row_key(r.row) is not None}

    opening_debt = debt_total(opening_rows)
    closing_debt = live_closing if live_closing is not None else debt_total(closing_rows)

    new_rows, payment_rows, credit_rows, cancel_rows, adjust_rows, removed_rows = [], [], [], [], [], []
    new_total = payments_total = credits_total = cancel_total = adjust_total = removed_total = ZERO

    for key, closing_r in closing_by_id.items():
        opening_r = opening_by_id.get(key)
        closing_amount = _row_amount(closing_r.row)
        closing_paid = _row_paid(closing_r.row)
        if opening_r is None:
            # A genuinely new invoice row this period.
            if closing_paid == "no":
                if closing_amount < 0:
                    credits_total += closing_amount
                    credit_rows.append(closing_r.row)
                else:
                    new_total += closing_amount
                    new_rows.append(closing_r.row)
            continue
        opening_amount = _row_amount(opening_r.row)
        opening_paid = _row_paid(opening_r.row)
        if opening_paid == "no" and closing_paid == "no":
            if closing_amount != opening_amount:
                adjust_total += closing_amount - opening_amount
                adjust_rows.append(closing_r.row)
        elif opening_paid == "no" and closing_paid == "yes":
            payments_total += opening_amount
            payment_rows.append(closing_r.row)
        elif opening_paid == "no" and closing_paid == "cancelled":
            cancel_total += opening_amount
            cancel_rows.append(closing_r.row)
        elif opening_paid != "no" and closing_paid == "no":
            # Reopened/corrected after being paid or cancelled -- treat as a debt-increasing adjustment.
            adjust_total += closing_amount
            adjust_rows.append(closing_r.row)

    for key, opening_r in opening_by_id.items():
        if key not in closing_by_id and _row_paid(opening_r.row) == "no":
            removed_total += _row_amount(opening_r.row)
            removed_rows.append(opening_r.row)

    bridge = BridgeResult(
        opening=opening_debt, closing=closing_debt,
        new_invoices=new_total, new_invoices_rows=new_rows,
        payments=payments_total, payments_rows=payment_rows,
        credit_notes=credits_total, credit_notes_rows=credit_rows,
        cancellations=cancel_total, cancellations_rows=cancel_rows,
        adjustments=adjust_total, adjustments_rows=adjust_rows,
        removed=removed_total, removed_rows=removed_rows,
    )
    predicted_closing = (
        bridge.opening + bridge.new_invoices - bridge.payments - bridge.credit_notes
        - bridge.cancellations - bridge.removed + bridge.adjustments
    )
    bridge.unexplained = bridge.closing - predicted_closing
    if removed_rows:
        names = ", ".join(sorted({r.get("project_name", "") for r in removed_rows})[:5])
        bridge.warnings.append(
            f"{len(removed_rows)} previously-unpaid invoice row(s) totalling {to_money(removed_total)} were "
            f"deleted from the ledger between the two snapshots ({names}); treated as a removal, not a "
            "payment or cancellation."
        )
    if bridge.unexplained != ZERO:
        bridge.warnings.append(
            f"Unexplained adjustment: {to_money(bridge.unexplained)} (closing debt does not fully reconcile "
            "from opening debt plus new invoices minus payments/credit notes/cancellations; see the rows above "
            "for everything that could be classified)."
        )
    return bridge


def classify_with_change_log(bridge: BridgeResult, change_log_rows: list[dict]) -> list[str]:
    """Best-effort cross-check: confirms bridge transitions against invoice_change_log actions where a
    matching entry exists; never treated as the state-at-time source itself (see load_invoice_change_log's
    docstring) -- only used to explain, and to flag low-confidence classifications."""
    logged_ids = {row.get("invoice_row_id") for row in change_log_rows or [] if row.get("invoice_row_id") is not None}
    unconfirmed = [r for r in (bridge.payments_rows + bridge.cancellations_rows) if r.get("invoice_row_id") not in logged_ids]
    if not unconfirmed:
        return []
    return [
        f"{len(unconfirmed)} payment/cancellation row(s) in this bridge have no matching invoice_change_log "
        "entry (the audit trigger may postdate the change, or it happened before the trigger was installed) "
        "-- classified by snapshot comparison only, not confirmed by the audit log."
    ]


def payments_from_bank_log(
    bank_payments: list[dict], allocations: list[dict], date_from: date, date_to: date, *, project_filter=None,
) -> tuple[Decimal, list[dict]]:
    """Payments actually received in [date_from, date_to] from the existing, already-reliable
    bank_payments/bank_payment_allocations audit log (the PRIMARY source for this tool -- not snapshot
    diffing, see get_payments_in_period's docstring)."""
    by_payment_id: dict = {}
    for alloc in allocations or []:
        by_payment_id.setdefault(alloc.get("payment_id"), []).append(alloc)
    matched_allocations = []
    total = ZERO
    for payment in bank_payments or []:
        raw_date = payment.get("payment_date")
        if not raw_date:
            continue
        try:
            pay_date = date.fromisoformat(str(raw_date)[:10])
        except Exception:
            continue
        if not (date_from <= pay_date <= date_to):
            continue
        payment_allocations = by_payment_id.get(payment.get("id"), [])
        if project_filter:
            payment_allocations = [a for a in payment_allocations if a.get("project_name") == project_filter]
            if not payment_allocations:
                continue
        if payment_allocations:
            for alloc in payment_allocations:
                total += to_money(alloc.get("amount_applied"))
                matched_allocations.append(alloc)
        else:
            total += to_money(payment.get("applied_amount") or payment.get("received_amount"))
    return total, matched_allocations
