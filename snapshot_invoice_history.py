"""Ask Data Phase 5: daily invoice history snapshot writer.

Captures today's (or an explicitly given day's) full invoices-table state into `invoice_snapshots`, one row
per invoice, so Ask Data's historical/trend tools can reconstruct "debt at date X" without ever touching the
live `invoices`/`projects` tables. This script is intentionally SEPARATE from Ask Data: Ask Data only ever
reads `invoice_snapshots`/`invoice_snapshot_runs`; only this script (run by the scheduled backup workflow)
ever writes to them.

Guarantees:
  * idempotent -- rerunning for the SAME day fully overwrites that day's rows (upsert keyed on
    (snapshot_date, invoice_row_id)) rather than duplicating them; `captured_at` always reflects the latest
    run.
  * never mutates `invoices`/`projects`/`bank_payments`/`invoice_change_log` -- read-only against all of
    those.
  * a failure here can NEVER affect the CRM: the only tables ever written to are the two new, purely
    additive history tables.
  * auditable -- a run-log row (`invoice_snapshot_runs`) is written BEFORE the data rows (status=
    'running'), so a crash mid-run leaves visible evidence instead of a silent gap, and is updated to
    'success'/'failed' at the end.

Usage:
    python snapshot_invoice_history.py                 # capture today (UTC), write to Supabase
    python snapshot_invoice_history.py --date 2026-10-20
    python snapshot_invoice_history.py --dry-run        # compute and print row count, write nothing

NOT yet scheduled/wired into .github/workflows/supabase-backup.yml -- see the Phase 5 rollout plan.
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))


def build_snapshot_rows(invoices, projects, captured_at: str) -> list[dict]:
    """Pure: one dict per invoice row, matching the invoice_snapshots schema exactly. No I/O.

    `captured_at` is a single UTC ISO timestamp computed ONCE by the caller for the whole run and stamped
    on every row explicitly -- never left to the table's DB-level default. The default only fires on a
    fresh INSERT; on the UPDATE branch of an upsert (a same-day rerun hitting the (snapshot_date,
    invoice_row_id) unique constraint), a column absent from the payload is left untouched, so relying on
    the DB default would silently keep the OLD captured_at forever after the first run. Passing it
    explicitly here, identical for every row, is what makes a rerun's captured_at actually refresh."""
    from services.ask_data.resolver import NetworkIndex, ProjectIndex

    project_index = ProjectIndex(projects)
    network_index = NetworkIndex(projects)
    rows = []
    for inv in invoices:
        if inv.db_id is None:
            continue  # never-saved/placeholder rows have no stable identity to snapshot
        canonical = project_index.canonical_name(inv.project_name)
        country = project_index.country(inv.project_name)
        network = network_index.network_of(canonical) if canonical else None
        rows.append({
            "invoice_row_id": inv.db_id,
            "invoice_number": int(inv.invoice_number) if inv.invoice_number else None,
            "project_name": canonical or inv.project_name,
            "maintenance_year": inv.maintenance_year,
            "year": inv.year,
            "for_month": inv.for_month or None,
            "invoice_type": inv.invoice_type or None,
            "payment_amount": float(inv.payment_amount or 0),
            "paid": str(inv.paid or "No"),
            "payment_date": inv.payment_date.date().isoformat() if inv.payment_date else None,
            "country_at_capture": country,
            "network_at_capture": network,
            "captured_at": captured_at,
        })
    return rows


def run(snapshot_date: datetime.date, *, dry_run: bool = False) -> int:
    from services.supabase_service import (
        load_invoices, load_projects, upsert_invoice_snapshot_rows, upsert_invoice_snapshot_run,
    )

    started_at = datetime.datetime.now(datetime.timezone.utc).isoformat()
    # One timestamp for the ENTIRE run, shared by every row AND the run-log -- see build_snapshot_rows()'s
    # docstring for why this must be explicit rather than relying on the table's DB-level default.
    captured_at = started_at
    print(f"Invoice snapshot for {snapshot_date.isoformat()} starting at {started_at}...")

    if not dry_run:
        upsert_invoice_snapshot_run(snapshot_date, status="running", started_at=started_at)

    try:
        invoices = load_invoices()
        projects = load_projects()
        rows = build_snapshot_rows(invoices, projects, captured_at)
    except Exception as exc:
        print(f"FAILED to build snapshot rows: {exc}")
        if not dry_run:
            upsert_invoice_snapshot_run(
                snapshot_date, status="failed", error_message=str(exc),
                completed_at=datetime.datetime.now(datetime.timezone.utc).isoformat(),
            )
        return 1

    print(f"Built {len(rows)} snapshot row(s) from {len(invoices)} invoice(s) / {len(projects)} project(s).")
    if dry_run:
        print("Dry run only -- nothing written.")
        return 0

    try:
        written = upsert_invoice_snapshot_rows(snapshot_date, rows)
        upsert_invoice_snapshot_run(
            snapshot_date, status="success", row_count=written,
            completed_at=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        )
    except Exception as exc:
        print(f"FAILED to write snapshot rows: {exc}")
        upsert_invoice_snapshot_run(
            snapshot_date, status="failed", error_message=str(exc),
            completed_at=datetime.datetime.now(datetime.timezone.utc).isoformat(),
        )
        return 1

    print(f"Wrote {written} row(s) for {snapshot_date.isoformat()}.")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--date", type=str, default=None, help="snapshot date YYYY-MM-DD (default: today, UTC)")
    parser.add_argument("--dry-run", action="store_true", help="compute and print row count; write nothing")
    args = parser.parse_args()

    snapshot_date = (
        datetime.date.fromisoformat(args.date) if args.date
        else datetime.datetime.now(datetime.timezone.utc).date()
    )
    try:
        return run(snapshot_date, dry_run=args.dry_run)
    except Exception as exc:  # never let a failure here raise past the script -- see module docstring
        print(f"Invoice snapshot job failed: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
