"""One-time script to remove duplicate invoice rows from Supabase.

Deduplication rules match the application save logic:
- rows WITH an invoice number use (invoice_number, project_name)
- rows WITHOUT an invoice number use (project_name, maintenance_year, year)

The oldest row (smallest id) is kept in each duplicate group.

Usage:
  python dedup_supabase_invoices.py
  python dedup_supabase_invoices.py --apply
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any, Iterable, Optional

from supabase import create_client


SECRETS_FILE = Path(__file__).parent / ".streamlit" / "secrets.toml"


def _load_supabase_credentials(secrets_path: Path) -> tuple[str, str]:
    if not secrets_path.exists():
        raise FileNotFoundError(f"Secrets file not found: {secrets_path}")

    current_section = None
    url = ""
    service_role_key = ""

    for raw_line in secrets_path.read_text().splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and line.endswith("]"):
            current_section = line[1:-1].strip()
            continue
        if current_section != "supabase":
            continue

        if "=" not in line:
            continue

        key, raw_value = line.split("=", 1)
        key = key.strip()
        value = raw_value.strip()
        if value.startswith('"'):
            value = value[1:]
        if value.endswith('"'):
            value = value[:-1]
        value = value.strip()
        if key == "url":
            url = value
        elif key == "service_role_key":
            service_role_key = value

    if not url or not service_role_key:
        raise RuntimeError("Could not read [supabase] url/service_role_key from secrets.toml")

    return url, service_role_key


def _normalize_project_name(project_name: Any) -> str:
    return str(project_name or "").strip().lower()


def _normalize_invoice_number(invoice_number: Any) -> Optional[str]:
    if invoice_number in (None, ""):
        return None
    try:
        return str(int(float(invoice_number)))
    except (TypeError, ValueError):
        value = str(invoice_number).strip()
        return value or None


def _normalize_year(year: Any) -> Optional[int]:
    if year in (None, ""):
        return None
    try:
        return int(year)
    except (TypeError, ValueError):
        return None


def _dedupe_key(row: dict[str, Any]) -> tuple:
    project_key = _normalize_project_name(row.get("project_name"))
    invoice_number = _normalize_invoice_number(row.get("invoice_number"))
    if invoice_number is not None and project_key:
        return ("numbered", invoice_number, project_key)
    return (
        "logical",
        project_key,
        str(row.get("maintenance_year") or "").strip(),
        _normalize_year(row.get("year")),
    )


def _batched(values: list[int], size: int) -> Iterable[list[int]]:
    for start in range(0, len(values), size):
        yield values[start:start + size]


def main() -> int:
    parser = argparse.ArgumentParser(description="Remove duplicate invoice rows from Supabase.")
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually delete duplicate rows. Without this flag, the script performs a dry run.",
    )
    args = parser.parse_args()

    url, service_role_key = _load_supabase_credentials(SECRETS_FILE)
    client = create_client(url, service_role_key)

    rows = (
        client.table("invoices")
        .select("id,invoice_number,project_name,maintenance_year,year")
        .order("id")
        .execute()
        .data
    )

    print(f"Loaded {len(rows)} invoice rows from Supabase")

    kept_by_key: dict[tuple, dict[str, Any]] = {}
    duplicate_groups: dict[tuple, list[dict[str, Any]]] = {}

    for row in rows:
        key = _dedupe_key(row)
        if key not in kept_by_key:
            kept_by_key[key] = row
            continue
        duplicate_groups.setdefault(key, [kept_by_key[key]]).append(row)

    duplicate_ids: list[int] = []
    for group in duplicate_groups.values():
        duplicate_ids.extend(int(row["id"]) for row in group[1:])

    print(f"Duplicate groups: {len(duplicate_groups)}")
    print(f"Rows to delete: {len(duplicate_ids)}")

    if duplicate_groups:
        print("Sample duplicate groups:")
        for key, group in list(duplicate_groups.items())[:10]:
            ids = [row["id"] for row in group]
            print(f"  keep id={ids[0]}, delete ids={ids[1:]}, key={key}")

    if not duplicate_ids:
        print("No duplicates found.")
        return 0

    if not args.apply:
        print("Dry run only. Re-run with --apply to delete duplicates.")
        return 0

    deleted = 0
    for batch in _batched(duplicate_ids, 200):
        client.table("invoices").delete().in_("id", batch).execute()
        deleted += len(batch)

    print(f"Deleted {deleted} duplicate invoice rows.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())