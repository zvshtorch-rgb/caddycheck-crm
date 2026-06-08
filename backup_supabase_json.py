"""Export Supabase tables to local JSON files.

Usage:
    py backup_supabase_json.py

The script reads .streamlit/secrets.toml, connects to Supabase via REST,
and writes timestamped JSON files under db_backups/<timestamp>/.
"""
from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

BASE_DIR = Path(__file__).resolve().parent
SECRETS_FILE = BASE_DIR / ".streamlit" / "secrets.toml"
DEFAULT_OUTPUT_DIR = BASE_DIR / "db_backups"
TABLES = [
    "projects",
    "invoices",
    "bank_payments",
    "bank_payment_allocations",
    "sent_invoices",
    "license_change_log",
    "project_change_log",
]
PAGE_SIZE = 1000


def _load_supabase_credentials(secrets_path: Path) -> tuple[str, str]:
    if not secrets_path.exists():
        raise FileNotFoundError(f"Secrets file not found: {secrets_path}")

    current_section = None
    url = ""
    service_role_key = ""

    for raw_line in secrets_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("[") and line.endswith("]"):
            current_section = line[1:-1].strip()
            continue
        if current_section != "supabase" or "=" not in line:
            continue

        key, raw_value = line.split("=", 1)
        key = key.strip()
        value = raw_value.strip().strip('"').strip()
        if key == "url":
            url = value
        elif key == "service_role_key":
            service_role_key = value

    if not url or not service_role_key:
        raise RuntimeError("Could not read [supabase] url/service_role_key from secrets.toml")

    return url, service_role_key


def _request_table_rows(base_url: str, headers: dict[str, str], table_name: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    offset = 0

    while True:
        response = requests.get(
            f"{base_url}/rest/v1/{table_name}",
            headers={**headers, "Range-Unit": "items", "Range": f"{offset}-{offset + PAGE_SIZE - 1}"},
            params={"select": "*", "order": "id"},
            timeout=60,
        )
        response.raise_for_status()
        batch = response.json()
        if not isinstance(batch, list) or not batch:
            break
        rows.extend(batch)
        if len(batch) < PAGE_SIZE:
            break
        offset += PAGE_SIZE

    return rows


def _write_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False, default=str), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Export Supabase tables to JSON backups.")
    parser.add_argument(
        "--output-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Directory where the timestamped backup folder will be created.",
    )
    args = parser.parse_args()

    url, service_role_key = _load_supabase_credentials(SECRETS_FILE)
    headers = {
        "apikey": service_role_key,
        "Authorization": f"Bearer {service_role_key}",
    }

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    output_root = Path(args.output_dir).expanduser().resolve()
    backup_dir = output_root / timestamp
    backup_dir.mkdir(parents=True, exist_ok=True)

    manifest: dict[str, Any] = {
        "created_at_utc": datetime.now(timezone.utc).isoformat(),
        "source_url": url,
        "tables": {},
    }

    for table_name in TABLES:
        rows = _request_table_rows(url, headers, table_name)
        table_file = backup_dir / f"{table_name}.json"
        _write_json(table_file, rows)
        manifest["tables"][table_name] = {
            "row_count": len(rows),
            "file": table_file.name,
        }
        print(f"Backed up {table_name}: {len(rows)} rows -> {table_file}")

    _write_json(backup_dir / "manifest.json", manifest)
    print(f"\nBackup complete: {backup_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
