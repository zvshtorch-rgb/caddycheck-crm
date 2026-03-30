#!/usr/bin/env python3
"""
One-time migration: import Excel data into Supabase.
Run from project root:  python migrate_to_supabase.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

SUPABASE_URL = "https://rdoxihpmghrvroddnkdi.supabase.co"
SUPABASE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
    ".eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJkb3hpaHBtZ2hydnJvZGRua2RpIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3MzY0NTExNywiZXhwIjoyMDg5MjIxMTE3fQ"
    ".umgghE4z-ClVQ0KY8LQJhJtbG2tYlVh0fY0d9JnYXBA"
)


def _get_client():
    from supabase import create_client
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def migrate():
    print("Loading data from Excel...")
    from services.excel_service import load_projects, load_invoices
    from config.settings import get_data_paths

    paths = get_data_paths()
    projects = load_projects(paths["projects_file"])
    invoices = load_invoices(paths["projects_file"])
    print(f"  Found {len(projects)} projects, {len(invoices)} invoices")

    client = _get_client()
    batch_size = 50

    # ── Migrate projects ───────────────────────────────────────────────────────
    print("Migrating projects...")
    project_rows = []
    for p in projects:
        row = {
            "project_name": p.project_name,
            "country": p.country or None,
            "num_cameras": p.num_cams or None,
            "payment_month": p.payment_month or None,
            "installation_year": p.installation_year,
            "project_approval": p.project_approval or None,
            "activation_date": p.activation_date.date().isoformat() if p.activation_date else None,
            "detection_type": p.detection_type or None,
            "cart_type": p.cart_type or None,
            "vim_version": p.vim_version or None,
            "status": p.status or None,
            "license_eop": p.license_eop.date().isoformat() if p.license_eop else None,
            "caddy_back": p.caddy_back or None,
        }
        for i in range(1, 10):
            row[f"m{i}y"] = (
                str(int(p.maintenance_invoice_numbers[i]))
                if i in p.maintenance_invoice_numbers else None
            )
        project_rows.append(row)

    for i in range(0, len(project_rows), batch_size):
        batch = project_rows[i:i + batch_size]
        client.table("projects").upsert(batch, on_conflict="project_name").execute()
        print(f"  Projects: {min(i + batch_size, len(project_rows))}/{len(project_rows)}")

    # ── Migrate invoices ───────────────────────────────────────────────────────
    print("Migrating invoices...")
    invoice_rows = []
    for inv in invoices:
        invoice_rows.append({
            "invoice_number": str(int(inv.invoice_number)) if inv.invoice_number else None,
            "project_name": inv.project_name,
            "maintenance_year": inv.maintenance_year or None,
            "payment_amount": inv.payment_amount,
            "cameras_number": int(inv.cameras_number) if inv.cameras_number else None,
            "payment_date": inv.payment_date.date().isoformat() if inv.payment_date else None,
            "paid": inv.paid or "No",
            "year": inv.year,
        })

    for i in range(0, len(invoice_rows), batch_size):
        batch = invoice_rows[i:i + batch_size]
        client.table("invoices").insert(batch).execute()
        print(f"  Invoices: {min(i + batch_size, len(invoice_rows))}/{len(invoice_rows)}")

    print("Migration complete!")


if __name__ == "__main__":
    migrate()
