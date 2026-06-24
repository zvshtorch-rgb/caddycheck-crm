"""Clear all invoice descriptions – set them to empty/None."""
import sys
from pathlib import Path

# Add parent directory to path so we can import modules
sys.path.insert(0, str(Path(__file__).parent))

from services.supabase_service import load_invoices, save_invoices


def clear_invoice_descriptions():
    """Load all invoices and clear their descriptions."""
    print("Loading invoices...")
    invoices = load_invoices()
    print(f"Loaded {len(invoices)} invoices")

    # Count invoices with descriptions
    invoices_with_desc = [inv for inv in invoices if inv.description]
    print(f"Found {len(invoices_with_desc)} invoices with descriptions")

    if not invoices_with_desc:
        print("No invoices with descriptions to clear.")
        return

    # Clear descriptions
    for inv in invoices:
        if inv.description:
            print(f"  Clearing: Invoice #{inv.invoice_number} - {inv.project_name}")
            inv.description = None

    # Save back
    print("\nSaving cleared invoices...")
    save_invoices(invoices)
    print(f"✓ Successfully cleared {len(invoices_with_desc)} invoice descriptions!")


if __name__ == "__main__":
    clear_invoice_descriptions()
