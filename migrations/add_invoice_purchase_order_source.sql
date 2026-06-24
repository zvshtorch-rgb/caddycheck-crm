-- Run this in Supabase SQL editor (once).
-- Links auto-created invoices back to the approved purchase order they came from
-- and marks them as drafts ("ready_to_send") for review before sending.

ALTER TABLE invoices
    ADD COLUMN IF NOT EXISTS source_type text,
    ADD COLUMN IF NOT EXISTS source_purchase_order_id uuid REFERENCES purchase_orders(id) ON DELETE SET NULL,
    ADD COLUMN IF NOT EXISTS auto_created boolean DEFAULT false,
    ADD COLUMN IF NOT EXISTS auto_created_at timestamptz,
    ADD COLUMN IF NOT EXISTS send_status text;

-- Enforce at most one invoice per purchase order (idempotency safety net).
CREATE UNIQUE INDEX IF NOT EXISTS invoices_source_purchase_order_id_key
    ON invoices (source_purchase_order_id)
    WHERE source_purchase_order_id IS NOT NULL;

NOTIFY pgrst, 'reload schema';
