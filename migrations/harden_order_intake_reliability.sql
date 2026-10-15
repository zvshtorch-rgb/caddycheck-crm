-- ============================================================
-- Order intake reliability hardening
-- Run ONCE in the Supabase SQL editor BEFORE enabling live polling.
-- Safe to re-run (idempotent). Adds columns, unique indexes and a
-- schema-version marker that the poller checks before it writes anything.
--
-- Pre-flight (optional): these must both return zero rows, otherwise the
-- unique indexes below cannot be created and this script aborts with a
-- message (nothing is changed):
--
--   SELECT incoming_email_id, COUNT(*) FROM purchase_orders
--    WHERE incoming_email_id IS NOT NULL GROUP BY 1 HAVING COUNT(*) > 1;
--   SELECT purchase_order_id, COUNT(*) FROM order_approvals
--    WHERE status = 'pending' GROUP BY 1 HAVING COUNT(*) > 1;
-- ============================================================

BEGIN;

DO $$
BEGIN
    IF EXISTS (
        SELECT 1 FROM purchase_orders
         WHERE incoming_email_id IS NOT NULL
         GROUP BY incoming_email_id HAVING COUNT(*) > 1
    ) THEN
        RAISE EXCEPTION 'Duplicate purchase_orders.incoming_email_id values exist; resolve them first (see pre-flight query at top of this file).';
    END IF;
    IF EXISTS (
        SELECT 1 FROM order_approvals
         WHERE status = 'pending'
         GROUP BY purchase_order_id HAVING COUNT(*) > 1
    ) THEN
        RAISE EXCEPTION 'Multiple pending order_approvals rows exist for one purchase order; resolve them first (see pre-flight query at top of this file).';
    END IF;
END $$;

-- ── incoming_order_emails ────────────────────────────────────────────────────
ALTER TABLE incoming_order_emails
    ADD COLUMN IF NOT EXISTS graph_message_id     TEXT,        -- immutable Graph message id
    ADD COLUMN IF NOT EXISTS pdf_sha256           TEXT,        -- flags possible duplicate PDFs only
    ADD COLUMN IF NOT EXISTS duplicate_of_email_id UUID REFERENCES incoming_order_emails(id) ON DELETE SET NULL,
    ADD COLUMN IF NOT EXISTS claimed_by           TEXT,        -- processing lease holder (run id)
    ADD COLUMN IF NOT EXISTS claimed_at           TIMESTAMPTZ, -- lease start
    ADD COLUMN IF NOT EXISTS attempt_count        INTEGER NOT NULL DEFAULT 0;

-- message_id is already UNIQUE; the immutable Graph id gets its own guarantee.
CREATE UNIQUE INDEX IF NOT EXISTS incoming_order_emails_graph_message_id_uniq
    ON incoming_order_emails(graph_message_id) WHERE graph_message_id IS NOT NULL;
CREATE INDEX IF NOT EXISTS incoming_order_emails_pdf_sha256_idx
    ON incoming_order_emails(pdf_sha256) WHERE pdf_sha256 IS NOT NULL;

-- ── purchase_orders: at most one order per intake email ──────────────────────
CREATE UNIQUE INDEX IF NOT EXISTS purchase_orders_incoming_email_uniq
    ON purchase_orders(incoming_email_id) WHERE incoming_email_id IS NOT NULL;

-- ── order_approvals: durable approval-email delivery state ───────────────────
-- Existing rows were created (and emailed) by the old poller, so they are
-- backfilled as 'sent'. The default is switched to 'pending' for new rows.
ALTER TABLE order_approvals
    ADD COLUMN IF NOT EXISTS email_status     TEXT NOT NULL DEFAULT 'sent'
        CHECK (email_status IN ('pending', 'sending', 'sent', 'failed')),
    ADD COLUMN IF NOT EXISTS email_attempts   INTEGER NOT NULL DEFAULT 0,
    ADD COLUMN IF NOT EXISTS email_sent_at    TIMESTAMPTZ,
    ADD COLUMN IF NOT EXISTS email_claimed_at TIMESTAMPTZ,
    ADD COLUMN IF NOT EXISTS email_last_error TEXT;

ALTER TABLE order_approvals ALTER COLUMN email_status SET DEFAULT 'pending';

UPDATE order_approvals
   SET email_sent_at = created_at
 WHERE email_status = 'sent' AND email_sent_at IS NULL;

-- One live token per purchase order; a resend rotates the token in place.
CREATE UNIQUE INDEX IF NOT EXISTS order_approvals_one_pending_per_po
    ON order_approvals(purchase_order_id) WHERE status = 'pending';

-- ── Schema-version marker (created last, so its presence proves the above) ───
CREATE OR REPLACE FUNCTION order_intake_schema_version()
RETURNS INTEGER
LANGUAGE sql
IMMUTABLE
AS $$ SELECT 20261015 $$;

REVOKE ALL ON FUNCTION order_intake_schema_version() FROM PUBLIC, anon, authenticated;
GRANT EXECUTE ON FUNCTION order_intake_schema_version() TO service_role;

NOTIFY pgrst, 'reload schema';

COMMIT;
