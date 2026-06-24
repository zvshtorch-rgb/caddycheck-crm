-- ============================================================
-- Customer Purchase-Order intake + CEO approval workflow (MVP)
-- Run this in the Supabase SQL editor (once).
--
-- NOTE: The existing `orders` table is the internal camera-ordering
-- table used across the CRM. To avoid breaking it, the customer
-- purchase-order intake uses a separate table named `purchase_orders`.
-- ============================================================

-- gen_random_uuid() is provided by pgcrypto (enabled by default on Supabase).
CREATE EXTENSION IF NOT EXISTS pgcrypto;

-- ── 1. Raw incoming emails that carried a purchase order ──────────────────────
CREATE TABLE IF NOT EXISTS incoming_order_emails (
    id                  UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    message_id          TEXT        NOT NULL UNIQUE,   -- provider message id (idempotency)
    provider            TEXT,                          -- imap | graph | gmail
    from_address        TEXT,
    subject             TEXT,
    received_at         TIMESTAMPTZ,
    body_text           TEXT,
    pdf_filename        TEXT,
    pdf_storage_bucket  TEXT,
    pdf_storage_path    TEXT,
    extracted_text      TEXT,                          -- best-effort PDF text
    processing_status   TEXT        NOT NULL DEFAULT 'received'
                        CHECK (processing_status IN ('received', 'processed', 'error')),
    error_message       TEXT,
    created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS incoming_order_emails_status_idx
    ON incoming_order_emails(processing_status);

-- ── 2. Structured purchase orders awaiting / completing approval ──────────────
CREATE TABLE IF NOT EXISTS purchase_orders (
    id                  UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    incoming_email_id   UUID        REFERENCES incoming_order_emails(id) ON DELETE SET NULL,
    order_reference     TEXT,                          -- customer PO number, if detected
    customer_name       TEXT,
    customer_email      TEXT,
    project_name        TEXT,
    amount              NUMERIC,
    currency            TEXT        DEFAULT 'EUR',
    summary             TEXT,                          -- short human description
    pdf_storage_bucket  TEXT,
    pdf_storage_path    TEXT,
    status              TEXT        NOT NULL DEFAULT 'pending_approval'
                        CHECK (status IN ('pending_approval', 'approved', 'rejected', 'needs_correction', 'superseded')),
    -- Revision tracking (handles a customer re-sending a revised PDF) ----------
    revision            INTEGER     NOT NULL DEFAULT 1,
    is_current          BOOLEAN     NOT NULL DEFAULT TRUE,   -- false once superseded
    superseded_by       UUID        REFERENCES purchase_orders(id) ON DELETE SET NULL,
    decided_at          TIMESTAMPTZ,
    decision_comment    TEXT,
    created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS purchase_orders_status_idx ON purchase_orders(status);
CREATE INDEX IF NOT EXISTS purchase_orders_project_idx ON purchase_orders(project_name);
CREATE INDEX IF NOT EXISTS purchase_orders_reference_idx ON purchase_orders(order_reference);
CREATE INDEX IF NOT EXISTS purchase_orders_current_idx ON purchase_orders(is_current);

-- ── 3. Secure approval tokens (one per approval request) ──────────────────────
-- Only the SHA-256 hash of the token is stored. The raw token lives only in
-- the email link. Tokens expire after 7 days and are single-use.
CREATE TABLE IF NOT EXISTS order_approvals (
    id                  UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    purchase_order_id   UUID        NOT NULL REFERENCES purchase_orders(id) ON DELETE CASCADE,
    token_hash          TEXT        NOT NULL UNIQUE,   -- sha256 hex of the raw token
    status              TEXT        NOT NULL DEFAULT 'pending'
                        CHECK (status IN ('pending', 'approved', 'rejected', 'needs_correction', 'expired')),
    decision            TEXT,                          -- approve | reject | request_correction
    decision_comment    TEXT,
    decided_by          TEXT,
    -- Audit trail -------------------------------------------------------------
    approved_by_ip      TEXT,                          -- best-effort client IP
    user_agent          TEXT,                          -- browser user-agent
    action_timestamp    TIMESTAMPTZ,                   -- when the decision was recorded
    expires_at          TIMESTAMPTZ NOT NULL,
    used_at             TIMESTAMPTZ,
    created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS order_approvals_po_idx ON order_approvals(purchase_order_id);
CREATE INDEX IF NOT EXISTS order_approvals_status_idx ON order_approvals(status);

-- ── 4. CRM notifications (shown on the dashboard) ─────────────────────────────
CREATE TABLE IF NOT EXISTS crm_notifications (
    id                  UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    category            TEXT        NOT NULL DEFAULT 'order_approval',
    title               TEXT        NOT NULL,
    message             TEXT,
    severity            TEXT        NOT NULL DEFAULT 'info'
                        CHECK (severity IN ('info', 'success', 'warning', 'error')),
    purchase_order_id   UUID        REFERENCES purchase_orders(id) ON DELETE CASCADE,
    is_read             BOOLEAN     NOT NULL DEFAULT FALSE,
    created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS crm_notifications_unread_idx ON crm_notifications(is_read);
CREATE INDEX IF NOT EXISTS crm_notifications_created_idx ON crm_notifications(created_at DESC);

-- ── 5. Row Level Security ─────────────────────────────────────────────────────
-- All application access goes through the Supabase *service_role* key, which
-- BYPASSES RLS. We still enable RLS and add an explicit deny-all for the public
-- anon/authenticated (browser) roles so these tables (PO data and approval
-- tokens) can never be read or written directly from a browser. The approval
-- page is "public" only in the sense that it needs no CRM login; the database
-- is still reached exclusively by the trusted backend service-role key.
ALTER TABLE incoming_order_emails ENABLE ROW LEVEL SECURITY;
ALTER TABLE purchase_orders       ENABLE ROW LEVEL SECURITY;
ALTER TABLE order_approvals       ENABLE ROW LEVEL SECURITY;
ALTER TABLE crm_notifications     ENABLE ROW LEVEL SECURITY;

ALTER TABLE incoming_order_emails FORCE ROW LEVEL SECURITY;
ALTER TABLE purchase_orders       FORCE ROW LEVEL SECURITY;
ALTER TABLE order_approvals       FORCE ROW LEVEL SECURITY;
ALTER TABLE crm_notifications     FORCE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS deny_all_incoming_order_emails ON incoming_order_emails;
DROP POLICY IF EXISTS deny_all_purchase_orders       ON purchase_orders;
DROP POLICY IF EXISTS deny_all_order_approvals       ON order_approvals;
DROP POLICY IF EXISTS deny_all_crm_notifications     ON crm_notifications;

-- Explicit deny-all for the anon and authenticated (browser) roles.
CREATE POLICY deny_all_incoming_order_emails ON incoming_order_emails
    FOR ALL TO anon, authenticated USING (false) WITH CHECK (false);
CREATE POLICY deny_all_purchase_orders ON purchase_orders
    FOR ALL TO anon, authenticated USING (false) WITH CHECK (false);
CREATE POLICY deny_all_order_approvals ON order_approvals
    FOR ALL TO anon, authenticated USING (false) WITH CHECK (false);
CREATE POLICY deny_all_crm_notifications ON crm_notifications
    FOR ALL TO anon, authenticated USING (false) WITH CHECK (false);

-- ── 4. Audit trail on approval decisions ──────────────────────────────────────
ALTER TABLE order_approvals
    ADD COLUMN IF NOT EXISTS approved_by_ip   TEXT,        -- client IP at decision time
    ADD COLUMN IF NOT EXISTS user_agent       TEXT,        -- client user-agent string
    ADD COLUMN IF NOT EXISTS action_timestamp TIMESTAMPTZ; -- when the decision was recorded

-- ── 5. Revision tracking (revised PDFs / multiple PDFs for the same order) ────
ALTER TABLE purchase_orders
    ADD COLUMN IF NOT EXISTS revision_number   INTEGER NOT NULL DEFAULT 1,
    ADD COLUMN IF NOT EXISTS previous_order_id UUID REFERENCES purchase_orders(id) ON DELETE SET NULL,
    ADD COLUMN IF NOT EXISTS superseded_by_id  UUID REFERENCES purchase_orders(id) ON DELETE SET NULL;

-- Allow a 'superseded' status (older revision replaced by a newer PDF).
ALTER TABLE purchase_orders DROP CONSTRAINT IF EXISTS purchase_orders_status_check;
ALTER TABLE purchase_orders
    ADD CONSTRAINT purchase_orders_status_check
    CHECK (status IN ('pending_approval', 'approved', 'rejected', 'needs_correction', 'superseded'));

CREATE INDEX IF NOT EXISTS purchase_orders_reference_idx ON purchase_orders(order_reference);

-- ── 6. CRM notifications (in-app feed shown on the dashboard) ─────────────────
CREATE TABLE IF NOT EXISTS crm_notifications (
    id              UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    category        TEXT        NOT NULL DEFAULT 'order_approval',
    title           TEXT        NOT NULL,
    body            TEXT,
    related_table   TEXT,                   -- e.g. 'purchase_orders'
    related_id      UUID,                   -- e.g. purchase_orders.id
    is_read         BOOLEAN     NOT NULL DEFAULT FALSE,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS crm_notifications_unread_idx
    ON crm_notifications(is_read, created_at DESC);

-- ── 7. Row Level Security ────────────────────────────────────────────────────
-- The app and the poller connect with the Supabase service_role key, which
-- bypasses RLS. Enabling RLS with only service_role policies therefore denies
-- all access through the public anon key (defense in depth).
ALTER TABLE incoming_order_emails ENABLE ROW LEVEL SECURITY;
ALTER TABLE purchase_orders      ENABLE ROW LEVEL SECURITY;
ALTER TABLE order_approvals      ENABLE ROW LEVEL SECURITY;
ALTER TABLE crm_notifications    ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS service_role_all ON incoming_order_emails;
CREATE POLICY service_role_all ON incoming_order_emails
    FOR ALL TO service_role USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS service_role_all ON purchase_orders;
CREATE POLICY service_role_all ON purchase_orders
    FOR ALL TO service_role USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS service_role_all ON order_approvals;
CREATE POLICY service_role_all ON order_approvals
    FOR ALL TO service_role USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS service_role_all ON crm_notifications;
CREATE POLICY service_role_all ON crm_notifications
    FOR ALL TO service_role USING (true) WITH CHECK (true);

-- Ensure the public anon role has no direct access to these tables.
REVOKE ALL ON incoming_order_emails, purchase_orders, order_approvals, crm_notifications FROM anon;

-- Reload PostgREST schema cache so the new tables are immediately queryable.
NOTIFY pgrst, 'reload schema';
