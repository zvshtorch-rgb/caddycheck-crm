-- ============================================================
-- Enable Row Level Security on all public tables.
-- The Streamlit CRM uses the SERVICE ROLE key which bypasses
-- RLS entirely, so no existing functionality is affected.
-- The anon key (used only by job_reporter.py) gets restricted
-- access to project_job_status only.
-- ============================================================

-- ── Core CRM tables (service role only – no anon access needed) ──────────────

ALTER TABLE public.projects              ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.invoices              ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.orders                ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.tickets               ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.ticket_attachments    ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.sent_invoices         ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.subscriptions         ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.bank_payments         ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.bank_payment_allocations ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.project_change_log    ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.license_change_log    ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.license_expiry_alert_log ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.purchase_order_approvals ENABLE ROW LEVEL SECURITY;

-- No policies added for the above tables → anon key gets DENIED by default.
-- Service role key bypasses RLS → Streamlit CRM works unchanged.

-- ── project_job_status (anon key used by job_reporter.py on 135 PCs) ────────

ALTER TABLE public.project_job_status    ENABLE ROW LEVEL SECURITY;

-- Drop the overly-broad "always true" policy if it exists
DROP POLICY IF EXISTS "Allow all"          ON public.project_job_status;
DROP POLICY IF EXISTS "anon_upsert"        ON public.project_job_status;
DROP POLICY IF EXISTS "Allow anon upsert"  ON public.project_job_status;
DROP POLICY IF EXISTS "anon insert"        ON public.project_job_status;
DROP POLICY IF EXISTS "anon update"        ON public.project_job_status;

-- Allow anon to INSERT new rows (first report from a new PC)
CREATE POLICY "anon_insert"
    ON public.project_job_status
    FOR INSERT
    TO anon
    WITH CHECK (true);

-- Allow anon to UPDATE (needed for upsert on conflict machine_name).
-- Agents can overwrite rows but cannot SELECT – they never see other PCs' data.
CREATE POLICY "anon_update"
    ON public.project_job_status
    FOR UPDATE
    TO anon
    USING   (true)
    WITH CHECK (true);

-- No SELECT policy for anon → agents cannot read any rows.
-- Service role SELECT is unaffected (bypasses RLS).
