-- Run this in Supabase SQL editor (once)
-- Adds invoice metadata columns used by Invoice Details and Debt Report sync.

ALTER TABLE invoices
ADD COLUMN IF NOT EXISTS invoice_type text,
ADD COLUMN IF NOT EXISTS for_month text,
ADD COLUMN IF NOT EXISTS sent_at text;
