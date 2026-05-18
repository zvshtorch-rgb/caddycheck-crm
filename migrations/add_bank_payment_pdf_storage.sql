-- Run this in the Supabase SQL editor (once)
-- Adds Supabase Storage references for the original uploaded bank transfer PDFs.

ALTER TABLE bank_payments
    ADD COLUMN IF NOT EXISTS pdf_storage_bucket text,
    ADD COLUMN IF NOT EXISTS pdf_storage_path   text;
