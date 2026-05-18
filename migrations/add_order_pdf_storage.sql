-- Run this in the Supabase SQL editor (once)
-- Adds Supabase Storage references for the original uploaded order PDFs.

ALTER TABLE orders
    ADD COLUMN IF NOT EXISTS pdf_storage_bucket text,
    ADD COLUMN IF NOT EXISTS pdf_storage_path   text;
