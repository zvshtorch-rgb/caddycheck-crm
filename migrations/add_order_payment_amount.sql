-- ============================================================
-- Run this in the Supabase SQL editor if the orders table already exists
-- ============================================================

ALTER TABLE orders
ADD COLUMN IF NOT EXISTS payment_amount NUMERIC;