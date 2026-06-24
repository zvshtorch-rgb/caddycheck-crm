-- Clear all invoice descriptions (set to NULL)
-- Run this in the Supabase SQL editor to remove all descriptions from invoices

UPDATE invoices
SET description = NULL
WHERE description IS NOT NULL;
