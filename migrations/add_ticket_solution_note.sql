-- Run this in the Supabase SQL editor (once)
-- Adds a dedicated solution note field to tickets.

ALTER TABLE tickets
    ADD COLUMN IF NOT EXISTS solution_note TEXT;

NOTIFY pgrst, 'reload schema';