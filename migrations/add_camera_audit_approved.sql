-- Run this in the Supabase SQL editor (once)
-- Adds an approval flag for accepted Camera Audit discrepancies.

ALTER TABLE projects
    ADD COLUMN IF NOT EXISTS camera_audit_approved BOOLEAN NOT NULL DEFAULT FALSE;

NOTIFY pgrst, 'reload schema';