-- Run this in the Supabase SQL editor (once)
-- Adds a saved remarks field for Camera Audit project notes.

ALTER TABLE projects
    ADD COLUMN IF NOT EXISTS camera_audit_remarks TEXT;

NOTIFY pgrst, 'reload schema';