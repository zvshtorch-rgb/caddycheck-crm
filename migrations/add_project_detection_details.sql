-- Run this in the Supabase SQL editor (once)
-- Adds structured per-detection camera counts to projects.

ALTER TABLE projects
    ADD COLUMN IF NOT EXISTS backtray_cameras INTEGER NOT NULL DEFAULT 0,
    ADD COLUMN IF NOT EXISTS topdown_cameras  INTEGER NOT NULL DEFAULT 0,
    ADD COLUMN IF NOT EXISTS pushout_cameras  INTEGER NOT NULL DEFAULT 0,
    ADD COLUMN IF NOT EXISTS sco_cameras      INTEGER NOT NULL DEFAULT 0;

NOTIFY pgrst, 'reload schema';
