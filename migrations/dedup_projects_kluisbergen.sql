-- ============================================================================
-- Fix duplicate project rows (e.g. "AD Kluisbergen" appearing twice).
--
-- Cause: two rows have project_name values that look identical but differ by a
-- hidden character (usually a trailing/leading space or non-breaking space).
-- Because the unique constraint is on the raw project_name, both rows coexist,
-- and the app (which trims names) cannot tell them apart to update or delete.
--
-- Run STEP 1 first to inspect. Then run STEP 2 to delete the unwanted row.
-- ============================================================================


-- ─────────────────────────────────────────────────────────────────────────
-- STEP 1: Inspect duplicates (trimmed name collides, raw name differs).
-- `ctid` is a stable per-row identifier you can delete by.
-- `name_len` reveals hidden characters: the longer one has the extra space.
-- ─────────────────────────────────────────────────────────────────────────
SELECT
    p.ctid,
    p.project_name,
    length(p.project_name)            AS name_len,
    p.num_cameras,
    p.detection_type,
    p.backtray_cameras,
    p.topdown_cameras,
    p.pushout_cameras,
    p.vim_version,
    p.status
FROM projects p
JOIN (
    SELECT btrim(project_name) AS tname
    FROM projects
    GROUP BY btrim(project_name)
    HAVING count(*) > 1
) d ON btrim(p.project_name) = d.tname
ORDER BY btrim(p.project_name), name_len;


-- ─────────────────────────────────────────────────────────────────────────
-- STEP 2: Delete the unwanted duplicate.
-- Copy the exact `ctid` value of the row you want to REMOVE from STEP 1's
-- output (it looks like (12,34)) and paste it below, then run.
--
-- Example: DELETE FROM projects WHERE ctid = '(12,34)';
-- ─────────────────────────────────────────────────────────────────────────
-- DELETE FROM projects WHERE ctid = 'PASTE_CTID_HERE';


-- ─────────────────────────────────────────────────────────────────────────
-- STEP 3 (optional): Normalize all project names so this cannot happen again.
-- Trims leading/trailing whitespace from every project name.
-- ─────────────────────────────────────────────────────────────────────────
-- UPDATE projects SET project_name = btrim(project_name)
-- WHERE project_name <> btrim(project_name);

-- NOTIFY pgrst, 'reload schema';
