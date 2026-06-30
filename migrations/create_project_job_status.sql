-- Per-PC job/camera usage reported by the Video Profiler agent (job_reporter.py).
-- One row per machine; updated on every report (upsert by machine_name).
--
-- The CRM joins this table against purchase_orders / projects to flag any
-- project whose active job (camera) count exceeds the approved quantity.

CREATE TABLE IF NOT EXISTS project_job_status (
    id              BIGSERIAL PRIMARY KEY,
    machine_name    TEXT NOT NULL UNIQUE,        -- stable PC hostname
    project_name    TEXT,                        -- mapped CRM project (nullable until mapped)
    owner           TEXT,                        -- "Owner" column from dbo.Jobs (e.g. CADDYCHECK)
    active_jobs     INTEGER NOT NULL DEFAULT 0,  -- jobs not yet completed (≈ active cameras)
    total_jobs      INTEGER NOT NULL DEFAULT 0,  -- all jobs in dbo.Jobs
    app_version     TEXT,
    agent_version   TEXT,
    reported_at     TIMESTAMPTZ NOT NULL DEFAULT now(),
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_project_job_status_project
    ON project_job_status (project_name);

CREATE INDEX IF NOT EXISTS idx_project_job_status_reported
    ON project_job_status (reported_at);
