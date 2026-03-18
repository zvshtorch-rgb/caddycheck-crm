-- ============================================================
-- Run this in the Supabase SQL editor (once, in order)
-- ============================================================

-- 1. subscriptions — one row per project, current license state
CREATE TABLE IF NOT EXISTS subscriptions (
    id              BIGSERIAL PRIMARY KEY,
    project_name    TEXT        NOT NULL UNIQUE REFERENCES projects(project_name) ON UPDATE CASCADE,
    valid_from      DATE,
    valid_until     DATE,
    cameras_allowed INTEGER,
    module_name     TEXT        NOT NULL DEFAULT 'Video Inform Profiler',
    status          TEXT        NOT NULL DEFAULT 'active'
                    CHECK (status IN ('active', 'expired', 'suspended')),
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 2. renewal_links — one-time-use tokens generated after each payment
CREATE TABLE IF NOT EXISTS renewal_links (
    id                 BIGSERIAL PRIMARY KEY,
    project_name       TEXT        NOT NULL,
    subscription_id    BIGINT      REFERENCES subscriptions(id),
    token              TEXT        NOT NULL UNIQUE,
    created_at         TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    expires_at         TIMESTAMPTZ NOT NULL,
    used_at            TIMESTAMPTZ,
    target_valid_until DATE        NOT NULL,
    cameras_allowed    INTEGER,
    status             TEXT        NOT NULL DEFAULT 'pending'
                       CHECK (status IN ('pending', 'used', 'expired')),
    invoice_number     TEXT,
    payment_amount     NUMERIC
);

CREATE INDEX IF NOT EXISTS renewal_links_token_idx    ON renewal_links(token);
CREATE INDEX IF NOT EXISTS renewal_links_project_idx  ON renewal_links(project_name);
CREATE INDEX IF NOT EXISTS subscriptions_project_idx  ON subscriptions(project_name);
