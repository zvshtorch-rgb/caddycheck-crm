-- ============================================================
-- Run this in the Supabase SQL editor (once)
-- ============================================================

CREATE TABLE IF NOT EXISTS orders (
    id                          BIGSERIAL PRIMARY KEY,
    order_number                TEXT,
    project_name                TEXT        NOT NULL,
    country                     TEXT,
    ordered_cameras             INTEGER     NOT NULL DEFAULT 0,
    payment_month               TEXT,
    installation_year           INTEGER,
    order_date                  DATE,
    requested_activation_date   DATE,
    status                      TEXT        NOT NULL DEFAULT 'New'
                                CHECK (status IN ('New', 'Ordered', 'In Progress', 'Installed', 'Active', 'Cancelled')),
    notes                       TEXT,
    source_filename             TEXT,
    source_archive_path         TEXT,
    created_at                  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at                  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS orders_order_number_idx ON orders(order_number);
CREATE INDEX IF NOT EXISTS orders_project_name_idx ON orders(project_name);
CREATE INDEX IF NOT EXISTS orders_status_idx ON orders(status);