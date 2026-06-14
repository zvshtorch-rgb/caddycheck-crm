-- Run this in the Supabase SQL editor (once)
-- Stores metadata for image and video attachments uploaded to support tickets.

CREATE TABLE IF NOT EXISTS ticket_attachments (
    id              BIGSERIAL PRIMARY KEY,
    ticket_id       BIGINT      NOT NULL REFERENCES tickets(id) ON DELETE CASCADE,
    file_name       TEXT        NOT NULL,
    file_type       TEXT,
    file_size       BIGINT,
    storage_bucket  TEXT        NOT NULL DEFAULT 'ticket-attachments',
    storage_path    TEXT        NOT NULL UNIQUE,
    uploaded_by     TEXT,
    notes           TEXT,
    uploaded_at     TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS ticket_attachments_ticket_id_idx ON ticket_attachments(ticket_id);
CREATE INDEX IF NOT EXISTS ticket_attachments_uploaded_at_idx ON ticket_attachments(uploaded_at);
