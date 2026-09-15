-- The bank_payments.invoice_number column was bigint, but multi-invoice
-- payments (the common case for batch SWIFT transfers) need to store a
-- comma-separated list of invoice numbers. Every such payment silently
-- failed to insert and fell back to ephemeral local storage, which is why
-- recent multi-invoice payments were missing from the "Saved Bank Payments"
-- table. Found + fixed during the 2026-09-15 PPS reconciliation audit.
alter table bank_payments
    alter column invoice_number type text using invoice_number::text;
