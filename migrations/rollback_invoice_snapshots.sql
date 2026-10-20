-- Rollback for migrations/create_invoice_snapshots.sql. NOT yet run.
-- Emergency use only -- this permanently deletes all captured historical snapshot rows and the job's
-- run-log. It does NOT touch `invoices`, `projects`, `bank_payments`, or `invoice_change_log`: those
-- operational/audit tables are never written to or read destructively by Ask Data Phase 5.

drop table if exists invoice_snapshot_runs;
drop table if exists invoice_snapshots;
