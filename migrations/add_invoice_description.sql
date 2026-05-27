ALTER TABLE public.invoices
ADD COLUMN IF NOT EXISTS description text;
