-- Run this in the Supabase SQL editor (once)

alter table if exists tickets
    add column if not exists subcategory text;

create index if not exists tickets_subcategory_idx on tickets(subcategory);