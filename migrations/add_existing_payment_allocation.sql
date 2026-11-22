-- Local implementation only. Install separately after explicit deployment approval.
-- An append-only full-row settlement, callable only by the trusted admin app backend.
begin;
set local lock_timeout = '5s';
do $$
begin
    if current_user in ('anon','authenticated','service_role','authenticator') then
        raise exception 'An independently authorized database administrator must install the release policy';
    end if;
end $$;
-- Independent release control. The migration installer owns it; backend/browser
-- roles may not activate it. No caller-set GUC or preview field authorizes release.
create table public.manual_allocation_release_policy (
    singleton boolean primary key default true check (singleton),
    enabled boolean not null default false
);
insert into public.manual_allocation_release_policy(singleton,enabled) values (true,false);
alter table public.manual_allocation_release_policy enable row level security;
revoke all on public.manual_allocation_release_policy from public, anon, authenticated, service_role;
grant select on public.manual_allocation_release_policy to service_role;

create or replace function public.add_invoice_to_existing_payment(
    p_payment_id bigint, p_invoice_row_id bigint, p_expected jsonb
) returns jsonb
language plpgsql
security invoker
set search_path = pg_catalog, public
set lock_timeout = '5s'
as $$
declare
    p public.bank_payments%rowtype;
    i public.invoices%rowtype;
    a public.bank_payment_allocations%rowtype;
    allocs jsonb;
    total numeric;
    numbers text;
    summary_numbers bigint[];
    actual_numbers bigint[];
    invoice_currency text;
    inserted_id bigint;
begin
    if current_user <> 'service_role' then
        raise exception 'Admin backend access is required';
    end if;
    if not exists(select 1 from public.manual_allocation_release_policy where singleton and enabled) then
        raise exception 'Manual invoice allocation is disabled by the database release policy';
    end if;
    if exists (
        select 1 from pg_trigger
        where tgrelid in ('public.bank_payments'::regclass, 'public.bank_payment_allocations'::regclass)
          and tgname in ('bank_payments_writer_guard', 'bank_allocations_writer_guard')
          and tgenabled <> 'D'
    ) then
        raise exception 'Adding existing-payment allocations is disabled while financial-writer guards are active';
    end if;
    if not exists(select 1 from pg_index where indexrelid=to_regclass('public.legacy_invoice_settlement_once')
                  and indisunique and indisvalid and indisready)
       or (select count(*) from pg_trigger where not tgisinternal and tgenabled='O'
           and (tgrelid,tgname) in (
               ('public.bank_payment_allocations'::regclass,'legacy_allocation_guard'),
               ('public.bank_payments'::regclass,'legacy_manual_payment_guard'),
               ('public.invoices'::regclass,'legacy_manual_invoice_guard'))) <> 3 then
        raise exception 'Coordinated legacy allocation protections must be installed and enabled first';
    end if;
    if p_payment_id <= 0 or p_invoice_row_id <= 0
       or p_payment_id is null or p_invoice_row_id is null
       or jsonb_typeof(p_expected) is distinct from 'object' then
        raise exception 'Valid payment, invoice row and reviewed snapshot are required';
    end if;

    -- Credit writes invoice -> payment. Fail fast if another writer has an opposite order.
    -- Ordinary ROW SHARE/ROW EXCLUSIVE relation locks remain compatible across unrelated rows.
    begin
        select * into i from public.invoices where id = p_invoice_row_id for update nowait;
        if not found then raise exception 'Existing invoice row no longer exists'; end if;
        select * into p from public.bank_payments where id = p_payment_id for update nowait;
        if not found then raise exception 'Saved payment no longer exists'; end if;
        perform 1 from public.bank_payment_allocations
          where payment_id = p.id order by id for update nowait;
    exception when lock_not_available then
        raise exception using errcode = '55P03',
            message = 'Payment or invoice is busy. Reload the payment and explicitly confirm a fresh preview';
    end;

    select coalesce(jsonb_agg(to_jsonb(x) order by x.id), '[]'::jsonb),
           coalesce(sum(x.amount_applied), 0)
      into allocs, total
      from public.bank_payment_allocations x where x.payment_id = p.id;
    if exists (
        select 1 from public.bank_payment_allocations x
        where x.invoice_row_id = i.id
           or (x.invoice_row_id is null and x.invoice_number::text = i.invoice_number::text
               and x.project_name = i.project_name)
    ) then
        -- A repeated click never writes again. A new preview rejects it explicitly.
        select * into a from public.bank_payment_allocations
          where payment_id = p.id and invoice_row_id = i.id order by id limit 1;
        if found and a.amount_applied = i.payment_amount and i.paid = 'Yes'
           and i.payment_date = p.payment_date
           and p_expected->'invoice'->>'id' = i.id::text
           and p_expected->'payment'->>'id' = p.id::text
           and (p_expected->'invoice'->>'payment_amount')::numeric = a.amount_applied
           and not exists (
               select 1 from public.bank_payment_allocations x
               where x.invoice_row_id = i.id and x.id <> a.id
           ) then
            return jsonb_build_object('status', 'already_allocated', 'payment', to_jsonb(p),
                                     'allocations', allocs, 'invoice', to_jsonb(i), 'allocation_id', a.id);
        end if;
        raise exception 'Invoice row is already allocated; duplicate settlement is blocked';
    end if;

    if p_expected->'payment' is distinct from to_jsonb(p)
       or p_expected->'invoice' is distinct from to_jsonb(i)
       or p_expected->'allocations' is distinct from allocs then
        raise exception 'Preview is stale. Refresh and explicitly confirm a new preview';
    end if;
    if p.payment_date is null then raise exception 'Saved payment date is required'; end if;
    if p.instructed_amount is null or p.instructed_amount::text in ('NaN', 'Infinity', '-Infinity')
       or p.instructed_amount <= 0 or p.instructed_amount >= 10000000000
       or p.instructed_amount <> round(p.instructed_amount, 2) then
        raise exception 'Positive valid instructed amount is required';
    end if;
    invoice_currency := case when to_jsonb(i) ? 'currency' then to_jsonb(i)->>'currency' else 'EUR' end;
    if p.currency is distinct from 'EUR' or invoice_currency is distinct from p.currency then
        raise exception 'Payment and invoice currencies must both be EUR and match';
    end if;
    if lower(btrim(coalesce(i.paid, ''))) not in ('', 'no') then
        raise exception 'Invoice row is not unpaid/payable';
    end if;
    if i.invoice_number is null or i.invoice_number::text !~ '^[0-9]+$' or i.invoice_number::numeric <= 0
       or i.payment_amount is null or i.payment_amount::text in ('NaN', 'Infinity', '-Infinity')
       or i.payment_amount <= 0 or i.payment_amount >= 10000000000
       or i.payment_amount <> round(i.payment_amount, 2) then
        raise exception 'Positive full existing invoice amount and number are required';
    end if;
    if exists (
        select 1 from public.bank_payment_allocations x where x.payment_id = p.id
        and (x.amount_applied is null or x.amount_applied <= 0
             or x.amount_applied::text in ('NaN', 'Infinity', '-Infinity')
             or x.amount_applied <> round(x.amount_applied, 2)
             or x.invoice_number is null or x.invoice_number <= 0)
    ) or exists (
        select 1 from public.bank_payment_allocations x where x.payment_id = p.id
          and x.invoice_row_id is not null group by x.invoice_row_id having count(*) > 1
    ) then
        raise exception 'Existing allocations are invalid or duplicated';
    end if;
    if p.applied_amount is distinct from total then
        raise exception 'Saved allocated total disagrees with actual allocations';
    end if;
    if coalesce(btrim(p.invoice_number::text), '') in ('', '0') then
        summary_numbers := '{}'::bigint[];
    else
        if btrim(p.invoice_number::text) !~ '^[0-9]+([[:space:]]*,[[:space:]]*[0-9]+)*$' then
            raise exception 'Saved invoice summary is invalid';
        end if;
        select array_agg(n order by n) into summary_numbers
          from (select btrim(value)::bigint n from unnest(string_to_array(p.invoice_number::text, ',')) value) s;
    end if;
    select coalesce(array_agg(n order by n), '{}'::bigint[]) into actual_numbers
      from (select distinct invoice_number n from public.bank_payment_allocations where payment_id = p.id) s;
    if summary_numbers is distinct from actual_numbers then
        raise exception 'Saved invoice summary disagrees with actual allocations';
    end if;
    if total + i.payment_amount > p.instructed_amount then
        raise exception 'Full invoice amount exceeds the remaining instructed balance';
    end if;

    insert into public.bank_payment_allocations
        (payment_id, invoice_row_id, invoice_number, project_name, maintenance_year, year, amount_applied,
         manual_invoice_before)
    values (p.id, i.id, i.invoice_number::bigint, i.project_name, i.maintenance_year, i.year, i.payment_amount,
            to_jsonb(i))
    returning id into inserted_id;
    update public.invoices set paid = 'Yes', payment_date = p.payment_date
      where id = i.id returning * into i;
    select sum(amount_applied) into total
      from public.bank_payment_allocations where payment_id = p.id;
    select string_agg(n::text, ',' order by n) into numbers
      from (select distinct invoice_number n from public.bank_payment_allocations where payment_id = p.id) s;
    update public.bank_payments set applied_amount = total, invoice_number = numbers,
        updated_at = timezone('utc', now())
      where id = p.id returning * into p;
    select jsonb_agg(to_jsonb(x) order by x.id) into allocs
      from public.bank_payment_allocations x where payment_id = p.id;
    return jsonb_build_object('status', 'added', 'payment', to_jsonb(p),
                             'allocations', allocs, 'invoice', to_jsonb(i), 'allocation_id', inserted_id);
end;
$$;

revoke all on function public.add_invoice_to_existing_payment(bigint, bigint, jsonb)
    from public, anon, authenticated;
grant execute on function public.add_invoice_to_existing_payment(bigint, bigint, jsonb) to service_role;
commit;
