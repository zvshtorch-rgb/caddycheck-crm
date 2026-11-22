-- LOCAL ONLY. Legacy row-splitting ledger, NOT the staged partial-allocation backend.
-- Install only after reviewed catalog/data preflight and separate deployment approval.
-- No historical links are repaired automatically. Reviewed removal/transition is required
-- before financial-writer cutover; never drop these protections while legacy writes continue.
begin;
set local lock_timeout = '5s';
set local search_path = pg_catalog, public;

do $$
begin
    if exists(select 1 from pg_attribute where attrelid='public.bank_payment_allocations'::regclass
              and attname='writer_operation_id' and not attisdropped)
       or exists(select 1 from pg_trigger
                 where tgrelid in ('public.bank_payments'::regclass,'public.bank_payment_allocations'::regclass)
                   and tgname in ('bank_payments_writer_guard','bank_allocations_writer_guard')) then
        raise exception 'Legacy protections cannot be installed with the financial-writer backend';
    end if;
    if exists(select 1 from pg_roles where rolname='caddycheck_legacy_correction') then
        raise exception 'Correction-owner role already exists; inspect its ownership/memberships first';
    end if;
    if exists(select 1 from pg_roles where rolname='caddycheck_legacy_scope')
       or exists(select 1 from pg_roles where rolname='caddycheck_legacy_install')
       or to_regnamespace('legacy_correction_private') is not null then
        raise exception 'Private correction scope already exists; inspect its ownership/ACL first';
    end if;
    if to_regclass('public.invoice_change_log') is null
       or not exists(select 1 from pg_trigger
                     where tgrelid='public.invoices'::regclass and tgname='invoices_audit'
                       and tgenabled='O' and not tgisinternal)
       or not exists(select 1 from pg_proc
                     where oid=to_regprocedure('public.invoice_audit_trigger()') and not prosecdef)
       or pg_get_serial_sequence('public.invoice_change_log','id') is null then
        raise exception 'Enabled security-invoker invoice audit trigger/table/ID sequence required';
    end if;
    if (select count(*) from pg_class where oid in (
          'public.invoices'::regclass,'public.bank_payments'::regclass,
          'public.bank_payment_allocations'::regclass,'public.invoice_change_log'::regclass)
          and relrowsecurity) <> 4 then
        raise exception 'RLS must already be enabled on all four financial/audit tables';
    end if;
end $$;

-- Preserve only the installer's pre-existing global data-role/superuser access.
-- pg_read_all_data/pg_write_all_data imply schema USAGE despite revoked ACLs.
create temporary table legacy_install_privilege_baseline on commit drop as
select current_user::text as installer,
       rolsuper as superuser,
       pg_has_role(current_user,'pg_read_all_data','USAGE') as read_all_data,
       pg_has_role(current_user,'pg_write_all_data','USAGE') as write_all_data
from pg_roles where rolname=current_user;

-- DDL locks the relation; validation/index construction cannot race allocation inserts.
alter table public.bank_payment_allocations
    add column manual_invoice_before jsonb;
alter table public.bank_payment_allocations
    alter column invoice_row_id set not null;
alter table public.bank_payment_allocations
    add constraint legacy_allocation_invoice_fk foreign key(invoice_row_id)
        references public.invoices(id) on update no action on delete no action;
create unique index legacy_invoice_settlement_once
    on public.bank_payment_allocations(invoice_row_id);
revoke truncate on public.bank_payment_allocations, public.bank_payments, public.invoices
    from public, anon, authenticated, service_role;

-- CREATEROLE automatically receives a bootstrap-granted ADMIN membership that a
-- non-superuser cannot revoke. A disposable creator contains those memberships;
-- dropping it removes them without retaining access to the final owners.
create role caddycheck_legacy_install nologin noinherit createrole;
do $$
declare installer text := current_user;
begin
    execute format('grant caddycheck_legacy_install to %I with set true, inherit false',installer);
    execute 'set local role caddycheck_legacy_install';
    -- NOBYPASSRLS/NOSUPERUSER are defaults, not privileged attribute assignments.
    create role caddycheck_legacy_correction nologin noinherit;
    create role caddycheck_legacy_scope nologin noinherit;
    execute format('grant caddycheck_legacy_scope, caddycheck_legacy_correction to %I with set true, inherit true',
                   installer);
    execute format('set local role %I',installer);
end $$;
-- Ownership transfers require CREATE on the destination schema.
grant create on schema public to caddycheck_legacy_scope, caddycheck_legacy_correction;
grant usage on schema public to caddycheck_legacy_correction;
grant usage on schema public to caddycheck_legacy_scope;
grant execute on function public.invoice_number_key(text) to caddycheck_legacy_correction;
grant select on public.bank_payments, public.bank_payment_allocations, public.invoices
    to caddycheck_legacy_correction;
grant delete on public.bank_payment_allocations to caddycheck_legacy_correction;
-- PostgreSQL requires an UPDATE privilege for SELECT FOR UPDATE, even before DELETE.
grant update(id) on public.bank_payment_allocations to caddycheck_legacy_correction;
grant update(invoice_number,applied_amount,updated_at) on public.bank_payments to caddycheck_legacy_correction;
grant update(paid,payment_date) on public.invoices to caddycheck_legacy_correction;

create schema legacy_correction_private authorization caddycheck_legacy_scope;
revoke all on schema legacy_correction_private from public, anon, authenticated, service_role;
grant usage on schema legacy_correction_private to caddycheck_legacy_correction;
do $$
declare installer text := current_user;
begin
    execute 'set local role caddycheck_legacy_scope';
    execute format('grant usage, create on schema legacy_correction_private to %I',installer);
    grant create on schema legacy_correction_private to caddycheck_legacy_correction;
    execute format('set local role %I',installer);
end $$;
create table legacy_correction_private.active_scope (
    backend_pid integer not null,
    transaction_id bigint not null,
    allocation_id bigint not null,
    payment_id bigint not null,
    invoice_row_id bigint not null,
    actor text not null,
    reason text not null,
    invoice_before jsonb not null,
    invoice_after jsonb not null,
    payment_after jsonb not null,
    primary key(backend_pid,transaction_id)
);
alter table legacy_correction_private.active_scope owner to caddycheck_legacy_scope;
revoke all on legacy_correction_private.active_scope from public, anon, authenticated, service_role,
    caddycheck_legacy_correction;

-- Only the outer contract function can create a scope. Policies never trust caller-set GUCs.
create function legacy_correction_private.permits(
    p_kind text, p_id bigint, p_row jsonb default null
) returns boolean language sql stable security definer
set search_path = pg_catalog, legacy_correction_private as $$
    select exists(select 1 from legacy_correction_private.active_scope s
        where s.backend_pid=pg_backend_pid() and s.transaction_id=txid_current()
          and case p_kind
            when 'payment' then p_id=s.payment_id
            when 'allocations' then p_id=s.payment_id
            when 'allocation_delete' then p_id=s.allocation_id
            when 'invoice' then p_id=s.invoice_row_id
            when 'invoice_update' then p_id=s.invoice_row_id and p_row=s.invoice_after
            when 'payment_update' then p_id=s.payment_id
                and p_row-'updated_at'=s.payment_after-'updated_at'
            when 'audit_insert' then p_id=s.invoice_row_id
                and p_row->'old_values'=s.invoice_before and p_row->'new_values'=s.invoice_after
                and p_row->'invoice_number'=s.invoice_after->'invoice_number'
                and p_row->'old_invoice_number'=s.invoice_before->'invoice_number'
                and p_row->'new_invoice_number'=s.invoice_after->'invoice_number'
                and p_row->'project_name'=s.invoice_after->'project_name'
                and p_row->>'changed_by'=s.actor and p_row->>'reason'=s.reason
                and p_row->>'source'='manual-allocation-reversal'
                and p_row->>'action'='payment_status_change'
            else false end)
$$;
alter function legacy_correction_private.permits(text,bigint,jsonb) owner to caddycheck_legacy_scope;
revoke all on function legacy_correction_private.permits(text,bigint,jsonb)
    from public, anon, authenticated, service_role;
grant execute on function legacy_correction_private.permits(text,bigint,jsonb) to caddycheck_legacy_correction;

-- Do not change existing RLS enablement or browser/service-role policies.
create policy legacy_correction_payment_read on public.bank_payments
    for select to caddycheck_legacy_correction
    using(legacy_correction_private.permits('payment',id));
create policy legacy_correction_payment_update on public.bank_payments
    for update to caddycheck_legacy_correction
    using(legacy_correction_private.permits('payment',id))
    with check(legacy_correction_private.permits('payment_update',id,to_jsonb(bank_payments)));
create policy legacy_correction_allocation_read on public.bank_payment_allocations
    for select to caddycheck_legacy_correction
    using(legacy_correction_private.permits('allocations',payment_id));
-- UPDATE privilege/policy is needed for FOR UPDATE; no stored allocation change is permitted.
create policy legacy_correction_allocation_lock on public.bank_payment_allocations
    for update to caddycheck_legacy_correction
    using(legacy_correction_private.permits('allocations',payment_id)) with check(false);
create policy legacy_correction_allocation_delete on public.bank_payment_allocations
    for delete to caddycheck_legacy_correction
    using(legacy_correction_private.permits('allocation_delete',id));
create policy legacy_correction_invoice_read on public.invoices
    for select to caddycheck_legacy_correction
    using(legacy_correction_private.permits('invoice',id));
create policy legacy_correction_invoice_update on public.invoices
    for update to caddycheck_legacy_correction
    using(legacy_correction_private.permits('invoice',id))
    with check(legacy_correction_private.permits('invoice_update',id,to_jsonb(invoices)));
create policy legacy_correction_invoice_audit_insert on public.invoice_change_log
    for insert to caddycheck_legacy_correction
    with check(legacy_correction_private.permits('audit_insert',invoice_row_id,to_jsonb(invoice_change_log)));
-- Restrictive fences keep a future permissive PUBLIC policy from broadening this role's scope.
create policy legacy_correction_payment_fence on public.bank_payments as restrictive
    for all to caddycheck_legacy_correction
    using(legacy_correction_private.permits('payment',id))
    with check(legacy_correction_private.permits('payment_update',id,to_jsonb(bank_payments)));
create policy legacy_correction_invoice_fence on public.invoices as restrictive
    for all to caddycheck_legacy_correction
    using(legacy_correction_private.permits('invoice',id))
    with check(legacy_correction_private.permits('invoice_update',id,to_jsonb(invoices)));
create policy legacy_correction_allocation_fence on public.bank_payment_allocations as restrictive
    for all to caddycheck_legacy_correction
    using(legacy_correction_private.permits('allocations',payment_id)) with check(false);
create policy legacy_correction_allocation_delete_fence on public.bank_payment_allocations as restrictive
    for delete to caddycheck_legacy_correction
    using(legacy_correction_private.permits('allocation_delete',id));
create policy legacy_correction_audit_fence on public.invoice_change_log as restrictive
    for insert to caddycheck_legacy_correction
    with check(legacy_correction_private.permits('audit_insert',invoice_row_id,to_jsonb(invoice_change_log)));
grant insert(action,invoice_row_id,invoice_number,old_invoice_number,new_invoice_number,project_name,
             old_values,new_values,changed_by,source,reason)
    on public.invoice_change_log to caddycheck_legacy_correction;
do $$
begin
    -- Identity generation needs no sequence ACL. Serial-style nextval defaults do.
    if exists(select 1 from pg_attribute where attrelid='public.invoice_change_log'::regclass
              and attname='id' and attidentity='') then
        execute format('grant usage on sequence %s to caddycheck_legacy_correction',
                       pg_get_serial_sequence('public.invoice_change_log','id'));
    end if;
end $$;

create table public.legacy_allocation_reversals (
    allocation_id bigint primary key,
    actor text not null check(btrim(actor)<>''),
    reason text not null check(btrim(reason)<>''),
    reversed_at timestamptz not null default clock_timestamp(),
    allocation jsonb not null,
    payment_before jsonb not null,
    payment_after jsonb not null,
    invoice_before jsonb not null,
    invoice_after jsonb not null
);
alter table public.legacy_allocation_reversals owner to caddycheck_legacy_correction;
revoke all on public.legacy_allocation_reversals from public, anon, authenticated, service_role;
grant select on public.legacy_allocation_reversals to service_role;

create function public.guard_legacy_allocation() returns trigger
language plpgsql set search_path = pg_catalog, public as $$
declare i public.invoices%rowtype;
begin
    if tg_op in ('UPDATE','DELETE') then
        perform 1 from public.bank_payments where id=old.payment_id for update nowait;
        if current_user='caddycheck_legacy_correction' then
            if tg_op<>'DELETE' or not legacy_correction_private.permits('allocation_delete',old.id) then
                raise exception 'Correction scope does not permit changing this allocation' using errcode='42501';
            end if;
        elsif exists(select 1 from public.bank_payment_allocations
                      where payment_id=old.payment_id and manual_invoice_before is not null) then
            raise exception 'Manually reconciled payment allocations are protected; use an authorized reversal';
        end if;
        if tg_op='DELETE' then return old; end if;
    end if;
    -- Native uniqueness closes the MVCC race even at repeatable-read isolation.
    -- NOWAIT prevents payment-first legacy writers waiting behind invoice-first allocators.
    select * into strict i from public.invoices where id=new.invoice_row_id for update nowait;
    perform 1 from public.bank_payments where id=new.payment_id for update nowait;
    if exists(select 1 from public.bank_payment_allocations
              where payment_id=new.payment_id and manual_invoice_before is not null)
       and new.manual_invoice_before is null then
        raise exception 'A protected payment accepts only explicitly reviewed manual allocations';
    end if;
    if new.manual_invoice_before is not null then
        if tg_op <> 'INSERT' or new.manual_invoice_before is distinct from to_jsonb(i)
           or lower(btrim(coalesce(i.paid,''))) not in ('','no')
           or new.amount_applied is distinct from i.payment_amount or i.payment_amount <= 0 then
            raise exception 'Manual allocation requires the original unpaid full-row snapshot';
        end if;
    end if;
    return new;
end $$;
create trigger legacy_allocation_guard before insert or update or delete
    on public.bank_payment_allocations for each row execute function public.guard_legacy_allocation();

create function public.guard_legacy_manual_payment() returns trigger
language plpgsql set search_path = pg_catalog, public as $$
declare total numeric; numbers text;
begin
    if current_user='caddycheck_legacy_correction' then
        if tg_op<>'UPDATE'
           or not legacy_correction_private.permits('payment_update',new.id,to_jsonb(new)) then
            raise exception 'Correction scope does not permit this payment change' using errcode='42501';
        end if;
        return new;
    end if;
    if exists(select 1 from public.bank_payment_allocations
              where payment_id=old.id and manual_invoice_before is not null) then
        if tg_op='DELETE' then
            raise exception 'Manually reconciled payment cannot be deleted; use an authorized reversal';
        end if;
        -- Preserve the existing attach-missing-PDF workflow, never replacement of an original PDF.
        if old.pdf_storage_path is null and new.pdf_storage_path is not null
           and coalesce(btrim(new.pdf_storage_bucket),'') <> ''
           and coalesce(btrim(new.pdf_storage_path),'') <> ''
           and (to_jsonb(new)-array['pdf_storage_bucket','pdf_storage_path'])
                 is not distinct from (to_jsonb(old)-array['pdf_storage_bucket','pdf_storage_path']) then
            return new;
        end if;
        select sum(amount_applied) into total
          from public.bank_payment_allocations where payment_id=old.id;
        select string_agg(n::text,',' order by n) into numbers from (
            select distinct invoice_number n from public.bank_payment_allocations where payment_id=old.id
        ) a;
        if (to_jsonb(new)-array['invoice_number','applied_amount','updated_at'])
              is distinct from (to_jsonb(old)-array['invoice_number','applied_amount','updated_at'])
           or new.applied_amount is distinct from total or new.invoice_number::text is distinct from numbers then
            raise exception 'Manually reconciled payment metadata/summary is protected; use an authorized reversal';
        end if;
    end if;
    if tg_op='DELETE' then return old; end if;
    return new;
end $$;
create trigger legacy_manual_payment_guard before update or delete on public.bank_payments
    for each row execute function public.guard_legacy_manual_payment();

create function public.guard_legacy_manual_invoice() returns trigger
language plpgsql set search_path = pg_catalog, public as $$
begin
    if current_user='caddycheck_legacy_correction' then
        if tg_op<>'UPDATE'
           or not legacy_correction_private.permits('invoice_update',new.id,to_jsonb(new)) then
            raise exception 'Correction scope does not permit this invoice change' using errcode='42501';
        end if;
        return new;
    end if;
    if exists(select 1 from public.bank_payment_allocations
              where invoice_row_id=old.id and manual_invoice_before is not null) then
        if tg_op='DELETE' then raise exception 'Manually settled invoice cannot be deleted'; end if;
        -- The initial manual RPC may make only the reviewed paid/date transition.
        if (to_jsonb(new)-array['paid','payment_date','description','sent_at','send_status'])
              is distinct from (to_jsonb(old)-array['paid','payment_date','description','sent_at','send_status'])
           or new.paid is distinct from 'Yes'
           or new.payment_date is distinct from (
               select p.payment_date from public.bank_payments p
               join public.bank_payment_allocations a on a.payment_id=p.id where a.invoice_row_id=old.id
           ) then
            raise exception 'Manually settled invoice finance/status is protected; use an authorized reversal';
        end if;
    end if;
    if tg_op='DELETE' then return old; end if;
    return new;
end $$;
create trigger legacy_manual_invoice_guard before update or delete on public.invoices
    for each row execute function public.guard_legacy_manual_invoice();

-- Financial DML runs only as the separate, RLS-subject correction owner.
create function legacy_correction_private.reverse_scoped(
    p_allocation_id bigint, p_expected jsonb, p_actor text, p_reason text, p_confirmed boolean
) returns jsonb language plpgsql security definer
set search_path = pg_catalog, public set lock_timeout = '5s' as $$
declare
    a public.bank_payment_allocations%rowtype;
    i public.invoices%rowtype;
    p public.bank_payments%rowtype;
    invoice_before jsonb;
    payment_before jsonb;
    allocs jsonb;
begin
    if p_confirmed is distinct from true or coalesce(btrim(p_actor),'')=''
       or coalesce(btrim(p_reason),'')='' or jsonb_typeof(p_expected) is distinct from 'object' then
        raise exception 'Explicit confirmation, actor, reason and reviewed snapshot are required';
    end if;
    if exists(select 1 from pg_trigger
              where tgrelid in ('public.bank_payments'::regclass,'public.bank_payment_allocations'::regclass)
                and tgname in ('bank_payments_writer_guard','bank_allocations_writer_guard') and tgenabled<>'D') then
        raise exception 'Legacy reversal is disabled while financial-writer guards are active';
    end if;
    if (select count(*) from pg_class where oid in (
          'public.invoices'::regclass,'public.bank_payments'::regclass,
          'public.bank_payment_allocations'::regclass,'public.invoice_change_log'::regclass)
          and relrowsecurity) <> 4 then
        raise exception 'Required correction/audit RLS is not enabled';
    end if;
    select * into a from public.bank_payment_allocations where id=p_allocation_id;
    if not found then raise exception 'Allocation is missing or already reversed; reload before confirming'; end if;
    select * into strict i from public.invoices where id=a.invoice_row_id for update nowait;
    select * into strict p from public.bank_payments where id=a.payment_id for update nowait;
    perform 1 from public.bank_payment_allocations where payment_id=p.id order by id for update nowait;
    select * into strict a from public.bank_payment_allocations where id=p_allocation_id;
    select jsonb_agg(to_jsonb(x) order by id) into allocs
      from public.bank_payment_allocations x where payment_id=p.id;
    if a.manual_invoice_before is null then raise exception 'Only a manual allocation can use this reversal'; end if;
    if p_expected->'payment' is distinct from to_jsonb(p)
       or p_expected->'invoice' is distinct from to_jsonb(i)
       or p_expected->'allocations' is distinct from allocs then
        raise exception 'Correction preview is stale; reload and explicitly confirm';
    end if;
    invoice_before:=to_jsonb(i); payment_before:=to_jsonb(p);
    perform set_config('caddycheck.actor',p_actor,true);
    perform set_config('caddycheck.source','manual-allocation-reversal',true);
    perform set_config('caddycheck.reason',p_reason,true);
    delete from public.bank_payment_allocations where id=a.id;
    update public.invoices set paid=a.manual_invoice_before->>'paid',
        payment_date=(a.manual_invoice_before->>'payment_date')::date
      where id=i.id returning * into i;
    update public.bank_payments set
        applied_amount=(select coalesce(sum(amount_applied),0) from public.bank_payment_allocations where payment_id=p.id),
        invoice_number=(select string_agg(n::text,',' order by n) from (
            select distinct invoice_number n from public.bank_payment_allocations where payment_id=p.id
        ) x), updated_at=timezone('utc',now())
      where id=p.id returning * into p;
    insert into public.legacy_allocation_reversals
        (allocation_id,actor,reason,allocation,payment_before,payment_after,invoice_before,invoice_after)
    values(a.id,p_actor,p_reason,to_jsonb(a),payment_before,to_jsonb(p),invoice_before,to_jsonb(i));
    select coalesce(jsonb_agg(to_jsonb(x) order by id),'[]'::jsonb) into allocs
      from public.bank_payment_allocations x where payment_id=p.id;
    return jsonb_build_object('status','reversed','payment',to_jsonb(p),'invoice',to_jsonb(i),
                             'allocations',allocs,'allocation_id',a.id);
end $$;
alter function legacy_correction_private.reverse_scoped(bigint,jsonb,text,text,boolean)
    owner to caddycheck_legacy_correction;
revoke all on function legacy_correction_private.reverse_scoped(bigint,jsonb,text,text,boolean)
    from public, anon, authenticated, service_role;
grant execute on function legacy_correction_private.reverse_scoped(bigint,jsonb,text,text,boolean)
    to caddycheck_legacy_scope;

-- This owner has no financial-table privileges, only the private scope table and inner RPC.
-- The reviewed snapshot supplies scope IDs; the inner RPC locks/rechecks the actual rows.
create function public.reverse_manual_payment_allocation(
    p_allocation_id bigint, p_expected jsonb, p_actor text, p_reason text, p_confirmed boolean
) returns jsonb language plpgsql security definer
set search_path = pg_catalog, legacy_correction_private set lock_timeout = '5s' as $$
declare
    a jsonb;
    invoice_after jsonb;
    payment_after jsonb;
    result jsonb;
    old_actor text := current_setting('caddycheck.actor',true);
    old_source text := current_setting('caddycheck.source',true);
    old_reason text := current_setting('caddycheck.reason',true);
begin
    if p_confirmed is distinct from true or coalesce(btrim(p_actor),'')=''
       or coalesce(btrim(p_reason),'')='' or jsonb_typeof(p_expected) is distinct from 'object'
       or p_expected->>'actor' is distinct from p_actor then
        raise exception 'Explicit confirmation, matching reviewed actor, reason and snapshot are required';
    end if;
    if jsonb_typeof(p_expected->'allocations') is distinct from 'array' then
        raise exception 'Reviewed allocation rows are required';
    end if;
    select value into strict a from jsonb_array_elements(p_expected->'allocations')
      where (value->>'id')::bigint=p_allocation_id;
    if jsonb_typeof(a->'manual_invoice_before') is distinct from 'object'
       or a->>'payment_id' is distinct from p_expected->'payment'->>'id'
       or a->>'invoice_row_id' is distinct from p_expected->'invoice'->>'id' then
        raise exception 'Reviewed manual allocation/payment/invoice do not match';
    end if;
    invoice_after := p_expected->'invoice' || jsonb_build_object(
        'paid',a->'manual_invoice_before'->'paid','payment_date',a->'manual_invoice_before'->'payment_date');
    payment_after := p_expected->'payment' || jsonb_build_object(
        'applied_amount',(select coalesce(sum((value->>'amount_applied')::numeric),0)
                         from jsonb_array_elements(p_expected->'allocations')
                         where (value->>'id')::bigint<>p_allocation_id),
        'invoice_number',(select string_agg(n,',' order by n::bigint) from (
            select distinct value->>'invoice_number' n from jsonb_array_elements(p_expected->'allocations')
            where (value->>'id')::bigint<>p_allocation_id) x));
    insert into legacy_correction_private.active_scope
        (backend_pid,transaction_id,allocation_id,payment_id,invoice_row_id,actor,reason,
         invoice_before,invoice_after,payment_after)
    values(pg_backend_pid(),txid_current(),p_allocation_id,(a->>'payment_id')::bigint,
           (a->>'invoice_row_id')::bigint,p_actor,p_reason,p_expected->'invoice',invoice_after,payment_after);
    result := legacy_correction_private.reverse_scoped(
        p_allocation_id,p_expected,p_actor,p_reason,p_confirmed);
    delete from legacy_correction_private.active_scope
      where backend_pid=pg_backend_pid() and transaction_id=txid_current();
    perform set_config('caddycheck.actor',coalesce(old_actor,''),true);
    perform set_config('caddycheck.source',coalesce(old_source,''),true);
    perform set_config('caddycheck.reason',coalesce(old_reason,''),true);
    return result;
end $$;
alter function public.reverse_manual_payment_allocation(bigint,jsonb,text,text,boolean)
    owner to caddycheck_legacy_scope;
revoke all on function public.reverse_manual_payment_allocation(bigint,jsonb,text,text,boolean)
    from public, anon, authenticated;
grant execute on function public.reverse_manual_payment_allocation(bigint,jsonb,text,text,boolean) to service_role;
revoke all on function public.guard_legacy_allocation(), public.guard_legacy_manual_payment(),
    public.guard_legacy_manual_invoice() from public, anon, authenticated, service_role;
-- Remove installation grants, then drop the disposable creator and its ADMIN links.
do $$
declare
    installer text := current_user;
    baseline record;
    expected_usage boolean;
begin
    select * into strict baseline from pg_temp.legacy_install_privilege_baseline;
    if baseline.installer is distinct from installer then
        raise exception 'Installer identity changed before cleanup';
    end if;
    execute 'set local role caddycheck_legacy_scope';
    execute format('revoke usage, create on schema legacy_correction_private from %I',installer);
    revoke create on schema legacy_correction_private from caddycheck_legacy_correction;
    execute format('set local role %I',installer);
    revoke create on schema public from caddycheck_legacy_scope, caddycheck_legacy_correction;
    execute 'set local role caddycheck_legacy_install';
    execute format('revoke caddycheck_legacy_scope, caddycheck_legacy_correction from %I',installer);
    execute format('set local role %I',installer);
    drop role caddycheck_legacy_install;
    expected_usage := baseline.superuser or baseline.read_all_data or baseline.write_all_data;
    if exists(select 1 from pg_roles where rolname='caddycheck_legacy_install')
       or exists(select 1 from pg_auth_members where roleid in (
          select oid from pg_roles where rolname in ('caddycheck_legacy_scope','caddycheck_legacy_correction'))
          or member in (select oid from pg_roles where rolname in (
              'caddycheck_legacy_scope','caddycheck_legacy_correction')))
       or (select count(*) from pg_roles where rolname in (
              'caddycheck_legacy_scope','caddycheck_legacy_correction')
              and not rolcanlogin and not rolbypassrls and not rolsuper
              and not rolcreaterole and not rolcreatedb and not rolinherit) <> 2
       or (select rolsuper from pg_roles where rolname=installer) is distinct from baseline.superuser
       or pg_has_role(installer,'pg_read_all_data','USAGE') is distinct from baseline.read_all_data
       or pg_has_role(installer,'pg_write_all_data','USAGE') is distinct from baseline.write_all_data
       or has_schema_privilege(installer,'legacy_correction_private','CREATE') is distinct from baseline.superuser
       or has_schema_privilege(installer,'legacy_correction_private','USAGE') is distinct from expected_usage
       or (not baseline.superuser and (
           pg_has_role(installer,'caddycheck_legacy_scope','SET')
           or pg_has_role(installer,'caddycheck_legacy_correction','SET')))
       or has_schema_privilege('caddycheck_legacy_scope','public','CREATE')
       or has_schema_privilege('caddycheck_legacy_correction','public','CREATE')
       or has_schema_privilege('caddycheck_legacy_correction','legacy_correction_private','CREATE')
       or exists(select 1 from pg_namespace n
                 cross join lateral aclexplode(coalesce(n.nspacl,acldefault('n',n.nspowner))) acl
                 where n.nspname='legacy_correction_private'
                   and acl.grantee not in (select oid from pg_roles where rolname in (
                       'caddycheck_legacy_scope','caddycheck_legacy_correction')))
       or (select nspowner from pg_namespace where nspname='legacy_correction_private')
           is distinct from (select oid from pg_roles where rolname='caddycheck_legacy_scope') then
        raise exception 'Installation-only owner permissions were not fully removed';
    end if;
    raise notice 'Installer cleanup complete: final-owner memberships=0, temporary creator absent; pre-existing global access retained: superuser=%, pg_read_all_data=%, pg_write_all_data=%, effective private-schema USAGE=%',
        baseline.superuser,baseline.read_all_data,baseline.write_all_data,expected_usage;
end $$;
commit;
