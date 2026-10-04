-- Apply Credit as ONE database transaction.
-- Run this once in the Supabase SQL editor (Dashboard -> SQL Editor).
--
-- The app builds the complete list of changes for an "Apply Credit" click (credit source rows, usage rows,
-- target invoice rows, remainder rows, payment record + allocations) and sends it here. Every statement runs
-- inside this single function call, i.e. one transaction: any failed precondition or error raises an
-- exception and Postgres rolls EVERYTHING back. Credit/target rows are only updated while they still have
-- the amount they had when the app loaded them and are still open (status No/blank), so stale data can
-- never spend the same credit twice.
--
-- Until this function exists the app falls back to a step-by-step apply with compensating rollback.

create or replace function public.apply_credit_atomic(p_plan jsonb)
returns jsonb
language plpgsql
set search_path = public
as $$
declare
    v_op         jsonb;
    v_kind       text;
    v_fields     jsonb;
    v_entry      jsonb;
    v_alloc      jsonb;
    v_n          integer;
    v_payment_id bigint;
    v_result_id  bigint := null;
begin
    if p_plan is null or jsonb_typeof(p_plan -> 'ops') is distinct from 'array' then
        raise exception 'apply_credit_atomic: plan must contain an "ops" array';
    end if;

    for v_op in select value from jsonb_array_elements(p_plan -> 'ops')
    loop
        v_kind := v_op ->> 'op';

        if v_kind = 'credit_consume' then
            update invoices set
                payment_amount = case when (v_op ->> 'mark_paid')::boolean then payment_amount
                                      else (v_op ->> 'new_amount')::numeric end,
                paid           = case when (v_op ->> 'mark_paid')::boolean then 'Yes' else paid end,
                payment_date   = case when (v_op ->> 'mark_paid')::boolean then (v_op ->> 'payment_date')::date
                                      else payment_date end,
                description    = v_op ->> 'description'
            where id = (v_op ->> 'id')::integer
              and round(payment_amount, 2) = round((v_op ->> 'expected_amount')::numeric, 2)
              and lower(btrim(coalesce(paid, ''))) in ('', 'no');
            get diagnostics v_n = row_count;
            if v_n <> 1 then
                raise exception 'apply_credit_atomic: credit row % changed since it was loaded or is no longer open', v_op ->> 'id';
            end if;

        elsif v_kind = 'target_update' then
            v_fields := v_op -> 'fields';
            update invoices set
                payment_amount = case when v_fields ? 'payment_amount' then (v_fields ->> 'payment_amount')::numeric
                                      else payment_amount end,
                paid           = case when v_fields ? 'paid' then v_fields ->> 'paid' else paid end,
                payment_date   = case when v_fields ? 'payment_date' then (v_fields ->> 'payment_date')::date
                                      else payment_date end,
                description    = case when v_fields ? 'description' then v_fields ->> 'description'
                                      else description end
            where id = (v_op ->> 'id')::integer
              and round(payment_amount, 2) = round((v_op ->> 'expected_amount')::numeric, 2)
              and lower(btrim(coalesce(paid, ''))) in ('', 'no');
            get diagnostics v_n = row_count;
            if v_n <> 1 then
                raise exception 'apply_credit_atomic: invoice row % changed since it was loaded or is no longer open', v_op ->> 'id';
            end if;

        elsif v_kind = 'insert_row' then
            insert into invoices (invoice_number, project_name, maintenance_year, payment_amount,
                                  payment_date, paid, year, invoice_type, description)
            values (v_op -> 'row' ->> 'invoice_number',
                    v_op -> 'row' ->> 'project_name',
                    v_op -> 'row' ->> 'maintenance_year',
                    (v_op -> 'row' ->> 'payment_amount')::numeric,
                    nullif(v_op -> 'row' ->> 'payment_date', '')::date,
                    coalesce(v_op -> 'row' ->> 'paid', 'No'),
                    nullif(v_op -> 'row' ->> 'year', '')::integer,
                    v_op -> 'row' ->> 'invoice_type',
                    v_op -> 'row' ->> 'description');

        elsif v_kind = 'save_payment' then
            v_entry := v_op -> 'entry';
            select id into v_payment_id from bank_payments
             where payment_fingerprint = v_entry ->> 'payment_fingerprint';
            if v_payment_id is null then
                insert into bank_payments (payment_date, invoice_number, source_name, source_kind, payment_fingerprint,
                                           instructed_amount, received_amount, applied_amount, fee_amount, currency,
                                           raw_text, parsed_payload, notes)
                values (nullif(v_entry ->> 'payment_date', '')::date, v_entry ->> 'invoice_number',
                        v_entry ->> 'source_name', v_entry ->> 'source_kind', v_entry ->> 'payment_fingerprint',
                        (v_entry ->> 'instructed_amount')::numeric, (v_entry ->> 'received_amount')::numeric,
                        (v_entry ->> 'applied_amount')::numeric, (v_entry ->> 'fee_amount')::numeric,
                        coalesce(v_entry ->> 'currency', 'EUR'), v_entry ->> 'raw_text',
                        coalesce(v_entry -> 'parsed_payload', '{}'::jsonb), v_entry ->> 'notes')
                returning id into v_payment_id;
            else
                update bank_payments set
                    payment_date   = nullif(v_entry ->> 'payment_date', '')::date,
                    invoice_number = v_entry ->> 'invoice_number',
                    applied_amount = (v_entry ->> 'applied_amount')::numeric,
                    parsed_payload = coalesce(v_entry -> 'parsed_payload', '{}'::jsonb),
                    notes          = v_entry ->> 'notes',
                    updated_at     = timezone('utc', now())
                where id = v_payment_id;
            end if;

            delete from bank_payment_allocations where payment_id = v_payment_id;
            for v_alloc in select value from jsonb_array_elements(coalesce(v_op -> 'allocations', '[]'::jsonb))
            loop
                insert into bank_payment_allocations (payment_id, invoice_row_id, invoice_number, project_name,
                                                      maintenance_year, year, amount_applied)
                values (v_payment_id,
                        (v_alloc ->> 'invoice_row_id')::bigint,
                        (v_alloc ->> 'invoice_number')::bigint,
                        v_alloc ->> 'project_name',
                        v_alloc ->> 'maintenance_year',
                        nullif(v_alloc ->> 'year', '')::integer,
                        (v_alloc ->> 'amount_applied')::numeric);
            end loop;
            v_result_id := v_payment_id;

        else
            raise exception 'apply_credit_atomic: unknown operation %', v_kind;
        end if;
    end loop;

    return jsonb_build_object('ok', true, 'payment_id', v_result_id);
end;
$$;

-- Only the backend (service role) may call it; it must not be reachable with the public anon key.
revoke all on function public.apply_credit_atomic(jsonb) from public;
revoke all on function public.apply_credit_atomic(jsonb) from anon, authenticated;
grant execute on function public.apply_credit_atomic(jsonb) to service_role;
