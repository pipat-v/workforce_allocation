do $$
declare
  constraint_name text;
begin
  select con.conname
    into constraint_name
  from pg_constraint con
  join pg_class rel on rel.oid = con.conrelid
  join pg_namespace nsp on nsp.oid = rel.relnamespace
  where nsp.nspname = 'public'
    and rel.relname = 'holidays'
    and con.contype = 'c'
    and pg_get_constraintdef(con.oid) like '%buddhist_holy_day%'
    and pg_get_constraintdef(con.oid) like '%public_holiday%'
    and pg_get_constraintdef(con.oid) like '%company_holiday%'
  limit 1;

  if constraint_name is not null then
    execute format('alter table public.holidays drop constraint %I', constraint_name);
  end if;
end $$;

alter table public.holidays
  add constraint holidays_type_check
  check (type in ('buddhist_holy_day', 'public_holiday', 'national_holiday', 'company_holiday'));
