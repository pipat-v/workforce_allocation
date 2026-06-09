create table if not exists public.employee_warnings (
  id bigint generated always as identity primary key,
  emp_id text not null,
  warn_date date not null,
  created_at timestamptz not null default now(),
  constraint employee_warnings_emp_date_unique unique (emp_id, warn_date)
);

alter table public.employee_warnings enable row level security;

grant select, insert, delete on public.employee_warnings to anon, authenticated;
grant usage, select on sequence public.employee_warnings_id_seq to anon, authenticated;

drop policy if exists "employee_warnings_public_all" on public.employee_warnings;
create policy "employee_warnings_public_all"
on public.employee_warnings for all
to anon, authenticated
using (true)
with check (true);
