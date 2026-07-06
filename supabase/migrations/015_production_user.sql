create table if not exists public.production_user (
  id bigint generated always as identity primary key,
  work_date date not null,
  emp_id text not null,
  employee_name text,
  dept text,
  job_site text,
  shift text,
  shift_start text,
  skill text not null,
  level integer not null default 0 check (level between 0 and 5),
  can_do boolean generated always as (level > 0) stored,
  source_run_id uuid references public.allocation_runs(id) on delete set null,
  sync_token uuid not null,
  synced_at timestamptz not null default now(),
  unique (work_date, emp_id, skill)
);

create index if not exists production_user_work_date_idx
on public.production_user (work_date);

create index if not exists production_user_emp_id_idx
on public.production_user (emp_id);

alter table public.production_user enable row level security;

grant select, insert, update, delete on public.production_user to anon, authenticated;
grant usage, select on sequence public.production_user_id_seq to anon, authenticated;

drop policy if exists "production_user_public_all" on public.production_user;
create policy "production_user_public_all"
on public.production_user for all
to anon, authenticated
using (true)
with check (true);

comment on table public.production_user is
'Employee skills expected to be available on each work date, excluding scheduled day off and planned leave.';
