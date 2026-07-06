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

-- Backfill dates that have already been processed. Re-running this migration is safe.
insert into public.production_user (
  work_date,
  emp_id,
  employee_name,
  dept,
  job_site,
  shift,
  shift_start,
  skill,
  level,
  source_run_id,
  sync_token,
  synced_at
)
select
  timestamp_rows.target_date::date,
  skills.emp_id,
  coalesce(skills.employee_name, timestamp_rows.name),
  coalesce(skills.dept, timestamp_rows.dept),
  skills.job_site,
  timestamp_rows.shift,
  timestamp_rows.shift_start,
  skills.skill,
  skills.level,
  timestamp_rows.run_id,
  gen_random_uuid(),
  now()
from public.timestamp_with_dept as timestamp_rows
join public.employee_skills as skills
  on skills.emp_id = timestamp_rows.emp_id
where timestamp_rows.target_date ~ '^\d{4}-\d{2}-\d{2}$'
  and coalesce(timestamp_rows.attendance_status, '') <> 'DayOff'
  and not exists (
    select 1
    from public.leave_records as leave_rows
    where leave_rows.emp_id = timestamp_rows.emp_id
      and leave_rows.leave_date = timestamp_rows.target_date::date
  )
on conflict (work_date, emp_id, skill)
do update set
  employee_name = excluded.employee_name,
  dept = excluded.dept,
  job_site = excluded.job_site,
  shift = excluded.shift,
  shift_start = excluded.shift_start,
  level = excluded.level,
  source_run_id = excluded.source_run_id,
  sync_token = excluded.sync_token,
  synced_at = excluded.synced_at;
