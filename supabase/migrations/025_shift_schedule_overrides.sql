create table if not exists public.shift_schedule_overrides (
  id uuid primary key default gen_random_uuid(),
  emp_id text not null,
  work_date date not null,
  dept text,
  job_site text,
  shift text not null,
  shift_start text,
  shift_end text,
  recorded_by text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (emp_id, work_date)
);

alter table public.shift_schedule_overrides enable row level security;

grant select, insert, update, delete on public.shift_schedule_overrides to anon, authenticated;

create policy "shift_schedule_overrides_all"
on public.shift_schedule_overrides for all
to anon, authenticated
using (true)
with check (true);
