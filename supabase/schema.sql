create extension if not exists "pgcrypto";

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text,
  role text not null default 'planner' check (role in ('admin', 'planner', 'viewer')),
  created_at timestamptz not null default now()
);

create table if not exists public.allocation_runs (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid references auth.users(id) on delete cascade,
  target_date date,
  status text not null default 'uploaded'
    check (status in ('uploaded', 'queued', 'processing', 'completed', 'failed')),
  scan_file_path text,
  master_file_path text,
  manpower_file_path text,
  skill_file_path text,
  dayoff_shift_file_path text,
  output_file_path text,
  solver_status text,
  error_message text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.master_data_files (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid references auth.users(id) on delete cascade,
  file_type text not null check (file_type in ('employee_master', 'manpower_plan', 'skill_matrix', 'dayoff_shift')),
  file_path text not null,
  original_filename text,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.allocation_results (
  id bigint generated always as identity primary key,
  run_id uuid not null references public.allocation_runs(id) on delete cascade,
  plan_id integer,
  emp_id text,
  name text,
  home_dept text,
  target_dept text,
  work_station text,
  shift text,
  required_skill text,
  emp_skill text,
  skill_match boolean,
  skill_level integer,
  attendance_status text,
  scan_in text,
  scan_out text,
  allocation_type text,
  created_at timestamptz not null default now()
);

create table if not exists public.gap_summaries (
  id bigint generated always as identity primary key,
  run_id uuid not null references public.allocation_runs(id) on delete cascade,
  date date,
  dept text,
  work_station text,
  shift text,
  required_hc integer,
  assigned_hc integer,
  gap integer,
  shortage integer,
  surplus integer,
  created_at timestamptz not null default now()
);

create table if not exists public.timestamp_with_dept (
  id bigint generated always as identity primary key,
  run_id uuid not null references public.allocation_runs(id) on delete cascade,
  target_date text,
  emp_id text not null,
  name text,
  dept text,
  position text,
  shift text,
  shift_start text,
  scan_in text,
  attendance_status text,
  minutes_late integer not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (run_id, emp_id)
);

create table if not exists public.employee_skills (
  id bigint generated always as identity primary key,
  emp_id text not null,
  employee_name text,
  dept text,
  job_site text,
  skill text not null,
  level integer not null default 0 check (level between 0 and 5),
  can_do boolean generated always as (level > 0) stored,
  source_file_id uuid references public.master_data_files(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (emp_id, skill)
);

create index if not exists employee_skills_emp_id_idx
on public.employee_skills (emp_id);

create index if not exists employee_skills_skill_idx
on public.employee_skills (skill);

alter table public.profiles enable row level security;
alter table public.allocation_runs enable row level security;
alter table public.master_data_files enable row level security;
alter table public.allocation_results enable row level security;
alter table public.gap_summaries enable row level security;
alter table public.timestamp_with_dept enable row level security;
alter table public.employee_skills enable row level security;

grant select, insert, update, delete on public.timestamp_with_dept to anon, authenticated;
grant usage, select on sequence public.timestamp_with_dept_id_seq to anon, authenticated;
grant select, insert, update, delete on public.employee_skills to anon, authenticated;
grant usage, select on sequence public.employee_skills_id_seq to anon, authenticated;

create policy "profiles_select_own"
on public.profiles for select
using (auth.uid() = id);

create policy "profiles_update_own"
on public.profiles for update
using (auth.uid() = id);

create policy "runs_owner_all"
on public.allocation_runs for all
using (auth.uid() = owner_id)
with check (auth.uid() = owner_id);

create policy "runs_public_all"
on public.allocation_runs for all
using (owner_id is null)
with check (owner_id is null);

create policy "master_files_owner_all"
on public.master_data_files for all
using (auth.uid() = owner_id)
with check (auth.uid() = owner_id);

create policy "master_files_public_all"
on public.master_data_files for all
using (owner_id is null)
with check (owner_id is null);

create policy "results_owner_select"
on public.allocation_results for select
using (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = allocation_results.run_id
      and runs.owner_id = auth.uid()
  )
);

create policy "results_public_select"
on public.allocation_results for select
using (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = allocation_results.run_id
      and runs.owner_id is null
  )
);

create policy "gaps_owner_select"
on public.gap_summaries for select
using (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = gap_summaries.run_id
      and runs.owner_id = auth.uid()
  )
);

create policy "gaps_public_select"
on public.gap_summaries for select
using (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = gap_summaries.run_id
      and runs.owner_id is null
  )
);

create policy "timestamp_with_dept_owner_all"
on public.timestamp_with_dept for all
to authenticated
using (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = timestamp_with_dept.run_id
      and runs.owner_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = timestamp_with_dept.run_id
      and runs.owner_id = auth.uid()
  )
);

create policy "timestamp_with_dept_public_all"
on public.timestamp_with_dept for all
to anon, authenticated
using (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = timestamp_with_dept.run_id
      and runs.owner_id is null
  )
)
with check (
  exists (
    select 1
    from public.allocation_runs runs
    where runs.id = timestamp_with_dept.run_id
      and runs.owner_id is null
  )
);

create policy "employee_skills_public_all"
on public.employee_skills for all
to anon, authenticated
using (true)
with check (true);

insert into storage.buckets (id, name, public)
values ('workforce-inputs', 'workforce-inputs', false)
on conflict (id) do nothing;

insert into storage.buckets (id, name, public)
values ('workforce-outputs', 'workforce-outputs', false)
on conflict (id) do nothing;

create policy "input_files_owner_select"
on storage.objects for select
using (
  bucket_id = 'workforce-inputs'
  and auth.uid()::text = (storage.foldername(name))[1]
);

create policy "input_files_owner_insert"
on storage.objects for insert
with check (
  bucket_id = 'workforce-inputs'
  and auth.uid()::text = (storage.foldername(name))[1]
);

create policy "input_files_public_select"
on storage.objects for select
using (
  bucket_id = 'workforce-inputs'
  and (storage.foldername(name))[1] = 'public'
);

create policy "input_files_public_insert"
on storage.objects for insert
with check (
  bucket_id = 'workforce-inputs'
  and (storage.foldername(name))[1] = 'public'
);

create policy "input_files_owner_update"
on storage.objects for update
using (
  bucket_id = 'workforce-inputs'
  and auth.uid()::text = (storage.foldername(name))[1]
)
with check (
  bucket_id = 'workforce-inputs'
  and auth.uid()::text = (storage.foldername(name))[1]
);

create policy "input_files_public_update"
on storage.objects for update
using (
  bucket_id = 'workforce-inputs'
  and (storage.foldername(name))[1] = 'public'
)
with check (
  bucket_id = 'workforce-inputs'
  and (storage.foldername(name))[1] = 'public'
);

create policy "output_files_owner_select"
on storage.objects for select
using (
  bucket_id = 'workforce-outputs'
  and auth.uid()::text = (storage.foldername(name))[1]
);

create policy "output_files_public_select"
on storage.objects for select
using (
  bucket_id = 'workforce-outputs'
  and (storage.foldername(name))[1] = 'public'
);
