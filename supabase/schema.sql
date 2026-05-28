create extension if not exists "pgcrypto";

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text,
  role text not null default 'planner' check (role in ('admin', 'planner', 'viewer')),
  created_at timestamptz not null default now()
);

create table if not exists public.allocation_runs (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  target_date date,
  status text not null default 'uploaded'
    check (status in ('uploaded', 'queued', 'processing', 'completed', 'failed')),
  scan_file_path text,
  master_file_path text,
  manpower_file_path text,
  skill_file_path text,
  output_file_path text,
  solver_status text,
  error_message text,
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

alter table public.profiles enable row level security;
alter table public.allocation_runs enable row level security;
alter table public.allocation_results enable row level security;
alter table public.gap_summaries enable row level security;

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

create policy "output_files_owner_select"
on storage.objects for select
using (
  bucket_id = 'workforce-outputs'
  and auth.uid()::text = (storage.foldername(name))[1]
);

