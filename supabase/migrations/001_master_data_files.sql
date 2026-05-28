create table if not exists public.master_data_files (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  file_type text not null check (file_type in ('employee_master', 'manpower_plan', 'skill_matrix')),
  file_path text not null,
  original_filename text,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.master_data_files enable row level security;

drop policy if exists "master_files_owner_all" on public.master_data_files;
create policy "master_files_owner_all"
on public.master_data_files for all
using (auth.uid() = owner_id)
with check (auth.uid() = owner_id);

drop policy if exists "input_files_owner_update" on storage.objects;
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

