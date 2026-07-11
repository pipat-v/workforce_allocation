alter table public.master_data_files
add column if not exists effective_until date;

create index if not exists master_data_files_type_effective_until_idx
on public.master_data_files (file_type, effective_until);
