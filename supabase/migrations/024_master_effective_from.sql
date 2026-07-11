alter table public.master_data_files
add column if not exists effective_from date;

update public.master_data_files
set effective_from = (created_at at time zone 'Asia/Bangkok')::date
where effective_from is null;

alter table public.master_data_files
alter column effective_from set default ((now() at time zone 'Asia/Bangkok')::date);

alter table public.master_data_files
alter column effective_from set not null;

create index if not exists master_data_files_type_effective_idx
on public.master_data_files (file_type, effective_from desc, created_at desc);
