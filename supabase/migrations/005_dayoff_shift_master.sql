alter table public.master_data_files
drop constraint if exists master_data_files_file_type_check;

alter table public.master_data_files
add constraint master_data_files_file_type_check
check (file_type in ('employee_master', 'manpower_plan', 'skill_matrix', 'dayoff_shift'));

alter table public.allocation_runs
add column if not exists dayoff_shift_file_path text;
