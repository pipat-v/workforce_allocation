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

alter table public.employee_skills enable row level security;

grant select, insert, update, delete on public.employee_skills to anon, authenticated;
grant usage, select on sequence public.employee_skills_id_seq to anon, authenticated;

drop policy if exists "employee_skills_public_all" on public.employee_skills;
create policy "employee_skills_public_all"
on public.employee_skills for all
to anon, authenticated
using (true)
with check (true);
