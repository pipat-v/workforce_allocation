create table if not exists public.employee_work_schedules (
  emp_id text primary key,
  employee_name text,
  dept text,
  job_site text,
  dayoff text,
  shift text,
  shift_start text,
  shift_end text,
  sync_token uuid not null,
  synced_at timestamptz not null default now()
);

alter table public.employee_work_schedules enable row level security;

grant select, insert, update, delete on public.employee_work_schedules to anon, authenticated;

drop policy if exists "employee_work_schedules_public_all" on public.employee_work_schedules;
create policy "employee_work_schedules_public_all"
on public.employee_work_schedules for all
to anon, authenticated
using (true)
with check (true);

create or replace function public.refresh_production_user(
  p_work_date date default ((now() at time zone 'Asia/Bangkok')::date)
)
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  v_sync_token uuid := gen_random_uuid();
  v_weekday text;
  v_count integer;
begin
  v_weekday := (array['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'])[extract(dow from p_work_date)::integer + 1];

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
    p_work_date,
    skills.emp_id,
    coalesce(schedules.employee_name, skills.employee_name),
    coalesce(schedules.dept, skills.dept),
    coalesce(schedules.job_site, skills.job_site),
    schedules.shift,
    schedules.shift_start,
    skills.skill,
    skills.level,
    null,
    v_sync_token,
    now()
  from public.employee_skills as skills
  join public.employee_work_schedules as schedules
    on schedules.emp_id = skills.emp_id
  where not exists (
    select 1
    from public.leave_records as leave_rows
    where leave_rows.emp_id = skills.emp_id
      and leave_rows.leave_date = p_work_date
  )
  and not (
    v_weekday = any (
      regexp_split_to_array(trim(coalesce(schedules.dayoff, '')), '[,/|[:space:]]+')
    )
    or (
      'พระ' = any (
        regexp_split_to_array(trim(coalesce(schedules.dayoff, '')), '[,/|[:space:]]+')
      )
      and exists (
        select 1
        from public.holidays
        where date = p_work_date
          and type = 'buddhist_holy_day'
      )
    )
  )
  on conflict (work_date, emp_id, skill)
  do update set
    employee_name = excluded.employee_name,
    dept = excluded.dept,
    job_site = excluded.job_site,
    shift = excluded.shift,
    shift_start = excluded.shift_start,
    level = excluded.level,
    source_run_id = null,
    sync_token = excluded.sync_token,
    synced_at = excluded.synced_at;

  delete from public.production_user
  where work_date = p_work_date
    and sync_token <> v_sync_token;

  select count(*) into v_count
  from public.production_user
  where work_date = p_work_date;

  return v_count;
end;
$$;

grant execute on function public.refresh_production_user(date) to anon, authenticated;

create extension if not exists pg_cron with schema pg_catalog;

do $$
begin
  if exists (select 1 from cron.job where jobname = 'refresh-production-user-daily') then
    perform cron.unschedule('refresh-production-user-daily');
  end if;
end
$$;

select cron.schedule(
  'refresh-production-user-daily',
  '5 17 * * *',
  $cron$select public.refresh_production_user((now() at time zone 'Asia/Bangkok')::date);$cron$
);

select public.refresh_production_user((now() at time zone 'Asia/Bangkok')::date);
