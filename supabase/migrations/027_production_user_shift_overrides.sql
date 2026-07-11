-- refresh_production_user() ไม่รู้จักกะพิเศษเฉพาะวันที่ตั้งไว้ล่วงหน้าใน shift_schedule_overrides เลย
-- (มันอ่านแค่ employee_work_schedules ซึ่งเก็บตารางกะ "ประจำ" ต่อคน ไม่มีมิติวันที่) ทำให้พนักงานที่ถูก
-- ตั้งกะพิเศษให้มาทำงานในวันหยุดประจำของตัวเอง ยังคงถูกตัดออกจาก production_user เหมือนเดิมทุกครั้งที่มี
-- การ sync ผ่านทาง RPC นี้ (ซึ่งเป็น path หลัก รันทุกครั้งที่บันทึก master data และรันทุกคืนผ่าน cron)
-- แก้โดย LEFT JOIN shift_schedule_overrides ด้วย emp_id + work_date แล้ว:
--   1) ใช้ dept/job_site/shift/shift_start จาก override ก่อน ถ้ามี
--   2) ถ้ามี override ของวันนี้ ให้ถือว่ามาทำงานแน่นอน ข้ามการเช็ควันหยุดประจำสัปดาห์/วันพระไปเลย
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
    coalesce(overrides.dept, schedules.dept, skills.dept),
    coalesce(overrides.job_site, schedules.job_site, skills.job_site),
    coalesce(overrides.shift, schedules.shift),
    coalesce(overrides.shift_start, schedules.shift_start),
    skills.skill,
    skills.level,
    null,
    v_sync_token,
    now()
  from public.employee_skills as skills
  join public.employee_work_schedules as schedules
    on schedules.emp_id = skills.emp_id
  left join public.shift_schedule_overrides as overrides
    on overrides.emp_id = skills.emp_id
    and overrides.work_date = p_work_date
  where not exists (
    select 1
    from public.leave_records as leave_rows
    where leave_rows.emp_id = skills.emp_id
      and leave_rows.leave_date = p_work_date
  )
  and (
    overrides.emp_id is not null
    or not (
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
