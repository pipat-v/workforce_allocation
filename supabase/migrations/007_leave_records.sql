create table if not exists public.leave_records (
  id uuid primary key default gen_random_uuid(),
  emp_id text not null,
  leave_date date not null,
  leave_type text not null check (leave_type in ('ลาป่วย', 'ลากิจ', 'ลาพักร้อน', 'ขาดงาน')),
  recorded_by text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (emp_id, leave_date)
);

alter table public.leave_records enable row level security;

grant select, insert, update, delete on public.leave_records to anon, authenticated;

create policy "leave_records_all"
on public.leave_records for all
to anon, authenticated
using (true)
with check (true);
