create table if not exists public.holidays (
  id uuid primary key default gen_random_uuid(),
  date date not null unique,
  name text not null default 'วันพระ',
  type text not null default 'buddhist_holy_day'
    check (type in ('buddhist_holy_day', 'public_holiday', 'company_holiday')),
  created_at timestamptz not null default now()
);

alter table public.holidays enable row level security;
grant select, insert, update, delete on public.holidays to anon, authenticated;

create policy "holidays_all"
on public.holidays for all
to anon, authenticated
using (true)
with check (true);
