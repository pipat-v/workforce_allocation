create table if not exists public.daily_confirmations (
  id uuid primary key default gen_random_uuid(),
  confirm_date date not null,
  dept text not null,
  confirmed_by text not null,
  confirmed_at timestamptz not null default now(),
  late_count int not null default 0,
  absent_count int not null default 0,
  leave_breakdown jsonb not null default '{}',
  unique (confirm_date, dept)
);

alter table public.daily_confirmations enable row level security;
grant select, insert, update, delete on public.daily_confirmations to anon, authenticated;

create policy "daily_confirmations_all"
on public.daily_confirmations for all
to anon, authenticated
using (true)
with check (true);
