create table if not exists public.ot_dashboard_settings (
  id text primary key default 'default',
  ot_targets jsonb not null default '{}'::jsonb,
  break_minutes jsonb not null default '{}'::jsonb,
  dept_managers jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.ot_dashboard_settings enable row level security;

grant select, insert, update on public.ot_dashboard_settings to anon, authenticated;

drop policy if exists "ot_dashboard_settings_all" on public.ot_dashboard_settings;
create policy "ot_dashboard_settings_all"
on public.ot_dashboard_settings for all
to anon, authenticated
using (true)
with check (true);
