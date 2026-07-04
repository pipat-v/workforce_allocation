alter table public.ot_dashboard_settings
  add column if not exists ot_target_avg numeric not null default 2;
