alter table public.timestamp_with_dept
  add column if not exists scan_out text;
