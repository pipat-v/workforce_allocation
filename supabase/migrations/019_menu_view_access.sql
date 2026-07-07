-- Idempotent re-run of 018 (that migration was written but never actually
-- applied to the live project — this caused every login attempt to fail with
-- "Could not find the table 'public.login_users' in the schema cache"),
-- plus a new menu_view_access column for the separate View/Edit permission split.

create table if not exists public.login_users (
  id bigint generated always as identity primary key,
  position text,
  username text not null,
  password text not null,
  menu_access text not null default 'All',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (username)
);

create table if not exists public.login_menus (
  id bigint generated always as identity primary key,
  menu_no numeric not null,
  menu_name text not null,
  updated_at timestamptz not null default now(),
  unique (menu_no)
);

-- menu_access is the EDIT grant (existing meaning, unchanged); this new column
-- is the VIEW grant. Same "All" | comma-list-of-menu_no format. Not enforced
-- anywhere yet — every page stays publicly viewable regardless of this value.
alter table public.login_users
  add column if not exists menu_view_access text not null default 'All';

comment on column public.login_users.menu_view_access is
'Per-menu VIEW grant, same "All" | comma-list format as menu_access (the EDIT grant). Not currently enforced — all pages remain publicly viewable regardless; stored for the Setting page''s per-user Edit/View matrix.';

alter table public.login_users enable row level security;
alter table public.login_menus enable row level security;

grant select, insert, update, delete on public.login_users to anon, authenticated;
grant usage, select on sequence public.login_users_id_seq to anon, authenticated;
grant select, insert, update, delete on public.login_menus to anon, authenticated;
grant usage, select on sequence public.login_menus_id_seq to anon, authenticated;

drop policy if exists "login_users_public_all" on public.login_users;
create policy "login_users_public_all"
on public.login_users for all
to anon, authenticated
using (true)
with check (true);

drop policy if exists "login_menus_public_all" on public.login_menus;
create policy "login_menus_public_all"
on public.login_menus for all
to anon, authenticated
using (true)
with check (true);

comment on table public.login_users is
'Login credentials for gated menus, synced from the "User ระบบคน.xlsx" Main sheet (ตำแหน่ง, User, Password, Menu).';
comment on table public.login_menus is
'Reference list of menu numbers/names, synced from the "User ระบบคน.xlsx" Detail sheet (หมายเลข, เมนู).';

insert into public.login_users (position, username, password, menu_access, menu_view_access)
values
  ('เถ้าแก่', 'Bew', '1234', 'All', 'All'),
  ('HR', 'Atom', '1234', 'All', 'All')
on conflict (username) do nothing;

insert into public.login_menus (menu_no, menu_name)
values
  (0, 'Dashboard'),
  (1, 'Upload Timestamp'),
  (2, 'ผลลัพธ์การจัดสรร'),
  (3, 'Timestamp With Dept'),
  (4, 'Master Data'),
  (4.1, 'Master Files'),
  (4.2, 'Manpower'),
  (4.3, 'วันพระ'),
  (4.4, 'วันหยุดประจำปี'),
  (4.5, 'Shift & Dayoff'),
  (4.6, 'ลาล่วงหน้า'),
  (5, 'Skill Matrix'),
  (6, 'Report & Dashboard'),
  (7, 'OT Dashboard'),
  (7.1, 'แผนภูมิ'),
  (7.2, 'สรุปรายหน่วยงาน'),
  (7.3, 'สรุปรายพนักงาน'),
  (8, 'Setting'),
  (9, 'คู่มือการใช้งาน')
on conflict (menu_no) do nothing;

-- Defensive backfill in case the column already existed with nulls from a
-- partial prior run (add column ... default already backfills new columns,
-- this is just a safety net).
update public.login_users set menu_view_access = 'All' where menu_view_access is null;
