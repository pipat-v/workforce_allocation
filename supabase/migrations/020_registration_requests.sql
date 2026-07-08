create table if not exists public.registration_requests (
  id bigint generated always as identity primary key,
  position text,
  username text not null,
  password text not null,
  status text not null default 'pending',
  created_at timestamptz not null default now()
);

alter table public.registration_requests enable row level security;

grant select, insert, update, delete on public.registration_requests to anon, authenticated;
grant usage, select on sequence public.registration_requests_id_seq to anon, authenticated;

drop policy if exists "registration_requests_public_all" on public.registration_requests;
create policy "registration_requests_public_all"
on public.registration_requests for all
to anon, authenticated
using (true)
with check (true);

comment on table public.registration_requests is
'Self-service sign-up requests from the login page, awaiting approval by HR/เถ้าแก่/ผู้จัดการ on the Setting page.';
