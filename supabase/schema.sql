-- Run in Supabase SQL Editor after creating your free project.
-- Replace the two placeholder UUIDs with the Auth user IDs for you and Dorothy
-- AFTER inviting/creating both users under Authentication > Users.
create table if not exists public.family_access (
  user_id uuid primary key references auth.users(id) on delete cascade,
  role text not null check (role in ('admin','viewer'))
);
create table if not exists public.family_meetings (
  id uuid primary key default gen_random_uuid(),
  title text not null check (length(title) between 1 and 100),
  starts_at timestamptz not null,
  ends_at timestamptz not null,
  status text not null check (status in ('green','yellow','red')),
  constraint valid_times check (ends_at > starts_at and ends_at <= starts_at + interval '24 hours')
);
create index if not exists family_meetings_start_idx on public.family_meetings(starts_at);
alter table public.family_access enable row level security;
alter table public.family_meetings enable row level security;
revoke all on public.family_access from anon;
revoke all on public.family_meetings from anon;
grant select on public.family_access to authenticated;
grant select,insert,update,delete on public.family_meetings to authenticated;
drop policy if exists "Access own role" on public.family_access;
create policy "Access own role" on public.family_access for select to authenticated using (user_id = (select auth.uid()));
drop policy if exists "Family can read meetings" on public.family_meetings;
create policy "Family can read meetings" on public.family_meetings for select to authenticated
using (exists (select 1 from public.family_access a where a.user_id=(select auth.uid())));
drop policy if exists "Only admin can insert" on public.family_meetings;
create policy "Only admin can insert" on public.family_meetings for insert to authenticated
with check (exists (select 1 from public.family_access a where a.user_id=(select auth.uid()) and a.role='admin'));
drop policy if exists "Only admin can update" on public.family_meetings;
create policy "Only admin can update" on public.family_meetings for update to authenticated
using (exists (select 1 from public.family_access a where a.user_id=(select auth.uid()) and a.role='admin'))
with check (exists (select 1 from public.family_access a where a.user_id=(select auth.uid()) and a.role='admin'));
drop policy if exists "Only admin can delete" on public.family_meetings;
create policy "Only admin can delete" on public.family_meetings for delete to authenticated
using (exists (select 1 from public.family_access a where a.user_id=(select auth.uid()) and a.role='admin'));
-- Once both Auth accounts exist, run these with their real IDs:
-- insert into public.family_access(user_id,role) values
--   ('YOUR_AUTH_USER_UUID','admin'),
--   ('DOROTHY_AUTH_USER_UUID','viewer')
-- on conflict (user_id) do update set role=excluded.role;
