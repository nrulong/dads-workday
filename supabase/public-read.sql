-- Public read-only availability. Run once in Supabase SQL Editor.
-- Anyone on the internet can read meeting titles, start/end times, and status.
-- Writes remain restricted to authenticated admin via existing RLS policies.
grant select on public.family_meetings to anon;
drop policy if exists "Public read meetings" on public.family_meetings;
create policy "Public read meetings"
on public.family_meetings for select to anon
using (true);
-- Do NOT grant anon insert/update/delete.
-- Admin editing still requires a Supabase Auth user assigned 'admin' in family_access.
