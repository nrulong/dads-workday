-- Run once in Supabase SQL Editor before enabling Apps Script.
alter table public.family_meetings add column if not exists outlook_event_id text;
create unique index if not exists family_meetings_outlook_event_id_unique
on public.family_meetings(outlook_event_id)
where outlook_event_id is not null;
-- Existing RLS stays enabled. Do not grant anonymous writes.
