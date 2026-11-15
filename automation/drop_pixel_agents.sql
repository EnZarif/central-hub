-- Remove the pixel_agents table.
--
-- Run once in the Supabase SQL editor:
-- https://supabase.com/dashboard/project/qwceyzswwtyqiozxnlah/sql/new
--
-- WHY THIS IS HERE
--
-- pixel_agents backed the "Pixel Office" widget, which showed live Claude
-- Code session status on the Hub's home page. Its policy was:
--
--     create policy "anon full access" on pixel_agents
--       for all using (true) with check (true);
--
-- The comment above that policy said it mirrored "the permissive anon-key
-- access already used by hub_data". That was not true. hub_data enforces
-- row level security properly: an unauthenticated read returns no rows and
-- an unauthenticated write is rejected. pixel_agents was the only table in
-- the project genuinely open to anyone.
--
-- Open to anyone means exactly that. The publishable key is in the page
-- source of a site on the public web, so a request carrying nothing else
-- returned rows like:
--
--     {"session_id":"b46a9cb9-...",
--      "project_path":"C:\\Users\\ZarifErtesam\\Downloads\\Personal\\...",
--      "status":"running","last_tool":"Bash"}
--
-- which discloses the account name, the local directory layout and when
-- the machine is being used. Writes and deletes were open too.
--
-- The widget has been removed from index.html and the watcher no longer
-- writes here, so nothing refers to this table any more. Its rows have
-- already been deleted; this drops the table itself.

drop table if exists public.pixel_agents;

-- Confirm it is gone: this should return no rows.
select table_name
from information_schema.tables
where table_schema = 'public' and table_name = 'pixel_agents';
