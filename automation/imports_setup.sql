-- Central Hub: the "imports" bucket that emailed files land in.
-- Run this once in the Supabase dashboard -> SQL Editor -> New query.
--
-- Shape: imports/<your user id>/inbox/<timestamp>__<original filename>
-- The user id prefix is what the policies below key off, so the path is not
-- cosmetic - keep the Apps Script writing to exactly this shape.

-- 1. The bucket. Private: the Hub page is public on the web, and these files
--    are customer job data.
insert into storage.buckets (id, name, public)
values ('imports', 'imports', false)
on conflict (id) do nothing;

-- 2. Read. You, signed in, can see only your own folder.
drop policy if exists "imports owner reads" on storage.objects;
create policy "imports owner reads"
  on storage.objects for select to authenticated
  using (
    bucket_id = 'imports'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

-- 3. Delete. So the Hub can clear a file once you are done with it.
drop policy if exists "imports owner deletes" on storage.objects;
create policy "imports owner deletes"
  on storage.objects for delete to authenticated
  using (
    bucket_id = 'imports'
    and (storage.foldername(name))[1] = auth.uid()::text
  );

-- 4. Deliberately NO insert policy for "authenticated".
--    Only the service role writes here, and the service key lives in the
--    Apps Script - never in index.html. That means a visitor who somehow got
--    a session could never seed a file into your morning list.

-- 5. Check it: should return the bucket, private = false.
select id, name, public from storage.buckets where id = 'imports';
