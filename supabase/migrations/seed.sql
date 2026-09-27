-- ==============================================================================
-- BCAway Supabase Seed: Initial Known BCA Faculty & Aliases
-- (Optional to apply: run manually via Supabase SQL Editor or psql)
-- ==============================================================================

insert into public.teachers (name, aliases) values
  ('Gomes', array[]::text[])
on conflict (name) do update set aliases = excluded.aliases;
