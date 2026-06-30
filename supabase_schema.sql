-- supabase_schema.sql
-- Run this in the Supabase dashboard → SQL Editor (one time).
-- It creates the private metadata table for quote conversions.
-- The storage bucket is created separately (see note at the bottom).

create table if not exists public.quote_conversions (
    id                   uuid primary key default gen_random_uuid(),
    created_at           timestamptz not null default now(),
    file_hash            text,
    source_filename      text,
    input_type           text,          -- 'pdf' or 'po_xlsx'
    supplier             text,
    job_code             text,
    activity_code        text,
    line_count           integer,
    subtotal_ex_gst      numeric(12,2),
    catalogue_row_count  integer,
    po_file_path         text,          -- path inside the 'quote-files' bucket
    catalogue_file_path  text
);

-- Keep the table private: enable Row Level Security and add NO public
-- policies. With RLS on and no policy, the anon/public key cannot read or
-- write. The app uses the service_role key, which bypasses RLS server-side.
alter table public.quote_conversions enable row level security;

-- Optional: index for browsing newest-first.
create index if not exists quote_conversions_created_at_idx
    on public.quote_conversions (created_at desc);

-- ---------------------------------------------------------------------------
-- STORAGE BUCKET (do this in the dashboard, not SQL):
--   Storage → New bucket → name: quote-files → Public bucket: OFF (private).
-- No bucket policies are needed: the app writes with the service_role key,
-- which bypasses storage RLS. Leaving the bucket private means the files are
-- reachable only from your Supabase dashboard / authenticated API.
-- ---------------------------------------------------------------------------
