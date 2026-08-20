-- Prezlab PPT QC: Supabase schema, mirroring the local SQLite store 1:1.
-- Apply in the Supabase SQL editor. The qc/store.py functions are the
-- storage interface; a Supabase driver implements the same functions over
-- these tables when SUPABASE_URL + SUPABASE_SERVICE_KEY are configured.
--
-- Confidentiality note (PRD): decks NEVER leave the machine. These tables
-- hold users, judgments, comments, and audit metadata/manifests only.
-- Recommended region: closest to UAE (currently ap-south / eu-central);
-- confirm against the IT-security data-residency ruling before adoption.

create table if not exists users (
    id bigint generated always as identity primary key,
    name text not null unique,
    role text not null check (role in ('designer', 'lead', 'admin')),
    pin_hash text,
    created_at timestamptz not null default now()
);

create table if not exists sessions (
    token text primary key,
    user_name text not null references users(name),
    created_at timestamptz not null default now()
);

create table if not exists comments (
    id bigint generated always as identity primary key,
    deck text not null,
    slide_index int not null,
    record_id text,
    author text not null references users(name),
    text text not null,
    created_at timestamptz not null default now()
);
create index if not exists ix_comments_deck on comments (deck);

create table if not exists audits (
    id bigint generated always as identity primary key,
    deck text not null,
    profile_id text not null,
    profile_version int not null,
    user_name text not null,
    slides int not null,
    errors int not null,
    warnings int not null,
    info int not null,
    arabic int not null,
    total int not null,
    kind text not null default 'audit',
    manifest jsonb not null,
    created_at timestamptz not null default now()
);
create index if not exists ix_audits_deck on audits (deck);

create table if not exists triage (
    id bigint generated always as identity primary key,
    record_id text not null,
    issue_type text not null,
    module text not null,
    severity text not null,
    confidence text not null,
    arabic_flag boolean not null,
    state text not null check (state in ('confirmed', 'false_positive', 'cleared')),
    deck text not null,
    profile_id text not null,
    author text,
    created_at timestamptz not null default now()
);
create index if not exists ix_triage_issue on triage (issue_type);
