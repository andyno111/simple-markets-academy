-- ─────────────────────────────────────────────────────────────────────────────
-- Simple Markets — Andrej's Performance Record
-- Run this entire file in Supabase → SQL Editor → New Query → Run
-- ─────────────────────────────────────────────────────────────────────────────

-- ACCOUNTS
create table if not exists public.accounts (
  id              text primary key default gen_random_uuid()::text,
  name            text        not null,
  type            text        not null check (type in ('Live','Funded','Challenge')),
  currency        text        not null default 'USD',
  initial_balance numeric     not null default 10000,
  risk_per_trade  numeric     not null default 100,
  created_at      timestamptz not null default now()
);

-- TRADES
create table if not exists public.trades (
  id          bigint generated always as identity primary key,
  date        date        not null,
  symbol      text        not null,
  account_id  text        references public.accounts(id) on delete set null,
  r           numeric     not null,
  type        text        not null check (type in ('Long','Short')),
  notes       text        not null default '',
  chart_link  text        not null default '',
  created_at  timestamptz not null default now()
);

-- ── ROW LEVEL SECURITY ────────────────────────────────────────────────────────
alter table public.accounts enable row level security;
alter table public.trades    enable row level security;

-- Public read
create policy "public_read_accounts" on public.accounts for select using (true);
create policy "public_read_trades"   on public.trades   for select using (true);

-- Anon write (we add proper auth in Step 3 — for now this lets the app write)
create policy "anon_insert_accounts" on public.accounts for insert with check (true);
create policy "anon_update_accounts" on public.accounts for update using (true);
create policy "anon_delete_accounts" on public.accounts for delete using (true);

create policy "anon_insert_trades"   on public.trades   for insert with check (true);
create policy "anon_update_trades"   on public.trades   for update using (true);
create policy "anon_delete_trades"   on public.trades   for delete using (true);

-- ── SEED YOUR FIRST ACCOUNTS ──────────────────────────────────────────────────
-- Edit these to match your real accounts, then uncomment and run.
-- insert into public.accounts (name, type, currency, initial_balance, risk_per_trade) values
--   ('Live Account',        'Live',      'USD', 25000,   250),
--   ('FTMO 100k Funded',    'Funded',    'USD', 100000,  500),
--   ('MyFXBook Challenge',  'Challenge', 'USD', 10000,   100);
