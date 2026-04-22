import { createClient } from '@supabase/supabase-js';

const URL  = import.meta.env.VITE_SUPABASE_URL;
const KEY  = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!URL || !KEY) {
  throw new Error('Missing VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY in your .env.local file.');
}

export const supabase = createClient(URL, KEY);

/* ── Data mappers (DB snake_case ↔ App camelCase) ─────────────────────────── */

export const mapAccount = a => ({
  id:             a.id,
  name:           a.name,
  type:           a.type,
  currency:       a.currency,
  initialBalance: Number(a.initial_balance),
  riskPercent:    Number(a.risk_percent),
});

export const mapTrade = t => ({
  id:          t.id,
  date:        t.date,
  symbol:      t.symbol,
  account:     t.account_id,
  r:           Number(t.r),
  type:        t.type || null,
  notes:       t.notes      || '',
  chartLink:   t.chart_link  || '',
  riskPercent: t.risk_percent ? Number(t.risk_percent) : null,
});

/* ── API helpers ──────────────────────────────────────────────────────────── */

export async function fetchAccounts() {
  const { data, error } = await supabase
    .from('accounts')
    .select('*')
    .order('created_at');
  if (error) throw error;
  return (data || []).map(mapAccount);
}

export async function fetchTrades() {
  const { data, error } = await supabase
    .from('trades')
    .select('*')
    .order('date')
    .order('created_at');
  if (error) throw error;
  return (data || []).map(mapTrade);
}

export async function insertTrade({ date, symbol, account, r, type, notes, chartLink, riskPercent }) {
  const { data, error } = await supabase
    .from('trades')
    .insert({ date, symbol, account_id: account, r, type: type || null, notes, chart_link: chartLink, risk_percent: riskPercent || null })
    .select()
    .single();
  if (error) throw error;
  return mapTrade(data);
}

export async function updateAccount(id, updates) {
  const db = {};
  if (updates.name            != null) db.name             = updates.name;
  if (updates.type            != null) db.type             = updates.type;
  if (updates.currency        != null) db.currency         = updates.currency;
  if (updates.initialBalance  != null) db.initial_balance  = updates.initialBalance;
  if (updates.riskPercent     != null) db.risk_percent     = updates.riskPercent;
  const { error } = await supabase.from('accounts').update(db).eq('id', id);
  if (error) throw error;
}

export async function deleteTrade(id) {
  const { error } = await supabase.from('trades').delete().eq('id', id);
  if (error) throw error;
}

export async function insertAccount({ name, type, currency, initialBalance, riskPercent }) {
  const { data, error } = await supabase
    .from('accounts')
    .insert({ name, type, currency, initial_balance: initialBalance, risk_percent: riskPercent })
    .select()
    .single();
  if (error) throw error;
  return mapAccount(data);
}

export async function updateAccountType(id, type) {
  const { error } = await supabase.from('accounts').update({ type }).eq('id', id);
  if (error) throw error;
}

export async function deleteAccount(id) {
  const { error } = await supabase.from('accounts').delete().eq('id', id);
  if (error) throw error;
}

/* ── Batch insert for Excel imports ─────────────────────────────────────── */

export async function batchInsertTrades(tradesArray, accountId) {
  const rows = tradesArray.map(t => ({
    date:       t.date,
    symbol:     t.symbol,
    account_id: accountId,
    r:          t.r,
    type:       t.type || null,
    notes:      t.notes     || '',
    chart_link: t.chartLink || '',
  }));

  const BATCH = 500;
  const results = [];
  for (let i = 0; i < rows.length; i += BATCH) {
    const { data, error } = await supabase
      .from('trades')
      .insert(rows.slice(i, i + BATCH))
      .select();
    if (error) throw error;
    results.push(...(data || []).map(mapTrade));
  }
  return results;
}