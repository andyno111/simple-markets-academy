import { useState, useMemo, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  AreaChart, Area, LineChart, Line, BarChart, Bar,
  RadarChart, Radar, PolarGrid, PolarAngleAxis,
  PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ReferenceLine, ResponsiveContainer, LabelList,
} from 'recharts';
import {
  supabase,
  fetchAccounts, fetchTrades,
  insertTrade, deleteTrade,
  insertAccount, updateAccountType, updateAccount, deleteAccount, batchInsertTrades,
} from './lib/supabase.js';

/* ── HELPERS ─────────────────────────────────────────────────────────────────── */
const fmtR = r => (r >= 0 ? '+' : '') + r.toFixed(2) + 'R';
const fmtU = n => (n >= 0 ? '+$' : '-$') + Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtPct = n => n.toFixed(1) + '%';

function fmtMon(k) { return k ? new Date(k + '-01').toLocaleDateString('en-US', { month: 'short', year: 'numeric' }) : '—'; }
function fmtRange(s, e) {
  if (!s) return '—';
  const sd = new Date(s), ed = new Date(e);
  const sm = sd.toLocaleDateString('en-US', { month: 'short' }), em = ed.toLocaleDateString('en-US', { month: 'short' });
  const sy = sd.getFullYear().toString().slice(2), ey = ed.getFullYear().toString().slice(2);
  if (s === e) return `${sm} ${sd.getDate()} '${sy}`;
  if (sm === em && sy === ey) return `${sm} '${sy}`;
  return `${sm}–${em} '${ey}`;
}

function applyPeriod(trades, period) {
  if (period === 'ALL') return trades;
  const now = new Date(), d = new Date(now);
  if (period === '1W') d.setDate(d.getDate() - 7);
  else if (period === '1M') d.setMonth(d.getMonth() - 1);
  else if (period === '3M') d.setMonth(d.getMonth() - 3);
  else if (period === 'YTD') { d.setMonth(0); d.setDate(1); }
  return trades.filter(t => new Date(t.date) >= d);
}

function computeStats(trades, breakEvenAsWin = false) {
  const sorted = trades.slice().sort((a, b) => a.date.localeCompare(b.date));
  const total = sorted.length, totalR = +sorted.reduce((s, t) => s + t.r, 0).toFixed(2);
  const wins = sorted.filter(t => breakEvenAsWin ? t.r >= 0 : t.r > 0).length;
  const losses = sorted.filter(t => t.r < 0).length;
  // When BE=wins OFF: exclude BE trades from denominator (industry standard)
  const wr = breakEvenAsWin
    ? (total ? wins / total * 100 : 0)
    : (wins + losses ? wins / (wins + losses) * 100 : 0);
  const wt = sorted.filter(t => t.r > 0), lt = sorted.filter(t => t.r < 0);
  const avgW = wt.length ? wt.reduce((s, t) => s + t.r, 0) / wt.length : 0;
  const avgL = lt.length ? Math.abs(lt.reduce((s, t) => s + t.r, 0) / lt.length) : 0;
  const pf = avgL > 0 ? (avgW * wins) / (avgL * losses) : avgW > 0 ? 99 : 0;
  const exp = total ? totalR / total : 0;
  let cW = 0, cL = 0, wSt = '', lSt = '';
  let bWS = { count: 0, start: '', end: '' }, wLS = { count: 0, start: '', end: '' };
  let peak = 0, cum = 0, maxDD = 0;
  for (const t of sorted) {
    const win = breakEvenAsWin ? t.r >= 0 : t.r > 0;
    if (win) { if (!cW) wSt = t.date; cW++; cL = 0; if (cW > bWS.count) bWS = { count: cW, start: wSt, end: t.date }; }
    else if (t.r < 0) { if (!cL) lSt = t.date; cL++; cW = 0; if (cL > wLS.count) wLS = { count: cL, start: lSt, end: t.date }; }
    else { cW = 0; cL = 0; }
    cum += t.r; if (cum > peak) peak = cum; maxDD = Math.max(maxDD, peak - cum);
  }
  let curDir = '', curCount = 0;
  for (let i = sorted.length - 1; i >= 0; i--) {
    const win = breakEvenAsWin ? sorted[i].r >= 0 : sorted[i].r > 0;
    if (!curDir) { if (win) { curDir = 'W'; curCount = 1; } else if (sorted[i].r < 0) { curDir = 'L'; curCount = 1; } else break; }
    else if (curDir === 'W' && win) curCount++;
    else if (curDir === 'L' && sorted[i].r < 0) curCount++;
    else break;
  }
  const mMap = {};
  for (const t of sorted) { const k = t.date.slice(0, 7); if (!mMap[k]) mMap[k] = 0; mMap[k] += t.r; }
  const mArr = Object.entries(mMap).map(([k, r]) => ({ key: k, r: +r.toFixed(2) }));
  const bestMon = mArr.length ? mArr.reduce((b, m) => m.r > b.r ? m : b, mArr[0]) : { key: '', r: 0 };
  const worstMon = mArr.length ? mArr.reduce((w, m) => m.r < w.r ? m : w, mArr[0]) : { key: '', r: 0 };
  const uDays = new Set(sorted.map(t => t.date)).size;
  const avgPerWeek = uDays ? +(sorted.length / (uDays / 5)).toFixed(1) : 0;
  return { total, wins, losses, wr, totalR, avgW, avgL, pf, exp, maxDD, bWS, wLS, curDir, curCount, bestMon, worstMon, avgPerWeek };
}

function buildMon(trades) {
  const m = {};
  for (const t of trades) { const k = t.date.slice(0, 7); m[k] = (m[k] || 0) + t.r; }
  return Object.entries(m).sort().map(([k, r]) => ({
    key: k,
    year: k.slice(0, 4),
    month: new Date(k + '-01').toLocaleDateString('en-US', { month: 'short' }),
    r: +r.toFixed(2),
  }));
}
function buildEq(trades, defaultRiskPercent = 1.0, init = 25000) {
  let bal = init, cumR = 0, n = 0;
  const c = [{ date: 'Start', n: 0, bal, cumR }];
  for (const t of trades) {
    const rp = t.riskPercent != null ? t.riskPercent : defaultRiskPercent;
    bal += t.r * (rp / 100) * init;
    cumR += t.r;
    n++;
    c.push({ date: t.date, n, bal: +bal.toFixed(2), cumR: +cumR.toFixed(2) });
  }
  return c;
}

function getWeekStart(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const m = new Date(d); m.setDate(d.getDate() + diff);
  return m.toISOString().slice(0, 10);
}
function buildDailyEqAndDD(trades, defaultRp = 1.0, init = 25000, weekly = false) {
  const byKey = {};
  for (const t of trades) {
    const key = weekly ? getWeekStart(t.date) : t.date;
    if (!byKey[key]) byKey[key] = [];
    byKey[key].push(t);
  }
  const keys = Object.keys(byKey).sort();
  let bal = init, peak = init;
  const rows = [];
  for (const k of keys) {
    for (const t of byKey[k]) {
      const rp = t.riskPercent != null ? t.riskPercent : defaultRp;
      bal += t.r * (rp / 100) * init;
    }
    peak = Math.max(peak, bal);
    rows.push({ date: k, bal: +bal.toFixed(2), dd: +(bal - peak).toFixed(2) });
  }
  return rows;
}
const fmtAxisDate = s => {
  if (!s || typeof s !== 'string') return '';
  const d = new Date(s + 'T00:00:00');
  return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
};
function computeRatios(trades, defaultRp = 1.0, init = 25000) {
  if (!trades || trades.length < 5) return { sharpe: null, sortino: null, calmar: null, recovery: null };
  const sorted = [...trades].sort((a, b) => a.date.localeCompare(b.date));
  const byDate = {};
  for (const t of sorted) { if (!byDate[t.date]) byDate[t.date] = 0; byDate[t.date] += t.r; }
  const dailyR = Object.values(byDate);
  const n = dailyR.length;
  if (n < 2) return { sharpe: null, sortino: null, calmar: null, recovery: null };
  const mean = dailyR.reduce((s, r) => s + r, 0) / n;
  const variance = dailyR.reduce((s, r) => s + (r - mean) ** 2, 0) / (n - 1);
  const stddev = Math.sqrt(variance);
  const downside = Math.sqrt(dailyR.filter(r => r < 0).reduce((s, r) => s + r * r, 0) / n);
  const sqrtN = Math.sqrt(252);
  const sharpe = stddev > 0 ? +(mean / stddev * sqrtN).toFixed(2) : null;
  const sortino = downside > 0 ? +(mean / downside * sqrtN).toFixed(2) : null;
  let cumR = 0, peak = 0, maxDD = 0, totalR = 0;
  for (const t of sorted) {
    cumR += t.r; totalR += t.r;
    if (cumR > peak) peak = cumR;
    if (peak - cumR > maxDD) maxDD = peak - cumR;
  }
  const firstDate = sorted[0].date, lastDate = sorted[sorted.length - 1].date;
  const daysDiff = Math.max(1, (new Date(lastDate + 'T00:00:00') - new Date(firstDate + 'T00:00:00')) / 86400000);
  const annR = totalR / daysDiff * 252;
  const calmar = maxDD > 0 ? +(annR / maxDD).toFixed(2) : null;
  const recovery = maxDD > 0 ? +(totalR / maxDD).toFixed(2) : null;
  return { sharpe, sortino, calmar, recovery };
}

/* ── Excel parser ─────────────────────────────────────────────────────────── */
// Column layout (0-indexed) in the "Overall" sheet:
//   1=PAIR  3=Date  5=Profit R  13=Closing Reason  14=ENTRY (TV link)
function parseXLSXFile(buffer) {
  const wb = XLSX.read(buffer, { type: 'array' });
  // Try "Overall" first, then fall back to first sheet
  const sheetName = wb.SheetNames.includes('Overall') ? 'Overall' : wb.SheetNames[0];
  if (!sheetName) throw new Error('Excel file has no sheets.');
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
  const trades = [];
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[1]) continue;
    const pair = row[1];
    const rawDate = row[3];
    const rawR = row[5];
    const closingReason = (row[13] || '').toString().trim().toUpperCase();
    const entry = row[14] || '';
    if (!pair || rawDate === undefined || rawDate === null || rawDate === '') continue;
    let date = '';
    if (typeof rawDate === 'number') {
      const d = XLSX.SSF.parse_date_code(rawDate);
      if (d) date = `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
    } else {
      const d = new Date(rawDate);
      if (!isNaN(d)) date = d.toISOString().slice(0, 10);
    }
    if (!date) continue;
    const r = closingReason === 'BE' ? 0 : +(typeof rawR === 'number' ? rawR : parseFloat(String(rawR)) || 0).toFixed(2);
    trades.push({
      date,
      symbol: String(pair).trim().toUpperCase().replace(/\s+/g, ''),
      r,
      type: null,
      notes: closingReason,
      chartLink: entry ? String(entry).trim() : '',
    });
  }
  return trades;
}

/* ── THEME ─────────────────────────────────────────────────────────────────── */
const T = dark => dark ? {
  bg: '#060c18', card: '#0b1422', cardHi: '#0e1829',
  border: '#162034', borderFaint: 'rgba(22,32,52,0.6)',
  text: '#dde3f0', textDim: '#8a9bbf', sub: '#4d6385',
  accent: '#00c9a7', gain: '#00c9a7',
  gainBg: 'rgba(0,201,167,0.08)', gainBorder: 'rgba(0,201,167,0.2)',
  loss: '#f5486a', lossBg: 'rgba(245,72,106,0.08)', lossBorder: 'rgba(245,72,106,0.2)',
  muted: '#070e1b', nav: '#050b16', navBorder: 'rgba(255,255,255,0.04)',
  navText: 'rgba(255,255,255,0.38)', navActive: '#fff', navActiveBg: 'rgba(0,201,167,0.08)',
  purple: '#7c6bff', purpleBg: 'rgba(124,107,255,0.08)', amber: '#f59e0b', sky: '#38bdf8',
} : {
  bg: '#eef2f9', card: '#ffffff', cardHi: '#f8faff',
  border: '#dce4f0', borderFaint: 'rgba(220,228,240,0.6)',
  text: '#0f172a', textDim: '#4b6080', sub: '#7e95b5',
  accent: '#0284c7', gain: '#059669',
  gainBg: 'rgba(5,150,105,0.07)', gainBorder: 'rgba(5,150,105,0.18)',
  loss: '#dc2626', lossBg: 'rgba(220,38,38,0.07)', lossBorder: 'rgba(220,38,38,0.18)',
  muted: '#f1f5fb', nav: '#0f172a', navBorder: 'rgba(255,255,255,0.05)',
  navText: 'rgba(255,255,255,0.42)', navActive: '#fff', navActiveBg: 'rgba(255,255,255,0.06)',
  purple: '#6d5de8', purpleBg: 'rgba(109,93,232,0.08)', amber: '#d97706', sky: '#0284c7',
};
const YC = ['#7c6bff', '#00c9a7', '#f59e0b', '#38bdf8', '#f5486a'];

/* ── APP ─────────────────────────────────────────────────────────────────────── */
export default function App() {
  const [dark, setDark] = useState(false);
  const [page, setPage] = useState('dashboard');
  const [trades, setTrades] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [loading, setLoading] = useState(true);
  const [dbError, setDbError] = useState(null);
  const [selectedAccount, setSelectedAccount] = useState('all');
  const [breakEvenAsWin, setBreakEvenAsWin] = useState(true);
  const [period, setPeriod] = useState('ALL');
  const [logPanelOpen, setLogPanelOpen] = useState(false);
  const [saving, setSaving] = useState(false);
  const [calendarMonth, setCalendarMonth] = useState({ y: new Date().getFullYear(), m: new Date().getMonth() });
  const [showNewAccount, setShowNewAccount] = useState(false);
  const [tradeSearch, setTradeSearch] = useState('');
  const [tradeSort, setTradeSort] = useState('date-desc');
  const [equityMode, setEquityMode] = useState('all');
  const [visibility, setVisibility] = useState({
    equity: true, pnl: true, winRate: true, totalR: true, profitFactor: true,
    streaks: true, drawdown: true, tradeHistory: true, calendar: true, triangle: true,
  });
  const [newTrade, setNewTrade] = useState({
    date: new Date().toISOString().slice(0, 10), symbol: '', account: '',
    r: null, type: 'Long', notes: '', chartLink: '', riskPercent: null,
  });
  const [newAccount, setNewAccount] = useState({ name: '', type: 'Challenge', currency: 'USD', initialBalance: 10000, riskPercent: 1.0 });
  const [isAdmin, setIsAdmin] = useState(() => { try { return localStorage.getItem('sm_admin') === '1'; } catch { return false; } });
  const [editAcct, setEditAcct] = useState(null); // { id, name, type, currency, initialBalance, riskPercent }
  const [showAdminModal, setShowAdminModal] = useState(false);
  const [adminPassword, setAdminPassword] = useState('');
  const [adminError, setAdminError] = useState('');
  const [importFiles, setImportFiles] = useState([]);
  const [importAccount, setImportAccount] = useState('');
  const [importProgress, setImportProgress] = useState(null);
  const [importError, setImportError] = useState('');
  const [chartYear, setChartYear] = useState(new Date().getFullYear().toString());
  const [dashMonthFilter, setDashMonthFilter] = useState(null); // 'YYYY-MM' or null
  const [monthPopup, setMonthPopup] = useState(null); // 'YYYY-MM' key for popup modal
  const [cardYears, setCardYears] = useState({}); // per-card year filter { cardLabel: 'all'|'2025'|... }
  const [totalRMode, setTotalRMode] = useState('R'); // 'R' or '%'
  const [dailyMode, setDailyMode] = useState('daily'); // 'daily' or 'weekly'
  const [equityDispMode, setEquityDispMode] = useState('R'); // 'R' or '%'
  const [aboutImgError, setAboutImgError] = useState(false);
  const [isMobile, setIsMobile] = useState(() => typeof window !== 'undefined' && window.innerWidth < 768);

  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener('resize', handler);
    return () => window.removeEventListener('resize', handler);
  }, []);

  const theme = useMemo(() => T(dark), [dark] /* theme object — see T() above */);

  /* ── LOAD FROM SUPABASE ──────────────────────────────────────────────────── */
  useEffect(() => {
    async function load() {
      setLoading(true);
      setDbError(null);
      try {
        const [a, t] = await Promise.all([fetchAccounts(), fetchTrades()]);
        setAccounts(a);
        setTrades(t);
        if (a.length) setNewTrade(p => ({ ...p, account: a[0].id, riskPercent: a[0].riskPercent ?? null }));
      } catch (e) {
        setDbError(e.message);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, []);

  /* ── REALTIME SUBSCRIPTION ───────────────────────────────────────────────── */
  useEffect(() => {
    const channel = supabase
      .channel('realtime-trades')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'trades' }, payload => {
        const rawTrade = payload.new;
        setTrades(prev => {
          if (prev.find(x => x.id === rawTrade.id)) return prev;
          return [...prev, {
            id: rawTrade.id,
            date: rawTrade.date,
            symbol: rawTrade.symbol,
            account: rawTrade.account_id,
            r: Number(rawTrade.r),
            type: rawTrade.type,
            notes: rawTrade.notes || '',
            chartLink: rawTrade.chart_link || '',
          }];
        });
      })
      .on('postgres_changes', { event: 'DELETE', schema: 'public', table: 'trades' }, payload => {
        setTrades(prev => prev.filter(x => x.id !== payload.old.id));
      })
      .subscribe((status, error) => {
        if (status === 'SUBSCRIBED') {
          console.log('[Realtime] Connected to trades channel.');
        }
        if (status === 'CHANNEL_ERROR') {
          console.error('[Realtime] Channel error — check that Realtime is enabled for the trades table in Supabase:', error);
        }
        if (status === 'TIMED_OUT') {
          console.warn('[Realtime] Connection timed out. New trades will not appear automatically until you refresh.');
        }
        if (status === 'CLOSED') {
          console.info('[Realtime] Channel closed.');
        }
      });
    return () => supabase.removeChannel(channel);
  }, []);

  /* ── DERIVED STATE ───────────────────────────────────────────────────────── */
  const filtered = useMemo(() => {
    let t = trades;
    if (selectedAccount !== 'all') t = t.filter(x => x.account === selectedAccount);
    t = applyPeriod(t, period);
    return t.slice().sort((a, b) => a.date.localeCompare(b.date));
  }, [trades, selectedAccount, period]);

  const kpiYears = useMemo(() => [...new Set(filtered.map(t => t.date.slice(0, 4)))].sort().reverse(), [filtered]);
  const st = useMemo(() => computeStats(filtered, breakEvenAsWin), [filtered, breakEvenAsWin]);
  const mon = useMemo(() => buildMon(filtered), [filtered]);

  const compareYears = useMemo(() => [...new Set(filtered.map(t => t.date.slice(0, 4)))].sort(), [filtered]);

  const eqTrades = useMemo(() => {
    if (equityMode === 'compare' || equityMode === 'all') return filtered;
    return filtered.filter(x => x.date.startsWith(equityMode));
  }, [filtered, equityMode]);

  const eq = useMemo(() => {
    const a = accounts.find(x => x.id === selectedAccount);
    return buildEq(eqTrades, a?.riskPercent ?? 0.5, a?.initialBalance ?? 25000);
  }, [eqTrades, accounts, selectedAccount]);

  const compareData = useMemo(() => {
    if (equityMode !== 'compare') return null;
    const byY = {};
    for (const y of compareYears) byY[y] = filtered.filter(t => t.date.startsWith(y)).slice().sort((a, b) => a.date.localeCompare(b.date));
    const maxLen = Math.max(...compareYears.map(y => byY[y].length), 0);
    const rows = [];
    for (let i = 0; i <= maxLen; i++) {
      const row = { trade: i };
      for (const y of compareYears) if (i <= byY[y].length) row[y] = +byY[y].slice(0, i).reduce((s, t) => s + t.r, 0).toFixed(2);
      rows.push(row);
    }
    return rows;
  }, [filtered, equityMode, compareYears]);

  const calData = useMemo(() => {
    const m = {};
    for (const t of filtered) { if (!m[t.date]) m[t.date] = { r: 0, n: 0 }; m[t.date].r += t.r; m[t.date].n++; }
    return m;
  }, [filtered]);

  const calCells = useMemo(() => {
    const { y, m } = calendarMonth, first = new Date(y, m, 1).getDay(), days = new Date(y, m + 1, 0).getDate();
    const c = [];
    for (let i = 0; i < first; i++) c.push(null);
    for (let d = 1; d <= days; d++) {
      const ds = `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      c.push({ d, ds, info: calData[ds] || null });
    }
    return c;
  }, [calendarMonth, calData]);

  const dispTrades = useMemo(() => {
    let t = filtered.filter(x => !tradeSearch || x.symbol.toLowerCase().includes(tradeSearch.toLowerCase()));
    if (tradeSort === 'date-desc') t = [...t].reverse();
    else if (tradeSort === 'r-desc') t = [...t].sort((a, b) => b.r - a.r);
    else if (tradeSort === 'r-asc') t = [...t].sort((a, b) => a.r - b.r);
    return t;
  }, [filtered, tradeSearch, tradeSort]);

  const recentTrades = useMemo(() => [...filtered].reverse(), [filtered]);
  const donut = useMemo(() => {
    const be = filtered.filter(t => t.r === 0).length;
    return [{ name: 'Wins', val: st.wins, fill: theme.gain }, { name: 'Losses', val: st.losses, fill: theme.loss }, ...(be ? [{ name: 'BE', val: be, fill: theme.sub }] : [])].filter(d => d.val > 0);
  }, [st, filtered, theme]);

  const tri = useMemo(() => {
    const nWR = Math.min(100, Math.round(st.wr / 70 * 100));
    const nWL = Math.min(100, Math.round(st.wins / Math.max(st.losses, 1) / 2 * 100));
    const nPF = Math.min(100, Math.round(Math.min(st.pf, 2.5) / 2.5 * 100));
    return { data: [{ axis: 'Win %', v: nWR }, { axis: 'W/L Ratio', v: nWL }, { axis: 'Profit Factor', v: nPF }], nWR, nWL, nPF, score: Math.round((nWR + nWL + nPF) / 3) };
  }, [st]);

  /* ── SHARED STYLES ──────────────────────────────────────────────────────── */
  const card = {
    background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 14, padding: '20px 22px',
    boxShadow: `0 1px 3px rgba(0,0,0,${dark ? 0.4 : 0.06}), inset 0 1px 0 rgba(255,255,255,${dark ? 0.03 : 0.7})`,
  };
  const sectionLabel = {
    fontSize: 10, fontWeight: 700, letterSpacing: 2, textTransform: 'uppercase',
    color: theme.sub, marginBottom: 16, display: 'flex', alignItems: 'center', gap: 8,
  };
  const inp = { padding: '9px 12px', borderRadius: 9, border: `1px solid ${theme.border}`, background: theme.muted, color: theme.text, fontFamily: "'Outfit',sans-serif", fontSize: 13.5, outline: 'none', width: '100%', boxSizing: 'border-box', transition: 'border-color .15s' };
  const btnP = { padding: '8px 18px', borderRadius: 9, fontSize: 13, fontWeight: 600, cursor: 'pointer', border: 'none', background: theme.accent, color: '#fff', fontFamily: "'Outfit',sans-serif", letterSpacing: 0.3 };
  const btnG = { padding: '7px 14px', borderRadius: 9, fontSize: 12.5, fontWeight: 500, cursor: 'pointer', background: 'transparent', border: `1px solid ${theme.border}`, color: theme.textDim, fontFamily: "'Outfit',sans-serif" };
  const lbl = { fontSize: 10, fontWeight: 700, letterSpacing: 1.5, textTransform: 'uppercase', color: theme.sub, marginBottom: 6, display: 'block' };

  /* ── SMALL COMPONENTS ───────────────────────────────────────────────────── */

  const SectionLabel = ({ children, accent = false }) => (
    <div style={sectionLabel}>
      {accent && <span style={{ display: 'inline-block', width: 3, height: 12, borderRadius: 2, background: theme.accent, flexShrink: 0 }} />}
      {children}
    </div>
  );

  const DirBadge = ({ type }) => {
    if (!type) return <span style={{ fontSize: 10, color: theme.sub }}>—</span>;
    return (
      <span style={{
        display: 'inline-flex', alignItems: 'center', gap: 3,
        fontSize: 10, padding: '2px 8px', borderRadius: 5, fontWeight: 700, letterSpacing: 0.5,
        background: type === 'Long' ? theme.gainBg : theme.lossBg,
        color: type === 'Long' ? theme.gain : theme.loss,
        border: `1px solid ${type === 'Long' ? theme.gainBorder : theme.lossBorder}`,
      }}>
        {type === 'Long' ? '▲' : '▼'} {type}
      </span>
    );
  };

  const AcctBadge = ({ type }) => {
    const s = { Live: { bg: theme.gainBg, c: theme.gain, border: theme.gainBorder }, Funded: { bg: theme.purpleBg, c: theme.purple, border: 'rgba(124,107,255,0.2)' }, Challenge: { bg: 'rgba(245,158,11,0.08)', c: theme.amber, border: 'rgba(245,158,11,0.2)' } }[type] || {};
    return <span style={{ fontSize: 9.5, padding: '2px 8px', borderRadius: 20, background: s.bg, color: s.c, fontWeight: 700, letterSpacing: 1, textTransform: 'uppercase', border: `1px solid ${s.border}` }}>{type}</span>;
  };


  const Tip = ({ active, payload, label, suffix = '' }) => {
    if (!active || !payload?.length) return null;
    const v = payload[0].value;
    return (
      <div style={{ background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 8, padding: '8px 12px' }}>
        <div style={{ fontSize: 10, color: theme.sub, marginBottom: 3 }}>{label}</div>
        <div style={{ fontSize: 13, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: v >= 0 ? theme.gain : theme.loss }}>
          {v >= 0 ? '+' : ''}{typeof v === 'number' ? v.toFixed(2) : v}{suffix}
        </div>
      </div>
    );
  };

  const MonBarTip = ({ active, payload }) => {
    if (!active || !payload?.length) return null;
    const d = payload[0].payload;
    const r = d.r;
    const dollars = r * riskDollar;
    return (
      <div style={{ background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 9, padding: '10px 14px', boxShadow: '0 4px 16px rgba(0,0,0,0.18)' }}>
        <div style={{ fontSize: 11, color: theme.sub, marginBottom: 5 }}>{new Date(d.key + '-01').toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}</div>
        <div style={{ fontSize: 15, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: r >= 0 ? theme.gain : theme.loss, marginBottom: 3 }}>
          {r >= 0 ? '+' : ''}{r.toFixed(2)}R
        </div>
        <div style={{ fontSize: 13, fontWeight: 600, fontFamily: "'JetBrains Mono',monospace", color: r >= 0 ? theme.gain : theme.loss, opacity: 0.8 }}>
          {fmtU(dollars)}
        </div>
      </div>
    );
  };

  const MultiTip = ({ active, payload, label }) => {
    if (!active || !payload?.length) return null;
    return (
      <div style={{ background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 8, padding: '10px 14px' }}>
        <div style={{ fontSize: 10, color: theme.sub, marginBottom: 6 }}>Trade #{label}</div>
        {payload.filter(p => p.value != null).map((p, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 3 }}>
            <div style={{ width: 18, height: 3, borderRadius: 2, background: p.color }} />
            <span style={{ fontSize: 11, color: theme.sub }}>{p.dataKey}</span>
            <span style={{ fontSize: 12, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: p.value >= 0 ? theme.gain : theme.loss, marginLeft: 'auto', paddingLeft: 12 }}>
              {p.value >= 0 ? '+' : ''}{p.value?.toFixed(2)}R
            </span>
          </div>
        ))}
      </div>
    );
  };

  const RRTip = ({ active, payload, label }) => {
    if (!active || !payload?.length) return null;
    const val = payload[0].value;
    const isPct = equityDispMode === '%';
    const dollars = isPct ? null : val * riskDollar;
    return (
      <div style={{ background: dark ? '#0b1628' : theme.card, border: `1px solid ${theme.border}`, borderRadius: 8, padding: '9px 13px', boxShadow: '0 4px 16px rgba(0,0,0,0.3)' }}>
        <div style={{ fontSize: 9.5, letterSpacing: 1.2, textTransform: 'uppercase', color: theme.sub, marginBottom: 5 }}>
          {label === 0 ? 'Start' : `Trade #${label}`}
        </div>
        <div style={{ fontSize: 15, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: val >= 0 ? theme.gain : theme.loss }}>
          {val >= 0 ? '+' : ''}{typeof val === 'number' ? val.toFixed(2) : val}{isPct ? '%' : 'R'}
        </div>
        {!isPct && visibility.equity && riskDollar > 0 && (
          <div style={{ fontSize: 11.5, fontFamily: "'JetBrains Mono',monospace", color: theme.sub, marginTop: 3 }}>
            {dollars >= 0 ? '+$' : '-$'}{Math.abs(dollars).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
          </div>
        )}
      </div>
    );
  };

  const EqTip = ({ active, payload, label }) => {
    if (!active || !payload?.length) return null;
    const v = payload[0].value;
    const init = eq.length ? eq[0].bal : (selAcctData?.initialBalance || 25000);
    const diff = v - init;
    return (
      <div style={{ background: dark ? '#0b1628' : theme.card, border: `1px solid ${theme.border}`, borderRadius: 8, padding: '9px 13px', boxShadow: '0 4px 16px rgba(0,0,0,0.3)' }}>
        <div style={{ fontSize: 9.5, letterSpacing: 1.2, textTransform: 'uppercase', color: theme.sub, marginBottom: 5 }}>
          {label === 0 ? 'Start' : `Trade #${label}`}
        </div>
        <div style={{ fontSize: 15, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: diff >= 0 ? theme.gain : theme.loss }}>
          ${v.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
        </div>
        <div style={{ fontSize: 11.5, fontFamily: "'JetBrains Mono',monospace", color: diff >= 0 ? theme.gain : theme.loss, marginTop: 3 }}>
          {diff >= 0 ? '+$' : '-$'}{Math.abs(diff).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
        </div>
      </div>
    );
  };

  const Tog = ({ on, fn }) => (
    <button onClick={fn} style={{ width: 34, height: 19, borderRadius: 10, cursor: 'pointer', background: on ? theme.accent : theme.border, border: 'none', position: 'relative', transition: 'background .2s', flexShrink: 0 }}>
      <div style={{ position: 'absolute', width: 13, height: 13, borderRadius: '50%', background: '#fff', top: 3, left: on ? 18 : 3, transition: 'left .2s' }} />
    </button>
  );

  const YrBtn = ({ label, active, onClick }) => (
    <button onClick={onClick} style={{ padding: isMobile ? '3px 7px' : '4px 11px', borderRadius: 6, fontSize: isMobile ? 10 : 11.5, fontWeight: 600, cursor: 'pointer', color: active ? '#fff' : theme.sub, background: active ? theme.accent : 'transparent', border: 'none', fontFamily: "'Outfit',sans-serif", transition: 'all .12s' }}>
      {label}
    </button>
  );

  const YrSelector = () => (
    <div style={{ display: 'flex', gap: 2, background: theme.muted, border: `1px solid ${theme.border}`, borderRadius: 7, padding: 2 }}>
      {[...compareYears, 'all', 'compare'].map(y => (
        <YrBtn key={y} label={y === 'all' ? 'All' : y === 'compare' ? 'Compare ✦' : y} active={equityMode === y} onClick={() => setEquityMode(y)} />
      ))}
    </div>
  );

  /* ── CRUD OPERATIONS ────────────────────────────────────────────────────── */
  const addTrade = useCallback(async () => {
    if (!newTrade.symbol || newTrade.r === null) return;
    setSaving(true);
    try {
      const saved = await insertTrade(newTrade);
      setTrades(p => [...p, saved]);
      setLogPanelOpen(false);
      setNewTrade(p => ({ ...p, symbol: '', r: null, notes: '', chartLink: '', riskPercent: null }));
    } catch (e) {
      alert('Error saving trade: ' + e.message);
    } finally {
      setSaving(false);
    }
  }, [newTrade]);

  const removeTrade = useCallback(async (id) => {
    try {
      await deleteTrade(id);
      setTrades(p => p.filter(x => x.id !== id));
    } catch (e) {
      alert('Error deleting trade: ' + e.message);
    }
  }, []);

  const addAccount = useCallback(async () => {
    if (!newAccount.name) return;
    try {
      const saved = await insertAccount(newAccount);
      setAccounts(p => [...p, saved]);
      setNewTrade(p => ({ ...p, account: saved.id }));
      setShowNewAccount(false);
      setNewAccount({ name: '', type: 'Challenge', currency: 'USD', initialBalance: 10000, riskPercent: 1.0 });
    } catch (e) {
      alert('Error adding account: ' + e.message);
    }
  }, [newAccount]);

  const promoteAccount = useCallback(async (id, newType) => {
    try {
      await updateAccountType(id, newType);
      setAccounts(p => p.map(x => x.id === id ? { ...x, type: newType } : x));
    } catch (e) {
      alert('Error updating account: ' + e.message);
    }
  }, []);

  const removeAccount = useCallback(async (id, name) => {
    if (!window.confirm(`Remove "${name}"? All trades on this account will be unlinked.`)) return;
    try {
      await deleteAccount(id);
      setAccounts(p => p.filter(x => x.id !== id));
    } catch (e) {
      alert('Error removing account: ' + e.message);
    }
  }, []);

  const saveAccountEdit = useCallback(async () => {
    if (!editAcct) return;
    try {
      await updateAccount(editAcct.id, editAcct);
      setAccounts(p => p.map(x => x.id === editAcct.id ? { ...x, ...editAcct } : x));
      setEditAcct(null);
    } catch (e) {
      alert('Error updating account: ' + e.message);
    }
  }, [editAcct]);

  const tryAdmin = useCallback(() => {
    const correct = import.meta.env.VITE_ADMIN_PASSWORD;
    if (!correct || adminPassword === correct) {
      setIsAdmin(true);
      try { localStorage.setItem('sm_admin', '1'); } catch { }
      setShowAdminModal(false);
      setAdminPassword('');
      setAdminError('');
    } else {
      setAdminError('Incorrect password. Try again.');
    }
  }, [adminPassword]);

  const handleImportFiles = useCallback(async (files) => {
    setImportError('');
    setImportProgress('parsing');
    try {
      const parsed = [];
      for (const file of Array.from(files)) {
        const buf = await file.arrayBuffer();
        const trades = parseXLSXFile(buf);
        parsed.push({ name: file.name, trades });
      }
      setImportFiles(parsed);
      if (parsed.length && accounts.length) setImportAccount(accounts[0].id);
    } catch (e) {
      setImportError(e.message);
    } finally {
      setImportProgress(null);
    }
  }, [accounts]);

  const doImport = useCallback(async () => {
    if (!importAccount || importFiles.length === 0) return;
    const allTrades = importFiles.flatMap(f => f.trades);
    setImportProgress('importing');
    setImportError('');
    try {
      const inserted = await batchInsertTrades(allTrades, importAccount);
      setTrades(p => [...p, ...inserted]);
      setImportFiles([]);
      setImportProgress('done');
    } catch (e) {
      setImportError(e.message);
      setImportProgress(null);
    }
  }, [importAccount, importFiles]);

  /* ── LOADING / ERROR SCREENS ────────────────────────────────────────────── */
  if (loading) return (
    <div style={{ position: 'fixed', inset: 0, background: 'radial-gradient(ellipse at 35% 25%, #0e2040 0%, #060d1c 45%, #020509 100%)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', fontFamily: "'Outfit',sans-serif", overflow: 'hidden' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700;800&display=swap');
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes ldFadeUp { from { opacity: 0; transform: translateY(18px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes ldScale { from { opacity: 0; transform: scale(0.82); } to { opacity: 1; transform: scale(1); } }
        @keyframes ldGlow { 0%,100%{box-shadow:0 0 50px rgba(0,201,167,.12),0 20px 60px rgba(0,0,0,.7),0 0 0 1px rgba(255,255,255,.06)} 50%{box-shadow:0 0 90px rgba(0,201,167,.28),0 20px 60px rgba(0,0,0,.7),0 0 0 1px rgba(255,255,255,.1)} }
        @keyframes ldShimmer { 0%{transform:translateX(-120%)} 100%{transform:translateX(220%)} }
        @keyframes ldPulse { 0%,100%{opacity:.35;transform:scale(1)} 50%{opacity:.7;transform:scale(1.04)} }
        @keyframes ldDot { 0%,80%,100%{transform:scale(0);opacity:0} 40%{transform:scale(1);opacity:1} }
        @keyframes ldBar { from{width:0} to{width:72%} }
      `}</style>

      {/* Ambient glow orb */}
      <div style={{ position: 'absolute', width: 700, height: 700, borderRadius: '50%', background: 'radial-gradient(circle, rgba(0,201,167,0.05) 0%, transparent 65%)', top: '50%', left: '50%', transform: 'translate(-50%,-50%)', animation: 'ldPulse 3.5s ease-in-out infinite', pointerEvents: 'none' }} />
      {/* Logo */}
      <div style={{ animation: 'ldScale 0.65s cubic-bezier(0.34,1.56,0.64,1) both', marginBottom: 30 }}>
        <img
          src="/sma-logo.png"
          alt="SMA Logo"
          style={{
            width: 120,
            height: 120,
            borderRadius: 24,
            display: 'block',
            animation: 'ldGlow 3s ease-in-out infinite',
          }}
        />
      </div>

      {/* Brand text */}
      <div style={{ animation: 'ldFadeUp 0.5s 0.2s both', textAlign: 'center', marginBottom: 6 }}>
        <div style={{ fontSize: 9, fontWeight: 700, letterSpacing: 4, textTransform: 'uppercase', color: '#00c9a7', marginBottom: 8 }}>Simple Markets</div>
        <div style={{ fontSize: 20, fontWeight: 700, color: '#dde3f0', letterSpacing: -0.5, lineHeight: 1.25 }}>Andrej's<br />Performance Record</div>
      </div>

      {/* Loading indicator */}
      <div style={{ animation: 'ldFadeUp 0.5s 0.4s both', marginTop: 32, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10 }}>
        <div style={{ width: 180, height: 2, background: 'rgba(255,255,255,0.06)', borderRadius: 2, overflow: 'hidden' }}>
          <div style={{ height: '100%', background: 'linear-gradient(90deg, #00c9a7, #38bdf8)', borderRadius: 2, animation: 'ldBar 2.2s cubic-bezier(0.4,0,0.2,1) forwards' }} />
        </div>
        <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.25)', letterSpacing: 1.5, textTransform: 'uppercase' }}>Loading data…</div>
      </div>
    </div>
  );

  if (dbError) return (
    <div style={{ minHeight: '100vh', background: '#07101f', display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: 12, fontFamily: "'Outfit',sans-serif", padding: 24 }}>
      <div style={{ fontSize: 32 }}>⚠️</div>
      <div style={{ color: '#f5486a', fontSize: 16, fontWeight: 700 }}>Database connection failed</div>
      <div style={{ color: '#5d7099', fontSize: 13, maxWidth: 400, textAlign: 'center' }}>{dbError}</div>
      <div style={{ color: '#5d7099', fontSize: 12 }}>Check your .env.local file and Supabase project status.</div>
    </div>
  );

  const selAcctData = accounts.find(x => x.id === selectedAccount) || accounts[0] || null;
  const riskDollar = selAcctData ? (selAcctData.riskPercent / 100) * selAcctData.initialBalance : 250;
  const ratios = computeRatios(filtered, selAcctData?.riskPercent ?? 1.0, selAcctData?.initialBalance ?? 25000);

  const RPILLS = [-2, -1.5, -1, -0.5, 0, 0.5, 1, 1.5, 2, 2.5, 3, 4, 5];
  const ALL_PAGES = [
    { id: 'dashboard', l: 'Dashboard', i: '▦' },
    { id: 'about', l: 'About & Strategy', i: '◈', public: true },
    { id: 'trades', l: 'Trade Log', i: '≡' },
    { id: 'calendar', l: 'Calendar', i: '◫' },
    { id: 'accounts', l: 'Accounts', i: '◎' },
    { id: 'settings', l: 'Settings', i: '⚙' },
  ];
  const PAGES = isAdmin ? ALL_PAGES : ALL_PAGES.filter(p => p.id === 'dashboard' || p.public);

  /* ── ADMIN MODAL ────────────────────────────────────────────────────────── */
  const AdminModal = () => (
    <>
      <div onClick={() => { setShowAdminModal(false); setAdminPassword(''); setAdminError(''); }} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.6)', zIndex: 200, backdropFilter: 'blur(4px)' }} />
      <div style={{ position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%,-50%)', width: 340, background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 16, padding: 28, zIndex: 201, display: 'flex', flexDirection: 'column', gap: 14, fontFamily: "'Outfit',sans-serif" }}>
        <div style={{ fontSize: 17, fontWeight: 700 }}>Admin Login</div>
        <div style={{ fontSize: 12, color: theme.sub }}>Enter your admin password to unlock editing capabilities and all pages.</div>
        <div>
          <label style={lbl}>Password</label>
          <input type="password" style={inp} value={adminPassword}
            onChange={e => { setAdminPassword(e.target.value); setAdminError(''); }}
            onKeyDown={e => e.key === 'Enter' && tryAdmin()}
            placeholder="Admin password…" autoFocus />
        </div>
        {adminError && <div style={{ fontSize: 12, color: theme.loss }}>{adminError}</div>}
        <div style={{ display: 'flex', gap: 8 }}>
          <button style={{ ...btnG, flex: 1 }} onClick={() => { setShowAdminModal(false); setAdminPassword(''); setAdminError(''); }}>Cancel</button>
          <button style={{ ...btnP, flex: 1 }} onClick={tryAdmin}>Unlock</button>
        </div>
      </div>
    </>
  );

  /* ── LOG PANEL ──────────────────────────────────────────────────────────── */
  const LogPanel = () => (
    <>
      <div onClick={() => setLogPanelOpen(false)} style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', zIndex: 100, backdropFilter: 'blur(3px)' }} />
      <div style={{ position: 'fixed', right: 0, top: 0, bottom: 0, width: isMobile ? '100%' : 400, background: theme.card, borderLeft: isMobile ? 'none' : `1px solid ${theme.border}`, zIndex: 101, padding: isMobile ? '20px 20px 40px' : 26, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ fontSize: 17, fontWeight: 700 }}>Log Trade</div>
          <button style={{ ...btnG, padding: '3px 9px' }} onClick={() => setLogPanelOpen(false)}>✕</button>
        </div>
        <div><label style={lbl}>Date</label><input type="date" style={inp} value={newTrade.date} onChange={e => setNewTrade(p => ({ ...p, date: e.target.value }))} /></div>
        <div><label style={lbl}>Symbol</label><input style={inp} value={newTrade.symbol} placeholder="EURUSD, XAUUSD…" onChange={e => setNewTrade(p => ({ ...p, symbol: e.target.value.toUpperCase() }))} /></div>
        <div>
          <label style={lbl}>Account</label>
          {accounts.length === 0
            ? <div style={{ fontSize: 12, color: theme.loss, padding: '8px 0' }}>No accounts yet — add one in the Accounts tab first.</div>
            : <select style={inp} value={newTrade.account} onChange={e => {
              const acct = accounts.find(a => a.id === e.target.value);
              setNewTrade(p => ({ ...p, account: e.target.value, riskPercent: acct?.riskPercent ?? null }));
            }}>
              {accounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          }
        </div>
        <div>
          <label style={lbl}>Direction</label>
          <div style={{ display: 'flex', gap: 8 }}>
            {['Long', 'Short'].map(dir => (
              <button key={dir} onClick={() => setNewTrade(p => ({ ...p, type: dir }))} style={{ flex: 1, padding: '8px', borderRadius: 8, border: `1.5px solid ${newTrade.type === dir ? theme.accent : theme.border}`, background: newTrade.type === dir ? `${theme.accent}18` : theme.muted, color: newTrade.type === dir ? theme.accent : theme.sub, fontFamily: "'Outfit',sans-serif", fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>
                {dir === 'Long' ? '▲ Long' : '▼ Short'}
              </button>
            ))}
          </div>
        </div>
        <div>
          <label style={lbl}>R Multiple</label>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 5, marginBottom: 8 }}>
            {RPILLS.map(r => {
              const sel = newTrade.r === r, col = r > 0 ? theme.gain : r < 0 ? theme.loss : theme.sub;
              return <button key={r} onClick={() => setNewTrade(p => ({ ...p, r }))} style={{ padding: '5px 9px', borderRadius: 6, border: `1.5px solid ${sel ? col : theme.border}`, background: sel ? `${col}18` : theme.muted, color: sel ? col : theme.sub, fontFamily: "'JetBrains Mono',monospace", fontSize: 11.5, fontWeight: 600, cursor: 'pointer' }}>{r > 0 ? '+' : ''}{r}R</button>;
            })}
          </div>
          <input type="number" step="0.1" style={inp} value={newTrade.r ?? ''} placeholder="Custom R…" onChange={e => setNewTrade(p => ({ ...p, r: parseFloat(e.target.value) || null }))} />
        </div>
        <div>
          <label style={lbl}>Risk % per trade</label>
          <input type="number" step="0.1" style={inp} value={newTrade.riskPercent ?? ''}
            placeholder={`Default: ${accounts.find(a => a.id === newTrade.account)?.riskPercent ?? 1}%`}
            onChange={e => setNewTrade(p => ({ ...p, riskPercent: e.target.value ? parseFloat(e.target.value) : null }))} />
          {newTrade.r !== null && (() => {
            const acct = accounts.find(a => a.id === newTrade.account);
            const rp = newTrade.riskPercent != null ? newTrade.riskPercent : (acct?.riskPercent ?? 1.0);
            const init = acct?.initialBalance ?? 25000;
            const pnl = newTrade.r * (rp / 100) * init;
            return (
              <div style={{ marginTop: 7, padding: '8px 12px', borderRadius: 8, background: pnl >= 0 ? theme.gainBg : theme.lossBg, border: `1px solid ${pnl >= 0 ? theme.gainBorder : theme.lossBorder}` }}>
                <span style={{ fontSize: 11, color: theme.sub, letterSpacing: 0.3 }}>P&L Preview: </span>
                <span style={{ fontSize: 13, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: pnl >= 0 ? theme.gain : theme.loss }}>
                  {pnl >= 0 ? '+$' : '-$'}{Math.abs(pnl).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </span>
                <span style={{ fontSize: 10, color: theme.sub, marginLeft: 6 }}>({rp}% of ${init.toLocaleString()})</span>
              </div>
            );
          })()}
        </div>
        <div><label style={lbl}>TradingView link</label><input style={inp} value={newTrade.chartLink} placeholder="https://tradingview.com/…" onChange={e => setNewTrade(p => ({ ...p, chartLink: e.target.value }))} /></div>
        <div><label style={lbl}>Notes</label><textarea style={{ ...inp, resize: 'vertical' }} rows={3} value={newTrade.notes} placeholder="Setup, emotions, mistakes…" onChange={e => setNewTrade(p => ({ ...p, notes: e.target.value }))} /></div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button style={{ ...btnG, flex: 1 }} onClick={() => setLogPanelOpen(false)}>Cancel</button>
          <button style={{ ...btnP, flex: 1, opacity: (!newTrade.symbol || newTrade.r === null || saving) ? .5 : 1 }} onClick={addTrade} disabled={saving}>
            {saving ? 'Saving…' : 'Save Trade'}
          </button>
        </div>
      </div>
    </>
  );

  /* ── ABOUT ──────────────────────────────────────────────────────────────── */
  const About = () => {
    const imgError = aboutImgError;
    const setImgError = setAboutImgError;
    const mechRows = [
      { label: 'Timeframe', value: '4-Hour (4H) Swing', sub: 'Swing positions held 1–5 days' },
      { label: 'Assets', value: 'FX Majors & Minors', sub: 'Liquid pairs only' },
      { label: 'Setup', value: 'Break of Structure (BoS)', sub: 'Market structure shifts only' },
      { label: 'Entry', value: '0.75 Fibonacci Retracement', sub: 'Single entry, no averaging' },
      { label: 'Risk : Reward', value: '1 : 3 — every single trade', sub: 'Fixed, never adjusted mid-trade' },
    ];
    const pillars = [
      { title: 'No interpretation needed', body: "The setup is binary. Either structure breaks and price returns to 0.75 — or it doesn't. There's no 'kind of' valid. That removes most of the mistakes traders make." },
      { title: 'Built around your life', body: 'Checking the 4H chart twice a day is enough. You\'re not watching ticks or setting alarms. The market does its work while you do yours.' },
      { title: 'The math is on your side', body: 'At 1:3 RR you only need to be right 34% of the time to break even. Most traders lose because of bad risk-reward, not bad entries. We fixed that first.' },
    ];
    const divider = { borderTop: `1px solid ${theme.border}`, margin: '44px 0' };
    return (
      <div style={{ maxWidth: 860, margin: '0 auto', fontFamily: "'Outfit',sans-serif", color: theme.text, paddingBottom: 60 }}>

        {/* ── HERO ── */}
        <div className="about-section" style={{ paddingBottom: 44, borderBottom: `1px solid ${theme.border}`, animationDelay: '0ms' }}>
          <div style={{ display: 'flex', alignItems: isMobile ? 'center' : 'flex-start', flexDirection: isMobile ? 'column' : 'row', gap: isMobile ? 20 : 32, flexWrap: 'wrap', textAlign: isMobile ? 'center' : 'left' }}>
            {/* Circular photo */}
            <div style={{ flexShrink: 0, paddingTop: 4 }}>
              {!imgError ? (
                <img
                  src="/andrej.jpg"
                  alt="Andrej"
                  onError={() => setImgError(true)}
                  style={{ width: 108, height: 108, borderRadius: '50%', objectFit: 'cover', objectPosition: 'center top', border: `3px solid ${theme.border}`, display: 'block' }}
                />
              ) : (
                <div style={{ width: 108, height: 108, borderRadius: '50%', background: dark ? '#0e1829' : '#dce4f0', border: `3px solid ${theme.border}`, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 36, fontWeight: 800, color: theme.accent }}>
                  A
                </div>
              )}
            </div>
            {/* Name + bio */}
            <div style={{ flex: 1, minWidth: isMobile ? 0 : 240 }}>
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 3, textTransform: 'uppercase', color: theme.accent, marginBottom: 10, fontFamily: "'JetBrains Mono',monospace" }}>
                Simple Markets Academy · Head Coach
              </div>
              <h1 style={{ margin: '0 0 14px', fontSize: 'clamp(26px,4vw,42px)', fontWeight: 800, lineHeight: 1.1, letterSpacing: -1, color: theme.text }}>
                Coach Andrej
              </h1>
              <div style={{ width: 40, height: 3, borderRadius: 2, background: theme.accent, marginBottom: 18 }} />
              <p style={{ margin: '0 0 12px', fontSize: 15, lineHeight: 1.85, color: theme.textDim }}>
                I coach at Simple Markets alongside two other head coaches. Before that I spent a long time doing what most traders do — adding more indicators, switching timeframes, looking for the edge in complexity. It wasn't there.
              </p>
              <p style={{ margin: 0, fontSize: 15, lineHeight: 1.85, color: theme.textDim }}>
                What I trade now is the opposite. Every trade on this page was taken with one ruleset, no exceptions. This is what the system actually produces in real market conditions.
              </p>
            </div>
          </div>
        </div>

        {/* ── STRATEGY MECHANICS ── */}
        <div className="about-section" style={{ paddingTop: 44, animationDelay: '80ms' }}>
          <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 3, textTransform: 'uppercase', color: theme.sub, marginBottom: 8, fontFamily: "'JetBrains Mono',monospace" }}>The Setup</div>
          <h2 style={{ margin: '0 0 10px', fontSize: 26, fontWeight: 800, letterSpacing: -0.5, color: theme.text }}>How Every Trade Is Found</h2>
          <p style={{ margin: '0 0 28px', fontSize: 14, color: theme.textDim, lineHeight: 1.75, maxWidth: 560 }}>
            Five rules. Every trade passes all five or it doesn't get taken. There's no discretion built in — it's a checklist.
          </p>
          <div style={{ border: `1px solid ${theme.border}`, borderRadius: 16, overflow: 'hidden', background: theme.card }}>
            {mechRows.map((r, i) => (
              <div key={r.label} style={{
                display: 'grid', gridTemplateColumns: '190px 1fr', alignItems: 'center',
                padding: '18px 24px', gap: 16,
                borderBottom: i < mechRows.length - 1 ? `1px solid ${theme.border}` : 'none',
                background: i % 2 === 0 ? 'transparent' : (dark ? 'rgba(255,255,255,0.014)' : 'rgba(0,0,0,0.014)'),
              }}>
                <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, textTransform: 'uppercase', color: theme.accent, fontFamily: "'JetBrains Mono',monospace", display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ display: 'inline-block', width: 5, height: 5, borderRadius: '50%', background: theme.accent, opacity: 0.7, flexShrink: 0 }} />
                  {r.label}
                </div>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: theme.text, marginBottom: 2 }}>{r.value}</div>
                  <div style={{ fontSize: 12, color: theme.sub }}>{r.sub}</div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* ── WHY IT WORKS ── */}
        <div style={divider} />
        <div className="about-section" style={{ animationDelay: '160ms' }}>
          <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 3, textTransform: 'uppercase', color: theme.sub, marginBottom: 8, fontFamily: "'JetBrains Mono',monospace" }}>The Logic</div>
          <h2 style={{ margin: '0 0 10px', fontSize: 26, fontWeight: 800, letterSpacing: -0.5, color: theme.text }}>Why This Works When Others Don't</h2>
          <p style={{ margin: '0 0 28px', fontSize: 14, color: theme.textDim, lineHeight: 1.75, maxWidth: 560 }}>
            Most traders fail at execution, not analysis. The system is built to remove the decisions that cause mistakes.
          </p>
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : 'repeat(3,1fr)', gap: isMobile ? 10 : 14 }}>
            {pillars.map((p, i) => {
              const pillarAccent = i === 0 ? theme.accent : i === 1 ? theme.purple : theme.amber;
              const pillarBg = i === 0
                ? (dark ? 'rgba(0,201,167,0.06)' : 'rgba(2,132,199,0.05)')
                : i === 1
                  ? (dark ? 'rgba(124,107,255,0.06)' : 'rgba(109,93,232,0.05)')
                  : (dark ? 'rgba(245,158,11,0.06)' : 'rgba(217,119,6,0.05)');
              return (
                <div key={p.title} style={{ background: pillarBg, border: `1px solid ${pillarAccent}28`, borderRadius: 14, padding: '24px 22px', position: 'relative', overflow: 'hidden' }}>
                  <div style={{ position: 'absolute', top: 0, left: 0, right: 0, height: 4, background: pillarAccent, borderRadius: '14px 14px 0 0' }} />
                  <div style={{ fontSize: 22, fontWeight: 800, color: pillarAccent, marginBottom: 8, marginTop: 6, opacity: 0.35, fontFamily: "'JetBrains Mono',monospace", lineHeight: 1 }}>
                    {String(i + 1).padStart(2, '0')}
                  </div>
                  <div style={{ fontSize: 14, fontWeight: 800, color: pillarAccent, marginBottom: 10, lineHeight: 1.3 }}>{p.title}</div>
                  <div style={{ fontSize: 13, color: theme.textDim, lineHeight: 1.8 }}>{p.body}</div>
                </div>
              );
            })}
          </div>
        </div>

        {/* ── TRANSPARENCY ── */}
        <div style={divider} />
        <div className="about-section" style={{ animationDelay: '240ms' }}>
          <div style={{
            background: dark ? 'rgba(0,201,167,0.06)' : 'rgba(2,132,199,0.05)',
            border: `1px solid ${dark ? 'rgba(0,201,167,0.18)' : 'rgba(2,132,199,0.18)'}`,
            borderRadius: 14,
            padding: '24px 30px',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
              <span style={{ display: 'inline-block', width: 8, height: 8, borderRadius: '50%', background: theme.accent, flexShrink: 0 }} />
              <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 2, textTransform: 'uppercase', color: theme.accent, fontFamily: "'JetBrains Mono',monospace" }}>On Transparency</div>
            </div>
            <p style={{ margin: 0, fontSize: 15, lineHeight: 1.85, color: theme.text }}>
              Every trade you see here was taken live. No back-tested hypotheticals, no demo account, no selective display. Good months and bad months both get logged. That's the only way this data means anything.
            </p>
          </div>
        </div>

        {/* ── CTA ── */}
        <div style={divider} />
        <div className="about-section" style={{ animationDelay: '300ms', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 24, background: dark ? `linear-gradient(135deg, ${theme.card} 60%, rgba(0,201,167,0.06) 100%)` : theme.card, border: `1px solid ${theme.border}`, borderRadius: 16, padding: '32px 36px' }}>
          <div style={{ flex: 1, minWidth: 240 }}>
            <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 3, textTransform: 'uppercase', color: theme.sub, marginBottom: 10, fontFamily: "'JetBrains Mono',monospace" }}>Get Started</div>
            <h3 style={{ margin: '0 0 10px', fontSize: 20, fontWeight: 800, color: theme.text, letterSpacing: -0.3 }}>The full strategy is free.</h3>
            <p style={{ margin: 0, fontSize: 14, color: theme.textDim, maxWidth: 380, lineHeight: 1.75 }}>
              The entry rules, exit rules, position sizing — all of it. Download the guide and trade the exact same system tracked here.
            </p>
          </div>
          <a href="#" style={{ display: 'inline-flex', alignItems: 'center', gap: 10, padding: '14px 28px', borderRadius: 12, background: theme.accent, color: '#fff', fontFamily: "'Outfit',sans-serif", fontSize: 14, fontWeight: 700, textDecoration: 'none', letterSpacing: 0.2, whiteSpace: 'nowrap', flexShrink: 0, boxShadow: `0 4px 20px ${theme.accent}33` }}>
            ↓ Download the Simple Markets Strategy
          </a>
        </div>

      </div>
    );
  };

  /* ── DASHBOARD ──────────────────────────────────────────────────────────── */
  const Dashboard = () => (
    <>
      {/* KPI strip */}
      <div style={isMobile ? { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 16 } : { display: 'grid', gridTemplateColumns: 'repeat(6,1fr)', gap: 10, marginBottom: 18 }}>
        {[
          {
            l: 'Net P&L', vis: 'pnl', noYr: false,
            getV: (cSt) => fmtU(cSt.totalR * riskDollar),
            getS: (cSt) => fmtR(cSt.totalR),
            isPos: (cSt) => cSt.totalR >= 0
          },
          {
            l: 'Total R', vis: 'totalR', noYr: false,
            getV: (cSt) => totalRMode === 'R'
              ? fmtR(cSt.totalR)
              : ((cSt.totalR * (selAcctData?.riskPercent ?? 1)) >= 0 ? '+' : '') + (cSt.totalR * (selAcctData?.riskPercent ?? 1)).toFixed(1) + '%',
            getS: (cSt) => `${cSt.total} trades`,
            isPos: (cSt) => cSt.totalR >= 0
          },
          {
            l: 'Win Rate', vis: 'winRate', noYr: false,
            getV: (cSt) => fmtPct(cSt.wr),
            getS: (cSt) => `${cSt.wins}W · ${cSt.losses}L`,
            isPos: (cSt) => cSt.wr >= 50
          },
          {
            l: 'Profit Factor', vis: 'profitFactor', noYr: true,
            getV: (cSt) => isFinite(cSt.pf) ? cSt.pf.toFixed(2) : '∞',
            getS: () => 'Gross W / Gross L',
            isPos: (cSt) => cSt.pf >= 1
          },
          {
            l: 'Expectancy', vis: null, noYr: true,
            getV: (cSt) => fmtR(cSt.exp),
            getS: () => 'Per trade',
            isPos: (cSt) => cSt.exp >= 0
          },
          {
            l: 'Max Drawdown', vis: 'drawdown', noYr: false,
            getV: (cSt) => '-' + cSt.maxDD.toFixed(2) + 'R',
            getS: (cSt) => `${cSt.total} trades`,
            neg: true
          },
        ].filter(k => !k.vis || isAdmin || visibility[k.vis]).map(k => {
          const cy = k.noYr ? 'all' : (cardYears[k.l] || 'all');
          const cFilt = cy === 'all' ? filtered : filtered.filter(t => t.date.startsWith(cy));
          const cSt = cy === 'all' ? st : computeStats(cFilt, breakEvenAsWin);
          const v = k.getV(cSt);
          const pos = k.isPos ? k.isPos(cSt) : false;
          const topColor = k.neg ? theme.loss : pos ? theme.gain : theme.loss;
          const cardIdx = [k.l].map((_, idx) => idx)[0] ?? 0;
          return (
            <div key={k.l} className="kpi-card" style={{
              ...card, padding: '15px 16px', borderTop: `2px solid ${topColor}`, position: 'relative', overflow: 'hidden', animationDelay: (([
                'Net P&L', 'Total R', 'Win Rate', 'Profit Factor', 'Expectancy', 'Max Drawdown'
              ].indexOf(k.l)) * 55) + 'ms'
            }}>
              <div style={{ position: 'absolute', top: 0, left: 0, right: 0, height: 40, background: `linear-gradient(180deg, ${topColor}0a 0%, transparent 100%)`, pointerEvents: 'none' }} />
              {/* Per-card year buttons — hidden for Profit Factor & Expectancy & mobile */}
              {!k.noYr && kpiYears.length > 0 && !isMobile && (
                <div style={{ display: 'flex', gap: 3, marginBottom: 9, flexWrap: 'wrap' }}>
                  {['all', ...kpiYears].map(y => (
                    <button key={y} onClick={() => setCardYears(p => ({ ...p, [k.l]: y }))}
                      style={{ padding: '2px 6px', borderRadius: 4, fontSize: 9, fontWeight: 700, cursor: 'pointer', background: cy === y ? theme.accent : 'transparent', color: cy === y ? '#fff' : theme.sub, border: `1px solid ${cy === y ? theme.accent : theme.border}`, fontFamily: "'Outfit',sans-serif", transition: 'all .1s', lineHeight: 1.5 }}>
                      {y === 'all' ? 'All' : "'" + y.slice(2)}
                    </button>
                  ))}
                </div>
              )}
              <div style={{ display: 'flex', alignItems: 'center', gap: 5, marginBottom: 8 }}>
                <div style={{ fontSize: 9.5, fontWeight: 700, letterSpacing: 1.8, textTransform: 'uppercase', color: theme.sub }}>{k.l}</div>
                {/* R / % mode toggle — Total R only */}
                {k.l === 'Total R' && (
                  <button onClick={() => setTotalRMode(m => m === 'R' ? '%' : 'R')} style={{ padding: '1px 6px', borderRadius: 3, fontSize: 8, fontWeight: 700, cursor: 'pointer', background: totalRMode === '%' ? `${theme.accent}20` : theme.muted, color: totalRMode === '%' ? theme.accent : theme.sub, border: `1px solid ${totalRMode === '%' ? theme.accent + '55' : theme.border}`, fontFamily: "'JetBrains Mono',monospace", transition: 'all .12s', lineHeight: 1.6 }}>
                    {totalRMode === 'R' ? 'R→%' : '%→R'}
                  </button>
                )}
              </div>
              {/* Animated value — key change triggers fade-in on year switch */}
              <div key={cy + v + totalRMode} style={{ fontSize: isMobile ? 15 : 20, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", lineHeight: 1, color: topColor, marginBottom: 4, animation: 'numFade 0.28s ease both' }}>{v}</div>
              <div style={{ fontSize: 10, color: theme.sub }}>{k.getS(cSt)}</div>
            </div>
          );
        })}
      </div>

      {/* Cumulative R + Performance Metrics */}
      {(isAdmin || visibility.equity) && (() => {
        const tickInterval = Math.max(0, Math.floor(eq.length / 8) - 1);
        const chartBg = dark ? '#050c18' : theme.card;
        const chartCard = { background: chartBg, border: `1px solid ${theme.border}`, borderRadius: 14, padding: isMobile ? '14px 12px 10px' : '18px 20px 14px', boxShadow: dark ? '0 2px 24px rgba(0,0,0,0.55), inset 0 1px 0 rgba(255,255,255,0.03)' : '0 2px 12px rgba(0,0,0,0.07)', minWidth: 0, overflow: 'hidden' };
        const axisProps = { tick: { fontSize: 8, fill: theme.sub, fontFamily: "'JetBrains Mono',monospace" }, tickLine: false, axisLine: false };
        const eqInit = eq.length ? eq[0].bal : (selAcctData?.initialBalance ?? 25000);
        // Derive cumulative % from balance for each point
        const eqData = eq.map(p => ({ ...p, cumPct: +((p.bal / eqInit - 1) * 100).toFixed(2) }));
        const isPct = equityDispMode === '%';
        const livePoint = eqData[eqData.length - 1];
        const liveVal = isPct ? livePoint?.cumPct : livePoint?.cumR;
        return (
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '62fr 38fr', gap: 14, marginBottom: 18, alignItems: 'stretch', width: '100%', boxSizing: 'border-box', minWidth: 0 }}>
            {/* LEFT — Cumulative R / % (full-width, taller, stretched domain) */}
            <div style={chartCard}>
              <div style={{ display: 'flex', flexDirection: isMobile ? 'column' : 'row', alignItems: isMobile ? 'flex-start' : 'center', gap: isMobile ? 10 : 8, marginBottom: 14 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 9, fontWeight: 700, letterSpacing: 2.5, textTransform: 'uppercase', color: theme.sub }}>
                    {equityMode === 'compare' ? 'R — Year Comparison' : equityMode !== 'all' ? `Cumulative ${isPct ? '%' : 'R'} — ${equityMode}` : `Cumulative ${isPct ? '%' : 'R'} — All Time`}
                  </div>
                  {equityMode !== 'compare' && eqData.length > 1 && (
                    <div style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: isMobile ? 18 : 22, fontWeight: 700, color: (liveVal ?? 0) >= 0 ? theme.gain : theme.loss, lineHeight: 1, marginTop: 4 }}>
                      {(liveVal ?? 0) >= 0 ? '+' : ''}{liveVal?.toFixed(2)}{isPct ? '%' : 'R'}
                    </div>
                  )}
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  {/* R / % toggle */}
                  {equityMode !== 'compare' && (
                    <div style={{ display: 'flex', gap: 2, background: theme.muted, borderRadius: 8, padding: 3 }}>
                      {['R', '%'].map(m => (
                        <button key={m} onClick={() => setEquityDispMode(m)}
                          style={{
                            fontSize: 11, fontWeight: 700, padding: '4px 11px', borderRadius: 6, border: 'none', cursor: 'pointer', fontFamily: "'JetBrains Mono',monospace", transition: 'all .15s ease',
                            background: equityDispMode === m ? theme.accent : 'transparent',
                            color: equityDispMode === m ? (dark ? '#000' : '#fff') : theme.sub,
                            boxShadow: equityDispMode === m ? `0 1px 6px ${theme.accent}50` : 'none',
                          }}>
                          {m}
                        </button>
                      ))}
                    </div>
                  )}
                  <YrSelector />
                </div>
              </div>
              {equityMode === 'compare' ? (
                <>
                  <div style={{ display: 'flex', gap: 10, marginBottom: 10 }}>
                    {compareYears.map((y, i) => (
                      <div key={y} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 11 }}>
                        <div style={{ width: 18, height: 3, borderRadius: 2, background: YC[i % YC.length] }} />
                        <span style={{ color: theme.sub, fontWeight: 600 }}>{y}</span>
                      </div>
                    ))}
                  </div>
                  <ResponsiveContainer width="100%" height={isMobile ? 170 : 380}>
                    <LineChart data={compareData} margin={{ top: 6, right: 4, bottom: isMobile ? 2 : 28, left: isMobile ? -14 : 4 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke={dark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.06)'} vertical={false} />
                      <XAxis dataKey="trade" {...axisProps} label={isMobile ? undefined : { value: 'Number of trades', position: 'insideBottom', offset: -12, fill: theme.accent, fontSize: 9, fontFamily: "'JetBrains Mono',monospace", letterSpacing: 1 }} tickFormatter={v => v === 0 ? '' : '#' + v} />
                      <YAxis {...axisProps} tickFormatter={v => (v >= 0 ? '+' : '') + v + 'R'} width={isMobile ? 28 : 36} />
                      <Tooltip content={<MultiTip />} />
                      {compareYears.map((y, i) => <Line key={y} type="monotone" dataKey={y} stroke={YC[i % YC.length]} strokeWidth={2.5} dot={false} connectNulls />)}
                    </LineChart>
                  </ResponsiveContainer>
                </>
              ) : (
                <ResponsiveContainer width="100%" height={isMobile ? 170 : 380}>
                  <AreaChart data={eqData} margin={{ top: 6, right: 4, bottom: isMobile ? 2 : 28, left: isMobile ? -14 : 4 }}>
                    <defs>
                      <linearGradient id="rrGrad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={theme.accent} stopOpacity={dark ? 0.32 : 0.22} />
                        <stop offset="65%" stopColor={theme.accent} stopOpacity={0.06} />
                        <stop offset="100%" stopColor={theme.accent} stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={dark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.06)'} vertical={false} />
                    <XAxis dataKey="n" {...axisProps} interval={tickInterval} tickFormatter={v => v === 0 ? '' : '#' + v}
                      label={isMobile ? undefined : { value: 'Number of trades', position: 'insideBottom', offset: -12, fill: theme.accent, fontSize: 9, fontFamily: "'JetBrains Mono',monospace", letterSpacing: 1 }} />
                    <YAxis {...axisProps}
                      tickFormatter={v => isPct ? (v >= 0 ? '+' : '') + v.toFixed(1) + '%' : (v >= 0 ? '+' : '') + v.toFixed(0) + 'R'}
                      width={isMobile ? 32 : (isPct ? 46 : 38)}
                      domain={[
                        dataMin => Math.floor(Math.min(dataMin, 0) - Math.max(Math.abs(dataMin) * 0.22, 2)),
                        dataMax => Math.ceil(dataMax + Math.max(dataMax * 0.08, 1)),
                      ]} />
                    <ReferenceLine y={0} stroke={dark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)'} strokeDasharray="4 3" />
                    <Tooltip content={<RRTip />} />
                    <Area type="monotone" dataKey={isPct ? 'cumPct' : 'cumR'} stroke={theme.accent} fill="url(#rrGrad)" strokeWidth={2.5} dot={false} />
                  </AreaChart>
                </ResponsiveContainer>
              )}
            </div>
            {/* RIGHT — Performance Metrics (moved from bottom) */}
            {(isAdmin || visibility.streaks) && (
              <div style={{ ...chartCard, overflowY: 'auto' }}>
                <div style={{ fontSize: 9, fontWeight: 700, letterSpacing: 2.5, textTransform: 'uppercase', color: theme.sub, marginBottom: 16, display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ display: 'inline-block', width: 3, height: 10, borderRadius: 2, background: theme.accent }} />
                  Performance Metrics
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
                  {[
                    { l: 'Current streak', v: st.curDir ? `${st.curDir === 'W' ? '+' : '-'}${st.curCount}${st.curDir}` : '—', d: 'Ongoing', c: st.curDir === 'W' ? theme.gain : st.curDir === 'L' ? theme.loss : theme.text },
                    { l: 'Best win streak', v: st.bWS.count ? `+${st.bWS.count}W` : '—', d: fmtRange(st.bWS.start, st.bWS.end), c: theme.gain },
                    { l: 'Worst loss streak', v: st.wLS.count ? `-${st.wLS.count}L` : '—', d: fmtRange(st.wLS.start, st.wLS.end), c: theme.loss },
                    { l: 'Best month', v: st.bestMon.key ? fmtR(st.bestMon.r) : '—', d: fmtMon(st.bestMon.key), c: theme.gain },
                    { l: 'Worst month', v: st.worstMon.key ? fmtR(st.worstMon.r) : '—', d: fmtMon(st.worstMon.key), c: theme.loss },
                    { l: 'Avg win', v: `+${st.avgW.toFixed(2)}R`, d: `Over ${st.wins} wins`, c: theme.gain },
                    { l: 'Avg loss', v: `-${st.avgL.toFixed(2)}R`, d: `Over ${st.losses} losses`, c: theme.loss },
                    { l: 'Avg trades/week', v: st.avgPerWeek || '—', d: 'Based on trading days', c: theme.text },
                    { l: 'Total trades', v: st.total, d: 'All time', c: theme.text },
                    { l: 'Expectancy', v: fmtR(st.exp), d: 'Per trade', c: st.exp >= 0 ? theme.gain : theme.loss },
                  ].map((m, i) => (
                    <div key={m.l} style={{ padding: isMobile ? '6px 0' : '8px 0', borderBottom: i < 9 ? `1px solid ${theme.border}33` : 'none' }}>
                      <div style={{ fontSize: 9, color: theme.sub, marginBottom: 1 }}>{m.l}</div>
                      <div style={{ display: 'flex', alignItems: 'baseline', gap: 6 }}>
                        <span style={{ fontSize: isMobile ? 13 : 15, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: m.c, lineHeight: 1 }}>{m.v}</span>
                        <span style={{ fontSize: 9, color: theme.sub }}>{m.d}</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        );
      })()}

      {/* Monthly + Donut */}
      <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '58fr 42fr', gap: 14, marginBottom: 18 }}>
        <div style={card}>
          {/* Header row: label + year nav + month filter badge */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12, flexWrap: 'wrap' }}>
            <div style={{ flex: 1 }}>
              <SectionLabel accent>Monthly R Breakdown</SectionLabel>
            </div>
            {dashMonthFilter && (
              <button onClick={() => setDashMonthFilter(null)} style={{ fontSize: 10, padding: '2px 8px', borderRadius: 5, border: `1px solid ${theme.accent}60`, background: `${theme.accent}12`, color: theme.accent, cursor: 'pointer', fontFamily: "'Outfit',sans-serif", fontWeight: 600 }}>
                {new Date(dashMonthFilter + '-01').toLocaleDateString('en-US', { month: 'short', year: 'numeric' })} ✕
              </button>
            )}
            <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              <button onClick={() => { const ys = [...new Set(mon.map(m => m.year))].sort(); const i = ys.indexOf(chartYear); if (i > 0) setChartYear(ys[i - 1]); }} style={{ ...btnG, padding: '3px 8px', fontSize: 12 }}>‹</button>
              <span style={{ fontSize: 12, fontWeight: 700, minWidth: 36, textAlign: 'center' }}>{chartYear}</span>
              <button onClick={() => { const ys = [...new Set(mon.map(m => m.year))].sort(); const i = ys.indexOf(chartYear); if (i < ys.length - 1) setChartYear(ys[i + 1]); }} style={{ ...btnG, padding: '3px 8px', fontSize: 12 }}>›</button>
            </div>
          </div>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart
              data={mon.filter(m => m.year === chartYear)}
              margin={{ top: 28, right: 8, bottom: 0, left: 4 }}
              barCategoryGap="28%"
              onClick={data => {
                if (data?.activePayload?.[0]) {
                  const key = data.activePayload[0].payload.key;
                  setMonthPopup(key);
                }
              }}>
              <CartesianGrid strokeDasharray="3 3" stroke={dark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.06)'} vertical={false} />
              <XAxis dataKey="month" tick={{ fontSize: 9, fill: theme.sub }} tickLine={false} axisLine={false} />
              <YAxis tick={{ fontSize: 8, fill: theme.sub }} tickLine={false} axisLine={false} width={28} />
              <Tooltip content={<MonBarTip />} />
              <Bar dataKey="r" radius={[4, 4, 0, 0]} maxBarSize={28} cursor="pointer">
                <LabelList dataKey="r" content={(props) => {
                  const { x, y, width, height, value } = props;
                  if (value === undefined || value === null) return null;
                  const color = value >= 0 ? theme.gain : theme.loss;
                  const label = (value >= 0 ? '+' : '') + value.toFixed(1) + 'R';
                  // Positive bars: label floats above the bar. Negative bars: label sits just inside the top edge (zero crossing).
                  const labelY = value >= 0 ? y - 7 : y + 15;
                  return (
                    <text x={x + width / 2} y={labelY} textAnchor="middle"
                      style={{ fontSize: 11, fontFamily: "'JetBrains Mono',monospace", fill: color, fontWeight: 700 }}>
                      {label}
                    </text>
                  );
                }} />
                {mon.filter(m => m.year === chartYear).map((d, i) => (
                  <Cell key={i}
                    fill={dashMonthFilter === d.key ? theme.amber : d.r >= 0 ? theme.gain : theme.loss}
                    opacity={dashMonthFilter && dashMonthFilter !== d.key ? 0.3 : 1}
                  />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
          {mon.filter(m => m.year === chartYear).length === 0 && (
            <div style={{ height: 100, display: 'flex', alignItems: 'center', justifyContent: 'center', color: theme.sub, fontSize: 13 }}>No trades in {chartYear}</div>
          )}
        </div>
        <div style={card}>
          <SectionLabel accent>Trade Distribution</SectionLabel>
          {donut.length ? (
            <>
              <ResponsiveContainer width="100%" height={130}>
                <PieChart>
                  <Pie data={donut} cx="50%" cy="50%" innerRadius={38} outerRadius={58} paddingAngle={3} dataKey="val">
                    {donut.map((d, i) => <Cell key={i} fill={d.fill} />)}
                  </Pie>
                  <Tooltip contentStyle={{ background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 8, fontSize: 12 }} />
                </PieChart>
              </ResponsiveContainer>
              <div style={{ display: 'flex', justifyContent: 'center', gap: 12, marginTop: 4 }}>
                {donut.map(d => (
                  <div key={d.name} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 11 }}>
                    <div style={{ width: 7, height: 7, borderRadius: 2, background: d.fill }} />
                    <span style={{ color: theme.sub }}>{d.name}</span>
                    <span style={{ fontWeight: 700, fontFamily: "'JetBrains Mono',monospace" }}>{d.val}</span>
                  </div>
                ))}
              </div>
            </>
          ) : <div style={{ height: 130, display: 'flex', alignItems: 'center', justifyContent: 'center', color: theme.sub, fontSize: 13 }}>No trades yet</div>}
        </div>
      </div>

      {/* Daily / Weekly Cumulative P&L + Drawdown */}
      {(isAdmin || visibility.equity) && (() => {
        const acct = accounts.find(x => x.id === selectedAccount);
        const ddRp = acct?.riskPercent ?? 1.0;
        const ddInit = acct?.initialBalance ?? 25000;
        const dailyData = buildDailyEqAndDD(filtered, ddRp, ddInit, false);
        const currentDD = dailyData.length ? dailyData[dailyData.length - 1].dd : 0;
        const currentBal = dailyData.length ? dailyData[dailyData.length - 1].bal : ddInit;
        const maxDD = dailyData.length ? Math.min(...dailyData.map(d => d.dd)) : 0;
        const maxDDDate = dailyData.find(d => d.dd === maxDD)?.date || '';
        const tickCount = Math.max(0, Math.floor(dailyData.length / 5) - 1);
        const chartBg = dark ? '#050c18' : theme.card;
        const cCard = { background: chartBg, border: `1px solid ${theme.border}`, borderRadius: 14, padding: '18px 20px 12px', boxShadow: dark ? '0 2px 24px rgba(0,0,0,0.55), inset 0 1px 0 rgba(255,255,255,0.03)' : '0 2px 12px rgba(0,0,0,0.07)' };
        const axP = { tick: { fontSize: 8, fill: theme.sub, fontFamily: "'JetBrains Mono',monospace" }, tickLine: false, axisLine: false };
        return (
          <div style={{ marginBottom: 18 }}>
            {/* Section header + Daily / Weekly toggle */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
              <div style={{ flex: 1, display: 'flex', alignItems: 'center', gap: 8, fontSize: 10, fontWeight: 700, letterSpacing: 2, textTransform: 'uppercase', color: theme.sub }}>
                <span style={{ display: 'inline-block', width: 3, height: 12, borderRadius: 2, background: theme.gain, flexShrink: 0 }} />
                Performance Analysis
              </div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 14 }}>
              {/* LEFT — Risk-Adjusted Ratios */}
              <div style={cCard}>
                <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 1.5, textTransform: 'uppercase', color: theme.textDim, marginBottom: 16 }}>
                  Risk-Adjusted Ratios
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                  {[
                    { key: 'sharpe', label: 'Sharpe Ratio', val: ratios.sharpe, desc: 'Return / Volatility' },
                    { key: 'sortino', label: 'Sortino Ratio', val: ratios.sortino, desc: 'Return / Downside Vol' },
                    { key: 'calmar', label: 'Calmar Ratio', val: ratios.calmar, desc: 'Ann. Return / Max DD' },
                    { key: 'recovery', label: 'Recovery Factor', val: ratios.recovery, desc: 'Total R / Max DD R' },
                  ].map(({ key, label, val, desc }) => {
                    const color = val === null ? theme.textDim
                      : key === 'recovery'
                        ? (val >= 3 ? theme.gain : val >= 1 ? theme.amber : theme.loss)
                        : (val >= 2 ? theme.gain : val >= 1 ? theme.amber : theme.loss);
                    const grade = val === null ? '—'
                      : key === 'recovery'
                        ? (val >= 3 ? 'Excellent' : val >= 1 ? 'Good' : val >= 0.5 ? 'Fair' : 'Poor')
                        : (val >= 2 ? 'Excellent' : val >= 1 ? 'Good' : val >= 0.5 ? 'Fair' : 'Poor');
                    const gradeColor = grade === 'Excellent' ? theme.gain : grade === 'Good' ? theme.amber : grade === 'Fair' ? theme.sky : theme.loss;
                    return (
                      <div key={key} style={{ background: dark ? 'rgba(255,255,255,0.03)' : 'rgba(0,0,0,0.025)', border: `1px solid ${theme.border}`, borderRadius: 10, padding: '14px 16px' }}>
                        <div style={{ fontSize: 11, fontWeight: 600, color: theme.textDim, marginBottom: 8 }}>{label}</div>
                        <div style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: 24, fontWeight: 700, color, lineHeight: 1, marginBottom: 10 }}>
                          {val !== null ? val.toFixed(2) : '—'}
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 5 }}>
                          <span style={{ display: 'inline-block', width: 7, height: 7, borderRadius: '50%', background: gradeColor, flexShrink: 0 }} />
                          <span style={{ fontSize: 12, fontWeight: 700, color: gradeColor }}>{grade}</span>
                        </div>
                        <div style={{ fontSize: 11, color: theme.textDim }}>{desc}</div>
                      </div>
                    );
                  })}
                </div>
                <div style={{ marginTop: 14, paddingTop: 10, borderTop: `1px solid ${theme.border}`, fontSize: 11, color: theme.textDim, lineHeight: 1.8 }}>
                  Sharpe / Sortino / Calmar: &gt;1 Good · &gt;2 Excellent&nbsp;&nbsp;·&nbsp;&nbsp;Recovery: &gt;1 Good · &gt;3 Excellent
                </div>
              </div>
              {/* RIGHT — Drawdown */}
              <div style={cCard}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 14 }}>
                  <div style={{ fontSize: 9, fontWeight: 700, letterSpacing: 2.5, textTransform: 'uppercase', color: theme.sub }}>Drawdown</div>
                  <div style={{ display: 'flex', gap: 6, alignItems: 'baseline' }}>
                    <span style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: 13, fontWeight: 700, color: theme.loss }}>{fmtU(maxDD)}</span>
                    <span style={{ fontSize: 10.5, color: theme.sub }}>({maxDD !== 0 ? Math.abs(maxDD / ddInit * 100).toFixed(1) : '0.0'}%)</span>
                  </div>
                </div>
                <ResponsiveContainer width="100%" height={180}>
                  <AreaChart data={dailyData} margin={{ top: 6, right: 6, bottom: 20, left: 4 }}>
                    <defs>
                      <linearGradient id="ddGrad" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={theme.loss} stopOpacity={dark ? 0.35 : 0.25} />
                        <stop offset="85%" stopColor={theme.loss} stopOpacity={0.03} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={dark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.06)'} vertical={false} />
                    <XAxis dataKey="date" {...axP} interval={tickCount} tickFormatter={fmtAxisDate} />
                    <YAxis {...axP} tickFormatter={v => v === 0 ? '0.00' : (v < 0 ? '-$' : '$') + Math.abs(v / 1000).toFixed(1) + 'k'} width={44} />
                    <ReferenceLine y={0} stroke={dark ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.12)'} strokeDasharray="4 4" />
                    <Tooltip contentStyle={{ background: dark ? '#0b1628' : theme.card, border: `1px solid ${theme.border}`, borderRadius: 8, fontSize: 12 }}
                      formatter={(v) => [fmtU(v), 'Drawdown']}
                      labelFormatter={l => fmtAxisDate(l)} />
                    <Area type="monotone" dataKey="dd" stroke={theme.loss} fill="url(#ddGrad)" strokeWidth={2} dot={false} />
                  </AreaChart>
                </ResponsiveContainer>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 10.5, color: theme.sub, marginTop: 4, paddingTop: 6, borderTop: `1px solid ${theme.borderFaint}` }}>
                  <span>Current: <span style={{ fontFamily: "'JetBrains Mono',monospace", color: currentDD < 0 ? theme.loss : theme.sub }}>{fmtU(currentDD)}</span></span>
                  {maxDDDate && <span>Max on {fmtAxisDate(maxDDDate)}</span>}
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {/* Triangle + Metrics */}
      <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '38fr 62fr', gap: 14, marginTop: 18, width: '100%', boxSizing: 'border-box', minWidth: 0 }}>
        {(isAdmin || visibility.triangle) && <div style={{ ...card, display: 'flex', flexDirection: 'column' }}>
          <SectionLabel accent>Edge Score</SectionLabel>
          <ResponsiveContainer width="100%" height={190}>
            <RadarChart data={tri.data} margin={{ top: 8, right: 34, bottom: 8, left: 34 }}>
              <PolarGrid stroke={theme.border} gridType="polygon" />
              <PolarAngleAxis dataKey="axis" tick={{ fontSize: 11, fill: theme.sub, fontFamily: "'Outfit',sans-serif", fontWeight: 600 }} />
              <Radar dataKey="v" stroke={theme.accent} fill={theme.accent} fillOpacity={0.18} strokeWidth={2.5} dot={{ r: 4, fill: theme.accent, strokeWidth: 0 }} />
              <Tooltip contentStyle={{ background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v, n) => [v + '/100', n]} />
            </RadarChart>
          </ResponsiveContainer>
          <div style={{ marginTop: 2 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: 6 }}>
              <span style={{ fontSize: 11, color: theme.sub, fontWeight: 600 }}>Composite score</span>
              <span style={{ fontSize: 22, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: tri.score >= 70 ? theme.gain : tri.score >= 45 ? theme.accent : theme.loss }}>
                {tri.score}<span style={{ fontSize: 11, color: theme.sub, fontWeight: 400 }}>/100</span>
              </span>
            </div>
            <div style={{ height: 8, borderRadius: 4, background: theme.border, overflow: 'hidden' }}>
              <div style={{ height: '100%', borderRadius: 4, width: `${tri.score}%`, background: tri.score >= 70 ? theme.gain : tri.score >= 45 ? theme.accent : theme.loss, transition: 'width .6s ease' }} />
            </div>
            <div style={{ display: 'flex', gap: 6, marginTop: 10 }}>
              {[{ l: 'Win %', v: tri.nWR }, { l: 'W/L', v: tri.nWL }, { l: 'P.F.', v: tri.nPF }].map(m => (
                <div key={m.l} style={{ flex: 1, textAlign: 'center', background: theme.muted, borderRadius: 8, padding: '6px 4px' }}>
                  <div style={{ fontSize: 9, color: theme.sub, letterSpacing: 1, marginBottom: 2, textTransform: 'uppercase' }}>{m.l}</div>
                  <div style={{ fontSize: 14, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: m.v >= 70 ? theme.gain : m.v >= 40 ? theme.accent : theme.loss }}>{m.v}</div>
                </div>
              ))}
            </div>
          </div>
        </div>}
        {(isAdmin || visibility.tradeHistory) && (() => {
          const capTrades = (dashMonthFilter
            ? recentTrades.filter(t => t.date.startsWith(dashMonthFilter))
            : recentTrades).slice(0, 25);
          const monthLabel = dashMonthFilter
            ? new Date(dashMonthFilter + '-01').toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
            : null;
          return (
            <div style={{ ...card, padding: 0, overflow: 'hidden', display: 'flex', flexDirection: 'column', maxHeight: 430 }}>
              {/* Header */}
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 18px', borderBottom: `1px solid ${theme.border}`, background: `linear-gradient(90deg, ${theme.cardHi} 0%, ${theme.card} 100%)`, flexShrink: 0 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 9 }}>
                  <span style={{ display: 'inline-block', width: 3, height: 13, borderRadius: 2, background: theme.accent }} />
                  <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: 2, textTransform: 'uppercase', color: theme.sub }}>
                    {monthLabel ? `Trades — ${monthLabel}` : 'Recent Trades'}
                  </span>
                  <span style={{ fontSize: 10, fontWeight: 700, padding: '2px 7px', borderRadius: 20, background: theme.gainBg, color: theme.gain, border: `1px solid ${theme.gainBorder}`, fontFamily: "'JetBrains Mono',monospace" }}>
                    {dashMonthFilter ? capTrades.length : filtered.length}
                  </span>
                </div>
                <button onClick={() => setPage('trades')} style={{ fontSize: 11, fontWeight: 600, color: theme.accent, background: 'transparent', border: 'none', cursor: 'pointer', letterSpacing: 0.3, fontFamily: "'Outfit',sans-serif" }}>
                  Full log →
                </button>
              </div>
              {/* Scrollable table — desktop | card list — mobile */}
              <div style={{ overflowY: 'auto', flex: 1 }}>
                {isMobile ? (
                  <div style={{ display: 'flex', flexDirection: 'column' }}>
                    {capTrades.map((t, idx) => {
                      const ta = accounts.find(x => x.id === t.account);
                      const rp = t.riskPercent != null ? t.riskPercent : (ta?.riskPercent ?? 1.0);
                      const profit = t.r * (rp / 100) * (ta?.initialBalance ?? 25000);
                      return (
                        <div key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '11px 16px', borderBottom: `1px solid ${theme.borderFaint}`, background: idx % 2 === 0 ? 'transparent' : dark ? 'rgba(255,255,255,0.012)' : 'rgba(0,0,0,0.012)' }}>
                          {/* Color strip */}
                          <div style={{ width: 3, height: 36, borderRadius: 2, flexShrink: 0, background: t.r > 0 ? theme.gain : t.r < 0 ? theme.loss : theme.sub }} />
                          {/* Symbol + date */}
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 2 }}>
                              <span style={{ fontSize: 13, fontWeight: 700, letterSpacing: 0.4 }}>{t.symbol}</span>
                              <DirBadge type={t.type} />
                            </div>
                            <div style={{ fontSize: 10, fontFamily: "'JetBrains Mono',monospace", color: theme.sub }}>{t.date}</div>
                          </div>
                          {/* R + profit */}
                          <div style={{ textAlign: 'right', flexShrink: 0 }}>
                            <div style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: 14, fontWeight: 700, color: t.r > 0 ? theme.gain : t.r < 0 ? theme.loss : theme.sub, lineHeight: 1 }}>{fmtR(t.r)}</div>
                            <div style={{ fontFamily: "'JetBrains Mono',monospace", fontSize: 10.5, color: profit >= 0 ? theme.gain : theme.loss, marginTop: 3 }}>
                              {(profit >= 0 ? '+$' : '-$') + Math.abs(profit).toLocaleString('en-US', { maximumFractionDigits: 0 })}
                            </div>
                          </div>
                          {t.chartLink && <a href={t.chartLink} target="_blank" rel="noreferrer" style={{ color: theme.accent, fontSize: 14, textDecoration: 'none', flexShrink: 0 }}>📊</a>}
                        </div>
                      );
                    })}
                    {capTrades.length === 0 && (
                      <div style={{ padding: '36px 0', textAlign: 'center', color: theme.sub, fontSize: 13 }}>
                        {dashMonthFilter ? 'No trades this month.' : 'No trades yet.'}
                      </div>
                    )}
                  </div>
                ) : (
                  <>
                    <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                      <thead style={{ position: 'sticky', top: 0, zIndex: 2 }}>
                        <tr style={{ background: theme.muted }}>
                          {['Date', 'Symbol', 'Dir', 'R', 'Profit $', ''].map(h => (
                            <th key={h} style={{ fontSize: 9, fontWeight: 700, letterSpacing: 1.6, textTransform: 'uppercase', color: theme.sub, padding: '8px 12px', textAlign: 'left', borderBottom: `1px solid ${theme.border}`, whiteSpace: 'nowrap' }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {capTrades.map((t, idx) => {
                          const ta = accounts.find(x => x.id === t.account);
                          const rp = t.riskPercent != null ? t.riskPercent : (ta?.riskPercent ?? 1.0);
                          const profit = t.r * (rp / 100) * (ta?.initialBalance ?? 25000);
                          return (
                            <tr key={t.id} className="trade-row" style={{ background: idx % 2 === 0 ? 'transparent' : dark ? 'rgba(255,255,255,0.012)' : 'rgba(0,0,0,0.012)' }}>
                              <td style={{ padding: '8px 12px', fontSize: 11, fontFamily: "'JetBrains Mono',monospace", color: theme.textDim, borderBottom: `1px solid ${theme.borderFaint}`, whiteSpace: 'nowrap' }}>{t.date}</td>
                              <td style={{ padding: '8px 12px', fontSize: 12.5, fontWeight: 700, letterSpacing: 0.4, borderBottom: `1px solid ${theme.borderFaint}` }}>{t.symbol}</td>
                              <td style={{ padding: '8px 12px', borderBottom: `1px solid ${theme.borderFaint}` }}><DirBadge type={t.type} /></td>
                              <td style={{ padding: '8px 12px', fontFamily: "'JetBrains Mono',monospace", fontSize: 13, fontWeight: 700, color: t.r > 0 ? theme.gain : t.r < 0 ? theme.loss : theme.sub, borderBottom: `1px solid ${theme.borderFaint}` }}>{fmtR(t.r)}</td>
                              <td style={{ padding: '8px 12px', fontFamily: "'JetBrains Mono',monospace", fontSize: 11.5, fontWeight: 600, color: profit > 0 ? theme.gain : profit < 0 ? theme.loss : theme.sub, borderBottom: `1px solid ${theme.borderFaint}`, whiteSpace: 'nowrap' }}>
                                {(profit >= 0 ? '+$' : '-$') + Math.abs(profit).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
                              </td>
                              <td style={{ padding: '8px 12px', borderBottom: `1px solid ${theme.borderFaint}` }}>
                                {t.chartLink ? <a href={t.chartLink} target="_blank" rel="noreferrer" style={{ color: theme.accent, fontSize: 11, fontWeight: 600, textDecoration: 'none' }}>📊</a> : <span style={{ color: theme.border, fontSize: 12 }}>—</span>}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                    {capTrades.length === 0 && (
                      <div style={{ padding: '36px 0', textAlign: 'center', color: theme.sub, fontSize: 13 }}>
                        {dashMonthFilter ? 'No trades this month.' : 'No trades yet.'}
                      </div>
                    )}
                  </>
                )}
              </div>
              {/* Footer */}
              <div style={{ padding: '10px 16px', borderTop: `1px solid ${theme.border}`, display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexShrink: 0 }}>
                <span style={{ fontSize: 11, color: theme.sub }}>
                  {dashMonthFilter ? `${capTrades.length} trades in ${monthLabel}` : filtered.length > 25 ? `Showing 25 of ${filtered.length}` : `${filtered.length} trades`}
                </span>
                <div style={{ display: 'flex', gap: 7 }}>
                  {dashMonthFilter && <button onClick={() => setDashMonthFilter(null)} style={{ ...btnG, fontSize: 11, padding: '4px 9px' }}>Clear</button>}
                  {isAdmin && <button onClick={() => setLogPanelOpen(true)} style={{ ...btnP, fontSize: 11, padding: '5px 12px' }}>+ Log Trade</button>}
                </div>
              </div>
            </div>
          );
        })()}
      </div>
    </>
  );



  /* ── TRADES PAGE ────────────────────────────────────────────────────────── */
  const Trades = () => (
    <>
      <button onClick={() => setPage('dashboard')} style={{ ...btnG, marginBottom: 14, fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 5 }}>← Dashboard</button>
      <div style={{ display: 'flex', gap: 8, marginBottom: 14, flexWrap: 'wrap' }}>
        <input style={{ ...inp, flex: 1, minWidth: 140 }} placeholder="Search symbol…" value={tradeSearch} onChange={e => setTradeSearch(e.target.value)} />
        <select style={{ ...inp, width: 'auto' }} value={tradeSort} onChange={e => setTradeSort(e.target.value)}>
          <option value="date-desc">Newest first</option>
          <option value="r-desc">Best R first</option>
          <option value="r-asc">Worst R first</option>
        </select>
        <button style={btnP} onClick={() => setLogPanelOpen(true)}>+ Log Trade</button>
      </div>
      <div style={{ ...card, padding: 0, overflow: 'hidden' }}>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr>{['Date', 'Symbol', 'Dir', 'Account', 'R', 'Profit $', 'Chart', ''].map(h => (
              <th key={h} style={{ fontSize: isMobile ? 9 : 10, fontWeight: 700, letterSpacing: 1.2, textTransform: 'uppercase', color: theme.sub, padding: isMobile ? '6px 8px' : '9px 12px', textAlign: 'left', borderBottom: `1px solid ${theme.border}`, whiteSpace: 'nowrap' }}>{h}</th>
            ))}</tr></thead>
            <tbody>
              {dispTrades.slice(0, 100).map(t => {
                const a = accounts.find(x => x.id === t.account);
                const tdS = { padding: isMobile ? '6px 8px' : '9px 12px', borderBottom: `1px solid ${theme.border}44` };
                return (
                  <tr key={t.id}>
                    <td style={{ ...tdS, fontSize: isMobile ? 10 : 11.5, fontFamily: "'JetBrains Mono',monospace" }}>{t.date}</td>
                    <td style={{ ...tdS, fontSize: isMobile ? 11 : 13, fontWeight: 600 }}>{t.symbol}</td>
                    <td style={tdS}>
                      <DirBadge type={t.type} />
                    </td>
                    <td style={{ ...tdS, fontSize: isMobile ? 10 : 11.5, color: theme.sub }}>{a?.name || '—'}</td>
                    <td style={{ ...tdS, fontFamily: "'JetBrains Mono',monospace", fontSize: isMobile ? 11 : 13, fontWeight: 700, color: t.r > 0 ? theme.gain : t.r < 0 ? theme.loss : theme.sub }}>{fmtR(t.r)}</td>
                    <td style={{ ...tdS, fontFamily: "'JetBrains Mono',monospace", fontSize: 12, fontWeight: 600, color: (() => { const ta = accounts.find(x => x.id === t.account); const rp = t.riskPercent != null ? t.riskPercent : (ta?.riskPercent ?? 1.0); const p = t.r * (rp / 100) * (ta?.initialBalance ?? 25000); return p > 0 ? theme.gain : p < 0 ? theme.loss : theme.sub; })() }}>
                      {(() => { const ta = accounts.find(x => x.id === t.account); const rp = t.riskPercent != null ? t.riskPercent : (ta?.riskPercent ?? 1.0); const p = t.r * (rp / 100) * (ta?.initialBalance ?? 25000); return (p >= 0 ? '+$' : '-$') + Math.abs(p).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 }); })()}
                    </td>
                    <td style={tdS}>{t.chartLink ? <a href={t.chartLink} target="_blank" rel="noreferrer" style={{ color: theme.accent, fontSize: 11 }}>📊 View</a> : <span style={{ color: theme.border }}>—</span>}</td>
                    <td style={tdS}>
                      {isAdmin && <button onClick={() => removeTrade(t.id)} style={{ fontSize: 10, color: theme.sub, background: 'none', border: 'none', cursor: 'pointer', padding: '2px 4px' }} title="Delete trade">✕</button>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {dispTrades.length === 0 && (
            <div style={{ padding: 40, textAlign: 'center', color: theme.sub, fontSize: 14 }}>
              {trades.length === 0 ? 'No trades yet — click "+ Log Trade" to add your first trade.' : 'No trades match your filter.'}
            </div>
          )}
          {dispTrades.length > 100 && <div style={{ padding: '10px 14px', fontSize: 12, color: theme.sub, borderTop: `1px solid ${theme.border}` }}>Showing 100 of {dispTrades.length} trades</div>}
        </div>
      </div>
    </>
  );

  /* ── CALENDAR ─────────────────────────────────────────────────────────────── */
  const Calendar = () => {
    const key = `${calendarMonth.y}-${String(calendarMonth.m + 1).padStart(2, '0')}`;
    const mt = filtered.filter(t => t.date.startsWith(key));
    const mR = mt.reduce((s, t) => s + t.r, 0), mW = mt.filter(t => t.r > 0).length;
    const nm = new Date(calendarMonth.y, calendarMonth.m, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    return (
      <>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 14 }}>
          <button style={btnG} onClick={() => setCalendarMonth(p => { const d = new Date(p.y, p.m - 1, 1); return { y: d.getFullYear(), m: d.getMonth() }; })}>‹ Prev</button>
          <div style={{ flex: 1, textAlign: 'center', fontWeight: 700, fontSize: 15 }}>{nm}</div>
          <button style={btnG} onClick={() => setCalendarMonth(p => { const d = new Date(p.y, p.m + 1, 1); return { y: d.getFullYear(), m: d.getMonth() }; })}>Next ›</button>
        </div>
        <div style={{ display: 'flex', gap: 10, marginBottom: 14 }}>
          {[{ l: 'Month R', v: fmtR(mR), c: mR >= 0 ? theme.gain : theme.loss }, { l: 'Trades', v: mt.length, c: theme.text }, { l: 'Win %', v: mt.length ? fmtPct(mW / mt.length * 100) : '—', c: theme.text }].map(s => (
            <div key={s.l} style={{ ...card, flex: 1, padding: '14px 16px' }}>
              <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.2, textTransform: 'uppercase', color: theme.sub, marginBottom: 6 }}>{s.l}</div>
              <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: s.c }}>{s.v}</div>
            </div>
          ))}
        </div>
        <div style={card}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7,1fr)', gap: 3, marginBottom: 6 }}>
            {['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'].map(d => <div key={d} style={{ fontSize: 9, fontWeight: 700, textAlign: 'center', color: theme.sub, padding: '3px 0', letterSpacing: 1 }}>{d}</div>)}
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(7,1fr)', gap: 3 }}>
            {calCells.map((cell, i) => {
              if (!cell) return <div key={`e${i}`} />;
              const { d, ds, info } = cell;
              return (
                <div key={ds} style={{ minHeight: 54, borderRadius: 6, padding: '5px 6px', border: `1px solid ${info ? (info.r > 0 ? `${theme.gain}44` : `${theme.loss}44`) : theme.border}`, background: info ? (info.r > 0 ? `${theme.gain}16` : `${theme.loss}16`) : theme.muted }}>
                  <div style={{ fontSize: 9, color: theme.sub }}>{d}</div>
                  {info && <><div style={{ fontSize: 11, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: info.r > 0 ? theme.gain : theme.loss, marginTop: 3 }}>{fmtR(info.r)}</div><div style={{ fontSize: 8, color: theme.sub, marginTop: 2 }}>{info.n}t</div></>}
                </div>
              );
            })}
          </div>
        </div>
      </>
    );
  };

  /* ── ACCOUNTS ─────────────────────────────────────────────────────────────── */
  const Accounts = () => (
    <>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 14 }}>
        <button style={btnP} onClick={() => setShowNewAccount(true)}>+ Add Account</button>
      </div>
      {accounts.length === 0 && (
        <div style={{ ...card, textAlign: 'center', padding: 48, color: theme.sub }}>
          <div style={{ fontSize: 32, marginBottom: 12 }}>◎</div>
          <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 6 }}>No accounts yet</div>
          <div style={{ fontSize: 13 }}>Add your Live, Funded, and Challenge accounts to get started.</div>
        </div>
      )}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3,1fr)', gap: 14 }}>
        {accounts.map(a => {
          const at = trades.filter(t => t.account === a.id), ar = at.reduce((s, t) => s + t.r, 0), aw = at.filter(t => t.r > 0).length;

          return (
            <div key={a.id} style={card}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 10 }}>
                <AcctBadge type={a.type} />
                <div style={{ display: 'flex', gap: 4 }}>
                  {a.type === 'Challenge' && <button style={{ ...btnG, fontSize: 10, padding: '2px 7px' }} onClick={() => promoteAccount(a.id, 'Funded')}>→ Funded</button>}
                  {a.type === 'Funded' && <button style={{ ...btnG, fontSize: 10, padding: '2px 7px' }} onClick={() => promoteAccount(a.id, 'Live')}>→ Live</button>}
                </div>
              </div>
              <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 3 }}>{a.name}</div>
              <div style={{ fontSize: 11.5, color: theme.sub, marginBottom: 14 }}>{a.currency} · {a.riskPercent}% risk/trade</div>
              {[{ l: 'Total R', v: fmtR(ar), c: ar >= 0 ? theme.gain : theme.loss }, { l: 'Net P&L', v: fmtU(ar * (a.riskPercent / 100) * a.initialBalance), c: ar >= 0 ? theme.gain : theme.loss }, { l: 'Win Rate', v: at.length ? fmtPct(aw / at.length * 100) : '—', c: theme.text }, { l: 'Trades', v: at.length, c: theme.text }].map(m => (
                <div key={m.l} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '7px 0', borderBottom: `1px solid ${theme.border}44` }}>
                  <span style={{ fontSize: 12, color: theme.sub }}>{m.l}</span>
                  <span style={{ fontSize: 13, fontWeight: 600, fontFamily: "'JetBrains Mono',monospace", color: m.c }}>{m.v}</span>
                </div>
              ))}
              <div style={{ marginTop: 12, display: 'flex', justifyContent: 'space-between', gap: 6 }}>
                <button style={{ ...btnG, fontSize: 11 }} onClick={() => setEditAcct({ id: a.id, name: a.name, type: a.type, currency: a.currency, initialBalance: a.initialBalance, riskPercent: a.riskPercent })}>✏ Edit</button>
                <button style={{ ...btnG, fontSize: 11, color: theme.loss, borderColor: theme.loss + '44' }} onClick={() => removeAccount(a.id, a.name)}>Remove</button>
              </div>
            </div>
          );
        })}
      </div>
      {showNewAccount && (
        <>
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', zIndex: 100 }} onClick={() => setShowNewAccount(false)} />
          <div style={{ position: 'fixed', right: 0, top: 0, bottom: 0, width: 360, background: theme.card, borderLeft: `1px solid ${theme.border}`, zIndex: 101, padding: 26, display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div style={{ fontSize: 17, fontWeight: 700 }}>Add Account</div>
            {[{ l: 'Account name', k: 'name', t: 'text', ph: 'e.g. FTMO 100k Challenge' }, { l: 'Initial balance', k: 'initialBalance', t: 'number', ph: '10000' }].map(f => (
              <div key={f.k}>
                <label style={lbl}>{f.l}</label>
                <input style={inp} type={f.t} value={newAccount[f.k]} placeholder={f.ph} onChange={e => setNewAccount(p => ({ ...p, [f.k]: f.t === 'number' ? +e.target.value : e.target.value }))} />
              </div>
            ))}
            {[{ l: 'Type', k: 'type', opts: ['Challenge', 'Funded', 'Live'] }, { l: 'Currency', k: 'currency', opts: ['USD', 'EUR', 'GBP'] }].map(f => (
              <div key={f.k}>
                <label style={lbl}>{f.l}</label>
                <select style={inp} value={newAccount[f.k]} onChange={e => setNewAccount(p => ({ ...p, [f.k]: e.target.value }))}>{f.opts.map(o => <option key={o}>{o}</option>)}</select>
              </div>
            ))}
            <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
              <button style={{ ...btnG, flex: 1 }} onClick={() => setShowNewAccount(false)}>Cancel</button>
              <button style={{ ...btnP, flex: 1 }} onClick={addAccount}>Add Account</button>
            </div>
          </div>
        </>
      )}
      {/* Edit Account panel */}
      {editAcct && (
        <>
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.5)', zIndex: 100 }} onClick={() => setEditAcct(null)} />
          <div style={{ position: 'fixed', right: 0, top: 0, bottom: 0, width: 360, background: theme.card, borderLeft: `1px solid ${theme.border}`, zIndex: 101, padding: 26, display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div style={{ fontSize: 17, fontWeight: 700 }}>Edit Account</div>
            {[
              { l: 'Account name', k: 'name', t: 'text' },
              { l: 'Initial balance', k: 'initialBalance', t: 'number' },
              { l: 'Default risk % / trade', k: 'riskPercent', t: 'number' },
            ].map(f => (
              <div key={f.k}>
                <label style={lbl}>{f.l}</label>
                <input style={inp} type={f.t} value={editAcct[f.k] ?? ''} onChange={e => setEditAcct(p => ({ ...p, [f.k]: f.t === 'number' ? +e.target.value : e.target.value }))} />
              </div>
            ))}
            {[{ l: 'Type', k: 'type', opts: ['Challenge', 'Funded', 'Live'] }, { l: 'Currency', k: 'currency', opts: ['USD', 'EUR', 'GBP'] }].map(f => (
              <div key={f.k}>
                <label style={lbl}>{f.l}</label>
                <select style={inp} value={editAcct[f.k]} onChange={e => setEditAcct(p => ({ ...p, [f.k]: e.target.value }))}>{f.opts.map(o => <option key={o}>{o}</option>)}</select>
              </div>
            ))}
            <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
              <button style={{ ...btnG, flex: 1 }} onClick={() => setEditAcct(null)}>Cancel</button>
              <button style={{ ...btnP, flex: 1 }} onClick={saveAccountEdit}>Save changes</button>
            </div>
          </div>
        </>
      )}
    </>
  );

  /* ── SETTINGS ─────────────────────────────────────────────────────────────── */
  const VIS = [
    { k: 'equity', l: 'Equity curve', d: 'Equity growth + year compare' },
    { k: 'pnl', l: 'Net P&L', d: 'Dollar profit & loss figures' },
    { k: 'totalR', l: 'Total R', d: 'Cumulative R-multiple score' },
    { k: 'winRate', l: 'Win rate', d: 'Percentage of winning trades' },
    { k: 'profitFactor', l: 'Profit factor', d: 'Gross profit ÷ gross loss' },
    { k: 'streaks', l: 'Streak metrics', d: 'Best and worst streak data' },
    { k: 'drawdown', l: 'Max drawdown', d: 'Peak-to-trough loss in R' },
    { k: 'tradeHistory', l: 'Trade history', d: 'Full public trade log' },
    { k: 'calendar', l: 'Calendar', d: 'Daily P&L calendar view' },
    { k: 'triangle', l: 'Edge score', d: 'Triangle radar & score bar' },
  ];
  const Settings = () => (
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, alignItems: 'start' }}>
      <div style={card}>
        <SectionLabel accent>Public Visibility</SectionLabel>
        <div style={{ fontSize: 12, color: theme.sub, marginBottom: 14 }}>Toggle what visitors see on your public track record page.</div>
        {VIS.map(item => (
          <div key={item.k} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 0', borderBottom: `1px solid ${theme.border}` }}>
            <div><div style={{ fontSize: 13, fontWeight: 600 }}>{item.l}</div><div style={{ fontSize: 11, color: theme.sub }}>{item.d}</div></div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 10, fontWeight: 600, color: visibility[item.k] ? theme.gain : theme.sub }}>{visibility[item.k] ? '● Public' : '○ Private'}</span>
              <Tog on={visibility[item.k]} fn={() => setVisibility(p => ({ ...p, [item.k]: !p[item.k] }))} />
            </div>
          </div>
        ))}
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
        <div style={card}>
          <SectionLabel accent>Profile</SectionLabel>
          {[{ l: 'Trader name', v: 'Andrej' }, { l: 'Website', v: 'simplemarketsacademy.com' }, { l: 'Public URL', v: 'trackrecord.simplemarketsacademy.com', ro: true }].map(f => (
            <div key={f.l} style={{ marginBottom: 12 }}>
              <label style={lbl}>{f.l}</label>
              <input style={{ ...inp, opacity: f.ro ? .55 : 1 }} defaultValue={f.v} readOnly={!!f.ro} />
            </div>
          ))}
          <div style={{ marginBottom: 12 }}>
            <label style={lbl}>Bio</label>
            <textarea style={{ ...inp, resize: 'vertical' }} rows={2} defaultValue="Simple Markets — professional prop trader & educator." />
          </div>
          <button style={{ ...btnP, width: '100%' }} onClick={() => { alert('Profile saved!'); }}>Save profile</button>
        </div>
        <div style={card}>
          <SectionLabel accent>Excel Bulk Import</SectionLabel>
          <div style={{ fontSize: 12, color: theme.sub, marginBottom: 14 }}>
            Import historical trades from your yearly spreadsheet journals. Reads the <strong style={{ color: theme.text }}>Overall</strong> sheet — columns: PAIR, Date, Profit R, Closing Reason, ENTRY link. BE trades are automatically logged as 0R. Direction is skipped. Select multiple files at once.
          </div>
          {/* File picker */}
          <label style={{ display: 'block', cursor: 'pointer' }}>
            <div style={{ border: `2px dashed ${theme.border}`, borderRadius: 10, padding: '24px 16px', textAlign: 'center', color: theme.sub, transition: 'border-color .15s' }}>
              <div style={{ fontSize: 22, marginBottom: 6 }}>📂</div>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 3, color: theme.text }}>
                {importProgress === 'parsing' ? 'Parsing files…' : 'Click to select .xlsx files'}
              </div>
              <div style={{ fontSize: 11 }}>You can select multiple files (one per year)</div>
            </div>
            <input type="file" accept=".xlsx,.xls" multiple style={{ display: 'none' }}
              onChange={e => {
                if (e.target.files?.length) {
                  handleImportFiles(e.target.files);
                  e.target.value = ''; // reset so same file can be re-selected
                }
              }} />
          </label>

          {/* Preview */}
          {importFiles.length > 0 && (
            <div style={{ marginTop: 14 }}>
              <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, textTransform: 'uppercase', color: theme.sub, marginBottom: 8 }}>Preview</div>
              {importFiles.map((f, fi) => (
                <div key={fi} style={{ marginBottom: 10, border: `1px solid ${theme.border}`, borderRadius: 8, overflow: 'hidden' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 12px', background: theme.muted }}>
                    <span style={{ fontSize: 12, fontWeight: 600 }}>{f.name}</span>
                    <span style={{ fontSize: 11, color: theme.sub, fontFamily: "'JetBrains Mono',monospace" }}>{f.trades.length} trades</span>
                  </div>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 11 }}>
                      <thead>
                        <tr style={{ background: theme.muted }}>
                          {['Date', 'Symbol', 'R', 'Reason', 'Chart'].map(h => (
                            <th key={h} style={{ padding: '5px 10px', textAlign: 'left', color: theme.sub, fontWeight: 700, letterSpacing: 1, fontSize: 9.5, textTransform: 'uppercase' }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {f.trades.slice(0, 5).map((t, ti) => (
                          <tr key={ti}>
                            <td style={{ padding: '5px 10px', fontFamily: "'JetBrains Mono',monospace", color: theme.textDim }}>{t.date}</td>
                            <td style={{ padding: '5px 10px', fontWeight: 600 }}>{t.symbol}</td>
                            <td style={{ padding: '5px 10px', fontFamily: "'JetBrains Mono',monospace", color: t.r > 0 ? theme.gain : t.r < 0 ? theme.loss : theme.sub }}>{t.r > 0 ? '+' : ''}{t.r}R</td>
                            <td style={{ padding: '5px 10px', color: theme.sub }}>{t.notes || '—'}</td>
                            <td style={{ padding: '5px 10px' }}>{t.chartLink ? <a href={t.chartLink} target="_blank" rel="noreferrer" style={{ color: theme.accent, fontSize: 10 }}>View</a> : <span style={{ color: theme.border }}>—</span>}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {f.trades.length > 5 && <div style={{ padding: '6px 10px', fontSize: 10, color: theme.sub }}>…and {f.trades.length - 5} more trades</div>}
                  </div>
                </div>
              ))}

              {/* Account selector */}
              <div style={{ marginBottom: 10 }}>
                <label style={lbl}>Import to account</label>
                <select style={inp} value={importAccount} onChange={e => setImportAccount(e.target.value)}>
                  {accounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
                </select>
              </div>

              {importError && <div style={{ fontSize: 12, color: theme.loss, marginBottom: 8 }}>{importError}</div>}
              {importProgress === 'done' && <div style={{ fontSize: 12, color: theme.gain, marginBottom: 8 }}>✓ Import complete!</div>}

              <div style={{ display: 'flex', gap: 8 }}>
                <button style={{ ...btnG, flex: 1 }} onClick={() => { setImportFiles([]); setImportProgress(null); setImportError(''); }}>Clear</button>
                <button
                  style={{ ...btnP, flex: 2, opacity: importProgress === 'importing' ? 0.6 : 1 }}
                  onClick={doImport}
                  disabled={importProgress === 'importing'}>
                  {importProgress === 'importing'
                    ? 'Importing…'
                    : `Import ${importFiles.reduce((s, f) => s + f.trades.length, 0)} trades`}
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );

  /* ── MONTH POPUP MODAL ──────────────────────────────────────────────────── */
  const MonthModal = () => {
    if (!monthPopup) return null;
    const monthTrades = filtered
      .filter(t => t.date.startsWith(monthPopup))
      .slice().sort((a, b) => a.date.localeCompare(b.date));
    const label = new Date(monthPopup + '-01').toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    const totalR = +monthTrades.reduce((s, t) => s + t.r, 0).toFixed(2);
    return (
      <>
        <div onClick={() => setMonthPopup(null)} className="modal-backdrop" style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,.65)', zIndex: 300, backdropFilter: 'blur(6px)' }} />
        <div className="modal-panel" style={{ position: 'fixed', top: '50%', left: '50%', transform: 'translate(-50%,-50%)', width: 580, maxWidth: 'calc(100vw - 32px)', maxHeight: '82vh', background: theme.card, border: `1px solid ${theme.border}`, borderRadius: 18, zIndex: 301, display: 'flex', flexDirection: 'column', overflow: 'hidden', fontFamily: "'Outfit',sans-serif", boxShadow: '0 24px 80px rgba(0,0,0,0.55)', color: theme.text }}>
          {/* Header */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '18px 22px 16px', borderBottom: `1px solid ${theme.border}`, background: `linear-gradient(135deg, ${theme.cardHi} 0%, ${theme.card} 100%)`, flexShrink: 0 }}>
            <div>
              <div style={{ fontSize: 17, fontWeight: 700, color: theme.text }}>{label}</div>
              <div style={{ fontSize: 11.5, color: theme.sub, marginTop: 3 }}>
                {monthTrades.length} trades ·{' '}
                <span style={{ fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: totalR >= 0 ? theme.gain : theme.loss }}>{fmtR(totalR)}</span>
              </div>
            </div>
            <button onClick={() => setMonthPopup(null)} style={{ ...btnG, padding: '5px 11px', fontSize: 13 }}>✕</button>
          </div>
          {/* Trade cards */}
          <div style={{ overflowY: 'auto', padding: '14px 18px', display: 'flex', flexDirection: 'column', gap: 8 }}>
            {monthTrades.length === 0 && (
              <div style={{ textAlign: 'center', padding: 40, color: theme.sub, fontSize: 14 }}>No trades in {label}.</div>
            )}
            {monthTrades.map(t => {
              const ta = accounts.find(x => x.id === t.account);
              const rp = t.riskPercent != null ? t.riskPercent : (ta?.riskPercent ?? 1.0);
              const init = ta?.initialBalance ?? 25000;
              const pnlDollar = t.r * (rp / 100) * init;
              const result = t.r > 0 ? 'WIN' : t.r < 0 ? 'LOSS' : 'BE';
              const rCol = t.r > 0 ? theme.gain : t.r < 0 ? theme.loss : theme.sub;
              const rBg = t.r > 0 ? theme.gainBg : t.r < 0 ? theme.lossBg : `${theme.sub}18`;
              const rBdr = t.r > 0 ? theme.gainBorder : t.r < 0 ? theme.lossBorder : `${theme.sub}40`;
              return (
                <div key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px', borderRadius: 11, border: `1px solid ${theme.border}`, background: dark ? 'rgba(255,255,255,0.025)' : theme.muted, color: theme.text }}>
                  {/* Symbol + date */}
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                      <span style={{ fontSize: 14, fontWeight: 700, letterSpacing: 0.4, color: theme.text }}>{t.symbol}</span>
                      <DirBadge type={t.type} />
                    </div>
                    <div style={{ fontSize: 10, color: theme.sub, fontFamily: "'JetBrains Mono',monospace" }}>{t.date}</div>
                  </div>
                  {/* Result badge */}
                  <span style={{ fontSize: 10, fontWeight: 700, padding: '3px 10px', borderRadius: 20, background: rBg, color: rCol, border: `1px solid ${rBdr}`, letterSpacing: 1, whiteSpace: 'nowrap' }}>{result}</span>
                  {/* R + $ */}
                  <div style={{ textAlign: 'right', minWidth: 78 }}>
                    <div style={{ fontSize: 14, fontWeight: 700, fontFamily: "'JetBrains Mono',monospace", color: rCol }}>{fmtR(t.r)}</div>
                    <div style={{ fontSize: 11, fontFamily: "'JetBrains Mono',monospace", color: pnlDollar >= 0 ? theme.gain : theme.loss, marginTop: 2 }}>
                      {pnlDollar >= 0 ? '+$' : '-$'}{Math.abs(pnlDollar).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
                    </div>
                  </div>
                  {/* Chart link */}
                  {t.chartLink
                    ? <a href={t.chartLink} target="_blank" rel="noreferrer" style={{ fontSize: 11, fontWeight: 600, color: theme.accent, textDecoration: 'none', whiteSpace: 'nowrap', padding: '5px 11px', borderRadius: 7, border: `1px solid ${theme.accent}44`, background: `${theme.accent}0e`, transition: 'all .12s' }}>Open Chart →</a>
                    : <span style={{ fontSize: 11, color: theme.border, width: 90, display: 'inline-block', textAlign: 'center' }}>—</span>
                  }
                </div>
              );
            })}
          </div>
        </div>
      </>
    );
  };

  /* ── SHELL ──────────────────────────────────────────────────────────────── */
  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;} body{margin:0;}
        select,input,textarea{color-scheme:${dark ? 'dark' : 'light'};}
        input:focus,textarea:focus,select:focus{border-color:${theme.accent}80!important;outline:none;}

        /* ── NAV ── */
        .nav-item{transition:background .15s,color .15s;}
        .nav-item:hover{color:rgba(255,255,255,0.85)!important;background:rgba(255,255,255,0.05)!important;}

        /* ── PAGE TRANSITION ── */
        @keyframes pageEnter{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
        .page-enter{animation:pageEnter .28s cubic-bezier(0.22,1,0.36,1) both;}

        /* ── KPI CARD STAGGER ── */
        @keyframes kpiEnter{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:translateY(0)}}
        .kpi-card{animation:kpiEnter .42s cubic-bezier(0.22,1,0.36,1) both;}

        /* ── KPI VALUE TICK ── */
        @keyframes numFade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}

        /* ── MODAL ANIMATIONS ── */
        @keyframes modalIn{from{opacity:0;transform:translate(-50%,-48%) scale(0.96)}to{opacity:1;transform:translate(-50%,-50%) scale(1)}}
        @keyframes backdropIn{from{opacity:0}to{opacity:1}}
        .modal-backdrop{animation:backdropIn .18s ease both;}
        .modal-panel{animation:modalIn .28s cubic-bezier(0.22,1,0.36,1) both;}

        /* ── BUTTON HOVER ── */
        .btn-primary{transition:transform .1s cubic-bezier(0.22,1,0.36,1),box-shadow .15s ease,opacity .1s;}
        .btn-primary:hover{transform:translateY(-1px);box-shadow:0 6px 18px ${theme.accent}40;}
        .btn-primary:active{transform:scale(0.97);box-shadow:none;}
        .btn-ghost{transition:border-color .12s,color .12s,background .12s;}
        .btn-ghost:hover{border-color:${theme.accent}70!important;color:${theme.text}!important;}

        /* ── ABOUT SECTION ENTRANCES ── */
        @keyframes aboutSection{from{opacity:0;transform:translateY(22px)}to{opacity:1;transform:translateY(0)}}
        .about-section{animation:aboutSection .5s cubic-bezier(0.22,1,0.36,1) both;}

        /* ── TABLE / TRADE ROW ── */
        tr.trade-row:hover td{background:${dark ? 'rgba(255,255,255,0.025)' : 'rgba(0,0,0,0.025)'}!important;}

        /* ── SCROLLBAR ── */
        ::-webkit-scrollbar{width:5px;height:5px;}
        ::-webkit-scrollbar-track{background:transparent;}
        ::-webkit-scrollbar-thumb{background:${theme.border};border-radius:3px;}

        /* ── REDUCED MOTION ── */
        @media(prefers-reduced-motion:reduce){
          .kpi-card,.page-enter,.modal-panel,.about-section{animation:none!important;}
          *{transition-duration:.01ms!important;}
        }

        /* ── MOBILE KPI ── */
        /* 2-column grid handled via inline styles */

        /* ── MOBILE BOTTOM NAV ── */
        .mob-nav{position:fixed;bottom:0;left:0;right:0;z-index:50;display:flex;align-items:stretch;backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);}
        .mob-nav-item{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;padding:10px 4px 12px;cursor:pointer;border:none;background:transparent;font-family:'Outfit',sans-serif;transition:opacity .15s;}
        .mob-nav-item:active{opacity:.6;}

        /* ── MOBILE HEADER ── */
        .mob-header{display:flex;align-items:center;gap:10px;padding:14px 16px 10px;border-bottom:1px solid;}
      `}</style>
      <div style={{ display: 'flex', minHeight: '100vh', background: theme.bg, color: theme.text, fontFamily: "'Outfit',sans-serif", overflowX: 'hidden', maxWidth: '100vw' }}>
        {/* SIDEBAR — hidden on mobile */}
        <div style={{ width: 228, background: theme.nav, display: isMobile ? 'none' : 'flex', flexDirection: 'column', height: '100vh', position: 'fixed', left: 0, top: 0, flexShrink: 0, borderRight: `1px solid ${theme.navBorder}`, zIndex: 20, overflowY: 'auto' }}>
          <div style={{ padding: '22px 22px 18px', borderBottom: `1px solid ${theme.navBorder}` }}>
            <div style={{ fontSize: 9.5, letterSpacing: 2.5, textTransform: 'uppercase', color: theme.accent, marginBottom: 5, fontWeight: 600 }}>Simple Markets</div>
            <div style={{ fontSize: 14, fontWeight: 700, color: '#fff', lineHeight: 1.35, letterSpacing: -0.2 }}>Andrej's<br />Performance Record</div>
          </div>
          <div style={{ padding: '16px 22px 6px', fontSize: 8.5, letterSpacing: 2.5, textTransform: 'uppercase', color: 'rgba(255,255,255,0.2)', fontWeight: 700 }}>Navigation</div>
          <div style={{ flex: 1 }}>
            {PAGES.map(p => (
              <div key={p.id} className="nav-item" onClick={() => setPage(p.id)} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 22px', cursor: 'pointer', fontSize: 13, fontWeight: page === p.id ? 600 : 400, color: page === p.id ? theme.navActive : theme.navText, background: page === p.id ? theme.navActiveBg : 'transparent', borderLeft: `2.5px solid ${page === p.id ? theme.accent : 'transparent'}`, letterSpacing: 0.2 }}>
                <span style={{ fontSize: 14, opacity: page === p.id ? 1 : 0.6 }}>{p.i}</span>{p.l}
              </div>
            ))}
          </div>
          <div style={{ padding: '14px 22px', borderTop: `1px solid ${theme.navBorder}` }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 10 }}>
              <span style={{ width: 6, height: 6, borderRadius: '50%', background: theme.gain, display: 'inline-block', boxShadow: `0 0 6px ${theme.gain}` }} />
              <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.25)', letterSpacing: 0.3 }}>{trades.length.toLocaleString()} trades · {accounts.length} accounts</span>
            </div>
            <button
              onClick={() => { if (isAdmin) { setIsAdmin(false); try { localStorage.removeItem('sm_admin'); } catch { } } else { setShowAdminModal(true); } }}
              style={{ width: '100%', padding: '7px 10px', borderRadius: 8, border: `1px solid ${isAdmin ? theme.accent + '60' : 'rgba(255,255,255,0.08)'}`, background: isAdmin ? `${theme.accent}12` : 'transparent', color: isAdmin ? theme.accent : 'rgba(255,255,255,0.28)', fontFamily: "'Outfit',sans-serif", fontSize: 11, fontWeight: 600, cursor: 'pointer', letterSpacing: 0.4, transition: 'all .15s' }}>
              {isAdmin ? '🔐 Exit Admin Mode' : '🔓 Admin Login'}
            </button>
          </div>
        </div>

        {/* MAIN */}
        <div style={{ flex: 1, minWidth: 0, width: '100%', overflowX: 'hidden', padding: isMobile ? '0 0 80px' : '24px 28px 60px', marginLeft: isMobile ? 0 : 228 }}>

          {/* ── MOBILE HEADER ── */}
          {isMobile ? (
            <div style={{ position: 'sticky', top: 0, zIndex: 30, background: theme.nav, borderBottom: `1px solid ${theme.navBorder}`, padding: '12px 16px', display: 'flex', alignItems: 'center', gap: 10 }}>
              <img src="/sma-logo.png" alt="SMA" style={{ width: 32, height: 32, borderRadius: 8, flexShrink: 0 }} />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: '#fff', letterSpacing: -0.2, lineHeight: 1 }}>
                  {PAGES.find(p => p.id === page)?.l}
                </div>
                <div style={{ fontSize: 10, color: theme.navText, marginTop: 2 }}>Simple Markets Academy</div>
              </div>
              <select style={{ padding: '6px 8px', borderRadius: 8, border: `1px solid ${theme.navBorder}`, background: 'rgba(255,255,255,0.05)', color: '#fff', fontFamily: "'Outfit',sans-serif", fontSize: 12, outline: 'none', maxWidth: 110 }} value={selectedAccount} onChange={e => setSelectedAccount(e.target.value)}>
                <option value="all">All accts</option>
                {accounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
              <button onClick={() => setDark(p => !p)} style={{ width: 32, height: 32, borderRadius: 8, border: `1px solid ${theme.navBorder}`, background: 'rgba(255,255,255,0.05)', cursor: 'pointer', fontSize: 14, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                {dark ? '☀️' : '🌙'}
              </button>
              {isAdmin && page !== 'settings' && (
                <button style={{ ...btnP, padding: '6px 12px', fontSize: 12, flexShrink: 0 }} onClick={() => setLogPanelOpen(true)}>+ Log</button>
              )}
            </div>
          ) : (
            /* ── DESKTOP HEADER ── */
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 22, flexWrap: 'wrap' }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 20, fontWeight: 700 }}>{PAGES.find(p => p.id === page)?.l}</div>
                <div style={{ fontSize: 11, color: theme.sub, marginTop: 1 }}>Last updated {new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}</div>
              </div>
              {['dashboard', 'trades'].includes(page) && (
                <div style={{ display: 'flex', gap: 2, background: theme.muted, border: `1px solid ${theme.border}`, borderRadius: 8, padding: 3 }}>
                  {['1W', '1M', '3M', 'YTD', 'ALL'].map(p => (
                    <button key={p} onClick={() => setPeriod(p)} style={{ padding: '4px 11px', borderRadius: 6, fontSize: 11.5, fontWeight: 600, cursor: 'pointer', color: period === p ? '#fff' : theme.sub, background: period === p ? theme.accent : 'transparent', border: 'none', fontFamily: "'Outfit',sans-serif", transition: 'all .12s' }}>{p}</button>
                  ))}
                </div>
              )}
              <select style={{ padding: '7px 12px', borderRadius: 8, border: `1px solid ${theme.border}`, background: theme.card, color: theme.text, fontFamily: "'Outfit',sans-serif", fontSize: 13, outline: 'none' }} value={selectedAccount} onChange={e => setSelectedAccount(e.target.value)}>
                <option value="all">All accounts</option>
                {accounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
              {['dashboard', 'trades'].includes(page) && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 7, fontSize: 12, color: theme.sub }}>
                  <span>BE=Win</span><Tog on={breakEvenAsWin} fn={() => setBreakEvenAsWin(p => !p)} />
                </div>
              )}
              {isAdmin && page !== 'settings' && <button style={btnP} onClick={() => setLogPanelOpen(true)}>+ Log Trade</button>}
              <button onClick={() => setDark(p => !p)} style={{ width: 34, height: 34, borderRadius: 8, border: `1px solid ${theme.border}`, background: theme.card, cursor: 'pointer', fontSize: 15, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                {dark ? '☀️' : '🌙'}
              </button>
            </div>
          )}

          {/* Mobile period + BE filter bar */}
          {isMobile && ['dashboard', 'trades'].includes(page) && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 16px', borderBottom: `1px solid ${theme.border}`, background: theme.card, overflowX: 'auto', scrollbarWidth: 'none' }}>
              <div style={{ display: 'flex', gap: 2, background: theme.muted, border: `1px solid ${theme.border}`, borderRadius: 8, padding: 3, flexShrink: 0 }}>
                {['1W', '1M', '3M', 'YTD', 'ALL'].map(p => (
                  <button key={p} onClick={() => setPeriod(p)} style={{ padding: '4px 10px', borderRadius: 6, fontSize: 11, fontWeight: 600, cursor: 'pointer', color: period === p ? '#fff' : theme.sub, background: period === p ? theme.accent : 'transparent', border: 'none', fontFamily: "'Outfit',sans-serif", transition: 'all .12s' }}>{p}</button>
                ))}
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: theme.sub, flexShrink: 0 }}>
                <span style={{ whiteSpace: 'nowrap' }}>BE=Win</span>
                <Tog on={breakEvenAsWin} fn={() => setBreakEvenAsWin(p => !p)} />
              </div>
            </div>
          )}

          <div key={page} className="page-enter" style={isMobile ? { padding: '16px 16px 0', width: '100%', boxSizing: 'border-box', overflowX: 'hidden' } : {}}>
            {page === 'dashboard' && Dashboard()}
            {page === 'about' && About()}
            {page === 'trades' && Trades()}
            {page === 'calendar' && Calendar()}
            {page === 'accounts' && Accounts()}
            {page === 'settings' && Settings()}
          </div>
        </div>

        {/* ── MOBILE BOTTOM NAV ── */}
        {isMobile && (
          <nav className="mob-nav" style={{ background: theme.nav, borderTop: `1px solid ${theme.navBorder}` }}>
            {PAGES.map(p => {
              const icons = { dashboard: '▦', about: '◈', trades: '≡', calendar: '◫', accounts: '◎', settings: '⚙' };
              const labels = { dashboard: 'Dashboard', about: 'About', trades: 'Trades', calendar: 'Calendar', accounts: 'Accounts', settings: 'Settings' };
              const active = page === p.id;
              return (
                <button
                  key={p.id}
                  className="mob-nav-item"
                  onClick={() => setPage(p.id)}
                  style={{ color: active ? theme.accent : theme.navText, position: 'relative' }}
                >
                  <span style={{ fontSize: 18, lineHeight: 1 }}>{icons[p.id]}</span>
                  <span style={{ fontSize: 9, fontWeight: active ? 700 : 400, letterSpacing: 0.3, textTransform: 'uppercase' }}>{labels[p.id]}</span>
                  {active && (
                    <div style={{ position: 'absolute', top: 0, left: '50%', transform: 'translateX(-50%)', width: 28, height: 2, borderRadius: 2, background: theme.accent }} />
                  )}
                </button>
              );
            })}
          </nav>
        )}
      </div>
      {logPanelOpen && LogPanel()}
      {showAdminModal && AdminModal()}
      {MonthModal()}
    </>
  );
}