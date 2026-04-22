# Andrej's Performance Record — Simple Markets

Live trading track record. Built with React + Vite + Supabase + Recharts.

---

## Local Setup

```bash
# 1. Install dependencies
npm install

# 2. Add your environment variables (already done if you got this from Claude)
# .env.local is already configured with your Supabase keys

# 3. Run the dev server
npm run dev
# → opens at http://localhost:5173
```

## Supabase Setup

1. Go to your Supabase project → **SQL Editor** → **New Query**
2. Paste the entire contents of `schema.sql` and click **Run**
3. Go to **Table Editor** to confirm `accounts` and `trades` tables exist

## Add Your First Accounts

In the app, go to **Accounts → + Add Account** and add your real accounts  
(Live, Funded, Challenge) with the correct initial balance and risk per trade.

Or uncomment and edit the seed section at the bottom of `schema.sql` and run it.

## Deploy to Vercel

```bash
# Push to GitHub first
git init && git add . && git commit -m "init"
git remote add origin https://github.com/YOUR_USERNAME/trackrecord.git
git push -u origin main
```

Then in Vercel:
1. Import the GitHub repo
2. Add environment variables:
   - `VITE_SUPABASE_URL` = your project URL
   - `VITE_SUPABASE_ANON_KEY` = your anon key
3. Deploy — Vercel auto-detects Vite

## Domain Setup (trackrecord.simplemarketsacademy.com)

In Vercel → Project → Settings → Domains → Add `trackrecord.simplemarketsacademy.com`  
Then add a CNAME record in your DNS provider pointing to `cname.vercel-dns.com`

---

## Project Structure

```
trackrecord/
├── src/
│   ├── App.jsx          ← Full UI + all logic
│   ├── main.jsx         ← React entry point
│   └── lib/
│       └── supabase.js  ← DB client + all API helpers
├── schema.sql           ← Run this in Supabase SQL Editor
├── .env.local           ← Your Supabase credentials (git-ignored)
├── .env.example         ← Template for other environments
├── index.html
├── vite.config.js
└── package.json
```
