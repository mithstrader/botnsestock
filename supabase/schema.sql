-- ══════════════════════════════════════════════════════════════
--  Nifty Precision Bot — Supabase Schema
--  Run this entire file in: Supabase → SQL Editor → New Query
-- ══════════════════════════════════════════════════════════════

-- ── 1. signals ─────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS signals (
  id            BIGSERIAL PRIMARY KEY,
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  signal_date   DATE        NOT NULL DEFAULT CURRENT_DATE,
  signal_time   TEXT        NOT NULL DEFAULT '',
  instrument    TEXT        NOT NULL,
  direction     TEXT        NOT NULL,   -- CALL | PUT
  rating        TEXT        NOT NULL,   -- ULTRA STRONG | STRONG | MEDIUM | WEAK
  score         INT         NOT NULL,
  ltp           NUMERIC(10,2),
  strike        INT,
  option_type   TEXT,                   -- CE | PE
  sl_level      NUMERIC(10,2),
  target_level  NUMERIC(10,2),
  sl_pts        NUMERIC(8,2),
  target_pts    NUMERIC(8,2),
  rsi           NUMERIC(6,2),
  vwap          NUMERIC(10,2),
  ema9          NUMERIC(10,2),
  ema21         NUMERIC(10,2),
  supertrend    INT,                    -- 1 | -1
  vix           NUMERIC(6,2),
  pcr           NUMERIC(6,3),
  fib_level     TEXT,
  fib_price     NUMERIC(10,2),
  fib_at_level  BOOLEAN DEFAULT FALSE,
  fib_bonus     INT     DEFAULT 0,
  orb_high      NUMERIC(10,2),
  orb_low       NUMERIC(10,2),
  pdh           NUMERIC(10,2),
  pdl           NUMERIC(10,2),
  cpr_top       NUMERIC(10,2),
  cpr_bottom    NUMERIC(10,2),
  cpr_narrow    BOOLEAN DEFAULT FALSE,
  ema_ok        BOOLEAN DEFAULT FALSE,
  vwap_ok       BOOLEAN DEFAULT FALSE,
  orb_ok        BOOLEAN DEFAULT FALSE,
  cpr_ok        BOOLEAN DEFAULT FALSE,
  pdh_ok        BOOLEAN DEFAULT FALSE,
  rsi_ok        BOOLEAN DEFAULT FALSE
);

-- ── 2. scans ───────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS scans (
  id               BIGSERIAL PRIMARY KEY,
  created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  scan_date        DATE        NOT NULL DEFAULT CURRENT_DATE,
  scan_time        TEXT        NOT NULL DEFAULT '',
  scan_number      INT         NOT NULL DEFAULT 0,
  vix              NUMERIC(6,2),
  pcr              NUMERIC(6,3),
  nifty50_score    INT,
  nifty50_dir      TEXT,
  nifty50_rsi      NUMERIC(6,2),
  nifty50_ltp      NUMERIC(10,2),
  banknifty_score  INT,
  banknifty_dir    TEXT,
  banknifty_rsi    NUMERIC(6,2),
  banknifty_ltp    NUMERIC(10,2),
  sensex_score     INT,
  sensex_dir       TEXT,
  sensex_rsi       NUMERIC(6,2),
  sensex_ltp       NUMERIC(10,2),
  signals_count    INT DEFAULT 0,
  footprint_data   JSONB         -- per-candle volume delta for footprint chart
);

-- If table already exists, add the column:
ALTER TABLE scans ADD COLUMN IF NOT EXISTS footprint_data JSONB;

-- ── 3. bot_status (single-row upsert) ──────────────────────────
CREATE TABLE IF NOT EXISTS bot_status (
  id               INT PRIMARY KEY DEFAULT 1,
  updated_at       TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  status           TEXT        NOT NULL DEFAULT 'stopped',  -- running | sleeping | stopped
  last_scan_at     TIMESTAMPTZ,
  scan_count       INT     DEFAULT 0,
  alerts_today     INT     DEFAULT 0,
  session_date     DATE,
  next_session_at  TIMESTAMPTZ,
  vix              NUMERIC(6,2),
  pcr              NUMERIC(6,3)
);

INSERT INTO bot_status (id) VALUES (1) ON CONFLICT (id) DO NOTHING;

-- ── 4. Row Level Security — public READ, bot writes via service key ─
ALTER TABLE signals    ENABLE ROW LEVEL SECURITY;
ALTER TABLE scans      ENABLE ROW LEVEL SECURITY;
ALTER TABLE bot_status ENABLE ROW LEVEL SECURITY;

-- Allow anyone to read (anon key used by GitHub Pages)
CREATE POLICY "anon_read_signals"    ON signals    FOR SELECT TO anon USING (true);
CREATE POLICY "anon_read_scans"      ON scans      FOR SELECT TO anon USING (true);
CREATE POLICY "anon_read_bot_status" ON bot_status FOR SELECT TO anon USING (true);

-- ── 5. Enable real-time on these tables ────────────────────────
-- Run these in SQL Editor as well:
ALTER PUBLICATION supabase_realtime ADD TABLE signals;
ALTER PUBLICATION supabase_realtime ADD TABLE scans;
ALTER PUBLICATION supabase_realtime ADD TABLE bot_status;

-- ── 6. Helpful indexes ─────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_signals_date       ON signals (signal_date DESC);
CREATE INDEX IF NOT EXISTS idx_signals_instrument ON signals (instrument);
CREATE INDEX IF NOT EXISTS idx_scans_date         ON scans   (scan_date DESC);
