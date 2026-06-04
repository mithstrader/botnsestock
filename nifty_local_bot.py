#!/usr/bin/env python3
"""
Nifty Precision Bot — LOCAL VERSION (run from your PC)
═══════════════════════════════════════════════════════
Run this on your Windows/Mac/Linux machine during market hours.
Data fetched directly from Yahoo Finance (works from India).

SETUP:
  1. pip install yfinance pandas numpy openpyxl requests schedule pytz plyer
  2. Fill TELEGRAM_TOKEN and CHAT_ID below
  3. Double-click run_bot.bat  (Windows)  OR  python nifty_local_bot.py

Instruments : Nifty50 | BankNifty | Sensex
Strategy    : ORB + VWAP + EMA9/21 + SuperTrend + RSI + CPR + PDH/PDL + Fib + VIX/PCR
R:R         : 1:3
"""

# ═══════════════════════════════════════════════════════════════
#  ⚙️  FILL THESE IN — that's all you need to change
# ═══════════════════════════════════════════════════════════════

# TELEGRAM_TOKEN = "YOUR_BOT_TOKEN_HERE"       # e.g. "7xxxxxxxxx:AAF..."
# CHAT_ID        = "YOUR_CHAT_ID_HERE"          # e.g. "123456789"
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "8794792642:AAE4r1BEYAyytGl_IXfZ8XwkebW7omTomGs")
CHAT_ID        = os.environ.get("CHAT_ID",        "670433968")

# ── Supabase (cloud backend) ──────────────────────────────────
# Get these from: Supabase → Project Settings → API
# Set as environment variables on Railway, or paste directly for local use.
SUPABASE_URL         = os.environ.get("SUPABASE_URL",         "")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")  # service_role key

# ── Excel (local logging, optional) ──────────────────────────
_default_excel = r"C:\NSE_Bot\logs" if platform.system() == "Windows" else "/tmp/nse_bot_logs"
EXCEL_FOLDER   = os.environ.get("EXCEL_FOLDER", _default_excel)

# ═══════════════════════════════════════════════════════════════
#  OPTIONAL SETTINGS
# ═══════════════════════════════════════════════════════════════

SCORE_ULTRA     = 85
SCORE_STRONG    = 75
SCORE_MEDIUM    = 65
SCAN_INTERVAL   = 5       # minutes between scans
MAX_TRADES_DAY  = 6
HEARTBEAT_SCANS = 12      # send alive ping every N scans (~60 min)
RESIGAL_GAP_MIN = 30      # minutes before re-alerting same instrument

INSTRUMENTS = {
    "NIFTY50":   {"ticker": "^NSEI",    "sl_pts": 40,  "lot_size": 25, "strike_step": 50,  "min_score": 75},
    "BANKNIFTY": {"ticker": "^NSEBANK", "sl_pts": 120, "lot_size": 15, "strike_step": 100, "min_score": 75},
    "SENSEX":    {"ticker": "^BSESN",   "sl_pts": 200, "lot_size": 10, "strike_step": 100, "min_score": 75},
}

RSI_PERIOD       = 14
RSI_CALL_LOW, RSI_CALL_HIGH = 40, 65
RSI_PUT_LOW,  RSI_PUT_HIGH  = 35, 60
CPR_NARROW_PCT   = 0.20
ORB_MINUTES      = 15
ST_PERIOD, ST_MULT = 10, 3.0
FIB_TOLERANCE_PCT = 0.15
FIB_RATIOS        = [0.236, 0.382, 0.500, 0.618, 0.786]
FIB_SCORE_MAP     = {0.618: 12, 0.500: 8, 0.382: 6, 0.786: 4, 0.236: 3}

VIX_EXTREME, VIX_NORMAL_MAX, VIX_CALM_MAX = 25.0, 20.0, 12.0
VIX_HIGH_PENALT, VIX_CALM_BONUS = 10, 5
PCR_VERY_BULL, PCR_BULL   = 1.5, 1.2
PCR_BEAR,      PCR_VERY_BEAR = 0.8, 0.6

MARKET_OPEN  = (9, 15)
MARKET_CLOSE = (15, 30)
BOT_START    = (9, 35)
LATE_SESSION = (14, 30)

# ═══════════════════════════════════════════════════════════════
#  IMPORTS
# ═══════════════════════════════════════════════════════════════

import os, sys, time, platform, subprocess, threading
from datetime import datetime, timedelta
from typing import Optional

import pytz, schedule, requests

try:
    from supabase import create_client as _sb_create
    _sb_available = True
except ImportError:
    _sb_available = False

IST = pytz.timezone("Asia/Kolkata")

try:
    import yfinance as yf
    import numpy as np
    import pandas as pd
    _yf = True
except ImportError:
    _yf = False
    print("❌ Missing: pip install yfinance pandas numpy")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    _openpyxl = True
except ImportError:
    _openpyxl = False
    print("⚠ openpyxl not found — Excel logging disabled. pip install openpyxl")

_OS = platform.system()

# ═══════════════════════════════════════════════════════════════
#  STATE
# ═══════════════════════════════════════════════════════════════

state = {
    "scan_count":     0,
    "alerts_sent":    0,
    "trades_today":   0,
    "vix":            None,
    "pcr":            None,
    "last_vix_fetch": None,
    "last_signal":    {},
    "data_errors":    {},
    "scores_seen":    [],
}

# ═══════════════════════════════════════════════════════════════
#  INDICATORS
# ═══════════════════════════════════════════════════════════════

def calc_ema(series, period):
    return series.ewm(span=period, adjust=False).mean()

def calc_rsi(series, period=14):
    delta = series.diff()
    gain  = delta.clip(lower=0).ewm(com=period-1, adjust=False).mean()
    loss  = (-delta.clip(upper=0)).ewm(com=period-1, adjust=False).mean()
    rs    = gain / loss.replace(0, 1e-10)
    return float((100 - 100/(1+rs)).iloc[-1])

def calc_vwap(df):
    tp  = (df["High"] + df["Low"] + df["Close"]) / 3
    vol = df["Volume"]
    if vol.sum() == 0:
        return float(tp.iloc[-1])
    return float((tp * vol).cumsum().iloc[-1] / vol.cumsum().iloc[-1])

def calc_atr(df, period=14):
    hl = df["High"] - df["Low"]
    hc = (df["High"] - df["Close"].shift()).abs()
    lc = (df["Low"]  - df["Close"].shift()).abs()
    return pd.concat([hl, hc, lc], axis=1).max(axis=1).ewm(com=period-1, adjust=False).mean()

def calc_supertrend(df, period=10, mult=3.0):
    if len(df) < period + 2:
        return 1 if float(df["Close"].iloc[-1]) > float(df["Close"].iloc[0]) else -1
    hl2 = (df["High"] + df["Low"]) / 2
    atr_ = calc_atr(df, period)
    upper_r = hl2 + mult * atr_
    lower_r = hl2 - mult * atr_
    upper = upper_r.copy()
    lower = lower_r.copy()
    trend = pd.Series(1, index=df.index, dtype=float)
    for i in range(1, len(df)):
        upper.iloc[i] = upper_r.iloc[i] if upper_r.iloc[i] < upper.iloc[i-1] or df["Close"].iloc[i-1] > upper.iloc[i-1] else upper.iloc[i-1]
        lower.iloc[i] = lower_r.iloc[i] if lower_r.iloc[i] > lower.iloc[i-1] or df["Close"].iloc[i-1] < lower.iloc[i-1] else lower.iloc[i-1]
        if   trend.iloc[i-1] == -1 and df["Close"].iloc[i] > upper.iloc[i]: trend.iloc[i] = 1
        elif trend.iloc[i-1] ==  1 and df["Close"].iloc[i] < lower.iloc[i]: trend.iloc[i] = -1
        else: trend.iloc[i] = trend.iloc[i-1]
    return int(trend.iloc[-1])

def calc_cpr(ph, pl, pc):
    pivot = (ph + pl + pc) / 3
    bc    = (ph + pl) / 2
    tc    = 2 * pivot - bc
    w     = abs(tc - bc) / pivot * 100
    return {"pivot": pivot, "top": max(tc,bc), "bottom": min(tc,bc),
            "width_pct": w, "narrow": w < CPR_NARROW_PCT}

def calc_fib(df_today, direction, orb_h, orb_l):
    empty = {"level":None,"price":None,"dist_pct":None,"at_fib":False,
             "score_bonus":0,"levels":{},"swing_high":None,"swing_low":None}
    if df_today is None or len(df_today) < 3:
        return empty
    ltp = float(df_today["Close"].iloc[-1])
    if direction == "CALL":
        sl, sh = orb_l, float(df_today["High"].max())
        move = sh - sl
        if move < 1: return empty
        raw = {r: sh - move * r for r in FIB_RATIOS}
    else:
        sh, sl = orb_h, float(df_today["Low"].min())
        move = sh - sl
        if move < 1: return empty
        raw = {r: sl + move * r for r in FIB_RATIOS}
    nearest = min(raw, key=lambda r: abs(raw[r]-ltp))
    dist_pct = abs(ltp - raw[nearest]) / ltp * 100
    at_fib   = dist_pct <= FIB_TOLERANCE_PCT
    return {
        "level":       f"{nearest:.3f}",
        "price":       round(raw[nearest], 1),
        "dist_pct":    round(dist_pct, 3),
        "at_fib":      at_fib,
        "score_bonus": FIB_SCORE_MAP.get(nearest, 0) if at_fib else 0,
        "levels":      {f"{r:.3f}": round(p, 1) for r, p in raw.items()},
        "swing_high":  round(sh, 1),
        "swing_low":   round(sl, 1),
    }

# ═══════════════════════════════════════════════════════════════
#  SUPABASE CLIENT + PUSH FUNCTIONS
# ═══════════════════════════════════════════════════════════════

_sb_client = None

def get_sb():
    global _sb_client
    if _sb_client is None and _sb_available and SUPABASE_URL and SUPABASE_SERVICE_KEY:
        try:
            _sb_client = _sb_create(SUPABASE_URL, SUPABASE_SERVICE_KEY)
        except Exception as e:
            print(f"  [SUPABASE] init error: {e}")
    return _sb_client

def sb_push_signal(sig):
    """Push a triggered signal to Supabase signals table."""
    client = get_sb()
    if client is None: return
    try:
        now = datetime.now(IST)
        client.table("signals").insert({
            "signal_date":  now.strftime("%Y-%m-%d"),
            "signal_time":  sig["time"],
            "instrument":   sig["instrument"],
            "direction":    sig["direction"],
            "rating":       sig["rating"],
            "score":        sig["score"],
            "ltp":          round(float(sig["ltp"]), 2),
            "strike":       int(sig["strike"]),
            "option_type":  sig["option_type"],
            "sl_level":     round(float(sig["sl_level"]), 2),
            "target_level": round(float(sig["target_level"]), 2),
            "sl_pts":       round(float(sig["sl_pts"]), 2),
            "target_pts":   round(float(sig["target_pts"]), 2),
            "rsi":          round(float(sig["rsi"]), 2),
            "vwap":         round(float(sig["vwap"]), 2),
            "ema9":         round(float(sig["e9"]), 2),
            "ema21":        round(float(sig["e21"]), 2),
            "supertrend":   int(sig["st_dir"]),
            "vix":          float(state["vix"]) if state["vix"] else None,
            "pcr":          float(state["pcr"]) if state["pcr"] else None,
            "fib_level":    sig["fib"]["level"],
            "fib_price":    sig["fib"]["price"],
            "fib_at_level": bool(sig["fib"]["at_fib"]),
            "fib_bonus":    int(sig["fib"]["score_bonus"]),
            "orb_high":     round(float(sig["orb_high"]), 2),
            "orb_low":      round(float(sig["orb_low"]), 2),
            "pdh":          round(float(sig["pdh"]), 2),
            "pdl":          round(float(sig["pdl"]), 2),
            "cpr_top":      round(float(sig["cpr"]["top"]), 2),
            "cpr_bottom":   round(float(sig["cpr"]["bottom"]), 2),
            "cpr_narrow":   bool(sig["cpr"]["narrow"]),
            "ema_ok":       bool(sig["conditions"].get("EMA 9/21")),
            "vwap_ok":      bool(sig["conditions"].get("VWAP")),
            "orb_ok":       bool(sig["conditions"].get("ORB")),
            "cpr_ok":       bool(sig["conditions"].get("CPR")),
            "pdh_ok":       bool(sig["conditions"].get("PDH/PDL")),
            "rsi_ok":       bool(sig["conditions"].get("RSI")),
        }).execute()
        print(f"  [SUPABASE] ✅ signal pushed for {sig['instrument']}")
    except Exception as e:
        print(f"  [SUPABASE] push_signal error: {e}")

def sb_push_scan(scan_no, results):
    """Push per-scan summary to Supabase scans table."""
    client = get_sb()
    if client is None: return
    try:
        now = datetime.now(IST)
        row = {
            "scan_date":     now.strftime("%Y-%m-%d"),
            "scan_time":     now.strftime("%H:%M"),
            "scan_number":   scan_no,
            "vix":           float(state["vix"]) if state["vix"] else None,
            "pcr":           float(state["pcr"]) if state["pcr"] else None,
            "signals_count": sum(1 for _, _, _, _, alerted in results if alerted),
        }
        for name, sig, _, _, _ in results:
            k = name.lower()
            if sig:
                row[f"{k}_score"] = int(sig["score"])
                row[f"{k}_dir"]   = sig["direction"]
                row[f"{k}_rsi"]   = round(float(sig["rsi"]), 2)
                row[f"{k}_ltp"]   = round(float(sig["ltp"]), 2)
        client.table("scans").insert(row).execute()
    except Exception as e:
        print(f"  [SUPABASE] push_scan error: {e}")

def sb_update_status(status="running"):
    """Update the single bot_status row."""
    client = get_sb()
    if client is None: return
    try:
        now = datetime.now(IST)
        row = {
            "id":           1,
            "updated_at":   now.isoformat(),
            "status":       status,
            "scan_count":   state["scan_count"],
            "alerts_today": state["alerts_sent"],
            "session_date": now.strftime("%Y-%m-%d"),
            "vix":          float(state["vix"]) if state["vix"] else None,
            "pcr":          float(state["pcr"]) if state["pcr"] else None,
        }
        if status == "running":
            row["last_scan_at"] = now.isoformat()
        if status == "sleeping":
            row["next_session_at"] = _next_market_open().isoformat()
        client.table("bot_status").upsert(row).execute()
    except Exception as e:
        print(f"  [SUPABASE] update_status error: {e}")

# ═══════════════════════════════════════════════════════════════
#  DATA FETCHING
# ═══════════════════════════════════════════════════════════════

_NSE_HDR = {
    "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept":          "application/json",
    "Referer":         "https://www.nseindia.com/",
}
_nse_sess = None

def _nse_session():
    global _nse_sess
    if _nse_sess is None:
        _nse_sess = requests.Session()
        _nse_sess.headers.update(_NSE_HDR)
        try: _nse_sess.get("https://www.nseindia.com", timeout=10)
        except: pass
    return _nse_sess

def fetch_vix_pcr():
    vix_val = pcr_val = None
    try:
        r = _nse_session().get("https://www.nseindia.com/api/allIndices", timeout=10)
        if r.status_code == 200:
            for item in r.json().get("data", []):
                if item.get("index") == "INDIA VIX":
                    vix_val = float(item["last"]); break
    except Exception as e:
        print(f"  [VIX] {e}")
    try:
        r = _nse_session().get(
            "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY", timeout=15)
        if r.status_code == 200:
            data  = r.json()["records"]["data"]
            ce_oi = sum(d["CE"]["openInterest"] for d in data if "CE" in d)
            pe_oi = sum(d["PE"]["openInterest"] for d in data if "PE" in d)
            pcr_val = round(pe_oi / ce_oi, 3) if ce_oi else None
    except Exception as e:
        print(f"  [PCR] {e}")
    return vix_val, pcr_val

def _normalise_df(df) -> Optional[pd.DataFrame]:
    """Flatten MultiIndex columns, convert index to IST, drop NaNs."""
    if df is None or df.empty:
        return None
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    # Rename 'Price' level leftover from some yf versions
    df.columns = [c.strip().capitalize() for c in df.columns]
    try:
        if df.index.tz is None:
            df.index = df.index.tz_localize("UTC").tz_convert(IST)
        else:
            df.index = df.index.tz_convert(IST)
    except Exception:
        pass
    df = df.dropna(subset=["Close", "High", "Low"])
    return df if len(df) >= 20 else None


def fetch_ohlcv(ticker: str) -> Optional[pd.DataFrame]:
    """Fetch 5-min OHLCV.  Tries Ticker.history() first (more robust),
    then falls back to yf.download()."""

    # ── Method 1: Ticker.history() ──────────────────────────────
    for period in ("5d", "7d", "10d"):
        try:
            tk  = yf.Ticker(ticker)
            df  = tk.history(period=period, interval="5m", auto_adjust=True)
            out = _normalise_df(df)
            if out is not None:
                return out
        except Exception as e:
            print(f"  [DATA] {ticker} Ticker.history period={period}: {e}")

    # ── Method 2: yf.download() fallback ────────────────────────
    for period in ("5d", "7d", "1mo"):
        try:
            df  = yf.download(ticker, period=period, interval="5m",
                              progress=False, auto_adjust=True,
                              threads=False)
            out = _normalise_df(df)
            if out is not None:
                return out
        except Exception as e:
            print(f"  [DATA] {ticker} download period={period}: {e}")

    return None

def split_sessions(df):
    today = datetime.now(IST).date()
    td = df[df.index.date == today].copy()
    prev_days = sorted(set(df.index.date))
    prev_days = [d for d in prev_days if d < today]
    prev = df[df.index.date == prev_days[-1]].copy() if prev_days else pd.DataFrame()
    return td, prev

def get_orb(df_today):
    if df_today.empty: return 0.0, 0.0
    end = df_today.index[0].replace(
        hour=MARKET_OPEN[0], minute=MARKET_OPEN[1]+ORB_MINUTES, second=0)
    orb = df_today[df_today.index < end]
    if orb.empty: orb = df_today.iloc[:3]
    return float(orb["High"].max()), float(orb["Low"].min())

# ═══════════════════════════════════════════════════════════════
#  SIGNAL ANALYSIS
# ═══════════════════════════════════════════════════════════════

def analyze(name, cfg):
    df = fetch_ohlcv(cfg["ticker"])
    if df is None or len(df) < 20:
        _data_error(name, f"No data from yfinance for {cfg['ticker']}")
        return None

    df_today, df_prev = split_sessions(df)
    if df_today.empty or len(df_today) < 5:
        _data_error(name, "No today's candles in data")
        return None
    if df_prev.empty:
        _data_error(name, "No previous day data (needed for CPR/PDH/PDL)")
        return None

    state["data_errors"][name] = 0

    pdh = float(df_prev["High"].max())
    pdl = float(df_prev["Low"].min())
    pdc = float(df_prev["Close"].iloc[-1])
    cpr = calc_cpr(pdh, pdl, pdc)

    close = df_today["Close"]
    ltp   = float(close.iloc[-1])
    e9    = float(calc_ema(close, 9).iloc[-1])
    e21   = float(calc_ema(close, 21).iloc[-1])
    vwap_ = calc_vwap(df_today)
    rsi_  = calc_rsi(close, RSI_PERIOD)
    st_   = calc_supertrend(df_today, ST_PERIOD, ST_MULT)
    orb_h, orb_l = get_orb(df_today)

    bull = sum([e9>e21, ltp>vwap_, st_==1, ltp>orb_h])
    bear = sum([e9<e21, ltp<vwap_, st_==-1, ltp<orb_l])
    direction   = "CALL" if bull >= bear else "PUT"
    option_type = "CE"   if direction == "CALL" else "PE"

    fib = calc_fib(df_today, direction, orb_h, orb_l)

    score, cond = 0, {}

    ema_ok  = (e9>e21 and direction=="CALL") or (e9<e21 and direction=="PUT")
    cond["EMA 9/21"] = ema_ok;  score += 15 if ema_ok else 0

    vwap_ok = (ltp>vwap_ and direction=="CALL") or (ltp<vwap_ and direction=="PUT")
    cond["VWAP"] = vwap_ok;     score += 15 if vwap_ok else 0

    st_ok  = (st_==1 and direction=="CALL") or (st_==-1 and direction=="PUT")
    cond["SuperTrend"] = st_ok; score += 15 if st_ok else 0

    orb_ok = (ltp>orb_h and direction=="CALL") or (ltp<orb_l and direction=="PUT")
    near_orb = abs(ltp-(orb_h if direction=="CALL" else orb_l))/ltp < 0.0005
    cond["ORB"] = orb_ok;       score += 20 if orb_ok else (8 if near_orb else 0)

    rsi_ok = (RSI_CALL_LOW<=rsi_<=RSI_CALL_HIGH) if direction=="CALL" else (RSI_PUT_LOW<=rsi_<=RSI_PUT_HIGH)
    cond["RSI"] = rsi_ok;       score += 10 if rsi_ok else 0

    cpr_ok = (ltp>cpr["top"] and cpr["narrow"]) if direction=="CALL" else (ltp<cpr["bottom"] and cpr["narrow"])
    cond["CPR"] = cpr_ok;       score += 10 if cpr_ok else (4 if cpr["narrow"] else 0)

    pdhl_ok = (ltp>pdh and direction=="CALL") or (ltp<pdl and direction=="PUT")
    cond["PDH/PDL"] = pdhl_ok;  score += 10 if pdhl_ok else 0

    if fib["at_fib"]:
        cond["Fib Pullback"] = True
        fib_label = f"0.{fib['level'].split('.')[1]} ({fib['price']:,.1f}) ✓ +{fib['score_bonus']}pts"
        score += fib["score_bonus"]
    else:
        cond["Fib Pullback"] = False
        fib_label = f"nearest {fib['level']} ({fib['price']:,.1f}) dist {fib['dist_pct']:.2f}%" if fib["level"] else "N/A"

    vix_v = state["vix"]
    if vix_v:
        if   vix_v > VIX_EXTREME:    score -= 30; cond["VIX"] = f"EXTREME {vix_v:.1f} ⛔"
        elif vix_v > VIX_NORMAL_MAX: score -= VIX_HIGH_PENALT; cond["VIX"] = f"HIGH {vix_v:.1f} ⚠"
        elif vix_v < VIX_CALM_MAX:   score += VIX_CALM_BONUS;  cond["VIX"] = f"CALM {vix_v:.1f} ✓"
        else:                          cond["VIX"] = f"NORMAL {vix_v:.1f}"

    pcr_v = state["pcr"]
    if pcr_v:
        if direction == "CALL":
            bonus = 5 if pcr_v>=PCR_VERY_BULL else (3 if pcr_v>=PCR_BULL else (-5 if pcr_v<=PCR_VERY_BEAR else 0))
        else:
            bonus = 5 if pcr_v<=PCR_VERY_BEAR else (3 if pcr_v<=PCR_BEAR else (-5 if pcr_v>=PCR_VERY_BULL else 0))
        score += bonus
        cond["PCR"] = f"{pcr_v:.2f}"

    score = max(0, min(100, score))

    now_ist = datetime.now(IST)
    late = now_ist.hour > LATE_SESSION[0] or (now_ist.hour == LATE_SESSION[0] and now_ist.minute >= LATE_SESSION[1])
    if late and score < SCORE_ULTRA:
        score = max(0, score - 10)
        cond["Session"] = "LATE ⚠ (-10)"

    conds_met = sum(1 for v in cond.values() if isinstance(v, bool) and v)

    if   score >= SCORE_ULTRA:  rating, rr = "ULTRA STRONG", "1:3"
    elif score >= SCORE_STRONG: rating, rr = "STRONG",       "1:3"
    elif score >= SCORE_MEDIUM: rating, rr = "MEDIUM",       "1:2"
    else:                       rating, rr = "WEAK",         "SKIP"

    sl_pts     = cfg["sl_pts"]
    target_pts = sl_pts * 3
    strike     = round(ltp / cfg["strike_step"]) * cfg["strike_step"]
    if direction == "CALL":
        sl_level, target_level = round(ltp-sl_pts,1), round(ltp+target_pts,1)
    else:
        sl_level, target_level = round(ltp+sl_pts,1), round(ltp-target_pts,1)

    return {
        "instrument": name, "ticker": cfg["ticker"],
        "direction": direction, "option_type": option_type,
        "ltp": ltp, "strike": strike, "sl_level": sl_level, "sl_pts": sl_pts,
        "target_level": target_level, "target_pts": target_pts,
        "score": score, "rating": rating, "rr": rr, "conds_met": conds_met,
        "conditions": cond, "e9": e9, "e21": e21, "vwap": vwap_, "rsi": rsi_,
        "st_dir": st_, "orb_high": orb_h, "orb_low": orb_l,
        "pdh": pdh, "pdl": pdl, "cpr": cpr, "fib": fib, "fib_label": fib_label,
        "time": now_ist.strftime("%H:%M"),
    }

# ═══════════════════════════════════════════════════════════════
#  NOTIFICATIONS
# ═══════════════════════════════════════════════════════════════

def send_telegram(msg: str):
    if not TELEGRAM_TOKEN or TELEGRAM_TOKEN == "YOUR_BOT_TOKEN_HERE":
        print(f"[TG - not configured] {msg[:80]}")
        return
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            data={"chat_id": CHAT_ID, "text": msg,
                  "parse_mode": "Markdown", "disable_web_page_preview": True},
            timeout=10)
        if r.status_code != 200:
            print(f"[TG] Error {r.status_code}: {r.text[:100]}")
    except Exception as e:
        print(f"[TG] {e}")

def _beep():
    try:
        if _OS == "Windows":
            import winsound; winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        elif _OS == "Darwin":
            subprocess.Popen(["afplay", "/System/Library/Sounds/Glass.aiff"],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except: pass

def _popup(title, msg, urgent=False):
    import tkinter as tk
    bg, accent = "#1e1e2e", "#e05c5c" if urgent else "#1D9E75"
    def _show():
        root = tk.Tk(); root.overrideredirect(True)
        root.attributes("-topmost", True); root.configure(bg=bg)
        w, h = 430, 115
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry(f"{w}x{h}+{sw-w-20}+{sh-h-60}")
        tk.Frame(root, bg=accent, height=4).pack(fill="x")
        tk.Label(root, text=title, font=("Segoe UI",11,"bold"),
                 fg="#fff", bg=bg, anchor="w", padx=12).pack(fill="x", pady=(6,0))
        tk.Label(root, text=msg, font=("Segoe UI",9), fg="#aaa", bg=bg,
                 anchor="w", padx=12, wraplength=410).pack(fill="x")
        tk.Button(root, text="✕", font=("Segoe UI",8), fg="#666", bg=bg,
                  bd=0, command=root.destroy).place(x=w-24, y=6)
        root.after(12000, root.destroy); root.mainloop()
    threading.Thread(target=_show, daemon=True).start()

def send_desktop(title, msg, urgent=False):
    _beep()
    if _OS == "Windows":
        try: _popup(title, msg, urgent); return
        except: pass
    print(f"\n{'='*55}\n  {title}\n  {msg}\n{'='*55}\n")

def startup_ping():
    now_str = datetime.now(IST).strftime("%d %b %Y %H:%M IST")
    send_telegram(
        f"🚀 *Nifty Precision Bot — Started*\n"
        f"⏰ {now_str}\n"
        f"🎯 Scanning: {', '.join(INSTRUMENTS)}\n"
        f"🔔 Alert threshold: score ≥ {SCORE_STRONG}\n"
        f"⏱ Every {SCAN_INTERVAL} min | 9:35 AM – 3:30 PM IST"
    )
    sb_update_status("running")

def heartbeat_ping():
    now_str = datetime.now(IST).strftime("%H:%M IST")
    recent  = state["scores_seen"][-6:]
    avg_sc  = round(sum(recent)/len(recent)) if recent else 0
    send_telegram(
        f"💓 *Nifty Precision Bot — Alive*\n"
        f"⏰ {now_str}\n"
        f"🔍 Scans: {state['scan_count']}  |  Alerts: {state['alerts_sent']}\n"
        f"📊 Avg score (recent): {avg_sc}/100\n"
        f"📈 VIX: {state['vix'] or '—'}  |  PCR: {state['pcr'] or '—'}\n"
        f"_No signal yet — watching…_"
    )

def _data_error(name, reason):
    cnt = state["data_errors"].get(name, 0) + 1
    state["data_errors"][name] = cnt
    print(f"  [{name}] data error #{cnt}: {reason}")
    if cnt == 1 or cnt % 6 == 0:
        send_telegram(f"⚠️ *Data issue — {name}*\n{reason}\nAttempt #{cnt}")

# ═══════════════════════════════════════════════════════════════
#  EXCEL LOGGING
# ═══════════════════════════════════════════════════════════════

try:
    os.makedirs(EXCEL_FOLDER, exist_ok=True)
except Exception:
    EXCEL_FOLDER = "/tmp"
EXCEL_FILE = os.path.join(EXCEL_FOLDER, "Nifty_Precision_Bot.xlsx")

_HEADERS = [
    "Date","Time","Instrument","Direction","Strike","Option",
    "LTP","Score /100","Rating","R:R","Conds Met /8",
    "EMA9","EMA21","VWAP","RSI(14)","SuperTrend",
    "ORB High","ORB Low","PDH","PDL",
    "CPR Top","CPR Bottom","CPR Narrow?",
    "Fib Level","Fib Price","Fib Dist %","Fib At Level?",
    "SL Level","SL Pts","Target Level","Target Pts",
    "VIX","PCR",
    "Entry Premium","Exit Premium","P&L (Rs)","Result","Notes",
]

def _init_excel():
    if not _openpyxl or os.path.exists(EXCEL_FILE): return
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Signals"; ws.freeze_panes = "A2"
    for i, h in enumerate(_HEADERS, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(horizontal="center")
    wb.save(EXCEL_FILE)

def log_excel(sig):
    if not _openpyxl: return
    try:
        _init_excel()
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb["Signals"]
        now = datetime.now(IST)
        cpr = sig["cpr"]
        ws.append([
            now.strftime("%Y-%m-%d"), now.strftime("%H:%M"),
            sig["instrument"], sig["direction"], sig["strike"], sig["option_type"],
            round(sig["ltp"],1), sig["score"], sig["rating"], sig["rr"], sig["conds_met"],
            round(sig["e9"],1), round(sig["e21"],1), round(sig["vwap"],1),
            round(sig["rsi"],1), "Bull" if sig["st_dir"]==1 else "Bear",
            round(sig["orb_high"],1), round(sig["orb_low"],1),
            round(sig["pdh"],1), round(sig["pdl"],1),
            round(cpr["top"],1), round(cpr["bottom"],1), "Yes" if cpr["narrow"] else "No",
            sig["fib"].get("level") or "", sig["fib"].get("price") or "",
            sig["fib"].get("dist_pct") or "", "Yes" if sig["fib"].get("at_fib") else "No",
            round(sig["sl_level"],1), sig["sl_pts"],
            round(sig["target_level"],1), sig["target_pts"],
            state.get("vix") or "", state.get("pcr") or "",
            "","","","","",
        ])
        last = ws.max_row
        ws.cell(row=last, column=4).font = Font(
            bold=True, color="276221" if sig["direction"]=="CALL" else "9C0006")
        wb.save(EXCEL_FILE)
        print(f"  [EXCEL] Saved → {EXCEL_FILE}")
    except Exception as e:
        print(f"  [EXCEL] {e}")

# ═══════════════════════════════════════════════════════════════
#  SIGNAL FORMAT
# ═══════════════════════════════════════════════════════════════

def format_signal(sig):
    d_icon = "▲ CALL" if sig["direction"] == "CALL" else "▼ PUT"
    icons  = {"ULTRA STRONG":"🔥","STRONG":"💪","MEDIUM":"⚡"}
    ri     = icons.get(sig["rating"],"")
    cpr    = sig["cpr"]
    clines = []
    for k, v in sig["conditions"].items():
        if   isinstance(v, bool): clines.append(f"  {'✅' if v else '❌'} {k}")
        else:                      clines.append(f"  ℹ️ {k}: {v}")
    return (
        f"{ri} *{sig['instrument']} — {d_icon} | {sig['rating']}*\n"
        f"📅 {datetime.now(IST).strftime('%d %b %Y %H:%M IST')}\n\n"
        f"💰 *LTP:* {sig['ltp']:,.1f}\n"
        f"🎫 *Strike:* {sig['strike']} {sig['option_type']}\n"
        f"📊 *Score:* {sig['score']}/100 ({sig['conds_met']}/8 conditions)\n"
        f"⚖️ *R:R:* {sig['rr']}\n\n"
        f"📍 *Trade Levels:*\n"
        f"  Entry  : {sig['ltp']:,.1f}\n"
        f"  SL     : {sig['sl_level']:,.1f}  (−{sig['sl_pts']} pts)\n"
        f"  Target : {sig['target_level']:,.1f}  (+{sig['target_pts']} pts)\n\n"
        f"📈 *Indicators:*\n"
        f"  EMA9/21 : {sig['e9']:,.0f} / {sig['e21']:,.0f}\n"
        f"  VWAP    : {sig['vwap']:,.1f}\n"
        f"  RSI(14) : {sig['rsi']:.1f}\n"
        f"  ST      : {'↑ Bullish' if sig['st_dir']==1 else '↓ Bearish'}\n"
        f"  ORB     : {sig['orb_high']:,.1f} H / {sig['orb_low']:,.1f} L\n"
        f"  PDH/PDL : {sig['pdh']:,.1f} / {sig['pdl']:,.1f}\n"
        f"  CPR     : {cpr['bottom']:,.1f}–{cpr['top']:,.1f} "
        f"({'Narrow 📈' if cpr['narrow'] else 'Wide 📉'})\n\n"
        f"🔢 *Fib:* {sig['fib_label']}\n\n"
        f"📋 *Conditions:*\n" + "\n".join(clines) +
        f"\n\n⚠️ _Trade at your own risk._"
    )

# ═══════════════════════════════════════════════════════════════
#  SCAN SUMMARY (sent after every scan)
# ═══════════════════════════════════════════════════════════════

def send_scan_summary(scan_no, time_str, results):
    lines = []
    for name, sig, rating, score, alerted in results:
        if sig is None:
            lines.append(f"`{name:<12}` ❓ NO DATA")
            continue
        di   = "▲" if sig["direction"] == "CALL" else "▼"
        fill = round(score / 10)
        bar  = "█" * fill + "░" * (10 - fill)
        icon = "🔥" if score>=SCORE_ULTRA else "💪" if score>=SCORE_STRONG else "⚡" if score>=SCORE_MEDIUM else "·"
        tag  = " ✅ *SIGNAL!*" if alerted else ""
        lines.append(
            f"`{name:<12}` {di} {sig['direction']} {icon} "
            f"`{score:3d}/100` `[{bar}]` RSI {sig['rsi']:.0f} | {sig['ltp']:,.0f}{tag}"
        )
    vix_s = f"{state['vix']:.1f}" if state["vix"] else "—"
    pcr_s = f"{state['pcr']:.2f}" if state["pcr"] else "—"
    alerted_any = any(r[4] for r in results)
    hdr = f"🔍 *Scan #{scan_no}* | {time_str} | {'🚨 Signal!' if alerted_any else 'No signal'}"
    send_telegram(
        f"{hdr}\n\n" + "\n".join(lines) +
        f"\n\n📈 VIX `{vix_s}` | PCR `{pcr_s}` | Alerts today: {state['alerts_sent']}"
    )

# ═══════════════════════════════════════════════════════════════
#  MAIN SCAN LOOP
# ═══════════════════════════════════════════════════════════════

def in_market_hours():
    now  = datetime.now(IST)
    if now.weekday() >= 5: return False
    mins = now.hour * 60 + now.minute
    return (BOT_START[0]*60+BOT_START[1]) <= mins <= (MARKET_CLOSE[0]*60+MARKET_CLOSE[1])

def refresh_vix_pcr():
    last = state.get("last_vix_fetch")
    now  = datetime.now(IST)
    if last is None or (now - last).total_seconds() > 900:
        vix_v, pcr_v = fetch_vix_pcr()
        if vix_v is not None: state["vix"] = vix_v
        if pcr_v is not None: state["pcr"] = pcr_v
        state["last_vix_fetch"] = now
        print(f"  [VIX] {state['vix']}   [PCR] {state['pcr']}")

def scan():
    if not in_market_hours(): return
    if state["trades_today"] >= MAX_TRADES_DAY:
        print("[LIMIT] Max daily trades reached"); return

    state["scan_count"] += 1
    now_str = datetime.now(IST).strftime("%H:%M IST")
    print(f"\n{'─'*55}\n[SCAN #{state['scan_count']}] {now_str}")

    if state["scan_count"] % HEARTBEAT_SCANS == 0:
        heartbeat_ping()

    refresh_vix_pcr()

    if state["vix"] and state["vix"] > VIX_EXTREME:
        msg = f"⛔ VIX={state['vix']:.1f} EXTREME — signals blocked"
        print(f"  [BLOCK] {msg}")
        send_telegram(f"🔍 *Scan #{state['scan_count']}* | {now_str}\n\n{msg}")
        return

    results = []
    for name, cfg in INSTRUMENTS.items():
        try:
            sig = analyze(name, cfg)
            if sig is None:
                results.append((name, None, "NO DATA", 0, False)); continue

            state["scores_seen"].append(sig["score"])
            state["scores_seen"] = state["scores_seen"][-30:]

            print(f"  [{name}] {sig['direction']} Score={sig['score']} {sig['rating']} "
                  f"RSI={sig['rsi']:.1f} LTP={sig['ltp']:,.0f}")

            alerted = False
            below_thresh = sig["rating"] in ("WEAK","MEDIUM") and sig["score"] < cfg["min_score"]

            if not below_thresh:
                now = datetime.now(IST)
                last = state["last_signal"].get(name, {})
                suppress = (last and
                            (now-last["time"]).total_seconds()/60 < RESIGAL_GAP_MIN and
                            last["direction"] == sig["direction"] and
                            sig["score"] < last["score"] + 10)
                if not suppress:
                    send_telegram(format_signal(sig))
                    log_excel(sig)
                    sb_push_signal(sig)
                    send_desktop(f"{name} {sig['direction']} {sig['rating']}",
                                 f"LTP {sig['ltp']:,.0f} | Score {sig['score']} | "
                                 f"SL {sig['sl_level']:,.0f} → T {sig['target_level']:,.0f}",
                                 urgent=(sig["score"] >= SCORE_ULTRA))
                    state["last_signal"][name] = {
                        "direction": sig["direction"], "score": sig["score"], "time": now}
                    state["alerts_sent"]  += 1
                    state["trades_today"] += 1
                    alerted = True
                    print(f"      ✅ ALERTED")

            results.append((name, sig, sig["rating"], sig["score"], alerted))

        except Exception as e:
            import traceback; traceback.print_exc()
            results.append((name, None, "ERROR", 0, False))

    send_scan_summary(state["scan_count"], now_str, results)
    sb_push_scan(state["scan_count"], results)
    sb_update_status("running")

def eod_summary():
    now_str = datetime.now(IST).strftime("%d %b %Y")
    send_telegram(
        f"📊 *Nifty Precision Bot — EOD {now_str}*\n\n"
        f"🔍 Scans      : {state['scan_count']}\n"
        f"📢 Signals    : {state['alerts_sent']}\n"
        f"📈 VIX (last) : {state['vix']}\n"
        f"📊 PCR (last) : {state['pcr']}\n"
    )
    send_desktop("Market Closed", f"Scans: {state['scan_count']} | Alerts: {state['alerts_sent']}")
    print(f"\n{'='*55}\nEOD | Scans: {state['scan_count']} | Alerts: {state['alerts_sent']}\n{'='*55}")
    sb_update_status("sleeping")


def _reset_daily_state():
    """Reset per-day counters when a new trading session begins."""
    state["scan_count"]     = 0
    state["alerts_sent"]    = 0
    state["trades_today"]   = 0
    state["last_signal"]    = {}
    state["scores_seen"]    = []
    state["vix"]            = None
    state["pcr"]            = None
    state["last_vix_fetch"] = None
    state["data_errors"]    = {}


def _next_market_open():
    """Return IST datetime of the next weekday 9:35 AM."""
    now  = datetime.now(IST)
    cand = now.replace(hour=BOT_START[0], minute=BOT_START[1], second=0, microsecond=0)
    if cand <= now:
        cand += timedelta(days=1)
    while cand.weekday() >= 5:          # skip Saturday / Sunday
        cand += timedelta(days=1)
    return cand

# ═══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════

def main():
    print("=" * 55)
    print("  Nifty Precision Bot — LOCAL  (runs 24×5, auto-restart)")
    print(f"  Instruments : {', '.join(INSTRUMENTS)}")
    print(f"  Excel log   : {EXCEL_FILE}")
    print(f"  Alert score : ≥{SCORE_STRONG} (Strong) | ≥{SCORE_ULTRA} (Ultra)")
    print("=" * 55)

    if TELEGRAM_TOKEN == "YOUR_BOT_TOKEN_HERE":
        print("\n⚠  TELEGRAM NOT CONFIGURED")
        print("   Open nifty_local_bot.py and fill in TELEGRAM_TOKEN and CHAT_ID\n")

    _init_excel()
    startup_ping()

    # Schedule recurring scans; EOD detection is done inside the while loop
    schedule.every(SCAN_INTERVAL).minutes.do(scan)

    # First scan immediately if already in market hours
    if in_market_hours():
        scan()

    eod_sent_date  = None   # date on which EOD summary was last sent
    active_date    = None   # date of the current / most recent trading session

    while True:
        try:
            schedule.run_pending()
            now    = datetime.now(IST)
            today  = now.date()
            in_mkt = in_market_hours()

            # ── New trading session started ──────────────────────
            if in_mkt and active_date != today:
                if active_date is not None:     # not the very first boot
                    print(f"\n🌅  New session: {today} — resetting daily state")
                    _reset_daily_state()
                    startup_ping()
                    scan()                      # immediate first scan of the day
                active_date = today

            # ── End of trading day ───────────────────────────────
            after_close = now.hour > 15 or (now.hour == 15 and now.minute >= 35)
            if after_close and now.weekday() < 5 and eod_sent_date != today:
                eod_summary()
                eod_sent_date = today
                nxt  = _next_market_open()
                diff = nxt - now
                hrs, leftover = divmod(int(diff.total_seconds()), 3600)
                mins = leftover // 60
                sleep_msg = (
                    f"💤 *Market Closed — Bot Sleeping*\n"
                    f"Next session: *{nxt.strftime('%a %d %b, %H:%M IST')}*\n"
                    f"Auto-resumes in {hrs}h {mins}m — no restart needed 🔄"
                )
                print(f"\n💤  Sleeping until {nxt.strftime('%a %d %b %H:%M IST')} "
                      f"({hrs}h {mins}m away)")
                send_telegram(sleep_msg)
                send_desktop("Market Closed",
                             f"Next: {nxt.strftime('%a %d %b %H:%M')} | "
                             f"{hrs}h {mins}m")

            # Sleep longer when outside market hours (saves CPU)
            time.sleep(20 if in_mkt else 60)

        except KeyboardInterrupt:
            print("\n\n[STOP] Ctrl-C received — goodbye!")
            break
        except Exception as exc:
            print(f"\n[MAIN-LOOP ERROR] {exc}")
            time.sleep(30)          # brief pause, then keep running


if __name__ == "__main__":
    main()
