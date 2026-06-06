@echo off
title Nifty Precision Bot
color 0A
cd /d "%~dp0"

echo ============================================
echo   Nifty Precision Bot - Local Runner
echo ============================================
echo.

:: ── Create venv if it doesn't exist ─────────────────────────
if not exist "nse_venv\Scripts\python.exe" (
    echo [Setup] Creating isolated Python environment...
    python -m venv nse_venv
    if errorlevel 1 (
        echo ERROR: Could not create venv. Is Python installed?
        pause & exit /b 1
    )
    echo [Setup] Installing packages with compatible numpy...
    nse_venv\Scripts\pip install --upgrade pip -q
    nse_venv\Scripts\pip install ^
        "yfinance>=0.2.55" ^
        "pandas>=2.0" ^
        "numpy>=1.26,<2" ^
        openpyxl requests schedule pytz plyer supabase -q
    echo [Setup] Done!
    echo.
)

echo Starting bot...  (9:35 AM to 3:30 PM IST, weekdays)
echo Press Ctrl+C to stop.
echo.

nse_venv\Scripts\python nifty_local_bot.py

echo.
echo Bot stopped. Press any key to close.
pause
