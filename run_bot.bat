@echo off
title Nifty Precision Bot (Local)
color 0A

echo ============================================
echo   Nifty Precision Bot - Local Runner
echo ============================================
echo.

:: Change to the folder where this .bat lives
cd /d "%~dp0"

:: ── Virtual-env setup (avoids Anaconda NumPy conflicts) ─────────────────────
set VENV_DIR=%~dp0nse_venv

if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo [Setup] Creating isolated Python environment (first-time only)...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo ERROR: Could not create virtual environment.
        echo Make sure Python 3.9+ is installed and on PATH.
        pause
        exit /b 1
    )
    echo [Setup] Installing packages...
    "%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip -q
    "%VENV_DIR%\Scripts\pip.exe" install -q ^
        "yfinance>=0.2.55" ^
        "pandas>=2.0" ^
        "numpy>=1.26,<2" ^
        openpyxl requests schedule pytz plyer
    echo [Setup] Done!
    echo.
) else (
    :: Silently upgrade yfinance in case it's stale
    "%VENV_DIR%\Scripts\pip.exe" install -q --upgrade "yfinance>=0.2.55" >nul 2>&1
)

:: ── Activate and run ────────────────────────────────────────────────────────
call "%VENV_DIR%\Scripts\activate.bat"

echo Starting bot... (runs 9:35 AM - 3:32 PM IST on weekdays)
echo Press Ctrl+C to stop manually.
echo.

python nifty_local_bot.py

echo.
echo Bot has exited. See you tomorrow!
pause
