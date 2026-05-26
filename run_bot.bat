@echo off
title Nifty Precision Bot (Local)
color 0A

echo ============================================
echo   Nifty Precision Bot - Local Runner
echo ============================================
echo.

:: Change to the folder where this .bat lives
cd /d "%~dp0"

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found! Please install Python 3.9+ and add to PATH.
    pause
    exit /b 1
)

:: Install / upgrade dependencies silently
echo Installing / checking dependencies...
pip install -q yfinance pandas numpy openpyxl requests schedule pytz plyer

echo.
echo Starting bot... (runs 9:35 AM - 3:32 PM IST on weekdays)
echo Press Ctrl+C to stop manually.
echo.

python nifty_local_bot.py

echo.
echo Bot has exited. See you tomorrow!
pause
