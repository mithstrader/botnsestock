@echo off
setlocal EnableDelayedExpansion
title NSE Bot — Local Run

:: ════════════════════════════════════════════════════════
::  run_local.cmd  —  Run NSE bot directly on this PC
::
::  First-time setup:
::    1. Copy  local_creds_template.cmd  →  local_creds.cmd
::    2. Fill in TELEGRAM_TOKEN and CHAT_ID in local_creds.cmd
::    3. Double-click this file (or pass mode as argument)
::
::  Usage:
::    run_local.cmd           — live scan (default)
::    run_local.cmd test      — quick connectivity test
::    run_local.cmd premarket — pre-market scan only
:: ════════════════════════════════════════════════════════

set MODE=%~1
if "%MODE%"=="" set MODE=live

echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║      NSE BOT — LOCAL RUN  v1.0               ║
echo  ╚══════════════════════════════════════════════╝
echo.

:: ── Load credentials ─────────────────────────────────────────────
if not exist "%~dp0local_creds.cmd" (
    echo  [ERROR] local_creds.cmd not found!
    echo.
    echo  Create it by copying the template:
    echo    copy "%~dp0local_creds_template.cmd" "%~dp0local_creds.cmd"
    echo.
    echo  Then open local_creds.cmd and fill in your:
    echo    TELEGRAM_TOKEN  ^(from @BotFather^)
    echo    CHAT_ID         ^(from @userinfobot^)
    echo.
    pause
    exit /b 1
)
call "%~dp0local_creds.cmd"

if "%TELEGRAM_TOKEN%"=="" (
    echo  [ERROR] TELEGRAM_TOKEN is empty in local_creds.cmd
    pause
    exit /b 1
)
if "%CHAT_ID%"=="" (
    echo  [ERROR] CHAT_ID is empty in local_creds.cmd
    pause
    exit /b 1
)

:: Show current IST time
for /f "delims=" %%T in ('powershell -NoProfile -Command "[System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([System.DateTime]::UtcNow,\"India Standard Time\").ToString('ddd dd-MMM-yyyy  hh:mm tt')"') do set IST_NOW=%%T

echo   Time (IST)   : %IST_NOW%
echo   Mode         : %MODE%
echo   Token loaded : YES  (first 8 chars: %TELEGRAM_TOKEN:~0,8%...)
echo   Chat ID      : %CHAT_ID%
echo   Excel logs   : %~dp0logs\
echo.

:: ── Find Python ──────────────────────────────────────────────────
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
if "%PYTHON%"=="" (
    where python3 >nul 2>&1 && set PYTHON=python3
)
if "%PYTHON%"=="" (
    echo  [ERROR] Python not found in PATH.
    echo  Install Python 3.9+ from https://python.org
    pause
    exit /b 1
)
for /f "delims=" %%V in ('%PYTHON% --version 2^>^&1') do set PY_VER=%%V
echo   Python       : %PY_VER%

:: ── Check required packages ──────────────────────────────────────
echo   Checking packages...
%PYTHON% -c "import requests, schedule, pytz, openpyxl" 2>nul
if errorlevel 1 (
    echo.
    echo  [SETUP] Installing required packages...
    %PYTHON% -m pip install requests schedule pytz openpyxl --quiet
    echo  [SETUP] Done.
)
echo.

:: ── Set environment variables for the bot ────────────────────────
set EXCEL_FOLDER=%~dp0logs

echo ──────────────────────────────────────────────────
echo.

:: ── Run in selected mode ──────────────────────────────────────────
if /i "%MODE%"=="test" (
    echo  Running TEST mode  ^(quick connectivity check^)...
    echo.
    %PYTHON% "%~dp0nse_options_bot_stock_cookies_xl_ex_enhanced.py" test
    echo.
    echo  Test complete. Check Telegram for the result.
    pause
    exit /b 0
)

if /i "%MODE%"=="premarket" (
    echo  Running PRE-MARKET scan only...
    echo.
    set PREMARKET_ONLY=true
    %PYTHON% "%~dp0nse_options_bot_stock_cookies_xl_ex_enhanced.py"
    echo.
    echo  Pre-market scan complete. Check Telegram.
    pause
    exit /b 0
)

:: ── Live mode: keep window open, bot runs until 3:28 PM ──────────
echo  Starting LIVE scan  ^(runs until 3:28 PM IST^)...
echo  Telegram signals will arrive every 5 minutes.
echo  Excel log: %~dp0logs\NSE_Bot_Master.xlsx
echo.
echo  Press Ctrl+C to stop early.
echo ──────────────────────────────────────────────────
echo.
%PYTHON% "%~dp0nse_options_bot_stock_cookies_xl_ex_enhanced.py"

echo.
echo ──────────────────────────────────────────────────
echo  Bot exited. Check Telegram for EOD summary.
echo  Excel: %~dp0logs\NSE_Bot_Master.xlsx
echo ──────────────────────────────────────────────────
echo.
pause
