@echo off
setlocal EnableDelayedExpansion
title NSE Bot Launcher

:: ════════════════════════════════════════════════════════
::  NSE Bot Launcher — triggers all 3 GitHub Actions
::  Double-click this file every morning before 9:30 AM IST
:: ════════════════════════════════════════════════════════

set REPO=mithstrader/botnsestock
set API=https://api.github.com/repos/%REPO%/actions/workflows
set BODY={"ref":"master"}

:: Workflow IDs
set WF_PREMARKET=281940257
set WF_NSE_BOT=280850925
set WF_NIFTY=281949914

echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║         NSE BOT LAUNCHER  v1.0               ║
echo  ║   github.com/mithstrader/botnsestock         ║
echo  ╚══════════════════════════════════════════════╝
echo.

:: ── Load token from bot_token.txt ──────────────────
if not exist "%~dp0bot_token.txt" (
    echo  [ERROR] bot_token.txt not found!
    echo.
    echo  Create a file called  bot_token.txt  in:
    echo  %~dp0
    echo  and paste your GitHub Personal Access Token inside it.
    echo  ^(token needs  repo + workflow  scopes^)
    echo.
    pause
    exit /b 1
)
set /p TOKEN=<"%~dp0bot_token.txt"
set TOKEN=%TOKEN: =%

if "%TOKEN%"=="" (
    echo  [ERROR] bot_token.txt is empty. Add your GitHub PAT.
    pause
    exit /b 1
)

:: Show current IST time
for /f "delims=" %%T in ('powershell -NoProfile -Command "[System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([System.DateTime]::UtcNow,\"India Standard Time\").ToString('ddd dd-MMM-yyyy  hh:mm tt')"') do set IST_NOW=%%T
echo   Current IST : %IST_NOW%
echo   Repo        : https://github.com/%REPO%
echo   Actions     : https://github.com/%REPO%/actions
echo.
echo ──────────────────────────────────────────────────

:: ── 1. NSE Pre-Market Scan ─────────────────────────
echo.
echo  [1/3]  Triggering NSE Pre-Market Scan...
curl.exe -s -o nul -w "       HTTP %%{http_code}" ^
  -X POST "%API%/%WF_PREMARKET%/dispatches" ^
  -H "Authorization: token %TOKEN%" ^
  -H "Accept: application/vnd.github.v3+json" ^
  -H "Content-Type: application/json" ^
  -d "%BODY%"
echo  (204 = OK)

:: ── 2. NSE Options Bot ─────────────────────────────
echo.
echo  [2/3]  Triggering NSE Options Bot (live scan)...
curl.exe -s -o nul -w "       HTTP %%{http_code}" ^
  -X POST "%API%/%WF_NSE_BOT%/dispatches" ^
  -H "Authorization: token %TOKEN%" ^
  -H "Accept: application/vnd.github.v3+json" ^
  -H "Content-Type: application/json" ^
  -d "%BODY%"
echo  (204 = OK)

:: ── 3. Nifty Precision Bot ─────────────────────────
echo.
echo  [3/3]  Triggering Nifty Precision Bot...
curl.exe -s -o nul -w "       HTTP %%{http_code}" ^
  -X POST "%API%/%WF_NIFTY%/dispatches" ^
  -H "Authorization: token %TOKEN%" ^
  -H "Accept: application/vnd.github.v3+json" ^
  -H "Content-Type: application/json" ^
  -d "%BODY%"
echo  (204 = OK)

:: ── Wait then show live run status ─────────────────
echo.
echo ──────────────────────────────────────────────────
echo   Waiting 6 seconds for runs to register...
powershell -NoProfile -Command "Start-Sleep -Seconds 6"

echo.
echo  Latest runs:
echo.
powershell -NoProfile -Command ^
  "$h=@{Authorization='token %TOKEN%';Accept='application/vnd.github.v3+json'};" ^
  "$r=Invoke-RestMethod 'https://api.github.com/repos/%REPO%/actions/runs?per_page=6' -Headers $h;" ^
  "foreach($run in $r.workflow_runs | Where-Object {$_.name -notlike '*pages*' -and $_.name -notlike '*Graph*'} | Select-Object -First 4){" ^
  "  $ist=[System.DateTimeOffset]::Parse($run.created_at).ToOffset([System.TimeSpan]::FromHours(5.5));" ^
  "  $icon=if($run.status -eq 'in_progress'){'>>> RUNNING'}elseif($run.conclusion -eq 'success'){'[  OK  ]'}else{'['+$run.status+']'};" ^
  "  Write-Host \"  $icon  $($run.name.PadRight(24)) $($ist.ToString('HH:mm')) IST\" }"

echo.
echo ──────────────────────────────────────────────────
echo.
echo  All bots are running on GitHub Actions.
echo  Check Telegram for signals every 5 minutes.
echo.
echo  Monitor: https://github.com/%REPO%/actions
echo.
pause
