@echo off
setlocal enableextensions enabledelayedexpansion

REM ============================================================
REM  Minsorterbank launcher for Windows.
REM  Double-click this file. It will:
REM    1. Check that Python is installed (and walk you through
REM       installing it if not).
REM    2. Set up a private folder of dependencies the first time
REM       (one-time, ~30 seconds on a normal internet connection).
REM    3. Start the Minsorterbank app and open it in your browser.
REM
REM  Nothing leaves your computer; the app runs on localhost.
REM ============================================================

title Minsorterbank

REM --- Move to this script's folder so paths work no matter where it's run from.
cd /d "%~dp0"

echo.
echo  Minsorterbank
echo  -------------
echo.

REM --- Step 1: find a Python interpreter.
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)

if not defined PY (
  echo  Python is not installed on this computer yet.
  echo.
  echo  I'm opening the Python download page in your browser now.
  echo  Please install it ^(click the big yellow "Download Python" button^)
  echo  and then double-click "Start-Minsorterbank.bat" again.
  echo.
  echo  IMPORTANT: on the first installer screen, tick the box that
  echo  says "Add python.exe to PATH" before clicking Install.
  echo.
  start "" https://www.python.org/downloads/
  echo  Press any key to close this window...
  pause >nul
  exit /b 1
)

REM --- Step 2: create a private virtual environment on first run.
if not exist ".venv\Scripts\python.exe" (
  echo  First-time setup: creating a private Python environment...
  %PY% -m venv .venv
  if errorlevel 1 (
    echo.
    echo  Could not create the virtual environment. Press any key to exit.
    pause >nul
    exit /b 1
  )
)

set "VENV_PY=.venv\Scripts\python.exe"

REM --- Step 3: make sure dependencies are installed (idempotent).
echo  Checking dependencies...
"%VENV_PY%" -m pip install --upgrade pip >nul 2>nul
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo  Could not install dependencies. Check your internet connection
  echo  and try again. Press any key to exit.
  pause >nul
  exit /b 1
)

REM --- Step 4: run the app. app.py auto-picks a free port and opens
REM     the browser to http://127.0.0.1:^<port^>.
echo.
echo  Starting Minsorterbank...
echo  Your browser will open in a moment.
echo  Keep this window open while you use the app.
echo  Close this window to quit.
echo.
"%VENV_PY%" app\app.py

echo.
echo  Minsorterbank stopped. Press any key to close this window.
pause >nul
endlocal
