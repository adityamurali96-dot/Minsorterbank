@echo off
title Minsorterbank
cd /d "%~dp0"

echo.
echo  Minsorterbank
echo  -------------
echo.

REM --- Locate Python ---
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo  ERROR: Python was not found on your PATH.
    echo  Install Python 3.8+ from https://www.python.org/downloads/
    echo  and make sure "Add python.exe to PATH" is checked.
    echo.
    pause
    exit /b 1
)

echo  Using: %PY%
%PY% --version
echo.

REM --- Create virtual environment if missing ---
if not exist ".venv\Scripts\python.exe" (
    echo  First-time setup: creating a private Python environment...
    %PY% -m venv .venv
    if errorlevel 1 (
        echo.
        echo  ERROR: Could not create virtual environment.
        echo  Try reinstalling Python from https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

set "VENV_PY=.venv\Scripts\python.exe"

REM --- Install dependencies ---
echo  Checking dependencies...
"%VENV_PY%" -m pip install -q -r requirements.txt
if errorlevel 1 (
    echo.
    echo  ERROR: Could not install dependencies. Check your internet connection.
    pause
    exit /b 1
)

REM --- Launch the Flask app on localhost ---
echo  Starting Minsorterbank on localhost...
echo  Your browser will open automatically.
echo  Keep this window open while using the app.
echo  Press Ctrl+C or close this window to quit.
echo.
"%VENV_PY%" app\app.py

echo.
if errorlevel 1 (
    echo  Minsorterbank exited with an error. See messages above.
) else (
    echo  Minsorterbank stopped.
)
echo.
pause
