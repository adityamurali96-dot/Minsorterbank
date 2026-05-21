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
call :find_python
if not defined PY (
  echo  Python is not installed on this computer yet.
  echo  I'll try to install it for you automatically. This may take a
  echo  few minutes and may ask for administrator permission.
  echo.

  REM --- Try winget first (built into Windows 10 1709+ and Windows 11).
  where winget >nul 2>nul
  if not errorlevel 1 (
    echo  Installing Python via winget...
    winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements --silent
  )

  call :find_python
  if not defined PY (
    REM --- Fallback: download the official installer and run it silently.
    echo  Downloading the official Python installer...
    set "PY_INSTALLER=%TEMP%\python-installer.exe"
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile $env:TEMP\python-installer.exe } catch { exit 1 }"
    if errorlevel 1 (
      echo.
      echo  Could not download Python automatically. Please check your
      echo  internet connection, then install Python manually from:
      echo    https://www.python.org/downloads/
      echo  Be sure to tick "Add python.exe to PATH" during install.
      start "" https://www.python.org/downloads/
      pause >nul
      exit /b 1
    )

    echo  Running the Python installer (this can take a minute)...
    "%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
    del "%PY_INSTALLER%" >nul 2>nul

    REM --- Refresh PATH in this session so the new python is visible.
    for /f "usebackq tokens=2,*" %%A in (`reg query "HKCU\Environment" /v PATH 2^>nul ^| findstr /i "PATH"`) do set "USER_PATH=%%B"
    for /f "usebackq tokens=2,*" %%A in (`reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul ^| findstr /i "PATH"`) do set "SYS_PATH=%%B"
    set "PATH=!SYS_PATH!;!USER_PATH!"

    call :find_python
  )

  if not defined PY (
    echo.
    echo  Python installation did not complete successfully.
    echo  Please install Python manually from https://www.python.org/downloads/
    echo  and tick "Add python.exe to PATH" during install, then run this
    echo  script again.
    start "" https://www.python.org/downloads/
    pause >nul
    exit /b 1
  )

  echo  Python installed successfully.
  echo.
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
exit /b 0

:find_python
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)
if not defined PY (
  if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
)
goto :eof
