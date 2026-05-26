@echo off
title Minsorterbank
cd /d "%~dp0"

echo.
echo  Minsorterbank
echo  -------------
echo.

REM --- Run the main flow as a subroutine. setlocal and all advanced
REM     features live inside :main, so even if something goes wrong
REM     the pause below still runs and the window stays open.
call :main
set EXITCODE=%ERRORLEVEL%

echo.
if "%EXITCODE%"=="0" (
  echo  Minsorterbank stopped.
) else (
  echo  ============================================================
  echo   Minsorterbank exited with an error (code %EXITCODE%).
  echo   Scroll up in this window to read the message above.
  echo  ============================================================
)
echo.
echo  Press any key to close this window...
pause >nul
exit %EXITCODE%


REM ============================================================
REM  Main flow (runs inside setlocal enabledelayedexpansion)
REM ============================================================
:main
setlocal enableextensions enabledelayedexpansion

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
    powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile $env:TEMP\python-installer.exe } catch { exit 1 }"
    if errorlevel 1 (
      echo.
      echo  Could not download Python automatically. Please check your
      echo  internet connection, then install Python manually from:
      echo    https://www.python.org/downloads/
      echo  Be sure to tick "Add python.exe to PATH" during install.
      start "" https://www.python.org/downloads/
      exit /b 1
    )

    echo  Running the Python installer (this can take a minute)...
    "!PY_INSTALLER!" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
    del "!PY_INSTALLER!" >nul 2>nul

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
    exit /b 1
  )

  echo  Python installed successfully.
  echo.
)

echo  Using Python: !PY!
call %PY% --version
echo.

REM --- Verify the Python version is at least 3.8.
call %PY% -c "import sys; sys.exit(0 if sys.version_info >= (3,8) else 1)" >nul 2>nul
if errorlevel 1 (
  echo  The Python found on this system is too old for Minsorterbank.
  echo  Minsorterbank requires Python 3.8 or later.
  echo  Please install a newer version from https://www.python.org/downloads/
  start "" https://www.python.org/downloads/
  exit /b 1
)

REM --- Step 2: create a private virtual environment on first run.
if not exist ".venv\Scripts\python.exe" (
  echo  First-time setup: creating a private Python environment...
  call %PY% -m venv .venv
  if errorlevel 1 (
    echo.
    echo  Could not create the virtual environment.
    echo  This usually means Python was installed without the "venv" module,
    echo  or the folder is read-only. Try reinstalling Python from
    echo  https://www.python.org/downloads/ with default options.
    exit /b 1
  )
)

set "VENV_PY=.venv\Scripts\python.exe"

if not exist "%VENV_PY%" (
  echo.
  echo  Could not find the private Python at "%VENV_PY%".
  echo  Try deleting the .venv folder next to this script and run again.
  exit /b 1
)

REM --- Step 3: show which required libraries are already installed,
REM     then run pip to install anything missing. pip is idempotent --
REM     packages that are already up to date will just print
REM     "Requirement already satisfied".
echo  Checking required libraries...
echo.
set "MISSING=0"
call :check_pkg pandas
call :check_pkg openpyxl
call :check_pkg xlrd
call :check_pkg lxml
call :check_pkg flask Flask
echo.

if "!MISSING!"=="0" (
  echo  All required libraries are already installed.
) else (
  echo  Installing missing libraries from requirements.txt...
  echo  ^(first run takes ~30 seconds, later runs are instant^)
  echo.
  "%VENV_PY%" -m pip install --upgrade pip >nul 2>nul
  "%VENV_PY%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo  Could not install dependencies. Check your internet connection
    echo  and try again.
    exit /b 1
  )
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
endlocal
exit /b %ERRORLEVEL%


REM ============================================================
REM  Helpers
REM ============================================================
:find_python
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && for /f "delims=" %%P in ('where python') do if not defined PY set "PY=""%%P"""
)
if not defined PY (
  for %%V in (313 312 311 310 39 38) do (
    if not defined PY (
      if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
    )
  )
)
goto :eof

REM  Usage: call :check_pkg <pip-name> [<import-name>]
REM  Prints e.g. "   pandas         OK" or "   pandas         missing"
REM  and bumps MISSING when the import fails. Always uses the venv
REM  python -- which is guaranteed to exist by the time we get here.
:check_pkg
set "PKG=%~1"
set "IMP=%~2"
if not defined IMP set "IMP=%~1"
".venv\Scripts\python.exe" -c "import !IMP!" >nul 2>nul
if errorlevel 1 (
  echo    !PKG!  - missing
  set /a MISSING+=1
) else (
  echo    !PKG!  - OK
)
goto :eof
