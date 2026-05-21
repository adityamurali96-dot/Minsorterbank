#!/bin/bash
# ============================================================
#  Minsorterbank launcher for macOS.
#  Double-click this file. It will:
#    1. Check that Python 3 is installed (and walk you through
#       installing it if not).
#    2. Set up a private folder of dependencies the first time
#       (one-time, ~30 seconds on a normal internet connection).
#    3. Start the Minsorterbank app and open it in your browser.
#
#  Nothing leaves your computer; the app runs on localhost.
# ============================================================

# Move to this script's folder so paths work no matter where it's
# run from (including when double-clicked from Finder).
cd "$(dirname "$0")"

# Always pause before exit so the Terminal window doesn't just
# vanish if something goes wrong -- the user needs to see the
# error message.
pause_before_exit() {
  local code=$?
  echo
  if [ "$code" -eq 0 ]; then
    echo " Minsorterbank stopped."
  else
    echo " ============================================================"
    echo "  Minsorterbank exited with an error (code $code)."
    echo "  Scroll up in this window to read the message above."
    echo " ============================================================"
  fi
  echo
  read -n 1 -s -r -p " Press any key to close this window..."
  echo
  exit "$code"
}
trap pause_before_exit EXIT

echo
echo " Minsorterbank"
echo " -------------"
echo

# --- Step 1: find python3.
if ! command -v python3 >/dev/null 2>&1; then
  osascript -e 'display dialog "Python 3 is not installed on this Mac yet.\n\nThe Python download page will open in your browser. Download the macOS installer, run it, then double-click Start-Minsorterbank.command again." buttons {"OK"} default button "OK" with icon caution with title "Minsorterbank"' >/dev/null 2>&1 || true
  open https://www.python.org/downloads/
  echo " Python 3 is not installed."
  echo " The download page has been opened in your browser."
  echo " After installing Python, double-click this file again."
  exit 1
fi

echo " Using Python: $(command -v python3)"
python3 --version
echo

# --- Step 2: create a private virtual environment on first run.
if [ ! -x ".venv/bin/python" ]; then
  echo " First-time setup: creating a private Python environment..."
  if ! python3 -m venv .venv; then
    echo
    echo " Could not create the virtual environment."
    echo " Try reinstalling Python from https://www.python.org/downloads/"
    exit 1
  fi
fi

VENV_PY="./.venv/bin/python"

if [ ! -x "$VENV_PY" ]; then
  echo
  echo " Could not find the private Python at $VENV_PY"
  echo " Try deleting the .venv folder next to this script and run again."
  exit 1
fi

# --- Step 3: report which required libraries are present, then run
#     pip to install anything missing. pip is idempotent -- packages
#     that are already up to date will just print "Requirement
#     already satisfied".
echo " Checking required libraries..."
echo
missing=0
check_pkg() {
  local pip_name="$1"
  local import_name="${2:-$1}"
  if "$VENV_PY" -c "import $import_name" >/dev/null 2>&1; then
    printf "   %-12s - OK\n" "$pip_name"
  else
    printf "   %-12s - missing\n" "$pip_name"
    missing=$((missing + 1))
  fi
}
check_pkg pandas
check_pkg openpyxl
check_pkg xlrd
check_pkg lxml
check_pkg Flask flask
echo

if [ "$missing" -eq 0 ]; then
  echo " All required libraries are already installed."
else
  echo " Installing missing libraries from requirements.txt..."
  echo " (first run takes ~30 seconds, later runs are instant)"
  echo
  "$VENV_PY" -m pip install --upgrade pip >/dev/null 2>&1 || true
  if ! "$VENV_PY" -m pip install -r requirements.txt; then
    echo
    echo " Could not install dependencies. Check your internet connection"
    echo " and try again."
    exit 1
  fi
fi

# --- Step 4: run the app. app.py auto-picks a free port and opens
#     the browser to http://127.0.0.1:<port>.
echo
echo " Starting Minsorterbank..."
echo " Your browser will open in a moment."
echo " Keep this Terminal window open while you use the app."
echo " Close this window to quit."
echo
"$VENV_PY" app/app.py
