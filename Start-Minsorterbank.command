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

set -e

# Move to this script's folder so paths work no matter where it's
# run from (including when double-clicked from Finder).
cd "$(dirname "$0")"

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
  echo
  read -n 1 -s -r -p " Press any key to close this window..."
  echo
  exit 1
fi

# --- Step 2: create a private virtual environment on first run.
if [ ! -x ".venv/bin/python" ]; then
  echo " First-time setup: creating a private Python environment..."
  python3 -m venv .venv
fi

VENV_PY="./.venv/bin/python"

# --- Step 3: make sure dependencies are installed (idempotent).
echo " Checking dependencies..."
"$VENV_PY" -m pip install --upgrade pip >/dev/null 2>&1 || true
"$VENV_PY" -m pip install -r requirements.txt

# --- Step 4: run the app. app.py auto-picks a free port and opens
#     the browser to http://127.0.0.1:<port>.
echo
echo " Starting Minsorterbank..."
echo " Your browser will open in a moment."
echo " Keep this Terminal window open while you use the app."
echo " Close this window to quit."
echo
"$VENV_PY" app/app.py
