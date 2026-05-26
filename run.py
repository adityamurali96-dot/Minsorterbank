"""
Fallback launcher for Minsorterbank.

Double-click this file if Start-Minsorterbank.bat / .cmd does not work
on your system.  Requires Python 3.8+ to already be installed.
"""

import os
import subprocess
import sys
from pathlib import Path


def main():
    os.chdir(Path(__file__).resolve().parent)

    if sys.version_info < (3, 8):
        print(
            f"Minsorterbank requires Python 3.8 or later.\n"
            f"You are running Python {sys.version}.\n"
            f"Download a newer version from https://www.python.org/downloads/"
        )
        return 1

    if os.name == "nt":
        venv_py = Path(".venv", "Scripts", "python.exe")
    else:
        venv_py = Path(".venv", "bin", "python")

    if not venv_py.exists():
        print("Creating virtual environment...")
        subprocess.check_call([sys.executable, "-m", "venv", ".venv"])

    print("Checking dependencies...")
    subprocess.check_call(
        [str(venv_py), "-m", "pip", "install", "-q", "-r", "requirements.txt"]
    )

    print("Starting Minsorterbank...\n")
    return subprocess.call([str(venv_py), os.path.join("app", "app.py")])


if __name__ == "__main__":
    try:
        code = main()
    except Exception:
        import traceback

        traceback.print_exc()
        code = 1

    if code:
        print("\nMinsorterbank exited with an error (see above).")
    else:
        print("\nMinsorterbank stopped.")
    try:
        input("Press Enter to close...")
    except EOFError:
        pass
    sys.exit(code or 0)
