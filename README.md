# Minsorterbank

Turn an Indian bank statement (ICICI / Axis / HDFC / SBI / generic Excel / CSV)
into a grouped Deposits / Withdrawals workbook.

## Use it

### Desktop app (recommended for everyday users)

Grab the latest download for your OS from
[Releases](../../releases/latest).

#### Windows

1. Download `Minsorterbank-windows-portable.zip`.
2. Right-click the zip &rarr; **Extract All** to a folder you like
   (e.g. your Desktop).
3. Open the extracted `Minsorterbank` folder.
4. Double-click **`Start-Minsorterbank.bat`**.
5. The first time only, it will ask you to install Python &mdash; just
   click the big yellow **Download Python** button on the page it opens,
   tick **&ldquo;Add python.exe to PATH&rdquo;** on the installer&rsquo;s
   first screen, then double-click `Start-Minsorterbank.bat` again.
6. Your browser will open at `http://127.0.0.1:<port>`. Upload your
   statement, get the sorted spreadsheet back.

Why not a `.exe`? Windows Defender and other antivirus tools aggressively
flag PyInstaller-packed executables as false positives. The launcher
approach uses real Python (which AVs trust), so it just works.

#### macOS (Apple Silicon &mdash; M1/M2/M3/M4)

1. Download `Minsorterbank-macos-arm64.dmg`.
2. Double-click it, then drag the `Minsorterbank` app into Applications.
3. Open Launchpad and click **Minsorterbank**.

> First-launch note on macOS: because the build isn&rsquo;t signed with an
> Apple Developer ID, macOS will say &ldquo;Apple could not verify
> Minsorterbank&rdquo;. Right-click the app &rarr; **Open** &rarr;
> **Open** to bypass that once.

The app launches a small local web UI in your browser, processes the file
entirely on your machine, and gives you back the sorted `.xlsx`. Nothing is
uploaded anywhere.

### CLI

```bash
python sort_statement.py path/to/statement.xls
```

### Run the web UI from source

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app/app.py
```

## Build installers locally

```bash
pip install pyinstaller
pyinstaller --clean -y Minsorterbank.spec
# dist/Minsorterbank.exe      (Windows: single self-contained .exe)
# dist/Minsorterbank          (Linux: single self-contained binary)
# dist/Minsorterbank.app      (macOS: single .app bundle)
```

## Releasing

Push a tag matching `v*` and GitHub Actions will build for Windows x64
and macOS arm64 (Apple Silicon), then publish them to a Release:

```bash
git tag v1.0.0
git push origin v1.0.0
```
