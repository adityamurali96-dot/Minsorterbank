# Minsorterbank

Turn an Indian bank statement (ICICI / Axis / HDFC / SBI / generic Excel / CSV)
into a grouped Deposits / Withdrawals workbook.

## Use it

### Desktop app (recommended for everyday users)

Grab the latest installer for your OS from
[Releases](../../releases/latest):

- **Windows** &mdash; `Minsorterbank-windows-x64.exe` &rarr; double-click to
  run. It&rsquo;s a single self-contained executable; nothing to install.
- **macOS (Apple Silicon &mdash; M1/M2/M3/M4)** &mdash;
  `Minsorterbank-macos-arm64.dmg` &rarr; open &rarr; drag the
  `Minsorterbank` app into Applications &rarr; launch.

The app launches a small local web UI in your browser, processes the file
entirely on your machine, and gives you back the sorted `.xlsx`. Nothing is
uploaded anywhere.

> First-launch note on macOS: because the build isn&rsquo;t signed with an
> Apple Developer ID, macOS will say &ldquo;Apple could not verify
> Minsorterbank&rdquo;. Right-click the app &rarr; **Open** &rarr;
> **Open** to bypass that once.

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
