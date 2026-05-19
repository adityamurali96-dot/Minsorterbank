# Minsorterbank

Turn an Indian bank statement (ICICI / Axis / HDFC / SBI / generic Excel / CSV)
into a grouped Deposits / Withdrawals workbook.

Runs entirely on your computer &mdash; nothing is uploaded anywhere.

## Use it (recommended for everyday users)

1. Go to [Releases](../../releases/latest) and download
   **`Minsorterbank.zip`**.
2. Right-click the zip &rarr; **Extract All** (Windows) or just
   double-click it (macOS).
3. Open the extracted `Minsorterbank` folder.
4. Double-click the launcher for your computer:
   - **Windows** &rarr; `Start-Minsorterbank.bat`
   - **macOS** &rarr; `Start-Minsorterbank.command`
5. A small Terminal / Command Prompt window appears (leave it open) and
   your browser opens the Minsorterbank page at a local address like
   `http://127.0.0.1:54321`. Upload your statement, download the sorted
   spreadsheet.
6. When you&rsquo;re done, close the Terminal / Command Prompt window
   to quit.

The launcher will guide you through installing Python the first time if
you don&rsquo;t already have it (one-time, ~1 minute). After that, every
run is instant.

### macOS first-launch note

The first time you double-click `Start-Minsorterbank.command`, macOS may
say &ldquo;Apple cannot check it for malicious software.&rdquo;
Right-click the file &rarr; **Open** &rarr; **Open**. After that once,
regular double-click works forever.

### Windows first-launch note

If you don&rsquo;t have Python installed yet, the launcher will open the
Python download page in your browser. Click the big yellow **Download
Python** button, run the installer, and on the **first installer
screen tick &ldquo;Add python.exe to PATH&rdquo;** before clicking
Install. Then double-click `Start-Minsorterbank.bat` again.

## Use it from the command line

```bash
python sort_statement.py path/to/statement.xls
```

## Run the web UI from source

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app/app.py
```

## Releasing

Push a tag matching `v*` (or `V*`) and GitHub Actions will build the
cross-platform portable zip and publish it on a Release:

```bash
git tag v1.0.0
git push origin v1.0.0
```
