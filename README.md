# Minsorterbank

Turn an Indian bank statement (ICICI / Axis / HDFC / SBI / generic Excel / CSV)
into a grouped Deposits / Withdrawals workbook.

Runs entirely on your computer &mdash; nothing is uploaded anywhere.

## Download &amp; run (everyday users)

### Step 1 &mdash; Download

1. Go to the top of this page: <https://github.com/adityamurali96-dot/Minsorterbank>
2. Click the green **&lt;&gt; Code** button.
3. In the dropdown, click **Download ZIP**.
4. A file called `Minsorterbank-main.zip` will land in your Downloads folder.

### Step 2 &mdash; Extract

- **Windows:** right-click the zip &rarr; **Extract All&hellip;** &rarr; pick a
  spot like your Desktop &rarr; **Extract**.
- **macOS:** double-click the zip. Finder will create a `Minsorterbank-main`
  folder next to it.

### Step 3 &mdash; Double-click the launcher

Open the extracted folder and double-click the launcher for your computer:

- **Windows** &rarr; `Start-Minsorterbank.bat`
- **macOS** &rarr; `Start-Minsorterbank.command`

A small Terminal / Command Prompt window will appear (leave it open) and
your browser will open the Minsorterbank page at a local address like
`http://127.0.0.1:54321`. Upload your bank statement, download the sorted
spreadsheet.

When you&rsquo;re done, close the Terminal / Command Prompt window to quit.

### First-launch notes

**Windows &mdash; if Python isn&rsquo;t installed yet:** the launcher will
open the Python download page in your browser. Click the big yellow
**Download Python** button, run the installer, and on the **first installer
screen tick &ldquo;Add python.exe to PATH&rdquo;** before clicking Install.
Then double-click `Start-Minsorterbank.bat` again.

**macOS &mdash; first-time Gatekeeper prompt:** macOS may say &ldquo;Apple
cannot check it for malicious software.&rdquo; Right-click
`Start-Minsorterbank.command` &rarr; **Open** &rarr; **Open**. After that
once, regular double-click works forever. If Python 3 isn&rsquo;t installed
yet, the launcher will pop up a dialog and open the Python download page;
install it, then double-click the file again.

The first run also installs the required Python libraries into a private
`.venv` folder next to the launcher (about 30 seconds, one-time). Every
run after that is instant.

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
