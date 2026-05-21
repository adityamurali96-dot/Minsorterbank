"""Local Flask frontend for sort_statement.

Run via the Start-Minsorterbank launcher for your OS, or directly:
`python app/app.py`. On startup it picks a free port, starts the Flask
server, then opens the default browser at the local URL.

Endpoints:
  GET  /            single-page upload UI
  POST /api/sort    multipart upload -> returns the sorted .xlsx
"""

from __future__ import annotations

import io
import os
import socket
import sys
import tempfile
import threading
import time
import traceback
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))

import sort_statement  # noqa: E402


app = Flask(
    __name__,
    template_folder=str(Path(__file__).resolve().parent / "templates"),
    static_folder=str(Path(__file__).resolve().parent / "static"),
)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


_PROFILE_BY_NAME = {p.name: p for p in sort_statement.PROFILES}


def _resolve_profile(choice: str, in_path):
    """Pick a profile from the user's dropdown choice, falling back to auto-detect."""
    if choice and choice != "auto":
        prof = _PROFILE_BY_NAME.get(choice)
        if prof is not None:
            return prof
    return sort_statement.detect_profile(in_path)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/sort", methods=["POST"])
def api_sort():
    if "file" not in request.files:
        return jsonify({"error": "No file provided."}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename."}), 400

    suffix = Path(f.filename).suffix or ".xls"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        f.save(tmp.name)
        in_path = Path(tmp.name)

    try:
        bank_choice = (request.form.get("bank") or "auto").strip().lower()
        profile = _resolve_profile(bank_choice, in_path)
        df = profile.parse(in_path)
        extracted = int(df["counterparty"].notna().sum())

        buf = io.BytesIO()
        out_path = in_path.with_suffix(".sorted.xlsx")
        sort_statement.write_workbook(df, out_path, bank_name=profile.name)
        buf.write(out_path.read_bytes())
        buf.seek(0)

        download_name = Path(f.filename).stem + "_sorted.xlsx"

        try:
            out_path.unlink()
        except OSError:
            pass

        response = send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=download_name,
        )
        response.headers["X-Bank-Profile"] = profile.name
        response.headers["X-Row-Count"] = str(len(df))
        response.headers["X-Extracted"] = str(extracted)
        response.headers["Access-Control-Expose-Headers"] = (
            "X-Bank-Profile, X-Row-Count, X-Extracted, Content-Disposition"
        )
        return response

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500
    finally:
        try:
            in_path.unlink()
        except OSError:
            pass


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _open_browser(url: str, delay: float = 1.2):
    def go():
        time.sleep(delay)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Thread(target=go, daemon=True).start()


def main():
    port = int(os.environ.get("MINSORTER_PORT", "0")) or _free_port()
    url = f"http://127.0.0.1:{port}"
    print(f"\n  Minsorterbank running at  {url}")
    print("  (Close this window to quit.)\n")
    if os.environ.get("MINSORTER_NO_BROWSER") != "1":
        _open_browser(url)
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)


if __name__ == "__main__":
    main()
