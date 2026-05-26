"""Local Flask frontend for sort_statement.

Run via the Start-Minsorterbank launcher for your OS, or directly:
`python app/app.py`. On startup it picks a free port, starts the Flask
server, then opens the default browser at the local URL.

Endpoints:
  GET  /             single-page upload UI
  POST /api/preview  multipart upload -> JSON of first rows + auto-detected
                     column mapping suggestions (for the mapping UI)
  POST /api/sort     multipart upload -> returns the sorted .xlsx; accepts
                     explicit column indices to bypass auto-detection
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
import merge_tabula  # noqa: E402


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


def _save_upload(file_storage):
    """Persist an uploaded file to a NamedTemporaryFile and return (path, suffix)."""
    suffix = Path(file_storage.filename).suffix or ".xls"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        file_storage.save(tmp.name)
        return Path(tmp.name), suffix


def _apply_hanging_merge(in_path: Path, suffix: str):
    """Merge wrapped Tabula rows if requested. Returns the path to parse from,
    plus a path to clean up (or None)."""
    if suffix.lower() != ".xlsx":
        raise ValueError("Tabula merge requires an .xlsx file.")
    merged_path = in_path.with_suffix(".merged.xlsx")
    merge_tabula.merge_statement(in_path, merged_path)
    return merged_path, merged_path


def _parse_int_field(name: str, allow_blank: bool = False):
    """Read a form field as an int. Returns None if absent (and allow_blank).
    Raises ValueError with a user-friendly message otherwise."""
    raw = (request.form.get(name) or "").strip()
    if not raw:
        if allow_blank:
            return None
        raise ValueError(f"Missing required field: {name}")
    try:
        return int(raw)
    except ValueError:
        raise ValueError(f"Invalid {name}: {raw!r}")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/preview", methods=["POST"])
def api_preview():
    if "file" not in request.files:
        return jsonify({"error": "No file provided."}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename."}), 400

    in_path, suffix = _save_upload(f)
    merged_path = None
    try:
        hanging = (request.form.get("hanging") or "no").strip().lower()
        bank_choice = (request.form.get("bank") or "auto").strip().lower()

        parse_path = in_path
        if hanging == "yes":
            try:
                parse_path, merged_path = _apply_hanging_merge(in_path, suffix)
            except ValueError as e:
                return jsonify({"error": str(e)}), 400

        _raw_text, raw_df = sort_statement._load_raw(parse_path)
        if raw_df is None or raw_df.shape[1] == 0:
            return jsonify({
                "error": "Couldn't read any tabular data from this file. "
                         "If it's a PDF export, try the 'Hanging rows' option."
            }), 422

        payload = sort_statement.build_preview(raw_df)
        profile = _resolve_profile(bank_choice, parse_path)
        payload["profile_name"] = profile.name
        return jsonify(payload)

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500
    finally:
        for p in (in_path, merged_path):
            if p is None:
                continue
            try:
                p.unlink()
            except OSError:
                pass


@app.route("/api/sort", methods=["POST"])
def api_sort():
    if "file" not in request.files:
        return jsonify({"error": "No file provided."}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename."}), 400

    in_path, suffix = _save_upload(f)

    merged_path = None
    try:
        bank_choice = (request.form.get("bank") or "auto").strip().lower()
        hanging = (request.form.get("hanging") or "no").strip().lower()

        try:
            header_row = _parse_int_field("header_row", allow_blank=True)
        except ValueError as e:
            return jsonify({"error": str(e)}), 400

        # Explicit column mapping path -- all five required if any are set.
        explicit_fields = ("data_start_row", "col_date", "col_remarks",
                           "col_withdrawal", "col_deposit")
        explicit_any = any((request.form.get(k) or "").strip() for k in explicit_fields)
        explicit_cols = None
        if explicit_any:
            try:
                explicit_cols = {
                    "data_start_row": _parse_int_field("data_start_row"),
                    "col_date": _parse_int_field("col_date"),
                    "col_remarks": _parse_int_field("col_remarks"),
                    "col_withdrawal": _parse_int_field("col_withdrawal"),
                    "col_deposit": _parse_int_field("col_deposit"),
                    "col_balance": _parse_int_field("col_balance", allow_blank=True),
                }
            except ValueError as e:
                return jsonify({"error": str(e)}), 400

        parse_path = in_path
        if hanging == "yes":
            try:
                parse_path, merged_path = _apply_hanging_merge(in_path, suffix)
            except ValueError as e:
                return jsonify({"error": str(e)}), 400

        profile = _resolve_profile(bank_choice, parse_path)

        if explicit_cols is not None:
            _raw_text, raw_df = sort_statement._load_raw(parse_path)
            if raw_df is None:
                return jsonify({
                    "error": "Couldn't read tabular data from this file."
                }), 422
            try:
                df = sort_statement.parse_with_explicit_columns(
                    raw_df,
                    data_start_row=explicit_cols["data_start_row"],
                    col_date=explicit_cols["col_date"],
                    col_remarks=explicit_cols["col_remarks"],
                    col_withdrawal=explicit_cols["col_withdrawal"],
                    col_deposit=explicit_cols["col_deposit"],
                    col_balance=explicit_cols["col_balance"],
                    extract_fn=profile.extract,
                )
            except ValueError as e:
                return jsonify({"error": str(e)}), 400
        else:
            try:
                df = profile.parse(parse_path, header_row=header_row)
            except sort_statement.HeaderNotFoundError as e:
                return jsonify({
                    "needs_header": True,
                    "error": str(e),
                    "preview": e.preview,
                }), 422
        extracted = int(df["counterparty"].notna().sum())

        buf = io.BytesIO()
        out_path = parse_path.with_suffix(".sorted.xlsx")
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
        if merged_path is not None:
            try:
                merged_path.unlink()
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
    if sys.version_info < (3, 8):
        print(
            f"\n  Minsorterbank requires Python 3.8 or later."
            f"\n  You are running Python {sys.version}."
            f"\n  Please install a newer Python from https://www.python.org/downloads/"
        )
        try:
            input("\n  Press Enter to close...")
        except EOFError:
            pass
        sys.exit(1)
    try:
        main()
    except Exception:
        traceback.print_exc()
        print("\n  Minsorterbank hit an unexpected error (see above).")
        try:
            input("  Press Enter to close...")
        except EOFError:
            pass
        sys.exit(1)
