"""
Excel Merger — Web Application
Requirements: flask, pandas, openpyxl
Run: python app.py   →  opens http://localhost:5050 automatically
"""

import io
import os
import uuid
import threading
import webbrowser

from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd

app = Flask(__name__)

EXCEL_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".xlsb"}

# In-memory store for merged files (token → bytes)
_store: dict[str, bytes] = {}
_lock = threading.Lock()


# ── Helpers ────────────────────────────────────────────────────────────── #

def _read_excel(raw_bytes: bytes, sheet_name, skip_rows: int) -> tuple[pd.DataFrame, str | None]:
    """
    Read an Excel file with three progressive fallback strategies.

    Some files (especially those exported by non-Microsoft tools) contain
    cells whose XML type is "n" (numeric) but whose stored value is a date
    string like "2020-09-03".  openpyxl's internal int() cast then blows up.

    Strategy 1 – standard pandas read (fastest, works for well-formed files)
    Strategy 2 – openpyxl data_only=True  (reads cached values, skips bad types)
    Strategy 3 – no dtype enforcement then stringify  (last resort)

    Returns (DataFrame, warning_message_or_None).
    """
    def _try(buf, **kwargs) -> pd.DataFrame:
        return pd.read_excel(buf, sheet_name=sheet_name,
                             skiprows=skip_rows, **kwargs)

    # ① Standard
    try:
        return _try(io.BytesIO(raw_bytes), dtype=str), None
    except Exception as e1:
        pass

    # ② data_only – bypasses malformed numeric/date cells
    try:
        df = _try(io.BytesIO(raw_bytes), dtype=str,
                  engine="openpyxl",
                  engine_kwargs={"data_only": True})
        return df, "⚠  Read in data_only mode (file has malformed cells)"
    except Exception as e2:
        pass

    # ③ No dtype, then stringify (openpyxl lets pandas infer types first)
    try:
        df = _try(io.BytesIO(raw_bytes),
                  engine="openpyxl",
                  engine_kwargs={"data_only": True})
        df = df.astype(str).replace({"nan": "", "NaT": "", "None": ""})
        return df, "⚠  Read without strict dtype (file has severely malformed cells)"
    except Exception as e3:
        # All strategies failed — re-raise the original error
        raise e1 from None


def _col_letter_to_index(letter: str) -> int:
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _resolve_skip_columns(raw: str, columns):
    if not raw.strip():
        return [], []
    cols = list(columns)
    to_drop, warnings = [], []
    for token in raw.split(","):
        token = token.strip()
        if not token:
            continue
        if token.replace(" ", "").isalpha() and len(token) <= 3:
            idx = _col_letter_to_index(token)
            if 0 <= idx < len(cols):
                to_drop.append(cols[idx])
            else:
                warnings.append(f"Column letter '{token}' out of range")
        elif token.lstrip("-").isdigit():
            idx = int(token)
            if 0 <= idx < len(cols):
                to_drop.append(cols[idx])
            else:
                warnings.append(f"Column index {idx} out of range")
        elif token in cols:
            to_drop.append(token)
        else:
            warnings.append(f"Column '{token}' not found")
    return to_drop, warnings


# ── Routes ─────────────────────────────────────────────────────────────── #

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/merge", methods=["POST"])
def merge():
    files          = request.files.getlist("files")
    sheet_name     = request.form.get("sheet_name", "").strip() or 0
    skip_cols_raw  = request.form.get("skip_cols", "").strip()
    remove_dups    = request.form.get("remove_duplicates") == "true"
    drop_empty     = request.form.get("drop_empty") == "true"

    try:
        skip_rows = int(request.form.get("skip_rows", 0) or 0)
    except ValueError:
        return jsonify({"success": False, "error": "Skip Rows must be a number.", "log": []})

    valid = [f for f in files if f.filename]
    if not valid:
        return jsonify({"success": False, "error": "No files were uploaded.", "log": []})

    log = []

    def entry(kind, msg):
        log.append({"k": kind, "m": msg})

    entry("head", f"Merging {len(valid)} file(s)")
    entry("meta", f"Sheet › {'first' if sheet_name == 0 else sheet_name}  |  Skip rows › {skip_rows}")
    if skip_cols_raw:
        entry("meta", f"Skip cols › {skip_cols_raw}")
    entry("meta", f"Dedup › {'on' if remove_dups else 'off'}  |  Drop empty › {'on' if drop_empty else 'off'}")
    entry("div",  "")

    frames, errors = [], []

    for f in valid:
        entry("file", f.filename)
        try:
            raw_bytes = f.read()
            df, read_warn = _read_excel(raw_bytes, sheet_name, skip_rows)
            if read_warn:
                entry("warn", f"  {read_warn}")

            if drop_empty:
                before = len(df)
                df.dropna(how="all", inplace=True)
                dropped = before - len(df)
                if dropped:
                    entry("info", f"  Dropped {dropped:,} fully-empty row(s)")

            to_drop, warns = _resolve_skip_columns(skip_cols_raw, df.columns)
            for w in warns:
                entry("warn", f"  ⚠  {w}")
            if to_drop:
                df.drop(columns=to_drop, inplace=True, errors="ignore")
                entry("info", f"  Removed cols: {', '.join(str(c) for c in to_drop)}")

            df.insert(0, "Source File", f.filename)
            frames.append(df)
            entry("ok", f"  ✓  {len(df):,} rows  ×  {len(df.columns) - 1:,} data columns")

        except Exception as exc:
            errors.append(f.filename)
            entry("err", f"  ✗  {exc}")

    if not frames:
        return jsonify({"success": False, "error": "No data could be read from any file.", "log": log})

    entry("div", "")
    entry("info", f"Concatenating {len(frames)} frame(s)…")
    merged = pd.concat(frames, ignore_index=True, sort=False)
    entry("info", f"Combined total: {len(merged):,} rows")

    if remove_dups:
        before = len(merged)
        data_cols = [c for c in merged.columns if c != "Source File"]
        merged.drop_duplicates(subset=data_cols, inplace=True)
        removed = before - len(merged)
        entry("info", f"Duplicates removed: {removed:,}  →  {len(merged):,} rows kept")

    buf = io.BytesIO()
    merged.to_excel(buf, index=False)
    buf.seek(0)

    token = str(uuid.uuid4())
    with _lock:
        _store[token] = buf.read()

    entry("ok", f"✓  Complete — {len(merged):,} rows from {len(frames)} file(s)")

    return jsonify({
        "success": True,
        "token":   token,
        "rows":    len(merged),
        "files":   len(frames),
        "errors":  len(errors),
        "log":     log,
    })


@app.route("/download/<token>")
def download(token):
    with _lock:
        data = _store.pop(token, None)
    if data is None:
        return "File not found or already downloaded.", 404
    return send_file(
        io.BytesIO(data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="merged_output.xlsx",
    )


# ── Entry point ────────────────────────────────────────────────────────── #

if __name__ == "__main__":
    port = 5050
    threading.Timer(1.2, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    app.run(debug=False, port=port, use_reloader=False)
