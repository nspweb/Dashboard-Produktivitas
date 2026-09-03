import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
from flask import Flask, request, jsonify

from _lib.column_match import process_dataframe

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024  # 15MB


def _json_safe_rows(rows):
    """Convert pandas missing values to JSON-compatible null values."""
    safe_rows = []
    for row in rows:
        safe_row = {}
        for key, value in row.items():
            try:
                missing = pd.isna(value)
                if not hasattr(missing, "__len__") and bool(missing):
                    safe_row[key] = None
                    continue
            except (TypeError, ValueError):
                pass
            safe_row[key] = value.item() if hasattr(value, "item") else value
        safe_rows.append(safe_row)
    return safe_rows


@app.route("/", defaults={"path": ""}, methods=["POST", "GET"])
@app.route("/<path:path>", methods=["POST", "GET"])
def handler(path):
    if request.method == "GET":
        return jsonify({"ok": True, "message": "POST file evaluasi (.xlsx) ke endpoint ini."})

    if "file" not in request.files:
        return jsonify({"error": "File tidak ditemukan pada request."}), 400

    file = request.files["file"]
    sheet_name = request.form.get("sheet_name") or None

    try:
        xls = pd.ExcelFile(file)
    except Exception:
        return jsonify({"error": "File tidak bisa dibaca. Pastikan formatnya .xlsx dan tidak rusak."}), 400

    sheet_names = xls.sheet_names
    if sheet_name is None and len(sheet_names) > 1:
        return jsonify({"needs_sheet_choice": True, "sheets": sheet_names})

    chosen_sheet = sheet_name or sheet_names[0]
    try:
        df_raw = pd.read_excel(xls, sheet_name=chosen_sheet)
    except Exception as e:
        return jsonify({"error": f"Gagal membaca sheet '{chosen_sheet}': {e}"}), 400

    df_raw.columns = [str(c) for c in df_raw.columns]

    try:
        result = process_dataframe(df_raw)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    result["rows"] = _json_safe_rows(result.get("rows", []))
    result["sheet_choice"] = chosen_sheet
    result["sheets"] = sheet_names
    result["total_rows"] = len(result["rows"])
    return jsonify(result)
