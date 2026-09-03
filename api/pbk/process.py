import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
from flask import Flask, request, jsonify

from _lib.pbk_clean import build_clean_df, Q_COLS

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024

EXPORT_COLS = ["timestamp", "nik", "nik_clean", "nama", "telp", "gender", "kabkota",
               "program", "tgl_mulai_fmt", "tgl_selesai_fmt", "rata_rata_skor"] + Q_COLS


def _clean_one(file_storage, is_pretest: bool):
    raw = pd.read_excel(file_storage)
    df = build_clean_df(raw, is_pretest=is_pretest)
    df = df[[c for c in EXPORT_COLS if c in df.columns]]
    return df.to_dict(orient="records")


@app.route("/", defaults={"path": ""}, methods=["POST", "GET"])
@app.route("/<path:path>", methods=["POST", "GET"])
def handler(path):
    if request.method == "GET":
        return jsonify({"ok": True, "message": "POST file pre_file dan/atau post_file (.xlsx) ke endpoint ini."})

    pre_file = request.files.get("pre_file")
    post_file = request.files.get("post_file")

    if not pre_file and not post_file:
        return jsonify({"error": "Unggah minimal salah satu file (Pre-Test atau Post-Test)."}), 400

    result = {}
    try:
        if pre_file:
            result["pre"] = _clean_one(pre_file, is_pretest=True)
        if post_file:
            result["post"] = _clean_one(post_file, is_pretest=False)
    except Exception as e:
        return jsonify({"error": f"Gagal memproses file: {e}"}), 400

    return jsonify(result)
