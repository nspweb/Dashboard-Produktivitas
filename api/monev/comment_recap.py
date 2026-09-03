import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from flask import Flask, request, jsonify, Response

from _lib.monev_report import generate_comment_recap

app = Flask(__name__)


@app.route("/", defaults={"path": ""}, methods=["POST", "GET"])
@app.route("/<path:path>", methods=["POST", "GET"])
def handler(path):
    if request.method == "GET":
        return jsonify({"ok": True, "message": "POST JSON ke endpoint ini untuk membuat rekap komentar."})

    body = request.get_json(silent=True) or {}
    rows = body.get("rows") or []
    info_col = body.get("info_col")
    comment_col_map = body.get("comment_col_map") or {}
    program_name = body.get("program_name") or "Program"

    if not rows:
        return jsonify({"error": "Tidak ada responden pada data terfilter."}), 400
    if not comment_col_map:
        return jsonify({"error": "Tidak ada kolom komentar yang terdeteksi."}), 400

    try:
        xlsx_bytes = generate_comment_recap(rows, info_col, comment_col_map)
    except Exception as e:
        return jsonify({"error": f"Gagal membuat rekap komentar: {e}"}), 500

    fname = f"Komentar_Peserta_{program_name[:40].strip().replace(' ', '_')}.xlsx"
    resp = Response(xlsx_bytes, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["Content-Disposition"] = f'attachment; filename="{fname}"'
    return resp
