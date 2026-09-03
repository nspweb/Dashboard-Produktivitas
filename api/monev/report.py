import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from flask import Flask, request, jsonify, Response

from _lib.column_match import SCORE_QUESTIONS
from _lib.monev_report import generate_official_report

app = Flask(__name__)

TEMPLATE_PATH = Path(__file__).resolve().parents[2] / "templates" / "template_laporan_resmi.xlsx"


@app.route("/", defaults={"path": ""}, methods=["POST", "GET"])
@app.route("/<path:path>", methods=["POST", "GET"])
def handler(path):
    if request.method == "GET":
        return jsonify({"ok": True, "message": "POST JSON ke endpoint ini untuk membuat laporan resmi."})

    if not TEMPLATE_PATH.exists():
        return jsonify({
            "error": "Template laporan resmi (template_laporan_resmi.xlsx) belum ada di folder /templates "
                     "pada proyek ini. Salin file template resmi Anda ke situ sebelum deploy."
        }), 500

    body = request.get_json(silent=True) or {}
    rows = body.get("rows") or []
    score_col_map = body.get("score_col_map") or {}
    program_name = body.get("program_name") or ""
    training_title = body.get("training_title") or ""
    date_text = body.get("date_text") or ""
    instructor_name = body.get("instructor_name") or ""

    missing = [f for f in ["program_name", "training_title", "date_text", "instructor_name"] if not body.get(f)]
    if missing:
        return jsonify({"error": f"Field wajib belum diisi: {', '.join(missing)}"}), 400

    try:
        xlsx_bytes, meta = generate_official_report(
            str(TEMPLATE_PATH),
            rows,
            score_col_map,
            SCORE_QUESTIONS,
            program_name=program_name,
            training_title=training_title,
            date_text=date_text,
            instructor_name=instructor_name,
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Gagal membuat laporan: {e}"}), 500

    fname = f"Laporan_Evaluasi_{program_name[:40].strip().replace(' ', '_')}.xlsx"
    resp = Response(xlsx_bytes, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["Content-Disposition"] = f'attachment; filename="{fname}"'
    resp.headers["X-Report-Meta"] = str(meta)
    return resp
