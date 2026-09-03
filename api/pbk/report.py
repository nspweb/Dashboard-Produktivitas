import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from flask import Flask, request, jsonify, Response

from _lib.pbk_clean import Q_COLS
from _lib.kolom_isian_report import generate_kolom_isian_report

app = Flask(__name__)

TEMPLATE_PATH = Path(__file__).resolve().parents[2] / "templates" / "laporan_kolom_isian_template.xlsx"


@app.route("/", defaults={"path": ""}, methods=["POST", "GET"])
@app.route("/<path:path>", methods=["POST", "GET"])
def handler(path):
    if request.method == "GET":
        return jsonify({"ok": True, "message": "POST JSON ke endpoint ini untuk membuat laporan Pre/Post-Test."})

    if not TEMPLATE_PATH.exists():
        return jsonify({
            "error": "Template laporan_kolom_isian_template.xlsx belum ada di folder /templates "
                     "pada proyek ini. Salin file template resmi Anda ke situ sebelum deploy."
        }), 500

    body = request.get_json(silent=True) or {}
    rows = body.get("rows") or []
    jenis_test = body.get("jenis_test") or "Pre-Test"
    training_title_input = body.get("training_title_input") or ""
    program_name = body.get("program_name") or ""
    tanggal_mulai = body.get("tanggal_mulai")
    tanggal_selesai = body.get("tanggal_selesai")

    missing = [f for f in ["training_title_input", "program_name", "tanggal_mulai", "tanggal_selesai"] if not body.get(f)]
    if missing:
        return jsonify({"error": f"Field wajib belum diisi: {', '.join(missing)}"}), 400

    try:
        xlsx_bytes, meta = generate_kolom_isian_report(
            template_path=str(TEMPLATE_PATH),
            rows=rows,
            q_cols=Q_COLS,
            jenis_test=jenis_test,
            training_title_input=training_title_input,
            program_name=program_name,
            tanggal_mulai=tanggal_mulai,
            tanggal_selesai=tanggal_selesai,
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"Gagal membuat laporan: {e}"}), 500

    fname_base = f"Laporan_{jenis_test.replace('-', '')}_{re.sub(r'[^A-Za-z0-9]+', '_', program_name).strip('_')}"
    resp = Response(xlsx_bytes, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["Content-Disposition"] = f'attachment; filename="{fname_base}.xlsx"'
    resp.headers["X-Report-Meta"] = str(meta)
    return resp
