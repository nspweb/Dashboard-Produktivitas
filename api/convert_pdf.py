import base64
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from flask import Flask, request, jsonify, Response

from _lib.pdf_convert import convert_xlsx_to_pdf

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024


@app.route("/", defaults={"path": ""}, methods=["POST", "GET"])
@app.route("/<path:path>", methods=["POST", "GET"])
def handler(path):
    if request.method == "GET":
        return jsonify({"ok": True, "message": "POST JSON {xlsx_base64, filename} ke endpoint ini."})

    body = request.get_json(silent=True) or {}
    xlsx_b64 = body.get("xlsx_base64")
    filename = body.get("filename") or "laporan.xlsx"

    if not xlsx_b64:
        return jsonify({"error": "xlsx_base64 tidak ada pada request."}), 400

    try:
        xlsx_bytes = base64.b64decode(xlsx_b64)
        pdf_bytes = convert_xlsx_to_pdf(xlsx_bytes, filename=filename)
    except RuntimeError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        return jsonify({"error": f"Gagal mengonversi ke PDF: {e}"}), 500

    out_name = Path(filename).stem + ".pdf"
    resp = Response(pdf_bytes, mimetype="application/pdf")
    resp.headers["Content-Disposition"] = f'attachment; filename="{out_name}"'
    return resp
