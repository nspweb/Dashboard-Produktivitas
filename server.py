"""Local Development Server untuk Dasbor Produktivitas & Monev BPVP.
Menjalankan frontend statis sekaligus semua API endpoint secara lokal tanpa perlu Vercel CLI.
"""
import os
import sys
from pathlib import Path
from flask import Flask, send_file, jsonify

ROOT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT_DIR))
sys.path.insert(0, str(ROOT_DIR / "api"))

# Load .env file jika ada
env_file = ROOT_DIR / ".env"
if env_file.exists():
    with open(env_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip("\"'"))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# Import handlers from api modules
from api.monev.process import handler as monev_process_handler
from api.monev.report import handler as monev_report_handler
from api.monev.comment_recap import handler as monev_comment_recap_handler
from api.pbk.process import handler as pbk_process_handler
from api.pbk.report import handler as pbk_report_handler
from api.convert_pdf import handler as convert_pdf_handler


# API Routes
@app.route("/api/monev/process", methods=["GET", "POST"])
def route_monev_process():
    return monev_process_handler("")


@app.route("/api/monev/report", methods=["GET", "POST"])
def route_monev_report():
    return monev_report_handler("")


@app.route("/api/monev/comment_recap", methods=["GET", "POST"])
def route_monev_comment_recap():
    return monev_comment_recap_handler("")


@app.route("/api/pbk/process", methods=["GET", "POST"])
def route_pbk_process():
    return pbk_process_handler("")


@app.route("/api/pbk/report", methods=["GET", "POST"])
def route_pbk_report():
    return pbk_report_handler("")


@app.route("/api/convert_pdf", methods=["GET", "POST"])
def route_convert_pdf():
    return convert_pdf_handler("")


# Static file routes
@app.route("/")
def index():
    return send_file(ROOT_DIR / "index.html")


@app.route("/<path:path>")
def static_files(path):
    target = ROOT_DIR / path
    if target.exists() and target.is_file():
        return send_file(target)
    return jsonify({"error": "File not found"}), 404


if __name__ == "__main__":
    port = 5000
    print("\n" + "=" * 58)
    print("  Dasbor Produktivitas & Monev BPVP - Local Server")
    print(f"  Buka di browser: http://localhost:{port}")
    print("=" * 58 + "\n")
    app.run(host="0.0.0.0", port=port, debug=True)
