"""Konversi .xlsx -> .pdf lewat CloudConvert API (pengganti LibreOffice, yang
tidak tersedia di lingkungan serverless Vercel).

Butuh environment variable CLOUDCONVERT_API_KEY (Project Settings > Environment
Variables di dashboard Vercel atau .env lokal). Dapatkan API key di https://cloudconvert.com/dashboard/api/v2/keys
"""
import base64
import os
import time
from pathlib import Path
import requests

CLOUDCONVERT_API = "https://api.cloudconvert.com/v2"


def convert_xlsx_to_pdf(xlsx_bytes: bytes, filename: str = "laporan.xlsx", timeout: int = 60) -> bytes:
    api_key = os.environ.get("CLOUDCONVERT_API_KEY")
    if not api_key:
        raise RuntimeError(
            "CLOUDCONVERT_API_KEY belum diset di environment variables. "
            "Tambahkan di Project Settings > Environment Variables Vercel atau file .env."
        )
    headers = {"Authorization": f"Bearer {api_key}"}

    clean_filename = Path(filename).name if filename else "laporan.xlsx"
    out_pdf_name = clean_filename.rsplit(".", 1)[0] + ".pdf"
    b64_content = base64.b64encode(xlsx_bytes).decode("ascii")

    # 1) Buat job langsung dengan payload base64 (tanpa multipart upload terpisah yang rentan timeout/gagal)
    job_payload = {
        "tasks": {
            "import-file": {
                "operation": "import/base64",
                "file": b64_content,
                "filename": clean_filename,
            },
            "convert-file": {
                "operation": "convert",
                "input": "import-file",
                "output_format": "pdf",
                "engine": "libreoffice",
                "filename": out_pdf_name,
                "pdf_a": False,
            },
            "export-file": {
                "operation": "export/url",
                "input": "convert-file",
            },
        }
    }
    r = requests.post(f"{CLOUDCONVERT_API}/jobs", json=job_payload, headers=headers, timeout=30)
    if not r.ok:
        err_msg = r.text
        try:
            err_json = r.json()
            err_msg = err_json.get("message") or err_msg
        except Exception:
            pass
        raise RuntimeError(f"CloudConvert error ({r.status_code}): {err_msg}")

    job = r.json()["data"]
    job_id = job["id"]

    # 2) Poll status job sampai selesai
    deadline = time.time() + timeout
    export_url = None
    while time.time() < deadline:
        jr = requests.get(f"{CLOUDCONVERT_API}/jobs/{job_id}", headers=headers, timeout=30)
        jr.raise_for_status()
        jdata = jr.json()["data"]
        status = jdata.get("status")

        if status == "finished":
            export_task = next((t for t in jdata.get("tasks", []) if t.get("name") == "export-file"), None)
            if export_task and export_task.get("result", {}).get("files"):
                export_url = export_task["result"]["files"][0]["url"]
            break

        if status == "error":
            # Ambil task penyebab utama kegagalan (bukan downstream task "Input task has failed")
            failed_tasks = [t for t in jdata.get("tasks", []) if t.get("status") == "error"]
            root_task = next(
                (t for t in failed_tasks if "Input task has failed" not in (t.get("message") or "")),
                None
            )
            failed = root_task or (failed_tasks[0] if failed_tasks else None)
            msg = failed.get("message", "Konversi gagal.") if failed else "Konversi gagal."
            raise RuntimeError(f"CloudConvert error: {msg}")

        time.sleep(1.5)

    if not export_url:
        raise RuntimeError("Konversi ke PDF melebihi batas waktu (timeout).")

    # 3) Unduh hasil PDF
    pdf_resp = requests.get(export_url, timeout=30)
    pdf_resp.raise_for_status()
    return pdf_resp.content
