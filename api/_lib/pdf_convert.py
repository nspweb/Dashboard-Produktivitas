"""Konversi .xlsx -> .pdf lewat CloudConvert API (pengganti LibreOffice, yang
tidak tersedia di lingkungan serverless Vercel).

Butuh environment variable CLOUDCONVERT_API_KEY (Project Settings > Environment
Variables di dashboard Vercel). Dapatkan API key di https://cloudconvert.com/dashboard/api/v2/keys
"""
import os
import time
import requests

CLOUDCONVERT_API = "https://api.cloudconvert.com/v2"


def convert_xlsx_to_pdf(xlsx_bytes: bytes, filename: str = "laporan.xlsx", timeout: int = 60) -> bytes:
    api_key = os.environ.get("CLOUDCONVERT_API_KEY")
    if not api_key:
        raise RuntimeError(
            "CLOUDCONVERT_API_KEY belum diset di environment variables Vercel. "
            "Tambahkan di Project Settings > Environment Variables, lalu redeploy."
        )
    headers = {"Authorization": f"Bearer {api_key}"}

    # 1) Buat job: import (upload) -> convert -> export (url)
    job_payload = {
        "tasks": {
            "import-file": {"operation": "import/upload"},
            "convert-file": {
                "operation": "convert",
                "input": "import-file",
                "output_format": "pdf",
                "engine": "libreoffice",
                "filename": "laporan.pdf",
                "pdf_a": False,
            },
            "export-file": {"operation": "export/url", "input": "convert-file"},
        }
    }
    r = requests.post(f"{CLOUDCONVERT_API}/jobs", json=job_payload, headers=headers, timeout=30)
    r.raise_for_status()
    job = r.json()["data"]

    import_task = next(t for t in job["tasks"] if t["name"] == "import-file")
    upload_form = import_task["result"]["form"]

    # 2) Upload file mentah ke URL yang diberikan CloudConvert
    files = {"file": (filename, xlsx_bytes)}
    up = requests.post(upload_form["url"], data=upload_form["parameters"], files=files, timeout=60)
    up.raise_for_status()

    # 3) Poll status job sampai selesai
    job_id = job["id"]
    deadline = time.time() + timeout
    export_url = None
    while time.time() < deadline:
        jr = requests.get(f"{CLOUDCONVERT_API}/jobs/{job_id}", headers=headers, timeout=30)
        jr.raise_for_status()
        jdata = jr.json()["data"]
        status = jdata["status"]
        if status == "finished":
            export_task = next(t for t in jdata["tasks"] if t["name"] == "export-file")
            export_url = export_task["result"]["files"][0]["url"]
            break
        if status == "error":
            failed = next((t for t in jdata["tasks"] if t["status"] == "error"), None)
            msg = failed.get("message", "Konversi gagal.") if failed else "Konversi gagal."
            raise RuntimeError(f"CloudConvert error: {msg}")
        time.sleep(1.5)

    if not export_url:
        raise RuntimeError("Konversi ke PDF melebihi batas waktu (timeout).")

    # 4) Unduh hasil PDF
    pdf_resp = requests.get(export_url, timeout=30)
    pdf_resp.raise_for_status()
    return pdf_resp.content
