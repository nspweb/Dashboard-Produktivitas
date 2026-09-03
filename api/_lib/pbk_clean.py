"""Pembersihan data Pre-Test / Post-Test, diporting dari dashboard_pbk.py
tanpa perubahan logika (deteksi kolom, normalisasi nama, parsing tanggal)."""
import re
import numpy as np
import pandas as pd

Q_COLS = [f"q{i}" for i in range(1, 11)]

QUESTION_SHORT = {f"q{i}": f"A{i}" for i in range(1, 11)}

QUESTION_LABELS = {
    "q1": "1. Pengetahuan terhadap materi pelatihan",
    "q2": "2. Wawasan terhadap perkembangan bidang kejuruan",
    "q3": "3. Keterampilan terhadap topik/tema materi",
    "q4": "4. Kemampuan mengidentifikasi kebutuhan kompetensi dunia kerja",
    "q5": "5. Kemampuan memanajemen waktu",
    "q6": "6. Kepercayaan diri menghadapi persaingan dunia kerja",
    "q7": "7. Motivasi meningkatkan kompetensi & pengembangan diri",
    "q8": "8. Kesiapan menghadapi dunia kerja",
    "q9": "9. Daya kreativitas menyelesaikan aktivitas kegiatan",
    "q10": "10. Efektivitas & efisiensi melaksanakan aktivitas kegiatan",
}

QUESTION_LABELS_EXPORT_PRE = {
    "q1": "1. Pengetahuan saya terhadap materi pelatihan",
    "q2": "2. Wawasan saya terhadap perkembangan bidang kejuruan pelatihan.",
    "q3": "3. Keterampilan saya terhadap topik / tema materi pelatihan",
    "q4": "4. Kemampuan saya dalam mengidentifikasi kebutuhan kompetensi dalam dunia kerja",
    "q5": "5. Kemampuan saya dalam memanajemen waktu.",
    "q6": "6. Kepercayaan diri saya dalam menghadapi persaingan dunia kerja.",
    "q7": "7. Motivasi saya dalam meningkatkan kompetensi dan pengembangan diri.",
    "q8": "8. Kesiapan saya dalam menghadapi dunia kerja",
    "q9": "9. Daya kreativitas saya dalam menyelesaikan aktivitas kegiatan sesuai bidang kejuruan yang sedang didalami.",
    "q10": "10. Efektivitas dan efisiensi saya dalam melaksanakan aktivitas kegiatan sesuai bidang kejuruan yang didalami",
}

QUESTION_LABELS_EXPORT_POST = {
    "q1": "1. Pengetahuan saya terhadap materi pelatihan.",
    "q2": "2. Wawasan saya terhadap perkembangan bidang Kejuruan Pelatihan.",
    "q3": "3. Keterampilan saya terhadap topik / tema materi pelatihan.",
    "q4": "4. Kemampuan saya dalam mengidentifikasi kebutuhan Kompetensi dalam Dunia Kerja",
    "q5": "5. Kemampuan saya dalam memanajemen waktu.",
    "q6": "6. Kepercayaan diri saya dalam menghadapi persaingan Dunia Kerja",
    "q7": "7. Motivasi saya dalam meningkatkan kompetensi dan pengembangan diri",
    "q8": "8. Kesiapan saya dalam menghadapi Dunia Kerja.",
    "q9": "9. Daya kreativitas saya dalam menyelesaikan aktivitas kegiatan sesuai bidang kejuruan yang sedang didalami.",
    "q10": "10. Efektivitas dan efisiensi saya dalam melaksanakan aktivitas kegiatan sesuai bidang kejuruan yang didalami.",
}

QUESTION_KEYS = {
    "q1": "pengetahuan saya terhadap materi pelatihan",
    "q2": "wawasan saya terhadap perkembangan bidang",
    "q3": "keterampilan saya terhadap topik",
    "q4": "kemampuan saya dalam mengidentifikasi kebutuhan kompetensi",
    "q5": "kemampuan saya dalam memanajemen waktu",
    "q6": "kepercayaan diri saya dalam menghadapi persaingan",
    "q7": "motivasi saya dalam meningkatkan kompetensi",
    "q8": "kesiapan saya dalam menghadapi dunia kerja",
    "q9": "daya kreativitas saya dalam menyelesaikan aktivitas",
    "q10": "efektivitas dan efisiensi saya dalam melaksanakan",
}

META_KEYS = {
    "timestamp": "timestamp",
    "nik": "nik",
    "nama": "nama lengkap",
    "telp": "tlp",
    "gender": "jenis kelamin",
    "kabkota": "kabupaten",
    "program": "program pelatihan",
    "tgl_mulai": "tanggal mulai",
    "tgl_selesai": "tanggal berakhir",
}

EXPORT_META_PRE = {
    "nik": "NIK",
    "nama": "Nama Lengkap",
    "telp": "No. TLP / Whatsapp",
    "gender": "Jenis Kelamin",
    "kabkota": "Kabupaten / Kota",
    "program": "Program Pelatihan Yang Diikuti",
    "tgl_mulai": "Tanggal Mulai",
    "tgl_selesai": "Tanggal Berakhir",
}

EXPORT_META_POST = {
    "nama": "Nama Lengkap",
    "nik": "NIK",
    "program": "Program Pelatihan Yang Diikuti",
}

_NAME_PARTICLES = {"bin", "binti", "van", "de", "al"}


def _norm(text: str) -> str:
    t = str(text).strip().lower()
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"[.:]+$", "", t)
    return t


def _find_column(df: pd.DataFrame, keyword: str):
    for col in df.columns:
        if keyword in _norm(col):
            return col
    return None


def titlecase_name(name) -> str:
    if pd.isna(name):
        return name
    text = str(name).strip()
    if not text:
        return text
    words = re.split(r"(\s+)", text.lower())
    out = []
    for w in words:
        if w.isspace() or w == "":
            out.append(w)
            continue
        if w in _NAME_PARTICLES:
            out.append(w)
        else:
            out.append("-".join(p[:1].upper() + p[1:] for p in w.split("-") if p != "") or w)
    return "".join(out)


def build_clean_df(raw: pd.DataFrame, is_pretest: bool) -> pd.DataFrame:
    colmap = {}
    for key, kw in META_KEYS.items():
        colmap[key] = _find_column(raw, kw)
    for key, kw in QUESTION_KEYS.items():
        colmap[key] = _find_column(raw, kw)

    out = pd.DataFrame()
    for key in ["timestamp", "nik", "nama", "telp", "gender", "kabkota",
                "program", "tgl_mulai", "tgl_selesai"] + Q_COLS:
        src = colmap.get(key)
        out[key] = raw[src] if src is not None else np.nan

    out["nik_clean"] = (
        out["nik"].astype(str).str.replace(r"\.0$", "", regex=True)
        .str.strip().str.replace(r"\s+", "", regex=True)
    )
    out.loc[out["nik_clean"].isin(["nan", "None", ""]), "nik_clean"] = np.nan

    for q in Q_COLS:
        out[q] = pd.to_numeric(out[q], errors="coerce")

    out["rata_rata_skor"] = out[Q_COLS].mean(axis=1, skipna=True)
    out["jenis_survei"] = "Pre-Test" if is_pretest else "Post-Test"

    out["nama"] = out["nama"].apply(titlecase_name)
    if "gender" in out:
        out["gender"] = out["gender"].astype(str).str.strip().str.title().replace("Nan", np.nan)
    if "kabkota" in out:
        out["kabkota"] = out["kabkota"].astype(str).str.strip().str.title().replace("Nan", np.nan)
    if "program" in out:
        out["program"] = out["program"].astype(str).str.strip().replace("nan", np.nan)
    if "telp" in out:
        out["telp"] = out["telp"].astype(str).str.strip().replace("nan", np.nan)

    def _parse_date_col(series):
        s = series.astype(str).str.strip()
        is_iso = s.str.match(r"^\d{4}-\d{1,2}-\d{1,2}")
        parsed = pd.to_datetime(s.where(~is_iso), errors="coerce", dayfirst=True)
        parsed_iso = pd.to_datetime(s.where(is_iso), errors="coerce", dayfirst=False)
        parsed = parsed.fillna(parsed_iso)
        return parsed

    for c in ["tgl_mulai", "tgl_selesai"]:
        parsed = _parse_date_col(out[c])
        out[c + "_fmt"] = parsed.dt.strftime("%Y-%m-%d")
        out[c + "_fmt"] = out[c + "_fmt"].fillna(out[c].astype(str).replace("nan", np.nan))

    out = out.replace({np.nan: None})
    return out
