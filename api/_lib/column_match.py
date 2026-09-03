"""Utilitas pencocokan kolom & definisi pertanyaan kanonik untuk sistem
REKAP MONEV (evaluasi penyelenggaraan pelatihan). Diporting dari app.py asli
(Streamlit) tanpa perubahan logika."""
import re
import difflib


def normalize(text: str) -> str:
    text = str(text).lower().replace("\n", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def best_match(canonical: str, columns, cutoff: float = 0.55):
    norm_canonical = normalize(canonical)
    norm_cols = {c: normalize(c) for c in columns}
    for original, norm in norm_cols.items():
        if norm == norm_canonical:
            return original
    for original, norm in norm_cols.items():
        if norm_canonical in norm or norm in norm_canonical:
            return original
    matches = difflib.get_close_matches(norm_canonical, list(norm_cols.values()), n=1, cutoff=cutoff)
    if matches:
        for original, norm in norm_cols.items():
            if norm == matches[0]:
                return original
    return None


PROGRAM_CANONICAL = "Program pelatihan yang diikuti"
TIMESTAMP_CANONICAL = "Timestamp"
INFO_SOURCE_CANONICAL = "Dari mana anda mendapatkan informasi tentang pelatihan"

SCORE_QUESTIONS = {
    "Info mudah didapat": "Apakah informasi pelatihan mudah untuk didapatkan?",
    "Pendaftaran mudah": "Apakah pendaftaran dan tahapannya mudah untuk dilakukan?",
    "Petunjuk pendaftaran jelas": "Apakah petunjuk tata cara pendaftaran jelas dan mudah dipahami?",
    "Program jelas": "Apakah program pelatihan jelas dan mudah dipahami?",
    "Program menarik": "Apakah program pelatihan menarik?",
    "Program bermanfaat": "Apakah program pelatihan bermanfaat?",
    "Kompetensi meningkat": "Apakah program pelatihan berhasil meningkatkan kompetensi Anda?",
    "Durasi sesuai": "Apakah durasi untuk menyelesaikan pelatihan sudah sesuai?",
    "Instruktur menguasai materi": "Apakah Instrukur menguasai program pelatihan yang disampaikan?",
    "Kemampuan penyampaian instruktur": "Bagaimana kemampuan Instruktur dalam menyampaikan program pelatihan?",
    "Kemampuan mengelola peserta": "Bagaimana kemampuan Instruktur dalam mengelola Peserta Pelatihan?",
    "Sikap & teladan instruktur": "Bagaimana sikap, disiplin, penampilan, dan teladan Instruktur selama pelatihan?",
    "Pelayanan petugas": "Bagaimana pelayanan petugas terhadap Peserta Pelatihan?",
    "Jadwal sesuai rencana": "Apakah pelaksanaan jadwal pelatihan sudah sesuai dengan rencana?",
    "Perlengkapan tepat waktu": "Apakah perlengkapan Peserta Pelatihan training material diberikan tepat waktu",
    "Sarana pelatihan memadai": "Apakah sarana prasarana fasilitas pelatihan sudah memadai",
    "Sarana penunjang memadai": "Apakah sarana prasarana fasilitas penunjang pelatihan sudah memadai",
}

COMMENT_LABELS = {
    "Komentar/saran program pelatihan": "Komentar dan saran Anda terhadap program pelatihan",
    "Komentar/saran instruktur": "Komentar dan saran Anda terhadap Instruktur",
    "Komentar/saran penyelenggaraan": "Komentar dan saran Anda terhadap penyelenggaraan pelatihan",
    "Keluhan penyelenggaraan": "Keluhan terhadap penyelenggaraan pelatihan",
}


def process_dataframe(df_raw):
    """Deteksi kolom program/timestamp/info/skor/komentar pada df mentah.
    Mengembalikan dict siap dipakai endpoint /api/monev/process."""
    import pandas as pd

    all_columns = list(df_raw.columns)

    program_col = best_match(PROGRAM_CANONICAL, all_columns)
    if not program_col:
        raise ValueError(
            "Kolom 'Program pelatihan yang diikuti' tidak ditemukan di sheet ini. "
            "Pastikan file yang diupload adalah hasil Google Form evaluasi pelatihan."
        )
    timestamp_col = best_match(TIMESTAMP_CANONICAL, all_columns)
    info_col = best_match(INFO_SOURCE_CANONICAL, all_columns)

    score_col_map = {}
    for label, question in SCORE_QUESTIONS.items():
        match = best_match(question, all_columns)
        if match:
            score_col_map[label] = match

    comment_col_map = {}
    for label, question in COMMENT_LABELS.items():
        match = best_match(question, all_columns)
        if match:
            comment_col_map[label] = match

    for label, col in list(score_col_map.items()):
        numeric_series = pd.to_numeric(df_raw[col], errors="coerce")
        if numeric_series.notna().sum() == 0:
            del score_col_map[label]
        else:
            df_raw[col] = numeric_series

    if timestamp_col:
        df_raw[timestamp_col] = pd.to_datetime(df_raw[timestamp_col], errors="coerce").astype(str)

    df_raw = df_raw.where(pd.notna(df_raw), None)

    return {
        "program_col": program_col,
        "timestamp_col": timestamp_col,
        "info_col": info_col,
        "score_col_map": score_col_map,
        "comment_col_map": comment_col_map,
        "rows": df_raw.to_dict(orient="records"),
    }
