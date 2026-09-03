"""Generator laporan resmi (format Kemnaker) untuk sistem REKAP MONEV.
Diporting dari report_generator.py asli -- logika inti TIDAK diubah, hanya
dependensi Streamlit & konversi PDF (LibreOffice) yang dihapus dari file ini."""
import re
import difflib
from pathlib import Path
from io import BytesIO
from copy import copy

import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage

PRIMARY_COLS = [get_column_letter(c) for c in range(3, 19)]   # C..R (16)
SECONDARY_COLS = [get_column_letter(c) for c in range(20, 28)]  # T..AA (8)

DEFAULT_CAPACITY = len(PRIMARY_COLS)                     # 16
MAX_RESPONDENTS = DEFAULT_CAPACITY + len(SECONDARY_COLS)  # 24

MASTER_SHEET = "KOLOM ISIAN"
VIEW_SHEETS = ["Materi Pelatihan", "Program Pelatihan", "INSTRUKTUR", "Penyelenggaraan"]

SHEET_TITLE_ROWS = {
    "KOLOM ISIAN": (6, 29),
    "Histogram": (6, 4),
    "Materi Pelatihan": (6, 29),
    "Program Pelatihan": (6, 29),
    "INSTRUKTUR": (7, 29),
    "Penyelenggaraan": (7, 29),
    "NILAI INST": (6, 6),
}

TITLE_TEXT_CELLS = {
    "Materi Pelatihan": "A6",
    "Program Pelatihan": "A6",
    "INSTRUKTUR": "A7",
    "Penyelenggaraan": "A7",
}

REPORT_TITLE = "HASIL ANALISIS EVALUASI"
HEADER_ROWS = [13, 22, 30, 38]

SECTION_LABELS = {
    "Materi Pelatihan": ["Info mudah didapat", "Pendaftaran mudah", "Petunjuk pendaftaran jelas", "Program jelas"],
    "Program Pelatihan": ["Program menarik", "Program bermanfaat", "Kompetensi meningkat", "Durasi sesuai"],
    "INSTRUKTUR": ["Instruktur menguasai materi", "Kemampuan penyampaian instruktur", "Kemampuan mengelola peserta", "Sikap & teladan instruktur"],
    "Penyelenggaraan": ["Pelayanan petugas", "Jadwal sesuai rencana", "Perlengkapan tepat waktu", "Sarana pelatihan memadai", "Sarana penunjang memadai"],
}

COMMENT_RECAP_HEADERS = [
    "NO.",
    "Dari mana anda mendapatkan informasi tentang pelatihan?",
    "Komentar dan saran Anda terhadap Pelatihan ",
    "Komentar dan saran Anda terhadap Instruktur",
    "Komentar dan saran Anda terhadap penyelenggaraan pelatihan ",
    "Keluhan terhadap penyelenggaraan pelatihan ",
]


def _normalize(text: str) -> str:
    text = str(text).lower().replace("\n", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _find_question_row(ws, question_text: str, cutoff: float = 0.8):
    target = _normalize(question_text)
    best_row, best_ratio = None, 0.0
    for row in ws.iter_rows(min_col=2, max_col=2):
        cell = row[0]
        if not cell.value:
            continue
        norm = _normalize(cell.value)
        if not (norm.startswith("apakah") or norm.startswith("bagaimana")):
            continue
        if norm == target:
            return cell.row
        ratio = difflib.SequenceMatcher(None, norm, target).ratio()
        if norm in target or target in norm:
            ratio = max(ratio, 0.9)
        if ratio > best_ratio:
            best_ratio, best_row = ratio, cell.row
    return best_row if best_ratio >= cutoff else None


def kategori_nilai(score: float) -> str:
    if score >= 5:
        return "SANGAT BAIK"
    if score >= 4:
        return "BAIK"
    if score >= 3:
        return "CUKUP"
    if score >= 2:
        return "KURANG"
    return "TIDAK BAIK"


def generate_comment_recap(
    rows: list,
    info_col: str | None,
    comment_col_map: dict,
    training_comment_label: str = "Komentar/saran program pelatihan",
    instructor_comment_label: str = "Komentar/saran instruktur",
    organizer_comment_label: str = "Komentar/saran penyelenggaraan",
    complaint_label: str = "Keluhan penyelenggaraan",
) -> bytes:
    n = len(rows)

    def get_col(label):
        col = comment_col_map.get(label)
        if col:
            return [("" if r.get(col) is None else str(r.get(col))) for r in rows]
        return [""] * n

    info_values = (
        [("" if r.get(info_col) is None else str(r.get(info_col))) for r in rows]
        if info_col else [""] * n
    )

    columns_data = [
        list(range(1, n + 1)),
        info_values,
        get_col(training_comment_label),
        get_col(instructor_comment_label),
        get_col(organizer_comment_label),
        get_col(complaint_label),
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Komentar"

    ws.append(COMMENT_RECAP_HEADERS)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[1].height = 45

    for i in range(n):
        ws.append([columns_data[c][i] for c in range(6)])

    widths = [5.5, 27.5, 27, 31, 32, 31]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperWidth = "215mm"
    ws.page_setup.paperHeight = "330mm"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = False
    ws.sheet_view.view = "pageBreakPreview"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _format_training_title(training_title: str) -> str:
    name = re.sub(r"\s+", " ", str(training_title).strip()).upper()
    if not name:
        return ""
    name = re.sub(r"^PELATIHAN BERBASIS KOMPETENSI KEJURUAN\s+", "", name)
    name = re.sub(r"^PELATIHAN KOMPETENSI BERBASIS\s+", "", name)
    return f"PELATIHAN BERBASIS KOMPETENSI KEJURUAN {name}"


def _format_program_title(program_name: str) -> str:
    name = re.sub(r"\s+", " ", str(program_name).strip()).upper()
    if not name:
        return ""
    name = re.sub(r"^PROGRAM PELATIHAN DASAR\s+", "", name)
    name = re.sub(r"\s*\([^)]*\)", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    if "PAKET" not in name:
        match = re.match(r"^(.*\S)\s+(\d+)$", name)
        if match:
            name = f"{match.group(1)} PAKET {match.group(2)}"
    return f"PROGRAM PELATIHAN DASAR {name}"


def _add_title_separator_line(ws, row: int = 6, min_col: int = 1, max_col: int = 29):
    thin_double = Side(style="double", color="000000")
    for col in range(min_col, max_col + 1):
        cell = ws.cell(row=row, column=col)
        existing = cell.border
        cell.border = Border(
            top=thin_double, bottom=existing.bottom,
            left=existing.left, right=existing.right,
        )


def _sync_titles_and_separator_lines(wb):
    for sheet_name, (row, max_col) in SHEET_TITLE_ROWS.items():
        if sheet_name not in wb.sheetnames:
            continue
        _add_title_separator_line(wb[sheet_name], row=row, max_col=max_col)
    for sheet_name, cell_coord in TITLE_TEXT_CELLS.items():
        if sheet_name not in wb.sheetnames:
            continue
        wb[sheet_name][cell_coord] = REPORT_TITLE


def _fix_secondary_header_numbers(ws):
    for row in HEADER_ROWS:
        if ws.cell(row=row, column=19).value != "JUMLAH":
            continue
        for i, col in enumerate(SECONDARY_COLS):
            ws[f"{col}{row}"] = DEFAULT_CAPACITY + i + 1


def _stabilize_master_table(ws):
    thin = Side(style="thin", color="000000")
    for row in range(13, 46):
        for col in range(3, 28):
            cell = ws.cell(row=row, column=col)
            border = cell.border
            cell.border = Border(
                left=border.left if border.left.style else thin,
                right=border.right if border.right.style else thin,
                top=border.top if border.top.style else thin,
                bottom=border.bottom if border.bottom.style else thin,
            )
            if row not in HEADER_ROWS:
                cell.alignment = copy(cell.alignment)
                cell.alignment = Alignment(
                    horizontal="center", vertical="center",
                    wrap_text=cell.alignment.wrap_text,
                )


def _configure_print_layout(ws):
    ws.page_setup.orientation = "landscape"
    # F4/Folio: 215 x 330 mm. Set physical dimensions because F4 is not
    # represented consistently by Excel/LibreOffice paper-size enums.
    ws.page_setup.paperSize = None
    ws.page_setup.paperWidth = "215mm"
    ws.page_setup.paperHeight = "330mm"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.2, right=0.2, top=0.35, bottom=0.35, header=0.15, footer=0.15)
    ws.print_options.horizontalCentered = True


def _update_divisor_everywhere(wb, n_respondents: int, use_secondary: bool):
    sum_range = "C{r}:R{r},T{r}:AA{r}" if use_secondary else "C{r}:R{r}"
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_col=19, max_col=19):
            cell = row[0]
            if isinstance(cell.value, str) and cell.value.startswith("=SUM(") and "/12" in cell.value:
                r = cell.row
                cell.value = f"=SUM({sum_range.format(r=r)})/{n_respondents}"


def _style_histogram_chart(wb, training_title: str, section_values=None):
    if "Histogram" not in wb.sheetnames:
        return
    ws = wb["Histogram"]
    if not ws._charts:
        return
    chart = ws._charts[0]
    try:
        runs = chart.title.tx.rich.p[0].r
        label = re.sub(r"\s+", " ", str(training_title).strip()).upper()
        for run in runs:
            if run.t and "NAMA PELATIHAN" in run.t:
                run.t = f" ({label})" if label else " (NAMA PELATIHAN)"
            if run.rPr is not None:
                run.rPr.u = None
                run.rPr.b = True
    except (AttributeError, IndexError):
        pass
    for ser in chart.series:
        ser.graphicalProperties.solidFill = "595959"
        ser.graphicalProperties.line.noFill = True

    if section_values:
        lo = min(section_values)
        hi = max(section_values)
    else:
        lo, hi = 0, 5
    import math
    axis_min = max(0, math.floor(lo * 10) / 10 - 0.2)
    axis_max = min(5, math.ceil(hi * 10) / 10 + 0.2)
    if axis_max - axis_min < 0.5:
        axis_max = min(5, axis_min + 0.5)
    chart.y_axis.scaling.min = round(axis_min, 1)
    chart.y_axis.scaling.max = round(axis_max, 1)
    chart.y_axis.majorUnit = 0.1


def _rewrite_cross_sheet_formulas(wb, row_map: dict, canonical_questions: dict, used_cols: list):
    hide_zero_format = "0;-0;;@"
    for sheet_name in VIEW_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        allowed_labels = SECTION_LABELS.get(sheet_name, list(row_map.keys()))
        for label in allowed_labels:
            if label not in row_map:
                continue
            master_row = row_map[label]
            question_text = canonical_questions.get(label)
            if not question_text:
                continue
            view_row = _find_question_row(ws, question_text)
            if not view_row:
                continue
            for col in used_cols:
                cell = ws[f"{col}{view_row}"]
                cell.value = f"='{MASTER_SHEET}'!{col}{master_row}"
                cell.number_format = hide_zero_format


def generate_official_report(
    template_path: str,
    rows: list,
    score_col_map: dict,
    canonical_questions: dict,
    program_name: str,
    training_title: str = "",
    date_text: str = "",
    instructor_name: str = "",
) -> tuple:
    n = len(rows)
    if n == 0:
        raise ValueError("Tidak ada responden pada data terfilter.")

    capped = min(n, MAX_RESPONDENTS)
    use_secondary = capped > DEFAULT_CAPACITY
    all_cols = PRIMARY_COLS + (SECONDARY_COLS if use_secondary else [])

    wb = openpyxl.load_workbook(template_path)
    ws_master = wb[MASTER_SHEET]

    # EMF logos in the original workbook are dropped by openpyxl on save.
    # Replace the master-sheet header image with a PNG that survives export.
    logo_path = Path(template_path).with_name("logo_kemnaker.png")
    if logo_path.exists():
        ws_master._images = []
        logo = ExcelImage(str(logo_path))
        logo.width, logo.height = 105, 92
        logo.anchor = "B1"
        ws_master.add_image(logo)

    _fix_secondary_header_numbers(ws_master)
    _stabilize_master_table(ws_master)
    _sync_titles_and_separator_lines(wb)
    for sheet in wb.worksheets:
        _configure_print_layout(sheet)

    ws_master["A6"] = REPORT_TITLE
    if training_title:
        ws_master["A7"] = _format_training_title(training_title)
    if program_name:
        ws_master["A8"] = _format_program_title(program_name)
    if date_text:
        ws_master["A9"] = date_text
    if instructor_name:
        ws_master["C31"] = f"Nama Instruktur : {instructor_name}"
        if "INSTRUKTUR" in wb.sheetnames:
            wb["INSTRUKTUR"]["C12"] = f"INSTRUKTUR : {instructor_name}"

    row_map = {}
    for label, question in canonical_questions.items():
        if label not in score_col_map:
            continue
        row = _find_question_row(ws_master, question)
        if row:
            row_map[label] = row

    values_by_label = {}
    for label, row in row_map.items():
        col_name = score_col_map[label]
        raw_values = [r.get(col_name) for r in rows if r.get(col_name) is not None]
        values = [int(round(float(v))) for v in raw_values][:capped]
        values_by_label[label] = values

        for col in all_cols:
            ws_master[f"{col}{row}"] = None
        for i, val in enumerate(values):
            cell = ws_master[f"{all_cols[i]}{row}"]
            cell.value = val
            cell.alignment = cell.alignment.copy(horizontal="center", vertical="center")

    _rewrite_cross_sheet_formulas(wb, row_map, canonical_questions, all_cols)

    section_avgs = []
    for labels in SECTION_LABELS.values():
        means = [sum(values_by_label[l]) / len(values_by_label[l]) for l in labels if l in values_by_label and values_by_label[l]]
        if means:
            section_avgs.append(sum(means) / len(means))
    _style_histogram_chart(wb, training_title or program_name, section_avgs)

    actual_n = max((len(v) for v in values_by_label.values()), default=capped)
    _update_divisor_everywhere(wb, actual_n, use_secondary)

    wb.calculation.fullCalcOnLoad = True

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue(), {
        "respondents_used": actual_n,
        "respondents_total_filtered": n,
        "capped": n > MAX_RESPONDENTS,
        "used_secondary_columns": use_secondary,
        "questions_matched": len(row_map),
        "questions_total": len(canonical_questions),
    }
