"""Business rules + Excel import/export for the magang interview dashboard.

The original Excel workbook is kept as the sole export template so users retain
all formatting, formulas, sheet titles, rubrics and Excel validations.
"""

from __future__ import annotations

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Any
import re

from openpyxl import load_workbook

TEMPLATE_PATH = Path(__file__).with_name("Template_Interview_Peserta_Magang.xlsx")
MAX_KANDIDAT = 150

# Exactly the same order as row 6 / columns A:Y in the supplied workbook.
KANDIDAT_FIELDS = [
    ("id", "ID Kandidat"),
    ("nama", "Nama Lengkap"),
    ("telepon", "No. HP / WA"),
    ("email", "Email"),
    ("alamat", "Alamat Domisili"),
    ("kota", "Kota / Kabupaten"),
    ("kampus", "Asal Perguruan Tinggi"),
    ("jenjang", "Jenjang"),
    ("prodi", "Program Studi / Jurusan"),
    ("semester", "Semester"),
    ("ipk", "IPK (0–4)"),
    ("status_kuliah", "Status Perkuliahan"),
    ("pengalaman_kerja", "Pengalaman Magang / Kerja"),
    ("pengalaman_organisasi", "Pengalaman Organisasi"),
    ("portofolio", "Proyek / Portofolio"),
    ("keahlian", "Keahlian Utama"),
    ("tools", "Tools / Software"),
    ("sertifikasi", "Sertifikasi / Pelatihan"),
    ("link", "Tautan LinkedIn / Portofolio"),
    ("bidang", "Posisi / Bidang Magang"),
    ("wfo", "Bersedia WFO"),
    ("durasi", "Durasi (bulan)"),
    ("mulai", "Dapat Mulai Tanggal"),
    ("motivasi", "Motivasi / Harapan Magang"),
    ("status_cv", "Status CV / Berkas"),
]

ASPEK = [
    ("komunikasi", "Komunikasi", 0.15),
    ("motivasi", "Motivasi", 0.15),
    ("teknis", "Kompetensi Teknis", 0.25),
    ("masalah", "Pemecahan Masalah", 0.15),
    ("kerja_sama", "Kerja Sama", 0.10),
    ("belajar", "Kemauan Belajar", 0.10),
    ("profesional", "Profesionalisme", 0.10),
]

INTERVIEW_FIELDS = [
    "id", "nama", "bidang", "tanggal", "batch", *[s[0] for s in ASPEK],
    "skor", "rekomendasi", "keputusan", "kelebihan", "peningkatan", "catatan", "pewawancara",
]

DECISIONS = ["Menunggu", "Lolos", "Cadangan", "Tidak Lolos"]
BATCH_PATTERN = re.compile(r"^(\d{4})-B(0[1-9]|[1-4][0-9]|50)$")


def make_batch_code(year: int, number: int) -> str:
    """Stable sortable batch ID; example: 2026-B01."""
    if not 2020 <= int(year) <= 2100 or not 1 <= int(number) <= 50:
        raise ValueError("Tahun batch harus 2020–2100 dan nomor batch 1–50.")
    return f"{int(year):04d}-B{int(number):02d}"


def batch_label(code: str | None) -> str:
    """Human-friendly representation for forms, ranking and Excel export."""
    if not code:
        return "Belum ditentukan"
    match = BATCH_PATTERN.fullmatch(str(code))
    return f"Batch {int(match.group(2))} — {match.group(1)}" if match else str(code)


def parse_batch_code(code: str | None, fallback_year: int) -> tuple[int, int]:
    match = BATCH_PATTERN.fullmatch(str(code or ""))
    return (int(match.group(1)), int(match.group(2))) if match else (fallback_year, 1)


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    return str(value).strip()


def _date_to_excel(value: Any):
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return value
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except (TypeError, ValueError):
        return None


def _number_or_none(value: Any, integer: bool = False):
    if value is None or value == "":
        return None
    try:
        return int(float(value)) if integer else float(value)
    except (ValueError, TypeError):
        return None


def _safe_cell(value: Any):
    # Prevent untrusted text from being interpreted as an Excel formula.
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _workbook(source: bytes | str | Path | BytesIO):
    if isinstance(source, bytes):
        source = BytesIO(source)
    return load_workbook(source, data_only=False)


def get_rubric_and_questions():
    wb = _workbook(TEMPLATE_PATH)
    sh = wb["Panduan & Rubrik"]
    weights = {}
    for i, (key, _, fallback) in enumerate(ASPEK):
        raw = _number_or_none(sh.cell(8 + i, 3).value)
        weights[key] = raw if raw is not None else fallback
    high = _number_or_none(sh["C19"].value)
    mid = _number_or_none(sh["C20"].value)
    questions = []
    for r in range(26, 38):
        questions.append({
            "no": sh.cell(r, 2).value,
            "area": _as_text(sh.cell(r, 3).value),
            "pertanyaan": _as_text(sh.cell(r, 4).value),
            "pendalaman": _as_text(sh.cell(r, 5).value),
            "dinilai": _as_text(sh.cell(r, 6).value),
        })
    scale = []
    for r in range(43, 48):
        scale.append({
            "nilai": sh.cell(r, 2).value,
            "kategori": _as_text(sh.cell(r, 3).value),
            "panduan": _as_text(sh.cell(r, 4).value),
        })
    return weights, (high if high is not None else 80), (mid if mid is not None else 65), questions, scale


def candidate_id_next(candidates: dict[str, dict]) -> str:
    used = [int(m.group(1)) for k in candidates if (m := re.fullmatch(r"MG-(\d+)", k))]
    nxt = max(used, default=0) + 1
    return f"MG-{nxt:03d}"


def read_xlsx(source: bytes | str | Path) -> tuple[dict, dict]:
    wb = _workbook(source)
    necessary = {"Data Kandidat", "Penilaian Interview"}
    if not necessary.issubset(set(wb.sheetnames)):
        raise ValueError("File harus berasal dari template interview dengan sheet Data Kandidat dan Penilaian Interview.")
    sh = wb["Data Kandidat"]
    candidates: dict[str, dict] = {}
    for row in sh.iter_rows(min_row=7, max_row=max(156, sh.max_row), min_col=1, max_col=25, values_only=True):
        if not _as_text(row[1]):
            continue
        rec = {key: _as_text(value) for (key, _), value in zip(KANDIDAT_FIELDS, row)}
        cid = rec["id"]
        if not cid:
            continue
        if cid in candidates:
            raise ValueError(f"ID kandidat duplikat: {cid}")
        candidates[cid] = rec
    sh = wb["Penilaian Interview"]
    interviews: dict[str, dict] = {}
    for row in sh.iter_rows(min_row=7, max_row=max(156, sh.max_row), min_col=1, max_col=19, values_only=True):
        cid = _as_text(row[0])
        if cid not in candidates:
            continue
        rec = {
            "id": cid,
            "tanggal": _as_text(row[3]),
            "batch": _as_text(row[18]),  # Optional 19th column; old Excel files remain supported.
            "keputusan": _as_text(row[13]) or "Menunggu",
            "kelebihan": _as_text(row[14]),
            "peningkatan": _as_text(row[15]),
            "catatan": _as_text(row[16]),
            "pewawancara": _as_text(row[17]),
        }
        if rec["batch"] and not BATCH_PATTERN.fullmatch(rec["batch"]):
            raise ValueError(f"Kode batch di Excel tidak valid untuk {cid}: {rec['batch']}")
        for i, (key, _, _) in enumerate(ASPEK):
            value = _number_or_none(row[4 + i], integer=True)
            rec[key] = value if value in range(1, 6) else 0
        interviews[cid] = rec
    return candidates, interviews


def calculate(interview: dict | None, weights: dict[str, float], high: float, mid: float):
    if not interview:
        return None, "Belum Dinilai"
    ratings = [interview.get(key) for key, _, _ in ASPEK]
    if not all(isinstance(x, (int, float)) and 1 <= x <= 5 for x in ratings):
        return None, "Belum Lengkap"
    total_weight = sum(weights.get(key, default) for key, _, default in ASPEK)
    if total_weight <= 0:
        return None, "Bobot Tidak Valid"
    score = round(20 * sum(interview[key] * weights.get(key, default) for key, _, default in ASPEK) / total_weight, 1)
    if score >= high:
        return score, "Direkomendasikan"
    if score >= mid:
        return score, "Dipertimbangkan"
    return score, "Belum Direkomendasikan"


def ranking_rows(candidates: dict, interviews: dict, weights: dict, high: float, mid: float, batch_filter: str | None = None):
    results = []
    for cid, person in candidates.items():
        val = interviews.get(cid)
        # None = all batches; empty string = historical records with no batch.
        if batch_filter is not None and (val or {}).get("batch", "") != batch_filter:
            continue
        score, status = calculate(val, weights, high, mid)
        if score is None:
            continue
        results.append({
            "ID": cid,
            "Nama Kandidat": person.get("nama", ""),
            "Perguruan Tinggi": person.get("kampus", ""),
            "Program Studi": person.get("prodi", ""),
            "Posisi Magang": person.get("bidang", ""),
            "IPK": person.get("ipk", ""),
            "Skor": score,
            "Batch": batch_label((val or {}).get("batch")),
            "Rekomendasi": status,
            "Keputusan Akhir": (val or {}).get("keputusan") or "Menunggu",
            "Tanggal Interview": (val or {}).get("tanggal", ""),
        })
    results.sort(key=lambda x: (-x["Skor"], x["Nama Kandidat"].lower()))
    previous_score, previous_rank = None, None
    for i, result in enumerate(results, 1):
        if result["Skor"] != previous_score:
            previous_rank = i
        result["Ranking"] = previous_rank
        previous_score = result["Skor"]
    return results


def export_xlsx(candidates: dict, interviews: dict) -> bytes:
    if len(candidates) > MAX_KANDIDAT or len(interviews) > MAX_KANDIDAT:
        raise ValueError(f"Template menampung maksimal {MAX_KANDIDAT} kandidat.")
    wb = _workbook(TEMPLATE_PATH)
    sh = wb["Data Kandidat"]
    # Retain ID placeholders + all original styling and data validation.
    for r in range(7, 157):
        for c in range(2, 26):
            sh.cell(r, c).value = None
    for i, (cid, data) in enumerate(sorted(candidates.items()), 7):
        for j, (field, _) in enumerate(KANDIDAT_FIELDS, 1):
            value = cid if field == "id" else data.get(field)
            if field in ("semester", "durasi"):
                value = _number_or_none(value, integer=True)
            elif field == "ipk":
                value = _number_or_none(value)
            elif field == "mulai":
                value = _date_to_excel(value)
            else:
                value = _safe_cell(value) if value else None
            sh.cell(i, j).value = value
    sh = wb["Penilaian Interview"]
    # Append a 19th column; original 18 columns and all their Excel formulas remain intact.
    # This also makes Excel backups round-trip the new batch field.
    from copy import copy
    sh.cell(6, 19).value = "Batch Interview"
    sh.cell(6, 19)._style = copy(sh.cell(6, 18)._style)
    sh.column_dimensions["S"].width = 23
    for r in range(7, 157):
        sh.cell(r, 19)._style = copy(sh.cell(r, 18)._style)
    writable_columns = [1, 4, 5, 6, 7, 8, 9, 10, 11, 14, 15, 16, 17, 18, 19]
    for r in range(7, 157):
        for c in writable_columns:
            sh.cell(r, c).value = None
    for row_num, (cid, interview) in enumerate(sorted(interviews.items()), 7):
        if cid not in candidates:
            continue
        sh.cell(row_num, 1).value = cid
        sh.cell(row_num, 4).value = _date_to_excel(interview.get("tanggal"))
        sh.cell(row_num, 19).value = _safe_cell(interview.get("batch") or "")
        for i, (key, _, _) in enumerate(ASPEK, 5):
            score = interview.get(key)
            sh.cell(row_num, i).value = score if score in (1, 2, 3, 4, 5) else None
        for col, key in [(14, "keputusan"), (15, "kelebihan"), (16, "peningkatan"), (17, "catatan"), (18, "pewawancara")]:
            value = interview.get(key)
            sh.cell(row_num, col).value = _safe_cell(value) if value else None
    # Excel calculates the template's existing lookup, scores, recommendations and dashboard on open.
    try:
        from openpyxl.workbook.properties import CalcProperties
        wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    except ImportError:
        pass
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
