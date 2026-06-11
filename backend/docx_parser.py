"""
docx_parser.py — Deka Research Questionnaire Parser v9
=======================================================
Changelog v9 (dari v8):

Fix 1 — _extract_kota_from_doc  (regex terbalik)
  Format data Deka: nama kota di kiri, kode di kanan (mis. 'Jakarta  1  CEK KUOTA').
  v8 hanya punya regex yg mengasumsikan kode di kiri ('1. Jakarta').
  v9: 3 strategi — (A) nested table col0=nama col1=kode, (B) baris "kode. nama",
  (C) satu baris panjang "Jakarta 1 Bogor 2 ...". Plus Pass 1 khusus mencari
  pertanyaan S6/S0a agar tidak terpengaruh tabel quota atau tabel lain.

Fix 2 — _is_screening_table  (Tukang v3 screening tidak terdeteksi)
  Dokumen Tukang menggunakan tabel dengan row-0 kosong/'SCREENING' dan row-1
  berisi 'No/Pertanyaan/Route'. v8 hanya lolos jika row-0 header langsung 'NO'.
  v9: 3 kondisi — strict (row0 NO/PERTANYAAN/ROUTE), lenient (row1 NO/PERTANYAAN/ROUTE),
  dan fallback (cek apakah 3 row pertama punya q_id valid di kolom 0).

Fix 3 — _parse_awareness_block  (label sub-pertanyaan hardcode)
  v8 selalu pakai label generik. v9 mengekstrak baris berformat
  "(SA)"/"(MA)"/"(S)"/"(M)" langsung dari teks sel, mem-pair baris
  berurutan sebagai sub a/b/c/d/e/f dengan teks asli dokumen.
  Fallback ke label generik hanya jika pola tidak ditemukan.

Fix 4 — _extract_inline_subs  (helper baru)
  Pertanyaan V/Z/Q4/Q6 dan lainnya yang memuat beberapa sub berurutan
  ("V1. ...(M)  V2. ...(M)  V3. ...(M)") sekarang di-split menjadi
  pertanyaan terpisah dengan ID eksplisit via helper _extract_inline_subs(),
  dipanggil di _parse_question_row sebelum fallback ke standard question.

Fix 5 — fallback tracking + parse notes
  Setiap pertanyaan yang masuk ke fallback (choices kosong, route tidak diketahui,
  atau tipe tidak terdeteksi) diberi flag "_parse_warning". Setelah konversi,
  app.py mengumpulkan peringatan ini dan mengembalikannya ke frontend sebagai
  "parse_notes" agar user tahu item apa yang perlu dicek manual.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import docx
from docx.table import Table, _Cell

# ── Pattern & konstanta ────────────────────────────────────────────────────────

_Q_ID_RE = re.compile(
    r'^('
    r'SEC\d+'
    r'|S\d+[a-zA-Z]?'
    r'|MD\d*[a-zA-Z0-9\-\.]*'
    r'|[A-Z]{1,2}\d+[a-zA-Z0-9\-\.]*'
    r'|[A-Z]{1,2}\d*\.[a-zA-Z0-9]?'
    r'|[VYZF]\d*[a-zA-Z0-9\-\.]*'
    r'|[A-Z]\d*'
    r')$',
    re.IGNORECASE
)

_HINT_MARKERS = [
    "KARTU BANTU", "TUNJUKKAN", "SPONTAN", "PROBE", "BACAKAN",
    "ROTASIKAN", "INTERVIEWER", "DP:", "NOTE:", "NOTE TO DP",
    "TANYAKAN KEPADA", "TANYAKAN JIKA", "DITANYAKAN JIKA",
    "CATAT USIA", "CATAT AKTUAL", "TRANSFER JAWABAN",
    "JAWABAN SPONTAN",
]

_STOP_WORDS     = {"AKHIRI WAWANCARA", "HENTIKAN WAWANCARA", "STOP", "STOP WAWANCARA"}
_CONTINUE_WORDS = {"LANJUTKAN", "CONTINUE", "CEK KUOTA"}

_NON_SECTION_RE = re.compile(
    r'^(TABEL ISIAN|PANEL KONTROL|KARTU BANTU P|NOTE:|CATATAN DP|DP:|'
    r'\[.*\]|AKHIR DARI|MASUKAN LIST)',
    re.IGNORECASE
)

# ── Brand & product choices ────────────────────────────────────────────────────

_GENERIC_BRAND_FALLBACK = [
    ("1", "Merek A"), ("2", "Merek B"), ("3", "Merek C"), ("4", "Merek D"),
    ("5", "Merek E"), ("98", "Lainnya"), ("99", "Tidak tahu"),
]

_BRAND_SLOTS: dict[str, list[tuple]] = {
    "primary":   list(_GENERIC_BRAND_FALLBACK),
    "secondary": list(_GENERIC_BRAND_FALLBACK),
    "tertiary":  list(_GENERIC_BRAND_FALLBACK),
}
_PRODUCT_LABELS: dict[str, str] = {
    "primary":   "Produk Utama",
    "secondary": "Produk Kedua",
    "tertiary":  "Produk Ketiga",
}
_CHOICES_PRODUK: list[tuple] = [
    ("1", "Produk Utama"), ("2", "Produk Kedua"), ("3", "Produk Ketiga"),
]
_CHOICES_JENIS_PRODUK: list[tuple] = [
    ("1", "Tipe A"), ("2", "Tipe B"), ("3", "Tipe C"),
    ("98", "Lainnya"), ("99", "Tidak Tahu"),
]


def _get_merek(slot: str = "primary") -> list[tuple]:
    return _BRAND_SLOTS.get(slot, _GENERIC_BRAND_FALLBACK)


_SKALA_10 = [(str(i), str(i)) for i in range(1, 11)]
_SKALA_11 = [(str(i), str(i)) for i in range(0, 11)]

_USIA_RANGE = [
    ("1","Di bawah 25 tahun"), ("2","25–29 tahun"), ("3","30–34 tahun"),
    ("4","35–39 tahun"), ("5","40–44 tahun"), ("6","45–50 tahun"),
    ("7","Di atas 50 tahun"),
]

_TIME_RANGE = [
    ("1","Di bawah 6 bulan yang lalu"), ("2","Di bawah 1 tahun yang lalu"),
    ("3","Di bawah 2 tahun yang lalu"), ("4","Di bawah 3 tahun yang lalu"),
    ("5","Lebih dari 3 tahun yang lalu"),
]

# ── Fallback warning tracking ──────────────────────────────────────────────────
# Diisi oleh parser selama proses; dikosongkan kembali di parse_questionnaire().
_PARSE_WARNINGS: list[dict] = []

def _warn(q_id: str, issue: str) -> None:
    # Deduplicate: skip if same id+issue already recorded
    if not any(w["id"] == q_id and w["issue"] == issue for w in _PARSE_WARNINGS):
        _PARSE_WARNINGS.append({"id": q_id, "issue": issue})


# ── Helpers ────────────────────────────────────────────────────────────────────

def _cell_text(cell: _Cell) -> str:
    return "\n".join(p.text for p in cell.paragraphs).strip()


def _detect_routing(text: str) -> str:
    t = text.strip().upper()
    for w in _STOP_WORDS:
        if w in t:
            return "STOP"
    for w in _CONTINUE_WORDS:
        if w in t:
            return "CONTINUE"
    return t if t else ""


def _normalize_route(raw: str) -> str:
    raw = raw.strip().upper().replace("\n", " ").replace("  ", " ")
    mapping = {
        "SA": "SA", "MA": "MA", "OE": "OE",
        "M": "MA", "S": "SA",
        "M (OE)": "MA", "MA OE": "MA",
        "SA MA": "SA+MA", "SA SA": "SA",
    }
    return mapping.get(raw, raw)


def _split_hint(text: str) -> tuple[str, str]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    q_lines, h_lines = [], []
    for ln in lines:
        upper = ln.upper()
        if any(upper.startswith(m) for m in _HINT_MARKERS):
            h_lines.append(ln)
        elif re.match(r'^\[.*\]$', ln):
            h_lines.append(ln)
        else:
            q_lines.append(ln)
    return " ".join(q_lines).strip(), " ".join(h_lines).strip()


def _make_q(q_id, section, subsection, question, hint="",
            route_type="SA", choices=None, skip_logic="",
            is_grid=False, is_header=False, raw_text="",
            parse_warning="") -> dict:
    q = {
        "id":          q_id,
        "section":     section,
        "subsection":  subsection,
        "question":    question,
        "hint":        hint,
        "route_type":  route_type,
        "choices":     choices or [],
        "skip_logic":  skip_logic,
        "is_grid":     is_grid,
        "is_header":   is_header,
        "raw_text":    raw_text,
    }
    if parse_warning:
        q["_parse_warning"] = parse_warning
        _warn(q_id, parse_warning)
    return q


def _choices_from_list(pairs: list[tuple[str, str]]) -> list[dict]:
    return [{"code": code, "label": label, "routing": ""} for code, label in pairs]


def _parse_simple_nested_table(tbl: Table) -> list[dict]:
    choices = []
    nc = len(tbl.columns)
    for row in tbl.rows:
        cells = [c.text.strip() for c in row.cells]
        if nc == 1:
            if cells[0] and not re.match(r'^\[|^ROTASIKAN|^DP:', cells[0]):
                choices.append({"code": "", "label": cells[0], "routing": ""})
        elif nc >= 2:
            c0, c1 = cells[0], cells[1] if len(cells) > 1 else ""
            routing = cells[2] if len(cells) > 2 else ""
            if re.match(r'^ROTASIKAN|^\[DP|^KARTU', c0, re.I):
                continue
            if re.match(r'^\d+$', c1.strip()):
                label, code = c0, c1
            elif re.match(r'^\d+$', c0.strip()):
                code, label = c0, c1
            else:
                label, code = c0, c1
            if label or code:
                choices.append({
                    "code":    code.strip(),
                    "label":   label.strip(),
                    "routing": _detect_routing(routing),
                })
    seen, deduped = set(), []
    for ch in choices:
        key = (ch["code"], ch["label"])
        if key not in seen and (ch["code"] or ch["label"]):
            seen.add(key)
            deduped.append(ch)
    return deduped


def _parse_brand_grid_table(nt: Table) -> dict[str, list[dict]]:
    if len(nt.rows) < 3:
        return {}
    rows = list(nt.rows)
    header_ids = [c.text.strip() for c in rows[0].cells]
    col_ids = [h.lower().replace("-","_").replace(" ","") for h in header_ids]
    result: dict[str, list[dict]] = {cid: [] for cid in col_ids if cid}
    for data_row in rows[2:]:
        cells = [c.text.strip() for c in data_row.cells]
        brand_label = cells[0]
        if not brand_label or re.match(r'^\[DP', brand_label, re.I):
            continue
        if brand_label.upper().startswith("TT"):
            brand_label = "TT/TA"
        for col_idx, cid in enumerate(col_ids):
            if not cid or col_idx >= len(cells):
                continue
            code = cells[col_idx].strip()
            if code and re.match(r'^\d+$', code):
                result.setdefault(cid, []).append({
                    "code": code, "label": brand_label, "routing": "",
                })
    return {k: v for k, v in result.items() if v}


def _extract_brand_lists_from_doc(doc: docx.Document) -> None:
    global _BRAND_SLOTS, _PRODUCT_LABELS, _CHOICES_PRODUK, _CHOICES_JENIS_PRODUK

    def _extract_col0_pairs(nt: Table, skip_rows: int = 2) -> list[tuple]:
        pairs = []
        for row in list(nt.rows)[skip_rows:]:
            cells = [c.text.strip() for c in row.cells]
            label = cells[0]
            code  = cells[1] if len(cells) > 1 else ""
            if not label:
                continue
            if re.match(r'^(ROTASIKAN|DP:|TT/)', label, re.I) and not code:
                continue
            pairs.append((code, label))
        return [(c, l) for c, l in pairs if c or l]

    def _infer_product_label(before_text: str) -> str:
        m = re.search(r'(?:merek|brand|produk)\s+([A-Za-z0-9/ ]{3,40})', before_text, re.I)
        if m:
            return m.group(1).strip().title()
        caps = re.findall(r'\b[A-Z][A-Za-z0-9]{2,}\b', before_text)
        if caps:
            return caps[0]
        return ""

    slots_assigned = {"primary": False, "secondary": False, "tertiary": False}

    for tbl in doc.tables:
        rows = list(tbl.rows)
        for i, row in enumerate(rows):
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 2:
                continue
            q_id   = cells[0].upper()
            q_cell = row.cells[1]

            if not cells[0]:
                for nt in q_cell.tables:
                    if len(nt.columns) < 6 or len(nt.rows) < 3:
                        continue
                    col_headers = [c.text.strip().upper() for c in nt.rows[0].cells]
                    pairs = _extract_col0_pairs(nt, skip_rows=2)
                    if not pairs:
                        continue
                    prev_text = rows[i-1].cells[1].text if i > 0 else ""

                    if "Q1" in col_headers and not slots_assigned["primary"]:
                        _BRAND_SLOTS["primary"] = pairs
                        lbl = _infer_product_label(prev_text)
                        if lbl: _PRODUCT_LABELS["primary"] = lbl
                        slots_assigned["primary"] = True
                    elif "V1" in col_headers and not slots_assigned["secondary"]:
                        _BRAND_SLOTS["secondary"] = pairs
                        lbl = _infer_product_label(prev_text)
                        if lbl: _PRODUCT_LABELS["secondary"] = lbl
                        slots_assigned["secondary"] = True
                    elif "Z1" in col_headers and not slots_assigned["tertiary"]:
                        _BRAND_SLOTS["tertiary"] = pairs
                        lbl = _infer_product_label(prev_text)
                        if lbl: _PRODUCT_LABELS["tertiary"] = lbl
                        slots_assigned["tertiary"] = True
                    elif not slots_assigned["primary"]:
                        _BRAND_SLOTS["primary"] = pairs
                        lbl = _infer_product_label(prev_text)
                        if lbl: _PRODUCT_LABELS["primary"] = lbl
                        slots_assigned["primary"] = True

                    if len(nt.columns) == 3:
                        h = [c.text.strip().upper() for c in nt.rows[0].cells]
                        if "S8D" in h or "S8E" in h:
                            pairs_sde = []
                            for r in list(nt.rows)[1:]:
                                c0 = r.cells[0].text.strip()
                                c1 = r.cells[1].text.strip() if len(r.cells) > 1 else ""
                                if c0 and c1:
                                    pairs_sde.append((c1, c0))
                            if pairs_sde:
                                _CHOICES_JENIS_PRODUK[:] = pairs_sde

            if q_id in ("S8C", "S8B", "S8A"):
                for nt in q_cell.tables:
                    if len(nt.columns) < 3 or len(nt.rows) < 3:
                        continue
                    pairs_s8 = []
                    for r in list(nt.rows)[1:]:
                        label = r.cells[0].text.strip()
                        code  = r.cells[1].text.strip() if len(r.cells) > 1 else ""
                        if label and code and re.match(r'^\d+$', code):
                            pairs_s8.append((code, label))
                    if pairs_s8:
                        _CHOICES_PRODUK[:] = pairs_s8

    if slots_assigned["primary"] and not slots_assigned["secondary"]:
        _BRAND_SLOTS["secondary"] = _BRAND_SLOTS["primary"]
        _BRAND_SLOTS["tertiary"]  = _BRAND_SLOTS["primary"]


# ── Fix 1: _extract_kota_from_doc — 3 strategi + Pass 1 S6/S0a ───────────────

def _extract_kota_from_doc(doc: docx.Document) -> list[tuple[str, str]]:
    """
    Ekstrak daftar kota dari pertanyaan screening kota (S6, S0a, S0b).

    Strategi (per sel / nested table):
    A) Nested table 3-col: col0=nama kota, col1=kode angka  (format Deka baru)
    B) Baris teks "kode. nama" — kode di kiri                (format lama)
    C) Baris tunggal panjang "NamaKota kode NamaKota kode …" (fallback)

    Pass 1: cari baris S6/S0a/S0b di tabel screening terlebih dahulu.
    Pass 2: scan semua tabel jika Pass 1 kosong.
    """
    KOTA_PATTERNS = [
        r'jabodetabek', r'jakarta', r'bogor', r'depok', r'tangerang', r'bekasi',
        r'bandung', r'cirebon', r'semarang', r'tegal', r'solo', r'surabaya',
        r'malang', r'yogyakarta', r'madiun', r'medan', r'palembang',
        r'makassar', r'denpasar', r'balikpapan', r'samarinda', r'pontianak',
        r'manado', r'pekanbaru', r'batam',
    ]
    kota_re = re.compile('|'.join(KOTA_PATTERNS), re.IGNORECASE)

    def _parse_nested_table_kota(nt: Table) -> list[tuple[str, str]]:
        """Format A: nested table dengan kolom nama | kode | (routing)."""
        result = []
        seen = set()
        for row in nt.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 2:
                continue
            c0, c1 = cells[0], cells[1]
            # Tentukan mana nama dan mana kode
            if re.match(r'^\d+$', c1) and kota_re.search(c0):
                # Format A: nama kiri, kode kanan (Deka baru — Jakarta | 1 | CEK KUOTA)
                name, code = c0, c1
            elif re.match(r'^\d+$', c0) and kota_re.search(c1):
                # Format B: kode kiri, nama kanan
                code, name = c0, c1
            else:
                continue
            key = name.upper()
            if key not in seen:
                seen.add(key)
                result.append((code, name))
        return result

    def _parse_text_kota(txt: str) -> list[tuple[str, str]]:
        """Format B+C: parsing dari teks baris."""
        result = []
        seen = set()
        lines = [l.strip() for l in re.split(r'[\n\t]', txt) if l.strip()]
        for ln in lines:
            # Format B: "1. Jakarta" atau "1 Jakarta"
            m = re.match(r'^(\d{1,2})[.\s]\s*(.+)$', ln)
            if m:
                code, name = m.group(1).strip(), m.group(2).strip()
                if kota_re.search(name) and name.upper() not in seen:
                    seen.add(name.upper())
                    result.append((code, name))
                continue
            # Format C baris panjang: "Jakarta 1 Bogor 2 ..."
            tokens = ln.split()
            i = 0
            while i < len(tokens) - 1:
                if kota_re.search(tokens[i]) and re.match(r'^\d+$', tokens[i+1]):
                    name, code = tokens[i], tokens[i+1]
                    key = name.upper()
                    if key not in seen:
                        seen.add(key)
                        result.append((code, name))
                    i += 2
                else:
                    i += 1
        return result

    def _collect_from_cell(cell: _Cell) -> list[tuple[str, str]]:
        found = []
        # Strategy A: nested table di dalam sel
        for nt in cell.tables:
            res = _parse_nested_table_kota(nt)
            if res:
                found.extend(res)
        # Strategy B/C: teks langsung
        if not found:
            found = _parse_text_kota(cell.text)
        return found

    # Pass 1: cari baris S6 / S0a / S0b di semua tabel
    KOTA_Q_IDS = {"S6", "S0A", "S0B", "S0"}
    for tbl in doc.tables:
        for row in tbl.rows:
            cells_t = [c.text.strip() for c in row.cells]
            if cells_t[0].upper() in KOTA_Q_IDS and len(row.cells) >= 2:
                res = _collect_from_cell(row.cells[1])
                if res:
                    return res

    # Pass 2: scan semua tabel
    found_all: list[tuple[str, str]] = []
    seen_all: set[str] = set()
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for code, name in _collect_from_cell(cell):
                    if name.upper() not in seen_all:
                        seen_all.add(name.upper())
                        found_all.append((code, name))
    if found_all:
        return found_all

    # Fallback default
    return [
        ('1', 'Jabodetabek'), ('2', 'Bandung'), ('3', 'Cirebon'),
        ('4', 'Semarang'), ('5', 'Tegal'), ('6', 'Solo'),
        ('7', 'Surabaya'), ('8', 'Malang'), ('9', 'Yogyakarta'),
        ('10', 'Madiun'), ('11', 'Lainnya'),
    ]


# ── Fix 2: _is_screening_table — 3 kondisi ────────────────────────────────────

def _is_screening_table(tbl: Table) -> bool:
    """
    Deteksi tabel screening 3-kolom.

    Kondisi 1 (strict): row0 = NO | PERTANYAAN | ROUTE
    Kondisi 2 (lenient): row0 kosong/teks, row1 = No | Pertanyaan | Route
    Kondisi 3 (fallback): kolom 0 dari 3 baris pertama (non-header) punya q_id valid
    """
    if len(tbl.columns) != 3 or len(tbl.rows) < 2:
        return False

    rows = list(tbl.rows)
    h0 = [c.text.strip().upper() for c in rows[0].cells]

    # Kondisi 1: header row 0 langsung berisi NO / PERTANYAAN / ROUTE
    if h0[0] in ("NO", "") and "PERTANYAAN" in h0[1] and "ROUTE" in h0[2]:
        return True

    # Kondisi 2: row 0 adalah judul ("SCREENING"), row 1 adalah header
    if len(rows) >= 2:
        h1 = [c.text.strip().upper() for c in rows[1].cells]
        if h1[0] in ("NO", "") and "PERTANYAAN" in h1[1] and "ROUTE" in h1[2]:
            return True

    # Kondisi 3: fallback — cek apakah ada q_id valid di 3-5 baris pertama (non-empty)
    valid_ids = 0
    checked = 0
    for row in rows[:8]:
        c0 = row.cells[0].text.strip()
        if not c0:
            continue
        checked += 1
        if _Q_ID_RE.match(c0):
            valid_ids += 1
        if checked >= 5:
            break
    if valid_ids >= 2:
        return True

    return False


def _is_main_table(tbl: Table) -> bool:
    if len(tbl.columns) != 3 or len(tbl.rows) < 3:
        return False
    for row in tbl.rows[:5]:
        cells = [c.text.strip() for c in row.cells]
        if "KUESIONER UTAMA" in cells[1].upper():
            return True
    return False


def _is_info_table(tbl: Table) -> bool:
    if len(tbl.rows) < 2:
        return False
    all_text = ""
    for row in list(tbl.rows)[:8]:
        for cell in row.cells:
            all_text += " " + cell.text.strip().upper()
    info_markers = ["NAMA RESPONDEN", "INTERVIEWER", "TANGGAL WAWANCARA",
                    "NOMOR KUESIONER", "WAKTU MULAI", "NOMOR RESPONDEN"]
    return sum(1 for m in info_markers if m in all_text) >= 2


def _extract_respondent_info(doc: docx.Document) -> list[dict]:
    info_tbl = None
    for tbl in doc.tables:
        if _is_info_table(tbl):
            info_tbl = tbl
            break

    standard_fields = [
        ("nama_responden",    "Nama Responden",             "yes"),
        ("nomor_kuesioner",   "Nomor Kuesioner",            "yes"),
        ("alamat",            "Alamat Lengkap",             "no"),
        ("kelurahan",         "Kelurahan",                  "no"),
        ("kecamatan",         "Kecamatan",                  "no"),
        ("rt",                "RT",                         "no"),
        ("rw",                "RW",                         "no"),
        ("kota",              "Kota",                       "yes"),
        ("telp_rumah",        "Nomor Telepon Rumah",        "no"),
        ("telp_hp",           "Nomor HP / Telepon Kantor",  "no"),
        ("email",             "Alamat Email",               "no"),
        ("nama_interviewer",  "Nama Interviewer",           "yes"),
        ("no_interviewer",    "No. Interviewer",            "yes"),
        ("tanggal_wawancara", "Tanggal Wawancara",          "yes"),
        ("waktu_mulai",       "Waktu Mulai",                "yes"),
        ("waktu_selesai",     "Waktu Selesai",              "no"),
    ]

    extra_fields: list[tuple] = []
    if info_tbl is not None:
        known_labels = {f[1].upper() for f in standard_fields}
        for row in info_tbl.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                if ":" in txt:
                    label_raw = txt.split(":")[0].strip()
                    if (label_raw and len(label_raw) > 3
                            and label_raw.upper() not in known_labels
                            and not any(m in label_raw.upper() for m in
                                        ["TANDA TANGAN", "DIPERIKSA", "SPV", "TL", "ESOMAR"])):
                        fname = re.sub(r'[^a-z0-9]', '_', label_raw.lower()).strip('_')
                        fname = re.sub(r'_+', '_', fname)
                        if fname and len(fname) > 2:
                            lbl_upper = label_raw.upper()
                            if lbl_upper not in known_labels:
                                known_labels.add(lbl_upper)
                                extra_fields.append((fname, label_raw, "no"))

    date_fields  = {"tanggal_wawancara"}
    time_fields  = {"waktu_mulai", "waktu_selesai"}
    kota_fields  = {"kota"}
    int_fields   = {"rt", "rw"}

    def _field_type(fname: str) -> str:
        if fname in date_fields:  return "date"
        if fname in time_fields:  return "time"
        if fname in kota_fields:  return "select_one list_kota"
        if fname in int_fields:   return "integer"
        return "text"

    questions: list[dict] = []
    questions.append({
        "id": "info_responden", "is_header": False,
        "section": "", "subsection": "",
        "question": "INFORMASI RESPONDEN", "hint": "",
        "route_type": "begin_group", "choices": [], "skip_logic": "",
    })

    all_fields = standard_fields + extra_fields
    for fname, flabel, freq in all_fields:
        questions.append({
            "id": fname, "is_header": False,
            "section": "INFORMASI RESPONDEN", "subsection": "",
            "question": flabel, "hint": "",
            "route_type": _field_type(fname), "choices": [], "skip_logic": "",
            "_required": freq,
        })

    questions.append({
        "id": "end_info_responden", "is_header": False,
        "section": "", "subsection": "",
        "question": "", "hint": "",
        "route_type": "end_group", "choices": [], "skip_logic": "",
    })

    return questions


def _find_questionnaire_tables(doc: docx.Document) -> tuple[Table | None, Table | None]:
    screening_tbl = main_tbl = None
    for tbl in doc.tables:
        if screening_tbl is None and _is_screening_table(tbl):
            screening_tbl = tbl
        elif main_tbl is None and _is_main_table(tbl):
            main_tbl = tbl
        if screening_tbl and main_tbl:
            break
    return screening_tbl, main_tbl


# ── Fix 3: _parse_awareness_block — ekstrak sub dari teks asli ───────────────

def _parse_awareness_block(q_id_prefix: str, cell: _Cell,
                           section: str, subsection: str,
                           merek_list: list[tuple],
                           product_name: str) -> list[dict]:
    """
    Parse blok awareness A1/A2/Y1/MD10/MD17.
    Ekstrak teks sub-pertanyaan langsung dari baris dokumen yang diakhiri (SA)/(MA)/(S)/(M).
    Fallback ke label generik jika pola tidak ditemukan.
    """
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)

    # ── Ekstrak sub dari teks asli ─────────────────────────────────────────
    # Cari baris yang berisi route marker di akhir: (SA) / (MA) / (S) / (M) / (OE)
    ROUTE_MARKER = re.compile(
        r'^(.+?)\s*\((SA|MA|S|M|OE)\)\s*$', re.IGNORECASE | re.DOTALL
    )
    inline_subs: list[tuple[str, str, str]] = []  # (letter, text, route)
    letter_seq = "abcdef"
    idx = 0

    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    for ln in lines:
        m = ROUTE_MARKER.match(ln)
        if not m:
            continue
        sub_text = m.group(1).strip()
        route_raw = m.group(2).upper()
        route = "SA" if route_raw in ("SA", "S") else "MA" if route_raw in ("MA", "M") else route_raw
        # Skip baris yang isinya instruksi DP
        if any(sub_text.upper().startswith(h) for h in _HINT_MARKERS):
            continue
        if len(sub_text) < 5:
            continue
        if idx < len(letter_seq):
            inline_subs.append((letter_seq[idx], sub_text, route))
            idx += 1

    # ── Fallback label generik ─────────────────────────────────────────────
    if not inline_subs:
        inline_subs = [
            ("a", f"Merek {product_name} yang pertama kali Anda ingat? (TOM)",
             "SA"),
            ("b", f"Merek {product_name} lain yang Anda ketahui? (Spontan)",
             "MA"),
            ("c", f"Merek {product_name} mana yang Anda ketahui? (Aided)",
             "MA"),
            ("d", f"Merek {product_name} yang paling Anda favoritkan? (Favorit 1)",
             "SA"),
            ("e", f"Merek {product_name} favorit kedua Anda? (Favorit 2)",
             "SA"),
        ]

    # ── Cek apakah ada sub-kolom sumber tahu (f) di nested table ──────────
    sumber_choices: list[dict] = []
    has_f = False
    for nt in cell.tables:
        h = [c.text.strip() for c in nt.rows[0].cells] if nt.rows else []
        if any(f'{q_id_prefix}f' in str(hh).lower() for hh in h):
            has_f = True
            for row in nt.rows[1:]:
                cells_nt = row.cells
                label = cells_nt[0].text.strip()
                code  = cells_nt[1].text.strip() if len(cells_nt) > 1 else ""
                if label and not re.match(r'^[A-Z]\d+', label) and code:
                    sumber_choices.append({"code": code, "label": label, "routing": ""})
            break

    if has_f and len(inline_subs) < 6:
        inline_subs.append((
            "f",
            f"Dari mana Anda pertama kali mengetahui merek {product_name}? (Sumber tahu)",
            "MA",
        ))

    merek_choices = _choices_from_list(merek_list)
    skip_matches = re.findall(
        r'(?:TANYAKAN|DITANYAKAN)\s+(?:JIKA|KEPADA)\s+[^\.]{5,80}', raw, re.I
    )
    skip_logic = "; ".join(skip_matches)

    items = []
    for suffix, label, route in inline_subs:
        full_id = f"{q_id_prefix}{suffix}"
        choices = sumber_choices if suffix == "f" else merek_choices
        hint_sub = hint if suffix == "a" else ""
        items.append(_make_q(
            q_id=full_id, section=section, subsection=subsection,
            question=label, hint=hint_sub, route_type=route,
            choices=choices,
            skip_logic=skip_logic if suffix == "a" else "",
            raw_text=raw,
        ))

    return items


# ── Fix 4: _extract_inline_subs — split multi-sub dari satu baris ─────────────

def _extract_inline_subs(q_id: str, raw_text: str,
                         section: str, subsection: str) -> list[dict] | None:
    """
    Coba split baris yang memuat beberapa sub-pertanyaan bernomor inline:
    "V1. ... (M)  V2. ... (M)  V3. ... (M)" atau
    "Q4a. ... (S)  Q4b. ... (OE)"
    atau "c1. ... (OE)"

    Return list[dict] jika ditemukan ≥2 sub, else None.
    """
    # Pola: ID bernomor diikuti teks dan route marker
    # Contoh: "V1.\tMerek ... (M)" / "Q4a. Merek ... (S)"
    sub_re = re.compile(
        r'([A-Za-z]{1,3}\d+[a-z]?)\s*[.\t]\s*(.+?)\s*\((SA|MA|S|M|OE)\)',
        re.IGNORECASE
    )
    matches = sub_re.findall(raw_text)

    if len(matches) < 2:
        return None

    # Filter sub yang prefix-nya cocok dengan q_id induk
    prefix = q_id.rstrip("0123456789").upper()
    filtered = [(sid, txt, rt) for sid, txt, rt in matches
                if sid.upper().startswith(prefix) or prefix in sid.upper()]

    if len(filtered) < 2:
        # Kalau prefix tidak cocok tapi ada banyak sub, tetap pakai semua
        if len(matches) >= 3:
            filtered = matches
        else:
            return None

    items = []
    for sid, sub_text, route_raw in filtered:
        route = "SA" if route_raw.upper() in ("SA", "S") else \
                "MA" if route_raw.upper() in ("MA", "M") else route_raw.upper()
        sub_text_clean = sub_text.strip()
        # Hapus instruksi trailing (mis. "Apalagi? Apalagi?")
        sub_text_clean = re.sub(r'\s+(Apalagi\??\s*)+$', '', sub_text_clean, flags=re.I).strip()

        skip_m = re.findall(
            r'(?:TANYAKAN|DITANYAKAN)\s+(?:JIKA|KEPADA)\s+[^\.]{5,60}',
            sub_text_clean, re.I
        )
        skip_logic = "; ".join(skip_m)

        items.append(_make_q(
            q_id=sid.lower(), section=section, subsection=subsection,
            question=sub_text_clean, hint="",
            route_type=route, choices=[],
            skip_logic=skip_logic, raw_text=raw_text,
        ))
    return items


# ── Special parsers (unchanged from v8, abbreviated docstrings) ────────────────

def _parse_usage_block(q_id_prefix: str, cell: _Cell,
                       section: str, subsection: str,
                       merek_list: list[tuple],
                       product_name: str) -> list[dict]:
    """Parse blok usage V/Z via inline sub extraction, fallback ke template."""
    raw = _cell_text(cell)

    # Try inline sub extraction first (Fix 4)
    inline = _extract_inline_subs(q_id_prefix.upper(), raw, section, subsection)
    if inline and len(inline) >= 3:
        # Inject merek choices where empty
        merek_choices = _choices_from_list(merek_list)
        for q in inline:
            if not q.get("choices") and q["route_type"] in ("SA", "MA"):
                # Check if text mentions "alasan"
                if re.search(r'alasan|mengapa', q["question"], re.I):
                    pass  # leave empty, will be filled from nested tables
                else:
                    q["choices"] = merek_choices
        # Fill alasan choices from nested tables
        alasan_choices: list[dict] = []
        for nt in cell.tables:
            ch = _parse_simple_nested_table(nt)
            if len(ch) > 1:
                alasan_choices = ch
                break
        for q in inline:
            if not q.get("choices") and re.search(r'alasan|mengapa', q["question"], re.I):
                q["choices"] = alasan_choices
        return inline

    # Fallback: template-based (v8 behavior)
    merek_choices = _choices_from_list(merek_list)
    usia_choices  = _choices_from_list(_USIA_RANGE)
    time_choices  = _choices_from_list(_TIME_RANGE)
    alasan_choices: list[dict] = []
    alasan_negatif_choices: list[dict] = []
    for nt in cell.tables:
        ch = _parse_simple_nested_table(nt)
        if not ch:
            continue
        first_label = ch[0].get("label","").lower()
        if "mudah didapat" in first_label or "terjangkau" in first_label:
            alasan_choices = ch
        elif "sulit didapat" in first_label or "mahal" in first_label:
            alasan_negatif_choices = ch

    p = q_id_prefix.lower()
    sub_defs = [
        (f"{p}1",   "MA", f"Merek {product_name} apa saja yang pernah Anda gunakan?",
                    "TUNJUKKAN KARTU BANTU", merek_choices),
        (f"{p}1a",  "OE", f"Mengapa Anda menggunakan merek-merek {product_name} tersebut?",
                    "PROBE", alasan_choices),
        (f"{p}2",   "MA", f"Merek {product_name} apa yang Anda gunakan dalam 1 tahun terakhir?",
                    "TUNJUKKAN KARTU BANTU", merek_choices),
        (f"{p}2a",  "OE", f"Mengapa Anda menggunakan merek-merek {product_name} (1 tahun terakhir)?",
                    "PROBE", alasan_choices),
        (f"{p}3",   "MA", f"Merek {product_name} apa yang Anda gunakan terakhir kali?",
                    "TUNJUKKAN KARTU BANTU", merek_choices),
        (f"{p}3a",  "OE", f"Mengapa Anda menggunakan merek {product_name} tersebut (terakhir kali)?",
                    "PROBE", alasan_choices),
        (f"{p}4a",  "SA", f"Merek {product_name} yang paling sering Anda gunakan saat ini? (BUMO)",
                    "TUNJUKKAN KARTU BANTU – satu merek", merek_choices),
        (f"{p}4b",  "SA", f"Sejak usia berapa pertama kali Anda membeli {product_name} merek tersebut?",
                    "PILIH RANGE USIA", usia_choices),
        (f"{p}4c",  "SA", f"Merek {product_name} apa yang Anda gunakan sebelum merek saat ini?",
                    "TUNJUKKAN KARTU BANTU", merek_choices),
        (f"{p}4c1", "MA", f"Mengapa Anda beralih dari merek sebelumnya ke merek {product_name} saat ini?",
                    "PROBE – semua alasan", alasan_choices),
        (f"{p}4d",  "SA", f"Kapan pertama kali Anda beralih ke merek {product_name} saat ini?",
                    "PILIH RANGE WAKTU", time_choices),
        (f"{p}4e",  "OE", f"Apa yang mendorong Anda pertama kali mencoba merek {product_name} saat ini?",
                    "SPONTAN", []),
        (f"{p}5",   "SA", f"Merek {product_name} terbaik menurut Anda?",
                    "TUNJUKKAN KARTU BANTU – satu merek", merek_choices),
    ]

    items = []
    for sub_id, route, label, hint_sub, choices in sub_defs:
        items.append(_make_q(
            q_id=sub_id, section=section, subsection=subsection,
            question=label, hint=hint_sub, route_type=route,
            choices=choices, raw_text=raw,
        ))
    return items


def _parse_q4_block(cell: _Cell, section: str, subsection: str,
                    merek_list: list[tuple]) -> list[dict]:
    """Pecah Q4 (BUMO) — inline extraction first, fallback to template."""
    raw = _cell_text(cell)

    # Try inline extraction
    inline = _extract_inline_subs("Q4", raw, section, subsection)
    if inline and len(inline) >= 3:
        merek_choices = _choices_from_list(merek_list)
        usia_choices  = _choices_from_list(_USIA_RANGE)
        time_choices  = _choices_from_list(_TIME_RANGE)
        alasan_choices: list[dict] = []
        for nt in cell.tables:
            ch = _parse_simple_nested_table(nt)
            if len(ch) > len(alasan_choices):
                alasan_choices = ch
        for q in inline:
            if not q.get("choices"):
                qid_l = q["id"].lower()
                if re.search(r'alasan|mengapa|ganti', q["question"], re.I):
                    q["choices"] = alasan_choices
                elif re.search(r'usia|umur', q["question"], re.I):
                    q["choices"] = usia_choices
                elif re.search(r'kapan|sejak|tahun lalu', q["question"], re.I):
                    q["choices"] = time_choices
                elif q["route_type"] in ("SA", "MA"):
                    q["choices"] = merek_choices
        return inline

    # Template fallback
    merek_choices = _choices_from_list(merek_list)
    usia_choices  = _choices_from_list(_USIA_RANGE)
    time_choices  = _choices_from_list(_TIME_RANGE)
    alasan_choices = []
    for nt in cell.tables:
        ch = _parse_simple_nested_table(nt)
        if len(ch) > len(alasan_choices):
            alasan_choices = ch

    pn = _PRODUCT_LABELS["primary"]
    return [
        _make_q("q4a",  section, subsection,
                f"Merek {pn} utama yang paling sering Anda gunakan saat ini? (BUMO)",
                "TUNJUKKAN KARTU BANTU – satu merek", "SA", merek_choices, raw_text=raw),
        _make_q("q4b",  section, subsection,
                f"Sejak usia berapa pertama kali Anda membeli {pn} merek tersebut?",
                "PILIH RANGE USIA", "SA", usia_choices, raw_text=raw),
        _make_q("q4c",  section, subsection,
                f"Merek {pn} apa yang Anda gunakan sebelum merek saat ini?",
                "TUNJUKKAN KARTU BANTU", "SA", merek_choices, raw_text=raw),
        _make_q("q4c1", section, subsection,
                f"Mengapa Anda beralih dari merek {pn} sebelumnya ke merek saat ini?",
                "PROBE – semua alasan", "MA", alasan_choices, raw_text=raw),
        _make_q("q4d",  section, subsection,
                f"Kapan pertama kali Anda beralih ke merek {pn} saat ini?",
                "PILIH RANGE WAKTU", "SA", time_choices, raw_text=raw),
        _make_q("q4e",  section, subsection,
                f"Apa yang mendorong Anda pertama kali mencoba merek {pn} saat ini?",
                "SPONTAN – terbuka", "OE", [], raw_text=raw),
    ]


def _parse_md1_block(cell: _Cell, section: str, subsection: str) -> list[dict]:
    raw = _cell_text(cell)
    platform_choices: list[dict] = []
    for nt in cell.tables:
        if len(nt.columns) >= 4:
            for row in nt.rows[2:]:
                cells = row.cells
                label = cells[0].text.strip()
                code  = cells[1].text.strip() if len(cells) > 1 else ""
                if label and not re.match(r'^MD\d', label) and code:
                    platform_choices.append({"code": code, "label": label, "routing": ""})
            break

    sub_defs = [
        ("md1a", "Sosial media yang Anda ketahui?",                 "SPONTAN + TUNJUKKAN KARTU"),
        ("md1b", "Sosial media yang Anda akses dalam 3 bulan terakhir?", "P3M"),
        ("md1c", "Sosial media yang Anda akses dalam 1 minggu terakhir?", "P1W"),
        ("md1d", "Sosial media yang Anda akses setiap hari?",       "SETIAP HARI"),
    ]
    return [
        _make_q(sid, section, subsection, label, hint, "MA",
                platform_choices, raw_text=raw)
        for sid, label, hint in sub_defs
    ]


def _parse_matrix_q(q_id: str, cell: _Cell, section: str, subsection: str,
                    merek_list: list[tuple], scale_type: str = "10") -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    items = []
    atribut_list: list[str] = []
    merek_choices = _choices_from_list(merek_list) if merek_list else []

    for nt in cell.tables:
        rows = list(nt.rows)
        if not rows:
            continue
        nc = len(nt.columns)
        if nc >= 4:
            for row in rows[1:]:
                atribut = row.cells[0].text.strip()
                if atribut and not re.match(r'^\[|^ROTASIKAN|^DP:', atribut, re.I):
                    atribut_list.append(atribut)
        elif nc == 2:
            for row in rows[1:]:
                atribut = row.cells[0].text.strip()
                if atribut and not re.match(r'^\[|^ROTASIKAN', atribut, re.I):
                    atribut_list.append(atribut)

    if not atribut_list:
        _warn(q_id, "Matrix atribut tidak ditemukan — fallback ke satu field")
        return [_make_q(q_id, section, subsection, q_text, hint, "", [], raw_text=raw,
                        parse_warning="Matrix atribut tidak ditemukan")]

    grp_id = f"grp_{q_id.lower()}"
    items.append(_make_q(grp_id, section, subsection, q_text, hint, "begin_group", [], raw_text=raw))
    skala_choices = _choices_from_list(_SKALA_10 if scale_type == "10" else _SKALA_11)

    for i, atribut in enumerate(atribut_list, 1):
        sub_id = f"{q_id.lower()}_{i:02d}"
        if merek_list and q_id.upper() in ("Q14C", "Q15A"):
            items.append(_make_q(sub_id, section, subsection,
                                 f"{atribut} — merek mana yang paling sesuai?",
                                 atribut, "SA", merek_choices, raw_text=atribut))
        else:
            items.append(_make_q(sub_id, section, subsection,
                                 atribut, hint, "SA", skala_choices, raw_text=atribut))

    items.append(_make_q(f"{grp_id}_end", section, subsection, "", "", "end_group", [], raw_text=""))
    return items


def _parse_q14a_block(q_id: str, cell: _Cell, section: str, subsection: str) -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    merek_choices = _choices_from_list(_get_merek("primary"))
    items = []
    atribut_list = []
    for nt in cell.tables:
        for row in nt.rows[1:]:
            val = row.cells[0].text.strip()
            if val and not re.match(r'^ROTASIKAN', val, re.I):
                atribut_list.append(val)

    if not atribut_list:
        return [_make_q(q_id, section, subsection, q_text, hint, "SA",
                        merek_choices, raw_text=raw,
                        parse_warning="Atribut tidak ditemukan di nested table")]

    items.append(_make_q(f"grp_{q_id.lower()}", section, subsection,
                         q_text, hint, "begin_group", [], raw_text=raw))
    for i, attr in enumerate(atribut_list, 1):
        items.append(_make_q(f"{q_id.lower()}_{i:02d}", section, subsection,
                             f"{q_text} — {attr}", attr, "SA", merek_choices, raw_text=attr))
    items.append(_make_q(f"grp_{q_id.lower()}_end", section, subsection,
                         "", "", "end_group", [], raw_text=""))
    return items


def _parse_q15b_block(cell: _Cell, section: str, subsection: str) -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    pn = _PRODUCT_LABELS["primary"]
    items = [
        _make_q("grp_q15b", section, subsection,
                f"Urutkan merek {pn} dari yang paling tua sampai paling baru",
                hint, "begin_group", [], raw_text=raw)
    ]
    for code, label in _get_merek("primary")[:10]:
        items.append(_make_q(f"q15b_{code}", section, subsection,
                             f"Peringkat ke berapa merek {label}? (1=tertua, 10=termuda)",
                             "Masukkan angka 1–10", "integer", [], raw_text=label))
    items.append(_make_q("grp_q15b_end", section, subsection,
                         "", "", "end_group", [], raw_text=""))
    return items


def _parse_slogan_q(q_id: str, cell: _Cell, section: str, subsection: str,
                    merek_list: list[tuple]) -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    if not q_text or "untuk merek apa" not in q_text.lower():
        q_text = "Untuk merek apa slogan tersebut?"
    merek_choices = _choices_from_list(merek_list)
    return [_make_q(q_id, section, subsection, q_text, hint, "SA",
                    merek_choices, raw_text=raw)]


def _parse_m1_block(cell: _Cell, section: str, subsection: str) -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    media_list: list[tuple] = []
    freq_choices: list[dict] = []

    for nt in cell.tables:
        if len(nt.columns) >= 8:
            rows = list(nt.rows)
            if len(rows) > 1:
                freq_labels = [c.text.strip() for c in rows[1].cells[3:]]
                freq_choices = [{"code": str(i+1), "label": lbl, "routing": ""}
                                for i, lbl in enumerate(freq_labels) if lbl]
            for row in rows[2:]:
                cells = row.cells
                media_label = cells[0].text.strip()
                code = cells[2].text.strip()
                if media_label and not re.match(r'^ROTASIKAN', media_label, re.I):
                    media_list.append((code or str(len(media_list)+1), media_label))
            break

    if not media_list:
        _warn("m1", "Daftar media tidak ditemukan — fallback ke satu field MA")
        return [_make_q("m1", section, subsection, q_text, hint, "MA", [], raw_text=raw,
                        parse_warning="Daftar media tidak ditemukan")]

    items = [_make_q("grp_m1", section, subsection, q_text, hint,
                     "begin_group", [], raw_text=raw)]
    m1a_choices = _choices_from_list(media_list)
    items.append(_make_q("m1a", section, subsection,
                         "Kegiatan media apa saja yang pernah Anda lakukan?",
                         "MA – pilih semua yang berlaku", "MA", m1a_choices, raw_text=raw))
    for code, media_label in media_list:
        safe = re.sub(r'[^a-z0-9]', '_', media_label.lower())[:20].strip('_')
        items.append(_make_q(f"m1b_{safe}", section, subsection,
                             f"Seberapa sering Anda {media_label.lower()}?",
                             "PILIH SATU FREKUENSI", "SA", freq_choices,
                             skip_logic=f"TANYAKAN JIKA m1a TERKODE {code}",
                             raw_text=media_label))
    items.append(_make_q("grp_m1_end", section, subsection, "", "", "end_group", [], raw_text=""))
    return items


def _parse_rating_q(q_id: str, cell: _Cell, section: str, subsection: str,
                    scale_size: int = 10) -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    skip_matches = re.findall(
        r'(?:TANYAKAN|DITANYAKAN)\s+(?:JIKA|KEPADA)\s+[^\.]{5,80}', raw, re.I
    )
    skip_logic = "; ".join(skip_matches)
    choices = _choices_from_list(_SKALA_11 if scale_size == 11 else _SKALA_10)
    return [_make_q(q_id, section, subsection, q_text, hint,
                    "SA", choices, skip_logic=skip_logic, raw_text=raw)]


def _parse_brand_statement_matrix(q_id: str, cell: _Cell,
                                   section: str, subsection: str) -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    items = []
    merek_list: list[tuple] = []
    stmt_list: list[str] = []

    for nt in cell.tables:
        rows = list(nt.rows)
        if not rows or len(nt.columns) < 8:
            continue
        h = [c.text.strip() for c in rows[0].cells]
        if any(re.match(r'^(ROTASIKAN|Alderon|Intilon|Grest)', hh, re.I) for hh in h):
            for idx, label in enumerate(h[1:], 1):
                if label and not re.match(r'^ROTASIKAN', label, re.I):
                    merek_list.append((str(idx), label))
            for row in rows[1:]:
                stmt = row.cells[0].text.strip()
                if stmt and not re.match(r'^ROTASIKAN|^DP:', stmt, re.I):
                    stmt_list.append(stmt)
            break

    if not merek_list or not stmt_list:
        _warn(q_id, "Brand statement matrix: merek/pernyataan tidak ditemukan — fallback SA")
        return [_make_q(q_id, section, subsection, q_text, hint,
                        "SA", _choices_from_list(_get_merek("primary")), raw_text=raw,
                        parse_warning="Brand statement matrix tidak terdeteksi")]

    merek_choices = _choices_from_list(merek_list)
    items.append(_make_q(f"grp_{q_id}", section, subsection,
                         q_text, hint, "begin_group", [], raw_text=raw))
    for i, stmt in enumerate(stmt_list, 1):
        items.append(_make_q(f"{q_id}_{i:02d}", section, subsection,
                             stmt, hint, "SA", merek_choices, raw_text=stmt))
    items.append(_make_q(f"grp_{q_id}_end", section, subsection,
                         "", "", "end_group", [], raw_text=""))
    return items


def _parse_media_matrix_q(q_id: str, cell: _Cell,
                           section: str, subsection: str) -> list[dict]:
    raw = _cell_text(cell)
    q_text, hint = _split_hint(raw)
    skala_choices = _choices_from_list(_SKALA_10)
    items = []
    media_list: list[tuple] = []

    for nt in cell.tables:
        rows = list(nt.rows)
        if not rows or len(nt.columns) < 10:
            continue
        for row in rows[1:]:
            cells_nt = [c.text.strip() for c in row.cells]
            label = cells_nt[0]
            if label and not re.match(r'^ROTASIKAN|^DP:|^$', label, re.I):
                safe = re.sub(r'[^a-z0-9]', '_', label.lower())[:20].strip('_')
                media_list.append((safe, label))
        break

    if not media_list:
        _warn(q_id, "Media matrix: daftar media tidak ditemukan — fallback SA skala")
        return [_make_q(q_id, section, subsection, q_text, hint,
                        "SA", skala_choices, raw_text=raw,
                        parse_warning="Media matrix tidak terdeteksi")]

    items.append(_make_q(f"grp_{q_id}", section, subsection,
                         q_text, hint, "begin_group", [], raw_text=raw))
    for safe, label in media_list:
        items.append(_make_q(f"{q_id}_{safe}", section, subsection,
                             f"{label} — {q_text.lower()}", hint,
                             "SA", skala_choices, raw_text=label))
    items.append(_make_q(f"grp_{q_id}_end", section, subsection,
                         "", "", "end_group", [], raw_text=""))
    return items


# ── Main row parser ────────────────────────────────────────────────────────────

def _parse_question_row(q_id: str, cell: _Cell, route_raw: str,
                        section: str, subsection: str) -> list[dict]:
    raw_text  = _cell_text(cell)
    q_text, hint_text = _split_hint(raw_text)
    route_type = _normalize_route(route_raw)
    q_id_upper = q_id.upper()

    # ── Awareness blocks ─────────────────────────────────────────────────────
    if q_id_upper == "A1":
        return _parse_awareness_block("a1", cell, section, subsection,
                                      _get_merek("primary"), _PRODUCT_LABELS["primary"])
    if q_id_upper == "A2":
        return _parse_awareness_block("a2", cell, section, subsection,
                                      _get_merek("secondary"), _PRODUCT_LABELS["secondary"])
    if q_id_upper == "Y1":
        return _parse_awareness_block("y1", cell, section, subsection,
                                      _get_merek("tertiary"), _PRODUCT_LABELS["tertiary"])
    # MD10, MD17 juga bisa awareness block
    if q_id_upper in ("MD10", "MD17"):
        markers = re.findall(r'\((SA|MA|S|M)\)', raw_text, re.I)
        if len(markers) >= 3:
            return _parse_awareness_block(q_id.lower(), cell, section, subsection,
                                          _get_merek("primary"), _PRODUCT_LABELS["primary"])

    # ── Usage blocks (V/Z) ──────────────────────────────────────────────────
    if q_id_upper == "V":
        return _parse_usage_block("V", cell, section, subsection,
                                  _get_merek("secondary"), _PRODUCT_LABELS["secondary"])
    if q_id_upper == "Z":
        return _parse_usage_block("Z", cell, section, subsection,
                                  _get_merek("tertiary"), _PRODUCT_LABELS["tertiary"])

    # ── Q4 BUMO block ────────────────────────────────────────────────────────
    if q_id_upper == "Q4":
        return _parse_q4_block(cell, section, subsection, _get_merek("primary"))

    # ── Q6 inline subs ───────────────────────────────────────────────────────
    if q_id_upper == "Q6" or q_id_upper == "V6" or q_id_upper == "Z6":
        inline = _extract_inline_subs(q_id_upper, raw_text, section, subsection)
        if inline and len(inline) >= 2:
            choices_list = []
            for nt in cell.tables:
                ch = _parse_simple_nested_table(nt)
                if ch:
                    choices_list = ch
                    break
            for q in inline:
                if not q.get("choices") and q["route_type"] in ("SA", "MA"):
                    q["choices"] = choices_list
            return inline

    # ── MD1 block ────────────────────────────────────────────────────────────
    if q_id_upper == "MD1":
        return _parse_md1_block(cell, section, subsection)

    # ── Matrix pertanyaan ─────────────────────────────────────────────────────
    if q_id_upper in ("Q9B",):
        return _parse_matrix_q("Q9b", cell, section, subsection, [], "10")
    if q_id_upper == "Q10":
        return _parse_matrix_q("Q10", cell, section, subsection, [], "10")
    if q_id_upper in ("Q14C",):
        return _parse_matrix_q("Q14c", cell, section, subsection, _get_merek("primary"), "10")
    if q_id_upper in ("Q15A",):
        return _parse_matrix_q("Q15a", cell, section, subsection, _get_merek("primary"), "10")

    # ── Q14a, Q14b ───────────────────────────────────────────────────────────
    if q_id_upper in ("Q14A",):
        return _parse_q14a_block("Q14a", cell, section, subsection)
    if q_id_upper in ("Q14B",):
        return _parse_q14a_block("Q14b", cell, section, subsection)

    # ── Q15b ranking merek ───────────────────────────────────────────────────
    if q_id_upper in ("Q15B",):
        return _parse_q15b_block(cell, section, subsection)

    # ── Slogan attribution ───────────────────────────────────────────────────
    slogan_match = re.match(r'^([QVZ]\d+[A-Z]?)B$', q_id_upper)
    if slogan_match:
        q_text_lower = q_text.lower()
        if "slogan" in q_text_lower or "untuk merek apa" in q_text_lower:
            merek = (_get_merek("secondary") if q_id_upper.startswith("V") else
                     _get_merek("tertiary")   if q_id_upper.startswith("Z") else
                     _get_merek("primary"))
            return _parse_slogan_q(q_id.lower(), cell, section, subsection, merek)

    # ── Q8a/Q8b/Q8c/Q8d — brand statement matrix ─────────────────────────────
    if q_id_upper in ("Q8A","Q8B","Q8C","Q8D"):
        return _parse_brand_statement_matrix(q_id.lower(), cell, section, subsection)

    # ── E5, M12, MD8, MD26A — matrix ─────────────────────────────────────────
    if q_id_upper == "E5":
        return _parse_matrix_q("E5", cell, section, subsection, [], "10")
    if q_id_upper == "M12":
        return _parse_media_matrix_q("m12", cell, section, subsection)
    if q_id_upper == "MD8":
        return _parse_media_matrix_q("md8", cell, section, subsection)
    if q_id_upper == "MD25":
        return [_make_q(q_id.lower(), section, subsection, q_text, hint_text,
                        "OE", [], raw_text=raw_text)]
    if q_id_upper == "MD26A":
        return _parse_matrix_q("MD26a", cell, section, subsection, [], "10")

    # ── Rating 1-10 / NPS ────────────────────────────────────────────────────
    nps_ids  = {"V10","Z10","MD12","MD19","Q13","Q12A","Q12B"}
    rate_ids = {
        "V7","V8","V9","V11","Z7","Z8","Z9","Z11","V13C","Z11C",
        "Q11","Q16C","Q17C","Q18C","Q19C","Q20C","Q22C","Q23C",
        "E4","E6","E7","F3","F4","S6",
    }
    if q_id_upper in nps_ids:
        return _parse_rating_q(q_id.lower(), cell, section, subsection, 11)
    if q_id_upper in rate_ids:
        return _parse_rating_q(q_id.lower(), cell, section, subsection, 10)

    # ── Auto-detect skala dari nested table ───────────────────────────────────
    for nt in cell.tables:
        if len(nt.columns) < 10:
            continue
        rows_nt = list(nt.rows)
        check_row = rows_nt[-1] if rows_nt else None
        if check_row:
            vals = [c.text.strip() for c in check_row.cells]
            nums = [v for v in vals if re.match(r'^\d+$', v)]
            if len(nums) >= 8:
                scale_size = 11 if "0" in vals and len(nt.columns) >= 11 else 10
                return _parse_rating_q(q_id.lower(), cell, section, subsection, scale_size)

    # ── Fix 4: Inline sub-pertanyaan (Q4, Q6 dan lainnya yg belum ditangkap) ─
    markers = re.findall(r'\((SA|MA|S|M|OE)\)', raw_text, re.I)
    if len(markers) >= 2:
        inline = _extract_inline_subs(q_id_upper, raw_text, section, subsection)
        if inline and len(inline) >= 2:
            # Coba inject choices dari nested tables
            all_nt_choices: list[dict] = []
            for nt in cell.tables:
                ch = _parse_simple_nested_table(nt)
                if ch:
                    all_nt_choices = ch
                    break
            for q in inline:
                if not q.get("choices") and q["route_type"] in ("SA", "MA") and all_nt_choices:
                    q["choices"] = all_nt_choices
            return inline

    # ── Standard question ─────────────────────────────────────────────────────
    all_choices: list[dict] = []
    for nt in cell.tables:
        if len(nt.columns) >= 10:
            continue
        all_choices.extend(_parse_simple_nested_table(nt))

    all_choices = [
        c for c in all_choices
        if not re.match(r'^\[DP|^ROTASIKAN|^DP:', c.get("label",""), re.I)
        and not re.match(r'^\[DP|^ROTASIKAN', c.get("code",""), re.I)
    ]

    is_grid = bool(re.search(r'\bROTASIKAN\b', raw_text, re.I))

    skip_matches = re.findall(
        r'(?:TANYAKAN|DITANYAKAN)\s+(?:JIKA|KEPADA)\s+[^\.]{5,80}',
        raw_text, re.I
    )
    skip_logic = "; ".join(skip_matches).strip()

    # Flag pertanyaan yang masuk fallback dengan choices kosong dan bukan OE
    warn_msg = ""
    _OE_IDS = {"e3a","e3b","e3c","md25"}
    if (not all_choices and route_type in ("SA", "MA") and not is_grid
            and q_id.lower() not in _OE_IDS):
        if not re.search(r'(KARTU BANTU|LIST BRAND|MASUKAN LIST)', raw_text, re.I):
            warn_msg = f"Choices kosong (route={route_type}) — perlu cek manual"

    return [_make_q(
        q_id=q_id, section=section, subsection=subsection,
        question=q_text, hint=hint_text,
        route_type=route_type, choices=all_choices,
        skip_logic=skip_logic, is_grid=is_grid, raw_text=raw_text,
        parse_warning=warn_msg,
    )]


# ── Table walker ───────────────────────────────────────────────────────────────

def _walk_table(tbl: Table, section_default: str) -> list[dict]:
    items: list[dict] = []
    section    = section_default
    subsection = ""

    brand_grid_registry: dict[str, list[dict]] = {}

    for row in tbl.rows:
        cells   = [c.text.strip() for c in row.cells]
        q_id    = cells[0]
        q_cell  = row.cells[1]

        if not q_id:
            for nt in q_cell.tables:
                if len(nt.columns) < 6 or len(nt.rows) < 3:
                    continue
                grid_data = _parse_brand_grid_table(nt)
                brand_grid_registry.update(grid_data)

        if q_id:
            for nt in q_cell.tables:
                if len(nt.columns) >= 4 and len(nt.rows) >= 2:
                    h0 = [c.text.strip().lower() for c in nt.rows[0].cells]
                    if any(re.match(r'^s8[abc]$', h) for h in h0):
                        grid_data = _parse_brand_grid_table(nt)
                        brand_grid_registry.update(grid_data)

    # Pass 2: parse questions
    for row in tbl.rows:
        cells  = [c.text.strip() for c in row.cells]
        q_id   = cells[0]
        route  = cells[2] if len(cells) > 2 else ""
        raw_q  = _cell_text(row.cells[1])

        if not q_id and not raw_q:
            continue

        if not q_id:
            text_up = raw_q.strip()
            if not text_up or _NON_SECTION_RE.match(text_up):
                continue
            if len(text_up) < 100:
                if text_up.upper() in ("SCREENING", "KUESIONER UTAMA"):
                    section    = text_up
                    subsection = ""
                else:
                    subsection = text_up
                items.append(_make_q(
                    "", section, subsection, text_up, "",
                    "info", [], is_header=True, raw_text=text_up,
                ))
            continue

        if not _Q_ID_RE.match(q_id):
            continue

        parsed = _parse_question_row(q_id, row.cells[1], route, section, subsection)

        for q in parsed:
            if not q.get("choices"):
                qid_norm = q["id"].lower().replace("-","_")
                qid_hyp  = q["id"].lower()
                chosen = (brand_grid_registry.get(qid_norm)
                          or brand_grid_registry.get(qid_hyp)
                          or brand_grid_registry.get(qid_norm.rstrip("_1").rstrip("_"))
                          or [])
                if chosen:
                    q["choices"] = chosen
                    # Remove any spurious parse_warning if choices now resolved
                    q.pop("_parse_warning", None)

            qid_up = q["id"].upper()
            if not q.get("choices"):
                if re.match(r'^Q6B-?1$', qid_up):
                    q["choices"] = _choices_from_list(_get_merek("primary"))
                elif re.match(r'^V6B-?1$', qid_up):
                    q["choices"] = _choices_from_list(_get_merek("secondary"))
                elif re.match(r'^Z6B-?1$', qid_up):
                    q["choices"] = _choices_from_list(_get_merek("tertiary"))

            if q["id"].lower() in ("s8a","s8b","s8c"):
                if not q.get("choices"):
                    q["choices"] = _choices_from_list(_CHOICES_PRODUK)

            if q["id"].lower() in ("s8d","s8e"):
                if not q.get("choices"):
                    q["choices"] = _choices_from_list(_CHOICES_JENIS_PRODUK)

            if not q.get("choices") and re.search(
                r'MASUKAN LIST BRAND|LIST BRAND PIPA', q.get("raw_text",""), re.I
            ):
                q["choices"] = _choices_from_list(_get_merek("primary"))

            if q["id"].lower() == "p7b" and not q.get("choices"):
                q["choices"] = _choices_from_list(_get_merek("primary"))

            if q["id"].lower() in ("e3a","e3b","e3c"):
                q["choices"]    = []
                q["route_type"] = "OE"

            if q["id"].lower() in ("e4","e6","e7","f3","f4") and not q.get("choices"):
                q["choices"]    = _choices_from_list(_SKALA_10)
                q["route_type"] = "SA"

        items.extend(parsed)

    return items


# ── Public API ─────────────────────────────────────────────────────────────────

def parse_questionnaire(docx_path: str) -> tuple[list[dict], list[dict]]:
    """
    Parse kuesioner dan kembalikan (questions, parse_warnings).
    parse_warnings: list of {id, issue} untuk pertanyaan yang perlu dicek manual.
    """
    global _PARSE_WARNINGS
    _PARSE_WARNINGS = []  # Reset untuk setiap parse baru

    doc = docx.Document(docx_path)

    _extract_brand_lists_from_doc(doc)
    kota_choices = _extract_kota_from_doc(doc)
    info_questions = _extract_respondent_info(doc)

    info_questions.insert(0, {
        "id": "_kota_choices_data",
        "is_header": True,
        "section": "", "subsection": "",
        "question": "", "hint": "",
        "route_type": "kota_choices_data",
        "choices": [{"code": c, "label": l, "routing": ""} for c, l in kota_choices],
        "skip_logic": "",
    })

    screening_tbl, main_tbl = _find_questionnaire_tables(doc)

    if screening_tbl is None and main_tbl is None:
        raise ValueError(
            "Tidak ditemukan tabel kuesioner 3-kolom. "
            "Pastikan file adalah kuesioner Deka Research format tabel Word."
        )

    questions: list[dict] = []
    questions.extend(info_questions)

    if screening_tbl:
        questions.extend(_walk_table(screening_tbl, "SCREENING"))

    if main_tbl:
        questions.extend(_walk_table(main_tbl, "KUESIONER UTAMA"))

    warnings = list(_PARSE_WARNINGS)
    return questions, warnings


def parse_to_json(docx_path: str, output_path: str | None = None) -> str:
    questions, warnings = parse_questionnaire(docx_path)
    q_ids = [q["id"] for q in questions if q["id"] and not q["is_header"]
             and q["route_type"] not in ("begin_group","end_group")]
    result = {
        "_meta": {
            "source":          Path(docx_path).name,
            "total_rows":      len(questions),
            "total_questions": len(q_ids),
            "question_ids":    q_ids,
            "parse_warnings":  warnings,
        },
        "questions": questions,
    }
    js = json.dumps(result, ensure_ascii=False, indent=2)
    if output_path:
        Path(output_path).write_text(js, encoding="utf-8")
    return js


if __name__ == "__main__":
    fpath = sys.argv[1] if len(sys.argv) > 1 else None
    opath = sys.argv[2] if len(sys.argv) > 2 else None
    if not fpath:
        print("Usage: python docx_parser.py <input.docx> [output.json]")
        sys.exit(1)
    js = parse_to_json(fpath, opath)
    d = json.loads(js)
    m = d["_meta"]
    print(f"✓  {m['total_questions']} pertanyaan → {opath or 'stdout'}")
    print(f"   First 15 IDs: {m['question_ids'][:15]}")
    if m["parse_warnings"]:
        print(f"\n⚠  {len(m['parse_warnings'])} item perlu dicek manual:")
        for w in m["parse_warnings"][:10]:
            print(f"   [{w['id']}] {w['issue']}")
