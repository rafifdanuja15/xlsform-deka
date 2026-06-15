"""
llm_client.py — Claude + Parallel Processing

Strategi:
- Pre-process teks: hapus noise, hemat 20-40% chars
- Chunk 4000 chars (lebih kecil = response lebih cepat per chunk)
- Parallel processing dengan ThreadPoolExecutor
- max_tokens=6000 per chunk (cukup untuk 25-35 pertanyaan)
- temperature=0.0 untuk konsistensi JSON
- max_retries=0, timeout per chunk (bukan total)
"""

import os
import re
import json
import logging
import httpx
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from openai import OpenAI

logger = logging.getLogger(__name__)

SUMOPOD_API_KEY  = os.environ.get("SUMOPOD_API_KEY", "")
SUMOPOD_BASE_URL = os.environ.get("SUMOPOD_BASE_URL", "https://ai.sumopod.com/v1")
SUMOPOD_MODEL    = os.environ.get("SUMOPOD_MODEL", "claude-sonnet-4-6")

API_TIMEOUT  = int(os.environ.get("API_TIMEOUT",  "600"))   # 10 menit per chunk
CHUNK_SIZE   = int(os.environ.get("CHUNK_SIZE",   "4000"))  # chars per chunk
MAX_WORKERS  = int(os.environ.get("MAX_WORKERS",  "3"))     # chunk paralel sekaligus
MAX_TOKENS   = int(os.environ.get("MAX_TOKENS",   "6000"))  # output token per chunk

PROMPT_PATH = os.path.join(os.path.dirname(__file__), "prompt.md")


# ── System prompt ringkas sebagai fallback ────────────────────────────────────

SYSTEM_PROMPT_FALLBACK = """Kamu adalah konverter kuesioner ke XLSForm KoboToolbox.

OUTPUT: JSON murni saja — tanpa markdown, tanpa penjelasan, tanpa komentar.

FORMAT WAJIB:
{
  "survey": [
    {"type":"...", "name":"...", "label":"...", "required":"yes",
     "relevant":"...", "constraint":"...", "constraint_message":"...",
     "hint":"...", "appearance":"...", "calculation":"..."}
  ],
  "choices": [
    {"list_name":"...", "name":"...", "label":"..."}
  ],
  "settings": {
    "form_title":"...", "form_id":"...", "version":"1", "default_language":"Indonesian"
  }
}

ATURAN:
- type valid: text, integer, decimal, select_one <list>, select_multiple <list>,
  note, begin_group, end_group, begin_repeat, end_repeat, calculate, date,
  time, datetime, geopoint, image, audio, video, barcode
- name: snake_case unik tanpa spasi
- required: "yes" atau omit jika tidak wajib
- relevant: ODK expression (${var} = 'value')
- Omit semua field yang kosong/tidak ada
- choices: isi hanya untuk select_one / select_multiple"""


def _load_system_prompt() -> str:
    if os.path.exists(PROMPT_PATH):
        with open(PROMPT_PATH, "r", encoding="utf-8") as f:
            content = f.read().strip()
        if content:
            return content + "\n\nPENTING: Output HANYA JSON murni. Omit semua field kosong."
    return SYSTEM_PROMPT_FALLBACK


def _make_client() -> OpenAI:
    """Buat client dengan timeout per-chunk dan tanpa auto-retry."""
    return OpenAI(
        api_key=SUMOPOD_API_KEY,
        base_url=SUMOPOD_BASE_URL,
        timeout=httpx.Timeout(API_TIMEOUT),
        max_retries=0,
    )


# ── Pre-processing ─────────────────────────────────────────────────────────────

def _clean_text(text: str) -> str:
    """Hapus noise dari teks — bisa hemat 20-40% chars sebelum dikirim ke AI."""

    # Normalise line endings
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    lines = text.split('\n')

    # Hapus trailing whitespace tiap baris
    lines = [l.rstrip() for l in lines]

    # Hapus baris yang hanya berisi karakter dekoratif
    lines = [l for l in lines if not re.match(r'^[\s\-_=\.~\*\|]{3,}$', l)]

    # Hapus nomor halaman umum
    lines = [l for l in lines if not re.match(
        r'^\s*(page\s+\d+(\s+of\s+\d+)?|halaman\s+\d+|\-\s*\d+\s*\-)\s*$',
        l, re.IGNORECASE
    )]

    # Hapus baris berulang (header/footer yang muncul >3x)
    counts = Counter(l.strip() for l in lines if len(l.strip()) > 8)
    repeated = {l for l, c in counts.items() if c > 3}
    lines = [l for l in lines if l.strip() not in repeated]

    # Gabung kembali, collapse blank lines berlebihan
    text = '\n'.join(lines)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'[ \t]{2,}', ' ', text)

    return text.strip()


def _split_into_chunks(text: str, chunk_size: int) -> list[str]:
    """Pecah di batas paragraf agar tidak potong di tengah pertanyaan."""
    if len(text) <= chunk_size:
        return [text]

    chunks = []
    start = 0
    while start < len(text):
        end = start + chunk_size
        if end >= len(text):
            chunks.append(text[start:])
            break

        # Cari pemisah terbaik: paragraf > baris baru > spasi
        cut = -1
        for sep in ['\n\n', '\n', ' ']:
            pos = text.rfind(sep, start + chunk_size // 2, end)
            if pos > start:
                cut = pos + len(sep)
                break

        end = cut if cut > start else end
        chunks.append(text[start:end])
        start = end

    return [c.strip() for c in chunks if c.strip()]


# ── Single chunk call ──────────────────────────────────────────────────────────

def _call_chunk(chunk: str, idx: int, total: int, system_prompt: str) -> dict:
    """
    Konversi satu chunk ke XLSForm JSON.
    Setiap call membuat client sendiri — aman untuk threading.
    """
    client = _make_client()

    if total == 1:
        part_note = "Konversi SEMUA pertanyaan dalam kuesioner ini."
    elif idx == 1:
        part_note = f"Ini bagian 1/{total} kuesioner. Sertakan settings form."
    elif idx == total:
        part_note = f"Ini bagian terakhir ({idx}/{total}). Settings boleh {{}}."
    else:
        part_note = f"Ini bagian {idx}/{total}. Settings boleh {{}}."

    user_msg = f"""Konversi teks kuesioner berikut ke format XLSForm KoboToolbox. {part_note}

TEKS KUESIONER:
---
{chunk}
---

WAJIB output JSON persis dengan struktur ini (TIDAK BOLEH format lain):
{{
  "survey": [
    {{
      "type": "text",
      "name": "nama_variabel",
      "label": "Label pertanyaan",
      "required": "yes",
      "relevant": "",
      "hint": "",
      "constraint": "",
      "constraint_message": "",
      "appearance": "",
      "calculation": ""
    }}
  ],
  "choices": [
    {{
      "list_name": "nama_list",
      "name": "kode_pilihan",
      "label": "Label pilihan"
    }}
  ],
  "settings": {{}}
}}

ATURAN WAJIB:
- type HARUS salah satu: text, integer, decimal, select_one <list_name>, select_multiple <list_name>, note, begin_group, end_group, begin_repeat, end_repeat, calculate, date, time, geopoint, image
- name: snake_case unik (contoh: q1_merek, s3_usia)
- Pertanyaan pilihan: type="select_one nama_list" dan isi choices dengan list_name yang sama
- Omit field yang kosong/tidak relevan
- Output HANYA JSON — tanpa markdown, tanpa penjelasan, tanpa ```

JSON:"""

    logger.info(f"[chunk {idx}/{total}] → {SUMOPOD_MODEL} | {len(chunk)} chars")

    response = client.chat.completions.create(
        model=SUMOPOD_MODEL,
        max_tokens=MAX_TOKENS,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": user_msg},
        ],
        temperature=0.0,
    )

    raw    = response.choices[0].message.content
    finish = response.choices[0].finish_reason
    logger.info(f"[chunk {idx}/{total}] ← {len(raw)} chars | finish={finish}")

    if finish == "length":
        logger.warning(
            f"[chunk {idx}/{total}] Output terpotong! "
            "Coba kurangi CHUNK_SIZE atau naikkan MAX_TOKENS."
        )

    result = _parse_llm_response(raw, idx)

    # Validasi: jika survey kosong, coba sekali lagi dengan prompt lebih keras
    if not result.get("survey"):
        logger.warning(f"[chunk {idx}/{total}] Survey kosong, retry dengan prompt eksplisit...")
        result = _retry_chunk(client, chunk, idx, total, system_prompt)

    return result


def _retry_chunk(client: OpenAI, chunk: str, idx: int, total: int,
                 system_prompt: str) -> dict:
    """Retry dengan few-shot example agar Claude tidak salah format."""
    retry_msg = f"""PENTING: Output kamu sebelumnya salah format. Kamu HARUS output XLSForm JSON.

Contoh output BENAR untuk pertanyaan "Apa merek pipa favorit?" dengan pilihan Rucika/Wavin:
{{
  "survey": [
    {{"type": "select_one merek_pipa", "name": "merek_favorit", "label": "Apa merek pipa favorit?"}}
  ],
  "choices": [
    {{"list_name": "merek_pipa", "name": "rucika", "label": "Rucika"}},
    {{"list_name": "merek_pipa", "name": "wavin",  "label": "Wavin"}}
  ],
  "settings": {{}}
}}

Sekarang konversi teks ini dengan format yang SAMA PERSIS:
---
{chunk}
---

JSON (langsung mulai dengan {{, tanpa ```, tanpa penjelasan):"""

    response = client.chat.completions.create(
        model=SUMOPOD_MODEL,
        max_tokens=MAX_TOKENS,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": retry_msg},
        ],
        temperature=0.0,
    )
    raw = response.choices[0].message.content
    logger.info(f"[chunk {idx}/{total}] retry ← {len(raw)} chars")
    return _parse_llm_response(raw, idx)


# ── Main entry point ───────────────────────────────────────────────────────────

def call_llm_for_xlsform(questionnaire_text: str) -> dict:
    """
    Konversi teks kuesioner ke XLSForm JSON.

    Flow:
    1. Bersihkan teks (hemat token)
    2. Pecah jadi chunk 4000 chars
    3. Proses paralel (MAX_WORKERS chunk sekaligus)
    4. Merge semua hasil
    """
    if not SUMOPOD_API_KEY:
        raise ValueError(
            "SUMOPOD_API_KEY belum dikonfigurasi. "
            "Isi file .env — dapatkan key di https://sumopod.com/dashboard/ai/keys"
        )

    # Step 1: Bersihkan
    original_len = len(questionnaire_text)
    clean = _clean_text(questionnaire_text)
    saved = original_len - len(clean)
    logger.info(
        f"Pre-process: {original_len:,} → {len(clean):,} chars "
        f"(hemat {saved:,} chars / {100*saved//original_len}%)"
    )

    # Step 2: Chunk
    chunks = _split_into_chunks(clean, CHUNK_SIZE)
    total  = len(chunks)
    logger.info(
        f"Chunks: {total} | model={SUMOPOD_MODEL} | "
        f"timeout={API_TIMEOUT}s/chunk | max_tokens={MAX_TOKENS} | "
        f"workers={MAX_WORKERS} (paralel)"
    )

    system_prompt = _load_system_prompt()

    # Step 3: Proses
    if total == 1:
        # Dokumen pendek — langsung
        result = _call_chunk(chunks[0], 1, 1, system_prompt)
        return result

    # Dokumen panjang — paralel
    results = [None] * total
    failed  = []

    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, total)) as pool:
        futures = {
            pool.submit(_call_chunk, chunk, i + 1, total, system_prompt): i
            for i, chunk in enumerate(chunks)
        }
        for future in as_completed(futures):
            i = futures[future]
            try:
                results[i] = future.result()
                logger.info(f"[chunk {i+1}/{total}] ✓")
            except Exception as e:
                logger.error(f"[chunk {i+1}/{total}] ✗ {e}")
                failed.append(i + 1)
                results[i] = {"survey": [], "choices": [], "settings": {}}

    if failed:
        logger.warning(f"Chunk yang gagal: {failed} — hasil mungkin tidak lengkap")

    if all(not r.get("survey") for r in results if r):
        raise ValueError(
            "Semua chunk gagal dikonversi. "
            "Periksa API key, koneksi, atau coba lagi."
        )

    # Step 4: Merge
    return _merge_results(results)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _merge_results(results: list[dict]) -> dict:
    all_survey  = []
    all_choices = []
    settings    = None

    for r in results:
        if not r:
            continue
        all_survey.extend(r.get("survey", []))
        all_choices.extend(r.get("choices", []))
        if settings is None and r.get("settings"):
            s = r["settings"]
            if any(v for v in s.values()):
                settings = s

    # Dedup choices (list_name + name)
    seen, deduped = set(), []
    for c in all_choices:
        key = (c.get("list_name", ""), c.get("name", ""))
        if key not in seen:
            seen.add(key)
            deduped.append(c)

    if not settings:
        settings = {
            "form_title": "Converted Form",
            "form_id": "converted_form",
            "version": "1",
            "default_language": "Indonesian",
        }

    logger.info(f"Merge selesai: {len(all_survey)} survey rows, {len(deduped)} choices")
    return {"survey": all_survey, "choices": deduped, "settings": settings}


def _parse_llm_response(raw: str, chunk_idx: int = 0) -> dict:
    """Parse JSON dari respons, handle markdown fences dan prefix teks."""
    cleaned = raw.strip()

    # Hapus markdown fences
    cleaned = re.sub(r'^```json\s*\n?', '', cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r'^```\s*\n?',    '', cleaned, flags=re.MULTILINE)
    cleaned = re.sub(r'\n?```\s*$',    '', cleaned, flags=re.MULTILINE)
    cleaned = cleaned.strip()

    # Jika ada teks sebelum {, buang
    brace = cleaned.find('{')
    if brace > 0:
        cleaned = cleaned[brace:]

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r'\{[\s\S]*\}', cleaned)
        if match:
            try:
                data = json.loads(match.group(0))
            except json.JSONDecodeError:
                logger.warning(f"[chunk {chunk_idx}] Gagal parse JSON. Preview: {raw[:300]}")
                return {"survey": [], "choices": [], "settings": {}}
        else:
            logger.warning(f"[chunk {chunk_idx}] Tidak ada JSON ditemukan. Preview: {raw[:300]}")
            return {"survey": [], "choices": [], "settings": {}}

    # Defaults
    data.setdefault("survey",  [])
    data.setdefault("choices", [])
    data.setdefault("settings", {})

    # Bersihkan field kosong dari setiap row
    data["survey"] = [
        {k: v for k, v in row.items() if v not in (None, "", [])}
        for row in data["survey"] if isinstance(row, dict)
    ]
    data["choices"] = [
        {k: v for k, v in row.items() if v not in (None, "", [])}
        for row in data["choices"] if isinstance(row, dict)
    ]

    return data
