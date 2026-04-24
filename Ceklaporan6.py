import os
import re
import json
import time
import PyPDF2
import anthropic
import streamlit as st
from docx import Document
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# ──────────────────────────────────────────────
# KONSTANTA
# ──────────────────────────────────────────────
MODEL      = "claude-sonnet-4-5"
MAX_TOKENS = 8192
MAX_CHARS  = 40000

# Nama sheet di Google Spreadsheet
SHEET_RIWAYAT  = "riwayat_audit"
SHEET_REFERENSI = "data_laporan"

GSPREAD_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

MODE_CONFIG = {
    "🔍 Pre-Check": {
        "key": "precheck",
        "desc": "Pengecekan cepat ~30 detik. Fokus pada isu kritikal.",
        "instruction": (
            "Lakukan pengecekan cepat (Pre-Check). Fokus pada isu kritikal: "
            "inkonsistensi nilai, tanggal, luas, dan alamat. Ringkas dalam 5-8 temuan utama."
        ),
    },
    "🧠 Deep Audit": {
        "key": "deepaudit",
        "desc": "Audit menyeluruh setiap bagian laporan.",
        "instruction": (
            "Lakukan audit mendalam dan menyeluruh. Periksa setiap bagian laporan secara detail. "
            "Identifikasi semua inkonsistensi, kejanggalan naratif, kesalahan penulisan angka, "
            "dan potensi kesalahan material."
        ),
    },
    "📐 KEPI/MAPPI": {
        "key": "mappi",
        "desc": "Pengecekan kesesuaian standar SPI & MAPPI.",
        "instruction": (
            "Periksa kesesuaian dengan standar KEPI/MAPPI dan SPI (Standar Penilaian Indonesia). "
            "Apakah semua elemen wajib ada? Apakah metode penilaian sesuai standar? "
            "Apakah pengungkapan dan asumsi sudah lengkap?"
        ),
    },
    "🏢 Multi-Objek": {
        "key": "multiobj",
        "desc": "Untuk laporan CBDK/PANI style dengan banyak properti.",
        "instruction": (
            "Ini adalah laporan multi-objek/multi-properti. "
            "LANGKAH 1: Identifikasi terlebih dahulu berapa dan apa saja objek properti yang ada. "
            "LANGKAH 2: Untuk SETIAP objek, cek konsistensi secara terpisah — "
            "jangan campurkan data antar objek. "
            "LANGKAH 3: Tandai temuan dengan nama objek yang relevan."
        ),
    },
}

CHECK_ITEMS_DEFAULT = [
    "Konsistensi Tanggal (inspeksi, penilaian, laporan)",
    "Konsistensi Luas (tanah, bangunan, GFA, NLA)",
    "Konsistensi Alamat & Lokasi",
    "Konsistensi Nilai (angka vs huruf, ringkasan vs kesimpulan)",
    "Kepemilikan & Nomor Sertifikat",
    "Koreksi & Konsistensi NJOP",
    "Kesesuaian Standar KEPI/MAPPI",
    "Analisis Pasar & Data Pembanding",
    "Pendekatan & Metode Penilaian",
    "Kelengkapan Narasi & Deskripsi Objek",
]

SEVERITY_CONFIG = {
    "kritikal": {"emoji": "🔴", "color": "#c0392b", "bg": "#fff0f0"},
    "minor":    {"emoji": "🟡", "color": "#d4860a", "bg": "#fff8e6"},
    "ok":       {"emoji": "🟢", "color": "#1a9e67", "bg": "#edfaf4"},
    "info":     {"emoji": "🔵", "color": "#1e6fbf", "bg": "#eef4ff"},
}

SYSTEM_PROMPT = """Kamu adalah expert QA auditor laporan penilaian properti di Indonesia.
Kamu memahami standar KEPI, MAPPI, dan SPI (Standar Penilaian Indonesia) dengan sangat baik.
Tugasmu menganalisis laporan penilaian dan menemukan inkonsistensi, kesalahan, atau ketidaksesuaian standar.

SELALU berikan output HANYA dalam format JSON yang valid, tanpa teks apapun di luar JSON.
Gunakan struktur PERSIS berikut:

{
  "report_type": "tunggal atau multi-objek",
  "properties": ["nama/deskripsi properti 1", "nama/deskripsi properti 2"],
  "summary": {
    "total_findings": 0,
    "kritikal": 0,
    "minor": 0,
    "ok": 0,
    "info": 0,
    "overall_score": 85,
    "executive_summary": "Ringkasan 2-3 kalimat hasil audit secara keseluruhan."
  },
  "findings": [
    {
      "id": "F001",
      "severity": "kritikal",
      "category": "Nilai",
      "title": "Judul singkat temuan (maks 10 kata)",
      "detail": "Penjelasan detail: apa yang ditemukan, di mana, dan mengapa ini menjadi masalah.",
      "page_hint": "Hal. 3 / Bagian II",
      "property": "nama properti (kosong jika berlaku untuk semua)"
    }
  ]
}

Panduan severity:
- kritikal: kesalahan yang dapat mempengaruhi nilai atau validitas laporan
- minor: ketidakkonsistenan kecil atau potensi perbaikan
- ok: elemen yang sudah benar dan sesuai standar
- info: catatan atau saran yang perlu diperhatikan

overall_score: 0-100, di mana 100 = sempurna tanpa temuan kritikal/minor."""


# ══════════════════════════════════════════════
# GOOGLE SHEETS — CONNECTION & HELPERS
# ══════════════════════════════════════════════

@st.cache_resource(show_spinner=False)
def get_gsheet_client():
    """
    Buat koneksi ke Google Sheets menggunakan credentials dari st.secrets.
    Kembalikan (client, spreadsheet, error_msg).
    """
    # Cek apakah secrets ada
    if "gcp_service_account" not in st.secrets:
        return None, None, "KEY_MISSING: 'gcp_service_account' tidak ditemukan di Secrets"
    if "spreadsheet_id" not in st.secrets:
        return None, None, "KEY_MISSING: 'spreadsheet_id' tidak ditemukan di Secrets"
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=GSPREAD_SCOPES)
        client = gspread.authorize(creds)
        spreadsheet_id = st.secrets["spreadsheet_id"]
        spreadsheet = client.open_by_key(spreadsheet_id)
        return client, spreadsheet, None
    except Exception as e:
        return None, None, str(e)


def get_or_create_sheet(spreadsheet, sheet_name: str, headers: list):
    """Ambil worksheet, buat jika belum ada, lengkapi header."""
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(headers))
        ws.append_row(headers)
    # Pastikan header ada di baris pertama
    existing = ws.row_values(1)
    if existing != headers:
        ws.insert_row(headers, 1)
    return ws


# ── RIWAYAT AUDIT ──────────────────────────────

RIWAYAT_HEADERS = [
    "audit_id", "timestamp", "files", "mode",
    "score", "kritikal", "minor",
    "report_type", "properties", "executive_summary", "findings_json"
]

def load_riwayat(spreadsheet) -> dict:
    """
    Baca sheet riwayat_audit → dict {audit_id: {...}}.
    Fallback ke session_state jika sheets tidak tersedia.
    """
    if spreadsheet is None:
        return st.session_state.get("riwayat_local", {})
    try:
        ws = get_or_create_sheet(spreadsheet, SHEET_RIWAYAT, RIWAYAT_HEADERS)
        records = ws.get_all_records()
        result = {}
        for row in records:
            aid = str(row.get("audit_id", ""))
            if not aid:
                continue
            try:
                findings = json.loads(row.get("findings_json", "[]"))
            except Exception:
                findings = []
            try:
                props = json.loads(row.get("properties", "[]"))
            except Exception:
                props = []
            result[aid] = {
                "timestamp":         row.get("timestamp", ""),
                "files":             row.get("files", "").split("|"),
                "mode":              row.get("mode", ""),
                "score":             int(row.get("score", 0)),
                "kritikal":          int(row.get("kritikal", 0)),
                "minor":             int(row.get("minor", 0)),
                "result": {
                    "report_type":   row.get("report_type", ""),
                    "properties":    props,
                    "summary": {
                        "overall_score":     int(row.get("score", 0)),
                        "kritikal":          int(row.get("kritikal", 0)),
                        "minor":             int(row.get("minor", 0)),
                        "executive_summary": row.get("executive_summary", ""),
                    },
                    "findings": findings,
                },
            }
        return result
    except Exception as e:
        st.warning(f"⚠️ Gagal membaca riwayat dari Sheets: {e}")
        return st.session_state.get("riwayat_local", {})


def save_riwayat_row(spreadsheet, audit_id: str, data: dict):
    """Tambahkan satu baris riwayat ke sheet."""
    if spreadsheet is None:
        # Fallback: simpan di session_state
        local = st.session_state.get("riwayat_local", {})
        local[audit_id] = data
        st.session_state["riwayat_local"] = local
        return
    try:
        ws = get_or_create_sheet(spreadsheet, SHEET_RIWAYAT, RIWAYAT_HEADERS)
        result  = data.get("result", {})
        summary = result.get("summary", {})
        row = [
            audit_id,
            data.get("timestamp", ""),
            "|".join(data.get("files", [])),
            data.get("mode", ""),
            summary.get("overall_score", 0),
            summary.get("kritikal", 0),
            summary.get("minor", 0),
            result.get("report_type", ""),
            json.dumps(result.get("properties", []), ensure_ascii=False),
            summary.get("executive_summary", ""),
            json.dumps(result.get("findings", []), ensure_ascii=False),
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
    except Exception as e:
        st.warning(f"⚠️ Gagal menyimpan riwayat ke Sheets: {e}")
        local = st.session_state.get("riwayat_local", {})
        local[audit_id] = data
        st.session_state["riwayat_local"] = local


def clear_riwayat(spreadsheet):
    """Hapus semua baris riwayat (kecuali header)."""
    if spreadsheet is None:
        st.session_state["riwayat_local"] = {}
        return
    try:
        ws = spreadsheet.worksheet(SHEET_RIWAYAT)
        ws.clear()
        ws.append_row(RIWAYAT_HEADERS)
    except Exception as e:
        st.warning(f"⚠️ Gagal menghapus riwayat: {e}")


# ── DATA LAPORAN / REFERENSI ────────────────────

REFERENSI_HEADERS = ["nama_laporan", "keterangan", "jumlah"]

def load_data_laporan(spreadsheet) -> dict:
    """
    Baca sheet data_laporan → dict {nama_laporan: {keterangan: jumlah}}.
    """
    if spreadsheet is None:
        return st.session_state.get("data_laporan_local", {})
    try:
        ws = get_or_create_sheet(spreadsheet, SHEET_REFERENSI, REFERENSI_HEADERS)
        records = ws.get_all_records()
        result = {}
        for row in records:
            lap  = str(row.get("nama_laporan", "")).strip()
            ket  = str(row.get("keterangan", "")).strip()
            jml  = row.get("jumlah", 0)
            if not lap:
                continue
            result.setdefault(lap, {})
            if ket:
                try:
                    result[lap][ket] = int(jml)
                except (ValueError, TypeError):
                    result[lap][ket] = 0
        return result
    except Exception as e:
        st.warning(f"⚠️ Gagal membaca referensi dari Sheets: {e}")
        return st.session_state.get("data_laporan_local", {})


def save_referensi_row(spreadsheet, nama_laporan: str, keterangan: str, jumlah: int):
    """Tambah satu baris ke sheet referensi."""
    if spreadsheet is None:
        local = st.session_state.get("data_laporan_local", {})
        local.setdefault(nama_laporan, {})[keterangan] = jumlah
        st.session_state["data_laporan_local"] = local
        return
    try:
        ws = get_or_create_sheet(spreadsheet, SHEET_REFERENSI, REFERENSI_HEADERS)
        ws.append_row([nama_laporan, keterangan, jumlah], value_input_option="USER_ENTERED")
    except Exception as e:
        st.warning(f"⚠️ Gagal menyimpan referensi: {e}")


def delete_referensi_row(spreadsheet, nama_laporan: str, keterangan: str):
    """Hapus baris dengan nama_laporan+keterangan tertentu."""
    if spreadsheet is None:
        local = st.session_state.get("data_laporan_local", {})
        if nama_laporan in local and keterangan in local[nama_laporan]:
            del local[nama_laporan][keterangan]
        st.session_state["data_laporan_local"] = local
        return
    try:
        ws = get_or_create_sheet(spreadsheet, SHEET_REFERENSI, REFERENSI_HEADERS)
        all_vals = ws.get_all_values()
        for i, row in enumerate(all_vals[1:], start=2):   # skip header
            if len(row) >= 2 and row[0] == nama_laporan and row[1] == keterangan:
                ws.delete_rows(i)
                break
    except Exception as e:
        st.warning(f"⚠️ Gagal menghapus baris referensi: {e}")


def add_laporan_baru(spreadsheet, nama_laporan: str):
    """Tambah laporan baru (baris placeholder tanpa keterangan)."""
    if spreadsheet is None:
        local = st.session_state.get("data_laporan_local", {})
        local.setdefault(nama_laporan, {})
        st.session_state["data_laporan_local"] = local
        return
    try:
        ws = get_or_create_sheet(spreadsheet, SHEET_REFERENSI, REFERENSI_HEADERS)
        ws.append_row([nama_laporan, "", ""], value_input_option="USER_ENTERED")
    except Exception as e:
        st.warning(f"⚠️ Gagal menambah laporan baru: {e}")


# ══════════════════════════════════════════════
# HELPER: EKSTRAKSI TEKS
# ══════════════════════════════════════════════

def extract_text_pdf(file) -> list:
    try:
        reader = PyPDF2.PdfReader(file)
        return [page.extract_text() or "" for page in reader.pages]
    except Exception as e:
        st.warning(f"⚠️ Gagal membaca PDF '{file.name}': {e}")
        return []

def extract_text_docx(file) -> list:
    try:
        doc = Document(file)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        pages, chunk = [], []
        for i, p in enumerate(paragraphs):
            chunk.append(p)
            if (i + 1) % 30 == 0:
                pages.append("\n".join(chunk))
                chunk = []
        if chunk:
            pages.append("\n".join(chunk))
        return pages
    except Exception as e:
        st.warning(f"⚠️ Gagal membaca DOCX '{file.name}': {e}")
        return []

def pages_to_text(pages: list, max_chars: int = MAX_CHARS) -> str:
    parts, total = [], 0
    for i, page in enumerate(pages, 1):
        chunk = f"\n--- Halaman {i} ---\n{page}"
        if total + len(chunk) > max_chars:
            parts.append("\n\n[... konten dipotong karena terlalu panjang ...]")
            break
        parts.append(chunk)
        total += len(chunk)
    return "".join(parts)


# ══════════════════════════════════════════════
# HELPER: CLAUDE API
# ══════════════════════════════════════════════

def recover_partial_json(raw_text: str):
    """
    Coba selamatkan JSON yang terpotong karena token limit.
    Menggunakan 3 strategi bertingkat.
    """
    import re as _re
    start = raw_text.find("{")
    if start == -1:
        return None
    partial = raw_text[start:]

    def close_json(text):
        cleaned = _re.sub(r',\s*"[^"]*"\s*:?\s*$', "", text.rstrip())
        cleaned = _re.sub(r',\s*"[^"]*"\s*$', "", cleaned.rstrip())
        stack, in_string, esc = [], False, False
        for ch in cleaned:
            if esc: esc = False; continue
            if ch == "\\" and in_string: esc = True; continue
            if ch == '"': in_string = not in_string; continue
            if in_string: continue
            if ch in "{[": stack.append(ch)
            elif ch in "}]":
                if stack: stack.pop()
        closing = "".join("]" if b == "[" else "}" for b in reversed(stack))
        return cleaned + closing

    # Strategi 1: close bracket
    try:
        parsed = json.loads(close_json(partial))
        parsed.setdefault("summary", {})
        parsed.setdefault("findings", [])
        parsed.setdefault("properties", [])
        parsed["_partial"] = True
        return parsed
    except Exception:
        pass

    # Strategi 2: buang findings, selamatkan summary
    fm = _re.search(r'"findings"\s*:\s*\[', partial)
    if fm:
        before = partial[:fm.start()]
        se = before.rfind("}")
        if se > 0:
            try:
                parsed = json.loads(before[:se+1] + ', "findings": []}')
                parsed["_partial"] = True
                return parsed
            except Exception:
                pass

    # Strategi 3: ekstrak field by field
    result = {"_partial": True, "findings": [], "properties": [], "summary": {}}
    m = _re.search(r'"report_type"\s*:\s*"([^"]*)"', partial)
    if m: result["report_type"] = m.group(1)
    pm = _re.search(r'"properties"\s*:\s*\[(.*?)\]', partial, _re.DOTALL)
    if pm: result["properties"] = _re.findall(r'"([^"]+)"', pm.group(1))
    for field in ["total_findings", "kritikal", "minor", "ok", "info", "overall_score"]:
        nm = _re.search('"'  + field + '"\\s*:\\s*(\\d+)', partial)
        if nm: result["summary"][field] = int(nm.group(1))
    em = _re.search(r'"executive_summary"\s*:\s*"([^"]*)"', partial)
    if em: result["summary"]["executive_summary"] = em.group(1)
    for fr in _re.findall(r'\{\s*"id"\s*:.*?"property"\s*:\s*"[^"]*"\s*\}', partial, _re.DOTALL):
        try: result["findings"].append(json.loads(fr))
        except Exception: pass
    if result.get("report_type") or result["summary"]:
        return result
    return None

def call_claude(api_key: str, mode_instruction: str, check_items: list, doc_text: str):
    client = anthropic.Anthropic(api_key=api_key)
    user_message = (
        f"{mode_instruction}\n\n"
        f"Item yang harus diperiksa:\n"
        + "\n".join(f"- {item}" for item in check_items)
        + f"\n\nKONTEN DOKUMEN:\n{doc_text}"
    )
    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )
    raw_text = response.content[0].text

    def try_parse(text):
        try:
            return json.loads(text)
        except (json.JSONDecodeError, ValueError):
            return None

    def fix_json(text):
        text = re.sub(r",\s*([}\]])", r"\1", text)
        text = re.sub(r"//[^\n]*", "", text)
        return text

    parsed = try_parse(raw_text)
    if parsed is None:
        stripped = re.sub(r"```(?:json)?\s*([\s\S]*?)```", r"\1", raw_text).strip()
        parsed = try_parse(stripped)
    if parsed is None:
        for m in re.finditer(r"\{[\s\S]*\}", raw_text):
            parsed = try_parse(m.group())
            if parsed:
                break
    if parsed is None:
        fixed = fix_json(raw_text)
        parsed = try_parse(fixed)
        if parsed is None:
            for m in re.finditer(r"\{[\s\S]*\}", fixed):
                parsed = try_parse(fix_json(m.group()))
                if parsed:
                    break
    # Lapis 5: recovery partial JSON (response terpotong karena token limit)
    if parsed is None:
        parsed = recover_partial_json(raw_text)

    if parsed is None:
        raise ValueError(
            f"Tidak bisa mem-parse JSON dari response Claude.\n"
            f"Raw (500 char pertama):\n{raw_text[:500]}"
        )
    return parsed, raw_text


# ══════════════════════════════════════════════
# HELPER: RENDER UI
# ══════════════════════════════════════════════

def render_finding_card(f: dict):
    sev   = f.get("severity", "info")
    cfg   = SEVERITY_CONFIG.get(sev, SEVERITY_CONFIG["info"])
    color = cfg["color"]
    bg    = cfg["bg"]
    cat   = f.get("category", "")
    prop  = f.get("property", "")
    title = f.get("title", "")
    detail= f.get("detail", "")
    hint  = f.get("page_hint", "")
    fid   = f.get("id", "")
    prop_tag = f' &nbsp;·&nbsp; <span style="color:#1e6fbf;">📌 {prop}</span>' if prop else ""
    st.markdown(f"""
<div style="background:{bg};border:1px solid {color}40;border-left:4px solid {color};
            border-radius:8px;padding:14px 16px;margin-bottom:10px;">
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
        <span style="background:{color}22;color:{color};border:1px solid {color};
                     font-size:10px;font-weight:700;padding:2px 8px;border-radius:4px;
                     text-transform:uppercase;letter-spacing:.5px;">{cfg["emoji"]} {sev}</span>
        <span style="color:#666;font-size:11px;font-family:monospace;">{cat}{prop_tag}</span>
        <span style="color:#666;font-size:11px;font-family:monospace;margin-left:auto;">{fid} &nbsp; {hint}</span>
    </div>
    <div style="font-size:14px;font-weight:600;color:#1a1a2e;margin-bottom:6px;">{title}</div>
    <div style="font-size:12px;color:#444;font-family:monospace;background:#ffffff;
                padding:8px 12px;border-radius:6px;line-height:1.6;
                border:1px solid #e0e0e0;">{detail}</div>
</div>""", unsafe_allow_html=True)


def render_summary_cards(s: dict):
    score = s.get("overall_score", 0)
    sc    = "#1a9e67" if score >= 80 else "#d4860a" if score >= 60 else "#c0392b"
    c1, c2, c3, c4 = st.columns(4)
    for col, num, label, color, bg in [
        (c1, score,              "Skor QC",    sc,        "#f8fafb"),
        (c2, s.get("kritikal",0),"🔴 Kritikal","#c0392b", "#fff0f0"),
        (c3, s.get("minor",0),  "🟡 Minor",   "#d4860a", "#fff8e6"),
        (c4, s.get("ok",0),     "🟢 Sesuai",  "#1a9e67", "#edfaf4"),
    ]:
        col.markdown(f"""
<div style="background:{bg};border:1px solid #dde3ea;border-radius:10px;padding:16px;
            text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.06);">
    <div style="font-size:32px;font-weight:800;color:{color};font-family:monospace;">{num}</div>
    <div style="font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:1px;">{label}</div>
</div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="CekLaporan v6 — KJPP SRR",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Sora', sans-serif; }
    .stApp { background: #f4f6f9; color: #1a1a2e; }
    section[data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #dde3ea; }
    section[data-testid="stSidebar"] * { color: #1a1a2e !important; }
    .block-container { padding-top: 2rem; max-width: 1100px; }
    div[data-testid="stFileUploader"] { border: 2px dashed #c9d2dc; border-radius: 10px; padding: 10px; background: #fff; }
    .stButton > button {
        background: #1a9e67; color: #fff; font-weight: 800;
        border: none; border-radius: 8px; padding: 10px 24px;
        font-family: 'Sora', sans-serif; transition: all .2s;
    }
    .stButton > button:hover { background: #147a50; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(26,158,103,0.3); }
    .stButton > button:disabled { background: #e5e7eb; color: #9ca3af; }
    .stTextInput input { background: #fff; color: #1a1a2e; border: 1px solid #dde3ea; border-radius: 6px; font-family: 'DM Mono', monospace; font-size: 13px; }
    .stTextInput input:focus { border-color: #1a9e67 !important; box-shadow: 0 0 0 2px rgba(26,158,103,0.15) !important; }
    .stSelectbox > div > div { background: #fff; color: #1a1a2e; border-color: #dde3ea; }
    .stCheckbox label { font-size: 13px; color: #374151 !important; }
    hr { border-color: #dde3ea; }
    h1,h2,h3 { color: #1a1a2e !important; font-family: 'Sora', sans-serif !important; }
    .stTabs [data-baseweb="tab-list"] { background: #fff; border-bottom: 2px solid #e5e7eb; border-radius: 8px 8px 0 0; }
    .stTabs [data-baseweb="tab"] { color: #6b7280; font-size: 13px; font-weight: 600; }
    .stTabs [aria-selected="true"] { color: #1a9e67 !important; border-bottom-color: #1a9e67 !important; }
    div[data-testid="stExpander"] { background: #fff; border: 1px solid #dde3ea; border-radius: 8px; }
    .stAlert { border-radius: 8px; }
    [data-testid="stRadio"] label { color: #374151 !important; }
    [data-testid="stCaption"] { color: #6b7280 !important; }
    .stMarkdown p { color: #374151; }
</style>""", unsafe_allow_html=True)

    # ── HEADER ──
    st.markdown("""
<div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;">
    <div style="background:#1a9e67;border-radius:10px;width:40px;height:40px;
                display:flex;align-items:center;justify-content:center;font-size:20px;">📋</div>
    <div>
        <h1 style="margin:0;font-size:24px;font-weight:800;letter-spacing:-0.5px;color:#1a1a2e;">
            Cek<span style="color:#1a9e67;">Laporan</span>
            <span style="font-size:12px;background:#edfaf4;color:#1a9e67;border:1px solid #1a9e67;
                         padding:2px 10px;border-radius:20px;font-weight:600;margin-left:8px;">v6.0 — AI Powered</span>
        </h1>
        <p style="margin:0;color:#6b7280;font-size:12px;font-family:'DM Mono',monospace;">
            KJPP Suwendho Rinaldy dan Rekan · Pengecekan Laporan Penilaian</p>
    </div>
</div>
<hr style="border-color:#dde3ea;">""", unsafe_allow_html=True)

    # ── KONEKSI GOOGLE SHEETS ──
    _, spreadsheet, gs_error = get_gsheet_client()
    gs_connected = spreadsheet is not None

    # ── LOAD DATA ──
    data_laporan = load_data_laporan(spreadsheet)
    riwayat      = load_riwayat(spreadsheet)

    # ════════════════════════════════
    # SIDEBAR
    # ════════════════════════════════
    with st.sidebar:
        # Status Google Sheets
        if gs_connected:
            st.markdown("""
<div style="background:#edfaf4;border:1px solid #b2e8d0;border-radius:8px;
            padding:10px 12px;margin-bottom:12px;font-size:12px;">
    🟢 <strong>Google Sheets</strong> terhubung<br>
    <span style="color:#6b7280;font-size:11px;">Data tersimpan permanen</span>
</div>""", unsafe_allow_html=True)
        else:
            error_detail = gs_error or "Secrets belum diisi"
            st.markdown(f"""
<div style="background:#fff8e6;border:1px solid #f5e0a0;border-radius:8px;
            padding:10px 12px;margin-bottom:12px;font-size:12px;">
    🟡 <strong>Google Sheets</strong> tidak terkonfigurasi<br>
    <span style="color:#6b7280;font-size:11px;">Data hanya tersimpan sementara (session)</span>
</div>""", unsafe_allow_html=True)
            with st.expander("🔍 Lihat detail error", expanded=True):
                st.code(error_detail, language="text")

        st.markdown("### 🔑 Claude API Key")
        api_key = st.text_input(
            "Masukkan API Key",
            type="password",
            placeholder="sk-ant-api03-...",
            help="Dapatkan API Key di https://console.anthropic.com",
        )
        if api_key:
            if api_key.startswith("sk-ant"):
                st.success("✅ Format key valid")
            else:
                st.error("❌ Format tidak valid (harus diawali sk-ant)")

        st.markdown("---")
        st.markdown("### ⚡ Mode Pengecekan")
        mode_label = st.radio("Pilih mode", options=list(MODE_CONFIG.keys()), label_visibility="collapsed")
        st.caption(MODE_CONFIG[mode_label]["desc"])

        st.markdown("---")
        st.markdown("### ✅ Item yang Dicek")
        selected_items = []
        for item in CHECK_ITEMS_DEFAULT:
            default_checked = item not in ["Analisis Pasar & Data Pembanding", "Pendekatan & Metode Penilaian"]
            if st.checkbox(item, value=default_checked, key=f"chk_{item}"):
                selected_items.append(item)

        st.markdown("---")
        st.markdown("### 📚 Referensi Laporan")
        laporan_names = list(data_laporan.keys())
        selected_ref  = st.selectbox("Laporan Referensi", ["(Tidak ada)"] + laporan_names + ["+ Tambah Baru"])

        if selected_ref not in ["(Tidak ada)", "+ Tambah Baru"] and selected_ref in data_laporan:
            with st.expander("Lihat referensi"):
                ref_data = data_laporan[selected_ref]
                if ref_data:
                    for k, v in ref_data.items():
                        st.write(f"• {k}: **{v}x**")
                else:
                    st.caption("Belum ada data referensi.")

        if selected_ref == "+ Tambah Baru":
            new_name = st.text_input("Nama laporan baru:")
            if st.button("Tambahkan") and new_name:
                if new_name not in data_laporan:
                    add_laporan_baru(spreadsheet, new_name)
                    st.success(f"✅ '{new_name}' ditambahkan")
                    st.rerun()
                else:
                    st.warning("Nama sudah ada.")

    # ════════════════════════════════
    # TABS
    # ════════════════════════════════
    tab_audit, tab_search, tab_history, tab_ref = st.tabs([
        "🤖 AI Audit", "🔍 Pencarian Teks", "📜 Riwayat Audit", "📁 Kelola Referensi"
    ])

    # ────────────────────────────────
    # TAB 1: AI AUDIT
    # ────────────────────────────────
    with tab_audit:
        st.markdown("#### 📁 Upload Laporan")
        uploaded_pdfs  = st.file_uploader("File PDF",  type="pdf",  accept_multiple_files=True, key="pdf_audit")
        uploaded_docxs = st.file_uploader("File DOCX", type="docx", accept_multiple_files=True, key="docx_audit")
        all_files = (uploaded_pdfs or []) + (uploaded_docxs or [])

        if all_files:
            st.markdown(f"**{len(all_files)} file siap dianalisis:**")
            for f in all_files:
                icon = "📄" if f.name.endswith(".pdf") else "📝"
                st.markdown(
                    f'<span style="font-family:monospace;font-size:12px;color:#6b7280;">'
                    f'{icon} {f.name} &nbsp;·&nbsp; {f.size//1024} KB</span>',
                    unsafe_allow_html=True
                )

        st.markdown("---")
        col_run, col_info = st.columns([2, 5])
        with col_run:
            run_disabled = not (api_key and api_key.startswith("sk-ant") and all_files and selected_items)
            run_btn = st.button("▶ Jalankan Analisis", disabled=run_disabled, use_container_width=True)
        with col_info:
            if not api_key:
                st.info("💡 Masukkan API Key di sidebar untuk memulai.")
            elif not all_files:
                st.info("💡 Upload minimal satu file laporan.")
            elif not selected_items:
                st.info("💡 Pilih minimal satu item pengecekan.")
            else:
                st.success(f"✅ Siap: {len(all_files)} file · {len(selected_items)} item · mode **{mode_label}**")

        if run_btn:
            st.markdown("---")

            with st.status("📖 Membaca dokumen...", expanded=True) as status:
                all_pages, file_info = [], []
                for f in all_files:
                    st.write(f"Membaca **{f.name}**...")
                    pages = extract_text_pdf(f) if f.name.endswith(".pdf") else extract_text_docx(f)
                    all_pages.extend(pages)
                    file_info.append(f"{f.name} ({len(pages)} hal.)")
                doc_text = pages_to_text(all_pages)
                status.update(label=f"✅ {len(all_pages)} halaman dibaca dari {len(all_files)} file", state="complete")

            with st.status("🧠 Claude sedang menganalisis laporan...", expanded=True) as status:
                st.write(f"Model: `{MODEL}` · Mode: **{mode_label}**")
                st.write(f"Teks dikirim: **{len(doc_text):,} karakter**")
                t_start = time.time()
                try:
                    result, raw_text = call_claude(
                        api_key, MODE_CONFIG[mode_label]["instruction"], selected_items, doc_text
                    )
                    elapsed = time.time() - t_start
                    status.update(label=f"✅ Analisis selesai dalam {elapsed:.1f} detik", state="complete")
                except Exception as e:
                    status.update(label="❌ Error", state="error")
                    st.error(f"**Error:** {e}")
                    st.stop()

            # ── SIMPAN RIWAYAT ──
            audit_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_riwayat_row(spreadsheet, audit_id, {
                "timestamp": datetime.now().strftime("%d %b %Y, %H:%M"),
                "files":     [f.name for f in all_files],
                "mode":      mode_label,
                "result":    result,
            })
            if gs_connected:
                st.toast("💾 Riwayat tersimpan ke Google Sheets", icon="✅")

            # ── TAMPILKAN HASIL ──
            st.markdown("---")
            st.markdown(f"### 📊 Hasil Audit — `{result.get('report_type','').upper()}`")
            st.caption(f"{' · '.join(file_info)} · {datetime.now().strftime('%d %b %Y, %H:%M')} · Mode: {mode_label}")

            # Peringatan jika response terpotong
            if result.get("_partial"):
                st.warning(
                    "⚠️ **Response Claude terpotong** karena laporan terlalu panjang. "
                    "Temuan yang berhasil dibaca tetap ditampilkan, namun mungkin tidak lengkap. "
                    "Coba gunakan mode **Pre-Check** untuk laporan besar, atau kurangi jumlah halaman."
                )

            render_summary_cards(result.get("summary", {}))
            st.markdown("<br>", unsafe_allow_html=True)

            exec_sum = result.get("summary", {}).get("executive_summary", "")
            if exec_sum:
                st.markdown(f"""
<div style="background:#fff;border:1px solid #dde3ea;border-radius:10px;padding:16px;
            margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,0.05);">
    <div style="font-size:10px;color:#6b7280;font-family:monospace;text-transform:uppercase;
                letter-spacing:1.5px;margin-bottom:8px;">RINGKASAN EKSEKUTIF</div>
    <p style="font-size:14px;line-height:1.8;color:#374151;margin:0;">{exec_sum}</p>
</div>""", unsafe_allow_html=True)

            properties = result.get("properties", [])
            if len(properties) > 1:
                st.markdown(f"""
<div style="background:#eef4ff;border:1px solid #bfcfee;border-radius:10px;padding:14px;margin-bottom:16px;">
    <div style="font-size:10px;color:#1e6fbf;font-family:monospace;text-transform:uppercase;
                letter-spacing:1.5px;margin-bottom:8px;">🏢 OBJEK PROPERTI TERDETEKSI ({len(properties)})</div>
    <div style="display:flex;flex-wrap:wrap;gap:6px;">
        {"".join(f'<span style="background:#fff;color:#1e6fbf;border:1px solid #bfcfee;font-size:11px;font-family:monospace;padding:3px 10px;border-radius:12px;">{p}</span>' for p in properties)}
    </div>
</div>""", unsafe_allow_html=True)

            findings = result.get("findings", [])
            if findings:
                filter_prop = None
                if len(properties) > 1:
                    sel = st.selectbox("Filter per objek:", ["Semua Objek"] + properties, key="prop_filter")
                    if sel != "Semua Objek":
                        filter_prop = sel
                grouped = {}
                for f in findings:
                    if filter_prop and f.get("property") and f["property"] != filter_prop:
                        continue
                    grouped.setdefault(f.get("severity", "info"), []).append(f)
                for sev in ["kritikal", "minor", "ok", "info"]:
                    group = grouped.get(sev, [])
                    if not group:
                        continue
                    cfg = SEVERITY_CONFIG[sev]
                    st.markdown(
                        f'<div style="font-size:11px;font-family:monospace;color:#6b7280;'
                        f'text-transform:uppercase;letter-spacing:1.5px;margin:16px 0 8px;">'
                        f'{cfg["emoji"]} {sev.upper()} ({len(group)})'
                        f'<span style="display:inline-block;height:1px;background:#dde3ea;'
                        f'width:200px;margin-left:10px;vertical-align:middle;"></span></div>',
                        unsafe_allow_html=True
                    )
                    for f in group:
                        render_finding_card(f)
            else:
                st.success("✅ Tidak ada temuan — laporan terlihat konsisten.")

            with st.expander("🔧 Raw JSON Output (untuk debugging)"):
                st.code(raw_text, language="json")

    # ────────────────────────────────
    # TAB 2: PENCARIAN TEKS
    # ────────────────────────────────
    with tab_search:
        st.markdown("#### 🔍 Pencarian Frasa Manual")
        st.caption("Upload dokumen dan cari frasa — dilengkapi highlight.")
        col_s1, col_s2 = st.columns([1, 2])
        with col_s1:
            up_pdf_s  = st.file_uploader("PDF",  type="pdf",  accept_multiple_files=True, key="pdf_search")
            up_docx_s = st.file_uploader("DOCX", type="docx", accept_multiple_files=True, key="docx_search")
            files_s = (up_pdf_s or []) + (up_docx_s or [])
            sel_laporan_s = st.selectbox("Laporan Referensi", ["(Tidak ada)"] + list(data_laporan.keys()), key="_sel_laporan_s2")
            if sel_laporan_s != "(Tidak ada)" and sel_laporan_s in data_laporan:
                st.markdown("**Wajib cek:**")
                for k, v in data_laporan[sel_laporan_s].items():
                    st.write(f"• {k}: {v}x")
        with col_s2:
            phrase = st.text_input("Frasa yang dicari:", placeholder="contoh: tanggal penilaian")
            if phrase and files_s:
                grand_total = 0
                for uf in files_s:
                    pages = extract_text_pdf(uf) if uf.name.endswith(".pdf") else extract_text_docx(uf)
                    pattern = re.compile(r"\s*".join(re.escape(w) for w in phrase.split()), re.IGNORECASE)
                    file_total = 0
                    st.markdown(f"**📄 {uf.name}**")
                    found_any = False
                    for i, page_txt in enumerate(pages, 1):
                        matches = pattern.findall(page_txt)
                        if matches:
                            found_any = True
                            file_total += len(matches)
                            highlighted = re.sub(
                                fr"({re.escape(phrase)})",
                                r'<mark style="background:#fbbf24;color:#000;padding:0 2px;">\1</mark>',
                                page_txt, flags=re.IGNORECASE
                            )
                            with st.expander(f"Halaman {i} — {len(matches)} kemunculan"):
                                st.markdown(highlighted, unsafe_allow_html=True)
                    if not found_any:
                        st.caption("Tidak ditemukan.")
                    else:
                        st.markdown(f"→ **Total di file ini: {file_total}x**")
                    grand_total += file_total
                st.markdown(f"---\n### Total semua file: **{grand_total}x**")
            elif phrase and not files_s:
                st.info("Upload file terlebih dahulu.")

    # ────────────────────────────────
    # TAB 3: RIWAYAT AUDIT
    # ────────────────────────────────
    with tab_history:
        st.markdown("#### 📜 Riwayat Audit")

        # Badge status storage
        if gs_connected:
            st.markdown('<span style="background:#edfaf4;color:#1a9e67;border:1px solid #1a9e67;font-size:11px;font-family:monospace;padding:2px 10px;border-radius:12px;">💾 Tersimpan di Google Sheets</span>', unsafe_allow_html=True)
        else:
            st.markdown('<span style="background:#fff8e6;color:#d4860a;border:1px solid #d4860a;font-size:11px;font-family:monospace;padding:2px 10px;border-radius:12px;">⚠️ Hanya di session (sementara)</span>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Reload fresh dari sheets
        riwayat = load_riwayat(spreadsheet)

        if not riwayat:
            st.info("Belum ada riwayat audit. Jalankan analisis AI terlebih dahulu.")
        else:
            st.caption(f"Total: **{len(riwayat)}** audit tersimpan")
            for audit_id in sorted(riwayat.keys(), reverse=True):
                r = riwayat[audit_id]
                score = r.get("score", 0)
                files_str = ", ".join(r.get("files", []))
                with st.expander(f"🕒 {r.get('timestamp','')}  ·  {files_str}  ·  Skor: {score}"):
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Skor QC",    score)
                    c2.metric("🔴 Kritikal", r.get("kritikal", 0))
                    c3.metric("🟡 Minor",    r.get("minor", 0))
                    c4.metric("Mode",        r.get("mode", ""))
                    exec_s = r.get("result", {}).get("summary", {}).get("executive_summary", "")
                    if exec_s:
                        st.caption(exec_s)
                    if st.button("Tampilkan Detail Temuan", key=f"hist_{audit_id}"):
                        for f in r.get("result", {}).get("findings", []):
                            render_finding_card(f)

            st.markdown("---")
            if st.button("🗑 Hapus Semua Riwayat"):
                clear_riwayat(spreadsheet)
                st.success("Riwayat dihapus.")
                st.rerun()

    # ────────────────────────────────
    # TAB 4: KELOLA REFERENSI
    # ────────────────────────────────
    with tab_ref:
        st.markdown("#### 📁 Kelola Data Referensi Laporan")
        st.caption("Simpan catatan frekuensi kemunculan frasa per jenis laporan sebagai baseline.")

        if gs_connected:
            st.markdown('<span style="background:#edfaf4;color:#1a9e67;border:1px solid #1a9e67;font-size:11px;font-family:monospace;padding:2px 10px;border-radius:12px;">💾 Tersimpan di Google Sheets</span>', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

        data_laporan = load_data_laporan(spreadsheet)

        if not data_laporan:
            st.info("Belum ada data referensi.")
        else:
            for lap_name, lap_data in data_laporan.items():
                with st.expander(f"📋 {lap_name}"):
                    if lap_data:
                        for k, v in lap_data.items():
                            c1, c2, c3 = st.columns([3, 1, 1])
                            c1.write(k)
                            c2.write(f"**{v}x**")
                            if c3.button("🗑", key=f"del_{lap_name}_{k}"):
                                delete_referensi_row(spreadsheet, lap_name, k)
                                st.rerun()
                    else:
                        st.caption("Belum ada keterangan.")
                    st.markdown("**Tambah keterangan:**")
                    ck, cv, cbtn = st.columns([3, 1, 1])
                    new_k = ck.text_input("Frasa",  key=f"nk_{lap_name}")
                    new_v = cv.number_input("Jumlah", min_value=0, step=1, key=f"nv_{lap_name}")
                    if cbtn.button("Tambah", key=f"nbtn_{lap_name}") and new_k:
                        save_referensi_row(spreadsheet, lap_name, new_k, new_v)
                        st.success("✅ Ditambahkan")
                        st.rerun()

        st.markdown("---")
        st.markdown("**Tambah Laporan Referensi Baru:**")
        c1, c2 = st.columns([3, 1])
        new_lap = c1.text_input("Nama laporan:", key="new_lap_name")
        if c2.button("Tambahkan", key="btn_new_lap") and new_lap:
            if new_lap not in data_laporan:
                add_laporan_baru(spreadsheet, new_lap)
                st.success(f"✅ '{new_lap}' ditambahkan")
                st.rerun()
            else:
                st.warning("Nama sudah ada.")

    # ── FOOTER ──
    st.markdown("""
<hr style="margin-top:40px;border-color:#dde3ea;">
<p style="text-align:center;font-size:12px;color:#9ca3af;font-family:'DM Mono',monospace;">
    CekLaporan v6.0 &nbsp;·&nbsp; KJPP SRR &nbsp;·&nbsp;
    Powered by Claude AI (claude-sonnet-4-5) &nbsp;·&nbsp; Created by HW
</p>""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
