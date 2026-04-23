import os
import re
import json
import time
import PyPDF2
import anthropic
import streamlit as st
from docx import Document
from datetime import datetime

# ──────────────────────────────────────────────
# KONSTANTA
# ──────────────────────────────────────────────
DATA_FILE   = "data_laporan.json"
HISTORY_FILE = "riwayat_audit.json"

MODEL       = "claude-sonnet-4-5"
MAX_TOKENS  = 4096
MAX_CHARS   = 60000   # batas karakter teks yang dikirim ke Claude

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
    "kritikal": {"emoji": "🔴", "color": "#e05c5c", "bg": "#3a1a1a"},
    "minor":    {"emoji": "🟡", "color": "#f0a500", "bg": "#3a2e00"},
    "ok":       {"emoji": "🟢", "color": "#2ecc8a", "bg": "#0a2e1e"},
    "info":     {"emoji": "🔵", "color": "#58a6ff", "bg": "#0d1f3a"},
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


# ──────────────────────────────────────────────
# HELPER: DATA
# ──────────────────────────────────────────────
def load_json(path: str) -> dict:
    if os.path.exists(path):
        try:
            return json.load(open(path, "r", encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_json(path: str, data: dict):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


# ──────────────────────────────────────────────
# HELPER: EKSTRAKSI TEKS
# ──────────────────────────────────────────────
def extract_text_pdf(file) -> list[str]:
    """Kembalikan list teks per halaman."""
    try:
        reader = PyPDF2.PdfReader(file)
        pages = []
        for page in reader.pages:
            txt = page.extract_text() or ""
            pages.append(txt)
        return pages
    except Exception as e:
        st.warning(f"⚠️ Gagal membaca PDF '{file.name}': {e}")
        return []

def extract_text_docx(file) -> list[str]:
    """Kembalikan list teks per paragraf (simulasi halaman)."""
    try:
        doc = Document(file)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        # Kelompokkan ~30 paragraf per "halaman" agar ada page_hint
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

def pages_to_text(pages: list[str], max_chars: int = MAX_CHARS) -> str:
    """Gabungkan halaman dengan penanda nomor halaman, potong jika melebihi max_chars."""
    parts = []
    total = 0
    for i, page in enumerate(pages, 1):
        chunk = f"\n--- Halaman {i} ---\n{page}"
        if total + len(chunk) > max_chars:
            parts.append("\n\n[... konten dipotong karena terlalu panjang ...]")
            break
        parts.append(chunk)
        total += len(chunk)
    return "".join(parts)


# ──────────────────────────────────────────────
# HELPER: CLAUDE API
# ──────────────────────────────────────────────
def call_claude(api_key: str, mode_instruction: str, check_items: list[str], doc_text: str) -> dict:
    """
    Kirim teks dokumen ke Claude dan kembalikan dict hasil parsing JSON.
    Raises exception jika API error.
    """
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

    # Parse JSON
    try:
        # Coba langsung
        return json.loads(raw_text), raw_text
    except json.JSONDecodeError:
        # Cari blok JSON di dalam teks
        m = re.search(r"\{[\s\S]*\}", raw_text)
        if m:
            return json.loads(m.group()), raw_text
        raise ValueError(f"Claude tidak mengembalikan JSON valid.\n\nRaw:\n{raw_text}")


# ──────────────────────────────────────────────
# HELPER: RENDER FINDINGS
# ──────────────────────────────────────────────
def render_finding_card(f: dict):
    sev  = f.get("severity", "info")
    cfg  = SEVERITY_CONFIG.get(sev, SEVERITY_CONFIG["info"])
    emoji = cfg["emoji"]
    color = cfg["color"]
    bg    = cfg["bg"]

    cat    = f.get("category", "")
    prop   = f.get("property", "")
    title  = f.get("title", "")
    detail = f.get("detail", "")
    hint   = f.get("page_hint", "")
    fid    = f.get("id", "")

    prop_tag = f' &nbsp;·&nbsp; <span style="color:#58a6ff;">📌 {prop}</span>' if prop else ""
    hint_tag = f'<span style="color:#7d8590;font-size:11px;float:right;">{hint}</span>' if hint else ""

    st.markdown(f"""
<div style="
    background:{bg};
    border:1px solid {color}40;
    border-left:4px solid {color};
    border-radius:8px;
    padding:14px 16px;
    margin-bottom:10px;
">
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
        <span style="
            background:{color}22;
            color:{color};
            border:1px solid {color};
            font-size:10px;
            font-weight:700;
            padding:2px 8px;
            border-radius:4px;
            text-transform:uppercase;
            letter-spacing:.5px;
        ">{emoji} {sev}</span>
        <span style="color:#7d8590;font-size:11px;font-family:monospace;">{cat}{prop_tag}</span>
        <span style="color:#7d8590;font-size:11px;font-family:monospace;margin-left:auto;">{fid} &nbsp; {hint}</span>
    </div>
    <div style="font-size:14px;font-weight:600;color:#e6edf3;margin-bottom:6px;">{title}</div>
    <div style="font-size:12px;color:#b1bac4;font-family:monospace;background:#0d1117;
                padding:8px 12px;border-radius:6px;line-height:1.6;">{detail}</div>
</div>
""", unsafe_allow_html=True)


def render_summary_cards(s: dict):
    score = s.get("overall_score", 0)
    score_color = "#2ecc8a" if score >= 80 else "#f0a500" if score >= 60 else "#e05c5c"

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
<div style="background:#161b22;border:1px solid #30363d;border-radius:10px;padding:16px;text-align:center;">
    <div style="font-size:32px;font-weight:800;color:{score_color};font-family:monospace;">{score}</div>
    <div style="font-size:10px;color:#7d8590;text-transform:uppercase;letter-spacing:1px;">Skor QC</div>
</div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
<div style="background:#161b22;border:1px solid #30363d;border-radius:10px;padding:16px;text-align:center;">
    <div style="font-size:32px;font-weight:800;color:#e05c5c;font-family:monospace;">{s.get('kritikal',0)}</div>
    <div style="font-size:10px;color:#7d8590;text-transform:uppercase;letter-spacing:1px;">🔴 Kritikal</div>
</div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
<div style="background:#161b22;border:1px solid #30363d;border-radius:10px;padding:16px;text-align:center;">
    <div style="font-size:32px;font-weight:800;color:#f0a500;font-family:monospace;">{s.get('minor',0)}</div>
    <div style="font-size:10px;color:#7d8590;text-transform:uppercase;letter-spacing:1px;">🟡 Minor</div>
</div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
<div style="background:#161b22;border:1px solid #30363d;border-radius:10px;padding:16px;text-align:center;">
    <div style="font-size:32px;font-weight:800;color:#2ecc8a;font-family:monospace;">{s.get('ok',0)}</div>
    <div style="font-size:10px;color:#7d8590;text-transform:uppercase;letter-spacing:1px;">🟢 Sesuai</div>
</div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# MAIN APP
# ──────────────────────────────────────────────
def main():
    # ── PAGE CONFIG ──
    st.set_page_config(
        page_title="CekLaporan v6 — KJPP SRR",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── GLOBAL CSS ──
    st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');
    html, body, [class*="css"] { font-family: 'Sora', sans-serif; }
    .stApp { background: #0d1117; color: #e6edf3; }
    section[data-testid="stSidebar"] { background: #161b22; border-right: 1px solid #30363d; }
    section[data-testid="stSidebar"] * { color: #e6edf3 !important; }
    .block-container { padding-top: 2rem; max-width: 1100px; }
    div[data-testid="stFileUploader"] { border: 2px dashed #30363d; border-radius: 10px; padding: 10px; }
    .stButton > button {
        background: #2ecc8a; color: #000; font-weight: 800;
        border: none; border-radius: 8px; padding: 10px 24px;
        font-family: 'Sora', sans-serif; transition: all .2s;
    }
    .stButton > button:hover { background: #1a9e67; transform: translateY(-1px); }
    .stButton > button:disabled { background: #1e2630; color: #7d8590; }
    .stTextInput input, .stTextInput input:focus {
        background: #0d1117; color: #e6edf3;
        border: 1px solid #30363d; border-radius: 6px;
        font-family: 'DM Mono', monospace; font-size: 13px;
    }
    .stTextInput input:focus { border-color: #2ecc8a !important; box-shadow: none !important; }
    .stSelectbox > div > div { background: #161b22; color: #e6edf3; border-color: #30363d; }
    .stCheckbox label { font-size: 13px; color: #b1bac4 !important; }
    hr { border-color: #30363d; }
    h1, h2, h3 { color: #e6edf3 !important; font-family: 'Sora', sans-serif !important; }
    .stTabs [data-baseweb="tab-list"] { background: #161b22; border-bottom: 1px solid #30363d; }
    .stTabs [data-baseweb="tab"] { color: #7d8590; font-size: 13px; font-weight: 600; }
    .stTabs [aria-selected="true"] { color: #2ecc8a !important; border-bottom-color: #2ecc8a !important; }
    .stTabs [data-baseweb="tab-panel"] { background: transparent; }
    div[data-testid="stExpander"] { background: #161b22; border: 1px solid #30363d; border-radius: 8px; }
    .stAlert { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

    # ── HEADER ──
    st.markdown("""
<div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;">
    <div style="background:#2ecc8a;border-radius:10px;width:40px;height:40px;
                display:flex;align-items:center;justify-content:center;font-size:20px;">📋</div>
    <div>
        <h1 style="margin:0;font-size:24px;font-weight:800;letter-spacing:-0.5px;">
            Cek<span style="color:#2ecc8a;">Laporan</span>
            <span style="font-size:12px;background:#0a2e1e;color:#2ecc8a;border:1px solid #2ecc8a;
                         padding:2px 10px;border-radius:20px;font-weight:600;margin-left:8px;">v6.0 — AI Powered</span>
        </h1>
        <p style="margin:0;color:#7d8590;font-size:12px;font-family:'DM Mono',monospace;">
            KJPP Suwendho Rinaldy dan Rekan · Pengecekan Laporan Penilaian</p>
    </div>
</div>
<hr>
""", unsafe_allow_html=True)

    # ── LOAD DATA ──
    data_laporan = load_json(DATA_FILE)
    riwayat      = load_json(HISTORY_FILE)

    # ════════════════════════════════
    # SIDEBAR
    # ════════════════════════════════
    with st.sidebar:
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
        mode_label = st.radio(
            "Pilih mode",
            options=list(MODE_CONFIG.keys()),
            label_visibility="collapsed",
        )
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
        selected_ref  = st.selectbox(
            "Laporan Referensi",
            ["(Tidak ada)"] + laporan_names + ["+ Tambah Baru"],
        )

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
                    data_laporan[new_name] = {}
                    save_json(DATA_FILE, data_laporan)
                    st.success(f"✅ '{new_name}' ditambahkan")
                else:
                    st.warning("Nama sudah ada.")

    # ════════════════════════════════
    # MAIN AREA — TABS
    # ════════════════════════════════
    tab_audit, tab_search, tab_history, tab_ref = st.tabs([
        "🤖 AI Audit",
        "🔍 Pencarian Teks",
        "📜 Riwayat Audit",
        "📁 Kelola Referensi",
    ])

    # ────────────────────────────────
    # TAB 1: AI AUDIT
    # ────────────────────────────────
    with tab_audit:
        st.markdown("#### 📁 Upload Laporan")
        uploaded_pdfs  = st.file_uploader("File PDF", type="pdf",  accept_multiple_files=True, key="pdf_audit")
        uploaded_docxs = st.file_uploader("File DOCX", type="docx", accept_multiple_files=True, key="docx_audit")
        all_files = (uploaded_pdfs or []) + (uploaded_docxs or [])

        if all_files:
            st.markdown(f"**{len(all_files)} file siap dianalisis:**")
            for f in all_files:
                icon = "📄" if f.name.endswith(".pdf") else "📝"
                st.markdown(
                    f'<span style="font-family:monospace;font-size:12px;color:#b1bac4;">'
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

        # ── PROSES ANALISIS ──
        if run_btn:
            st.markdown("---")

            # 1. Ekstraksi teks
            with st.status("📖 Membaca dokumen...", expanded=True) as status:
                all_pages = []
                file_info = []
                for f in all_files:
                    st.write(f"Membaca **{f.name}**...")
                    if f.name.endswith(".pdf"):
                        pages = extract_text_pdf(f)
                    else:
                        pages = extract_text_docx(f)
                    all_pages.extend(pages)
                    file_info.append(f"{f.name} ({len(pages)} hal.)")

                doc_text = pages_to_text(all_pages)
                status.update(label=f"✅ {len(all_pages)} halaman dibaca dari {len(all_files)} file", state="complete")

            # 2. Panggil Claude
            with st.status("🧠 Claude sedang menganalisis laporan...", expanded=True) as status:
                st.write(f"Model: `{MODEL}` · Mode: **{mode_label}**")
                st.write(f"Teks dikirim: **{len(doc_text):,} karakter**")
                st.write(f"Item pengecekan: {len(selected_items)} item")

                t_start = time.time()
                try:
                    result, raw_text = call_claude(
                        api_key,
                        MODE_CONFIG[mode_label]["instruction"],
                        selected_items,
                        doc_text,
                    )
                    elapsed = time.time() - t_start
                    status.update(
                        label=f"✅ Analisis selesai dalam {elapsed:.1f} detik",
                        state="complete"
                    )
                except Exception as e:
                    status.update(label="❌ Error", state="error")
                    st.error(f"**Error:** {e}")
                    st.stop()

            # ── SIMPAN KE RIWAYAT ──
            audit_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            riwayat[audit_id] = {
                "timestamp": datetime.now().strftime("%d %b %Y, %H:%M"),
                "files": [f.name for f in all_files],
                "mode": mode_label,
                "score": result.get("summary", {}).get("overall_score", 0),
                "kritikal": result.get("summary", {}).get("kritikal", 0),
                "minor": result.get("summary", {}).get("minor", 0),
                "result": result,
            }
            save_json(HISTORY_FILE, riwayat)

            # ── TAMPILKAN HASIL ──
            st.markdown("---")
            st.markdown(f"### 📊 Hasil Audit — `{result.get('report_type','').upper()}`")
            st.caption(
                f"{' · '.join(file_info)} &nbsp;·&nbsp; "
                f"{datetime.now().strftime('%d %b %Y, %H:%M')} &nbsp;·&nbsp; "
                f"Mode: {mode_label}"
            )

            # Summary cards
            render_summary_cards(result.get("summary", {}))
            st.markdown("<br>", unsafe_allow_html=True)

            # Executive summary
            exec_sum = result.get("summary", {}).get("executive_summary", "")
            if exec_sum:
                st.markdown(f"""
<div style="background:#161b22;border:1px solid #30363d;border-radius:10px;padding:16px;margin-bottom:16px;">
    <div style="font-size:10px;color:#7d8590;font-family:monospace;text-transform:uppercase;
                letter-spacing:1.5px;margin-bottom:8px;">RINGKASAN EKSEKUTIF</div>
    <p style="font-size:14px;line-height:1.8;color:#b1bac4;margin:0;">{exec_sum}</p>
</div>""", unsafe_allow_html=True)

            # Multi-objek: tampilkan daftar properti
            properties = result.get("properties", [])
            if len(properties) > 1:
                st.markdown(f"""
<div style="background:#0d1f3a;border:1px solid #58a6ff40;border-radius:10px;padding:14px;margin-bottom:16px;">
    <div style="font-size:10px;color:#58a6ff;font-family:monospace;text-transform:uppercase;
                letter-spacing:1.5px;margin-bottom:8px;">🏢 OBJEK PROPERTI TERDETEKSI ({len(properties)})</div>
    <div style="display:flex;flex-wrap:wrap;gap:6px;">
        {"".join(f'<span style="background:#0d1117;color:#58a6ff;border:1px solid #58a6ff;font-size:11px;font-family:monospace;padding:3px 10px;border-radius:12px;">{p}</span>' for p in properties)}
    </div>
</div>""", unsafe_allow_html=True)

            # ── FINDINGS ──
            findings = result.get("findings", [])
            if findings:
                # Filter per properti jika multi-objek
                filter_prop = None
                if len(properties) > 1:
                    prop_options = ["Semua Objek"] + properties
                    sel = st.selectbox("Filter per objek:", prop_options, key="prop_filter")
                    if sel != "Semua Objek":
                        filter_prop = sel

                # Kelompokkan per severity
                grouped: dict[str, list] = {}
                for f in findings:
                    if filter_prop and f.get("property") and f["property"] != filter_prop:
                        continue
                    sev = f.get("severity", "info")
                    grouped.setdefault(sev, []).append(f)

                for sev in ["kritikal", "minor", "ok", "info"]:
                    group = grouped.get(sev, [])
                    if not group:
                        continue
                    cfg = SEVERITY_CONFIG[sev]
                    st.markdown(
                        f'<div style="font-size:11px;font-family:monospace;color:#7d8590;'
                        f'text-transform:uppercase;letter-spacing:1.5px;margin:16px 0 8px;">'
                        f'{cfg["emoji"]} {sev.upper()} ({len(group)})'
                        f'<span style="display:inline-block;height:1px;background:#30363d;'
                        f'width:200px;margin-left:10px;vertical-align:middle;"></span></div>',
                        unsafe_allow_html=True
                    )
                    for f in group:
                        render_finding_card(f)
            else:
                st.success("✅ Tidak ada temuan — laporan terlihat konsisten.")

            # Raw JSON expandable
            with st.expander("🔧 Raw JSON Output (untuk debugging)"):
                st.code(raw_text, language="json")

    # ────────────────────────────────
    # TAB 2: PENCARIAN TEKS (dari Ceklaporan5)
    # ────────────────────────────────
    with tab_search:
        st.markdown("#### 🔍 Pencarian Frasa Manual")
        st.caption("Upload dokumen dan cari frasa — dilengkapi highlight. Fitur asli dari Ceklaporan5.")

        col_s1, col_s2 = st.columns([1, 2])
        with col_s1:
            up_pdf_s  = st.file_uploader("PDF", type="pdf",  accept_multiple_files=True, key="pdf_search")
            up_docx_s = st.file_uploader("DOCX", type="docx", accept_multiple_files=True, key="docx_search")
            files_s = (up_pdf_s or []) + (up_docx_s or [])

            if st.session_state.get("_sel_laporan_s") is None:
                st.session_state["_sel_laporan_s"] = list(data_laporan.keys())[0] if data_laporan else None

            sel_laporan_s = st.selectbox(
                "Laporan Referensi",
                ["(Tidak ada)"] + list(data_laporan.keys()),
                key="_sel_laporan_s2"
            )
            if sel_laporan_s != "(Tidak ada)" and sel_laporan_s in data_laporan:
                st.markdown("**Wajib cek:**")
                for k, v in data_laporan[sel_laporan_s].items():
                    st.write(f"• {k}: {v}x")

        with col_s2:
            phrase = st.text_input("Frasa yang dicari:", placeholder="contoh: tanggal penilaian")
            if phrase and files_s:
                grand_total = 0
                for uf in files_s:
                    if uf.name.endswith(".pdf"):
                        pages = extract_text_pdf(uf)
                    else:
                        pages = extract_text_docx(uf)

                    pattern = re.compile(
                        r"\s*".join(re.escape(w) for w in phrase.split()),
                        re.IGNORECASE
                    )
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
                                r'<mark style="background:#f0a500;color:#000;padding:0 2px;">\1</mark>',
                                page_txt,
                                flags=re.IGNORECASE,
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
        if not riwayat:
            st.info("Belum ada riwayat audit. Jalankan analisis AI terlebih dahulu.")
        else:
            # Tampilkan dari terbaru
            for audit_id in sorted(riwayat.keys(), reverse=True):
                r = riwayat[audit_id]
                score  = r.get("score", 0)
                sc_col = "#2ecc8a" if score >= 80 else "#f0a500" if score >= 60 else "#e05c5c"
                files_str = ", ".join(r.get("files", []))

                with st.expander(
                    f"🕒 {r.get('timestamp','')} &nbsp;·&nbsp; {files_str} &nbsp;·&nbsp; Skor: {score}"
                ):
                    col_h1, col_h2, col_h3, col_h4 = st.columns(4)
                    col_h1.metric("Skor QC", score)
                    col_h2.metric("🔴 Kritikal", r.get("kritikal", 0))
                    col_h3.metric("🟡 Minor",    r.get("minor", 0))
                    col_h4.metric("Mode",        r.get("mode", ""))

                    result_h = r.get("result", {})
                    exec_s   = result_h.get("summary", {}).get("executive_summary", "")
                    if exec_s:
                        st.caption(exec_s)

                    if st.button("Tampilkan Detail Temuan", key=f"hist_{audit_id}"):
                        for f in result_h.get("findings", []):
                            render_finding_card(f)

            if st.button("🗑 Hapus Semua Riwayat"):
                save_json(HISTORY_FILE, {})
                st.success("Riwayat dihapus.")
                st.rerun()

    # ────────────────────────────────
    # TAB 4: KELOLA REFERENSI
    # ────────────────────────────────
    with tab_ref:
        st.markdown("#### 📁 Kelola Data Referensi Laporan")
        st.caption("Simpan catatan frekuensi kemunculan frasa per jenis laporan sebagai baseline.")

        if not data_laporan:
            st.info("Belum ada data referensi. Tambah laporan baru di bawah.")
        else:
            for lap_name, lap_data in data_laporan.items():
                with st.expander(f"📋 {lap_name}"):
                    if lap_data:
                        for k, v in lap_data.items():
                            c1, c2, c3 = st.columns([3, 1, 1])
                            c1.write(k)
                            c2.write(f"**{v}x**")
                            if c3.button("🗑", key=f"del_{lap_name}_{k}"):
                                del data_laporan[lap_name][k]
                                save_json(DATA_FILE, data_laporan)
                                st.rerun()
                    else:
                        st.caption("Belum ada keterangan.")

                    st.markdown("**Tambah keterangan:**")
                    c_k, c_v, c_btn = st.columns([3, 1, 1])
                    new_k = c_k.text_input("Frasa", key=f"nk_{lap_name}")
                    new_v = c_v.number_input("Jumlah", min_value=0, step=1, key=f"nv_{lap_name}")
                    if c_btn.button("Tambah", key=f"nbtn_{lap_name}") and new_k:
                        data_laporan[lap_name][new_k] = new_v
                        save_json(DATA_FILE, data_laporan)
                        st.success("✅ Ditambahkan")
                        st.rerun()

        st.markdown("---")
        st.markdown("**Tambah Laporan Referensi Baru:**")
        c1, c2 = st.columns([3, 1])
        new_lap = c1.text_input("Nama laporan:", key="new_lap_name")
        if c2.button("Tambahkan", key="btn_new_lap") and new_lap:
            if new_lap not in data_laporan:
                data_laporan[new_lap] = {}
                save_json(DATA_FILE, data_laporan)
                st.success(f"✅ '{new_lap}' ditambahkan")
                st.rerun()
            else:
                st.warning("Nama sudah ada.")

    # ── FOOTER ──
    st.markdown("""
<hr style="margin-top:40px;">
<p style="text-align:center;font-size:12px;color:#7d8590;font-family:'DM Mono',monospace;">
    CekLaporan v6.0 &nbsp;·&nbsp; KJPP SRR &nbsp;·&nbsp;
    Powered by Claude AI (claude-sonnet-4-5) &nbsp;·&nbsp; Created by HW
</p>
""", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
