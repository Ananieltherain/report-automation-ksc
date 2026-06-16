import streamlit as st
import json
import tempfile
import os
import pytesseract

# Sesuaikan jalur ini dengan lokasi instalasi Tesseract Anda
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Sekarang Anda bisa menggunakannya untuk OCR
# teks = pytesseract.image_to_string('gambar.png')

from datetime import datetime

# ──────────────────────────────────────────────────────────────
# IMPORT ENGINE
# ──────────────────────────────────────────────────────────────
try:
    from lhs_engine import (
        generate_lhs,
        DAFTAR_ASURANSI, DAFTAR_SURVEYOR, DAFTAR_USAGE,
        DAFTAR_COVERAGE, DAFTAR_CASE, PASAL_PSAKBI,
    )
except ImportError as e:
    st.error(f"❌ `lhs_engine.py` tidak ditemukan.\nDetail: {e}")
    st.stop()

# st_extractor opsional — tidak crash jika belum install OCR
try:
    from st_extractor import extract_from_pdf
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# ──────────────────────────────────────────────────────────────
# KONFIGURASI HALAMAN
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="LHS Generator KSC",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────────────────────
# CSS TAMBAHAN (mobile-friendly improvements)
# ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Badge OCR */
.badge-ocr {
    background: #166534;
    color: white;
    border-radius: 6px;
    padding: 2px 8px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .5px;
}
.badge-manual {
    background: #6B7280;
    color: white;
    border-radius: 6px;
    padding: 2px 8px;
    font-size: 11px;
}
/* Info box extract */
.extract-box {
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 12px;
    font-size: 13px;
}
</style>
""", unsafe_allow_html=True)

st.title("⚡ LHS Generator")
st.caption("PT. Karya Solusi Cemerlang — Web Mobile-Friendly")

# ──────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ──────────────────────────────────────────────────────────────
DEFAULTS = {
    "saksi_list"          : [{} for _ in range(4)],
    "analisa_list"        : ["" for _ in range(3)],
    "claimable"           : "✅ CLAIMABLE",
    # field yang bisa diisi dari OCR
    "asuransi"            : DAFTAR_ASURANSI[0],
    "case"                : DAFTAR_CASE[0],
    "tertanggung"         : "",
    "nopolis"             : "",
    "alamat_tertanggung"  : "",
    "periode"             : "",
    "coverage"            : DAFTAR_COVERAGE[0],
    "usage"               : DAFTAR_USAGE[0],
    "surveyor"            : DAFTAR_SURVEYOR[0],
    "merk"                : "",
    "nopol"               : "",
    "tahun_kendaraan"     : "",
    "no_ka"               : "",
    "no_sin"              : "",
    "surat_kuasa"         : "",
    "tanggal_surat_kuasa" : "",
    "surat_tugas"         : "",
    "tanggal_surat_tugas" : "",
    "penyebab"            : "",
    "dol"                 : "",
    "dol_time"            : "",
    "alamat_tkp"          : "",
    "keterangan_tkp"      : "",
    "dol_lhs"             : datetime.today().strftime("%d %B %Y"),
    "kronologis_kejadian" : "",
    # meta
    "_ocr_filled"         : [],   # field yang terisi via OCR (untuk badge)
}

for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ──────────────────────────────────────────────────────────────
# HELPER: LOAD JSON
# ──────────────────────────────────────────────────────────────
def load_from_json(data: dict):
    """Isi session state dari dict (JSON load)."""
    for k, v in data.items():
        if not k.startswith("_"):
            st.session_state[k] = v
    if "saksi_list" in data:
        st.session_state.saksi_list = data["saksi_list"]
    if "analisa_list" in data:
        st.session_state.analisa_list = data["analisa_list"]


# ──────────────────────────────────────────────────────────────
# HELPER: ISI DARI HASIL OCR
# ──────────────────────────────────────────────────────────────
def apply_ocr_result(ocr: dict):
    """
    Isi session state dari hasil extract PDF.
    Hanya field yang berhasil diextract (tidak kosong) yang diisi.
    Field dropdown (asuransi, surveyor) dicari match terdekat.
    """
    filled = []

    # Field teks biasa — langsung assign
    text_fields = [
        "tertanggung", "nopolis", "alamat_tertanggung", "periode",
        "merk", "nopol", "tahun_kendaraan", "no_ka", "no_sin",
        "surat_kuasa", "tanggal_surat_kuasa",
        "surat_tugas", "tanggal_surat_tugas", "dol",
    ]
    for f in text_fields:
        val = ocr.get(f, "").strip()
        if val:
            st.session_state[f] = val
            filled.append(f)

    # Asuransi — cari match di DAFTAR_ASURANSI (case-insensitive)
    asuransi_ocr = ocr.get("asuransi", "").strip()
    if asuransi_ocr:
        match = next(
            (a for a in DAFTAR_ASURANSI
             if asuransi_ocr.upper() in a.upper() or a.upper() in asuransi_ocr.upper()),
            None
        )
        if match:
            st.session_state["asuransi"] = match
            filled.append("asuransi")
        else:
            # Tidak ada di list → tambahkan sementara ke list (hanya di session)
            st.session_state["asuransi_ocr_raw"] = asuransi_ocr
            filled.append("asuransi (manual)")

    # Surveyor — cari match di DAFTAR_SURVEYOR
    surveyor_ocr = ocr.get("surveyor", "").strip()
    if surveyor_ocr:
        # Coba match nama depan atau belakang
        match = next(
            (s for s in DAFTAR_SURVEYOR
             if any(part.upper() in s.upper()
                    for part in surveyor_ocr.split() if len(part) > 1)),
            None
        )
        if match:
            st.session_state["surveyor"] = match
            filled.append("surveyor")
        else:
            st.session_state["surveyor_ocr_raw"] = surveyor_ocr
            filled.append("surveyor (manual)")

    st.session_state["_ocr_filled"] = filled


# ──────────────────────────────────────────────────────────────
# BADGE HELPER
# ──────────────────────────────────────────────────────────────
def ocr_badge(field_key: str) -> str:
    """Tampilkan badge OCR jika field diisi via extract."""
    if field_key in st.session_state.get("_ocr_filled", []):
        return " 🟢"
    return ""


# ──────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📂 Manajemen Data")

    # ── Upload JSON
    uploaded_json = st.file_uploader("Load Data JSON", type=["json"])
    if uploaded_json:
        if st.button("📥 Muat JSON ini"):
            try:
                data = json.load(uploaded_json)
                load_from_json(data)
                st.success("✅ Data dimuat!")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal: {e}")

    st.divider()

    # ── Upload ST/SK PDF untuk OCR
    st.subheader("📄 Auto-Extract dari ST/SK")

    if not OCR_AVAILABLE:
        st.warning(
            "OCR tidak tersedia.\n\n"
            "Install: `pip install pytesseract pdf2image pillow`\n"
            "dan `tesseract-ocr` di sistem."
        )
    else:
        st.caption(
            "Upload PDF gabungan ST KSC + SK Asuransi. "
            "Field akan terisi otomatis dari dokumen."
        )
        uploaded_pdf = st.file_uploader(
            "Upload PDF ST/SK", type=["pdf"], key="pdf_upload"
        )

        if uploaded_pdf:
            if st.button("🔍 Extract & Isi Form", type="primary", use_container_width=True):
                with st.spinner("Memproses OCR... (15-30 detik)"):
                    try:
                        # Simpan ke temp file
                        with tempfile.NamedTemporaryFile(
                            delete=False, suffix=".pdf"
                        ) as tmp:
                            tmp.write(uploaded_pdf.read())
                            tmp_pdf_path = tmp.name

                        ocr_result = extract_from_pdf(tmp_pdf_path)
                        os.unlink(tmp_pdf_path)

                        apply_ocr_result(ocr_result)
                        n = len(st.session_state["_ocr_filled"])
                        st.success(f"✅ {n} field berhasil diisi otomatis!")

                        # Tampilkan ringkasan
                        with st.expander("Lihat hasil extract"):
                            for f in st.session_state["_ocr_filled"]:
                                key = f.replace(" (manual)", "")
                                val = st.session_state.get(key, "—")
                                st.write(f"• **{f}**: `{val}`")

                        st.rerun()

                    except Exception as e:
                        st.error(f"OCR gagal: {e}")

    st.divider()

    # ── Info field OCR
    if st.session_state.get("_ocr_filled"):
        st.caption("🟢 = diisi otomatis dari PDF")

    st.divider()
    st.info(
        "💡 Tombol **Download** muncul setelah klik "
        "'Siapkan Laporan' di bagian bawah."
    )


# ──────────────────────────────────────────────────────────────
# TABS FORMULIR UTAMA
# ──────────────────────────────────────────────────────────────

# Banner info OCR jika sudah diextract
if st.session_state.get("_ocr_filled"):
    n = len(st.session_state["_ocr_filled"])
    st.success(
        f"🟢 **{n} field terisi otomatis dari PDF.** "
        "Tandai 🟢 = dari OCR. Periksa & lengkapi yang masih kosong.",
        icon="✅"
    )

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏢 Identitas", "🚗 Kejadian & Ref", "🧑‍🤝‍🧑 Saksi", "🔍 Analisa", "⚖️ Kesimpulan"
])

# ──────────────────────────────────────────────────────────────
# TAB 1: IDENTITAS POLIS & KENDARAAN
# ──────────────────────────────────────────────────────────────
with tab1:
    st.subheader("Data Polis & Tertanggung")

    # Asuransi dropdown — handle jika dari OCR tidak ada di list
    asuransi_raw = st.session_state.get("asuransi_ocr_raw", "")
    if asuransi_raw:
        st.info(
            f"🟡 Asuransi dari OCR: **{asuransi_raw}** — tidak ditemukan di daftar. "
            "Pilih yang paling sesuai di bawah, atau tambahkan ke `DAFTAR_ASURANSI` di `lhs_engine.py`."
        )

    asuransi = st.selectbox(
        f"Asuransi{ocr_badge('asuransi')}",
        DAFTAR_ASURANSI,
        index=DAFTAR_ASURANSI.index(st.session_state["asuransi"])
              if st.session_state["asuransi"] in DAFTAR_ASURANSI else 0,
        key="asuransi",
    )
    case_type = st.selectbox(
        "Jenis Klaim (Case)",
        DAFTAR_CASE,
        index=DAFTAR_CASE.index(st.session_state["case"])
              if st.session_state["case"] in DAFTAR_CASE else 0,
        key="case",
    )
    tertanggung = st.text_input(
        f"Tertanggung{ocr_badge('tertanggung')}",
        key="tertanggung",
    )
    nopolis = st.text_input(
        f"Nomor Polis{ocr_badge('nopolis')}",
        key="nopolis",
    )
    alamat_tertanggung = st.text_area(
        f"Alamat Tertanggung{ocr_badge('alamat_tertanggung')}",
        key="alamat_tertanggung",
    )

    col1, col2 = st.columns(2)
    with col1:
        periode = st.text_input(
            f"Periode Polis{ocr_badge('periode')}",
            key="periode",
            placeholder="01/01/2024 s/d 01/01/2025",
        )
        coverage = st.selectbox(
            "Jaminan (Coverage)",
            DAFTAR_COVERAGE,
            index=DAFTAR_COVERAGE.index(st.session_state["coverage"])
                  if st.session_state["coverage"] in DAFTAR_COVERAGE else 0,
            key="coverage",
        )
    with col2:
        usage = st.selectbox(
            "Penggunaan (Usage)",
            DAFTAR_USAGE,
            index=DAFTAR_USAGE.index(st.session_state["usage"])
                  if st.session_state["usage"] in DAFTAR_USAGE else 0,
            key="usage",
        )

        # Surveyor dropdown — handle OCR raw
        surveyor_raw = st.session_state.get("surveyor_ocr_raw", "")
        if surveyor_raw:
            st.caption(
                f"🟡 Surveyor dari OCR: **{surveyor_raw}** "
                "— tidak cocok di daftar, pilih manual."
            )
        surveyor = st.selectbox(
            f"Surveyor{ocr_badge('surveyor')}",
            DAFTAR_SURVEYOR,
            index=DAFTAR_SURVEYOR.index(st.session_state["surveyor"])
                  if st.session_state["surveyor"] in DAFTAR_SURVEYOR else 0,
            key="surveyor",
        )

    st.subheader("Data Kendaraan")
    merk = st.text_input(
        f"Merk / Jenis Kendaraan{ocr_badge('merk')}",
        placeholder="Contoh: Honda Vario 125",
        key="merk",
    )
    nopol = st.text_input(
        f"No. Polisi{ocr_badge('nopol')}",
        placeholder="Contoh: BD4452WL",
        key="nopol",
    )

    col3, col4, col5 = st.columns(3)
    with col3:
        tahun_kendaraan = st.text_input(
            f"Tahun{ocr_badge('tahun_kendaraan')}",
            key="tahun_kendaraan",
        )
    with col4:
        no_ka = st.text_input(
            f"No Rangka{ocr_badge('no_ka')}",
            key="no_ka",
        )
    with col5:
        no_sin = st.text_input(
            f"No Mesin{ocr_badge('no_sin')}",
            key="no_sin",
        )


# ──────────────────────────────────────────────────────────────
# TAB 2: KEJADIAN & REFERENSI SURAT
# ──────────────────────────────────────────────────────────────
with tab2:
    st.subheader("Data Kejadian (DOL)")

    penyebab = st.text_input(
        "Penyebab Kejadian",
        placeholder="Contoh: tabrakan dari arah belakang",
        key="penyebab",
    )

    col6, col7 = st.columns(2)
    with col6:
        dol = st.text_input(
            f"Hari / Tanggal Kejadian{ocr_badge('dol')}",
            placeholder="Senin, 1 Januari 2024",
            key="dol",
        )
    with col7:
        dol_time = st.text_input(
            "Waktu Kejadian",
            placeholder="08.30 WIB",
            key="dol_time",
        )

    alamat_tkp = st.text_input("Lokasi TKP", key="alamat_tkp")
    keterangan_tkp = st.text_area("Keterangan TKP", key="keterangan_tkp")
    dol_lhs = st.text_input(
        "Tanggal Laporan LHS",
        key="dol_lhs",
    )

    st.subheader("Referensi Surat Tugas & Kuasa")

    # Banner jika sudah ada dari OCR
    if any(f in st.session_state.get("_ocr_filled", [])
           for f in ["surat_kuasa", "surat_tugas", "tanggal_surat_kuasa", "tanggal_surat_tugas"]):
        st.success("🟢 Nomor & tanggal surat sudah terisi dari PDF.")

    col8, col9 = st.columns(2)
    with col8:
        surat_kuasa = st.text_input(
            f"No. Surat Kuasa Asuransi{ocr_badge('surat_kuasa')}",
            key="surat_kuasa",
        )
        tgl_surat_kuasa = st.text_input(
            f"Tgl Surat Kuasa{ocr_badge('tanggal_surat_kuasa')}",
            key="tanggal_surat_kuasa",
        )
    with col9:
        surat_tugas = st.text_input(
            f"No. Surat Tugas KSC{ocr_badge('surat_tugas')}",
            key="surat_tugas",
        )
        tgl_surat_tugas = st.text_input(
            f"Tgl Surat Tugas{ocr_badge('tanggal_surat_tugas')}",
            key="tanggal_surat_tugas",
        )


# ──────────────────────────────────────────────────────────────
# TAB 3: SAKSI DINAMIS
# ──────────────────────────────────────────────────────────────
with tab3:
    st.subheader("Saksi-Saksi di Lapangan")
    st.caption("Minimal 1, maksimal 10 saksi.")

    for i in range(len(st.session_state.saksi_list)):
        with st.expander(f"👤 Saksi {i+1}", expanded=(i == 0)):
            saksi = st.session_state.saksi_list[i]

            saksi["nama"] = st.text_input(
                "Nama Saksi", value=saksi.get("nama", ""), key=f"s_nama_{i}"
            )

            c1, c2 = st.columns(2)
            with c1:
                saksi["role"] = st.text_input(
                    "Role / Status", value=saksi.get("role", ""), key=f"s_role_{i}"
                )
                saksi["usia"] = st.text_input(
                    "Usia (th)", value=saksi.get("usia", ""), key=f"s_usia_{i}"
                )
            with c2:
                saksi["pekerjaan"] = st.text_input(
                    "Pekerjaan", value=saksi.get("pekerjaan", ""), key=f"s_kerja_{i}"
                )
                saksi["tanggal_interview"] = st.text_input(
                    "Tgl Interview", value=saksi.get("tanggal_interview", ""),
                    key=f"s_tglint_{i}"
                )

            saksi["alamat"] = st.text_area(
                "Alamat", value=saksi.get("alamat", ""), key=f"s_alamat_{i}"
            )
            saksi["keterangan"] = st.text_area(
                "Keterangan Saksi", value=saksi.get("keterangan", ""),
                key=f"s_ket_{i}", height=120,
                help="Tekan Enter untuk paragraf baru — akan tampil sebagai baris baru di dokumen."
            )

            if st.button(f"🗑️ Hapus Saksi {i+1}", key=f"del_saksi_{i}"):
                if len(st.session_state.saksi_list) > 1:
                    st.session_state.saksi_list.pop(i)
                    st.rerun()
                else:
                    st.warning("Minimal harus ada 1 saksi.")

    if len(st.session_state.saksi_list) < 10:
        if st.button("➕ Tambah Saksi"):
            st.session_state.saksi_list.append({})
            st.rerun()


# ──────────────────────────────────────────────────────────────
# TAB 4: ANALISA DINAMIS
# ──────────────────────────────────────────────────────────────
with tab4:
    st.subheader("Analisa Fakta")
    st.caption("Setiap poin = satu baris di laporan.")

    for i in range(len(st.session_state.analisa_list)):
        ca1, ca2 = st.columns([10, 1])
        with ca1:
            st.session_state.analisa_list[i] = st.text_input(
                f"Poin {i+1}",
                value=st.session_state.analisa_list[i],
                key=f"analisa_{i}",
            )
        with ca2:
            st.write("")
            st.write("")
            if st.button("❌", key=f"del_analisa_{i}"):
                st.session_state.analisa_list.pop(i)
                st.rerun()

    if len(st.session_state.analisa_list) < 10:
        if st.button("➕ Tambah Poin Analisa"):
            st.session_state.analisa_list.append("")
            st.rerun()


# ──────────────────────────────────────────────────────────────
# TAB 5: KESIMPULAN
# ──────────────────────────────────────────────────────────────
with tab5:
    st.subheader("Kronologis & Kesimpulan")

    kronologis_kejadian = st.text_area(
        "Kronologis Kejadian",
        height=160,
        key="kronologis_kejadian",
        help="Tekan Enter untuk paragraf baru.",
    )

    claimable_status = st.radio(
        "Hasil Klaim:",
        ["✅ CLAIMABLE", "❌ UNCLAIMABLE"],
        key="claimable",
    )
    is_claimable = (claimable_status == "✅ CLAIMABLE")

    pasal_unclaimable = ""
    if not is_claimable:
        st.error("⚠️ Laporan ditandai UNCLAIMABLE")
        pasal_labels = [f"{p} — {desc}" for p, desc in PASAL_PSAKBI]
        selected_pasal = st.selectbox("Pilih Pasal PSAKBI:", [""] + pasal_labels)
        if selected_pasal:
            pasal_unclaimable = selected_pasal.split(" — ")[0]


# ──────────────────────────────────────────────────────────────
# GENERATE & DOWNLOAD
# ──────────────────────────────────────────────────────────────
st.divider()
st.subheader("⚙️ Proses Dokumen")

TEMPLATE_NAME = "LHS FINAL - Template.docx"
template_path_local = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), TEMPLATE_NAME
)

uploaded_template = st.file_uploader(
    f"Upload Template (Opsional — default: '{TEMPLATE_NAME}' di server)",
    type=["docx"],
    key="template_upload",
)

if st.button(
    "📄 Siapkan Laporan", type="primary", use_container_width=True
):
    # ── Validasi
    errors = []
    if not st.session_state.get("tertanggung", "").strip():
        errors.append("Tertanggung belum diisi")
    if not st.session_state.get("nopol", "").strip():
        errors.append("No. Polisi belum diisi")
    saksi_bersih = [
        s for s in st.session_state.saksi_list
        if s.get("nama", "").strip()
    ]
    if not saksi_bersih:
        errors.append("Minimal 1 saksi harus memiliki nama")

    if errors:
        for e in errors:
            st.warning(f"⚠️ {e}")
        st.stop()

    analisa_bersih = [
        a.strip() for a in st.session_state.analisa_list if a.strip()
    ]

    # ── Compile data
    collected_data = {
        "asuransi"            : st.session_state.get("asuransi", ""),
        "case"                : st.session_state.get("case", ""),
        "tertanggung"         : st.session_state.get("tertanggung", ""),
        "nopolis"             : st.session_state.get("nopolis", ""),
        "merk"                : st.session_state.get("merk", ""),
        "nopol"               : st.session_state.get("nopol", ""),
        "tahun_kendaraan"     : st.session_state.get("tahun_kendaraan", ""),
        "no_ka"               : st.session_state.get("no_ka", ""),
        "no_sin"              : st.session_state.get("no_sin", ""),
        "surat_kuasa"         : st.session_state.get("surat_kuasa", ""),
        "tanggal_surat_kuasa" : st.session_state.get("tanggal_surat_kuasa", ""),
        "surat_tugas"         : st.session_state.get("surat_tugas", ""),
        "tanggal_surat_tugas" : st.session_state.get("tanggal_surat_tugas", ""),
        "surveyor"            : st.session_state.get("surveyor", ""),
        "alamat_tertanggung"  : st.session_state.get("alamat_tertanggung", ""),
        "periode"             : st.session_state.get("periode", ""),
        "coverage"            : st.session_state.get("coverage", ""),
        "usage"               : st.session_state.get("usage", ""),
        "penyebab"            : st.session_state.get("penyebab", ""),
        "dol"                 : st.session_state.get("dol", ""),
        "dol_time"            : st.session_state.get("dol_time", ""),
        "alamat_tkp"          : st.session_state.get("alamat_tkp", ""),
        "keterangan_tkp"      : st.session_state.get("keterangan_tkp", ""),
        "dol_lhs"             : st.session_state.get("dol_lhs", ""),
        "kronologis_kejadian" : st.session_state.get("kronologis_kejadian", ""),
        "claimable"           : is_claimable,
        "pasal_unclaimable"   : pasal_unclaimable,
        "saksi_list"          : saksi_bersih,
        "analisa_list"        : analisa_bersih,
    }

    # ── Resolve template
    active_template = ""
    if uploaded_template:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as t:
            t.write(uploaded_template.read())
            active_template = t.name
    elif os.path.exists(template_path_local):
        active_template = template_path_local
    else:
        st.error(
            f"❌ Template `{TEMPLATE_NAME}` tidak ditemukan di server. "
            "Silakan upload manual via input di atas."
        )
        st.stop()

    # ── Generate dokumen
    try:
        safe_name = (
            st.session_state.get("tertanggung", "LHS")
            .replace(" ", "_")[:30]
            + "_"
            + st.session_state.get("nopol", "").replace(" ", "")
        )
        out_filename = f"LHS_{safe_name}.docx"

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            output_path = tmp.name

        generate_lhs(active_template, output_path, collected_data)

        with open(output_path, "rb") as f:
            docx_bytes = f.read()

        # Bersihkan temp
        os.unlink(output_path)
        if uploaded_template:
            try:
                os.unlink(active_template)
            except Exception:
                pass

        json_str = json.dumps(collected_data, ensure_ascii=False, indent=2)

        st.success("✅ Laporan berhasil dibuat! Silakan unduh.")

        cd1, cd2 = st.columns(2)
        with cd1:
            st.download_button(
                "⬇️ Download Dokumen (.docx)",
                data=docx_bytes,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument"
                     ".wordprocessingml.document",
                use_container_width=True,
            )
        with cd2:
            st.download_button(
                "💾 Backup Data (.json)",
                data=json_str,
                file_name=f"Data_LHS_{safe_name}.json",
                mime="application/json",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"❌ Gagal membuat dokumen: {e}")
        st.exception(e)