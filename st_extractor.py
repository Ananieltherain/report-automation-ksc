"""
st_extractor.py — Ekstraksi otomatis field LHS dari PDF Surat Tugas
PT. Karya Solusi Cemerlang

Cara kerja:
  1. User upload 2 file: ST KSC (halaman 1) + SK Asuransi (halaman 2)
     ATAU 1 file PDF yang sudah digabung (seperti contoh)
  2. OCR dengan Tesseract (gratis, lokal, tidak perlu API)
  3. Regex extract semua field yang bisa diambil otomatis
  4. Return dict langsung kompatibel dengan generate_lhs()

Field yang di-auto-extract:
  ✅ surat_tugas          → dari ST KSC
  ✅ tanggal_surat_tugas  → dari ST KSC
  ✅ surat_kuasa          → dari SK Asuransi (nomor di header)
  ✅ tanggal_surat_kuasa  → dari SK Asuransi
  ✅ asuransi             → dari SK Asuransi
  ✅ surveyor             → dari ST KSC (Memberi Tugas Kepada → Nama)
  ✅ tertanggung          → dari ST KSC / SK (INSURED / Nama Tertanggung)
  ✅ alamat_tertanggung   → dari ST KSC
  ✅ merk                 → dari ST KSC / SK
  ✅ tahun_kendaraan      → dari ST KSC / SK
  ✅ no_ka                → dari ST KSC / SK
  ✅ no_sin               → dari ST KSC / SK
  ✅ nopol                → dari ST KSC / SK
  ✅ nopolis              → dari ST KSC / SK
  ✅ dol                  → dari SK (TANGGAL KEJADIAN)

Field yang TETAP diisi manual (tidak ada di ST/SK):
  ✏️  case, coverage, usage, periode, penyebab,
  ✏️  dol_time, alamat_tkp, keterangan_tkp, dol_lhs,
  ✏️  kronologis, saksi_list, analisa_list, claimable

Dependencies: pip install pytesseract pdf2image pillow
System: tesseract-ocr harus terinstall
"""

import re
import os
import pytesseract

# Sesuaikan jalur ini dengan lokasi instalasi Tesseract Anda
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Sekarang Anda bisa menggunakannya untuk OCR
# teks = pytesseract.image_to_string('gambar.png')

from typing import Optional

# ──────────────────────────────────────────────────────────────
# OCR ENGINE
# ──────────────────────────────────────────────────────────────

def _pdf_to_text(pdf_path: str) -> list[str]:
    """
    Convert PDF ke list of text strings (satu per halaman).
    Gunakan Tesseract OCR — gratis, lokal, tidak perlu internet.
    Return: ['teks halaman 1', 'teks halaman 2', ...]
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract
    except ImportError as e:
        raise ImportError(
            f"Install dulu: pip install pytesseract pdf2image pillow\n{e}"
        )

    pages = convert_from_path(pdf_path, dpi=250)
    texts = []
    for page in pages:
        text = pytesseract.image_to_string(page, lang='eng')
        texts.append(text)
    return texts


def _image_to_text(image_path: str) -> str:
    """OCR dari file gambar (PNG/JPG)."""
    try:
        from PIL import Image
        import pytesseract
        img = Image.open(image_path)
        return pytesseract.image_to_string(img, lang='eng')
    except ImportError as e:
        raise ImportError(f"Install dulu: pip install pytesseract pillow\n{e}")


# ──────────────────────────────────────────────────────────────
# REGEX HELPERS
# ──────────────────────────────────────────────────────────────

def _find(pattern: str, text: str, group: int = 1,
          flags: int = re.IGNORECASE) -> str:
    """Cari pattern, return group. Empty string jika tidak ketemu."""
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else ''


def _clean(text: str) -> str:
    """Bersihkan whitespace berlebih."""
    return re.sub(r'\s+', ' ', text).strip()


# ──────────────────────────────────────────────────────────────
# CORE EXTRACTOR
# ──────────────────────────────────────────────────────────────

def _extract_from_st_ksc(text: str) -> dict:
    """
    Extract field dari ST KSC (halaman PT. Karya Solusi Cemerlang).
    Pola: Nomor: XXX/SVY.../KSC/...
    """
    result = {}

    # No ST KSC — pola: Nomor: .../KSC/...
    result['surat_tugas'] = _find(
        r'Nomor\s*:\s*([A-Z0-9/\-\.]+/KSC/[A-Z0-9/\-\.]+)', text
    )

    # Tanggal ST — "Jakarta, DD Bulan YYYY"
    result['tanggal_surat_tugas'] = _find(
        r'Jakarta,\s+(\d{1,2}\s+\w+\s+\d{4})', text
    )

    # Surveyor — dari blok "Memberi Tugas Kepada ... Nama : XXX"
    blok = re.search(
        r'Memberi Tugas Kepada.*?Nama\s*:\s*([A-Z][A-Z .]+?)(?:\nJabatan|\nAlamat)',
        text, re.DOTALL | re.IGNORECASE
    )
    if blok:
        result['surveyor'] = _clean(blok.group(1))
    else:
        # Fallback: cari nama setelah "Surveyor" label
        result['surveyor'] = _find(r'Jabatan\s*:\s*Surveyor[^a-z]*\n[^:]+:\s*([A-Z][A-Z .]+)', text)

    # Asuransi — dari "PT. Asuransi XXX Nomor :"
    result['asuransi'] = _find(
        r'(PT\.?\s+(?:Asuransi\s+)?[A-Za-z ]+?)\s+Nomor\s*:', text
    )

    # No Surat Kuasa Asuransi — dari "PT. Asuransi ... Nomor : XXX tanggal"
    result['surat_kuasa'] = _find(
        r'Nomor\s*:\s*([A-Z0-9/\-\.]+/\d{4})\s+tanggal', text
    )

    # Tanggal Surat Kuasa — dari "... Nomor : XXX tanggal DD Bulan YYYY"
    result['tanggal_surat_kuasa'] = _find(
        r'tanggal\s+(\d{1,2}\s+\w+\s+\d{4})\s+untuk', text
    )

    # Tertanggung
    result['tertanggung'] = _find(
        r'Nama\s+Tertanggung\s*:\s*(.+)', text
    )

    # Alamat tertanggung — multiline sampai baris berikutnya
    m = re.search(
        r'Alamat\s*:\s*((?:[A-Z][^\n]+\n?){1,3}?)(?:\nMerk|\nNo|\nTahun|\nNomor)',
        text, re.IGNORECASE
    )
    if m:
        result['alamat_tertanggung'] = _clean(m.group(1))
    else:
        result['alamat_tertanggung'] = ''

    # Merk / Type
    result['merk'] = _find(r'Merk\s*/\s*Type\s*:\s*(.+)', text)

    # Tahun
    result['tahun_kendaraan'] = _find(
        r'Tahun\s*[/,]\s*(?:Warna\s*[:/]\s*)?(\d{4})', text
    )

    # No Rangka & Mesin — "MH.../JMK..."
    m = re.search(
        r'No\s+Rangka\s*/\s*Mesin\s*:\s*([A-Z0-9]+)\s*/\s*([A-Z0-9]+)',
        text, re.IGNORECASE
    )
    if m:
        result['no_ka']  = m.group(1).strip()
        result['no_sin'] = m.group(2).strip()
    else:
        result['no_ka'] = result['no_sin'] = ''

    # No Polisi kendaraan
    result['nopol'] = _find(
        r'Nomor\s+Polisi\s*:\s*([A-Z]{1,3}\s*\d+\s*[A-Z]{0,4})', text
    )

    # No Polis Asuransi
    result['nopolis'] = _find(
        r'Nomor\s+Polis\s+Asuransi\s*:\s*([A-Z0-9\-]+)', text
    )

    # Periode Polis — bisa "-" atau tanggal range
    m = re.search(r'Periode\s+Polis\s*:\s*([^\n]+)', text, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        result['periode'] = '' if val == '-' else val
    else:
        result['periode'] = ''

    return result


def _extract_from_sk_asuransi(text: str) -> dict:
    """
    Extract field dari SK Asuransi (format berbeda per perusahaan).
    Pola umum: field UPPERCASE seperti POLIS, INSURED, MERK/TYPE, dsb.
    """
    result = {}

    # No SK Asuransi — nomor di header (biasanya baris ke-2 setelah judul)
    # Pola: XX/MV/XXX/YYYY atau XX/CLM/XXX/YYYY
    result['surat_kuasa'] = _find(
        r'^([A-Z0-9]+/(?:MV|CLM|AUTO|MOT|PKB)[^/]*/[^/]+/\d{4})$',
        text, flags=re.MULTILINE
    )

    # Tanggal SK Asuransi
    result['tanggal_surat_kuasa'] = _find(
        r'Jakarta,\s+(\d{1,2}\s+\w+\s+\d{4})', text
    )

    # Asuransi — dari "kami PT XXX,"
    result['asuransi'] = _find(
        r'kami\s+(PT\.?\s+[A-Za-z ]+?),', text
    )
    if not result['asuransi']:
        # Fallback: dari nama perusahaan di header
        result['asuransi'] = _find(
            r'^(PT\.?\s+[A-Za-z ]+)\n', text, flags=re.MULTILINE
        )

    # Tertanggung
    result['tertanggung'] = _find(r'INSURED\s*:\s*(.+)', text)

    # Merk
    result['merk'] = _find(r'MERK\s*/\s*TYPE\s*:\s*(.+)', text)

    # Tahun
    result['tahun_kendaraan'] = _find(r'TAHUN\s+PEMBUATAN\s*:\s*(\d{4})', text)

    # No Polisi
    result['nopol'] = _find(
        r'NO\.\s*POLISI\s*:\s*([A-Z]{1,3}\s*\d+\s*[A-Z]{0,4})', text
    )

    # No Mesin / Rangka — urutan bisa terbalik tergantung format SK
    m = re.search(
        r'NO\.\s*MESIN\s*/\s*RANGKA\s*:\s*([A-Z0-9]+)\s*/\s*([A-Z0-9]+)',
        text, re.IGNORECASE
    )
    if m:
        # Format ini: MESIN / RANGKA → group1=mesin, group2=rangka
        result['no_sin'] = m.group(1).strip()
        result['no_ka']  = m.group(2).strip()
    else:
        m = re.search(
            r'NO\.\s*RANGKA\s*/\s*MESIN\s*:\s*([A-Z0-9]+)\s*/\s*([A-Z0-9]+)',
            text, re.IGNORECASE
        )
        if m:
            result['no_ka']  = m.group(1).strip()
            result['no_sin'] = m.group(2).strip()
        else:
            result['no_ka'] = result['no_sin'] = ''

    # No Polis Asuransi
    result['nopolis'] = _find(r'POLIS\s*:\s*([A-Z0-9\-]+)', text)

    # Tanggal Kejadian (DOL) — OCR kadang insert "—" atau ">" sebagai noise
    result['dol'] = _find(
        r'TANGGAL\s+KEJADIAN[\s\-\—\>:]+([A-Za-z0-9 ]+)', text
    )

    return result


def _merge_results(st_data: dict, sk_data: dict) -> dict:
    """
    Gabungkan hasil ekstraksi ST KSC + SK Asuransi.
    ST KSC lebih dipercaya untuk data KSC (no ST, surveyor, tanggal ST).
    SK Asuransi lebih dipercaya untuk no SK asuransi dan tanggal kejadian.
    Fallback ke sumber lain jika salah satu kosong.
    """
    merged = {}

    # Field prioritas ST KSC
    ksc_priority = ['surat_tugas', 'tanggal_surat_tugas', 'surveyor',
                    'tertanggung', 'alamat_tertanggung', 'periode']
    for f in ksc_priority:
        merged[f] = st_data.get(f) or sk_data.get(f, '')

    # Field prioritas SK Asuransi
    sk_priority = ['surat_kuasa', 'tanggal_surat_kuasa', 'dol']
    for f in sk_priority:
        merged[f] = sk_data.get(f) or st_data.get(f, '')

    # Field yang bisa dari keduanya (ambil yang lebih lengkap)
    common = ['asuransi', 'merk', 'tahun_kendaraan', 'nopol', 'nopolis',
              'no_ka', 'no_sin']
    for f in common:
        v_st = st_data.get(f, '')
        v_sk = sk_data.get(f, '')
        # Pilih yang lebih panjang (biasanya lebih lengkap)
        merged[f] = v_st if len(v_st) >= len(v_sk) else v_sk

    # Bersihkan trailing/leading whitespace
    for k in merged:
        if isinstance(merged[k], str):
            merged[k] = merged[k].strip()

    return merged


# ──────────────────────────────────────────────────────────────
# SMART DETECTION: ST KSC vs SK Asuransi
# ──────────────────────────────────────────────────────────────

def _detect_page_type(text: str) -> str:
    """
    Deteksi apakah halaman ini ST KSC atau SK Asuransi.
    Return: 'st_ksc' | 'sk_asuransi' | 'unknown'
    """
    if 'KARYA SOLUSI CEMERLANG' in text.upper() and 'SURAT TUGAS' in text.upper():
        # Cek apakah ini ST KSC (dari perusahaan kita)
        if re.search(r'/KSC/', text):
            return 'st_ksc'
    if re.search(r'INSURED\s*:', text, re.IGNORECASE):
        return 'sk_asuransi'
    if re.search(r'Memberi Tugas Kepada', text, re.IGNORECASE):
        return 'st_ksc'
    if re.search(r'Nama\s+Tertanggung\s*:', text, re.IGNORECASE):
        return 'st_ksc'
    return 'unknown'


# ──────────────────────────────────────────────────────────────
# PUBLIC API
# ──────────────────────────────────────────────────────────────

def extract_from_pdf(pdf_path: str) -> dict:
    """
    Extract semua field dari PDF (bisa 1 atau 2 halaman).
    Deteksi otomatis mana ST KSC dan mana SK Asuransi.

    Return: dict siap pakai untuk generate_lhs()
            + key tambahan '_ocr_raw' (list text per halaman, untuk debug)
            + key '_confidence' (field mana yang berhasil diextract)
    """
    texts = _pdf_to_text(pdf_path)

    st_text = ''
    sk_text = ''

    for text in texts:
        ptype = _detect_page_type(text)
        if ptype == 'st_ksc':
            st_text = text
        elif ptype == 'sk_asuransi':
            sk_text = text
        else:
            # Jika tidak terdeteksi, coba assign berdasarkan urutan
            if not st_text:
                st_text = text
            elif not sk_text:
                sk_text = text

    st_data = _extract_from_st_ksc(st_text) if st_text else {}
    sk_data = _extract_from_sk_asuransi(sk_text) if sk_text else {}

    result = _merge_results(st_data, sk_data)

    # Meta info untuk UI
    filled  = [k for k, v in result.items() if v]
    empty   = [k for k, v in result.items() if not v]
    result['_ocr_pages']   = len(texts)
    result['_filled']      = filled
    result['_empty']       = empty

    return result


def extract_from_images(st_image_path: str = None,
                        sk_image_path: str = None) -> dict:
    """
    Extract dari file gambar terpisah (PNG/JPG).
    Satu atau keduanya bisa None.
    """
    st_text = _image_to_text(st_image_path) if st_image_path else ''
    sk_text = _image_to_text(sk_image_path) if sk_image_path else ''

    st_data = _extract_from_st_ksc(st_text) if st_text else {}
    sk_data = _extract_from_sk_asuransi(sk_text) if sk_text else {}

    result = _merge_results(st_data, sk_data)
    filled = [k for k, v in result.items() if v and not k.startswith('_')]
    result['_filled'] = filled
    result['_empty']  = [k for k, v in result.items() if not v and not k.startswith('_')]
    return result


# ──────────────────────────────────────────────────────────────
# CLI TEST
# ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    import sys, json

    pdf = sys.argv[1] if len(sys.argv) > 1 else 'from_jpgs__1_.pdf'
    if not os.path.exists(pdf):
        print(f"File tidak ditemukan: {pdf}")
        sys.exit(1)

    print(f"Memproses: {pdf}")
    result = extract_from_pdf(pdf)

    print(f"\n✅ Berhasil extract {len(result['_filled'])} field:")
    fields_display = [k for k in result if not k.startswith('_')]
    for k in fields_display:
        v = result[k]
        status = '✅' if v else '⬜'
        print(f"  {status} {k:25s}: '{v}'")

    if result['_empty']:
        needs_manual = [k for k in result['_empty'] if not k.startswith('_')]
        print(f"\n⬜ Perlu diisi manual: {needs_manual}")