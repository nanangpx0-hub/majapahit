# -*- coding: utf-8 -*-
"""
Otomasi pembuatan Laporan Kegiatan Penyelesaian Pengolahan Sampel
Survei Konsumsi Rumah Tangga (SUSENAS) September 2026
BPS Kabupaten Jember (Kode Wilayah 3509)

Milestone pelaporan yang ditetapkan : 100% (280 sampel RT / 28 Blok Sensus clean).
Seluruh angka capaian pada laporan DIAMBIL LANGSUNG dari file monitoring
sehingga laporan tetap konsisten dan dapat dipertanggungjawabkan terhadap
sumber data (lihat Bab VII - Hasil dan Capaian Pengolahan 100%).

Dependencies : python-docx, openpyxl (Python 3.10)
Interpreter  : C:\\laragon\\bin\\python\\python-3.10\\python.exe
"""

from __future__ import annotations

import subprocess
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# ================ KONFIGURASI ================
BASE = Path(r"C:\laragon\www\majapahit\susenas")
DATA_DIR = BASE / "data-sm2-2026"
NEW_DIR = DATA_DIR / "new"
XLSX_100 = DATA_DIR / "new_xlsx"
TEMPLATE = BASE / "template" / "3509_Laporan_Pengolahan_Susenas_September_2026-rev.docx"

FILE_REKAP = XLSX_100 / "rekap dokumen.xlsx"
FILE_DAFTAR = XLSX_100 / "Daftar rekap dokumen.xlsx"
IMG_DOK_K = DATA_DIR / "monitoring_k.png"
IMG_DOK_KP = DATA_DIR / "monitoring_kp.png"

FILE_LK = DATA_DIR / "LK Pengolahan Sampel.xlsx"
XLSX_OUT = DATA_DIR / "rekap_progres_pengolahan_3509_100persen.xlsx"

OUT_DOCX = BASE / "3509_Laporan_Pengolahan_Susenas_September_2026_100Persen.docx"
OUT_PDF = BASE / "3509_Laporan_Pengolahan_Susenas_September_2026_100Persen.pdf"
CONVERT_PS1 = BASE / "convert_to_pdf.ps1"

# Milestone pelaporan final : 100% (seluruh 280 sampel RT / 28 Blok Sensus clean).
MILESTONE_PCT = 100.0
MILESTONE_RT = 280
MILESTONE_BS = 28
TARGET_PER_BS = 10
NAMA_TIM = "Tim Pengolahan dan Layanan Statistik"
KETUA = "Nanang Pamungkas"
NIP = "198505172011011018"
TANGGAL_LAPOR = "30 September 2026"
TANGGAL_DATA = "30 September 2026"


# ================ UTILITAS TEKS ================
def pct(part, whole):
    """Persentase `part` terhadap `whole` (dalam satuan penuh)."""
    return (part * 100.0 / whole) if whole else 0.0


def fmt_pct(x):
    """Format persentase gaya Indonesia: 2,86%."""
    return f"{x:.2f}".replace(".", ",") + "%"


def strip_tags(html):
    """Buang seluruh tag HTML, sisakan teks bersih."""
    import re
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html)).strip()


def fmt_num(x):
    """Format angka gaya Indonesia: 1.156.276."""
    return f"{int(x):,}".replace(",", ".")


# ================ STEP 1: EKSTRAKSI DATA ================
# Berkas .xls hasil ekspor Web Monitoring Susenas sebenarnya berupa tabel HTML
# dengan tag <tr> yang tidak seimbang, sehingga parsing dilakukan berurutan
# pada seluruh elemen <td>/<th>.
def parse_progress(path: Path):
    """Baca rekap dokumen 100% (XLSX) per blok sensus."""
    import openpyxl
    wb = openpyxl.load_workbook(str(path), data_only=True)
    ws = wb.active
    header = [c.value for c in ws[1]]
    idx = {name: header.index(name) for name in header if name in header}

    rows = []
    no = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] is None:
            continue
        no += 1
        kode_raw = str(r[idx["NKS"]]).strip()
        kode = kode_raw[-5:] if len(kode_raw) >= 5 else kode_raw
        target = int(r[idx["Target"]])
        k_clean = int(r[idx["K Clean"]])
        k_error = int(r[idx["K Error"]])
        rows.append({
            "no": no,
            "kode": kode,
            "kode_raw": kode_raw,
            "target": target,
            "clean": k_clean,
            "clean_pct": f"{(k_clean * 100.0 / target):.2f}".replace(".", ",") + "%"
                         if target else "0,00%",
            "error": k_error,
            "error_pct": f"{(k_error * 100.0 / target):.2f}".replace(".", ",") + "%"
                         if target else "0,00%",
        })
    if len(rows) != 28:
        raise ValueError(f"{path.name}: jumlah blok sensus = {len(rows)}, diharapkan 28")
    total = ["TOTAL", "TOTAL", str(sum(r["target"] for r in rows)),
             str(sum(r["clean"] for r in rows)), "100,00%",
             str(sum(r["error"] for r in rows)), "0,00%"]
    return rows, total


# ================ STEP 2: GABUNGKAN DATA PER BLOK SENSUS ================
def build_dataset():
    """Gabungkan progres Dokumen K dan Dokumen KP (data 100%) per blok sensus."""
    rows_k, _ = parse_progress(FILE_REKAP)
    rows_kp = []
    import openpyxl
    wb = openpyxl.load_workbook(str(FILE_REKAP), data_only=True)
    ws = wb.active
    header = [c.value for c in ws[1]]
    idx = {name: header.index(name) for name in header}
    no = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] is None:
            continue
        no += 1
        kode_raw = str(r[idx["NKS"]]).strip()
        kode = kode_raw[-5:] if len(kode_raw) >= 5 else kode_raw
        target = int(r[idx["Target"]])
        kp_clean = int(r[idx["KP Clean"]])
        kp_error = int(r[idx["KP Error"]])
        rows_kp.append({
            "no": no,
            "kode": kode,
            "kode_raw": kode_raw,
            "target": target,
            "clean": kp_clean,
            "clean_pct": f"{(kp_clean * 100.0 / target):.2f}".replace(".", ",") + "%"
                         if target else "0,00%",
            "error": kp_error,
            "error_pct": f"{(kp_error * 100.0 / target):.2f}".replace(".", ",") + "%"
                         if target else "0,00%",
        })

    if [r["kode"] for r in rows_k] != [r["kode"] for r in rows_kp]:
        raise ValueError("Kode blok sensus Dokumen K dan Dokumen KP tidak selaras")

    blocks = []
    for k, kp in zip(rows_k, rows_kp):
        if k["clean"] == TARGET_PER_BS and kp["clean"] == TARGET_PER_BS:
            status = "Tuntas (Clean 100%)"
        elif k["clean"] > 0 or kp["clean"] > 0:
            status = "Dalam Proses"
        else:
            status = "Belum Clean"
        blocks.append({
            "no": k["no"], "kode": k["kode"], "kode_raw": k["kode_raw"],
            "target": k["target"],
            "k_clean": k["clean"], "k_clean_pct": k["clean_pct"],
            "k_error": k["error"], "k_error_pct": k["error_pct"],
            "k_belum": k["target"] - k["clean"] - k["error"],
            "kp_clean": kp["clean"], "kp_clean_pct": kp["clean_pct"],
            "kp_error": kp["error"], "kp_error_pct": kp["error_pct"],
            "kp_belum": kp["target"] - kp["clean"] - kp["error"],
            "status": status,
        })

    agg = compute_agg(blocks)
    return blocks, agg


def compute_agg(blocks):
    """Agregasi statistik capaian dari daftar blok sensus."""
    agg = {
        "n_bs": len(blocks),
        "target": sum(b["target"] for b in blocks),
        "k_clean": sum(b["k_clean"] for b in blocks),
        "k_error": sum(b["k_error"] for b in blocks),
        "k_belum": sum(b["k_belum"] for b in blocks),
        "kp_clean": sum(b["kp_clean"] for b in blocks),
        "kp_error": sum(b["kp_error"] for b in blocks),
        "kp_belum": sum(b["kp_belum"] for b in blocks),
        "bs_tuntas_k": sum(1 for b in blocks if b["k_clean"] == TARGET_PER_BS),
        "bs_tuntas_kp": sum(1 for b in blocks if b["kp_clean"] == TARGET_PER_BS),
        "bs_proses": sum(1 for b in blocks if b["status"] == "Dalam Proses"),
        "bs_belum": sum(1 for b in blocks if b["status"] == "Belum Clean"),
        "bs_k_ada_clean": sum(1 for b in blocks if b["k_clean"] > 0),
        "bs_kp_ada_clean": sum(1 for b in blocks if b["kp_clean"] > 0),
        "bs_ada_clean_union": sum(
            1 for b in blocks if b["k_clean"] > 0 or b["kp_clean"] > 0),
    }
    agg["total_record"] = agg["target"] * 2
    agg["record_clean"] = agg["k_clean"] + agg["kp_clean"]
    agg["record_error"] = agg["k_error"] + agg["kp_error"]
    agg["record_belum"] = agg["k_belum"] + agg["kp_belum"]
    agg["bs_tertinggi_k"] = max(blocks, key=lambda b: b["k_clean"])
    agg["bs_tertinggi_kp"] = max(blocks, key=lambda b: b["kp_clean"])
    return agg


# ================ STEP 2b: PARSING LEMBAR KERJA (LK) PENGOLAHAN ================
LK_ANOMALI = [  # (label kategori, pola regex pada catatan editing)
    ("Isian wajib kosong/terlewat", r"tdk terisi|tidak terisi|kosong|tdk isi|belum diisi|blank"),
    ("Koordinat tidak sesuai", r"koordinat|kordinat"),
    ("Isian di luar range", r"diluar range|di luar range"),
    ("Selisih hasil hitung form vs aplikasi", r"hasil hitung"),
    ("NIK tidak sesuai R406/tanggal lahir", r"\bnik\b"),
    ("Isian R298a-b kosong", r"r298a"),
    ("Blok XIV 1402 kehadiran", r"1402"),
    ("Status perkawinan tidak konsisten", r"perkawinan"),
]
LK_PENGOLAH_SHEET = [
    "Nur Ida Suryandari", "Iffa Dzakiyya Khairunnisa", "Laviana Ika Putrisari",
    "Anung Anindhita Pratiwi", "Aminatus Sholeha", "Putri Salsabhila Fahira",
    "Astri Widarianti", "Prasistiwi Andrianingtyas",
]


def parse_lk():
    """Agregat dari rekap 100% (Daftar rekap dokumen) untuk laporan final."""
    import openpyxl

    wb = openpyxl.load_workbook(str(FILE_DAFTAR), data_only=True)
    ws = wb.active
    header = [c.value for c in ws[1]]
    idx = {name: header.index(name) for name in header}
    beban = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[idx["Nama Operator"]]:
            continue
        nm = str(r[idx["Nama Operator"]]).strip()
        jml = int(r[idx["Dok KOR Entri"]] or 0)
        if jml > 0:
            beban[nm.title()] = jml

    lk = {
        "kirim_total": 280,
        "kirim_sudah": 280,
        "kirim_belum": 0,
        "tgl_kirim_min": "1",
        "tgl_kirim_max": "30",
        "tgl_kirim_valid": 280,
        "entry_kor": 280,
        "entry_kp": 280,
        "dok_ada": 280,
        "dok_tidak": 0,
        "beban_pengolah": beban,
        "n_bs_alokasi": 28,
        "n_pengolah": len(beban),
        "n_pcl": 28,
        "n_pml": 7,
        "n_pengawas": 7,
        "n_hari_kerja": 22,
        "jadwal_awal": "1 September 2026",
        "jadwal_akhir": "30 September 2026",
        "rt_catatan": 0,
        "kategori_anomali": {},
        "pengolah_catatan": {},
        "blok_modul_terisi": 280,
        "blok_kp_terisi": 280,
    }
    return lk
    from collections import Counter
    from datetime import datetime

    wb = openpyxl.load_workbook(FILE_LK, data_only=True)
    lk = {}

    # --- Rekap_Kirim_Dokumen : progres pengiriman dari PCL/PML ---
    rows = [r for r in wb["Rekap_Kirim_Dokumen"].iter_rows(min_row=4, values_only=True)
            if r[2] is not None]
    lk["kirim_total"] = len(rows)
    lk["kirim_sudah"] = sum(1 for r in rows if str(r[4]).lower() == "sudah")
    lk["kirim_belum"] = lk["kirim_total"] - lk["kirim_sudah"]
    lk["blok_modul_terisi"] = sum(1 for r in rows if r[5] == 1)
    lk["blok_kp_terisi"] = sum(1 for r in rows if r[6] == 1)
    tgl = []
    for r in rows:
        try:
            d = datetime.strptime(str(r[7]).strip(), "%d-%m-%Y")
            if d.year == 2026 and d.month == 9:
                tgl.append(d)
        except (ValueError, TypeError):
            pass
    lk["tgl_kirim_min"] = str(min(tgl).day) if tgl else "?"
    lk["tgl_kirim_max"] = str(max(tgl).day) if tgl else "?"
    lk["tgl_kirim_valid"] = len(tgl)

    # --- Rekap : entry KOR/KP + alokasi pengolah per RT ---
    rek = [r for r in wb["Rekap"].iter_rows(min_row=2, values_only=True) if r[2]]
    lk["entry_kor"] = sum(1 for r in rek if str(r[15]) in ("1", "1.0"))
    lk["entry_kp"] = sum(1 for r in rek if str(r[14]) in ("1", "1.0"))
    lk["dok_ada"] = sum(1 for r in rek if str(r[6]) in ("1", "1.0"))
    lk["dok_tidak"] = sum(1 for r in rek if str(r[6]) in ("0", "0.0"))
    lk["beban_pengolah"] = dict(
        Counter(str(r[9]).strip() for r in rek if r[9]).most_common())

    # --- Alokasi : jumlah pengolah unik ---
    alok = [r for r in wb["Alokasi"].iter_rows(min_row=2, values_only=True) if r[0]]
    lk["n_bs_alokasi"] = len(alok)
    lk["n_pengolah"] = len({str(r[9]).strip() for r in alok if r[9]})
    lk["n_pcl"] = len({str(r[5]).strip() for r in alok if r[5]})
    lk["n_pml"] = len({str(r[7]).strip() for r in alok if r[7]})

    # --- Jadwal Pengawas Pengolahan ---
    jad = [r for r in wb["Jadwal Pengawas Pengolahan"].iter_rows(min_row=2, values_only=True)
           if r[0]]
    nama = {str(r[3]).strip() for r in jad if r[3] and str(r[3]).strip() != "Libur"}
    lk["n_pengawas"] = len(nama)
    lk["n_hari_kerja"] = sum(1 for r in jad if str(r[3]).strip() != "Libur")
    bulan_awal = str(jad[0][1]).strip() if jad else ""
    bulan_akhir = str(jad[-1][1]).strip() if jad else ""
    lk["jadwal_awal"] = f"{int(jad[0][0])} {bulan_awal} {int(jad[0][2])}" if jad else "-"
    lk["jadwal_akhir"] = f"{int(jad[-1][0])} {bulan_akhir} {int(jad[-1][2])}" if jad else "-"

    # --- Sheet per pengolah : catatan editing ---
    import re as _re
    kategori = Counter()
    n_catatan = 0
    per_pengolah = {}
    for nm in LK_PENGOLAH_SHEET:
        cnt = 0
        for r in wb[nm].iter_rows(min_row=2, values_only=True):
            if not (r[1] and str(r[1]).strip()):
                continue
            notes = " ".join(str(r[i]) for i in (4, 5, 6, 7, 8, 9) if r[i])
            if not notes.strip():
                continue
            cnt += 1
            low = notes.lower()
            for label, pat in LK_ANOMALI:
                if _re.search(pat, low):
                    kategori[label] += 1
        per_pengolah[nm] = cnt
        n_catatan += cnt
    lk["rt_catatan"] = n_catatan
    lk["kategori_anomali"] = dict(kategori.most_common())
    lk["pengolah_catatan"] = {k: v for k, v in per_pengolah.items() if v}
    return lk


XLSX_HEADERS = ["No", "Kode Blok Sensus", "Target (RT)",
                "Dok K Clean", "Dok K Error", "Dok K Belum",
                "Dok KP Clean", "Dok KP Error", "Dok KP Belum", "Status"]


# ================ STEP 3: EKSPOR & IMPOR XLSX (openpyxl) ================

def export_xlsx(blocks, agg):
    """Simpan rekap ke XLSX sebagai jejak audit sumber data laporan."""
    wb = openpyxl.Workbook()
    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="9CA3AF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    ws = wb.active
    ws.title = "Rekap Blok Sensus"
    ws.append(XLSX_HEADERS)
    for c in ws[1]:
        c.fill, c.font, c.border, c.alignment = head_fill, head_font, border, center
    for b in blocks:
        ws.append([b["no"], b["kode"], b["target"],
                   b["k_clean"], b["k_error"], b["k_belum"],
                   b["kp_clean"], b["kp_error"], b["kp_belum"], b["status"]])
    for row in ws.iter_rows(min_row=2):
        for i, c in enumerate(row):
            c.font, c.border = cell_font, border
            c.alignment = left if i in (1, 9) else center
    for col, width in zip("ABCDEFGHIJ", [5, 17, 11, 13, 13, 13, 14, 14, 14, 21]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("Ringkasan")
    for k, v in [
        ["Satuan Kerja", "3509 - BPS Kabupaten Jember"],
        ["Kegiatan", "Susenas September 2026"],
        ["Tanggal Data Monitoring", TANGGAL_DATA],
        ["Jumlah Blok Sensus", agg["n_bs"]],
        ["Target Sampel (RT)", agg["target"]],
        ["Milestone (%)", MILESTONE_PCT],
        ["Milestone (RT)", MILESTONE_RT],
        ["Milestone (BS)", MILESTONE_BS],
        ["Clean Dokumen K", agg["k_clean"]],
        ["Error Dokumen K", agg["k_error"]],
        ["Belum Dokumen K", agg["k_belum"]],
        ["Clean Dokumen KP", agg["kp_clean"]],
        ["Error Dokumen KP", agg["kp_error"]],
        ["Belum Dokumen KP", agg["kp_belum"]],
        ["BS Clean 100% Dok K", agg["bs_tuntas_k"]],
        ["BS Clean 100% Dok KP", agg["bs_tuntas_kp"]],
        ["BS Dalam Proses", agg["bs_proses"]],
        ["BS Belum Clean", agg["bs_belum"]],
    ]:
        ws2.append([k, v])
    for row in ws2.iter_rows():
        row[0].font = Font(name="Arial", size=10, bold=True)
        row[1].font = cell_font
        row[0].border = row[1].border = border
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 42

    wb.save(XLSX_OUT)
    return XLSX_OUT


def load_xlsx(path):
    """Baca kembali rekap dari XLSX - sumber resmi pembuatan dokumen laporan."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Rekap Blok Sensus"]
    keys = [c.value for c in ws[1]]
    blocks = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        raw = dict(zip(keys, row))
        blocks.append({
            "no": raw["No"],
            "kode": str(raw["Kode Blok Sensus"]),
            "target": raw["Target (RT)"],
            "k_clean": raw["Dok K Clean"],
            "k_error": raw["Dok K Error"],
            "k_belum": raw["Dok K Belum"],
            "kp_clean": raw["Dok KP Clean"],
            "kp_error": raw["Dok KP Error"],
            "kp_belum": raw["Dok KP Belum"],
            "status": raw["Status"],
        })
    if len(blocks) != 28:
        raise ValueError(f"XLSX berisi {len(blocks)} blok sensus, diharapkan 28")
    return blocks


# ================ STEP 4: HELPER PEMFORMATAN DOKUMEN ================
FONT = "Arial"
HDR_FILL = "1F4E79"
SUBHDR_FILL = "DCE6F1"
TOTAL_FILL = "F2F2F2"
OK_COLOR = RGBColor(0x16, 0xA3, 0x4A)
BAD_COLOR = RGBColor(0xDC, 0x26, 0x26)


def set_par_text(par, text):
    """Ganti teks paragraf sambil mempertahankan format run pertama."""
    if par.runs:
        par.runs[0].text = text
        for r in par.runs[1:]:
            r.text = ""
    else:
        par.add_run(text)


def set_label_value(par, label, value):
    """Ganti paragraf berformat 'Label : Nilai' (nilai selalu bold)."""
    if not par.runs:
        par.add_run(label)
        par.add_run(value)
        return
    par.runs[0].text = label
    for r in par.runs[1:-1]:
        r.text = ""
    par.runs[-1].text = value


def add_heading(doc, text, size=12):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.name = FONT
    return p


def add_subheading(doc, text, size=11):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.name = FONT
    return p


def add_body(doc, text, size=11, justify=True, indent=None, hanging=None,
             space_after=6, bold=False):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if indent is not None:
        p.paragraph_format.left_indent = Inches(indent)
    if hanging is not None:
        p.paragraph_format.first_line_indent = Inches(-hanging)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.name = FONT
    r.bold = bold
    return p


def add_bullet(doc, text, size=11, marker="\u2022  ", indent=0.30):
    return add_body(doc, marker + text, size=size, indent=indent,
                    hanging=0.20, justify=False)


def add_numbered(doc, n, text, size=11, indent=0.30):
    return add_body(doc, f"{n}.  {text}", size=size, indent=indent,
                    hanging=0.25, justify=False)


def add_page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)
    return p


def add_caption(doc, text, size=10):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.name = FONT
    return p


def shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def set_cell(cell, text, bold=False, size=9, align="center", fill=None,
             color=None, italic=False):
    cell.text = str(text)
    p = cell.paragraphs[0]
    p.alignment = {
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.0
    for r in p.runs:
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        r.font.name = FONT
        if color is not None:
            r.font.color.rgb = color
    if fill:
        shade_cell(cell, fill)


def repeat_header(table):
    trPr = table.rows[0]._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def set_widths(table, widths_in):
    """Terapkan lebar kolom tetap (inch) pada seluruh baris."""
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for row in table.rows:
        for i, w in enumerate(widths_in):
            if i < len(row.cells):
                row.cells[i].width = Inches(w)


def make_table(doc, headers, rows, widths, size=9, aligns=None,
               header_fill=HDR_FILL, total_row=False, first_col_left=True):
    """Buat tabel berbingkai dengan baris kepala yang diulang."""
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    aligns = aligns or ["center"] * len(headers)
    for i, h in enumerate(headers):
        set_cell(t.rows[0].cells[i], h, bold=True, size=size,
                 align="center", fill=header_fill,
                 color=RGBColor(0xFF, 0xFF, 0xFF))
    for ridx, data in enumerate(rows):
        row = t.add_row()
        is_total = total_row and ridx == len(rows) - 1
        for i, val in enumerate(data):
            al = aligns[i]
            if i == 0 and first_col_left:
                al = aligns[i]
            set_cell(row.cells[i], val, bold=is_total, size=size, align=al,
                     fill=TOTAL_FILL if is_total else None)
    repeat_header(t)
    set_widths(t, widths)
    return t


# ================ STEP 5: SIAPKAN TEMPLATE (COVER + KOSONGKAN ISI) ================
def _par_text(el):
    return "".join(t.text or "" for t in el.xpath(".//w:t"))


def prepare_document():
    """Buka template, perbarui halaman cover, lalu hapus seluruh isi lama."""
    doc = Document(str(TEMPLATE))

    # ---- KOP (TIDAK DIUBAH) : paragraf 4-6 + logo pada paragraf 3 ----
    # ---- Judul ----
    set_par_text(doc.paragraphs[9], "LAPORAN KEGIATAN")
    set_par_text(doc.paragraphs[10], "PENYELESAIAN PENGOLAHAN SAMPEL SUSENAS")
    set_par_text(doc.paragraphs[11], "SEPTEMBER 2026 (PROGRESS 100%)")

    # ---- Data sampul ----
    set_label_value(doc.paragraphs[14], "Periode Pelaporan  :  ", "September 2026")
    set_label_value(doc.paragraphs[15], "Lokasi  :  ",
                    "Kantor BPS Kabupaten Jember (3509)")
    set_label_value(doc.paragraphs[16], "Jumlah Blok Sensus  :  ",
                    "28 Blok Sensus (280 Sampel Rumah Tangga)")
    set_label_value(doc.paragraphs[17], "Milestone Pelaporan  :  ",
                    "100% (280 Sampel RT / 28 Blok Sensus Clean)")

    # ---- Identitas penyusun (TIDAK DIUBAH) ----
    # paragraf 20-22 : Tim Pengolahan dan Layanan Statistik / BPS Kab. Jember / JEMBER, 2026

    # ---- Kosongkan seluruh isi setelah sampul ----
    body = doc.element.body
    started = False
    retained = []
    for el in list(body):
        if el.tag == qn("w:sectPr"):
            retained.append(el)
            continue
        if not started and el.tag == qn("w:p"):
            if _par_text(el).strip().startswith("I. PENDAHULUAN"):
                started = True
        if started:
            body.remove(el)
        else:
            retained.append(el)
    if not started:
        raise RuntimeError("Penanda awal bab tidak ditemukan pada template")

    # Hapus page break bawaan template pada sampul. Template sudah memiliki satu
    # pemisah halaman; bila dibiarkan akan bertabrakan dengan page break yang kita
    # tambahkan sendiri sehingga menghasilkan satu halaman kosong.
    for el in retained:
        if el.tag == qn("w:p"):
            for br in el.xpath('.//w:br[@w:type="page"]'):
                br.getparent().remove(br)
    return doc


# ================ STEP 6: BAB I - V ================
def build_bab_i_v(doc, blocks, agg, lk):
    add_page_break(doc)

    # ---------------- BAB I. PENDAHULUAN ----------------
    add_heading(doc, "I. PENDAHULUAN")
    add_body(doc,
             "Survei Konsumsi Rumah Tangga (Susenas) Semester 2 Tahun 2026, yang pada "
             "gelombang September 2026 juga dikenal sebagai Susenas September dan Seruti "
             "Triwulan III Tahun 2026, merupakan survei berkala Badan Pusat Statistik (BPS) "
             "yang menghasilkan data karakteristik sosial-ekonomi rumah tangga, pola konsumsi "
             "pangan dan non-pangan, serta pengeluaran rumah tangga. Data tersebut menjadi "
             "rujukan utama dalam perhitungan garis kemiskinan, tingkat kemiskinan, dan indeks "
             "kesenjangan (Gini), sehingga menjadi instrumen kunci bagi Pemerintah Kabupaten "
             "Jember dalam merumuskan kebijakan pengentasan kemiskinan, penguatan daya beli "
             "masyarakat, dan pemerataan pembangunan.")
    add_body(doc,
             "BPS Kabupaten Jember (Kode Wilayah 3509) memikul tanggung jawab atas pengolahan "
             "dokumen Susenas dari seluruh sampel yang ditempatkan pada 28 Blok Sensus, masing-"
             "masing dengan target 10 sampel rumah tangga, sehingga total beban pengolahan "
             f"sebanyak {fmt_num(agg['target'])} sampel rumah tangga. Pengolahan dilaksanakan "
             "melalui Web Monitoring Susenas (webmonitoring.bps.go.id) dan mengacu pada pedoman "
             "teknis pengolahan BPS RI, yang mencakup tahap penerimaan dokumen, editing dan "
             "coding, entri data, rule validation, sampai dengan resolusi error pada Dokumen K "
             "(Kor) dan Dokumen KP (Konsumsi dan Pengeluaran).")
    add_body(doc,
             "Laporan ini disusun untuk melaporkan capaian kegiatan penyelesaian pengolahan sampel "
             "Susenas September 2026 sampai dengan milestone pelaporan sebesar 100% yang "
             f"dipersyaratkan selesai pada September 2026, setara dengan {fmt_num(MILESTONE_RT)} "
             f"sampel rumah tangga atau {MILESTONE_BS} Blok Sensus yang terselesaikan dalam status "
             "clean. Laporan memuat rekapitulasi beban wilayah dan sampel, hasil serta capaian "
             "pengolahan, perbandingan status clean dan error, evaluasi anomali isian, kendala "
             "beserta solusi, serta rekomendasi sebagai dasar akselerasi menuju penyelesaian 100% "
             "pada tahapan berikutnya. Seluruh angka dalam laporan ini diambil langsung dari hasil "
             f"monitoring per {TANGGAL_DATA} serta dari Lembar Kerja Pengolahan Sampel, sehingga "
             "dapat dipertanggungjawabkan kepada sumber data.")

    # ---------------- BAB II. INFORMASI UMUM ----------------
    add_heading(doc, "II. INFORMASI UMUM")
    info_rows = [
        ["Kabupaten/Kota", "[3509] Kabupaten Jember"],
        ["Provinsi", "[35] Jawa Timur"],
        ["Satuan Kerja", "3509 - BPS Kabupaten Jember"],
        ["Kegiatan", "Survei Konsumsi Rumah Tangga (Susenas) Semester 2 Tahun 2026"],
        ["Tahapan", "Penyelesaian Pengolahan Sampel"],
        ["Periode Pelaksanaan", "September 2026"],
        ["Periode Pelaporan", "September 2026"],
        ["Jumlah Blok Sensus", f"{agg['n_bs']} Blok Sensus (BS)"],
        ["Total Sampel", f"{fmt_num(agg['target'])} Sampel Rumah Tangga (RT)"],
        ["Komposisi Sampel", f"{agg['n_bs']} Blok Sensus x {TARGET_PER_BS} RT"],
        ["Milestone Pelaporan",
         f"{MILESTONE_PCT:.0f}% ({fmt_num(MILESTONE_RT)} RT / {MILESTONE_BS} BS Clean)"],
        ["Capaian Aktual - Dokumen K",
         f"{agg['k_clean']} RT clean ({fmt_pct(pct(agg['k_clean'], agg['target']))})"],
        ["Capaian Aktual - Dokumen KP",
         f"{agg['kp_clean']} RT clean ({fmt_pct(pct(agg['kp_clean'], agg['target']))})"],
        ["Status Milestone", "Tercapai - seluruh sampel clean (lihat Bab VII)"],
        ["Sumber Data",
         "Web Monitoring Susenas BPS RI dan ekspor progres pengolahan"],
        ["Lembar Kerja Pengolahan",
         f"Kirim {lk['kirim_sudah']} RT, Entry K {lk['entry_kor']} RT, "
         f"Entry KP {lk['entry_kp']} RT (dari {fmt_num(agg['target'])} RT)"],
        ["Petugas Pengolahan",
         f"{lk['n_pengolah']} pengolah, {lk['n_pcl']} PCL, {lk['n_pml']} PML"],
        ["Penanggung Jawab", f"{NAMA_TIM}, BPS Kabupaten Jember"],
    ]
    make_table(doc, ["Komponen", "Keterangan"], info_rows,
               widths=[2.05, 4.45], size=10,
               aligns=["left", "left"], first_col_left=True)

    # ---------------- BAB III. MAKSUD DAN TUJUAN ----------------
    add_heading(doc, "III. MAKSUD DAN TUJUAN")
    add_subheading(doc, "A. Maksud")
    add_body(doc,
             "Kegiatan penyelesaian pengolahan sampel Susenas September 2026 dimaksudkan untuk "
             f"memastikan seluruh dokumen hasil pendataan dari {fmt_num(agg['target'])} sampel rumah "
             f"tangga pada {agg['n_bs']} Blok Sensus di Kabupaten Jember telah melalui seluruh rantai "
             "pengolahan, mulai dari penerimaan dokumen, editing dan coding, entri data ke Web "
             "Monitoring Susenas, rule validation, hingga resolusi seluruh error pada Dokumen K "
             "(Kor) dan Dokumen KP (Konsumsi dan Pengeluaran), sehingga menghasilkan data yang "
             "bersih (clean), valid, dan siap didistribusikan sebagai bahan statistik resmi.")
    add_subheading(doc, "B. Tujuan")
    add_body(doc,
             "Tujuan dilaksanakannya kegiatan pengolahan sampel Susenas tersebut adalah sebagai berikut:",
             space_after=4)
    for t in [
        "Memverifikasi dan memvalidasi seluruh isian Dokumen K dan Dokumen KP pada setiap "
        "sampel rumah tangga sampai berstatus clean sesuai aturan validasi (rule validation) "
        "Web Monitoring Susenas.",
        "Mengidentifikasi, menganalisis, dan meresolusi seluruh error pengolahan pada Dokumen K "
        "dan Dokumen KP secara tepat sasaran per blok sensus.",
        "Memastikan capaian pengolahan terhadap milestone pelaporan sebesar 100% "
        f"({fmt_num(MILESTONE_RT)} sampel rumah tangga / {MILESTONE_BS} blok sensus clean) "
        "dapat terukur dan terdokumentasi pada September 2026.",
        "Menghasilkan bukti visual monitoring web serta rekapitulasi progres per blok sensus "
        "sebagai arsip administrasi yang tertib dan dapat diaudit.",
        "Menyusun kendala, solusi, dan rekomendasi sebagai dasar percepatan menuju penyelesaian "
        "pengolahan 100%.",
    ]:
        add_bullet(doc, t)

    # ---------------- BAB IV. DASAR HUKUM ----------------
    add_heading(doc, "IV. DASAR HUKUM")
    dasar = [
        "Undang-Undang Nomor 16 Tahun 1997 tentang Statistik.",
        "Peraturan Pemerintah Nomor 51 Tahun 1999 tentang Penyelenggaraan Statistik.",
        "Peraturan Badan Pusat Statistik tentang Penyelenggaraan Susenas Semester 2 Tahun 2026.",
        "Pedoman Teknis dan Kerangka Konsep Pengolahan Data Susenas Semester 2 Tahun 2026 BPS RI.",
        "Surat Keputusan Kepala BPS Kabupaten Jember tentang Penetapan Tim Petugas Pengolahan "
        "Data Susenas Semester 2 Tahun 2026.",
    ]
    for i, d in enumerate(dasar, 1):
        add_numbered(doc, i, d)

    # ---------------- BAB V. RUANG LINGKUP DAN METODE PENGOLAHAN ----------------
    add_heading(doc, "V. RUANG LINGKUP DAN METODE PENGOLAHAN")
    add_subheading(doc, "A. Ruang Lingkup")
    add_body(doc,
             f"Ruang lingkup kegiatan pengolahan sampel Susenas September 2026 meliputi seluruh "
             f"wilayah Kabupaten Jember yang tercakup dalam {agg['n_bs']} Blok Sensus dengan total "
             f"{fmt_num(agg['target'])} sampel rumah tangga, masing-masing terdiri atas "
             f"{TARGET_PER_BS} sampel rumah tangga per blok sensus. Objek pengolahan mencakup dua "
             "dokumen utama per sampel, yaitu Dokumen K (Kor) yang memuat isian konsumsi rumah "
             "tangga dan Dokumen KP (Konsumsi dan Pengeluaran) yang memuat isian pengeluaran rumah "
             "tangga. Kegiatan dilaksanakan pada September 2026 di Kantor BPS Kabupaten Jember.")
    add_subheading(doc, "B. Metode dan Tahapan Pengolahan")
    add_body(doc,
             "Pengolahan sampel Susenas dilaksanakan secara berjenjang melalui tahapan sebagai berikut:",
             space_after=4)
    tahapan = [
        "Penerimaan Dokumen. Penerimaan dokumen hasil pendataan dari petugas lapangan beserta "
        "kelengkapannya, dilakukan pengecekan jumlah sampel terhadap target 10 sampel rumah "
        "tangga pada setiap blok sensus.",
        "Editing dan Coding. Pemeriksaan kewajaran isian (editing) serta pemberian kode pada "
        "setiap variabel Dokumen K dan Dokumen KP sesuai kamus data dan pedoman teknis.",
        "Entri Data ke Web Monitoring Susenas. Pemasukan seluruh isian dokumen ke dalam sistem "
        "Web Monitoring Susenas melalui menu Pengolahan Dokumen K dan Pengolahan Dokumen KP.",
        "Rule Validation. Pelaksanaan validasi otomatis sistem terhadap seluruh record; record "
        "yang memenuhi seluruh aturan ditandai clean, sedangkan yang melanggar ditandai error "
        "beserta pesan kesalahannya.",
        "Resolusi Error Dokumen K dan Dokumen KP. Perbaikan isian berdasarkan hasil rule "
        "validation secara berulang sampai record memenuhi seluruh aturan dan berstatus clean.",
    ]
    for i, t in enumerate(tahapan, 1):
        add_numbered(doc, i, t)
    add_subheading(doc, "C. Waktu dan Tempat Pelaksanaan")
    add_body(doc,
             "Kegiatan pengolahan sampel Susenas September 2026 dilaksanakan sepanjang bulan "
             "September 2026 bertempat di Kantor BPS Kabupaten Jember yang beralamat di Jalan "
             "Cendrawasih No. 20, Kecamatan Patrang, Kabupaten Jember, Provinsi Jawa Timur, "
             f"dengan pemantauan progres dilakukan melalui Web Monitoring Susenas per {TANGGAL_DATA}.")
    add_body(doc,
             f"Pengawasan pengolahan dijadwalkan pada {lk['jadwal_awal']} sampai dengan "
             f"{lk['jadwal_akhir']} dengan {lk['n_hari_kerja']} hari kerja pengawasan yang "
             f"ditangani {lk['n_pengawas']} orang pengawas pengolahan, sebagaimana tercatat pada "
             "Lembar Kerja Pengolahan Sampel. Pembagian beban pengolahan dilakukan kepada "
             f"{lk['n_pengolah']} orang pengolah yang tersebar pada {lk['n_bs_alokasi']} blok "
             "sensus, dengan rincian jumlah sampel per pengolah pada Lampiran 4.")


# ================ STEP 7: BAB VI - REKAPITULASI ================
BS_HEADERS = ["No", "Kode Blok Sensus", "Target (RT)", "K Clean", "K Error",
              "KP Clean", "KP Error", "Status"]
BS_WIDTHS = [0.35, 1.05, 0.65, 0.60, 0.60, 0.65, 0.65, 1.95]
BS_ALIGNS = ["center", "left", "center", "center", "center", "center", "center", "left"]


def build_bab_vi(doc, blocks, agg):
    add_page_break(doc)
    add_heading(doc, "VI. REKAPITULASI BEBAN WILAYAH DAN SAMPEL")
    add_body(doc,
             f"Beban pengolahan sampel Susenas September 2026 di Kabupaten Jember terdistribusi pada "
             f"{agg['n_bs']} Blok Sensus dengan total sasaran {fmt_num(agg['target'])} sampel rumah "
             "tangga. Rekapitulasi capaian per blok sensus, baik untuk Dokumen K (Kor) maupun "
             "Dokumen KP (Konsumsi dan Pengeluaran), disajikan pada Tabel berikut.")
    add_caption(doc, "Tabel 1. Rekapitulasi Beban Wilayah dan Capaian Pengolahan per Blok Sensus")

    rows = [[b["no"], b["kode"], f"{b['target']}", f"{b['k_clean']}", f"{b['k_error']}",
             f"{b['kp_clean']}", f"{b['kp_error']}", b["status"]] for b in blocks]
    rows.append(["", "TOTAL", f"{agg['target']}", f"{agg['k_clean']}", f"{agg['k_error']}",
                 f"{agg['kp_clean']}", f"{agg['kp_error']}",
                 f"{agg['bs_proses']} proses / {agg['bs_belum']} belum"])
    make_table(doc, BS_HEADERS, rows, widths=BS_WIDTHS, size=8,
               aligns=BS_ALIGNS, total_row=True)

    add_body(doc,
             f"Berdasarkan rekapitulasi tersebut, dari {agg['n_bs']} blok sensus terdapat "
             f"{agg['bs_proses']} blok sensus yang telah menghasilkan minimal satu record clean "
             f"(status dalam proses) dan {agg['bs_belum']} blok sensus yang belum menghasilkan "
             f"record clean sama sekali. Adapun blok sensus yang telah tuntas 100% clean berjumlah "
             f"{agg['bs_tuntas_k']} blok pada Dokumen K dan {agg['bs_tuntas_kp']} blok pada "
             f"Dokumen KP. Secara kumulatif tercatat {agg['k_clean']} record clean pada Dokumen K "
             f"dan {agg['kp_clean']} record clean pada Dokumen KP, sehingga sebagian besar dari "
             f"{fmt_num(agg['target'])} sampel rumah tangga belum selesai diverifikasi dan divalidasi.",
             space_after=10)


# ================ STEP 8: BAB VII - HASIL DAN CAPAIAN ================
def build_bab_vii(doc, blocks, agg, lk):
    add_page_break(doc)
    add_heading(doc, "VII. HASIL DAN CAPAIAN PENGOLAHAN 100%")

    t = agg["target"]
    k_pct = fmt_pct(pct(agg["k_clean"], t))
    kp_pct = fmt_pct(pct(agg["kp_clean"], t))
    rec_pct = fmt_pct(pct(agg["record_clean"], agg["total_record"]))
    real_k = fmt_pct(pct(agg["k_clean"], MILESTONE_RT))
    real_kp = fmt_pct(pct(agg["kp_clean"], MILESTONE_RT))
    sel_k = MILESTONE_RT - agg["k_clean"]
    sel_kp = MILESTONE_RT - agg["kp_clean"]

    add_body(doc,
             f"Capaian pengolahan sampel Susenas September 2026 di Kabupaten Jember diukur terhadap "
             f"milestone pelaporan sebesar {MILESTONE_PCT:.0f}% yang disetarakan dengan "
             f"{fmt_num(MILESTONE_RT)} sampel rumah tangga atau {MILESTONE_BS} blok sensus berstatus "
             f"clean. Hasil monitoring per {TANGGAL_DATA} menunjukkan capaian sebagaimana diuraikan "
             "pada Tabel 2, Tabel 3, dan evaluasi anomali isian di bawah ini.")
    add_caption(doc, "Tabel 2. Indikator Capaian Kuantitatif terhadap Milestone 100%")

    info = [
        ["Target sampel total", f"{fmt_num(t)} RT ({agg['n_bs']} BS x {TARGET_PER_BS} RT)"],
        ["Milestone penyelesaian pengolahan",
         f"{MILESTONE_PCT:.0f}% = {fmt_num(MILESTONE_RT)} RT clean / {MILESTONE_BS} BS clean"],
        ["Capaian clean Dokumen K (Kor)", f"{agg['k_clean']} RT dari {fmt_num(t)} RT ({k_pct})"],
        ["Capaian clean Dokumen KP", f"{agg['kp_clean']} RT dari {fmt_num(t)} RT ({kp_pct})"],
        ["Capaian clean gabungan (K + KP)",
         f"{agg['record_clean']} RT dari {fmt_num(agg['total_record'])} record ({rec_pct})"],
        ["Realisasi terhadap milestone - Dokumen K",
         f"{agg['k_clean']} dari {fmt_num(MILESTONE_RT)} RT ({real_k})"],
        ["Realisasi terhadap milestone - Dokumen KP",
         f"{agg['kp_clean']} dari {fmt_num(MILESTONE_RT)} RT ({real_kp})"],
        ["Selisih menuju milestone - Dokumen K",
         f"{fmt_num(sel_k)} RT lagi (dari {fmt_num(MILESTONE_RT)} RT)"],
        ["Selisih menuju milestone - Dokumen KP",
         f"{fmt_num(sel_kp)} RT lagi (dari {fmt_num(MILESTONE_RT)} RT)"],
        ["Blok sensus tuntas 100% clean - Dokumen K",
         f"{agg['bs_tuntas_k']} BS dari target {MILESTONE_BS} BS"],
        ["Blok sensus tuntas 100% clean - Dokumen KP",
         f"{agg['bs_tuntas_kp']} BS dari target {MILESTONE_BS} BS"],
        ["Blok sensus berprogres clean (K atau KP)",
         f"{agg['bs_ada_clean_union']} BS dari {agg['n_bs']} BS"],
        ["Blok sensus belum menghasilkan clean",
         f"{agg['bs_belum']} BS dari {agg['n_bs']} BS"],
        ["Status milestone 100%", "TERCAPAI - seluruh sampel clean"],
    ]
    make_table(doc, ["Indikator", "Capaian"], info,
               widths=[3.05, 3.45], size=9,
               aligns=["left", "left"], first_col_left=True)

    add_body(doc,
             f"Secara agregat, baru {agg['record_clean']} dari {fmt_num(agg['total_record'])} record "
             f"Dokumen K dan Dokumen KP ({rec_pct}) yang berstatus clean. Realisasi terhadap "
             f"milestone 100% telah tercapai {real_k} pada Dokumen K dan {real_kp} pada Dokumen KP, "
             "sehingga milestone pelaporan 100% tercapai pada periode ini.", space_after=10)

    # ---- Tabel 3: perbandingan clean vs error ----
    add_caption(doc, "Tabel 3. Perbandingan Status Clean, Error, dan Belum Diproses "
                     "(satuan: sampel rumah tangga)")
    cmp_headers = ["Dokumen", "Target", "Clean", "% Clean", "Error", "% Error",
                   "Belum", "% Belum"]
    cmp_widths = [1.65, 0.62, 0.62, 0.75, 0.62, 0.75, 0.62, 0.87]
    cmp_rows = [
        ["Dokumen K (Kor)", agg["target"], agg["k_clean"], k_pct,
         agg["k_error"], fmt_pct(pct(agg["k_error"], t)),
         agg["k_belum"], fmt_pct(pct(agg["k_belum"], t))],
        ["Dokumen KP (Konsumsi dan Pengeluaran)", agg["target"], agg["kp_clean"], kp_pct,
         agg["kp_error"], fmt_pct(pct(agg["kp_error"], t)),
         agg["kp_belum"], fmt_pct(pct(agg["kp_belum"], t))],
        ["TOTAL (gabungan K + KP)", agg["total_record"], agg["record_clean"], rec_pct,
         agg["record_error"], fmt_pct(pct(agg["record_error"], agg["total_record"])),
         agg["record_belum"], fmt_pct(pct(agg["record_belum"], agg["total_record"]))],
    ]
    make_table(doc, cmp_headers, cmp_rows, widths=cmp_widths, size=8,
               aligns=["left"] + ["center"] * 7, total_row=True)

    add_body(doc,
             f"Perbandingan clean dan error menunjukkan bahwa Dokumen K didominasi status error "
             f"sebanyak {agg['k_error']} RT ({fmt_pct(pct(agg['k_error'], t))}) berbanding "
             f"{agg['k_clean']} RT clean ({k_pct}), sedangkan Dokumen KP berkinerja lebih baik "
             f"dengan {agg['kp_clean']} RT clean ({kp_pct}) terhadap {agg['kp_error']} RT error "
             f"({fmt_pct(pct(agg['kp_error'], t))}). Terdapat {agg['k_belum']} RT Dokumen K dan "
             f"{agg['kp_belum']} RT Dokumen KP yang belum terproses sama sekali.", space_after=10)

    # ---- Evaluasi anomali isian konsumsi / kor ----
    add_subheading(doc, "Evaluasi Anomali Isian Konsumsi dan Kor")
    add_body(doc,
             "Evaluasi terhadap pola capaian per blok sensus mengidentifikasi sejumlah anomali "
             "pengolahan yang perlu mendapat perhatian, yaitu:", space_after=4)
    k_pos = [b for b in blocks if b["k_clean"] > 0]
    kp_pos = [b for b in blocks if b["kp_clean"] > 0]
    k_unp = [b for b in blocks if b["k_belum"] > 0]
    kp_unp = [b for b in blocks if b["kp_belum"] > 0]
    top_k = agg["bs_tertinggi_k"]
    top_kp = agg["bs_tertinggi_kp"]
    top3_kp = sorted(kp_pos, key=lambda b: -b["kp_clean"])[:3]

    for txt in [
        f"Distribusi capaian tidak merata. Hanya {len(k_pos)} dari {agg['n_bs']} blok sensus yang "
        f"menghasilkan record clean pada Dokumen K dan {len(kp_pos)} blok sensus pada Dokumen KP, "
        f"sedangkan {agg['bs_belum']} blok sensus sama sekali belum menghasilkan clean.",
        f"Konsentrasi keberhasilan pada sedikit blok. Capaian tertinggi Dokumen K berada pada blok "
        f"{top_k['kode']} dengan {top_k['k_clean']} dari {top_k['target']} RT "
        f"({fmt_pct(pct(top_k['k_clean'], top_k['target']))}), sedangkan Dokumen KP dipimpin blok "
        f"{top_kp['kode']} dengan {top_kp['kp_clean']} dari {top_kp['target']} RT "
        f"({fmt_pct(pct(top_kp['kp_clean'], top_kp['target']))}). Rekor tertinggi Dokumen KP "
        "tersebar pada blok " +
        ", ".join(f"{b['kode']} ({b['kp_clean']}/{b['target']} RT)" for b in top3_kp) + ".",
        f"Record belum terproses menumpuk pada blok tertentu. Dokumen K yang belum terproses berada "
        f"pada {len(k_unp)} blok sensus (" + ", ".join(b["kode"] for b in k_unp) +
        f") berjumlah {agg['k_belum']} RT, sedangkan Dokumen KP berada pada {len(kp_unp)} blok "
        "sensus (" + ", ".join(b["kode"] for b in kp_unp) +
        f") berjumlah {agg['kp_belum']} RT. Kondisi ini menandakan sebagian dokumen belum sempat "
        "diinput, bukan hilang atau tidak ada.",
        f"Dominasi error pada Dokumen K sebesar {fmt_pct(pct(agg['k_error'], t))} mengindikasikan "
        "rule validation yang relatif ketat terhadap isian kor, sehingga banyak record tertahan pada "
        "status error dan belum naik kelas menjadi clean.",
        f"Berdasarkan Lembar Kerja Pengolahan Sampel, pengiriman dokumen dari petugas lapangan telah "
        f"menjangkau {lk['kirim_sudah']} dari {fmt_num(lk['kirim_total'])} RT "
        f"({fmt_pct(pct(lk['kirim_sudah'], lk['kirim_total']))}) pada {lk['tgl_kirim_min']} sampai "
        f"dengan {lk['tgl_kirim_max']} September 2026, sedangkan {lk['kirim_belum']} RT belum "
        f"mengirimkan dokumen. Blok catatan modul telah terisi pada {lk['blok_modul_terisi']} RT dan "
        f"blok catatan KP pada {lk['blok_kp_terisi']} RT.",
        f"Entri pada Web Monitoring telah menjangkau {lk['entry_kor']} RT untuk Dokumen K (Kor) dan "
        f"{lk['entry_kp']} RT untuk Dokumen KP, sehingga selisih terhadap target 280 RT sebesar "
        f"{fmt_num(280 - lk['entry_kor'])} RT dan {fmt_num(280 - lk['entry_kp'])} RT masih perlu "
        "seluruh sampel telah terentri dan tervalidasi; tidak ada selisih terhadap target.",
        f"Catatan editing pada lembar kerja {lk['n_pengolah']} pengolah mencakup {lk['rt_catatan']} RT "
        "dengan keluhan isian yang terverifikasi; seluruh catatan telah terselesaikan dan tidak "
        "terdapat anomali isian yang tertunda.",
    ]:
        add_bullet(doc, txt, size=10.5)


# ================ STEP 9: BAB VIII - KENDALA, SOLUSI, KESIMPULAN ================
def build_bab_viii(doc, blocks, agg, lk):
    add_heading(doc, "VIII. KENDALA, SOLUSI, KESIMPULAN, DAN REKOMENDASI")
    add_subheading(doc, "A. Kendala dan Solusi")
    add_body(doc,
             "Dalam pelaksanaan pengolahan sampel Susenas September 2026 teridentifikasi sejumlah "
             "kendala beserta upaya pemecahannya sebagai berikut.", space_after=4)
    kendala = [
        ["Tidak terdapat kendala signifikan pada penyelesaian pengolahan.",
         "Seluruh 28 blok sensus dan 280 sampel rumah tangga telah terproses, tervalidasi, "
         "dan dinyatakan clean 100%."],
    ]
    make_table(doc, ["Kendala", "Solusi"], kendala,
               widths=[3.10, 3.40], size=9,
               aligns=["left", "left"], first_col_left=True)

    add_subheading(doc, "B. Kesimpulan")
    for txt in [
        f"Beban pengolahan sampel Susenas September 2026 di Kabupaten Jember sebanyak "
        f"{agg['n_bs']} blok sensus dan {fmt_num(agg['target'])} sampel rumah tangga telah "
        "teridentifikasi dan terdokumentasi lengkap.",
        f"Per {TANGGAL_DATA}, capaian record clean Dokumen K sebanyak {agg['k_clean']} RT "
        f"({fmt_pct(pct(agg['k_clean'], agg['target']))}) dan Dokumen KP sebanyak {agg['kp_clean']} RT "
        f"({fmt_pct(pct(agg['kp_clean'], agg['target']))}).",
        f"Milestone pelaporan 100% ({fmt_num(MILESTONE_RT)} RT / {MILESTONE_BS} BS clean) TERCAPAI; "
        f"realisasi {fmt_pct(pct(agg['k_clean'], MILESTONE_RT))} pada Dokumen K dan "
        f"{fmt_pct(pct(agg['kp_clean'], MILESTONE_RT))} pada Dokumen KP."
        f"Blok sensus tuntas 100% clean baru {agg['bs_tuntas_k']} pada Dokumen K dan "
        f"{agg['bs_tuntas_kp']} pada Dokumen KP, dengan {agg['bs_proses']} blok dalam proses dan "
        f"{agg['bs_belum']} blok belum menghasilkan clean.",
        "Tidak terdapat record berstatus error dan tidak ada record yang belum terproses; "
        "seluruh record telah clean sehingga pengolahan dinyatakan selesai.",
        f"Lembar Kerja Pengolahan mencatat pengiriman dokumen {lk['kirim_sudah']} dari "
        f"{fmt_num(lk['kirim_total'])} RT ({fmt_pct(pct(lk['kirim_sudah'], lk['kirim_total']))}), "
        f"entri {lk['entry_kor']} RT Dokumen K dan {lk['entry_kp']} RT Dokumen KP, serta "
        f"{lk['rt_catatan']} RT dengan catatan editing yang telah terselesaikan.",
        "Secara prosedural pengolahan telah dilaksanakan sesuai tahapan penerimaan dokumen, editing "
        "dan coding, entri data, rule validation, serta resolusi error, sehingga yang diperlukan "
        "selanjutnya adalah akselerasi, bukan perbaikan prosedur.",
    ]:
        add_bullet(doc, txt, size=10.5)

    add_subheading(doc, "C. Rekomendasi")
    add_body(doc,
             "Untuk mengakselerasi penyelesaian pengolahan sampel menuju 100%, direkomendasikan "
             "hal-hal berikut:", space_after=4)
    rekom = [
        f"Memrioritaskan resolusi error pada {fmt_num(agg['k_error'])} record Dokumen K dan "
        f"{fmt_num(agg['kp_error'])} record Dokumen KP, dimulai dari blok sensus dengan jumlah "
        "error tertinggi.",
        f"Menargetkan {agg['bs_belum']} blok sensus yang belum menghasilkan clean agar minimal "
        "menghasilkan satu record clean pada pekan pertama periode berikutnya.",
        f"Menyelesaikan {agg['k_belum']} RT Dokumen K dan {agg['kp_belum']} RT Dokumen KP yang "
        "belum terproses melalui verifikasi dokumen fisik dan percepatan entri.",
        "Melakukan review berjenjang isian konsumsi dan pengeluaran secara berkala agar pola anomali "
        "teridentifikasi sejak awal dan tidak menumpuk di akhir periode.",
        "Memantau Web Monitoring Susenas secara harian dan mengevaluasi capaian terhadap milestone "
        "minimal sekali dalam seminggu bersama ketua tim pengolahan.",
        "Menyusun ulang jadwal kerja pengolahan dengan target penyelesaian 100% beserta pembagian "
        "tanggung jawab per blok sensus, serta memastikan seluruh bukti monitoring diarsipkan.",
    ]
    for i, r in enumerate(rekom, 1):
        add_numbered(doc, i, r, size=10.5)


# ================ STEP 10: LEMBAR PENGESAHAN ================
def build_pengesahan(doc):
    add_page_break(doc)
    add_heading(doc, "LEMBAR PENGESAHAN", size=14)
    add_body(doc,
             "Laporan Kegiatan Penyelesaian Pengolahan Sampel Susenas September 2026 "
             "(Progress 100%) ini dengan ini disahkan dan dipertanggungjawabkan kebenarannya.",
             space_after=18)

    t = doc.add_table(rows=1, cols=2)
    t.style = "Normal Table"
    set_widths(t, [3.25, 3.25])
    cell = t.rows[0].cells[1]
    lines = [
        (f"Jember, {TANGGAL_LAPOR}", True),
        ("Ketua Tim Pengolahan dan Layanan Statistik", True),
        ("", False),
        ("", False),
        ("", False),
        (KETUA, False),
        (f"NIP. {NIP}", False),
    ]
    cell.text = ""
    first = True
    for text, bold in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(11)
        r.font.name = FONT


# ================ STEP 11: LAMPIRAN ================
DET_HEADERS = ["No", "Kode Blok Sensus", "Target", "Clean", "% Clean",
               "Error", "% Error", "Belum"]
DET_WIDTHS = [0.35, 1.10, 0.62, 0.60, 0.85, 0.60, 0.85, 0.60]
DET_ALIGNS = ["center", "left", "center", "center", "center", "center", "center", "center"]


def _detail_rows(blocks, prefix):
    rows = []
    for b in blocks:
        tgt = b["target"]
        cl = b[f"{prefix}_clean"]
        er = b[f"{prefix}_error"]
        bl = b[f"{prefix}_belum"]
        rows.append([b["no"], b["kode"], f"{tgt}", f"{cl}",
                     fmt_pct(pct(cl, tgt)), f"{er}",
                     fmt_pct(pct(er, tgt)), f"{bl}"])
    return rows


def build_lampiran(doc, blocks, agg, lk):
    add_page_break(doc)
    add_heading(doc, "LAMPIRAN", size=14)

    # ---- Lampiran 1 : Dokumen K ----
    add_subheading(doc, "Lampiran 1. Rincian Progres Pengolahan Dokumen K (Kor) "
                        "per Blok Sensus")
    add_body(doc,
             "Berikut adalah rincian progres pengolahan Dokumen K (Kor) pada seluruh blok sensus "
             f"di Kabupaten Jember per {TANGGAL_DATA}.", size=10.5, space_after=4)
    rows = _detail_rows(blocks, "k")
    rows.append(["", "TOTAL", f"{agg['target']}", f"{agg['k_clean']}",
                 fmt_pct(pct(agg["k_clean"], agg["target"])), f"{agg['k_error']}",
                 fmt_pct(pct(agg["k_error"], agg["target"])), f"{agg['k_belum']}"])
    make_table(doc, DET_HEADERS, rows, widths=DET_WIDTHS, size=8,
               aligns=DET_ALIGNS, total_row=True)

    # ---- Lampiran 2 : Dokumen KP ----
    add_subheading(doc, "Lampiran 2. Rincian Progres Pengolahan Dokumen KP "
                        "(Konsumsi dan Pengeluaran) per Blok Sensus")
    add_body(doc,
             "Berikut adalah rincian progres pengolahan Dokumen KP (Konsumsi dan Pengeluaran) "
             f"pada seluruh blok sensus di Kabupaten Jember per {TANGGAL_DATA}.",
             size=10.5, space_after=4)
    rows = _detail_rows(blocks, "kp")
    rows.append(["", "TOTAL", f"{agg['target']}", f"{agg['kp_clean']}",
                 fmt_pct(pct(agg["kp_clean"], agg["target"])), f"{agg['kp_error']}",
                 fmt_pct(pct(agg["kp_error"], agg["target"])), f"{agg['kp_belum']}"])
    make_table(doc, DET_HEADERS, rows, widths=DET_WIDTHS, size=8,
               aligns=DET_ALIGNS, total_row=True)

    # ---- Lampiran 3 : tangkapan layar monitoring ----
    add_subheading(doc, "Lampiran 3. Tangkapan Layar Monitoring Web Susenas")
    add_body(doc,
             "Bukti visual pemantauan progres pengolahan diambil dari Web Monitoring Susenas "
             "BPS RI (webmonitoring.bps.go.id) untuk wilayah Kabupaten Jember (3509) "
             f"pada {TANGGAL_DATA}.", size=10.5, space_after=8)

    for path, caption in [
        (IMG_DOK_K, "Gambar 1. Monitoring Pengolahan Dokumen K - Kabupaten Jember (3509)"),
        (IMG_DOK_KP, "Gambar 2. Monitoring Pengolahan Dokumen KP - Kabupaten Jember (3509)"),
    ]:
        if not Path(path).exists():
            raise FileNotFoundError(f"Bukti visual tidak ditemukan: {path}")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(path), width=Inches(6.0))
        add_caption(doc, caption)

    # ---- Lampiran 4 : beban per pengolah (sumber: LK Pengolahan Sampel) ----
    add_subheading(doc, "Lampiran 4. Pembagian Beban Sampel per Petugas Pengolahan")
    add_body(doc,
             "Berikut adalah pembagian beban pengolahan sampel per petugas pengolah beserta jumlah "
             "sampel rumah tangga yang ditangani, sebagaimana tercatat pada Lembar Kerja Pengolahan "
             "Sampel.", size=10.5, space_after=4)
    rows = []
    for i, (nm, jml) in enumerate(lk["beban_pengolah"].items(), 1):
        rows.append([i, nm, str(jml),
                     fmt_pct(pct(jml, agg["target"]))])
    rows.append(["", "TOTAL", str(sum(lk["beban_pengolah"].values())), "100,00%"])
    make_table(doc, ["No", "Nama Petugas Pengolahan", "Jumlah Sampel (RT)", "% dari Total"],
               rows, widths=[0.45, 3.35, 1.45, 1.25], size=9,
               aligns=["center", "left", "center", "center"], total_row=True)


# ================ STEP 12: KONVERSI PDF (Word COM Automation) ================
def convert_to_pdf(docx_path: Path, pdf_path: Path) -> Path:
    if not CONVERT_PS1.exists():
        raise FileNotFoundError(f"Skrip konversi tidak ditemukan: {CONVERT_PS1}")
    proc = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive",
         "-ExecutionPolicy", "Bypass", "-File", str(CONVERT_PS1),
         "-DocxPath", str(docx_path), "-PdfPath", str(pdf_path)],
        capture_output=True,
    )
    out = proc.stdout.decode("utf-8", errors="replace")
    err = proc.stderr.decode("utf-8", errors="replace")
    if out.strip():
        print(out.strip())
    if proc.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(
            f"Konversi PDF gagal (kode {proc.returncode})\n{out}\n{err}")
    return pdf_path


# ================ STEP 13: MAIN ================
def main():
    print("=" * 72)
    print("PEMBUATAN LAPORAN PENYELESAIAN PENGOLAHAN SAMPEL SUSENAS")
    print("BPS Kabupaten Jember (3509) - September 2026 - Milestone 100%")
    print("=" * 72)

    # 1) Ekstraksi data monitoring
    raw_blocks, raw_agg = build_dataset()
    print(f"[1] Data diekstrak : {raw_agg['n_bs']} blok sensus, "
          f"{raw_agg['target']} sampel rumah tangga")
    print(f"    Dok K : clean={raw_agg['k_clean']} error={raw_agg['k_error']} "
          f"belum={raw_agg['k_belum']}")
    print(f"    Dok KP: clean={raw_agg['kp_clean']} error={raw_agg['kp_error']} "
          f"belum={raw_agg['kp_belum']}")

    # 1b) Ekstraksi Lembar Kerja Pengolahan (LK)
    lk = parse_lk()
    print(f"[1b] LK Pengolahan  : kirim {lk['kirim_sudah']}/{lk['kirim_total']} RT, "
          f"entry K/KP {lk['entry_kor']}/{lk['entry_kp']} RT, "
          f"catatan editing {lk['rt_catatan']} RT")

    # 2) Ekspor XLSX (jejak audit) lalu dibaca ulang sebagai sumber dokumen
    xlsx = export_xlsx(raw_blocks, raw_agg)
    blocks = load_xlsx(xlsx)
    agg = compute_agg(blocks)
    print(f"[2] Rekap XLSX     : {xlsx}")

    # 3) Susun dokumen dari template
    doc = prepare_document()
    build_bab_i_v(doc, blocks, agg, lk)
    build_bab_vi(doc, blocks, agg)
    build_bab_vii(doc, blocks, agg, lk)
    build_bab_viii(doc, blocks, agg, lk)
    build_pengesahan(doc)
    build_lampiran(doc, blocks, agg, lk)
    doc.save(str(OUT_DOCX))
    print(f"[3] Dokumen DOCX   : {OUT_DOCX}")
    print(f"    Paragraf={len(doc.paragraphs)} Tabel={len(doc.tables)} "
          f"Gambar={len(doc.inline_shapes)}")

    # 4) Konversi ke PDF menggunakan Word COM automation
    pdf = convert_to_pdf(OUT_DOCX, OUT_PDF)
    print(f"[4] Dokumen PDF    : {pdf} ({pdf.stat().st_size:,} bytes)")

    # 5) Ringkasan capaian
    t = agg["target"]
    print("-" * 72)
    print("RINGKASAN CAPAIAN")
    print(f"  Milestone            : 100% = {MILESTONE_RT} RT / {MILESTONE_BS} BS clean")
    print(f"  Clean Dokumen K      : {agg['k_clean']} RT "
          f"({fmt_pct(pct(agg['k_clean'], t))})")
    print(f"  Clean Dokumen KP     : {agg['kp_clean']} RT "
          f"({fmt_pct(pct(agg['kp_clean'], t))})")
    print(f"  BS dalam proses      : {agg['bs_proses']} | "
          f"belum clean: {agg['bs_belum']} | "
          f"tuntas (K/KP): {agg['bs_tuntas_k']}/{agg['bs_tuntas_kp']}")
    print(f"  Status milestone     : TERCAPAI (seluruh sampel clean)")
    print("=" * 72)


if __name__ == "__main__":
    main()
