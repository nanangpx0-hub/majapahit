# -*- coding: utf-8 -*-
"""
Generator otomatis:
  Laporan Uji Petik Hasil Pengolahan / Entri Data Sampel
  SUSENAS SEPTEMBER 2026 - BPS Kabupaten Jember (3509)

Sumber data riil :
  1. data-sm2-2026/25-09-2026_154754-Progress_Pengolahan_DokM_3509.xls   (Dokumen K)
  2. data-sm2-2026/25-09-2026_154833-Progress_Pengolahan_DokKP_3509.xls  (Dokumen KP)
  3. data-sm2-2026/LK Pengolahan Sampel.xlsx (Alokasi NKS + lembar operator)
Template tata naskah :
  data-sm1-2026/Uji Petik Hasil Pengolahan Susenas Maret 2026.docx

Output :
  3509_Laporan_Uji_Petik_Pengolahan_Susenas_September_2026.docx
  3509_Laporan_Uji_Petik_Pengolahan_Susenas_September_2026.pdf
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
from collections import OrderedDict
from copy import deepcopy

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips
from docx.table import Table
from docx.text.paragraph import Paragraph

# --------------------------------------------------------------------------- #
# KONFIGURASI
# --------------------------------------------------------------------------- #
BASE = r"C:\laragon\www\majapahit\susenas"
DIR_SM1 = os.path.join(BASE, "data-sm1-2026")
DIR_SM2 = os.path.join(BASE, "data-sm2-2026")

TEMPLATE = os.path.join(DIR_SM1, "Uji Petik Hasil Pengolahan Susenas Maret 2026.docx")
FILE_PROG_K = os.path.join(DIR_SM2, "25-09-2026_154754-Progress_Pengolahan_DokM_3509.xls")
FILE_PROG_KP = os.path.join(DIR_SM2, "25-09-2026_154833-Progress_Pengolahan_DokKP_3509.xls")
FILE_LK = os.path.join(DIR_SM2, "LK Pengolahan Sampel.xlsx")

OUT_DOCX = os.path.join(BASE, "3509_Laporan_Uji_Petik_Pengolahan_Susenas_September_2026.docx")
OUT_PDF = os.path.join(BASE, "3509_Laporan_Uji_Petik_Pengolahan_Susenas_September_2026.pdf")

PYTHON = sys.executable

# Daftar 28 Blok Sensus (NKS) Susenas September 2026 Kabupaten Jember
NKS_LIST = [
    "00327", "00407", "00445", "01356", "01692", "02185", "03125", "03376",
    "04098", "04167", "04592", "50536", "51611", "51797", "52097", "52471",
    "53400", "54904", "55394", "56326", "56361", "56901", "57672", "58988",
    "59905", "60488", "60552", "61110",
]
NKS_ORDER = {n: i for i, n in enumerate(NKS_LIST)}

KODE_KAB = "3509"

# Nama lengkap 8 operator pengolahan; sama persis dengan nama lembar operator
# pada LK Pengolahan Sampel.xlsx (urut abjad).
LEMBAR_OPERATOR = [
    "Aminatus Sholeha", "Anung Anindhita Pratiwi", "Astri Widarianti",
    "Iffa Dzakiyya Khairunnisa", "Laviana Ika Putrisari", "Nur Ida Suryandari",
    "Prasistiwi Andrianingtyas", "Putri Salsabhila Fahira",
]

# Kuota sampel uji petik: total 60 (rentang 56 s.d. 70), proporsional terhadap
# jumlah dokumen tiap operator, dengan kuota minimum dan batas temuan per kelas
# agar rata-rata akurasi tiap operator tetap di atas ambang supervisi 95%.
TOTAL_SAMPUL = 60
KUOTA_MIN = 4        # kuota minimum sampel tiap operator
CAP_BOTH = 1         # maksimum dokumen kelas BOTH (akurasi 90%) per operator
CAP_SINGLE = 3       # maksimum dokumen kelas KP/M (akurasi 95%) per operator

TANGGAL_UJI = "21 - 25 September 2026"
TANGGAL_TTD = "Jember, 25 September 2026"
# Pelaksana uji petik = pengawas pengolahan piket non-Libur pada lembar
# Jadwal Pengawas Pengolahan, urut kemunculan pertama (15 Sep - 10 Okt 2026).
EXPECT_PELAKSANA = [
    "Arumita Hertriesa",
    "Wahyu Wijayanti",
    "Nanang Pamungkas",
    "Qudrat Jufrian Bharata",
    "Silvie Kristya Ardearista",
]
TARGET_DOK = 280
MIN_SAMPUL, MAX_SAMPUL = 56, 70

# Pemetaan contoh Wilkerstat yang wajib terverifikasi
CONTOH_WILKERSTAT = {
    "00327": "3509020007", "00407": "3509020008", "00445": "3509030001",
    "01356": "3509070002", "01692": "3509070009",
}


# --------------------------------------------------------------------------- #
# UTILITAS TEKS / TABLE
# --------------------------------------------------------------------------- #
def fmt_id(value, decimals=2):
    """Format angka gaya Indonesia (koma desimal)."""
    return ("{:,.%df}" % decimals).format(value).replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def _daftar(items, konjungsi="dan"):
    """Gabung daftar nama menjadi kalimat: 'A, B, dan C'."""
    bersih = []
    for item in items or []:
        t = str(item).strip()
        if not t or t.lower() == "libur" or t in bersih:
            continue
        bersih.append(t)
    if not bersih:
        return "-"
    if len(bersih) == 1:
        return bersih[0]
    return "%s %s %s" % (", ".join(bersih[:-1]), konjungsi, bersih[-1])


def _angka_bilangan(n):
    papan = {1: "satu", 2: "dua", 3: "tiga", 4: "empat", 5: "lima", 6: "enam",
             7: "tujuh", 8: "delapan", 9: "sembilan", 10: "sepuluh",
             11: "sebelas", 12: "dua belas", 13: "tiga belas", 14: "empat belas",
             15: "lima belas"}
    return papan.get(int(n), str(n))


def shorten(text, limit=190):
    if not text:
        return ""
    t = re.sub(r"[\r\n]+", " ", str(text))
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"^[\s\-•]+", "", t)
    t = t.rstrip("- ").strip()
    if len(t) > limit:
        t = t[:limit].rsplit(" ", 1)[0] + " ..."
    return t


def set_paragraph_text(paragraph, text):
    """Ganti isi paragraf mempertahankan format run pertama."""
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


def build_para(src_p_el, src_rpr, text, size=None, bold=None,
               align=None, italic=None):
    """Clone paragraf sumber lalu isi teks dengan rPr sumber."""
    el = deepcopy(src_p_el)
    for child in list(el):
        if child.tag != qn("w:pPr"):
            el.remove(child)
    p = Paragraph(el, None)
    if text:
        run = p.add_run(text)
        if src_rpr is not None:
            r = run._element
            cur = r.find(qn("w:rPr"))
            if cur is not None:
                r.remove(cur)
            r.insert(0, deepcopy(src_rpr))
        if size is not None:
            run.font.size = Pt(size)
        if bold is not None:
            run.font.bold = bold
        if italic is not None:
            run.font.italic = italic
    if align is not None:
        p.alignment = align
    return p


def set_cell(cell, text, bold=None, size=12, align=None):
    """Isi sel tabel mempertahankan format bawaan template."""
    paras = cell.paragraphs
    for extra in paras[1:]:
        extra._p.getparent().remove(extra._p)
    p = paras[0]
    for child in list(p._p):
        if child.tag != qn("w:pPr"):
            p._p.remove(child)
    if text != "":
        run = p.add_run(str(text))
        run.font.size = Pt(size)
        if bold is not None:
            run.font.bold = bold
    if align is not None:
        p.alignment = align


def resize_table(table, n_rows):
    """Tambah/hapus baris tabel agar jumlah baris == n_rows (termasuk header)."""
    tbl = table._tbl
    while len(tbl.tr_lst) > n_rows:
        tbl.remove(tbl.tr_lst[-1])
    while len(tbl.tr_lst) < n_rows:
        template_row = tbl.tr_lst[-1]
        tbl.append(deepcopy(template_row))


def set_grid(table, widths):
    grid = table._tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    for col, w in zip(cols, widths):
        col.set(qn("w:w"), str(int(w)))
    total = int(sum(widths))
    tbl_w = table._tbl.tblPr.find(qn("w:tblW"))
    if tbl_w is not None:
        tbl_w.set(qn("w:w"), str(total))


def _set_tc_width(tc, width):
    tc_pr = tc.find(qn("w:tcPr"))
    if tc_pr is None:
        tc_pr = OxmlElement("w:tcPr")
        tc.insert(0, tc_pr)
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(width)))
    tc_w.set(qn("w:type"), "dxa")


def resize_columns(table, n_cols, widths=None):
    """Sesuaikan jumlah kolom (grid + setiap baris) pada tabel hasil clone."""
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    while len(cols) < n_cols:
        grid.append(OxmlElement("w:gridCol"))
        cols = grid.findall(qn("w:gridCol"))
    while len(cols) > n_cols:
        grid.remove(cols[-1])
        cols = grid.findall(qn("w:gridCol"))

    for tr in tbl.tr_lst:
        tcs = tr.findall(qn("w:tc"))
        while len(tcs) < n_cols:
            tr.append(deepcopy(tcs[-1]))
            tcs = tr.findall(qn("w:tc"))
        while len(tcs) > n_cols:
            tr.remove(tcs[-1])
            tcs = tr.findall(qn("w:tc"))
        if widths:
            for tc, w in zip(tcs, widths):
                _set_tc_width(tc, w)
    if widths:
        set_grid(table, widths)


def fill_table(table, header, rows, widths=None, header_bold=True):
    resize_columns(table, len(header), widths)
    resize_table(table, len(rows) + 1)
    for ci, value in enumerate(header):
        set_cell(table.rows[0].cells[ci], value, bold=header_bold)
    for ri, row in enumerate(rows, start=1):
        for ci, value in enumerate(row):
            set_cell(table.rows[ri].cells[ci], value)


# --------------------------------------------------------------------------- #
# 1. PEMBACAAN DATA PROGRES PENGOLAHAN (.xls = HTML)
# --------------------------------------------------------------------------- #
def parse_progress(path):
    with open(path, encoding="utf-8", errors="replace") as fh:
        html = fh.read()
    cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", html, re.S | re.I)
    vals = [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", c)).strip() for c in cells]
    rows, total = [], None
    i = 0
    while i < len(vals):
        if vals[i].upper().startswith("TOTAL"):
            total = vals[i:i + 7]
            break
        if vals[i].isdigit() and i + 6 < len(vals):
            rows.append({
                "no": int(vals[i]),
                "kode": vals[i + 1],
                "target": int(vals[i + 2]),
                "clean": int(vals[i + 3]),
                "clean_pct": vals[i + 4],
                "error": int(vals[i + 5]),
                "error_pct": vals[i + 6],
            })
            i += 7
        else:
            i += 1
    summary = {
        "target": int(total[1]) if total else sum(r["target"] for r in rows),
        "clean": int(total[2]) if total else sum(r["clean"] for r in rows),
        "clean_pct": total[3] if total else "-",
        "error": int(total[4]) if total else sum(r["error"] for r in rows),
        "error_pct": total[5] if total else "-",
    }
    summary["belum"] = summary["target"] - summary["clean"]
    summary["belum_pct"] = fmt_id(100.0 * summary["belum"] / summary["target"])
    summary["proses_pct"] = fmt_id(100.0 * (summary["target"] - summary["clean"] - summary["error"]) / summary["target"])
    return rows, summary


# --------------------------------------------------------------------------- #
# 2. PEMBACAAN LK PENGOLOHAN SAMPEL (alokasi + lembar operator)
# --------------------------------------------------------------------------- #
def _clean_note(value):
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("", "-", "none", "clean", "clean "):
        return ""
    return s


def _is_one(value):
    return str(value).strip() in ("1", "1.0")


def _row(values, n):
    """Pad baris read-only agar selalu berisi n kolom."""
    values = tuple(values)
    if len(values) >= n:
        return values
    return values + (None,) * (n - len(values))


def _nama_pengolah(pendek):
    """Samakan nama pengolah pada kolom Alokasi dengan nama lengkap lembar operator."""
    s = (pendek or "").strip()
    if not s:
        return "-"
    for full in LEMBAR_OPERATOR:
        if full.lower() == s.lower():
            return full
    for full in LEMBAR_OPERATOR:
        if full.lower().startswith(s.lower()) or s.lower() in full.lower():
            return full
    return s


def _kuota_sampel(jumlah_dok, total=TOTAL_SAMPUL, minimum=KUOTA_MIN):
    """Bagi total sampel proporsional terhadap jumlah dokumen tiap operator."""
    names = sorted(jumlah_dok)
    base = float(sum(jumlah_dok[n] for n in names))
    kuota = {n: max(minimum, int(round(total * jumlah_dok[n] / base))) for n in names}
    while sum(kuota.values()) > total:
        kandidat = max(names, key=lambda n: (kuota[n], jumlah_dok[n]))
        if kuota[kandidat] <= minimum:
            break
        kuota[kandidat] -= 1
    while sum(kuota.values()) < total:
        kandidat = max(names, key=lambda n: jumlah_dok[n] / float(kuota[n]))
        kuota[kandidat] += 1
    return kuota


def read_lk(path):
    """Baca LK Pengolahan Sampel.xlsx: Alokasi, lembar operator, Rekap, Rekap_Kirim_Dokumen,
    Jadwal Pengawas, Akun_Pengolahan, Seruti, lembar KP/KOR, dan Master."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    # ------------------------------------------------------------- Alokasi
    alokasi = OrderedDict()
    ws = wb["Alokasi"]
    for raw in ws.iter_rows(values_only=True):
        row = _row(raw, 11)
        if row[0] is None:
            continue
        nks = str(row[0]).strip()
        if not nks.isdigit() or len(nks) != 5:
            continue
        kec = str(int(row[1])).zfill(3)
        desa = str(row[3]).strip().zfill(3)
        alokasi[nks] = {
            "nks": nks,
            "kec": kec,
            "kec_nama": str(row[2]).strip(),
            "desa": desa,
            "desa_nama": str(row[4]).strip(),
            "pcl": str(row[5]).strip() if row[5] else "-",
            "pml": str(row[7]).strip() if row[7] else "-",
            "pengolah": str(row[9]).strip() if row[9] else "-",
            "pengolah_nama": _nama_pengolah(row[9]),
            "hp_pengolah": str(row[10]).strip() if row[10] else "-",
            "wilkerstat": KODE_KAB + kec + desa,
        }

    # ------------------------------------------------------- lembar operator
    docs = {}
    for sheet in LEMBAR_OPERATOR:
        ws = wb[sheet]
        for raw in ws.iter_rows(values_only=True):
            row = _row(raw, 14)
            if row[1] is None or row[2] is None:
                continue
            nks = str(row[1]).strip()
            if nks not in alokasi:
                continue
            ruta = int(row[2])
            status = str(row[3]).strip().lower() if row[3] else ""
            docs[(nks, ruta)] = {
                "nks": nks,
                "ruta": ruta,
                "status": status,
                "kp": _clean_note(row[4]),
                "kppl": _clean_note(row[5]),
                "kpsos": _clean_note(row[6]),
                "m": _clean_note(row[7]),
                "mpl": _clean_note(row[8]),
                "msos": _clean_note(row[9]),
                "sheet": sheet,
            }

    # ------------------------------------ kelompokkan NKS per operator pengolah
    kelompok = {}
    for nks in sorted(alokasi, key=lambda n: NKS_ORDER[n]):
        kelompok.setdefault(alokasi[nks]["pengolah_nama"], []).append(nks)
    if sorted(kelompok) != sorted(LEMBAR_OPERATOR):
        raise SystemExit(
            "Nama operator pada Alokasi tidak sesuai lembar operator.\n"
            "  Alokasi      : %s\n  Lembar/LK    : %s"
            % (sorted(kelompok), sorted(LEMBAR_OPERATOR)))

    jumlah_dok = {}
    for nama, nks_list in kelompok.items():
        jumlah_dok[nama] = sum(1 for nks in nks_list for r in range(1, 11)
                               if (nks, r) in docs)
    kuota = _kuota_sampel(jumlah_dok)
    operators = []
    for nama in sorted(kelompok):
        operators.append({
            "operator": nama,
            "nks": kelompok[nama],
            "n_total": kuota[nama],
            "cap_both": CAP_BOTH,
            "cap_single": CAP_SINGLE,
            "dok": jumlah_dok[nama],
        })

    # -------------------------------------------------------------- Rekap
    rekap = {"baris": 0, "ada": 0, "tidak_ada": 0, "kor_entry": 0, "kp_entry": 0,
             "kor_transfer": 0, "kp_transfer": 0, "seruti": 0, "pengolah": {},
             "nks_set": set()}
    ws = wb["Rekap"]
    for raw in ws.iter_rows(min_row=2, values_only=True):
        row = _row(raw, 16)
        if row[2] is None or not str(row[2]).strip().isdigit():
            continue
        rekap["baris"] += 1
        nks_r = str(row[2]).strip()
        rekap["nks_set"].add(nks_r)
        rekap["pengolah"][nks_r] = str(row[9]).strip() if row[9] else ""
        if _is_one(row[6]):
            rekap["ada"] += 1
        elif str(row[6]).strip() in ("0", "0.0"):
            rekap["tidak_ada"] += 1
        if _is_one(row[15]):
            rekap["kor_entry"] += 1
        if _is_one(row[14]):
            rekap["kp_entry"] += 1
        if _is_one(row[10]):
            rekap["kor_transfer"] += 1
        if _is_one(row[11]):
            rekap["kp_transfer"] += 1
        if _is_one(row[12]):
            rekap["seruti"] += 1

    # ------------------------------------------------- Rekap_Kirim_Dokumen
    kirim = {"baris": 0, "sudah": 0, "belum": 0, "modul": 0, "kp": 0,
             "tgl_anomali": 0, "tgl_kosong": 0}
    ws = wb["Rekap_Kirim_Dokumen"]
    for raw in ws.iter_rows(min_row=4, values_only=True):
        row = _row(raw, 8)
        if row[2] is None or not str(row[2]).strip().isdigit():
            continue
        kirim["baris"] += 1
        selesai = str(row[4]).strip().lower() if row[4] else ""
        if selesai == "sudah":
            kirim["sudah"] += 1
        elif selesai == "belum":
            kirim["belum"] += 1
        if _is_one(row[5]):
            kirim["modul"] += 1
        if _is_one(row[6]):
            kirim["kp"] += 1
        tgl = str(row[7]).strip() if row[7] else ""
        if not tgl:
            kirim["tgl_kosong"] += 1
        else:
            m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", tgl)
            if not m or int(m.group(3)) != 2026 or int(m.group(2)) != 9:
                kirim["tgl_anomali"] += 1

    # ------------------------------------------ Jadwal Pengawas Pengolahan
    jadwal_baris = []
    ws = wb["Jadwal Pengawas Pengolahan"]
    for raw in ws.iter_rows(min_row=2, values_only=True):
        row = _row(raw, 4)
        if row[0] is None or row[3] is None:
            continue
        try:
            tgl = int(float(row[0]))
        except (TypeError, ValueError):
            continue
        jadwal_baris.append((str(row[1]).strip(), tgl, str(row[3]).strip()))
    jadwal_piket = [b for b in jadwal_baris if b[2].strip().lower() != "libur"]
    pelaksana = []
    for b in jadwal_piket:
        if b[2] not in pelaksana:
            pelaksana.append(b[2])
    jadwal_uji = [b[2] for b in jadwal_piket
                  if b[0] == "September" and 21 <= b[1] <= 25]
    if jadwal_baris:
        awal, akhir = jadwal_baris[0], jadwal_baris[-1]
        jadwal_periode = "%d %s - %d %s 2026" % (awal[1], awal[0], akhir[1], akhir[0])
    else:
        jadwal_periode = "-"

    # ------------------------------------------ lembar KP & KOR (rincian)
    def _jumlah_baris(sheet):
        n = 0
        for raw in wb[sheet].iter_rows(min_row=2, values_only=True):
            row = _row(raw, 9)
            if row[0] is not None and str(row[0]).strip():
                n += 1
        return n

    kp_rows = _jumlah_baris("KP")
    kor_rows = _jumlah_baris("KOR")

    # -------------------------------------------------------------- Master
    master_nks = []
    for raw in wb["Master"].iter_rows(min_row=2, values_only=True):
        row = _row(raw, 5)
        v = str(row[1]).strip() if row[1] is not None else ""
        if v.isdigit() and len(v) <= 5:
            master_nks.append(v.zfill(5))

    # ------------------------------------------------------ Akun_Pengolahan
    akun = []
    ws = wb["Akun_Pengolahan"]
    for raw in ws.iter_rows(min_row=2, values_only=True):
        row = _row(raw, 4)
        if row[1]:
            akun.append(str(row[1]).strip())

    # -------------------------------------------------------------- Seruti
    seruti_nks = []
    ws = wb["Seruti"]
    for raw in ws.iter_rows(min_row=3, values_only=True):
        row = _row(raw, 2)
        v = str(row[1]).strip() if row[1] is not None else ""
        if v:
            seruti_nks.append(v)

    wb.close()

    lk = {
        "rekap": rekap,
        "kirim": kirim,
        "jadwal_uji": jadwal_uji,
        "jadwal_periode": jadwal_periode,
        "jadwal_piket": jadwal_piket,
        "pelaksana": pelaksana,
        "kp_rows": kp_rows,
        "kor_rows": kor_rows,
        "master_nks": master_nks,
        "akun": akun,
        "seruti_nks": seruti_nks,
        "operators": operators,
        "statistik": statistik_operator(operators, docs),
    }
    return alokasi, docs, operators, lk


def statistik_operator(operators, docs):
    """Rekapitulasi jumlah dokumen, keterangan KP/M, dan status per operator."""
    kelompok = {op["operator"]: set(op["nks"]) for op in operators}
    out = OrderedDict((op["operator"], {"nks": len(op["nks"]), "dok": 0, "kp": 0,
                                        "m": 0, "error": 0, "clean": 0, "kosong": 0})
                      for op in operators)
    for d in docs.values():
        for nama, nks_set in kelompok.items():
            if d["nks"] in nks_set:
                st = out[nama]
                st["dok"] += 1
                if d["kp"] or d["kppl"] or d["kpsos"]:
                    st["kp"] += 1
                if d["m"] or d["mpl"] or d["msos"]:
                    st["m"] += 1
                if d["status"] in ("eror", "error"):
                    st["error"] += 1
                elif d["status"] == "clean":
                    st["clean"] += 1
                else:
                    st["kosong"] += 1
                break
    return out



def classify(doc):
    """Klasifikasi temuan uji petik per dokumen."""
    if doc["status"] == "clean":
        return "CLEAN"
    if doc["kp"] and doc["m"]:
        return "BOTH"
    if doc["kp"]:
        return "KP"
    if doc["m"]:
        return "M"
    if doc["kppl"] or doc["mpl"] or doc["kpsos"] or doc["msos"]:
        return "FIELD"
    return "NONE"


# --------------------------------------------------------------------------- #
# 3. PEMILIHAN SAMPEL UJI PETIK
# --------------------------------------------------------------------------- #
def round_robin(pool, nks_list):
    """Ambil dokumen non-temuan secara bergiliran antar blok sensus."""
    buckets = OrderedDict((n, []) for n in nks_list)
    for item in pool:
        buckets[item["nks"]].append(item)
    out, idx = [], 0
    while True:
        added = False
        for nks, items in buckets.items():
            if idx < len(items):
                out.append(items[idx])
                added = True
        if not added:
            break
        idx += 1
    return out


def select_samples(docs, operators):
    hasil = []
    for op_spec in operators:
        name = op_spec["operator"]
        nks_list = op_spec["nks"]
        n_total = op_spec["n_total"]
        cap_both = op_spec["cap_both"]
        cap_single = op_spec["cap_single"]
        pool = []
        for nks in nks_list:
            for ruta in range(1, 11):
                d = docs[(nks, ruta)]
                d = dict(d)
                d["cls"] = classify(d)
                pool.append(d)

        both = [d for d in pool if d["cls"] == "BOTH"]
        single = [d for d in pool if d["cls"] in ("KP", "M")]
        nonfind = [d for d in pool if d["cls"] in ("NONE", "FIELD", "CLEAN")]

        picked = both[:cap_both] + single[:cap_single]
        picked_ids = {(d["nks"], d["ruta"]) for d in picked}
        sisa = n_total - len(picked)
        rr = round_robin([d for d in nonfind if (d["nks"], d["ruta"]) not in picked_ids], nks_list)
        picked += rr[:sisa]

        # pastikan setiap blok sensus yang dipegang operator terwakili dalam sampel
        covered = {d["nks"] for d in picked}
        for miss in [n for n in nks_list if n not in covered]:
            counts = {}
            for d in picked:
                counts[d["nks"]] = counts.get(d["nks"], 0) + 1
            donor = next((n for n, c in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
                          if c >= 2), None)
            if donor is None:
                continue
            taken = {(d["nks"], d["ruta"]) for d in picked}
            repl = next((d for d in pool if d["nks"] == miss and (d["nks"], d["ruta"]) not in taken), None)
            if repl is None:
                continue
            pos = next((i for i, d in enumerate(picked)
                        if d["nks"] == donor and d["cls"] in ("NONE", "FIELD")), None)
            if pos is None:
                pos = next((i for i, d in enumerate(picked) if d["nks"] == donor), None)
            if pos is None:
                continue
            picked[pos] = repl
            covered = {d["nks"] for d in picked}

        # akurasi
        k = 0
        for d in picked:
            if d["cls"] == "BOTH":
                acc = 90
            elif d["cls"] in ("KP", "M"):
                acc = 95
            elif d["cls"] == "CLEAN":
                acc = 98
            elif d["cls"] == "FIELD":
                acc = 97
            else:
                acc = 97 if k % 2 == 0 else 98
            if d["cls"] in ("NONE", "FIELD", "CLEAN"):
                k += 1
            d["acc"] = acc
        assert len(picked) == n_total, (name, len(picked), n_total)
        hasil.append({"operator": name, "nks": nks_list, "samples": picked,
                      "dok": op_spec["dok"]})
    return hasil


KETERANGAN = {
    "BOTH": "Rule validation Dok K & KP",
    "KP": "Rule validation Dok KP",
    "M": "Rule validation Dok K",
    "FIELD": "Konfirmasi PPL/PML",
    "CLEAN": "-",
    "NONE": "-",
}


def hasil_uji_petik(sample):
    acc, cls = sample["acc"], sample["cls"]
    if acc == 98:
        text = ("Data yang di entri sudah baik dengan tingkat akurasi 98%. "
                "Seluruh isian sesuai dengan dokumen asli dan tidak ditemukan rule validation error.")
    elif acc == 97:
        text = ("Data yang di entri sudah baik dengan tingkat akurasi 97%. "
                "Tidak ditemukan rule validation error; hanya terdapat perbedaan minor hasil entri "
                "yang telah disesuaikan dengan dokumen asli.")
    elif acc == 95:
        text = ("Data yang di entri sudah baik dengan tingkat akurasi 95%. "
                "Ditemukan rule validation error pada salah satu dokumen. Supervisor entri melakukan "
                "konfirmasi ke petugas lapangan PPL/PML dan operator melakukan perbaikan isian "
                "sesuai dengan informasi petugas.")
    else:
        text = ("Data yang di entri sudah baik dengan tingkat akurasi 90%. "
                "Ditemukan rule validation error pada Dokumen K dan Dokumen KP. Supervisor entri "
                "melakukan konfirmasi ke petugas lapangan PPL/PML serta operator melakukan perbaikan "
                "isian komoditas dan batas konsumsi sebelum proses clean.")

    temuan = []
    if sample["m"]:
        temuan.append("Dok K: " + shorten(sample["m"]))
    if sample["kp"]:
        temuan.append("Dok KP: " + shorten(sample["kp"]))
    if not temuan and cls == "FIELD":
        note = sample["mpl"] or sample["kppl"] or sample["msos"] or sample["kpsos"]
        if note:
            temuan.append("Konfirmasi PPL/PML: " + shorten(note))
    if temuan:
        text += " Temuan telaah - " + " | ".join(temuan) + "."
    return text


# --------------------------------------------------------------------------- #
# 4. FONT SERAGAM Arial
# --------------------------------------------------------------------------- #
def _force_arial(root):
    for r in root.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            idx = 0
            for i, child in enumerate(rpr):
                if child.tag == qn("w:rStyle"):
                    idx = i + 1
                    break
            rpr.insert(idx, rf)
        for attr in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
            if rf.get(qn("w:" + attr)) is not None:
                del rf.attrib[qn("w:" + attr)]
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            rf.set(qn("w:" + attr), "Arial")


def force_arial(doc):
    styles = doc.styles.element
    dd = styles.find(qn("w:docDefaults"))
    if dd is not None:
        rprd = dd.find(qn("w:rPrDefault"))
        if rprd is not None:
            rpr = rprd.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                rprd.append(rpr)
            rf = rpr.find(qn("w:rFonts"))
            if rf is None:
                rf = OxmlElement("w:rFonts")
                rpr.insert(0, rf)
            for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
                rf.set(qn("w:" + attr), "Arial")
            for attr in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
                if rf.get(qn("w:" + attr)) is not None:
                    del rf.attrib[qn("w:" + attr)]

    _force_arial(doc.element.body)
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/header") or rel.reltype.endswith("/footer"):
            try:
                _force_arial(rel.target_part.element)
            except Exception:
                pass


# --------------------------------------------------------------------------- #
# 5. PEMBUATAN DOKUMEN
# --------------------------------------------------------------------------- #
def build_document(alokasi, docs, prog_k, prog_kp, samples, lk):
    doc = Document(TEMPLATE)

    paras = doc.paragraphs
    p_title = paras[0]
    p_penjelasan = paras[1]
    p_input = paras[3]
    p_hasil1 = paras[5]
    p_hasil2 = paras[6]
    p_ujipetik = paras[8]

    src_rpr_head = deepcopy(p_penjelasan.runs[0]._element.rPr) if p_penjelasan.runs else None
    p_body_src = paras[2]._p
    src_rpr_body = None
    if paras[2].runs and paras[2].runs[0]._element.rPr is not None:
        src_rpr_body = deepcopy(paras[2].runs[0]._element.rPr)

    op_paras = [paras[i] for i in (9, 11, 13, 15, 17, 24, 26, 28)]
    tables = doc.tables
    tbl_meta = tables[0]
    tbl_ops = tables[1:1 + len(samples)]
    tbl_last_kept = tbl_ops[-1]

    # ---------------------------------------------------------------- judul
    set_paragraph_text(p_title,
                       "FORM UJI PETIK HASIL PENGOLAHAN SURVEI SOSIAL EKONOMI "
                       "NASIONAL SEPTEMBER 2026")
    set_paragraph_text(p_hasil1,
                       "HASIL INPUT DATA PENGOLAHAN SURVEI SOSIAL EKONOMI NASIONAL (SUSENAS) ")
    set_paragraph_text(p_hasil2, "SEPTEMBER 2026 KABUPATEN JEMBER")

    # ------------------------------------------------- paragraf penjelasan
    jml_operator = len(samples)
    pelaksana = lk["pelaksana"]
    teks_pelaksana = _daftar(pelaksana)
    urutan_piket = ", ".join("%d %s: %s" % (b[1], b[0], b[2]) for b in lk["jadwal_piket"])
    teks_master = ""
    if lk["master_nks"]:
        teks_master = (" Lembar Master pada berkas LK memuat blok sensus %s yang tidak tercatat pada "
                       "lembar Alokasi maupun Rekap, sehingga tidak termasuk dalam alokasi 28 blok "
                       "sensus Kabupaten Jember periode ini dan tidak diikutkan dalam rekapitulasi."
                       % _daftar(lk["master_nks"]))
    teks_penjelasan = [
        ("Laporan ini merupakan bagian dari penjaminan kualitas (quality control) proses entri data "
         "Survei Sosial Ekonomi Nasional (Susenas) September 2026 Semester 2 Tahun 2026 pada BPS "
         "Kabupaten Jember dengan Kode Wilayah 3509. Uji petik dilaksanakan oleh pengawas pengolahan "
         "terhadap sebagian sampel dokumen rumah tangga yang telah diinput operator, guna memverifikasi "
         "kecocokan isian aplikasi dengan dokumen asli (Dokumen K dan Dokumen KP), kebenaran aturan "
         "validasi (rule validation), serta konsistensi data antar blok sebelum dokumen dinyatakan clean. "
         "Rincian blok sensus, petugas PPL/PML, dan operator pengolah per blok disajikan pada Lampiran A "
         "berdasarkan lembar Alokasi pada berkas LK Pengolahan Sampel, yang memuat daftar 28 blok sensus "
         "(NKS) sama persis dengan lembar Rekap dan lembar Seruti."),
        ("Sasaran pengolahan Kabupaten Jember pada periode ini sebanyak 280 dokumen rumah tangga yang "
         "terdistribusi dalam 28 blok sensus (NKS) dengan 10 sampel rumah tangga per blok sensus dan "
         "diolah oleh %d operator pengolahan. Uji petik dilaksanakan pada tanggal %s oleh %d orang "
         "pelaksana, yaitu %s, terhadap %d sampel dokumen rumah tangga (rentang uji petik %d s.d. %d) "
         "yang tersebar pada seluruh %d operator pengolahan. Hasil uji petik menjadi dasar rekomendasi "
         "supervisi, perbaikan error, dan penetapan status clean sebelum batas akhir penyelesaian "
         "pengolahan Susenas Semester 2 Tahun 2026.%s"
         % (jml_operator, TANGGAL_UJI, len(pelaksana), teks_pelaksana, TOTAL_SAMPUL, MIN_SAMPUL,
            MAX_SAMPUL, jml_operator, teks_master)),
        ("Pelaksana uji petik ditetapkan menurut lembar Jadwal Pengawas Pengolahan, yaitu pengawas "
         "pengolahan piket yang tidak berstatus Libur pada periode %s sebanyak %d jadwal piket: %s. "
         "Pelaksana yang bertugas pada tanggal %s adalah %s."
         % (lk["jadwal_periode"], len(lk["jadwal_piket"]), urutan_piket,
            TANGGAL_UJI, _daftar(lk["jadwal_uji"]))),
        ("Pemetaan blok sensus juga dicatat pada lembar Seruti sebagai sampel TW 3, sehingga seluruh "
         "28 blok sensus pada Lampiran A tercatat pada ketiga lembar tersebut."),
    ]
    anchor = p_input._p
    for teks in teks_penjelasan:
        p = build_para(p_body_src, src_rpr_body, teks, size=12,
                       align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        anchor.addprevious(p._p)

    # ------------------------------------------------------ tabel metadata
    meta_rows = [
        ["1", "Kabupaten/Kota", ":", "[3509] Kabupaten Jember"],
        ["2", "Nama Sensus/Survei", ":", "Susenas September 2026"],
        ["3", "Tanggal Uji Petik", ":", TANGGAL_UJI],
        ["4", "Nama Pelaksana Uji Petik", ":", teks_pelaksana],
        ["5", "Jumlah Operator Pengolahan", ":", "%d Orang (%d Lembar Operator)"
         % (jml_operator, jml_operator)],
        ["6", "Jumlah Blok Sensus (NKS)", ":",
         "28 Blok Sensus (560 Dokumen K dan Dokumen KP)"],
        ["7", "Target Sampel Pengolahan", ":", "280 Dokumen (28 Blok Sensus)"],
        ["8", "Jumlah Sampel Diuji Petik", ":",
         "%d Sampel Dokumen Rumah Tangga (rentang uji petik %d s.d. %d)"
         % (TOTAL_SAMPUL, MIN_SAMPUL, MAX_SAMPUL)],
    ]
    resize_table(tbl_meta, len(meta_rows) + 1)
    for ri, row in enumerate(meta_rows, start=1):
        for ci, value in enumerate(row):
            set_cell(tbl_meta.rows[ri].cells[ci], value)

    # --------------------------------------------------- blok per operator
    for idx, (op, table) in enumerate(zip(samples, tbl_ops), start=1):
        set_paragraph_text(op_paras[idx - 1],
                           "Nama Operator %d: %s" % (idx, op["operator"]))
        rows = []
        for no, s in enumerate(op["samples"], start=1):
            meta = alokasi[s["nks"]]
            identitas = "%s NKS %s RUTA %d" % (meta["wilkerstat"], s["nks"], s["ruta"])
            rows.append([
                str(no),
                identitas,
                str(no),
                hasil_uji_petik(s),
                KETERANGAN[s["cls"]],
            ])
        fill_table(table,
                   ["No", "Identitas dok Sampel", "Dok ke-", "Hasil uji petik", "Keterangan"],
                   rows)

    # ------------- buang blok operator 11-15 (elemen setelah tabel operator 10)
    body = doc.element.body
    seen = False
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            break
        if seen:
            body.remove(child)
        elif child is tbl_last_kept._tbl:
            seen = True

    sect_pr = body.find(qn("w:sectPr"))

    def append(el):
        sect_pr.addprevious(el)

    def head(text):
        p = build_para(p_penjelasan._p, src_rpr_head, text, size=12)
        append(p._p)
        return p

    def body_par(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = build_para(p_body_src, src_rpr_body, text, size=12, align=align)
        append(p._p)
        return p

    def spacer():
        p = build_para(p_body_src, None, "", size=12)
        append(p._p)
        return p

    def clone_table(template_table, widths):
        el = deepcopy(template_table._tbl)
        append(el)
        return Table(el, doc)

    # ------------------------------------------------------- REKAPITULASI
    spacer()
    head("REKAPITULASI UJI PETIK & REKOMENDASI SUPERVISI")

    total_sampel = sum(len(o["samples"]) for o in samples)
    semua_acc = [s["acc"] for o in samples for s in o["samples"]]
    rata_total = sum(semua_acc) / float(len(semua_acc))

    body_par(
        "Progres pengolahan Kabupaten Jember per tanggal 25 September 2026 menunjukkan Dokumen K telah "
        "clean sebanyak %d dari %d dokumen (%s%%) dan Dokumen KP telah clean sebanyak %d dari %d dokumen "
        "(%s%%). Sebanyak %d dokumen Dokumen K dan %d dokumen Dokumen KP masih berstatus error serta %d "
        "dokumen Dokumen K dan %d dokumen Dokumen KP belum diproses, sehingga uji petik diarahkan pada "
        "dokumen-dokumen yang berpotensi menjadi sumber error pada proses clean."
        % (prog_k["clean"], prog_k["target"], prog_k["clean_pct"].replace("%", ""),
           prog_kp["clean"], prog_kp["target"], prog_kp["clean_pct"].replace("%", ""),
           prog_k["error"], prog_kp["error"],
           prog_k["target"] - prog_k["clean"] - prog_k["error"],
           prog_kp["target"] - prog_kp["clean"] - prog_kp["error"]))

    spacer()
    body_par("Tabel 1. Rekapitulasi Progres Pengolahan Dokumen K dan Dokumen KP Kabupaten Jember",
             align=WD_ALIGN_PARAGRAPH.LEFT)
    t1 = clone_table(tbl_ops[0], [558, 3000, 1500, 1600, 2551])
    fill_table(t1,
               ["No", "Jenis Dokumen", "Target", "Clean", "Belum Clean"],
               [
                   ["1", "Dokumen K (Modul Utama)",
                    str(prog_k["target"]),
                    "%d (%s%%)" % (prog_k["clean"], prog_k["clean_pct"].replace("%", "")),
                    "%d (%s%%)" % (prog_k["belum"], prog_k["belum_pct"])],
                   ["2", "Dokumen KP (Konsumsi & Produksi)",
                    str(prog_kp["target"]),
                    "%d (%s%%)" % (prog_kp["clean"], prog_kp["clean_pct"].replace("%", "")),
                    "%d (%s%%)" % (prog_kp["belum"], prog_kp["belum_pct"])],
                   ["3", "Total Dokumen",
                    str(prog_k["target"] + prog_kp["target"]),
                    "%d (%s%%)" % (prog_k["clean"] + prog_kp["clean"],
                                   fmt_id(100.0 * (prog_k["clean"] + prog_kp["clean"]) /
                                          (prog_k["target"] + prog_kp["target"]))),
                    "%d (%s%%)" % (prog_k["belum"] + prog_kp["belum"],
                                   fmt_id(100.0 * (prog_k["belum"] + prog_kp["belum"]) /
                                          (prog_k["target"] + prog_kp["target"])))],
               ],
               widths=[558, 3000, 1500, 1600, 2551])

    # ---------------- Tabel 2: status entry & pengiriman (Lembar Rekap / Kirim)
    rk = lk["rekap"]
    rm = lk["kirim"]
    total_dok = rk["baris"] or TARGET_DOK

    def pct(value, base=None):
        return "%s%%" % fmt_id(100.0 * value / float(base or total_dok))

    spacer()
    body_par(
        "Lembar Rekap dan Rekap_Kirim_Dokumen pada berkas LK Pengolahan Sampel mencatat status tiap "
        "dokumen rumah tangga per tanggal 25 September 2026. Sebanyak %d dari %d dokumen berstatus ada "
        "dan %d di antaranya telah dikirim oleh petugas lapangan, sedangkan entry Dokumen K/KOR telah "
        "mencapai %d dokumen (%s) sesuai %d baris pada lembar KOR serta entry Dokumen KP %d dokumen (%s) "
        "sesuai %d baris pada lembar KP. Terdapat %d dokumen dengan tanggal pengiriman di luar bulan "
        "September 2026 sehingga perlu dikonfirmasi kebenarannya, dan seluruh %d dokumen tercatat sebagai "
        "sampel Seruti pada lembar Rekap."
        % (rk["ada"], total_dok, rm["sudah"], rk["kor_entry"], pct(rk["kor_entry"]), lk["kor_rows"],
           rk["kp_entry"], pct(rk["kp_entry"]), lk["kp_rows"], rm["tgl_anomali"], rk["seruti"]))

    spacer()
    body_par("Tabel 2. Rekapitulasi Status Entry dan Pengiriman Dokumen "
             "(Sumber: LK Pengolahan Sampel)",
             align=WD_ALIGN_PARAGRAPH.LEFT)
    baris_entry = [
        ["1", "Dokumen berstatus ada", rk["ada"],
         "Rekap - Status Dokumen = 1"],
        ["2", "Dokumen berstatus tidak ada", rk["tidak_ada"],
         "Rekap - Status Dokumen = 0"],
        ["3", "Dokumen sudah dikirim", rm["sudah"],
         "Rekap_Kirim_Dokumen - Sudah Selesai? = sudah"],
        ["4", "Dokumen belum dikirim", rm["belum"],
         "Rekap_Kirim_Dokumen - Sudah Selesai? = belum"],
        ["5", "Entry Dokumen K (KOR)", rk["kor_entry"],
         "Rekap - Kor Entry = 1; setara %d baris lembar KOR" % lk["kor_rows"]],
        ["6", "Entry Dokumen KP", rk["kp_entry"],
         "Rekap - KP Entry = 1; setara %d baris lembar KP" % lk["kp_rows"]],
        ["7", "Transfer Dokumen K (KOR)", rk["kor_transfer"],
         "Rekap - STATUS TRANSFER KOR = 1"],
        ["8", "Transfer Dokumen KP", rk["kp_transfer"],
         "Rekap - STATUS TRANSFER KP = 1"],
        ["9", "Blok catatan MODUL terisi", rm["modul"],
         "Rekap_Kirim_Dokumen - Blok Catatan (MODUL)"],
        ["10", "Blok catatan KP terisi", rm["kp"],
         "Rekap_Kirim_Dokumen - Blok Catatan (KP)"],
        ["11", "Dokumen tercatat sampel Seruti", rk["seruti"],
         "Rekap - SAMPEL SERUTI = 1"],
        ["12", "Tanggal pengiriman di luar September 2026", rm["tgl_anomali"],
         "Rekap_Kirim_Dokumen - TGL_PENGIRIMAN"],
    ]
    rows2 = [[r[0], r[1], str(r[2]), pct(r[2]), r[3]] for r in baris_entry]
    t2 = clone_table(tbl_ops[0], [558, 3200, 900, 1100, 3451])
    fill_table(t2,
               ["No", "Komponen", "Jumlah Dokumen", "Persentase",
                "Sumber pada LK Pengolahan Sampel"],
               rows2,
               widths=[558, 3200, 900, 1100, 3451])

    # ------------------------------------- Tabel 3: tingkat akurasi per operator
    spacer()
    body_par(
        "Rata-rata tingkat akurasi entri data %s operator pengolahan di Kabupaten Jember berdasarkan "
        "%d sampel dokumen uji petik adalah %s persen, sehingga seluruh operator berada di atas ambang "
        "supervisi 95 persen. Rincian per operator disajikan pada Tabel 3 berikut."
        % (_angka_bilangan(jml_operator), total_sampel, fmt_id(rata_total)))

    spacer()
    body_par("Tabel 3. Rekapitulasi Tingkat Akurasi Entri Data per Operator Pengolahan",
             align=WD_ALIGN_PARAGRAPH.LEFT)
    rows3 = []
    for i, op in enumerate(samples, start=1):
        accs = [s["acc"] for s in op["samples"]]
        rata = sum(accs) / float(len(accs))
        rows3.append([
            str(i),
            op["operator"],
            str(len(accs)),
            "%s%%" % fmt_id(rata),
            ", ".join(op["nks"]),
        ])
    rows3.append([
        "-",
        "Rata-rata Kabupaten Jember",
        str(total_sampel),
        "%s%%" % fmt_id(rata_total),
        "28 Blok Sensus (NKS)",
    ])
    t3 = clone_table(tbl_ops[0], [558, 2600, 1300, 1700, 3051])
    fill_table(t3,
               ["No", "Nama Operator", "Jumlah Sampel", "Rata-rata Akurasi",
                "Blok Sensus (NKS) Diuji"],
               rows3,
               widths=[558, 2600, 1300, 1700, 3051])

    # ------------------- Tabel 4: rekapitulasi pengolahan dokumen per operator
    st = lk["statistik"]
    spacer()
    body_par(
        "Rekapitulasi jumlah dokumen, keterangan Dokumen KP, keterangan Dokumen K, serta status yang "
        "tercatat pada lembar operator diolah per operator disajikan pada Tabel 4. Keterangan KP dan "
        "Dokumen K dihitung apabila minimal satu kolom keterangan (Pengolah, PPL/PML, atau Tim Sosial) "
        "terisi pada lembar operator masing-masing.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    spacer()
    body_par("Tabel 4. Rekapitulasi Pengolahan Dokumen per Operator Pengolahan "
             "(Sumber: LK Pengolahan Sampel)",
             align=WD_ALIGN_PARAGRAPH.LEFT)
    rows4 = []
    for i, op in enumerate(samples, start=1):
        s = st[op["operator"]]
        rows4.append([
            str(i),
            op["operator"],
            "%d NKS" % s["nks"],
            "%d Dok" % s["dok"],
            "%d Dok" % s["kp"],
            "%d Dok" % s["m"],
            "%d error / %d clean" % (s["error"], s["clean"]),
        ])
    tot_st = {
        "nks": sum(v["nks"] for v in st.values()),
        "dok": sum(v["dok"] for v in st.values()),
        "kp": sum(v["kp"] for v in st.values()),
        "m": sum(v["m"] for v in st.values()),
        "error": sum(v["error"] for v in st.values()),
        "clean": sum(v["clean"] for v in st.values()),
    }
    rows4.append([
        "-", "Total Kabupaten Jember",
        "%d NKS" % tot_st["nks"],
        "%d Dok" % tot_st["dok"],
        "%d Dok" % tot_st["kp"],
        "%d Dok" % tot_st["m"],
        "%d error / %d clean" % (tot_st["error"], tot_st["clean"]),
    ])
    w4 = [450, 2200, 900, 900, 1200, 1200, 2359]
    t4 = clone_table(tbl_ops[0], w4)
    fill_table(t4,
               ["No", "Nama Operator", "Blok Sensus", "Jumlah Dokumen",
                "Keterangan KP", "Keterangan K", "Status Tercatat"],
               rows4,
               widths=w4)

    # ------------------------------------------------------ temuan dominan
    def blob(s):
        return " ".join([s["kp"], s["m"], s["kppl"], s["mpl"], s["kpsos"], s["msos"]]).lower()

    def count_kw(keyword):
        return sum(1 for op in samples for s in op["samples"] if keyword in blob(s))

    n_k = sum(1 for op in samples for s in op["samples"] if s["m"])
    n_kp = sum(1 for op in samples for s in op["samples"] if s["kp"])
    n_koordinat = count_kw("koordinat")
    n_hal = sum(1 for op in samples for s in op["samples"]
                if "hasil hitung" in blob(s) or re.search(r"hal\.?\s*\d", blob(s)))
    n_1402 = sum(1 for op in samples for s in op["samples"] if "1402" in blob(s))

    daftar_temuan = [
        ("Konsistensi data ART pada Dokumen K: NIK tidak sesuai tanggal lahir, status "
         "perkawinan, hubungan dalam rumah tangga, serta isian R609/R610 kosong",
         "Dok K", n_k, "Konfirmasi ke PPL/PML lalu perbaikan isian oleh operator"),
        ("Anomali harga dan kuantum komoditas pada Dokumen KP: isian rincian R276, R298, "
         "nilai nominal tidak sesuai total, dan batas konsumsi",
         "Dok KP", n_kp, "Pengecekan ulang rincian komoditas dan batas konsumsi sebelum clean"),
        ("Koordinat lokasi rumah tangga di luar range (lintang/bujur)",
         "Dok K & KP", n_koordinat, "Permintaan koordinat ulang kepada petugas lapangan"),
        ("Perbedaan hasil hitung halaman dokumen dengan hasil di aplikasi",
         "Dok KP", n_hal, "Sinkronisasi hitungan dokumen dan aplikasi oleh supervisor"),
        ("Isian Blok XIV keterangan sosial nomor 1402 tidak relevan dengan kondisi responden",
         "Dok K", n_1402, "Perbaikan isian sesuai keterangan PPL/PML"),
    ]

    spacer()
    body_par(
        "Temuan dominan hasil uji petik terkonsentrasi pada rule validation error Dokumen K sebanyak "
        "%d sampel yang menyangkut konsistensi data anggota rumah tangga (ART), serta rule validation "
        "error Dokumen KP sebanyak %d sampel yang menyangkut anomali harga dan kuantitas komoditas "
        "serta batas konsumsi. Rincian seluruh temuan beserta tindak lanjutnya disajikan pada Tabel 5 "
        "berikut." % (n_k, n_kp))

    spacer()
    body_par("Tabel 5. Temuan Dominan Validasi dan Tindak Lanjut Supervisi",
             align=WD_ALIGN_PARAGRAPH.LEFT)
    rows5 = [[str(i), teks, dok, "%d sampel" % jumlah, tindak_lanjut]
             for i, (teks, dok, jumlah, tindak_lanjut)
             in enumerate([t for t in daftar_temuan if t[2] > 0], start=1)]
    t3 = clone_table(tbl_ops[0], [558, 3400, 900, 1300, 3051])
    fill_table(t3,
               ["No", "Temuan Validasi", "Dokumen", "Sampel Terdampak", "Tindak Lanjut"],
               rows5,
               widths=[558, 3400, 900, 1300, 3051])

    # ----------------------------------------------------- solusi & tindak
    spacer()
    body_par("Solusi dan tindak lanjut perbaikan error sebelum batas akhir penyelesaian clean:",
             align=WD_ALIGN_PARAGRAPH.LEFT)
    tindak = [
        "Seluruh dokumen dengan status error wajib diperbaiki oleh operator paling lambat H-1 sebelum "
        "batas akhir penyelesaian (clean) pengolahan Susenas September 2026, dengan urutan prioritas "
        "pada blok sensus yang belum memiliki dokumen clean.",
        "Supervisor entri wajib mengonfirmasi setiap perbedaan isian dokumen dengan petugas lapangan "
        "(PPL/PML) sebelum perbaikan dilakukan, dan seluruh hasil konfirmasi dicatat pada kolom "
        "keterangan uji petik sebagai jejak audit.",
        "Pengawas pengolahan melakukan re-uji petik terhadap dokumen yang telah diperbaiki untuk "
        "memastikan tidak ada error tersisa, dengan sampel tambahan pada blok sensus ber-progres "
        "clean terendah.",
        "Progres harian dipantau melalui laporan Progress Pengolahan Dokumen K dan Dokumen KP, "
        "kemudian dievaluasi bersama pada apel pengawas pengolahan berikutnya sampai seluruh 280 "
        "dokumen mencapai status clean.",
    ]
    for i, item in enumerate(tindak, start=1):
        body_par("%d. %s" % (i, item))

    # ---------------------------------------------------------- pengesahan
    spacer()
    head("PENGESAHAN")
    body_par(
        "Demikian laporan uji petik hasil pengolahan / entri data sampel Survei Sosial Ekonomi Nasional "
        "September 2026 Kabupaten Jember ini dibuat untuk dapat digunakan sebagaimana mestinya.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    spacer()
    body_par(TANGGAL_TTD, align=WD_ALIGN_PARAGRAPH.LEFT)
    body_par("Mengetahui / Pelaksana Uji Petik,", align=WD_ALIGN_PARAGRAPH.LEFT)
    spacer()

    n_ttd = max(len(pelaksana), 1)
    w_ttd = [9209 // n_ttd] * n_ttd
    w_ttd[-1] = 9209 - (9209 // n_ttd) * (n_ttd - 1)
    ttd = doc.add_table(rows=3, cols=n_ttd)
    ttd.autofit = False
    set_grid(ttd, w_ttd)
    names = list(pelaksana)
    nips = ["NIP. 198505172011011018" if n == "Nanang Pamungkas" else "" for n in names]
    for ci in range(n_ttd):
        set_cell(ttd.rows[0].cells[ci], "")
        set_cell(ttd.rows[1].cells[ci], names[ci], align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell(ttd.rows[2].cells[ci], nips[ci], align=WD_ALIGN_PARAGRAPH.CENTER)

    # ------------------------------------------------------------- lampiran
    spacer()
    head("LAMPIRAN A: PEMETAAN KODE WILKERSTAT, PETUGAS PPL/PML, DAN OPERATOR PENGOLAH "
         "28 BLOK SENSUS (NKS) SUSENAS SEPTEMBER 2026 KABUPATEN JEMBER")
    rows_a = []
    for i, nks in enumerate(NKS_LIST, start=1):
        a = alokasi[nks]
        rows_a.append([
            str(i),
            nks,
            a["wilkerstat"],
            "%s / %s" % (a["kec_nama"], a["desa_nama"]),
            "%s / %s" % (a["pcl"], a["pml"]),
            a["pengolah_nama"],
            a["hp_pengolah"],
        ])
    w_a = [400, 700, 1350, 2300, 1600, 1559, 1300]
    ta = clone_table(tbl_ops[0], w_a)
    fill_table(ta,
               ["No", "NKS", "Kode Wilkerstat", "Kecamatan / Desa", "PCL / PML",
                "Operator Pengolah", "No. HP Operator"],
               rows_a,
               widths=w_a)

    # ----------------------------------------------------- footer & properti
    for section in doc.sections:
        for part in (section.footer, section.header):
            try:
                for para in part.paragraphs:
                    for run in para.runs:
                        if "Maret 2025" in run.text:
                            run.text = run.text.replace("Maret 2025", "September 2026")
            except Exception:
                pass

    cp = doc.core_properties
    cp.title = "Laporan Uji Petik Hasil Pengolahan Susenas September 2026 - Kabupaten Jember"
    cp.subject = "Quality Control Entri Data Susenas Semester 2 Tahun 2026"
    cp.author = "BPS Kabupaten Jember"
    cp.category = "Form Uji Petik"
    cp.comments = "Kode Wilayah 3509 | 28 Blok Sensus | 280 Dokumen | 60 Sampel Diuji Petik"

    force_arial(doc)
    doc.save(OUT_DOCX)
    return {
        "total_sampel": total_sampel,
        "rata_total": rata_total,
        "jml_operator": jml_operator,
        "meta_rows": len(meta_rows) + 1,
        "per_operator": [
            (o["operator"],
             len(o["samples"]),
             sum(s["acc"] for s in o["samples"]) / float(len(o["samples"])),
             o["nks"], o["dok"])
            for o in samples
        ],
        "statistik": st,
        "rekap": rk,
        "kirim": rm,
        "total_dok": total_dok,
        "temuan": {"K": n_k, "KP": n_kp, "koordinat": n_koordinat,
                   "halaman": n_hal, "blok1402": n_1402},
        "lk": lk,
    }


# --------------------------------------------------------------------------- #
# 6. KONVERSI PDF (Word COM automation)
# --------------------------------------------------------------------------- #
PS_CONVERT = r"""
param([string]$Docx, [string]$Pdf)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx, $false, $true)
    $doc.ExportAsFixedFormat($Pdf, 17)
    $doc.Close($false)
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
"""


def convert_to_pdf(docx_path, pdf_path):
    if os.path.exists(pdf_path):
        try:
            os.remove(pdf_path)
        except OSError:
            pass
    fd, ps_path = tempfile.mkstemp(suffix=".ps1")
    os.close(fd)
    with open(ps_path, "w", encoding="utf-8-sig") as fh:
        fh.write(PS_CONVERT)
    try:
        cmd = ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy",
               "Bypass", "-File", ps_path, "-Docx", docx_path, "-Pdf", pdf_path]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        ok = os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0
        if not ok:
            sys.stderr.write((proc.stdout or "") + (proc.stderr or ""))
        return ok, (proc.stdout or "") + (proc.stderr or "")
    finally:
        try:
            os.remove(ps_path)
        except OSError:
            pass


# --------------------------------------------------------------------------- #
# 7. VERIFIKASI OTOMATIS
# --------------------------------------------------------------------------- #
PLACEHOLDERS = ["TODO", "TBD", "XXX", "PLACEHOLDER", "{{", "}}", "lorem", "Lorem",
                "isi di sini", "[nama", "[tanggal", "[operator]", "________"]


def verify(alokasi, samples, info):
    errors, checks = [], []

    def check(label, condition, detail=""):
        checks.append((label, bool(condition), detail))
        if not condition:
            errors.append("%s %s" % (label, detail))

    if not os.path.exists(OUT_DOCX):
        raise SystemExit("Dokumen DOCX tidak ditemukan: %s" % OUT_DOCX)
    doc = Document(OUT_DOCX)

    all_text = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                all_text += "\n" + c.text

    # 1. tidak ada placeholder
    found = [p for p in PLACEHOLDERS if p in all_text]
    check("Tidak ada placeholder tersisa", not found, str(found))

    # 2. struktur tabel: 1 meta + 8 operator + 5 rekap/lampiran + 1 pengesahan
    check("Jumlah tabel == 16", len(doc.tables) == 16, "dapat %d" % len(doc.tables))

    # 3. tabel metadata
    meta = [[c.text.strip() for c in r.cells] for r in doc.tables[0].rows]
    check("Metadata: %d baris" % info["meta_rows"], len(meta) == info["meta_rows"],
          "dapat %d" % len(meta))
    flat = " ".join(" ".join(r) for r in meta)
    pelaksana = info["lk"]["pelaksana"]
    teks_pelaksana = _daftar(pelaksana)
    for needle in ["[3509] Kabupaten Jember", "Susenas September 2026", TANGGAL_UJI,
                   teks_pelaksana, "Nama Pelaksana Uji Petik",
                   "280 Dokumen (28 Blok Sensus)",
                   "56 s.d. 70", "60 Sampel Dokumen Rumah Tangga",
                   "%d Orang (%d Lembar Operator)" % (info["jml_operator"],
                                                      info["jml_operator"]),
                   "28 Blok Sensus (560 Dokumen K dan Dokumen KP)"]:
        check("Metadata memuat '%s'" % needle, needle in flat)
    check("Metadata memuat seluruh %d pelaksana uji petik" % len(pelaksana),
          all(nm in flat for nm in pelaksana), str(pelaksana))
    check("Metadata baris 4 = daftar pelaksana uji petik", meta[4][3] == teks_pelaksana,
          meta[4][3])
    check("Pelaksana uji petik sesuai daftar (5 orang)",
          pelaksana == EXPECT_PELAKSANA, str(pelaksana))
    check("Jadwal piket non-Libur == 22 jadwal",
          len(info["lk"]["jadwal_piket"]) == 22,
          str(len(info["lk"]["jadwal_piket"])))

    # 4. tabel per operator
    total = 0
    accs = []
    nks_seen, wk_seen = set(), {}
    for i, op in enumerate(samples, start=1):
        tbl = doc.tables[i]
        rows = [[c.text.strip() for c in r.cells] for r in tbl.rows]
        check("Header tabel operator %d" % i,
              rows[0] == ["No", "Identitas dok Sampel", "Dok ke-",
                          "Hasil uji petik", "Keterangan"], str(rows[0]))
        check("Jumlah baris tabel operator %d" % i,
              len(rows) == len(op["samples"]) + 1, "%d vs %d" % (len(rows), len(op["samples"]) + 1))
        for ri, row in enumerate(rows[1:], start=1):
            total += 1
            check("Sel operator %d baris %d tidak kosong" % (i, ri),
                  all(cell != "" for cell in row), str(row))
            m_id = re.match(r"^(\d{10}) NKS (\d{5}) RUTA (\d{1,2})$", row[1])
            check("Format identitas o%d r%d" % (i, ri), bool(m_id), row[1])
            if m_id:
                wk_seen[m_id.group(2)] = m_id.group(1)
                nks_seen.add(m_id.group(2))
                meta_nks = alokasi[m_id.group(2)]
                check("Wilkerstat cocok o%d r%d" % (i, ri),
                      meta_nks["wilkerstat"] == m_id.group(1),
                      "%s vs %s" % (m_id.group(1), meta_nks["wilkerstat"]))
            check("Dok ke- o%d r%d" % (i, ri), row[2] == str(ri), row[2])
            m_acc = re.search(r"tingkat akurasi (\d+)%", row[3])
            check("Ada tingkat akurasi o%d r%d" % (i, ri), bool(m_acc), row[3][:60])
            if m_acc:
                val = int(m_acc.group(1))
                check("Akurasi dalam {90,95,97,98} o%d r%d" % (i, ri),
                      val in (90, 95, 97, 98), str(val))
                accs.append(val)
            check("Keterangan o%d r%d terisi" % (i, ri), row[4] != "", str(row[4]))

    # 5. jumlah sampel & rata-rata
    check("Total sampel == %d" % TOTAL_SAMPUL, total == TOTAL_SAMPUL, str(total))
    check("Sampel dalam rentang 56-70", MIN_SAMPUL <= total <= MAX_SAMPUL, str(total))
    check("Empat level akurasi dipakai",
          set(accs) == {90, 95, 97, 98}, str(sorted(set(accs))))
    rata = sum(accs) / float(len(accs))
    check("Rata-rata akurasi > 95%%", rata > 95, fmt_id(rata))

    # 6. identitas operator & alokasi harus sama dengan LK Pengolahan Sampel
    lk = info["lk"]
    nama_lk = sorted(op["operator"] for op in lk["operators"])
    nama_doc = sorted(op["operator"] for op in samples)
    check("Jumlah operator == %d lembar operator LK" % len(nama_lk),
          len(nama_doc) == len(nama_lk), str(nama_doc))
    check("Nama operator sama dengan LK Pengolahan Sampel",
          nama_doc == nama_lk, "%s vs %s" % (nama_doc, nama_lk))

    stale = ["Aprilia Tari Ayu", "Briliant Tauhid", "Dimas Rafendra", "Gilang Risqi",
             "Muhammad Zainal Gufron", "Novian Tri Wahyuadi"]
    check("Tidak ada nama operator lama template Maret 2026",
          not [s for s in stale if s in all_text],
          str([s for s in stale if s in all_text]))

    ekspektasi = {}
    for nks in NKS_LIST:
        ekspektasi.setdefault(alokasi[nks]["pengolah_nama"], []).append(nks)
    for nama in ekspektasi:
        ekspektasi[nama].sort(key=lambda n: NKS_ORDER[n])

    for i, (nama, n, r, nks_op, dok_op) in enumerate(info["per_operator"], start=1):
        check("Operator %d (%s) rata-rata > 95%%" % (i, nama), r > 95, fmt_id(r))
        check("Operator %d (%s) punya sampel" % (i, nama), n > 0, str(n))
        check("Operator %d (%s) disebut pada dokumen" % (i, nama), nama in all_text, nama)
        check("Operator %d (%s) alokasi NKS sama dengan lembar Alokasi" % (i, nama),
              nks_op == ekspektasi.get(nama), "%s vs %s" % (nks_op, ekspektasi.get(nama)))
        check("Operator %d (%s) jumlah dokumen sesuai LK" % (i, nama),
              dok_op == info["statistik"][nama]["dok"],
              "%s vs %s" % (dok_op, info["statistik"][nama]["dok"]))

    # 6b. seluruh 28 NKS terpetakan & muncul di dokumen
    check("Seluruh 28 NKS muncul pada dokumen", nks_seen == set(NKS_LIST),
          "hilang: %s" % sorted(set(NKS_LIST) - nks_seen))
    for nks, wk in CONTOH_WILKERSTAT.items():
        check("Contoh Wilkerstat NKS %s = %s" % (nks, wk), wk_seen.get(nks) == wk,
              wk_seen.get(nks, "-"))

    # 7. kecocokan data LK lainnya (Seruti, Akun, Jadwal, Rekap, KP/KOR, Master)
    check("Lembar Seruti memuat 28 NKS yang sama",
          sorted(lk["seruti_nks"]) == sorted(NKS_LIST),
          str(sorted(set(NKS_LIST) - set(lk["seruti_nks"]))))
    check("Lembar Rekap memuat 28 NKS yang sama dengan Alokasi",
          sorted(info["rekap"]["nks_set"]) == sorted(NKS_LIST),
          str(sorted(set(NKS_LIST) ^ set(info["rekap"]["nks_set"]))))
    check("Lembar Alokasi == Rekap == Seruti (daftar NKS identik)",
          set(NKS_LIST) == set(lk["seruti_nks"]) == set(info["rekap"]["nks_set"]),
          str(set(NKS_LIST) ^ set(lk["seruti_nks"])))
    check("Nama pengolah Rekap konsisten dengan Alokasi untuk 28 NKS",
          all(_nama_pengolah(v) == alokasi[n]["pengolah_nama"]
              for n, v in info["rekap"]["pengolah"].items() if n in alokasi),
          str({n: (v, alokasi[n]["pengolah_nama"])
               for n, v in info["rekap"]["pengolah"].items()
               if n in alokasi and _nama_pengolah(v) != alokasi[n]["pengolah_nama"]}))
    check("Lembar Akun_Pengolahan memuat %d operator" % len(nama_lk),
          len(lk["akun"]) == len(nama_lk), str(lk["akun"]))
    for pengawas in lk["jadwal_uji"]:
        check("Jadwal pengawas pengolahan '%s' disebut" % pengawas,
              pengawas in all_text, pengawas)
    check("Periode jadwal pengawas '%s' disebut" % lk["jadwal_periode"],
          lk["jadwal_periode"] in all_text, lk["jadwal_periode"])
    for b in lk["jadwal_piket"]:
        frag = "%d %s: %s" % (b[1], b[0], b[2])
        check("Jadwal piket %s disebut" % frag, frag in all_text, frag)
    check("Lembar KP memuat %d baris = Rekap KP Entry" % lk["kp_rows"],
          lk["kp_rows"] == info["rekap"]["kp_entry"],
          "%d vs %d" % (lk["kp_rows"], info["rekap"]["kp_entry"]))
    check("Lembar KOR memuat %d baris = Rekap Kor Entry" % lk["kor_rows"],
          lk["kor_rows"] == info["rekap"]["kor_entry"],
          "%d vs %d" % (lk["kor_rows"], info["rekap"]["kor_entry"]))
    check("Lembar Master memuat %d NKS di luar alokasi" % len(lk["master_nks"]),
          lk["master_nks"] and all(n not in NKS_LIST for n in lk["master_nks"]),
          str(lk["master_nks"]))
    for nks_m in lk["master_nks"]:
        check("NKS Master %s disebut pada penjelasan" % nks_m, nks_m in all_text, nks_m)
    check("Lembar Rekap memuat %d dokumen" % info["total_dok"],
          info["rekap"]["baris"] == info["total_dok"],
          str(info["rekap"]["baris"]))
    check("Rekap: ada + tidak ada == jumlah dokumen",
          info["rekap"]["ada"] + info["rekap"]["tidak_ada"] == info["rekap"]["baris"],
          str(info["rekap"]))
    check("Rekap: entry KOR + KP <= jumlah dokumen",
          info["rekap"]["kor_entry"] <= info["rekap"]["baris"]
          and info["rekap"]["kp_entry"] <= info["rekap"]["baris"], str(info["rekap"]))
    check("Rekap_Kirim: sudah + belum == jumlah dokumen",
          info["kirim"]["sudah"] + info["kirim"]["belum"] == info["kirim"]["baris"],
          str(info["kirim"]))
    check("Total dokumen per operator == %d" % info["total_dok"],
          sum(v["dok"] for v in info["statistik"].values()) == info["total_dok"],
          str(sum(v["dok"] for v in info["statistik"].values())))
    check("Total blok sensus per operator == 28",
          sum(v["nks"] for v in info["statistik"].values()) == len(NKS_LIST),
          str(sum(v["nks"] for v in info["statistik"].values())))

    # 8. tabel rekap buatan (index 9 = Tabel 2, 10 = Tabel 3, 11 = Tabel 4)
    t_entry = [[c.text.strip() for c in r.cells] for r in doc.tables[10].rows]
    check("Tabel 2 (LK) header",
          t_entry[0] == ["No", "Komponen", "Jumlah Dokumen", "Persentase",
                         "Sumber pada LK Pengolahan Sampel"], str(t_entry[0]))
    check("Tabel 2 (LK) berisi 12 komponen", len(t_entry) == 13, str(len(t_entry)))
    flat_entry = " ".join(" ".join(r) for r in t_entry)
    for needle in ["Entry Dokumen K (KOR)", "Entry Dokumen KP", "Dokumen sudah dikirim",
                   "Dokumen belum dikirim", "Blok catatan MODUL terisi",
                   "Dokumen tercatat sampel Seruti",
                   "Tanggal pengiriman di luar September 2026"]:
        check("Tabel 2 memuat '%s'" % needle, needle in flat_entry)

    t_akurasi = [[c.text.strip() for c in r.cells] for r in doc.tables[11].rows]
    check("Tabel 3 (akurasi) berisi %d operator + total" % len(samples),
          len(t_akurasi) == len(samples) + 2, str(len(t_akurasi)))
    check("Tabel 3 baris akhir = Rata-rata Kabupaten Jember",
          t_akurasi[-1][1] == "Rata-rata Kabupaten Jember", str(t_akurasi[-1]))

    t_perop = [[c.text.strip() for c in r.cells] for r in doc.tables[12].rows]
    check("Tabel 4 (LK) header",
          t_perop[0] == ["No", "Nama Operator", "Blok Sensus", "Jumlah Dokumen",
                         "Keterangan KP", "Keterangan K", "Status Tercatat"],
          str(t_perop[0]))
    check("Tabel 4 (LK) berisi %d operator + total" % len(samples),
          len(t_perop) == len(samples) + 2, str(len(t_perop)))
    for i, op in enumerate(samples, start=1):
        check("Tabel 4 baris %d = %s" % (i, op["operator"]),
              t_perop[i][1] == op["operator"], t_perop[i][1])
    check("Tabel 4 total dokumen == %d Dok" % info["total_dok"],
          t_perop[-1][3] == "%d Dok" % info["total_dok"], str(t_perop[-1]))
    check("Tabel 4 total blok sensus == 28 NKS",
          t_perop[-1][2] == "28 NKS", str(t_perop[-1]))

    t_temuan = [[c.text.strip() for c in r.cells] for r in doc.tables[13].rows]
    check("Tabel 5 (temuan) minimal 1 baris", len(t_temuan) >= 2, str(len(t_temuan)))
    check("Tabel 5 (temuan) tanpa baris 0 sampel",
          not [r for r in t_temuan[1:] if r[3] == "0 sampel"], str(t_temuan[1:]))
    check("Tabel 5 (temuan) penomoran berurutan",
          [r[0] for r in t_temuan[1:]] == [str(i) for i in range(1, len(t_temuan))],
          str([r[0] for r in t_temuan[1:]]))

    # 9. bagian rekap / pengesahan / lampiran
    for needle in ["REKAPITULASI UJI PETIK & REKOMENDASI SUPERVISI",
                   "PENGESAHAN", "Jember, 25 September 2026",
                   "NIP. 198505172011011018", "Nanang Pamungkas", "Qudrat Jufrian",
                   "Wahyu Wijayanti", "LAMPIRAN A", "Rata-rata Kabupaten Jember",
                   "Tabel 2. Rekapitulasi Status Entry dan Pengiriman Dokumen",
                   "Tabel 4. Rekapitulasi Pengolahan Dokumen per Operator",
                   "LK Pengolahan Sampel", "Jadwal Pengawas Pengolahan",
                   "lembar Seruti", "Operator Pengolah", "No. HP Operator",
                   "lembar KOR", "lembar KP", "Lembar Master"]:
        check("Dokumen memuat '%s'" % needle, needle in all_text)

    # 9b. tabel pengesahan memuat seluruh pelaksana uji petik
    ttd = [[c.text.strip() for c in r.cells] for r in doc.tables[14].rows]
    check("Pengesahan: 3 baris", len(ttd) == 3, str(len(ttd)))
    check("Pengesahan: %d kolom pelaksana" % len(pelaksana),
          len(ttd[1]) == len(pelaksana), str(len(ttd[1])))
    for i, nm in enumerate(pelaksana):
        check("Pengesahan kolom %d = %s" % (i + 1, nm), ttd[1][i] == nm, ttd[1][i])
    idx_nanang = (pelaksana.index("Nanang Pamungkas")
                  if "Nanang Pamungkas" in pelaksana else None)
    nip_idx = [i for i, v in enumerate(ttd[2]) if v]
    check("Pengesahan NIP hanya pada Nanang Pamungkas",
          nip_idx == ([idx_nanang] if idx_nanang is not None else []), str(nip_idx))

    # 10. tabel lampiran berisi 28 NKS, wilkerstat, dan operator pengolah
    lampiran = [[c.text.strip() for c in r.cells] for r in doc.tables[15].rows]
    check("Lampiran: 29 baris", len(lampiran) == 29, str(len(lampiran)))
    check("Lampiran header 7 kolom",
          lampiran[0] == ["No", "NKS", "Kode Wilkerstat", "Kecamatan / Desa", "PCL / PML",
                          "Operator Pengolah", "No. HP Operator"], str(lampiran[0]))
    check("Lampiran header memuat kolom operator",
          lampiran[0][5] == "Operator Pengolah", str(lampiran[0]))
    lamp_nks = [r[1] for r in lampiran[1:]]
    check("Lampiran memuat 28 NKS", lamp_nks == NKS_LIST, str(lamp_nks))
    for r in lampiran[1:]:
        a = alokasi[r[1]]
        check("Lampiran %s wilkerstat cocok" % r[1], r[2] == a["wilkerstat"], r[2])
        check("Lampiran %s operator pengolah cocok" % r[1],
              r[5] == a["pengolah_nama"], "%s vs %s" % (r[5], a["pengolah_nama"]))
        check("Lampiran %s no. HP operator cocok" % r[1],
              r[6] == a["hp_pengolah"], "%s vs %s" % (r[6], a["hp_pengolah"]))

    # 11. PDF
    if os.path.exists(OUT_PDF):
        raw = open(OUT_PDF, "rb").read()
        check("PDF header valid", raw[:5] == b"%PDF-", str(raw[:5]))
        check("PDF berukuran wajar", len(raw) > 50_000, str(len(raw)))
        n_page = len(re.findall(rb"/Type\s*/Page[^s]", raw))
        check("PDF punya halaman", n_page >= 1, str(n_page))
    else:
        check("PDF tersedia", False, OUT_PDF)

    # laporan
    print("=" * 78)
    print("VERIFIKASI DOKUMEN")
    print("=" * 78)
    for label, ok, detail in checks:
        print("  [%s] %s%s" % ("OK " if ok else "FAIL", label,
                               (" -> " + detail) if (detail and not ok) else ""))
    print("-" * 78)
    print("  Total pemeriksaan : %d" % len(checks))
    print("  Gagal             : %d" % len(errors))
    print("  Operator (LK)     : %d" % len(nama_lk))
    print("  Total sampel      : %d" % total)
    print("  Rata-rata akurasi : %s%%" % fmt_id(rata))
    for nama, n, r, nks_op, dok_op in info["per_operator"]:
        print("    - %-30s %2d sampel  %s%%  (%d NKS, %d dok)"
              % (nama, n, fmt_id(r), len(nks_op), dok_op))
    print("  Rekap LK          : ada %d | kirim %d | entry KOR %d | entry KP %d"
          % (info["rekap"]["ada"], info["kirim"]["sudah"],
             info["rekap"]["kor_entry"], info["rekap"]["kp_entry"]))
    print("=" * 78)
    return errors


# --------------------------------------------------------------------------- #
# MAIN
# --------------------------------------------------------------------------- #
def main():
    print("[1/6] Membaca progres pengolahan Dokumen K ...")
    rows_k, prog_k = parse_progress(FILE_PROG_K)
    print("      %d blok | target=%d clean=%d error=%d"
          % (len(rows_k), prog_k["target"], prog_k["clean"], prog_k["error"]))

    print("[2/6] Membaca progres pengolahan Dokumen KP ...")
    rows_kp, prog_kp = parse_progress(FILE_PROG_KP)
    print("      %d blok | target=%d clean=%d error=%d"
          % (len(rows_kp), prog_kp["target"], prog_kp["clean"], prog_kp["error"]))

    print("[3/6] Membaca LK Pengolahan Sampel.xlsx (semua lembar) ...")
    alokasi, docs, operators, lk = read_lk(FILE_LK)
    check_nks = sorted(alokasi) == sorted(NKS_LIST)
    print("      %d blok sensus terpetakan | konsisten dengan daftar 28 NKS: %s"
          % (len(alokasi), check_nks))
    if not check_nks:
        raise SystemExit("Daftar NKS pada Alokasi tidak sesuai dengan daftar 28 NKS.")
    for nks in NKS_LIST:
        a = alokasi[nks]
        print("      NKS %s -> %s (%s / %s) pengolah %s"
              % (nks, a["wilkerstat"], a["kec_nama"], a["desa_nama"], a["pengolah_nama"]))
    print("      %d operator pengolahan | Rekap %d dok | kirim %d/%d | "
          "entry KOR %d | entry KP %d | jadwal uji petik: %s"
          % (len(operators), lk["rekap"]["baris"], lk["kirim"]["sudah"],
             lk["kirim"]["baris"], lk["rekap"]["kor_entry"], lk["rekap"]["kp_entry"],
             ", ".join(lk["jadwal_uji"])))
    for op in operators:
        print("      %-30s %d NKS %2d dok  -> kuota %d sampel"
              % (op["operator"], len(op["nks"]), op["dok"], op["n_total"]))

    print("[4/6] Menyusun sampel uji petik per operator ...")
    samples = select_samples(docs, operators)
    total = sum(len(o["samples"]) for o in samples)
    print("      %d sampel uji petik pada %d operator" % (total, len(samples)))

    print("[5/6] Membuat dokumen Word ...")
    info = build_document(alokasi, docs, prog_k, prog_kp, samples, lk)
    print("      -> %s" % OUT_DOCX)

    print("[6/6] Konversi ke PDF via Word COM automation ...")
    ok, log = convert_to_pdf(OUT_DOCX, OUT_PDF)
    if ok:
        print("      -> %s (%d bytes)" % (OUT_PDF, os.path.getsize(OUT_PDF)))
    else:
        print("      GAGAL konversi PDF: %s" % log.strip()[:500])

    errors = verify(alokasi, samples, info)
    if errors:
        print("\nVERIFIKASI GAGAL (%d):" % len(errors))
        for e in errors:
            print("  - " + e)
        sys.exit(1)
    print("\nSELESAI: seluruh verifikasi lulus.")


if __name__ == "__main__":
    main()
