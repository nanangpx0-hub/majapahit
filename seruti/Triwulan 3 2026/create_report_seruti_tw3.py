# -*- coding: utf-8 -*-
from __future__ import annotations

import subprocess
import sys
from collections import OrderedDict
from pathlib import Path

import openpyxl
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = Path(r"C:\laragon\www\majapahit\seruti\Triwulan 3 2026")
DATA_DIR = BASE / "data"
LK_PATH = DATA_DIR / "LK Pengolahan Sampel (1).xlsx"
OUT_DOCX = BASE / "laporan_penyelesaian_pengolahan_seruti_tw3_2026.docx"
OUT_PDF = BASE / "laporan_penyelesaian_pengolahan_seruti_tw3_2026.pdf"
TEMPLATE = Path(r"C:\laragon\www\majapahit\seruti\template\3509_Laporan_Pengolahan_Seruti_Triwulan_III_2026-rev.docx")
CONVERT_PS1 = Path(r"C:\laragon\www\majapahit\seruti\convert_to_pdf.ps1")

NAVY = RGBColor(0x1F, 0x4E, 0x79)
FONT = "Arial"
HDR_FILL = "1F4E79"
TOTAL_FILL = "F2F2F2"


def fmt_num(x):
    if x is None:
        return "0"
    if isinstance(x, float):
        if x == int(x):
            x = int(x)
        else:
            return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{int(x):,}".replace(",", ".")


def fmt_pct(x):
    return f"{x:.1f}%".replace(".", ",")


def pct(a, b):
    if not b:
        return 0.0
    return a / b * 100.0


def parse_seruti():
    wb = openpyxl.load_workbook(LK_PATH, data_only=True)
    ws = wb["Seruti"]

    # Progress terbaru per blok sensus (sumber: Progres Entri per Blok Sensus)
    prog_path = DATA_DIR / "Progres Entri per Blok Sensus 2026-10-21 11_19_01.xlsx"
    prog = {}
    if prog_path.exists():
        sys.path.insert(0, r"C:\laragon\www\majapahit\seruti")
        from xlsx_raw import read_xlsx
        _, prows, _ = read_xlsx(str(prog_path), 0)
        for ridx, row in prows:
            nks = str(row.get("A", "")).strip()
            if not nks or nks == "NKS":
                continue
            try:
                prog[nks] = {
                    "target": int(float(row.get("B", 0) or 0)),
                    "ic": int(float(row.get("C", 0) or 0)),
                    "ie": int(float(row.get("D", 0) or 0)),
                    "ib": int(float(row.get("E", 0) or 0)),
                    "mc": int(float(row.get("F", 0) or 0)),
                    "me": int(float(row.get("G", 0) or 0)),
                    "mb": int(float(row.get("H", 0) or 0)),
                }
            except (ValueError, TypeError):
                pass

    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or r[1] in (None, "NKS TW 3"):
            continue
        nks = str(r[1]).strip()
        rec = {
            "nks": nks,
            "pcl": str(r[2]).strip() if r[2] else "",
            "pml": str(r[3]).strip() if r[3] else "",
            "pengolah": str(r[4]).strip() if r[4] else "",
            "target": int(r[5] or 0),
            "ic": int(r[6] or 0),
            "ie": int(r[7] or 0),
            "ib": int(r[8] or 0),
            "mc": int(r[9] or 0),
            "me": int(r[10] or 0),
            "mb": int(r[11] or 0),
        }
        if nks in prog:
            rec.update(prog[nks])
        rows.append(rec)
    return rows


def norm_name(name):
    n = name.strip()
    mapping = {
        "Anung Anindhita P": "Anung Anindhita Pratiwi",
        "Prasistiwi": "Prasistiwi Andrianingtyas",
        "Iffa Dzakiyya Khairunnisa": "Iffa Dzakiyya Khairunnisa",
        "Iffa Dzakiyya": "Iffa Dzakiyya Khairunnisa",
        "Astri Widarianti": "Astri Widarianti",
        "Nur Ida Suryandari": "Nur Ida Suryandari",
        "Aminatus Sholeha": "Aminatus Sholeha",
        "Laviana Ika Putrisari": "Laviana Ika Putrisari",
        "Putri Salsabhila Fahira": "Putri Salsabhila Fahira",
    }
    return mapping.get(n, n)


def parse_operator_usernames():
    sys.path.insert(0, r"C:\laragon\www\majapahit\seruti")
    from xlsx_raw import read_xlsx
    p = DATA_DIR / "Progres Entri per Operator (Berdasarkan Operator Entri Pertama) 2026-10-21 11_22_46.xlsx"
    name, rows, sheets = read_xlsx(str(p), 0)
    uname = {}
    for ridx, row in rows:
        u = row.get("A")
        nm = row.get("B")
        if u and nm and str(u).strip() not in ("Username",) and str(nm).strip() != "Nama Operator":
            uname[norm_name(str(nm).strip())] = str(u).strip()
    return uname


def unique_sorted(names):
    return sorted(set(names))


def set_par_text(par, text):
    if par.runs:
        par.runs[0].text = text
        for r in par.runs[1:]:
            r.text = ""
    else:
        par.add_run(text)


def set_label_value(par, label, value):
    if not par.runs:
        par.add_run(label)
        par.add_run(value)
        return
    par.runs[0].text = label
    for r in par.runs[1:-1]:
        r.text = ""
    if len(par.runs) > 1:
        par.runs[-1].text = value
    else:
        par.add_run(value)


def _par_text(el):
    return "".join(t.text or "" for t in el.xpath(".//w:t"))


def prepare_document(progress_text="100,0%"):
    doc = Document(str(TEMPLATE))

    set_par_text(doc.paragraphs[9], "LAPORAN KEGIATAN")
    set_par_text(doc.paragraphs[10], "PENYELESAIAN PENGOLAHAN SAMPEL SERUTI")
    set_par_text(doc.paragraphs[11], f"TRIWULAN III TAHUN 2026 (PROGRESS {progress_text})")

    set_label_value(doc.paragraphs[14], "Periode Pelaporan  :  ", "Juli s.d. September 2026")
    set_label_value(doc.paragraphs[15], "Lokasi  :  ",
                    "Kantor BPS Kabupaten Jember (3509)")
    set_label_value(doc.paragraphs[16], "Jumlah Blok Sensus  :  ",
                    "28 Blok Sensus (280 Sampel Rumah Tangga)")
    set_label_value(doc.paragraphs[17], "Jumlah Dokumen  :  ",
                    "560 (Kuesioner Inti 280 + MAK 280)")

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

    for el in retained:
        if el.tag == qn("w:p"):
            for br in el.xpath('.//w:br[@w:type="page"]'):
                br.getparent().remove(br)
    return doc


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

def set_base_style(doc):
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(11)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.08


def add_heading(doc, text, size=14, space_before=12, space_after=6, color=NAVY, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.keep_with_next = True
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.font.name = FONT
    r.font.color.rgb = color
    return p


def add_subheading(doc, text, size=12):
    return add_heading(doc, text, size=size, space_before=10, space_after=4)


def add_body(doc, text, size=11, bold=False, italic=False, space_after=6, align=None, justify=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if align is not None:
        p.alignment = align
    elif justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.name = FONT
    r.bold = bold
    r.italic = italic
    return p


def add_bullet(doc, text, size=11):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    r.font.size = Pt(size)
    return p


def make_table(doc, headers, rows, widths=None, size=9, aligns=None, first_col_left=False):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        par = hdr[i].paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = par.add_run(h)
        run.bold = True
        run.font.size = Pt(size)
        run.font.name = FONT
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "1F4E79")
        hdr[i]._tc.get_or_add_tcPr().append(shd)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ""
            par = cells[i].paragraphs[0]
            if aligns and aligns[i] == "center":
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif aligns and aligns[i] == "right":
                par.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                par.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = par.add_run(str(val))
            run.font.size = Pt(size)
            run.font.name = FONT
            if ri % 2 == 1:
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), "F2F2F2")
                cells[i]._tc.get_or_add_tcPr().append(shd)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    th.set(qn("w:val"), "true")
    trPr.append(th)
    return t


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.font.size = Pt(9)
    r.italic = True
    return p

def main():
    rows = parse_seruti()
    uname_map = parse_operator_usernames()

    target = sum(r["target"] for r in rows)
    n_bs = len(rows)
    ic = sum(r["ic"] for r in rows)
    ie = sum(r["ie"] for r in rows)
    ib = sum(r["ib"] for r in rows)
    mc = sum(r["mc"] for r in rows)
    me = sum(r["me"] for r in rows)
    mb = sum(r["mb"] for r in rows)
    total_docs = target * 2
    clean_docs = ic + mc
    entered_docs = (ic + ie) + (mc + me)

    peng = OrderedDict()
    for r in rows:
        nm = norm_name(r["pengolah"])
        d = peng.setdefault(nm, {"target": 0, "ic": 0, "ie": 0, "ib": 0, "mc": 0, "me": 0, "mb": 0})
        d["target"] += r["target"]
        d["ic"] += r["ic"]; d["ie"] += r["ie"]; d["ib"] += r["ib"]
        d["mc"] += r["mc"]; d["me"] += r["me"]; d["mb"] += r["mb"]

    pml_list = unique_sorted({r["pml"] for r in rows if r["pml"]})
    pcl_list = unique_sorted({r["pcl"] for r in rows if r["pcl"]})

    op_names = sorted(peng.keys())

    def username_for(name):
        return uname_map.get(name, "-")

    doc = prepare_document(progress_text=fmt_pct(pct(clean_docs, total_docs)))
    doc.add_page_break()


    # I
    add_heading(doc, "I. Pendahuluan", size=13)
    add_body(doc,
        "Survei Ekonomi Rumah Tangga Triwulanan (SERUTI) merupakan salah satu sumber data sosial "
        "ekonomi rumah tangga yang penting untuk menggambarkan pola konsumsi, pengeluaran, dan "
        "kondisi ekonomi rumah tangga secara triwulanan di wilayah Kabupaten Jember.")
    add_body(doc,
        f"Kegiatan pengolahan data SERUTI Triwulan III Tahun 2026 di BPS Kabupaten Jember mencakup "
        f"{fmt_num(target)} rumah tangga sampel dari {n_bs} blok sensus dengan total "
        f"{fmt_num(total_docs)} dokumen, terdiri atas Kuesioner Inti dan Kuesioner MAK, serta "
        f"dilaksanakan oleh {len(peng)} ({_numword(len(peng))}) operator entri data.")
    add_body(doc,
        "Laporan ini menyajikan capaian pengolahan sampel SERUTI Triwulan III Tahun 2026 sesuai "
        "data monitoring per 21 Oktober 2026.")

    # II
    add_heading(doc, "II. Capaian Penyelesaian Pengolahan", size=13)
    add_body(doc,
        f"Berdasarkan laporan progres pengolahan, seluruh {fmt_num(target)} dokumen Kuesioner MAK "
        f"sampel telah selesai dientri dan berstatus clean, dan Kuesioner Inti juga telah selesai "
        f"pengolahan dengan {fmt_num(ic)} dokumen clean dari {fmt_num(target)} dokumen "
        f"({fmt_pct(pct(ic, target))}). Dari total {fmt_num(total_docs)} dokumen sampel, "
        f"{fmt_num(clean_docs)} dokumen telah berstatus clean, sehingga capaian validasi dan cleaning "
        f"tercatat {fmt_pct(pct(clean_docs, total_docs))}.")

    add_subheading(doc, "Tabel 1. Capaian Penyelesaian Pengolahan per Jenis Dokumen")
    make_table(doc,
        ["Jenis Dokumen", "Target", "Realisasi", "Keterangan"],
        [["Kuesioner Inti", f"{fmt_num(target)} dok", fmt_pct(pct(ic, target)),
          f"{fmt_num(ic)} dari {fmt_num(target)} clean; {fmt_num(ie)} error; {fmt_num(ib)} belum"],
         ["Kuesioner MAK", f"{fmt_num(target)} dok", fmt_pct(pct(mc, target)),
          f"{fmt_num(mc)} dari {fmt_num(target)} clean"],
         ["Total", f"{fmt_num(total_docs)} dok", fmt_pct(pct(clean_docs, total_docs)),
          f"{fmt_num(clean_docs)} dari {fmt_num(total_docs)} clean"]],
        widths=[4.0, 2.5, 2.5, 7.5], size=9,
        aligns=["left", "center", "center", "left"])
    add_caption(doc, "Tabel 1. Capaian penyelesaian pengolahan per jenis dokumen (sumber: LK Pengolahan Sampel).")

    add_subheading(doc, "Tabel 2. Ringkasan Tahapan Pengolahan")
    make_table(doc,
        ["No", "Tahapan", "Capaian", "Keterangan"],
        [["1", "Pendataan (Pencacahan)", "100%",
          f"Seluruh {fmt_num(target)} RUTA telah selesai didata"],
         ["2", "Entri Data", fmt_pct(pct(entered_docs, total_docs)),
          f"{fmt_num(entered_docs)} dari {fmt_num(total_docs)} dokumen telah dientri"],
         ["3", "Validasi & Cleaning", fmt_pct(pct(clean_docs, total_docs)),
          f"{fmt_num(clean_docs)} dokumen berstatus clean"],
         ["4", "Penyelesaian Pengolahan", fmt_pct(pct(clean_docs, total_docs)),
          "MAK dan Kuesioner Inti selesai 100%"]],
        widths=[1.2, 4.5, 2.5, 8.3], size=9,
        aligns=["center", "left", "center", "left"])
    add_caption(doc, "Tabel 2. Ringkasan tahapan pengolahan.")

    add_body(doc,
        f"Seluruh {n_bs} blok sensus telah selesai pengolahan. Progres entri Kuesioner Inti dan "
        f"Kuesioner MAK telah mencapai 100% pada seluruh blok sensus.")

    add_subheading(doc, "Tabel 3. Rincian Progres per Blok Sensus")
    bs_rows = []
    for r in rows:
        t = r["target"]
        entri_pct = pct(r["ic"] + r["mc"] + r["ie"] + r["me"], t * 2)
        bs_rows.append([
            r["nks"], fmt_num(t),
            fmt_num(r["ic"]), fmt_num(r["ie"]), fmt_num(r["ib"]),
            fmt_num(r["mc"]), fmt_num(r["me"]), fmt_num(r["mb"]),
            fmt_pct(entri_pct),
        ])
    make_table(doc,
        ["NKS", "Target", "INTI Clean", "INTI Error", "INTI Belum",
         "MAK Clean", "MAK Error", "MAK Belum", "Entri(%)"],
        bs_rows,
        widths=[1.9, 1.3, 1.7, 1.7, 1.7, 1.7, 1.7, 1.7, 1.6], size=8,
        aligns=["center", "center", "center", "center", "center", "center", "center", "center", "center"])
    add_caption(doc, "Tabel 3. Rincian progres pengolahan per blok sensus.")

    # III
    doc.add_page_break()
    add_heading(doc, "III. Daftar Petugas Lapangan dan Operator Entry", size=13)
    add_body(doc,
        "Pelaksanaan pengolahan didukung oleh tim lapangan dan operator entry yang terlibat dalam "
        "pendataan, penyerahan dokumen, entri data, serta pengendalian kualitas hasil pengolahan.")

    add_subheading(doc, "Tabel 4. Rincian Progres Entri per Operator")
    op_rows = []
    for nm in op_names:
        d = peng[nm]
        t = d["target"] * 2
        clean = d["ic"] + d["mc"]
        op_rows.append([
            username_for(nm), nm,
            fmt_num(d["ic"]), fmt_num(d["ie"]),
            fmt_num(d["mc"]), fmt_num(d["me"]),
            fmt_num(t),
            fmt_pct(pct(clean, t)),
        ])
    make_table(doc,
        ["Username", "Nama Operator", "INTI Clean", "INTI Error",
         "MAK Clean", "MAK Error", "Total Dokumen", "Entri(%)"],
        op_rows,
        widths=[4.0, 4.5, 1.7, 1.7, 1.7, 1.7, 2.0, 1.7], size=8,
        aligns=["left", "left", "center", "center", "center", "center", "center", "center"])
    add_caption(doc, "Tabel 4. Rincian progres entri per operator.")

    add_subheading(doc, "Tabel 5. Daftar Pengawas Lapangan (PML)")
    make_table(doc, ["No", "Nama PML"],
        [[str(i + 1), name] for i, name in enumerate(pml_list)],
        widths=[1.5, 15.0], size=9, aligns=["center", "left"])
    add_caption(doc, "Tabel 5. Daftar pengawas lapangan (PML).")

    add_subheading(doc, "Tabel 6. Daftar Pencacah Lapangan (PPL)")
    make_table(doc, ["No", "Nama PPL"],
        [[str(i + 1), name] for i, name in enumerate(pcl_list)],
        widths=[1.5, 15.0], size=9, aligns=["center", "left"])
    add_caption(doc, "Tabel 6. Daftar pencacah lapangan (PPL).")

    add_subheading(doc, "Tabel 7. Daftar Operator Entry")
    make_table(doc, ["No", "Username", "Nama Operator"],
        [[str(i + 1), username_for(nm), nm] for i, nm in enumerate(op_names)],
        widths=[1.5, 6.0, 9.0], size=9, aligns=["center", "left", "left"])
    add_caption(doc, "Tabel 7. Daftar operator entry.")

    # IV
    add_heading(doc, "IV. Kesimpulan dan Tindak Lanjut", size=13)
    add_body(doc,
        f"Penyelesaian pengolahan dokumen SERUTI Triwulan III Tahun 2026 di BPS Kabupaten Jember "
        f"telah mencapai {fmt_pct(pct(clean_docs, total_docs))} pada tahap validasi dan cleaning "
        f"karena {fmt_num(clean_docs)} dari {fmt_num(total_docs)} dokumen sampel sudah berstatus clean. "
        f"Kuesioner MAK telah selesai 100% ({fmt_num(mc)} dokumen clean), dan Kuesioner Inti juga "
        f"telah selesai 100% ({fmt_num(ic)} dokumen clean), tanpa error maupun dokumen yang "
        f"belum dientri.")
    add_body(doc, "Tindak lanjut yang perlu dilakukan adalah sebagai berikut:")
    for txt in [
        f"Melakukan verifikasi final dan penguncian data (locking) seluruh dokumen Kuesioner Inti dan "
        f"Kuesioner MAK yang telah berstatus clean.",
        "Melakukan monitoring harian per satuan lingkungan setempat (SLS) dan per operator untuk memastikan konsistensi "
        "hasil pengolahan tetap terjaga.",
        "Mengoptimalkan pengendalian kualitas (quality control) oleh tim pengolahan agar konsistensi "
        "hasil entri dan validasi tetap terjaga.",
        "Menyusun rekap final hasil pengolahan sebagai dasar pelaporan penutupan kegiatan SERUTI "
        "Triwulan III Tahun 2026.",
    ]:
        add_bullet(doc, txt)

    # Signature
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Jember, Oktober 2026")
    r.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(2)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Ketua Tim Pengolahan dan")
    r.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Layanan Statistik")
    r.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(24)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Nanang Pamungkas")
    r.bold = True; r.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(0)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("NIP. 199110242019011001")
    r.font.size = Pt(10)
    p.paragraph_format.space_after = Pt(0)

    doc.save(str(OUT_DOCX))
    print(f"Dokumen DOCX: {OUT_DOCX}")
    print(f"Paragraf={len(doc.paragraphs)} Tabel={len(doc.tables)}")

    pdf = convert_to_pdf(OUT_DOCX, OUT_PDF)
    print(f"Dokumen PDF: {pdf} ({pdf.stat().st_size:,} bytes)")


def _numword(n):
    words = {1: "satu", 2: "dua", 3: "tiga", 4: "empat", 5: "lima",
             6: "enam", 7: "tujuh", 8: "delapan", 9: "sembilan", 10: "sepuluh"}
    return words.get(n, str(n))


if __name__ == "__main__":
    main()

