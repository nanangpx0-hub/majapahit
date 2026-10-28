import openpyxl
from collections import defaultdict, Counter
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import copy
import re

# ================ STEP 1: EXTRACT DATA FROM EXCEL ================
wb = openpyxl.load_workbook('C:/laragon/www/majapahit/sakernas/data/PPL-1784774073385 (1).xlsx', data_only=True)
ws = wb.active

kecamatan_data = defaultdict(list)
pml_set = set()
ppl_set = set()
all_rows = []

for row in range(2, ws.max_row + 1):
    prov = ws.cell(row, 1).value
    kab = ws.cell(row, 2).value
    kec = ws.cell(row, 3).value
    desa = ws.cell(row, 4).value
    sls = ws.cell(row, 5).value
    subsls = ws.cell(row, 6).value
    pml_email = ws.cell(row, 7).value
    ppl_email = ws.cell(row, 8).value
    
    entry = {
        'prov': str(prov or ''),
        'kab': str(kab or ''),
        'kec': str(kec or ''),
        'desa': str(desa or ''),
        'sls': str(sls or ''),
        'subsls': str(subsls or ''),
        'pml': str(pml_email or ''),
        'ppl': str(ppl_email or ''),
    }
    kecamatan_data[kec].append(entry)
    if pml_email: pml_set.add(pml_email)
    if ppl_email: ppl_set.add(ppl_email)
    all_rows.append(entry)

# ================ STEP 2: DEFINE KECAMATAN MAPPING ================
KECAMATAN_MAP = {
    '010': 'Kencong', '020': 'Gumukmas', '030': 'Puger',
    '040': 'Wuluhan', '050': 'Ambulu', '060': 'Tempurejo',
    '070': 'Silo', '080': 'Mayang', '090': 'Mumbulsari',
    '100': 'Jenggawah', '110': 'Ajung', '120': 'Rambipuji',
    '140': 'Umbulsari', '160': 'Jombang', '170': 'Sumberbaru',
    '180': 'Tanggul', '190': 'Bangsalsari', '200': 'Panti',
    '230': 'Pakusari', '240': 'Kalisat', '250': 'Ledokombo',
    '260': 'Sumberjambe', '270': 'Sukowono',
    '710': 'Kaliwates', '720': 'Sumbersari', '730': 'Patrang'
}

def full_kode_kec(kec_short):
    return f'3509{kec_short}'

def nama_kecamatan(kec_short):
    return KECAMATAN_MAP.get(kec_short, f'Kec. {kec_short}')

# ================ STEP 3: COMPUTE STATISTICS ================
sorted_kec = sorted(kecamatan_data.keys(), key=lambda x: int(x))
total_sls = len(all_rows)
total_kec = len(sorted_kec)
total_pml = len(pml_set)
total_ppl = len(ppl_set)

kec_stats = {}
for k in sorted_kec:
    entries = kecamatan_data[k]
    desa_set = set(e['desa'] for e in entries)
    kec_stats[k] = {
        'nama': nama_kecamatan(k),
        'kode': full_kode_kec(k),
        'jumlah_tugas': len(entries),
        'jumlah_desa': len(desa_set),
        'pml_set': set(e['pml'] for e in entries if e['pml']),
        'ppl_set': set(e['ppl'] for e in entries if e['ppl']),
    }

pml_kec = defaultdict(set)
for k in sorted_kec:
    for e in kecamatan_data[k]:
        if e['pml']:
            pml_kec[e['pml']].add(k)

# ================ STEP 4: CREATE DOCUMENT FROM TEMPLATE ================
doc = Document('C:/laragon/www/majapahit/sakernas/contoh/3509_Laporan_Assignment_PPL_PML_SE2026-rev.docx')

def set_paragraph_text(para, new_text):
    if para.runs:
        first_run = para.runs[0]
        for run in para.runs[1:]:
            run.text = ''
        first_run.text = new_text
    else:
        para.add_run(new_text)

def replace_in_runs(para, old_text, new_text):
    for run in para.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)

# ================ STEP 5: MODIFY PARAGRAPHS BY INDEX ================
# P000-P003: empty lines (keep)
# P004-P006: Header (keep)
# P009: LAPORAN KEGIATAN (keep)
# P010: ASSIGNMENT WILAYAH PPL DAN PML (keep)
# P011: PERSIAPAN SENSUS EKONOMI 2026 -> SAKERNAS SEMESTER 2 TAHUN 2026
set_paragraph_text(doc.paragraphs[11], 'SAKERNAS SEMESTER 2 TAHUN 2026')

# P014: Periode Pelaporan
set_paragraph_text(doc.paragraphs[14], 'Periode Pelaporan  :  Semester 2 Tahun 2026 (Juli s.d. Desember)')

# P015: Lokasi (keep)
# P016: Jumlah Wilayah
set_paragraph_text(doc.paragraphs[16], f'Jumlah Wilayah  :  {total_kec} Kecamatan se-Kabupaten Jember')

# P017: Jumlah Tugas
set_paragraph_text(doc.paragraphs[17], f'Jumlah Tugas  :  {total_sls} Blok Sensus')

# P020-P022: keep

# P026: Pendahuluan paragraph 1
ref_text = doc.paragraphs[26].text
new_text = ref_text.replace('Sensus Ekonomi 2026 (SE2026)', 'Survei Angkatan Kerja Nasional Semester 2 Tahun 2026 (Sakernas Semester 2 2026)')
new_text = new_text.replace('SE2026', 'Sakernas Semester 2 2026')
new_text = new_text.replace('Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')
set_paragraph_text(doc.paragraphs[26], new_text)

# P027: Pendahuluan paragraph 2
for run in doc.paragraphs[27].runs:
    run.text = run.text.replace('SE2026', 'Sakernas Semester 2 2026')

# P036: Tujuan
for run in doc.paragraphs[36].runs:
    run.text = run.text.replace('SE2026', 'Sakernas Semester 2 2026')

# P046: Dasar hukum #3
p46 = doc.paragraphs[46]
set_paragraph_text(p46, '3. Peraturan Badan Pusat Statistik Nomor 7 Tahun 2024 tentang Penyelenggaraan Survei Angkatan Kerja Nasional (Sakernas) 2026.')

# P047: Dasar hukum #4
p47 = doc.paragraphs[47]
for run in p47.runs:
    run.text = run.text.replace('Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')

# P048: Dasar hukum #5
p48 = doc.paragraphs[48]
for run in p48.runs:
    run.text = run.text.replace('Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')

# P052: Ruang Lingkup - change 31 -> 26
p52 = doc.paragraphs[52]
for run in p52.runs:
    run.text = run.text.replace('31 kecamatan', f'{total_kec} kecamatan')
    run.text = run.text.replace('31 Kecamatan', f'{total_kec} Kecamatan')

# P062: Waktu dan Tempat Pelaksanaan 
p62 = doc.paragraphs[62]
for run in p62.runs:
    run.text = run.text.replace('April sampai dengan Juni 2026', 'Juli sampai dengan Desember 2026')

# P067: Summary after table 1
p67 = doc.paragraphs[67]
for run in p67.runs:
    run.text = run.text.replace('31 kecamatan', f'{total_kec} kecamatan')
    run.text = run.text.replace('31 Kecamatan', f'{total_kec} Kecamatan')
    run.text = run.text.replace('1.156.276', str(total_sls))

# P076: Kesimpulan intro
p76 = doc.paragraphs[76]
for run in p76.runs:
    run.text = run.text.replace('Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')

# P078: Kesimpulan bullet 1
p78 = doc.paragraphs[78]
for run in p78.runs:
    run.text = run.text.replace('31 kecamatan', f'{total_kec} kecamatan')

# P079: Kesimpulan bullet 2
p79 = doc.paragraphs[79]
for run in p79.runs:
    run.text = run.text.replace('1.156.276', str(total_sls))

# P080: Kesimpulan bullet 3
p80 = doc.paragraphs[80]
for run in p80.runs:
    run.text = run.text.replace('7 PML', f'{total_pml} PML')
    run.text = run.text.replace('15 PPL', f'{total_ppl} PPL')

# P083: Rekomendasi intro
p83 = doc.paragraphs[83]
for run in p83.runs:
    run.text = run.text.replace('SE2026', 'Sakernas Semester 2 2026')

# P096: Lampiran description
p96 = doc.paragraphs[96]
for run in p96.runs:
    run.text = run.text.replace('SE2026', 'Sakernas Semester 2 2026')

# ================ STEP 6: MODIFY TABLE 0 - INFORMASI UMUM ================
table0 = doc.tables[0]
table0.rows[2].cells[2].text = 'Sakernas Semester 2 Tahun 2026'
table0.rows[4].cells[2].text = 'Juli s.d. Desember 2026'
table0.rows[6].cells[2].text = f'{total_kec} Kecamatan'
table0.rows[7].cells[2].text = f'{total_sls} Blok Sensus'
table0.rows[8].cells[2].text = '100% Terselesaikan'

# ================ STEP 7: MODIFY TABLE 1 - DAFTAR WILAYAH ASSIGNMENT ================
table1 = doc.tables[1]
while len(table1.rows) > 1:
    table1._tbl.remove(table1.rows[1]._tr)
for idx, k in enumerate(sorted_kec, 1):
    stats = kec_stats[k]
    row = table1.add_row()
    row.cells[0].text = str(idx)
    row.cells[1].text = full_kode_kec(k)
    row.cells[2].text = stats['nama']
    row.cells[3].text = str(stats['jumlah_tugas'])
    row.cells[4].text = 'Terselesaikan'

# ================ STEP 8: MODIFY TABLE 2 - HASIL DAN CAPAIAN ================
table2 = doc.tables[2]
table2.rows[1].cells[1].text = f'{total_kec} Kecamatan'
table2.rows[2].cells[1].text = str(total_sls)
table2.rows[3].cells[1].text = f'{total_kec} Kecamatan (100%)'
table2.rows[4].cells[1].text = '0 Kecamatan (0%)'
table2.rows[5].cells[1].text = f'{total_pml} Orang'
table2.rows[6].cells[1].text = f'{total_ppl} Orang'

# ================ STEP 9: MODIFY TABLE 3 - TANDA TANGAN ================
table3 = doc.tables[3]
# Update the date in the signature block
cell_text = table3.rows[0].cells[1].text
cell_text = cell_text.replace('Juni 2026', 'Desember 2026')
table3.rows[0].cells[1].text = cell_text

# ================ STEP 10: MODIFY TABLE 4 - LAMPIRAN ================
table4 = doc.tables[4]
while len(table4.rows) > 1:
    table4._tbl.remove(table4.rows[1]._tr)
for idx, k in enumerate(sorted_kec, 1):
    stats = kec_stats[k]
    row = table4.add_row()
    row.cells[0].text = str(idx)
    row.cells[1].text = full_kode_kec(k)
    row.cells[2].text = stats['nama']
    row.cells[3].text = str(stats['jumlah_tugas'])
    row.cells[4].text = '0'
    row.cells[5].text = 'SELESAI'

# ================ STEP 11: FINAL CLEANUP - REPLACE ANY REMAINING OLD TERMS ================
for para in doc.paragraphs:
    for run in para.runs:
        run.text = run.text.replace('SE2026', 'Sakernas Semester 2 2026')
        run.text = run.text.replace('Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')
        run.text = run.text.replace('sensus ekonomi', 'Survei Angkatan Kerja Nasional (Sakernas)')

for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            # Only update text in cells that have old terminology
            if 'SE2026' in cell.text or 'Sensus Ekonomi 2026' in cell.text:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.text = run.text.replace('SE2026', 'Sakernas Semester 2 2026')
                        run.text = run.text.replace('Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')

# ================ STEP 12: SAVE ================
output_path = 'C:/laragon/www/majapahit/sakernas/3509_Laporan_Assignment_PPL_PML_Sakernas_2026.docx'
doc.save(output_path)
print(f"Document saved to: {output_path}")
print(f"\n=== LAPORAN SUMMARY ===")
print(f"Kegiatan: Sakernas Semester 2 Tahun 2026")
print(f"Lokasi: BPS Kabupaten Jember (3509)")
print(f"Jumlah Kecamatan: {total_kec}")
print(f"Jumlah Tugas (SLS): {total_sls}")
print(f"Jumlah PML: {total_pml}")
print(f"Jumlah PPL: {total_ppl}")
