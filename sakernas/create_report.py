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
    """Convert 3-digit code to full 7-digit BPS code"""
    return f'3509{kec_short}'

def nama_kecamatan(kec_short):
    return KECAMATAN_MAP.get(kec_short, f'Kec. {kec_short}')

# ================ STEP 3: COMPUTE STATISTICS ================
sorted_kec = sorted(kecamatan_data.keys(), key=lambda x: int(x))
total_sls = len(all_rows)
total_kec = len(sorted_kec)
total_pml = len(pml_set)
total_ppl = len(ppl_set)

# Per-kecamatan stats
kec_stats = {}
for k in sorted_kec:
    entries = kecamatan_data[k]
    desa_set = set(e['desa'] for e in entries)
    sls_count = len(entries)
    kec_stats[k] = {
        'nama': nama_kecamatan(k),
        'kode': full_kode_kec(k),
        'jumlah_tugas': sls_count,
        'jumlah_desa': len(desa_set),
        'pml_set': set(e['pml'] for e in entries if e['pml']),
        'ppl_set': set(e['ppl'] for e in entries if e['ppl']),
    }

# PML statistics per kecamatan
pml_kec = defaultdict(set)
ppl_names = {}
for k in sorted_kec:
    for e in kecamatan_data[k]:
        if e['pml']:
            pml_kec[e['pml']].add(k)

# ================ STEP 4: CREATE THE DOCUMENT FROM TEMPLATE ================
doc = Document('C:/laragon/www/majapahit/sakernas/contoh/3509_Laporan_Assignment_PPL_PML_SE2026-rev.docx')

def replace_text_in_para(para, old_text, new_text):
    """Replace text in a paragraph while preserving formatting"""
    for run in para.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)

def set_paragraph_text(para, new_text):
    """Set full text of a paragraph preserving first run's format"""
    if para.runs:
        first_run = para.runs[0]
        for run in para.runs[1:]:
            run.text = ''
        first_run.text = new_text
    else:
        para.add_run(new_text)

# ================ STEP 5: MODIFY PARAGRAPHS ================
# Define replacements for all SE2026 -> Sakernas references
replacement_rules = {
    'PERSIAPAN SENSUS EKONOMI 2026': 'SAKERNAS SEMESTER 2 TAHUN 2026',
    'SE2026': 'Sakernas Semester 2 2026',
    'Sensus Ekonomi 2026 (SE2026)': 'Survei Angkatan Kerja Nasional Semester 2 2026 (Sakernas Semester 2 2026)',
    'Sensus Ekonomi 2026': 'Sakernas Semester 2 2026',
    'sensus ekonomi': 'Survei Angkatan Kerja Nasional (Sakernas)',
}

# Modify specific paragraphs based on content
for para in doc.paragraphs:
    text = para.text
    
    if 'Jumlah Wilayah' in text:
        set_paragraph_text(para, f'Jumlah Wilayah  :  {total_kec} Kecamatan se-Kabupaten Jember')
    elif 'Jumlah Keluarga' in text or 'Jumlah Keluarga/ Usaha Tugas' in text:
        set_paragraph_text(para, f'Jumlah Tugas  :  {total_sls} Blok Sensus')
    elif 'Januari 2026' in text or 'Juni 2026' in text or 'Periode Pelaporan' in text:
        set_paragraph_text(para, 'Periode Pelaporan  :  Juli - Desember 2026')
    elif 'Tim Pengolahan dan Layanan Statistik' in text:
        set_paragraph_text(para, 'Tim Pengolahan dan Layanan Statistik')
    elif 'JEMBER, 2026' in text or 'JEMBER, 2026' in text:
        set_paragraph_text(para, 'JEMBER, 2026')
    
    # Replace content for SE2026 -> Sakernas in all paragraphs
    for old, new in replacement_rules.items():
        if old in para.text:
            for run in para.runs:
                run.text = run.text.replace(old, new)

# ================ STEP 6: MODIFY TABLE 0 - INFORMASI UMUM ================
table0 = doc.tables[0]
# Row 2: Kegiatan
table0.rows[2].cells[2].text = 'Sakernas Semester 2 Tahun 2026'
# Row 4: Periode
table0.rows[4].cells[2].text = 'Juli s.d. Desember 2026'
# Row 6: Jumlah Kecamatan
table0.rows[6].cells[2].text = f'{total_kec} Kecamatan'
# Row 7: Total Wilayah Tugas
table0.rows[7].cells[2].text = f'{total_sls} Blok Sensus'
# Row 8: Status
table0.rows[8].cells[2].text = '100% Terselesaikan'

# ================ STEP 7: MODIFY TABLE 1 - DAFTAR WILAYAH ASSIGNMENT ================
table1 = doc.tables[1]

# Clear all data rows (keep header)
while len(table1.rows) > 1:
    table1._tbl.remove(table1.rows[1]._tr)

# Add data rows
for idx, k in enumerate(sorted_kec, 1):
    stats = kec_stats[k]
    row = table1.add_row()
    row.cells[0].text = str(idx)
    row.cells[1].text = full_kode_kec(k)
    row.cells[2].text = stats['nama']
    row.cells[3].text = str(stats['jumlah_tugas'])
    row.cells[4].text = 'Terselesaikan'

# ================ STEP 8: MODIFY SUMMARY PARAGRAPH AFTER TABLE 1 ================
# Find the paragraph after table 1's mention of '31 kecamatan' or '1.156.276'
for para in doc.paragraphs:
    if '1.156.276' in para.text or '1.156.276' in para.text:
        new_text = para.text.replace('31 kecamatan', f'{total_kec} kecamatan')
        new_text = new_text.replace('1.156.276', str(total_sls))
        set_paragraph_text(para, new_text)
    elif 'seluruh 31 kecamatan' in para.text:
        new_text = para.text.replace('31 kecamatan', f'{total_kec} kecamatan')
        set_paragraph_text(para, new_text)
    elif '31 kecamatan' in para.text and ('terselesaikan' in para.text or 'seluruh' in para.text or 'seluruh' in para.text):
        new_text = para.text.replace('31 kecamatan', f'{total_kec} kecamatan')
        if '1.156.276' in new_text:
            new_text = new_text.replace('1.156.276', str(total_sls))
        set_paragraph_text(para, new_text)
    elif '7 PML' in para.text:
        new_text = para.text.replace('7 PML', f'{total_pml} PML')
        new_text = new_text.replace('15 PPL', f'{total_ppl} PPL')
        set_paragraph_text(para, new_text)

# Replace numbers in kesimpulan paragraphs
for para in doc.paragraphs:
    t = para.text
    if '31 kecamatan di Kabupaten Jember' in t:
        set_paragraph_text(para, t.replace('31', str(total_kec)))
    elif '1.156.276 unit tugas' in t:
        set_paragraph_text(para, t.replace('1.156.276', str(total_sls)))

# ================ STEP 9: MODIFY TABLE 2 - HASIL DAN CAPAIAN ================
table2 = doc.tables[2]
# Row 1: Jumlah Kecamatan
table2.rows[1].cells[1].text = f'{total_kec} Kecamatan'
# Row 2: Total Unit Tugas
table2.rows[2].cells[1].text = str(total_sls)
# Row 3: Wilayah Terselesaikan
table2.rows[3].cells[1].text = f'{total_kec} Kecamatan (100%)'
# Row 4: Wilayah Belum Terselesaikan
table2.rows[4].cells[1].text = '0 Kecamatan (0%)'
# Row 5: Jumlah PML
table2.rows[5].cells[1].text = f'{total_pml} Orang'
# Row 6: Jumlah PPL
table2.rows[6].cells[1].text = f'{total_ppl} Orang'

# ================ STEP 10: MODIFY TABLE 4 - LAMPIRAN ================
table4 = doc.tables[4]

# Clear all data rows
while len(table4.rows) > 1:
    table4._tbl.remove(table4.rows[1]._tr)

# Add data rows
for idx, k in enumerate(sorted_kec, 1):
    stats = kec_stats[k]
    row = table4.add_row()
    row.cells[0].text = str(idx)
    row.cells[1].text = full_kode_kec(k)
    row.cells[2].text = stats['nama']
    row.cells[3].text = str(stats['jumlah_tugas'])
    row.cells[4].text = '0'
    row.cells[5].text = 'SELESAI'

# ================ STEP 11: UPDATE NARRATIVE TEXT IN ALL PARAGRAPHS ================
# Update specific paragraph texts that reference old numbers
for para in doc.paragraphs:
    t = para.text
    
    # Update the paragraph that mentions '31 kecamatan' in Ruang Lingkup
    if '31 kecamatan' in t and ('lingkup' in t.lower() or 'mencakup' in t.lower() or 'seluruh' in t.lower() or 'setiap' in t.lower()):
        for run in para.runs:
            run.text = run.text.replace('31 kecamatan', f'{total_kec} kecamatan')
            run.text = run.text.replace('31 Kecamatan', f'{total_kec} Kecamatan')
    
    # Replace 'Persiapan SE2026' or 'persiapan SE2026' with 'Sakernas Semester 2 2026'
    for run in para.runs:
        run.text = run.text.replace('Persiapan SE2026', 'Sakernas Semester 2 2026')
        run.text = run.text.replace('persiapan SE2026', 'Sakernas Semester 2 2026')
        run.text = run.text.replace('persiapan Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')
        run.text = run.text.replace('SE2026,', 'Sakernas Semester 2 2026,')
        run.text = run.text.replace('SE2026.', 'Sakernas Semester 2 2026.')
        run.text = run.text.replace('SE2026\n', 'Sakernas Semester 2 2026\n')
    
    # Replace dasar hukum point 3
    if 'Peraturan Badan Pusat Statistik Nomor 5 Tahun 2023 tentang Penyelenggaraan Sensus Ekonomi 2026' in t:
        for run in para.runs:
            run.text = run.text.replace('Peraturan Badan Pusat Statistik Nomor 5 Tahun 2023 tentang Penyelenggaraan Sensus Ekonomi 2026', 
                                        'Peraturan Badan Pusat Statistik Nomor 8 Tahun 2024 tentang Penyelenggaraan Survei Angkatan Kerja Nasional')
    # Replace point 4
    if 'Surat Keputusan Kepala BPS Kabupaten Jember tentang Penetapan Petugas Lapangan Sensus Ekonomi 2026' in t:
        for run in para.runs:
            run.text = run.text.replace('Sensus Ekonomi 2026', 'Sakernas Semester 2 2026')
    
    # Replace point 5
    if 'Pedoman Teknis Pelaksanaan Assignment Wilayah Sensus Ekonomi 2026 BPS RI' in t:
        for run in para.runs:
            run.text = run.text.replace('Pedoman Teknis Pelaksanaan Assignment Wilayah Sensus Ekonomi 2026 BPS RI',
                                        'Pedoman Teknis Pelaksanaan Assignment Wilayah Sakernas Semester 2 2026 BPS RI')
    
    # Replace 'Persiapan Sensus Ekonomi 2026' title
    if 'PERSIAPAN SENSUS EKONOMI 2026' in t:
        for run in para.runs:
            run.text = run.text.replace('PERSIAPAN SENSUS EKONOMI 2026', 
                                        'SAKERNAS SEMESTER 2 TAHUN 2026')

# ================ STEP 12: UPDATE LAMPIRAN TEXT ================
for para in doc.paragraphs:
    if 'sistem aplikasi assignment SE2026' in para.text:
        for run in para.runs:
            run.text = run.text.replace('sistem aplikasi assignment SE2026', 
                                        'sistem aplikasi assignment Sakernas 2026')
    elif 'Sensus Ekonomi 2026 di BPS' in para.text:
        for run in para.runs:
            run.text = run.text.replace('Sensus Ekonomi 2026 di BPS', 
                                        'Sakernas Semester 2 2026 di BPS')
    elif 'dalam rangka persiapan Sensus Ekonomi 2026 di BPS' in para.text:
        for run in para.runs:
            run.text = run.text.replace('dalam rangka persiapan Sensus Ekonomi 2026 di BPS',
                                        'dalam rangka Sakernas Semester 2 2026 di BPS')

# ================ STEP 13: SAVE THE DOCUMENT ================
output_path = 'C:/laragon/www/majapahit/sakernas/3509_Laporan_Assignment_PPL_PML_Sakernas_2026.docx'
doc.save(output_path)
print(f"Document saved to: {output_path}")

# Print summary for verification
print("\n=== LAPORAN SUMMARY ===")
print(f"Kegiatan: Sakernas Semester 2 Tahun 2026")
print(f"Lokasi: BPS Kabupaten Jember (3509)")
print(f"Jumlah Kecamatan: {total_kec}")
print(f"Jumlah Tugas (SLS): {total_sls}")
print(f"Jumlah PML: {total_pml}")
print(f"Jumlah PPL: {total_ppl}")
print(f"\nDaftar Kecamatan:")
for k in sorted_kec:
    s = kec_stats[k]
    print(f"  {s['kode']} - {s['nama']}: {s['jumlah_tugas']} SLS, {s['jumlah_desa']} desa")
print(f"\nDaftar PML ({total_pml}):")
for p in sorted(pml_set):
    k_list = sorted(pml_kec[p])
    print(f"  {p}: {len(k_list)} kec ({', '.join(k_list)})")
print(f"\nDaftar PPL ({total_ppl}):")
for p in sorted(ppl_set):
    print(f"  {p}")
