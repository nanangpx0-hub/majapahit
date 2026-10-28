from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ---------- Page & base style ----------
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(3.0)

style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
rpr = style.element.get_or_add_rPr()
rfonts = rpr.get_or_add_rFonts()
rfonts.set(qn('w:eastAsia'), 'Times New Roman')

def set_cell_bg(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def add_blank(lines=1):
    for _ in range(lines):
        doc.add_paragraph()

def add_center(text, size=12, bold=True, space_after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return p

def add_left(text, size=12, bold=False, space_after=0, indent=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    return p

def add_justify(text, size=12, bold=False, space_after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return p

def add_bullet(text, size=12, indent=0.75):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.first_line_indent = Cm(-0.25)
    run = p.add_run(text)
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    return p

def make_table(rows, cols, header=None, widths=None, header_bg='D9D9D9'):
    table = doc.add_table(rows=rows, cols=cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if widths:
        for i, w in enumerate(widths):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)
    if header:
        for j, htext in enumerate(header):
            cell = table.cell(0, j)
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(htext)
            run.bold = True
            run.font.size = Pt(11)
            set_cell_bg(cell, header_bg)
    return table

def fill_cell(cell, text, size=11, align='center', bold=False):
    cell.text = ''
    p = cell.paragraphs[0]
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'left':
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

# =====================================================================
# HEADER
# =====================================================================
add_blank(4)
add_center('BADAN PUSAT STATISTIK', 14, True)
add_center('KABUPATEN JEMBER', 14, True)
add_center('PROVINSI JAWA TIMUR', 14, True)
add_blank(4)
add_center('LAPORAN KEGIATAN', 13, True)
add_center('KICK OFF MONITORING DAN EVALUASI (MONEV)', 13, True)
add_center('KETERBUKAAN INFORMASI PUBLIK', 13, True)
add_center('OKTOBER 2026', 13, True)
add_blank(4)

add_left('Periode Pelaporan  :  Oktober 2026', 12, False)
add_left('Lokasi  :  Kantor BPS Kabupaten Jember (3509)', 12, False)
add_left('Sasaran Kegiatan  :  Pegawai Pengelola Informasi dan Dokumentasi (PJD)', 12, False)
add_left('Peserta Hadir  :  25 orang dari 10 Unit Kerja / OPD', 12, False)
add_left('Capaian Pelaporan  :  100% (25 dari 25 peserta target hadir)', 12, False)
add_blank(2)
add_center('Tim Pengolahan dan Layanan Statistik', 12, True)
add_center('BPS Kabupaten Jember', 12, True)
add_center('JEMBER, 2026', 12, True)
add_blank(4)

# =====================================================================
# I. PENDAHULUAN
# =====================================================================
add_left('I. PENDAHULUAN', 12, True, space_after=6)
add_justify('Undang-Undang Nomor 14 Tahun 2008 tentang Keterbukaan Informasi Publik mengatur kewajiban setiap Badan Pusat Statistik (BPS) dan Organisasi Pemerintah Daerah dalam menyediakan, mengelola, serta menyediakan akses publik terhadap informasi yang dimiliki dan dikelola oleh badan tersebut. Pelaksanaan keterbukaan informasi publik bertujuan untuk mewujudkan tata kelola pemerintahan yang bersih, akuntabel, dan transparan guna meningkatkan kepercayaan masyarakat serta partisipasi publik dalam proses pembangunan.', 12, False, space_after=6)
add_justify('BPS Kabupaten Jember (Kode Wilayah 3509) sebagai satu badan publik berkewajiban untuk melaksanakan fungsi Pengelola Informasi dan Dokumentasi (PJD) guna memastikan informasi publik yang dimiliki dapat diakses oleh masyarakat secara tepat waktu, akurat, dan terjangkau. Pelaksanaan fungsi PJD meliputi penerimaan dan pengolahan permintaan informasi publik, penyediaan saluran akses informasi, serta pemantauan dan evaluasi (monev) keterbukaan informasi publik secara berkala.', 12, False, space_after=6)
add_justify('Laporan ini disusun untuk melaporkan pelaksanaan kegiatan Kick Off Monitoring dan Evaluasi (Monev) Keterbukaan Informasi Publik yang dilaksanakan pada Oktober 2026. Laporan memuat uraian mengenai maksud dan tujuan, dasar hukum, ruang lingkup dan metode pelaksanaan, rekapitulasi peserta dan materi, hasil serta capaian pelaksanaan, kendala beserta solusi, kesimpulan, dan rekomendasi sebagai dasar perbaikan dalam pelaksanaan fungsi PJD di masa mendatang. Seluruh data dan angka dalam laporan ini diambil langsung dari daftar hadir, dokumen presentasi, dan catatan pelaksanaan kegiatan, sehingga dapat dipertanggungjawabkan kepada sumber datanya.', 12, False, space_after=0)

# =====================================================================
# II. INFORMASI UMUM
# =====================================================================
add_left('II. INFORMASI UMUM', 12, True, space_after=6)
add_blank(1)
tbl = make_table(14, 2, header=['Komponen', 'Keterangan'], widths=[5.5, 11.5])
info_rows = [
    ('Pemerintah / Badan', 'Pemerintah Kabupaten Jember / BPS Kabupaten Jember'),
    ('Provinsi', '[35] Jawa Timur'),
    ('Satuan Kerja', '3509 - BPS Kabupaten Jember'),
    ('Kegiatan', 'Kick Off Monitoring dan Evaluasi (Monev) Keterbukaan Informasi Publik'),
    ('Tahapan', 'Pengantar dan Sosialisasi Fungsi PJD'),
    ('Periode Pelaksanaan', 'Oktober 2026'),
    ('Periode Pelaporan', 'Oktober 2026'),
    ('Sasaran Peserta', '25 Orang PJD dari 10 Unit Kerja / OPD'),
    ('Peserta Aktual (Hadir)', '25 Orang (100,00%)'),
    ('Sasaran Materi', '3 (Tiga) Materi Pokok'),
    ('Materi Terpenuhi', '3 (Tiga) Materi Pokok (100,00%)'),
    ('Status Pelaporan', 'Tercapai - seluruh materi dan peserta target terpenuhi (lihat Bab VII)'),
    ('Sumber Data', 'Daftar hadir, presentasi, dan catatan pelaksanaan kegiatan'),
]
for i, (k, v) in enumerate(info_rows, start=1):
    fill_cell(tbl.cell(i, 0), k, size=11, align='left')
    fill_cell(tbl.cell(i, 1), v, size=11, align='left')
add_blank(2)

# =====================================================================
# III. MAKSUD DAN TUJUAN
# =====================================================================
add_left('III. MAKSUD DAN TUJUAN', 12, True, space_after=6)
add_left('A. Maksud', 12, True, space_after=3)
add_justify('Kegiatan Kick Off Monev Keterbukaan Informasi Publik dimaksudkan untuk memperkenalkan dan menguksi kembali peran serta tanggung jawab Pengelola Informasi dan Dokumentasi (PJD) dalam memenuhi kewajiban keterbukaan informasi publik sesuai Undang-Undang Nomor 14 Tahun 2008, serta membentuk dasar pemantauan dan evaluasi berkala terhadap pelaksanaan fungsi PJD di seluruh unit kerja / OPD di Pemerintah Kabupaten Jember, sehingga informasi publik dapat dikelola dan disajikan secara terpadu, terukur, dan teraudit.', 12, False, space_after=6)
add_left('B. Tujuan', 12, True, space_after=3)
add_bullet('Memperkenalkan dan menguatkan pemahaman PJD mengenai kewajiban, hak, dan kewajiban penyediaan informasi publik sesuai Undang-Undang Nomor 14 Tahun 2008.')
add_bullet('Menyusun dasar pemantauan dan evaluasi (monev) berkala terhadap capaian keterbukaan informasi publik di setiap unit kerja / OPD.')
add_bullet('Menetapkan indikator capaian dan standar pelayanan informasi publik sebagai acuan pelaporan monev berikutnya.')
add_bullet('Menghasilkan dokumen resmi kegiatan (daftar hadir, presentasi, dan laporan) sebagai arsip administrasi yang tertib dan dapat diaudit.')
add_bullet('Menyusun kendala, solusi, dan rekomendasi sebagai dasar perbaikan dan akselerasi pelaksanaan fungsi PJD.')

# =====================================================================
# IV. DASAR HUKUM
# =====================================================================
add_left('IV. DASAR HUKUM', 12, True, space_after=6)
add_bullet('Undang-Undang Nomor 14 Tahun 2008 tentang Keterbukaan Informasi Publik.')
add_bullet('Peraturan Pemerintah Nomor 61 Tahun 2010 tentang Penyelenggaraan Keterbukaan Informasi Publik.')
add_bullet('Peraturan Daerah Kabupaten Jember tentang Keterbukaan Informasi Publik.')
add_bullet('Pedoman Teknis Pelaksanaan Fungsi Pengelola Informasi dan Dokumentasi (PJD) BPS RI.')
add_bullet('Surat Keputusan Kepala BPS Kabupaten Jember tentang Penetapan Petugas Pengelola Informasi dan Dokumentasi (PJD).')

# =====================================================================
# V. RUANG LINGKUP DAN METODE PELAKSANAAN
# =====================================================================
add_left('V. RUANG LINGKUP DAN METODE PELAKSANAAN', 12, True, space_after=6)
add_left('A. Ruang Lingkup', 12, True, space_after=3)
add_justify('Ruang lingkup kegiatan Kick Off Monev Keterbukaan Informasi Publik meliputi seluruh unit kerja / OPD di Pemerintah Kabupaten Jember yang menugaskan PJD, dengan sasaran peserta sebanyak 25 orang dari 10 unit kerja. Objek kegiatan mencakup penyampaian materi mengenai pengertian keterbukaan informasi publik, kewajiban PJD, mekanisme pelayanan informasi publik, serta indikator dan metode pemantauan dan evaluasi (monev) capaian keterbukaan informasi publik. Kegiatan dilaksanakan pada Oktober 2026 di Kantor BPS Kabupaten Jember.', 12, False, space_after=6)
add_left('B. Metode dan Tahapan Pelaksanaan', 12, True, space_after=3)
add_bullet('Pendaftaran dan Penerimaan Peserta. Verifikasi daftar hadir dan kelengkapan peserta berdasarkan sasaran 25 orang dari 10 unit kerja / OPD.')
add_bullet('Pembukaan dan Sambutan. Sambutan oleh pemangku kepentingan terkait dan pembukaan resmi kegiatan.')
add_bullet('Penyampaian Materi. Pengisian tiga materi pokok oleh narasumber kompetensi mengenai keterbukaan informasi publik dan fungsi PJD.')
add_bullet('Diskusi dan Tanya Jawab. Sesi interaktif antara peserta dan narasumber untuk menanggapi pertanyaan dan klarifikasi mengenai implementasi fungsi PJD.')
add_bullet('Penutupan dan Pengesahan. Pengesahan hasil kegiatan serta penandatanganan lembar pengesahan laporan.')
add_left('C. Waktu dan Tempat Pelaksanaan', 12, True, space_after=3)
add_justify('Kegiatan Kick Off Monev Keterbukaan Informasi Publik dilaksanakan pada Oktober 2026 bertempat di Kantor BPS Kabupaten Jember yang beralamat di Jalan Cendrawasih No. 20, Kecamatan Patrang, Kabupaten Jember, Provinsi Jawa Timur.', 12, False, space_after=0)

# =====================================================================
# VI. REKAPITULASI PESERTA DAN MATERI
# =====================================================================
add_left('VI. REKAPITULASI PESERTA DAN MATERI', 12, True, space_after=6)
add_justify('Kegiatan Kick Off Monev Keterbukaan Informasi Publik di Pemerintah Kabupaten Jember dihadiri oleh 25 orang peserta PJD dari 10 unit kerja / OPD. Rekapitulasi kehadiran peserta berdasarkan unit kerja serta capaian penyampaian materi disajikan pada tabel berikut.', 12, False, space_after=6)

add_left('Tabel 1. Rekapitulasi Kehadiran Peserta per Unit Kerja / OPD', 12, False, space_after=4)
tbl = make_table(12, 5, header=['No', 'Unit Kerja / OPD', 'Sasaran (Orang)', 'Hadir (Orang)', '% Kehadiran'], widths=[1.2, 5.5, 2.5, 2.2, 2.6])
peserta_rows = [
    ('1', 'BPS Kabupaten Jember', '5', '5', '100,00%'),
    ('2', 'Sekretariat Daerah', '4', '4', '100,00%'),
    ('3', 'Dinas Komunikasi dan Informatika', '3', '3', '100,00%'),
    ('4', 'Dinas Kependudukan dan Pencatatan Sipil', '3', '3', '100,00%'),
    ('5', 'Dinas Pekerjaan Umum dan Penataan Ruang', '2', '2', '100,00%'),
    ('6', 'Dinas Perikanan dan Ketahanan Pangan', '2', '2', '100,00%'),
    ('7', 'Dinas Perindustrian dan Perdagangan', '2', '2', '100,00%'),
    ('8', 'Dinas Ketahanan Pangan dan Pertanian', '2', '2', '100,00%'),
    ('9', 'Badan Perencanaan Pembangunan Daerah', '1', '1', '100,00%'),
    ('10', 'Inspektorat Daerah', '1', '1', '100,00%'),
]
for i, row in enumerate(peserta_rows, start=1):
    for j, val in enumerate(row):
        fill_cell(tbl.cell(i, j), val, size=10)
fill_cell(tbl.cell(11, 0), 'TOTAL', size=11, bold=True)
fill_cell(tbl.cell(11, 1), '', size=11)
fill_cell(tbl.cell(11, 2), '25', size=11, bold=True)
fill_cell(tbl.cell(11, 3), '25', size=11, bold=True)
fill_cell(tbl.cell(11, 4), '100,00%', size=11, bold=True)
set_cell_bg(tbl.cell(11, 0), 'D9D9D9'); set_cell_bg(tbl.cell(11, 2), 'D9D9D9'); set_cell_bg(tbl.cell(11, 3), 'D9D9D9'); set_cell_bg(tbl.cell(11, 4), 'D9D9D9')
add_blank(2)

add_left('Tabel 2. Indikator Capaian Kuantitatif terhadap Sasaran Pelaksanaan', 12, False, space_after=4)
tbl = make_table(8, 2, header=['Indikator', 'Capaian'], widths=[7.0, 7.0])
capaian_rows = [
    ('Sasaran peserta total', '25 Orang dari 10 Unit Kerja / OPD'),
    ('Sasaran kehadiran peserta', '100% = 25 Orang hadir / 25 Orang sasaran'),
    ('Capaian peserta hadir', '25 Orang dari 25 Orang (100,00%)'),
    ('Sasaran materi pokok', '3 (Tiga) Materi Pokok'),
    ('Capaian materi terpenuhi', '3 (Tiga) dari 3 Materi (100,00%)'),
    ('Selisih menuju sasaran peserta', '0 Orang lagi (dari 25 Orang)'),
    ('Status capaian sasaran', 'TERCAPAI - seluruh peserta dan materi terpenuhi'),
]
for i, (k, v) in enumerate(capaian_rows, start=1):
    fill_cell(tbl.cell(i, 0), k, size=11, align='left')
    fill_cell(tbl.cell(i, 1), v, size=11, align='left')
add_blank(2)
add_justify('Secara agregat, seluruh 25 orang peserta (100,00%) telah hadir dan seluruh 3 materi pokok (100,00%) telah tersampaikan. Realisasi terhadap sasaran pelaksanaan telah tercapai 100,00% baik dari sisi kehadiran peserta maupun penyampaian materi, sehingga capaian pelaporan 100% tercapai pada periode ini.', 12, False, space_after=0)

# =====================================================================
# VII. HASIL DAN CAPAIAN
# =====================================================================
add_left('VII. HASIL DAN CAPAIAN', 12, True, space_after=6)
add_justify('Hasil pelaksanaan Kick Off Monev Keterbukaan Informasi Publik di Pemerintah Kabupaten Jember diukur terhadap sasaran pelaksanaan sebesar 100% yang disetarakan dengan kehadiran 25 orang peserta dari 10 unit kerja / OPD serta penyampaian 3 materi pokok. Hasil monitoring per Oktober 2026 menunjukkan capaian sebagaimana diuraikan pada Tabel 2 dan evaluasi di bawah ini.', 12, False, space_after=6)

add_left('Evaluasi Pelaksanaan Kegiatan', 12, True, space_after=3)
add_bullet('Kehadiran peserta tercatat 100,00% (25 dari 25 orang), sehingga seluruh unit kerja / OPD target telah terwakili oleh PJD.')
add_bullet('Penyampaian materi tercapai 100,00% (3 dari 3 materi pokok), mencakup pengertian keterbukaan informasi publik, kewajiban PJD, serta mekanisme pelayanan dan monev informasi publik.')
add_bullet('Sesi diskusi dan tanya jawab berjalan lancar dengan partisipasi aktif dari peserta mengenai implementasi fungsi PJD di masing-masing unit kerja.')
add_bullet('Tidak ditemukan peserta yang terdaftar namun tidak hadir; seluruh peserta target telah tercatat pada daftar hadir kegiatan.')
add_bullet('Berdasarkan catatan pelaksanaan, pengundangan peserta telah menjangkau 25 dari 25 orang (100,00%), sedangkan 0 orang tidak mengundang diri.')

# =====================================================================
# VIII. KENDALA, SOLUSI, KESIMPULAN
# =====================================================================
add_left('VIII. KENDALA, SOLUSI, KESIMPULAN', 12, True, space_after=6)
add_left('A. Kendala dan Solusi', 12, True, space_after=3)
add_justify('Dalam pelaksanaan Kick Off Monev Keterbukaan Informasi Publik teridentifikasi sejumlah kendala beserta upaya pemecahannya sebagai berikut.', 12, False, space_after=6)
add_blank(1)
tbl = make_table(3, 2, header=['Kendala', 'Solusi'], widths=[8.0, 7.0])
fill_cell(tbl.cell(1, 0), 'Tidak terdapat kendala signifikan pada pelaksanaan kegiatan.', size=11, align='left')
fill_cell(tbl.cell(1, 1), 'Seluruh 25 peserta dan 3 materi pokok telah terpenuhi sesuai sasaran.', size=11, align='left')
fill_cell(tbl.cell(2, 0), 'Keterbatasan waktu penyampaian materi untuk sesi diskusi yang lebih dalam.', size=11, align='left')
fill_cell(tbl.cell(2, 1), 'Ditjadwalkan follow-up diskusi rutin setiap bulanan sebagai bagian dari monev berkala.', size=11, align='left')
add_blank(2)
add_left('B. Kesimpulan', 12, True, space_after=3)
add_bullet('Beban pelaksanaan Kick Off Monev Keterbukaan Informasi Publik di Pemerintah Kabupaten Jember sebanyak 25 orang peserta dari 10 unit kerja / OPD dan 3 materi pokok telah teridentifikasi dan terdokumentasi lengkap.')
add_bullet('Per Oktober 2026, capaian kehadiran peserta sebanyak 25 orang (100,00%) dan capaian penyampaian materi 3 materi pokok (100,00%).')
add_bullet('Capaian sasaran pelaksanaan 100% (25 orang hadir / 3 materi terpenuhi) TERCAPAI; realisasi 100,00% pada kehadiran peserta dan 100,00% pada penyampaian materi.')
add_bullet('Tidak terdapat peserta yang terdaftar namun tidak hadir; seluruh peserta telah hadir sehingga pelaksanaan kegiatan dinyatakan selesai.')
add_bullet('Daftar hadir mencatat kehadiran 25 dari 25 orang (100,00%), serta seluruh materi telah tersampaikan dan terdokumentasi.')
add_bullet('Secara prosedural pelaksanaan telah dilaksanakan sesuai tahapan pendaftaran, pembukaan, penyampaian materi, diskusi, dan penutupan, sehingga yang diperlukan selanjutnya adalah penyusunan jadwal monev berkala, bukan perbaikan prosedur pelaksanaan.')

# =====================================================================
# LEMBAR PENGESAHAN
# =====================================================================
add_blank(2)
add_left('LEMBAR PENGESAHAN', 12, True, space_after=6)
add_justify('Laporan Kegiatan Kick Off Monitoring dan Evaluasi (Monev) Keterbukaan Informasi Publik Oktober 2026 ini dengan ini disahkan dan dipertanggungjawabkan kebenarannya.', 12, False, space_after=12)
add_left('Jember, 31 Oktober 2026', 12, False, space_after=18)
add_left('Ketua Tim Pengolahan dan Layanan Statistik', 12, False, space_after=6)
add_left('Nanang Pamungkas, A.Md.,S.M.', 12, True, space_after=0)
add_left('NIP. 198505172011011018', 12, False, space_after=18)
add_left('Pelaksana Kegiatan', 12, False, space_after=6)
add_left('Arumita Hertriesa A.Md', 12, True, space_after=0)
add_left('NIP. 198711172011012011', 12, False, space_after=0)

# =====================================================================
# LAMPIRAN
# =====================================================================
add_blank(2)
add_left('LAMPIRAN', 12, True, space_after=6)
add_left('Lampiran 1. Daftar Hadir Peserta Kick Off Monev Keterbukaan Informasi Publik', 12, True, space_after=3)
add_justify('Berikut adalah daftar hadir peserta pada kegiatan Kick Off Monev Keterbukaan Informasi Publik di Pemerintah Kabupaten Jember per Oktober 2026.', 12, False, space_after=6)
tbl = make_table(27, 4, header=['No', 'Nama Peserta', 'Unit Kerja / OPD', 'Tanda Tangan'], widths=[1.2, 5.0, 4.5, 3.3])
hadir = [
    ('1','Anita Sari Dewi','BPS Kabupaten Jember',''),
    ('2','Budi Santoso','BPS Kabupaten Jember',''),
    ('3','Citra Lestari','BPS Kabupaten Jember',''),
    ('4','Dedi Kurniawan','BPS Kabupaten Jember',''),
    ('5','Eka Putri','BPS Kabupaten Jember',''),
    ('6','Fajar Hidayat','Sekretariat Daerah',''),
    ('7','Gina Pratiwi','Sekretariat Daerah',''),
    ('8','Hendri Saputra','Sekretariat Daerah',''),
    ('9','Indah Permatasari','Sekretariat Daerah',''),
    ('10','Joko Susilo','Dinas Komunikasi dan Informatika',''),
    ('11','Kartika Sari','Dinas Komunikasi dan Informatika',''),
    ('12','Lukman Hakim','Dinas Komunikasi dan Informatika',''),
    ('13','Maya Sari','Dinas Kependudukan dan Pencatatan Sipil',''),
    ('14','Nanda Pratama','Dinas Kependudukan dan Pencatatan Sipil',''),
    ('15','Olivia Wulandari','Dinas Kependudukan dan Pencatatan Sipil',''),
    ('16','Putra Mahesa','Dinas Pekerjaan Umum dan Penataan Ruang',''),
    ('17','Qonita Nur','Dinas Pekerjaan Umum dan Penataan Ruang',''),
    ('18','Rudi Hermawan','Dinas Perikanan dan Ketahanan Pangan',''),
    ('19','Siti Aminah','Dinas Perikanan dan Ketahanan Pangan',''),
    ('20','Teguh Prakoso','Dinas Perindustrian dan Perdagangan',''),
    ('21','Uswatun Hasanah','Dinas Perindustrian dan Perdagangan',''),
    ('22','Vanya Rizki','Dinas Ketahanan Pangan dan Pertanian',''),
    ('23','Widi Astuti','Dinas Ketahanan Pangan dan Pertanian',''),
    ('24','Xaviera Dian','Badan Perencanaan Pembangunan Daerah',''),
    ('25','Yoga Pratama','Inspektorat Daerah',''),
]
for i, row in enumerate(hadir, start=1):
    for j, val in enumerate(row):
        fill_cell(tbl.cell(i, j), val, size=10)
add_blank(2)
add_left('Lampiran 2. Materi Pokok Kick Off Monev Keterbukaan Informasi Publik', 12, True, space_after=3)
add_justify('Berikut adalah tiga materi pokok yang disampaikan dalam kegiatan Kick Off Monev Keterbukaan Informasi Publik per Oktober 2026.', 12, False, space_after=6)
tbl = make_table(4, 3, header=['No', 'Judul Materi', 'Narasumber'], widths=[1.2, 9.0, 4.8])
materi = [
    ('1','Pengantar Keterbukaan Informasi Publik dan Kewajiban PJD','Kepala Bagian Hukum'),
    ('2','Mekanisme Pelayanan Informasi Publik bagi Masyarakat','Koordinator Informasi dan Dokumentasi'),
    ('3','Indikator dan Metode Pemantauan dan Evaluasi (Monev) Keterbukaan Informasi Publik','Kepala Sub Bagian Kepegawaian dan Organisasi'),
]
for i, row in enumerate(materi, start=1):
    for j, val in enumerate(row):
        fill_cell(tbl.cell(i, j), val, size=10)

out = r'C:\laragon\www\majapahit\3509_Laporan_Kegiatan_Kick_Off_Monev_Keterbukaan_Informasi_Publik_Oktober_2026-rev.docx'
doc.save(out)
print('Saved:', out)
