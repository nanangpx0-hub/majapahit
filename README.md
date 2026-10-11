# Dokumentasi Proyek Majapahit
**Badan Pusat Statistik Kabupaten Jember (Kode Wilayah: 3509)**  
**Tim Pengolahan dan Layanan Statistik**

---

## Ringkasan Proyek

Repository **majapahit** adalah workspace otomasi pembuatan laporan statistik resmi BPS Kabupaten Jember. Setiap sub-direktori merepresentasikan satu kegiatan survei dengan pipeline Python yang menghasilkan dokumen Word (`.docx`) dan PDF terformat secara programatik — tanpa ketik manual.

| Direktori | Kegiatan | Status |
|:---|:---|:---|
| `seruti/` | SERUTI TW III 2026 — Survei Ekonomi Rumah Tangga Triwulanan | ✅ 100% Clean |
| `susenas/` | SUSENAS September 2026 — Survei Sosial Ekonomi Nasional | ✅ 50% milestone |
| `sakernas/` | SAKERNAS 2026 — Survei Angkatan Kerja Nasional | ✅ Selesai |

---

## Struktur Direktori

```
majapahit/
├── README.md                        ← Dokumentasi utama proyek ini
├── AGENTS.md                        ← Kebijakan kerja untuk AI agent
├── CLAUDE.md                        ← Kebijakan kerja untuk Claude
├── GEMINI.md                        ← Kebijakan kerja untuk Gemini
├── opencode.json                    ← Konfigurasi permissions OpenCode AI
├── .github/hooks/                   ← Hook persetujuan tool VS Code Local (Windows)
│
├── seruti/                          ← Laporan SERUTI TW III 2026
│   ├── create_report_seruti_tw3.py  ← Generator DOCX utama
│   ├── 3509_Laporan_Penyelesaian_Seruti_TW3_2026.docx
│   ├── 3509_Laporan_Penyelesaian_Seruti_TW3_2026.pdf
│   ├── DOKUMENTASI_PENGOLAHAN_SERUTI_TW3_2026.md
│   ├── logo_bps.png
│   ├── tanda_tangan_nanang_pamungkas.png
│   ├── tabulasi/                    ← File Excel tabulasi dari server BPS
│   └── Triwulan 3 2026/             ← Screenshot dashboard progres
│
├── susenas/                         ← Laporan SUSENAS September 2026
│   ├── create_report_susenas.py     ← Generator DOCX + rekap Excel
│   ├── create_laporan_uji_petik_september_2026.py
│   ├── 3509_Laporan_Pengolahan_Susenas_September_2026_50Persen.docx
│   ├── 3509_Laporan_Pengolahan_Susenas_September_2026_50Persen.pdf
│   ├── 3509_Laporan_Uji_Petik_Pengolahan_Susenas_September_2026.docx
│   ├── convert_to_pdf.ps1           ← Script konversi Word → PDF via COM
│   ├── data-sm2-2026/               ← Data monitoring progres pengolahan
│   └── template/                    ← Template DOCX referensi
│
└── sakernas/                        ← Laporan SAKERNAS 2026
    ├── create_report.py             ← Generator DOCX dari data Excel
    ├── create_report_v2.py          ← Versi revisi generator
    ├── verify_report.py             ← Skrip verifikasi isi laporan
    ├── 3509_Laporan_Assignment_PPL_PML_Sakernas_2026.docx
    ├── 3509_Laporan_Assignment_PPL_PML_Sakernas_2026.pdf
    └── data/
        └── PPL-*.xlsx               ← Data assignment PPL/PML
```

---

## Modul 1 — SERUTI (Survei Ekonomi Rumah Tangga Triwulanan)

### Deskripsi
Laporan kegiatan penyelesaian pengolahan sampel dan hasil tabulasi SERUTI Triwulan III Tahun 2026 untuk BPS Kabupaten Jember.

### Metadata Kegiatan
| Field | Value |
|:---|:---|
| Periode | Juli s.d. September 2026 |
| Cut-off data | 11 Oktober 2026 pukul 22:00:02 WIB |
| Blok Sensus (NKS) | 28 |
| Total Sampel RUTA | 280 Rumah Tangga |
| Total Dokumen | 560 (280 Inti + 280 MAK) |
| Status | **100,00% Clean — nihil error** |

### Capaian Pengolahan
| Jenis Dokumen | Target | Clean | Error | Status |
|:---|:---:|:---:|:---:|:---:|
| Kuesioner Inti | 280 | 280 | 0 | 100% ✅ |
| Kuesioner MAK | 280 | 280 | 0 | 100% ✅ |
| **Total** | **560** | **560** | **0** | **100% ✅** |

### Cara Regenerasi Laporan
```powershell
# 1. Generate dokumen DOCX
python C:\laragon\www\majapahit\seruti\create_report_seruti_tw3.py

# 2. Konversi ke PDF (memerlukan Microsoft Word terinstall)
powershell -NoProfile -ExecutionPolicy Bypass `
    -File C:\laragon\www\majapahit\susenas\convert_to_pdf.ps1 `
    -DocxPath "C:\laragon\www\majapahit\seruti\3509_Laporan_Penyelesaian_Seruti_TW3_2026.docx" `
    -PdfPath "C:\laragon\www\majapahit\seruti\3509_Laporan_Penyelesaian_Seruti_TW3_2026.pdf"
```

### Highlight Temuan Tabulasi
- Pengeluaran konsumsi per kapita TW III: **Rp 4.190.104,-**
- Dominasi makanan & minuman non-alkohol: **35,94%** (Rp 1.506.052,-)
- Lonjakan pendidikan TW III: **+148% dari TW II** (tahun ajaran baru)
- Iuran arisan sebagai instrumen keuangan komunal terbesar: **Rp 2,333 triliun**

---

## Modul 2 — SUSENAS (Survei Sosial Ekonomi Nasional)

### Deskripsi
Laporan pengolahan sampel SUSENAS September 2026 pada milestone 50% untuk BPS Kabupaten Jember, beserta laporan uji petik pengolahan.

### Metadata Kegiatan
| Field | Value |
|:---|:---|
| Periode | September 2026 |
| Milestone Pelaporan | 50% (140 RT / 14 Blok Sensus) |
| Total Populasi Sampel | 280 RT / 28 Blok Sensus |
| File Monitoring | `25-09-2026_154754-Progress_Pengolahan_DokM_3509.xls` |

### File Output
- `3509_Laporan_Pengolahan_Susenas_September_2026_50Persen.docx/.pdf` — Laporan pengolahan 50%
- `3509_Laporan_Uji_Petik_Pengolahan_Susenas_September_2026.docx/.pdf` — Laporan uji petik

### Cara Regenerasi Laporan
```powershell
# Laporan pengolahan 50%
python C:\laragon\www\majapahit\susenas\create_report_susenas.py

# Laporan uji petik
python C:\laragon\www\majapahit\susenas\create_laporan_uji_petik_september_2026.py
```

---

## Modul 3 — SAKERNAS (Survei Angkatan Kerja Nasional)

### Deskripsi
Laporan assignment PPL (Pencacah Lapangan) dan PML (Pengawas Lapangan) untuk SAKERNAS 2026, dihasilkan dari data Excel assignment yang diunduh dari sistem BPS.

### Sumber Data
- `data/PPL-*.xlsx` — File Excel berisi data assignment PPL dan PML per kecamatan

### Cara Regenerasi Laporan
```powershell
# Generator versi 1
python C:\laragon\www\majapahit\sakernas\create_report.py

# Generator versi 2 (revisi)
python C:\laragon\www\majapahit\sakernas\create_report_v2.py

# Verifikasi hasil
python C:\laragon\www\majapahit\sakernas\verify_report.py
```

---

## Stack Teknologi

| Komponen | Detail |
|:---|:---|
| Language | Python 3.10 |
| Interpreter | `C:\laragon\bin\python\python-3.10\python.exe` |
| Library DOCX | `python-docx` |
| Library Excel | `openpyxl` |
| Konversi PDF | PowerShell + Microsoft Word COM (`convert_to_pdf.ps1`) |
| IDE | Kiro (VS Code-based) |
| Shell | PowerShell (`pwsh`) |

### Dependensi Python
```bash
pip install python-docx openpyxl
```

---

## Pipeline Otomasi Dokumen

Setiap modul mengikuti pola yang sama:

```
Data Excel/PNG  →  Script Python  →  DOCX terformat  →  PDF (via Word COM)
(sumber data)      (generator)       (output utama)      (output cetak)
```

Skrip Python menggunakan `python-docx` untuk membangun dokumen dari nol secara programatik: kop surat, tabel, grafik, narasi analisis, dan tanda tangan digital dirender langsung dari kode tanpa template manual.

---

## Konversi PDF

Script `susenas/convert_to_pdf.ps1` digunakan bersama oleh semua modul. Script ini menjalankan Microsoft Word via COM automation untuk mengkonversi `.docx` ke `.pdf` dengan fidelitas tinggi.

```powershell
# Penggunaan generik
.\susenas\convert_to_pdf.ps1 `
    -DocxPath "path\ke\file.docx" `
    -PdfPath  "path\ke\output.pdf"
```

---

## Kebijakan Auto-Approve AI Agent

Workspace ini mengaktifkan auto-approve untuk pemanggilan tool pada harness **Local VS Code** melalui hook workspace, dan OpenCode melalui `opencode.json`. Lihat:

- `.github/hooks/auto-approve-all.json` dan `.github/hooks/auto-approve-all.ps1` — VS Code Local di Windows, mengizinkan setiap pemanggilan tool sebelum dijalankan
- `AGENTS.md`, `CLAUDE.md`, dan `GEMINI.md` — instruksi perilaku agent; file instruksi saja tidak mengubah kontrol persetujuan aplikasi
- `opencode.json` — OpenCode AI
- `.kiro/steering/auto-approve.md` dan `.kiro/hooks/auto-approve-all.json` — instruksi/hook khusus Kiro

**Catatan penggunaan dan keamanan:** hook Local berjalan saat `chat.useHooks` aktif dan workspace dipercaya. Hook ini tidak otomatis berlaku untuk Copilot Agent Host, Claude, Codex, Kiro, atau harness lain; masing-masing memakai kontrol dan format hook sendiri. Auto-approve mengizinkan tindakan tool tanpa konfirmasi, termasuk perubahan file dan perintah terminal yang berisiko. Kontrol persetujuan tingkat aplikasi/organisasi tetap berlaku.

---

## Penanggung Jawab

**Nanang Pamungkas, A.Md., S.M.**  
NIP. 198505172011011018  
Ketua Tim Pengolahan dan Layanan Statistik  
BPS Kabupaten Jember  
Jl. Sentot Prawirodirdjo No. 24, Jember 68131  
bps3509@bps.go.id | (0331) 485586

---

*Dokumentasi ini dibuat otomatis — Jember, 11 Oktober 2026*
