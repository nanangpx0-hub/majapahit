---
inclusion: always
---

# Auto-Approve & Full Autonomous Execution — Workspace Policy

## Status: AKTIF — Auto Allow All

Semua AI agent yang bekerja di workspace ini (`c:\laragon\www\majapahit`) beroperasi dalam mode **Full Autonomous & Auto-Approve**. Tidak ada konfirmasi manual yang diperlukan.

### Tindakan yang Disetujui Otomatis (100%)

| Kategori | Cakupan |
|:---|:---|
| **Terminal / Shell** | Semua perintah PowerShell, Bash, Python, CLI |
| **File System** | Baca, tulis, edit, overwrite, hapus file sementara |
| **Data Processing** | Excel, PDF, Word, JSON, CSV, database lokal |
| **Background Process** | Subagent, background tasks, proses paralel |
| **Web** | Pencarian web, fetch URL, dokumentasi |

### Aturan Eksekusi

- Jalankan rencana kerja dari **hulu ke hilir** tanpa jeda konfirmasi.
- Jangan meminta izin untuk melanjutkan langkah berikutnya.
- Verifikasi hasil secara mandiri sebelum memberikan laporan akhir.
- Tidak ada dialog interaktif, modal, atau pertanyaan teknis yang membutuhkan persetujuan pengguna.

### Interpreter Python

```
C:\laragon\bin\python\python-3.10\python.exe
```
