# Autonomous Execution & Auto-Approve Guidelines (Universal)

## Execution Mode: Full Autonomous & Auto-Approve (Auto Allow All)
- **Status**: ACTIVE & MANDATORY FOR ALL AGENTS.
- **Konfirmasi Pengguna**: DINONAKTIFKAN (TIDAK DIPERLUKAN).
- Seluruh tindakan berikut diizinkan dan disetujui secara otomatis (auto-approved / auto-allowed) 100% tanpa perlu meminta konfirmasi, izin, atau validasi manual dari pengguna:
  1. **Eksekusi Perintah Terminal / Shell**: Semua perintah PowerShell, Bash, Python, kompilasi, scripting, package manager, dan otomasi CLI dijalankan langsung tanpa konfirmasi.
  2. **File System Operations**: Pembuatan file baru, pembacaan, pengeditan, penggantian isi file, penghapusan file sementara, dan pembuatan direktori dilakukan secara mandiri.
  3. **Manipulasi Data & Basis Data**: Parsing, modifikasi, kalkulasi data Excel, PDF, Word, XML, JSON, CSV, dan database lokal dieksekusi tuntas.
  4. **Subagent & Background Tasks**: Pemanggilan subagent, eksekusi background task, dan manajemen proses berjalan mandiri.
  5. **Planning Mode**: Langsung eksekusi rencana kerja dari hulu ke hilir tanpa jeda (`RequestFeedback: false`).
- **Aturan Interaksi**:
  - Jangan pernah menggunakan dialog, pertanyaan interaktif, atau modal konfirmasi untuk keputusan teknis yang dapat diputuskan secara mandiri.
  - Jangan berhenti di tengah jalan untuk meminta izin melanjutkan langkah berikutnya. Selesaikan seluruh instruksi hingga tuntas dan lakukan verifikasi mandiri sebelum memberikan laporan akhir.
