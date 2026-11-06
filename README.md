# Sistem Interview Peserta Magang — Streamlit

Aplikasi ini mengubah **Template_Interview_Peserta_Magang.xlsx** menjadi sistem pengisian berbasis web tanpa perlu database pada tahap awal.

## Fitur

1. **Mulai di Sini** — Dashboard kandidat, rekomendasi, keputusan pewawancara, ringkasan.
2. **Data Kandidat** — Form 25 kolom sama persis dengan sheet Excel, dengan ID MG-001 dan seterusnya. Bisa tambah/edit kandidat.
3. **Penilaian Interview** — Tujuh aspek dengan bobot sesuai Excel, catatan hasil wawancara, rekomendasi dan keputusan akhir.
4. **Summary & Ranking** — Peringkat otomatis, filter rekomendasi dan posisi, grafik, download CSV.
5. **Panduan & Rubrik** — 12 pertanyaan wawancara, deskripsi bobot serta skala penilaian dari sheet Excel.

## Menjalankan di komputer

Pastikan Python 3.10+ terpasang, kemudian:

```bash
pip install -rrequirements.txt
streamlit run app.py
```

Aplikasi akan terbuka di browser. Untuk deploy ke **Streamlit Community Cloud**, buat GitHub repository dan upload **keempat file**: `app.py`, `data_store.py`, `requirements.txt`, `Template_Interview_Peserta_Magang.xlsx`. Pilih `app.py` sebagai Main file path saat deploy.

## Penting: penyimpanan data

**Versi ini tidak menggunakan database.** Data hanya disimpan dalam memori sesi Streamlit dan **belum permanen** saat Anda menekan tombol "Simpan". Sebelum menutup tab, menyegarkan browser, atau redeploy, klik **Download Backup Excel** di sidebar untuk menyimpan data. Ketika ingin melanjutkan, upload file Excel tersebut melalui sidebar, lalu klik **Muat Data Excel**. Impor akan mengganti data yang sudah ada dalam sesi saat itu.

File Excel yang diunduh mempertahankan keempat sheet, format, rumus, bobot dan ambang nilai asli. Excel biasanya menghitung rumus saat file dibuka.

## Privasi dan keamanan

Data pelamar berisi informasi pribadi, misalnya alamat, kontak, IPK dan catatan seleksi. Gunakan aplikasi di jaringan/perangkat tepercaya. **Jangan deploy sebagai situs publik tanpa kontrol akses.** Jika menggunakan Streamlit Cloud, set `APP_PASSWORD` di **App → Settings → Secrets**:

```toml
APP_PASSWORD = "ganti-dengan-sandi-yang-kuat-dan-unik"
```

Sandi sederhana ini hanya penghalang akses tingkat awal, bukan sistem autentikasi multiuser atau pengamanan enterprise. Jangan simpan sandi di GitHub/repository. Untuk penggunaan bersama atau penyimpanan otomatis jangka panjang, gunakan autentikasi yang memadai dan basis data dengan kontrol akses.

## Ketentuan penilaian

- Komunikasi 15%; Motivasi 15%; Kompetensi Teknis 25%; Pemecahan Masalah 15%; Kerja Sama 10%; Kemauan Belajar 10%; Profesionalisme 10%.
- Skor akhir = jumlah (skor 1–5 × bobot) × 20.
- ≥80 direkomendasikan; ≥65 dipertimbangkan; di bawah 65 belum direkomendasikan.
- Semua tujuh aspek harus diisi sebelum sistem menghitung skor dan memasukkan kandidat ke ranking.
- Keputusan akhir (Lolos/Cadangan/Tidak Lolos/Menunggu) selalu ditentukan pewawancara.
- Template Excel mendukung maksimal 150 kandidat.

## Struktur berkas

```text
app.py                                # UI Streamlit
data_store.py                         # Baca/tulis Excel, bobot, ranking
requirements.txt                      # Dependensi
Template_Interview_Peserta_Magang.xlsx # Format Excel asli
```
