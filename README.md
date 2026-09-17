# Dashboard Monitoring Pembangunan ZI — Provinsi Gorontalo

Paket ini berisi:
- `app.py` — dashboard web Streamlit.
- `Dashboard_Monitoring_ZI_Gorontalo.xlsx` — snapshot Excel dengan Tab 1 Dashboard dan Tab 2 Progress ZI Satker.
- `sample_data/` — dua file LKE contoh yang digunakan untuk mengisi demo.
- `.streamlit/secrets.toml.example` — contoh konfigurasi password admin.

## Fitur dashboard web

### Viewer
- Dashboard ringkasan:
  1. Jumlah satker diusulkan
  2. Jumlah satker yang memenuhi syarat WBK
  3. Rata-rata Nilai RB
  4. Rata-rata tindak lanjut
  5. Diagram tindak lanjut per satker
  6. Peringkat Nilai RB
- Progress Satker:
  - dropdown satker
  - Nilai Area Perubahan dan Komponen Hasil
  - Status tindak lanjut per komponen
  - persentase tindak lanjut per komponen pengungkit
  - klik/expand komponen untuk melihat detail setiap poin LKE pada level terkecil, misalnya `A.1.i.a`, `A.2.i.a`, dst.

### Admin
- Login khusus admin.
- Upload satu atau beberapa file LKE `.xlsx`.
- Jika nama satker belum ada → ditambahkan.
- Jika nama satker sudah ada → data satker diperbarui.
- Data tersimpan di SQLite lokal (`data/zi_dashboard.db`).

## Menjalankan

1. Install Python 3.10+.
2. Jalankan:
   ```bash
   pip install -r requirements.txt
   ```
3. Salin `.streamlit/secrets.toml.example` menjadi `.streamlit/secrets.toml`.
4. Ganti password contoh dengan password admin sendiri.
5. Jalankan:
   ```bash
   streamlit run app.py
   ```
jalankan: py -m streamlit run app.py

## Catatan keamanan

Login hanya diperlukan untuk menu Admin karena admin dapat mengubah data dashboard. Viewer tidak perlu login.

Untuk penggunaan internal sederhana, password admin sudah cukup untuk prototype. Untuk deployment resmi/internet, lebih baik menggunakan autentikasi akun organisasi (misalnya Google Workspace/SSO) dan membatasi siapa yang memiliki role Admin.

## Catatan data

Parser membaca sheet `Utama` dan `Jawaban` dari file LKE. File sebaiknya sudah dibuka dan disimpan di Excel terlebih dahulu agar nilai formula pada sheet `Utama` tersedia.

Persentase tindak lanjut dihitung dari 3 checklist pada setiap poin:
- Perbaikan Catatan TPI
- Sudah ada eviden 2024–2025
- Kesesuaian Penamaan Eviden

Baris total/agregat tidak dihitung sebagai poin tindak lanjut.
