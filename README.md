# Praktikum Mahadata — Penjualan Toko

Contoh topik: analisis pola penjualan alat tulis, makanan, dan minuman selama 2025.
Data simulasi dibuat dengan Python (seed 42); bukan data toko sungguhan.
Pastikan topik ini berbeda dari teman sekelas sebelum dikumpulkan.

## Menjalankan di VS Code

1. Ekstrak ZIP terlebih dahulu. Buka folder `Project_Mahadata` lewat **File > Open Folder**.
2. Pastikan Python 3.10 atau lebih baru dan ekstensi Python dari Microsoft sudah tersedia.
3. Buka **Terminal > New Terminal** pada folder tersebut.
4. Instal library sekali saja (memerlukan internet):

   ```powershell
   python -m pip install -r requirements.txt
   ```

5. Jalankan:

   ```powershell
   python analisis_data.py
   ```

   Bisa juga buka `analisis_data.py`, pilih interpreter Python yang sama lewat
   **Python: Select Interpreter**, kemudian klik **Run Python File**.

Program otomatis membuat `database.db` berisi 1.500 transaksi bila belum ada,
membaca database dengan Pandas, membersihkan data, menghitung omzet, dan mengekspor hasil.
Database yang sudah ada tidak dibuat ulang; menjalankan ulang tidak menggandakan transaksi.
File hasil ekspor dan grafik diperbarui pada setiap run. Tutup Excel sebelum menjalankan ulang.

## Hasil otomatis

- `database.db`: sumber SQLite dengan tabel `transaksi`.
- `hasil_export.xlsx`: ringkasan, agregasi bulanan/kategori/produk, dan 1.500 baris data bersih.
- `hasil_export.csv`: data bersih; CSV adalah hasil ekspor, bukan sumber utama.
- `grafik/dashboard_penjualan.png`: grafik garis, batang, dan lingkaran dalam satu dashboard.
- `kesimpulan.txt`: hasil analisis singkat.

Semua file berada di folder script meskipun terminal dijalankan dari folder lain.
Grafik tampil sebagai jendela Matplotlib; tutup jendela setelah selesai.
Untuk melihat grafik di editor VS Code, klik PNG pada Explorer.
Excel dihasilkan dari VS Code; membacanya memerlukan aplikasi Excel/LibreOffice atau ekstensi viewer.
Apabila lingkungan tidak mendukung jendela grafik, gunakan `python analisis_data.py --tanpa-tampil`,
lalu buka PNG. File Excel dan CSV tetap dibuat.

## Penjelasan untuk presentasi

1. Tunjukkan struktur tabel dan jumlah data di SQLite. Satu baris adalah satu transaksi produk.
2. Jelaskan `siapkan_database`, kemudian `pd.read_sql_query` sebagai pembaca sumber utama.
3. Pembersihan: hapus duplikat ID, rapikan spasi, ubah tipe tanggal/angka,
   isi 30 metode pembayaran kosong dengan `Tidak diketahui`, dan saring nilai tidak valid.
4. Rumus omzet = jumlah unit × harga satuan. `groupby` merangkum per bulan, kategori, dan produk.
5. Run script, tunjukkan grafik serta file Excel/CSV, lalu jelaskan kesimpulan yang benar-benar muncul.
6. Jelaskan bahwa omzet bukan laba dan hasil simulasi tidak membuktikan tren pasar nyata.

Target video 3–7 menit; 16:9, minimal 720p, layar kerja dan kamera setengah badan.
Unggah Unlisted ke YouTube dan kumpulkan tautan sesuai instruksi guru.
