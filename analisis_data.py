"""Praktikum Mahadata: analisis penjualan toko dari SQLite.

Jalankan: python analisis_data.py
Pengujian tanpa jendela grafik: python analisis_data.py --tanpa-tampil
Semua data adalah SIMULASI, bukan transaksi toko sungguhan.
"""

import argparse
import random
import sqlite3
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import matplotlib


FOLDER = Path(__file__).resolve().parent
DATABASE = FOLDER / "database.db"


def siapkan_database():
    """Buat 1.500 transaksi sekali saja; database yang ada tetap digunakan."""
    if DATABASE.exists():
        return
    rng = random.Random(42)
    produk = [
        ("Buku Tulis", "Alat Tulis", 6000),
        ("Pulpen", "Alat Tulis", 3500),
        ("Pensil", "Alat Tulis", 2500),
        ("Air Mineral", "Minuman", 4000),
        ("Teh Kemasan", "Minuman", 5000),
        ("Susu Kotak", "Minuman", 7000),
        ("Roti", "Makanan", 6000),
        ("Biskuit", "Makanan", 8000),
        ("Keripik", "Makanan", 10000),
    ]
    transaksi = []
    for nomor in range(1, 1501):
        nama, kategori, harga = rng.choice(produk)
        tanggal = date(2025, 1, 1) + timedelta(days=rng.randrange(365))
        # Sisipkan spasi dan metode kosong untuk demonstrasi pembersihan.
        metode = rng.choice(["Tunai", "QRIS", "Debit"])
        transaksi.append((nomor, tanggal.isoformat(), f" {nama} ", kategori,
                          rng.randint(1, 8), harga,
                          None if nomor % 50 == 0 else metode))
    with sqlite3.connect(DATABASE) as koneksi:
        koneksi.execute("""CREATE TABLE transaksi (
            id_transaksi INTEGER PRIMARY KEY,
            tanggal TEXT NOT NULL,
            produk TEXT NOT NULL,
            kategori TEXT NOT NULL,
            jumlah INTEGER NOT NULL CHECK(jumlah > 0),
            harga_satuan INTEGER NOT NULL CHECK(harga_satuan > 0),
            metode_pembayaran TEXT
        )""")
        koneksi.executemany("INSERT INTO transaksi VALUES (?, ?, ?, ?, ?, ?, ?)", transaksi)


def baca_dan_bersihkan():
    with sqlite3.connect(DATABASE) as koneksi:
        data = pd.read_sql_query("SELECT * FROM transaksi ORDER BY id_transaksi", koneksi)
    awal = len(data)
    kosong = int(data["metode_pembayaran"].isna().sum())
    data = data.drop_duplicates(subset="id_transaksi").copy()
    for kolom in ["produk", "kategori"]:
        data[kolom] = data[kolom].astype("string").str.strip()
    data["metode_pembayaran"] = data["metode_pembayaran"].fillna("Tidak diketahui").str.strip()
    data["tanggal"] = pd.to_datetime(data["tanggal"], errors="coerce")
    for kolom in ["jumlah", "harga_satuan"]:
        data[kolom] = pd.to_numeric(data[kolom], errors="coerce")
    valid = (data["tanggal"].notna() & data["jumlah"].gt(0)
             & data["harga_satuan"].gt(0) & data["produk"].notna()
             & data["kategori"].notna() & data["produk"].ne("")
             & data["kategori"].ne(""))
    data = data.loc[valid].copy()
    if len(data) < 1000:
        raise ValueError("Data valid kurang dari 1.000 baris. Periksa isi database.")
    data["omzet"] = data["jumlah"] * data["harga_satuan"]
    data["bulan"] = data["tanggal"].dt.strftime("%Y-%m")
    catatan = f"Baris awal: {awal}; baris valid: {len(data)}; metode kosong diisi: {kosong}."
    return data, catatan


def analisis(data):
    bulanan = data.groupby("bulan", as_index=False).agg(
        omzet=("omzet", "sum"), jumlah_unit=("jumlah", "sum"),
        transaksi=("id_transaksi", "count"))
    kategori = data.groupby("kategori", as_index=False).agg(omzet=("omzet", "sum"))
    produk = data.groupby("produk", as_index=False).agg(
        jumlah_unit=("jumlah", "sum"), omzet=("omzet", "sum"))
    produk = produk.sort_values("jumlah_unit", ascending=False).reset_index(drop=True)
    return bulanan, kategori, produk


def ekspor_excel(data, bulanan, kategori, produk, catatan):
    # ExcelWriter dipakai agar ekspor Excel otomatis dari satu script Python.
    from openpyxl.styles import Alignment, Font, PatternFill

    ringkasan = pd.DataFrame({
        "Indikator": ["Sumber", "Periode", "Jumlah transaksi", "Total unit",
                      "Omzet (Rp)", "Rata-rata omzet per transaksi (Rp)", "Pembersihan"],
        "Nilai": ["Data simulasi; SQLite database.db; seed 42", "Januari–Desember 2025",
                  len(data), int(data["jumlah"].sum()), int(data["omzet"].sum()),
                  float(data["omzet"].mean()), catatan],
    })
    with pd.ExcelWriter(FOLDER / "hasil_export.xlsx", engine="openpyxl",
                        datetime_format="DD/MM/YYYY") as penulis:
        for nama, tabel in [("Ringkasan", ringkasan), ("Bulanan", bulanan),
                            ("Kategori", kategori), ("Produk", produk), ("Data Bersih", data)]:
            tabel.to_excel(penulis, sheet_name=nama, index=False)
            sheet = penulis.sheets[nama]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            sheet.sheet_view.showGridLines = False
            for cell in sheet[1]:
                cell.fill = PatternFill("solid", fgColor="123B5D")
                cell.font = Font(color="FFFFFF", bold=True)
            sheet.row_dimensions[1].height = 26
            for kolom in sheet.columns:
                huruf = kolom[0].column_letter
                panjang = max(len(str(c.value or "")) for c in kolom)
                sheet.column_dimensions[huruf].width = min(max(panjang + 3, 16), 65)
                for cell in kolom[1:]:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00' if isinstance(cell.value, float) else '#,##0'
            if nama == "Ringkasan":
                sheet.column_dimensions["A"].width = 42
                sheet.column_dimensions["B"].width = 65
                sheet.row_dimensions[8].height = 44
    data.to_csv(FOLDER / "hasil_export.csv", index=False, encoding="utf-8-sig")


def buat_grafik(bulanan, kategori, produk, tampil):
    import matplotlib.pyplot as plt

    folder_grafik = FOLDER / "grafik"
    folder_grafik.mkdir(exist_ok=True)
    plt.rcParams.update({"font.size": 11, "axes.spines.top": False,
                         "axes.spines.right": False})
    fig, axes = plt.subplots(1, 3, figsize=(17, 6), layout="constrained")
    fig.suptitle("Analisis Penjualan Toko • Data Simulasi 2025", fontsize=19, weight="bold")
    axes[0].plot(bulanan["bulan"], bulanan["omzet"] / 1_000_000,
                 marker="o", color="#087F8C", linewidth=2)
    axes[0].set(title="Tren omzet bulanan", xlabel="Bulan", ylabel="Omzet (juta rupiah)")
    axes[0].tick_params(axis="x", rotation=60)
    axes[0].grid(axis="y", alpha=0.2)
    urut = produk.sort_values("jumlah_unit")
    axes[1].barh(urut["produk"], urut["jumlah_unit"], color="#168AAD")
    axes[1].set(title="Produk terlaris berdasarkan unit", xlabel="Jumlah unit terjual")
    axes[2].pie(kategori["omzet"], labels=kategori["kategori"], autopct="%1.1f%%",
                colors=["#123B5D", "#19A7A0", "#F4B942"], startangle=90)
    axes[2].set_title("Kontribusi omzet per kategori")
    fig.savefig(folder_grafik / "dashboard_penjualan.png", dpi=160, facecolor="white")
    if tampil:
        print("Grafik ditampilkan. Tutup jendela grafik untuk mengakhiri program.")
        plt.show()
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tanpa-tampil", action="store_true", help="Simpan grafik tanpa membuka jendela")
    args = parser.parse_args()
    if args.tanpa_tampil:
        matplotlib.use("Agg")
    siapkan_database()
    data, catatan = baca_dan_bersihkan()
    bulanan, kategori, produk = analisis(data)
    ekspor_excel(data, bulanan, kategori, produk, catatan)
    puncak = bulanan.loc[bulanan["omzet"].idxmax()]
    kesimpulan = (
        "ANALISIS PENJUALAN TOKO - DATA SIMULASI\n"
        f"{catatan}\n"
        f"Total omzet: Rp {data['omzet'].sum():,.0f}\n"
        f"Produk terlaris berdasarkan unit: {produk.iloc[0]['produk']} "
        f"({produk.iloc[0]['jumlah_unit']:,} unit).\n"
        f"Bulan omzet tertinggi: {puncak['bulan']} (Rp {puncak['omzet']:,.0f}).\n"
        "Omzet adalah jumlah x harga satuan, bukan laba.\n"
        "Hasil hanya menggambarkan data simulasi, bukan kondisi pasar nyata.\n"
    )
    (FOLDER / "kesimpulan.txt").write_text(kesimpulan, encoding="utf-8")
    print(kesimpulan)
    print(f"Excel, CSV, dan kesimpulan tersimpan di: {FOLDER}")
    buat_grafik(bulanan, kategori, produk, tampil=not args.tanpa_tampil)
    print(f"Selesai. Grafik tersimpan di: {FOLDER / 'grafik'}")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as error:
        raise SystemExit("Tidak bisa menulis output. Tutup file Excel/CSV yang sedang terbuka, lalu run ulang.") from error
