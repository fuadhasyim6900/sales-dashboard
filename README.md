# Sales Intelligence Dashboard

Dashboard Streamlit untuk distributor bahan bangunan/cat: omset, AO, customer aktif,
NOO, telemarketing, area, opportunity, dan KPI efisiensi.

## Struktur folder repo

```
sales-dashboard/
├── app.py                # aplikasi utama
├── requirements.txt      # dependency Python
├── .streamlit/
│   └── config.toml       # tema tampilan (opsional)
└── data/
    ├── Data_Omset_Power_Query_tahun_2026.xlsx   # data transaksi (WAJIB)
    └── DATA_TARGET_FUAD.xlsx                     # data target (opsional)
```

Aplikasi otomatis membaca kedua file di atas dari folder `data/` — **tidak perlu upload manual** setiap kali dibuka.
Nama file harus sama persis (case sensitive di Linux/Streamlit Cloud). Kalau nama file datamu beda,
tinggal ganti nilai `DEFAULT_TRANSAKSI_FILE` / `DEFAULT_TARGET_FILE` di bagian atas `app.py`.

## 1. Push ke GitHub

```bash
cd sales-dashboard
git init
git add .
git commit -m "Initial commit: sales intelligence dashboard"
git branch -M main
git remote add origin https://github.com/USERNAME/NAMA-REPO.git
git push -u origin main
```

> Catatan ukuran file: GitHub membatasi file hingga 100 MB per file (soft limit disarankan <25 MB
> agar clone/push cepat). File data yang Anda punya saat ini (~2.2 MB dan ~50 KB) aman untuk di-push langsung.

## 2. Deploy ke Streamlit Community Cloud

1. Buka https://share.streamlit.io lalu login dengan akun GitHub.
2. Klik **New app**.
3. Pilih repo, branch `main`, dan **Main file path**: `app.py`.
4. Klik **Deploy**. Streamlit akan otomatis membaca `requirements.txt` dan menjalankan `app.py`.
5. Setelah selesai build (biasanya 1-3 menit), dashboard bisa diakses via URL publik yang diberikan.

## 3. Update data di kemudian hari

Setiap kali ada data omset baru:

```bash
# ganti file di folder data/ dengan file terbaru (nama file tetap sama)
cp /path/ke/Data_Omset_Power_Query_tahun_2026_terbaru.xlsx data/Data_Omset_Power_Query_tahun_2026.xlsx
git add data/Data_Omset_Power_Query_tahun_2026.xlsx
git commit -m "Update data omset"
git push
```

Streamlit Community Cloud otomatis rebuild & reload aplikasi setiap ada push baru ke branch yang di-deploy.
Anda juga tetap bisa memilih **"Upload manual"** di sidebar aplikasi kalau ingin mencoba file lain
tanpa mengubah isi repo (misalnya untuk simulasi/testing sebelum data final di-commit).

## 4. Menjalankan secara lokal (opsional, untuk testing sebelum push)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Kolom yang dibutuhkan data transaksi

`KOTA, TGL FKTR, SUPP, NOMINAL, DEPO, SALES, MONTH, TANGGAL, DIVISI, AREA, SBD/NON, KD GRUP, KD-SUPP, TELE, NOO/NPD`

## Kolom yang dibutuhkan data target (opsional)

`NAMA SALESMAN, DIVISI, SALES YG COVER, DEPO, SUPPLIER, JAN...DEC`
