# Idle Lambda Finder

Tool untuk mencari channel/lambda DWDM yang masih kosong (idle) dari export
"Manage WDM Trail" (.xlsx). Bandingkan kolom **Source Channel** di semua file
input dengan channel plan standar **Pak Rupas**, lalu
hasilkan satu file Excel mapping USED vs IDLE.

Otomatis menangani file export yang rusak (error `sharedStrings.xml` hilang).

## Cara pakai

1. **Sekali saja**: double-click `setup.bat` *tapi pasti kena block*.
2. Taruh semua file export `Manage_WDM_Trail_*.xlsx` ke folder `input/raw/`.
3. Double-click `run.bat`.
4. **Cara kedua** *kalau step diatas kena block* <br>
run script ini lewat cmd di directory folder ini
  ```
  cd path\ke\folder\find-idle-lambda
  python -m venv .venv
  .venv\Scripts\activate
  pip install -r requirements.txt
  ```
5. Kalau sudah lanjut run script ini di cmd:
```
.venv\Scripts\activate
python src\find_idle_lambda.py --config config\config.yaml
```
6. Kemudian tunggu sambil kedip mata, Hasil ada di `output/Lambda_Mapping_<timestamp>.xlsx`, berisi:
   - **Sheet "Lambda Mapping"**: 80 baris channel, dengan Wavelength, Frequency,
     Status (USED/IDLE, diwarnai hijau/merah), dan info trail yang memakainya
     (Name, Source, Sink, file asal).
   - **Sheet "Ringkasan"**: total channel, jumlah used/idle, persentase, dan
     daftar lengkap channel yang idle.
   - **Sheet "Flex Grid 100G"**: pasangan channel 50G idle yang berurutan
     (misal channel 5 & 6), dengan center frequency/wavelength slot 100G-nya
     (titik tengah + rentang ±50GHz).

## Struktur folder

```
idle-lambda-finder/
  run.bat              ← double-click untuk jalankan
  setup.bat            ← double-click sekali di awal
  config/
    config.yaml         ← channel plan & pengaturan folder
  src/
    find_idle_lambda.py ← logika utama
    utils.py             ← baca config, perbaiki file rusak, parsing channel
  input/raw/            ← taruh file export di sini
  output/               ← hasil mapping otomatis tersimpan di sini
```

## Konfigurasi (`config/config.yaml`)

- `channel_column`: nama kolom yang dibaca (default `"Source Channel"`)
- `info_columns`: kolom tambahan yang ikut ditampilkan di hasil (Name, Source, Sink)
- `total_channels`, `freq_channel1_thz`, `spacing_ghz`: parameter channel plan.
  Default diset untuk grid **C-band 50GHz, 80 channel** (channel 1 =
  196.05 THz, channel 80 = 192.10 THz).
  Kalau nanti perlu grid lain (misal L-band atau 96 channel), tinggal ubah
  tiga angka ini.

## Cara kerja parsing channel

Kolom "Source Channel" berformat `C\<no channel>\<wavelength nm>\<freq THz>`,
contoh `C\74\1558.17\192.400`. Script mengambil angka setelah band pertama
(`74`) sebagai nomor channel. Baris dengan channel `00`/`0`/`-` dianggap
kosong dan dilewati.

## Kalau ada file baru yang gagal dibaca

Script sudah otomatis menambal file yang kena bug "sharedStrings.xml hilang"
(umum terjadi pada export dari sistem NMS tertentu). Kalau tetap gagal dengan
error lain, cek pesan `[ERROR]` di layar — biasanya karena nama kolom
"Source Channel" berbeda, bisa disesuaikan di `config.yaml`.

## Pengembangkan lebih lanjut

- **Grid berbeda per link/span**: kalau nanti perlu idle lambda per rute
  (bukan gabungan semua file), tinggal grouping berdasarkan `Name`/link
  sebelum dibandingkan ke channel plan.
- **Export ke PDF/Word**: tinggal konversi hasil mapping pakai `python-docx`
  atau render ke PDF dari data yang sama.
- **Bungkus jadi .exe**: aman dibungkus PyInstaller karena tidak ada
  dependency Excel/COM (`pyinstaller --onefile --noupx src\find_idle_lambda.py`).
- Waduh lupa nambahin Channel Flexgrid nya cuy.
