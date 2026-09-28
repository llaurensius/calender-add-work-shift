# Google Calendar Work Shift Importer (Excel to .ICS)

Script automasi Python untuk membaca jadwal kerja shift bulanan dari file Excel (`.xlsx`) dan mengonversinya menjadi file iCalendar (`.ics`) standar yang siap diimpor ke Google Calendar.

Dibuat dengan **100% Modul Bawaan Python** (*Zero External Dependencies*). Tidak memerlukan instalasi library pihak ketiga seperti `openpyxl`.

---

## 📋 Fitur Utama

- **Deteksi Otomatis:** Otomatis mendeteksi nama bulan dan tahun dari nama file (misal: `oktober-2026.xlsx`, `november-2026.xlsx`).
- **Penanganan Shift Lintas Hari:** Shift Malam (`NM1`: 23:00 - 07:00 WIB) otomatis berakhir di keesokan harinya (H+1).
- **Alarm / Pengingat Otomatis:** Menyertakan alarm notifikasi kalender (default: 60 menit sebelum shift dimulai).
- **Strategi Warna Kalender:** Menghasilkan file terpisah (`shift-kerja.ics` dan `dayoff.ics`) agar hari libur dapat tampil dengan warna **Merah** di Google Calendar.
- **Fleksibel untuk Rutinitas Bulanan:** Menerima argumen input/output dinamis untuk digunakan setiap bulan.

---

## 🕒 Pemetaan Shift

| Kode Excel | Ringkasan | Judul Event Kalender | Waktu | Durasi | Notifikasi |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **NP4** | Shift Pagi | `Shift Pagi - Kerja` | 07:00 - 15:00 WIB | 8 Jam | 60 Menit sebelumnya |
| **NS1** | Shift Siang | `Shift Siang - Kerja` | 15:00 - 23:00 WIB | 8 Jam | 60 Menit sebelumnya |
| **NM1** | Shift Malam | `Shift Malam - Kerja` | 23:00 - 07:00 WIB (+1) | 8 Jam (Lintas Hari) | 60 Menit sebelumnya |
| **dayoff** | Libur | `Libur` | Seharian (*All-day*) | 24 Jam | - |

---

## 🚀 Cara Penggunaan

### 1. Menjalankan Otomatis (Default)
Cukup jalankan perintah berikut, script akan otomatis mendeteksi file `.xlsx` di folder saat ini:
```bash
python3 export_shifts.py
```

Atau tentukan nama file secara spesifik:
```bash
python3 export_shifts.py oktober-2026.xlsx
```

### 2. Opsi Perintah Tambahan (Fleksibilitas Bulanan)
```bash
# Menentukan bulan dan tahun secara manual (jika nama file tidak berpola bulan-tahun)
python3 export_shifts.py jadwal.xlsx --month 11 --year 2026

# Mengatur waktu alarm pengingat (misal 30 menit sebelum shift, atau 0 untuk menonaktifkan)
python3 export_shifts.py oktober-2026.xlsx --reminder 30

# Menyimpan output ke folder tertentu
python3 export_shifts.py oktober-2026.xlsx --outdir ./hasil-kalender

# Mengubah prefix nama file hasil ekspor
python3 export_shifts.py oktober-2026.xlsx --prefix okt26
```

---

## 📁 File Output yang Dihasilkan

Setelah script dijalankan, akan tercipta 3 file:
1. `[prefix]-shift-kerja.ics` $\rightarrow$ Seluruh Shift Pagi, Siang, dan Malam.
2. `[prefix]-dayoff.ics` $\rightarrow$ Seluruh hari Libur (*All-day event*).
3. `[prefix]-semua-jadwal.ics` $\rightarrow$ File gabungan seluruh shift + libur.

---

## 📱 Panduan Impor ke Google Calendar (Agar Libur Berwarna Merah)

Google Calendar mewarnai event berdasarkan kalender tempat file diimpor. Ikuti langkah mudah berikut:

1. **Buka Google Calendar** di browser komputer ([calendar.google.com](https://calendar.google.com)).
2. **Buat 2 Kalender Khusus** (Cukup lakukan sekali):
   - Di panel kiri, pada bagian **Other calendars** (Kalender lain), klik tanda **`+`** $\rightarrow$ **Create new calendar**.
   - Kalender 1: Beri nama **`Jadwal Kerja`** (pilih warna yang diinginkan, misal Biru/Default).
   - Kalender 2: Beri nama **`Hari Libur`** (klik menu titik tiga di sebelah nama kalender $\rightarrow$ pilih warna **Merah**).
3. **Impor File Kalender:**
   - Klik ikon **Settings** (roda gigi di kanan atas) $\rightarrow$ pilih **Settings** $\rightarrow$ di panel kiri pilih **Import & export**.
   - Pada bagian **Select file from your computer**:
     - Pilih file `*-shift-kerja.ics`, pada pilihan **Add to calendar** pilih **`Jadwal Kerja`** $\rightarrow$ Klik **Import**.
     - Pilih file `*-dayoff.ics`, pada pilihan **Add to calendar** pilih **`Hari Libur`** $\rightarrow$ Klik **Import**.
4. **Selesai!** Seluruh jadwal kerja tampil seragam dan hari libur langsung berwarna merah mencolok.
