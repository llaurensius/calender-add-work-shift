# User Stories & Spesifikasi Teknis: Google Calendar Work Shift Importer

Dokumen ini mendokumentasikan spesifikasi kebutuhan pengguna (*User Stories*), aturan bisnis (*business rules*), serta keputusan arsitektur teknis untuk pembuatan automasi impor jadwal kerja shift ke Google Calendar.

---

## 1. Ringkasan & Tujuan Proyek

Mengembangkan script automasi Python ringan tanpa dependensi eksternal (*zero-dependency / standard library*) yang membaca file jadwal bulanan Excel (seperti `oktober-2026.xlsx`, `november-2026.xlsx`, dst.) dan mengonversinya menjadi file iCalendar (`.ics`).

Pengguna dapat menggunakan script ini setiap bulan dengan fleksibilitas input dan output yang disesuaikan. File hasil ekspor dapat langsung diimpor ke **Kalender Khusus** di Google Calendar sehingga jadwal kerja seragam dalam satu warna dan hari libur tampil dengan warna merah secara terpisah.

---

## 2. Definisi Shift & Pemetaan Event

Berdasarkan keputusan pengguna:

| Kode Excel | Nama Shift | Judul Event Kalender | Waktu Mulai | Waktu Selesai | Karakteristik Waktu | Reminder |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **NP4** | Shift Pagi | `Shift Pagi - Kerja` | 07:00 WIB | 15:00 WIB | 8 Jam (hari sama) | 60 Menit sebelum mulai |
| **NS1** | Shift Siang | `Shift Siang - Kerja` | 15:00 WIB | 23:00 WIB | 8 Jam (hari sama) | 60 Menit sebelum mulai |
| **NM1** | Shift Malam | `Shift Malam - Kerja` | 23:00 WIB | 07:00 WIB (+1) | 8 Jam (**lintas hari / H+1**) | 60 Menit sebelum mulai |
| **dayoff** | Libur | `Libur` | Seharian | Seharian | *All-Day Event* (24 jam) | Opsional / Tanpa alarm kerja |

> **Zona Waktu:** `Asia/Jakarta` (WIB / UTC+7).

---

## 3. Strategi Pewarnaan & Struktur File Output (Opsi B)

Karena Google Calendar menentukan warna event berdasarkan kalender tujuan impor file `.ics`:

1. **`[prefix]-shift-kerja.ics`**
   - Berisi seluruh shift kerja aktif (`Shift Pagi - Kerja`, `Shift Siang - Kerja`, `Shift Malam - Kerja`).
   - Dilengkapi alarm/reminder (default 60 menit sebelum shift).
   - Diimpor ke Kalender Khusus di Google Calendar (misal: **"Jadwal Kerja"**) dengan 1 warna seragam.
2. **`[prefix]-dayoff.ics`**
   - Berisi seluruh jadwal libur seharian dengan judul `Libur`.
   - Diimpor ke Kalender Khusus kedua (misal: **"Hari Libur"**) yang warnanya diset **Merah**.
3. **`[prefix]-semua-jadwal.ics`** *(Bonus)*
   - File gabungan seluruh shift dan libur jika sewaktu-waktu ingin impor sekaligus ke dalam 1 kalender saja.

---

## 4. User Stories

### US-01: Pembacaan File Excel Tanpa Ketergantungan Eksternal
- **Sebagai** pengguna,
- **Saya ingin** menjalankan script tanpa perlu menginstall library pihak ketiga (`pip install openpyxl`),
- **Sehingga** script selalu siap pakai di komputer mana pun secara praktis dan ringan.
- **Kriteria Penerimaan:**
  - Script membaca format XLSX secara langsung menggunakan modul bawaan Python (`zipfile`, `xml.etree.ElementTree`).
  - Mampu membaca baris tanggal (Baris 1) dan baris jadwal shift (Baris 2) dinamis (28 s/d 31 hari).

### US-02: Konversi Jadwal Shift Kerja Lintas Hari & Notifikasi Pengingat
- **Sebagai** pengguna,
- **Saya ingin** jadwal shift kerja memiliki judul `Shift [Pagi/Siang/Malam] - Kerja`, jam kerja akurat, dan alarm notifikasi 60 menit sebelumnya,
- **Sehingga** saya selalu diingatkan sebelum waktu dinas dimulai.
- **Kriteria Penerimaan:**
  - Shift NP4: 07:00 - 15:00 WIB.
  - Shift NS1: 15:00 - 23:00 WIB.
  - Shift NM1: 23:00 - 07:00 WIB keesokan harinya (H+1).
  - Menyertakan komponen `VALARM` standar iCalendar (RFC 5545).

### US-03: Konversi Hari Libur Menjadi All-Day Event "Libur"
- **Sebagai** pengguna,
- **Saya ingin** hari libur dibuat sebagai event seharian penuh dengan judul `Libur`,
- **Sehingga** mudah dikenali dan bisa diwarnai merah melalui kalender khusus.
- **Kriteria Penerimaan:**
  - Format `DTSTART;VALUE=DATE` dan `DTEND;VALUE=DATE` (inklusif/eksklusif standar iCalendar).
  - Judul event: `Libur`.

### US-04: Fleksibilitas Penggunaan Rutin Bulanan
- **Sebagai** pengguna,
- **Saya ingin** dapat menggunakan script ini setiap bulan dengan fleksibel (misal otomatis deteksi bulan & tahun dari nama file, atau lewat argumen CLI),
- **Sehingga** saya tidak perlu mengedit kode program setiap kali menerima jadwal bulan baru.
- **Kriteria Penerimaan:**
  - Otomatis mendeteksi nama bulan & tahun dari nama file (misal: `oktober-2026.xlsx`, `november-2026.xlsx`).
  - Menyediakan argumen baris perintah (`--month`, `--year`, `--output-dir`, `--reminder`, `--prefix`).
  - Menampilkan ringkasan jumlah shift dan libur yang berhasil diekspor di terminal.
