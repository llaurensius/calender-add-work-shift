#!/usr/bin/env python3
"""
Google Calendar Work Shift Importer (Excel to .ICS)
Mengonversi jadwal shift bulanan dari file Excel (.xlsx) ke file iCalendar (.ics).

Fitur:
- 100% menggunakan modul bawaan Python (Zero external dependencies).
- Otomatis mendeteksi bulan dan tahun dari nama file (misal: oktober-2026.xlsx).
- Menghasilkan file terpisah: shift kerja, hari libur (dayoff), dan file gabungan.
- Menangani shift malam (NM1) yang melewati tengah malam (H+1).
- Menyertakan notifikasi alarm pengingat sebelum shift dimulai.
"""

import os
import sys
import re
import zipfile
import argparse
from datetime import datetime, timedelta, timezone
import xml.etree.ElementTree as ET

# Pemetaan nama bulan (Bahasa Indonesia & Bahasa Inggris)
MONTH_MAPPING = {
    "januari": 1, "january": 1, "jan": 1,
    "februari": 2, "february": 2, "pebruari": 2, "feb": 2,
    "maret": 3, "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "mei": 5, "may": 5,
    "juni": 6, "june": 6, "jun": 6,
    "juli": 7, "july": 7, "jul": 7,
    "agustus": 8, "august": 8, "agu": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "oktober": 10, "october": 10, "okt": 10, "oct": 10,
    "november": 11, "nov": 11,
    "desember": 12, "december": 12, "des": 12, "dec": 12
}

# Aturan Shift
SHIFT_CONFIG = {
    "NP4": {
        "title": "Shift Pagi - Kerja",
        "description": "Jadwal Kerja Shift Pagi (NP4: 07:00 - 15:00 WIB)",
        "start_time": (7, 0),
        "end_time": (15, 0),
        "day_offset": 0,
        "is_all_day": False,
        "is_dayoff": False,
    },
    "NS1": {
        "title": "Shift Siang - Kerja",
        "description": "Jadwal Kerja Shift Siang (NS1: 15:00 - 23:00 WIB)",
        "start_time": (15, 0),
        "end_time": (23, 0),
        "day_offset": 0,
        "is_all_day": False,
        "is_dayoff": False,
    },
    "NM1": {
        "title": "Shift Malam - Kerja",
        "description": "Jadwal Kerja Shift Malam (NM1: 23:00 - 07:00 WIB keesokan harinya)",
        "start_time": (23, 0),
        "end_time": (7, 0),
        "day_offset": 1,  # Melewati tengah malam ke H+1
        "is_all_day": False,
        "is_dayoff": False,
    },
    "DAYOFF": {
        "title": "Libur",
        "description": "Hari Libur (Day Off)",
        "is_all_day": True,
        "is_dayoff": True,
    }
}


def col2num(col_str: str) -> int:
    """Mengubah huruf kolom Excel (A, B, ..., Z, AA, dst.) menjadi indeks angka 1-based."""
    num = 0
    for c in col_str.upper():
        if 'A' <= c <= 'Z':
            num = num * 26 + (ord(c) - ord('A')) + 1
    return num


def parse_xlsx_pure(file_path: str):
    """Membaca sheet pertama file .xlsx menggunakan zipfile dan XML bawaan."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File tidak ditemukan: {file_path}")

    with zipfile.ZipFile(file_path, "r") as z:
        # 1. Baca shared strings jika ada
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in tree.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                # Mendukung elemen <t> tunggal atau potongan formatted text <r><t>
                text_parts = [t.text for t in si.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t") if t.text]
                shared_strings.append("".join(text_parts))

        # 2. Cari file worksheet
        sheet_candidates = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        if not sheet_candidates:
            raise ValueError("Tidak ditemukan worksheet dalam file .xlsx ini.")
        sheet_path = sorted(sheet_candidates)[0]
        
        sheet_tree = ET.fromstring(z.read(sheet_path))
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        
        rows_data = {}
        for r in sheet_tree.findall(".//s:row", ns):
            r_idx = int(r.attrib.get("r", 0))
            rows_data[r_idx] = {}
            for c in r.findall("s:c", ns):
                coord = c.attrib.get("r", "")
                m = re.match(r"([A-Za-z]+)([0-9]+)", coord)
                if not m:
                    continue
                col_idx = col2num(m.group(1))
                val_elem = c.find("s:v", ns)
                val = val_elem.text if val_elem is not None else ""
                
                # Cek jika bertipe shared string
                if c.attrib.get("t") == "s" and val.isdigit():
                    idx = int(val)
                    val = shared_strings[idx] if idx < len(shared_strings) else ""
                elif c.attrib.get("t") == "inlineStr":
                    is_t = c.find(".//s:t", ns)
                    val = is_t.text if is_t is not None else ""
                
                rows_data[r_idx][col_idx] = str(val).strip()

    return rows_data


def detect_month_year_from_filename(filename: str):
    """Mendeteksi bulan dan tahun dari nama file (contoh: oktober-2026.xlsx)."""
    base = os.path.basename(filename).lower()
    
    # Deteksi tahun 4 digit (misal: 2026, 2027)
    year_match = re.search(r"\b(20[2-9][0-9])\b", base)
    year = int(year_match.group(1)) if year_match else None

    # Deteksi bulan
    month = None
    for name, m_num in MONTH_MAPPING.items():
        if re.search(rf"\b{name}\b", base):
            month = m_num
            break

    # Alternatif format angka: misal 2026-10 atau 10-2026
    if not month:
        num_match = re.search(r"[-_](0?[1-9]|1[0-2])[-_]", base)
        if num_match:
            month = int(num_match.group(1))

    return month, year


def format_ics_datetime(dt: datetime) -> str:
    """Format tanggal jam lokal YYYYMMDDTHHMMSS."""
    return dt.strftime("%Y%m%dT%H%M%S")


def format_ics_date(dt: datetime) -> str:
    """Format tanggal seharian YYYYMMDD."""
    return dt.strftime("%Y%m%d")


def build_ics_content(calendar_name: str, events: list, reminder_minutes: int = 60) -> str:
    """Menyusun file .ics sesuai standar RFC 5545 dengan timezone Asia/Jakarta."""
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Work Shift Importer//ID",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{calendar_name}",
        "X-WR-TIMEZONE:Asia/Jakarta",
        "BEGIN:VTIMEZONE",
        "TZID:Asia/Jakarta",
        "X-LIC-LOCATION:Asia/Jakarta",
        "BEGIN:STANDARD",
        "TZOFFSETFROM:+0700",
        "TZOFFSETTO:+0700",
        "TZNAME:WIB",
        "DTSTART:19700101T000000",
        "END:STANDARD",
        "END:VTIMEZONE",
    ]

    now_stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

    for ev in events:
        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{ev['uid']}")
        lines.append(f"DTSTAMP:{now_stamp}")
        lines.append(f"SUMMARY:{ev['summary']}")
        if ev.get("description"):
            lines.append(f"DESCRIPTION:{ev['description']}")
        lines.append("STATUS:CONFIRMED")

        if ev.get("is_all_day"):
            # All-day event: DTSTART & DTEND dengan format DATE (DTEND adalah hari berikutnya)
            lines.append(f"DTSTART;VALUE=DATE:{format_ics_date(ev['start'])}")
            lines.append(f"DTEND;VALUE=DATE:{format_ics_date(ev['end'])}")
            lines.append("TRANSP:TRANSPARENT")
        else:
            lines.append(f"DTSTART;TZID=Asia/Jakarta:{format_ics_datetime(ev['start'])}")
            lines.append(f"DTEND;TZID=Asia/Jakarta:{format_ics_datetime(ev['end'])}")
            lines.append("TRANSP:OPAQUE")

            # Tambahkan Alarm/Pengingat jika disetel
            if reminder_minutes > 0:
                lines.append("BEGIN:VALARM")
                lines.append("ACTION:DISPLAY")
                lines.append(f"DESCRIPTION:Pengingat: {ev['summary']}")
                lines.append(f"TRIGGER:-PT{reminder_minutes}M")
                lines.append("END:VALARM")

        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def process_schedule(file_path: str, year: int, month: int, reminder_minutes: int = 60):
    """Mengekstrak shift dari Excel dan menghasilkan list event."""
    rows = parse_xlsx_pure(file_path)

    # Baris 1: Tanggal, Baris 2: Shift
    date_row = rows.get(1, {})
    shift_row = rows.get(2, {})

    if not date_row or not shift_row:
        raise ValueError("Format file Excel tidak sesuai. Baris 1 harus Tanggal dan Baris 2 harus Jadwal.")

    work_events = []
    dayoff_events = []
    all_events = []

    # Kolom tanggal mulai dari kolom B (indeks 2)
    col_indices = sorted([c for c in date_row.keys() if c >= 2])
    
    for c in col_indices:
        day_val = date_row.get(c, "")
        if not day_val:
            continue
        try:
            day_num = int(float(day_val))
        except ValueError:
            continue

        raw_shift = shift_row.get(c, "").strip()
        if not raw_shift:
            continue

        shift_key = raw_shift.upper()
        if "OFF" in shift_key or "LIBUR" in shift_key:
            shift_key = "DAYOFF"

        if shift_key not in SHIFT_CONFIG:
            print(f"⚠️  Peringatan: Kode shift '{raw_shift}' pada tanggal {day_num} tidak dikenali, dilewati.")
            continue

        config = SHIFT_CONFIG[shift_key]
        
        # Buat tanggal event
        try:
            start_date = datetime(year, month, day_num)
        except ValueError as e:
            print(f"⚠️  Peringatan: Tanggal tidak valid {year}-{month}-{day_num}: {e}")
            continue

        uid_base = f"{shift_key.lower()}-{year}{month:02d}{day_num:02d}@shiftcalendar"

        if config["is_all_day"]:
            # All-day event (Libur): End date adalah tanggal berikutnya
            end_date = start_date + timedelta(days=1)
            event_obj = {
                "uid": uid_base,
                "summary": config["title"],
                "description": config.get("description", ""),
                "start": start_date,
                "end": end_date,
                "is_all_day": True,
                "is_dayoff": True,
                "raw_code": raw_shift,
                "day": day_num,
            }
            dayoff_events.append(event_obj)
            all_events.append(event_obj)
        else:
            s_hour, s_min = config["start_time"]
            e_hour, e_min = config["end_time"]
            day_off = config.get("day_offset", 0)

            dt_start = datetime(year, month, day_num, s_hour, s_min)
            dt_end = datetime(year, month, day_num, e_hour, e_min) + timedelta(days=day_off)

            event_obj = {
                "uid": f"{uid_base}-{s_hour:02d}{s_min:02d}",
                "summary": config["title"],
                "description": config.get("description", ""),
                "start": dt_start,
                "end": dt_end,
                "is_all_day": False,
                "is_dayoff": False,
                "raw_code": raw_shift,
                "day": day_num,
            }
            work_events.append(event_obj)
            all_events.append(event_obj)

    return work_events, dayoff_events, all_events


def main():
    parser = argparse.ArgumentParser(
        description="Konversi Jadwal Kerja Shift Excel (.xlsx) ke Google Calendar iCalendar (.ics)"
    )
    parser.add_argument(
        "file",
        nargs="?",
        default=None,
        help="Path file Excel (.xlsx). Jika dikosongkan, script akan mencari otomatis file .xlsx di folder saat ini."
    )
    parser.add_argument("-m", "--month", type=int, help="Nomor bulan (1-12). Jika tidak diisi, otomatis deteksi dari nama file.")
    parser.add_argument("-y", "--year", type=int, help="Tahun 4 digit (misal: 2026). Jika tidak diisi, otomatis deteksi dari nama file.")
    parser.add_argument("-r", "--reminder", type=int, default=60, help="Menit pengingat/alarm sebelum shift dimulai (default: 60 menit, set 0 untuk mematikan).")
    parser.add_argument("-o", "--outdir", default=".", help="Direktori penyimpanan file .ics (default: folder saat ini).")
    parser.add_argument("-p", "--prefix", help="Prefix nama file output .ics (default: otomatis sesuai nama file/bulan).")

    args = parser.parse_args()

    # 1. Tentukan file Excel input
    input_file = args.file
    if not input_file:
        xlsx_files = [f for f in os.listdir(".") if f.endswith(".xlsx") and not f.startswith("~$")]
        if not xlsx_files:
            print("❌ Tidak ditemukan file .xlsx di folder ini. Harap tentukan nama file:")
            print("   python3 export_shifts.py nama-file.xlsx")
            sys.exit(1)
        # Prioritaskan file yang ada bulan/tahun
        input_file = xlsx_files[0]
        print(f"ℹ️  Menggunakan file otomatis: {input_file}")

    if not os.path.exists(input_file):
        print(f"❌ File tidak ditemukan: {input_file}")
        sys.exit(1)

    # 2. Deteksi atau gunakan Tahun dan Bulan
    det_month, det_year = detect_month_year_from_filename(input_file)
    month = args.month or det_month
    year = args.year or det_year

    if not month or not year:
        now = datetime.now()
        month = month or now.month
        year = year or now.year
        print(f"⚠️  Bulan/Tahun tidak terdeteksi otomatis. Menggunakan default: Bulan {month}, Tahun {year}")
    else:
        print(f"📅 Target Jadwal: Bulan {month} Tahun {year}")

    # 3. Tentukan output prefix dan direktori
    os.makedirs(args.outdir, exist_ok=True)
    if args.prefix:
        prefix = args.prefix
    else:
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        prefix = base_name

    # 4. Proses konversi
    work_events, dayoff_events, all_events = process_schedule(input_file, year, month, args.reminder)

    if not all_events:
        print("❌ Tidak ada event shift yang berhasil diproses.")
        sys.exit(1)

    # 5. Tulis file .ics
    shift_file = os.path.join(args.outdir, f"{prefix}-shift-kerja.ics")
    dayoff_file = os.path.join(args.outdir, f"{prefix}-dayoff.ics")
    combined_file = os.path.join(args.outdir, f"{prefix}-semua-jadwal.ics")

    # Shift kerja
    with open(shift_file, "w", encoding="utf-8") as f:
        f.write(build_ics_content("Jadwal Shift Kerja", work_events, args.reminder))

    # Day off
    with open(dayoff_file, "w", encoding="utf-8") as f:
        f.write(build_ics_content("Hari Libur", dayoff_events, reminder_minutes=0))

    # Gabungan
    with open(combined_file, "w", encoding="utf-8") as f:
        f.write(build_ics_content("Semua Jadwal", all_events, args.reminder))

    # 6. Tampilkan ringkasan
    print("\n" + "=" * 55)
    print("✅ BERHASIL MENGHASILKAN FILE KALENDER (.ICS)")
    print("=" * 55)
    print(f"📊 Ringkasan Event ({len(all_events)} Hari):")
    np4_count = sum(1 for e in work_events if e["raw_code"].upper() == "NP4")
    ns1_count = sum(1 for e in work_events if e["raw_code"].upper() == "NS1")
    nm1_count = sum(1 for e in work_events if e["raw_code"].upper() == "NM1")
    off_count = len(dayoff_events)
    print(f"   • Shift Pagi (NP4)  : {np4_count} hari")
    print(f"   • Shift Siang (NS1) : {ns1_count} hari")
    print(f"   • Shift Malam (NM1) : {nm1_count} hari (melewati tengah malam)")
    print(f"   • Libur (Day Off)   : {off_count} hari")
    print(f"   • Notifikasi Alarm  : {args.reminder} menit sebelum jam shift")

    print("\n📁 File Output Tersimpan di:")
    print(f"   1. {shift_file}")
    print(f"      👉 Impor ke Kalender Khusus 'Jadwal Kerja' (1 warna seragam)")
    print(f"   2. {dayoff_file}")
    print(f"      👉 Impor ke Kalender Khusus 'Hari Libur' (atur warna kalender jadi MERAH)")
    print(f"   3. {combined_file}")
    print(f"      👉 (Opsional) File gabungan seluruh shift + libur")

    print("\n💡 Panduan Singkat Impor ke Google Calendar:")
    print("   1. Buka Google Calendar di browser (https://calendar.google.com).")
    print("   2. Di menu kiri samping 'Other calendars' (Kalender lain), klik tanda '+' -> 'Create new calendar'.")
    print("      - Buat kalender: 'Jadwal Kerja' (beri warna pilihan Anda).")
    print("      - Buat kalender: 'Hari Libur' (klik titik tiga di samping kalender -> pilih warna MERAH).")
    print("   3. Masuk ke Settings (ikon roda gigi) -> 'Import & export'.")
    print(f"      - Unggah '{os.path.basename(shift_file)}' -> pilih kalender tujuan: 'Jadwal Kerja'.")
    print(f"      - Unggah '{os.path.basename(dayoff_file)}' -> pilih kalender tujuan: 'Hari Libur'.")
    print("   4. Selesai! Jadwal kerja rapi dan semua hari libur langsung berwarna merah.")
    print("=" * 55 + "\n")


if __name__ == "__main__":
    main()
