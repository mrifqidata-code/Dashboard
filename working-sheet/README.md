# Slot Map Weekly Meeting (Excel)

Working sheet untuk meeting mingguan Center Manager × Senior Coach.

- `Slot_Map_Weekly_Meeting_W40.xlsx` — file untuk meeting W40 (data per 26 Sep 2026).
- `build_working_sheet.py` — generator. Input: CSV tab Reguler Class Availability (kolom A:L),
  misalnya dari Sheet "Feed Slot Map (dari Exboard)". Memakai blok Slot Map dari
  `../exboard-slot-map/build_slot_map_tab.py`, jadi tampilannya sama dengan tab Slot Map di Exboard.

```
python build_working_sheet.py feed.csv Slot_Map_Weekly_Meeting_W41.xlsx --week "W41 · 5–11 Okt 2026" --asof "3 Okt 2026"
```

Isi workbook:
- Cara Pakai · Ringkasan
- Satu tab per center: Slot Map (sama dengan Exboard, center tetap), Daftar Diskusi (kelas 0 murid &
  slot kosong, saran coach available, kolom keputusan), tabel Coach.
- Reguler Class Availability: salinan tab Exboard (A:L). Semua rumus membaca tab ini; paste data
  minggu baru ke sini dan semua angka ikut menghitung ulang.
