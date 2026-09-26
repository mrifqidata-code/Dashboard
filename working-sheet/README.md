# Slot Map Weekly Meeting (Excel)

Working sheet untuk meeting mingguan Center Manager × Senior Coach.

- `Slot_Map_Weekly_Meeting_W40.xlsx` — file untuk meeting W40 (data per 26 Sep 2026).
- `build_working_sheet.py` — generator. Input: CSV dari Sheet "Feed Slot Map (dari Exboard)".

```
python build_working_sheet.py feed.csv Slot_Map_Weekly_Meeting_W41.xlsx --week "W41 · 5–11 Okt 2026" --asof "3 Okt 2026"
```

Isi workbook: Cara Pakai · Ringkasan · satu tab per center (Slot Map per lane berwarna status,
Daftar Diskusi kelas 0 murid & slot kosong dengan saran coach available, tabel Coach) · Master Data.
