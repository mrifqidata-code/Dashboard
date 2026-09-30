# Slot Map Kelas

Dashboard slot map kelas per center dan lane. Setiap sel (jam × hari) menampilkan sekaligus:
age group, keterisian (terisi/kapasitas + %), dan nama coach.

## Sumber data

- Google Sheet **Feed Slot Map (dari Exboard)** — `1TmXlMMf5waCAEn1dS9W7_MKzK028Kkfmvwy5RhBbbb0`.
  Sel A1 berisi `IMPORTRANGE` ke tab **Reguler Class Availability** di Exboard.
- Halaman membaca feed itu lewat connector Google Drive (`download_file_content`, CSV)
  setiap kali dibuka, sama seperti dashboard Rekap Mingguan Area.
- Di luar claude.ai, feed bisa diunduh sebagai CSV lalu dibuka lewat tombol **Buka CSV**.

## Catatan data

Di tab Reguler Class Availability saat ini blok Lane 2–5 bergeser ke kanan
(Lane n bergeser n−1 kolom). Parser mendeteksi pergeseran dari posisi kolom Slot, jadi:

| Lane | Tidak ada di sumber |
|------|---------------------|
| 2 | coach |
| 3 | status availability, coach |
| 4 | jumlah terisi, status availability |
| 5 | kapasitas, jumlah terisi |

Begitu kolom di sumber dirapikan, dashboard otomatis menampilkan data lengkap.
