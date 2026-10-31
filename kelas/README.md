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

## Coach tersedia (klik slot)

Klik/tap sel slot mana pun untuk membuka daftar coach yang bisa mengajar di jam itu:

- **Coach tersedia**: coach yang punya kelas di center itu, masuk hari itu (minimal satu kelas hari itu),
  dan tidak sedang mengajar. Diurutkan: yang sudah mengajar age group slot itu dulu, lalu yang slotnya
  ada di sela jadwal hari itu, lalu jeda terpendek. Tiap coach menampilkan jadwalnya hari itu dengan
  posisi slot yang diklik (kotak putus-putus).
- **Sedang mengajar di jam ini**: bentrok dengan kelas lain.
- **Libur**: tidak punya kelas sama sekali hari itu.

Kelas reguler 45 menit (`DUR`), jadi kelas yang jam mulainya berselisih kurang dari 45 menit dihitung bentrok.
Lane memakai dua jadwal yang bergeser 15 menit, jadi kelas di lane lain yang mulai 15 atau 30 menit lebih awal
atau lebih lambat juga bentrok. Coach center lain tidak ditampilkan.
Slot Tutup (kapasitas 0, tanpa murid) dan tulisan CLOSE di kolom coach bukan kelas.
