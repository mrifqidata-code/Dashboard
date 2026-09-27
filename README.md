# Dashboard Area

Aplikasi web ringan (tanpa server, tanpa build) untuk membuka dashboard dari HP:

| Halaman | Isi | Sumber data |
|---|---|---|
| `rekap/` | Rekap Area harian / mingguan / bulanan: acquisition, revenue, CVR, DP → FP per center, tren, salin ringkasan WA | **Feed Rekap Harian** (tab Feed Harian, belum dibuat), **Feed Rekap Bulanan** (tab 02 Acquisition), **Feed Rekap Mingguan Area** |
| `sa/` | Acquisition per SA: CVR SU→Paid, CVR Baby Tiny, Full Pay Rate, FP SD, Revenue per Student Advisor, filter center, urutkan per metrik, detail funnel & alasan gagal | Sheet **Feed Acquisition SA (dari Exboard)** (tab 02B Acquisition SA) |
| `slot-map/` | Slot Map Kelas: age group, keterisian & coach per slot | Sheet **Feed Slot Map (dari Exboard)** |

`index.html` langsung membuka dashboard yang terakhir dilihat. Bisa dipasang ke home screen
(manifest + service worker), dan data terakhir tetap tampil saat offline.

## Cara data dibaca

Halaman membaca Sheet feed langsung dari browser (`/export?format=csv`, cadangan `/gviz/tq`).
Karena itu **ketiga Sheet feed harus dibagikan "Siapa saja yang memiliki link" (Viewer)**.
Siapa pun yang tahu ID Sheet bisa membaca isinya, termasuk angka revenue dan nama coach.
Sheet Exboard sendiri tetap privat (feed mengambilnya lewat `IMPORTRANGE`).

Kalau dibuka lewat claude.ai (artifact), `slot-map/` tetap membaca lewat connector Google Drive seperti sebelumnya.

## Pasang

1. Bagikan ketiga Sheet feed: Bagikan → Akses umum → *Siapa saja yang memiliki link* → Viewer.
2. GitHub → repo ini → **Settings → Pages** → *Deploy from a branch* → pilih branch yang berisi file ini, folder `/ (root)`.
3. Buka `https://mrifqidata-code.github.io/Dashboard/` di HP:
   - Android (Chrome): menu ⋮ → **Tambahkan ke layar utama / Instal aplikasi**.
   - iPhone (Safari): tombol Bagikan → **Tambahkan ke Layar Utama**.

## Catatan

- Rekap versi aplikasi hanya untuk melihat data. Input manual, catatan coaching SA/CM, mode bulanan, dan
  pengaturan threshold tetap di dashboard Rekap Mingguan Area di claude.ai.
- Threshold status memakai nilai bawaan (Acquisition/Revenue hijau ≥100%, merah <90%; CVR hijau ≥45%, merah <35%;
  DP → FP hijau ≥80%, merah <70%).
- Revenue di tab 02B Acquisition SA tercatat dalam ribuan rupiah; halaman SA mengalikannya ×1.000.
- Warna di halaman SA: CVR / CVR Baby Tiny / Full Pay Rate hijau bila ≥ rata-rata area, merah bila 30% di bawahnya.
- CVR Baby Tiny = (paid Baby + paid Tiny) / (show up Baby + show up Tiny), dari bagian B tab 02B
  (paid per segmen = show up × CVR segmen). Kolom BT di sheet tidak dipakai karena pembaginya berbeda.
- Rekap: begitu tab **Feed Harian** ada di Exboard dan ID feed-nya diisi di `SRC.day` (`rekap/index.html`),
  harian, mingguan dan bulanan semuanya dihitung dari tabel itu (Acquisition = FP SD, Revenue dari Revenue Record).
  Sebelum itu: mingguan dari feed mingguan (Acquisition = Paid DP + Paid Full), bulanan dari tab 02 Acquisition
  (hanya Reporting Month yang sedang dipilih di Exboard). Target bulan/minggu berjalan dipro-rata sampai hari ini.
- Data ditarik ulang otomatis saat aplikasi dibuka lagi setelah 5 menit.
