# Dashboard Area

Aplikasi web ringan (tanpa server, tanpa build) untuk membuka dua dashboard dari HP:

| Halaman | Isi | Sumber data |
|---|---|---|
| `rekap/` | Rekap Mingguan Area: acquisition, revenue, CVR, DP → FP per center, tren 8 minggu, salin ringkasan WA | Sheet **Feed Rekap Mingguan Area (dari Exboard)** |
| `slot-map/` | Slot Map Kelas: age group, keterisian & coach per slot | Sheet **Feed Slot Map (dari Exboard)** |

`index.html` langsung membuka dashboard yang terakhir dilihat. Bisa dipasang ke home screen
(manifest + service worker), dan data terakhir tetap tampil saat offline.

## Cara data dibaca

Halaman membaca Sheet feed langsung dari browser (`/export?format=csv`, cadangan `/gviz/tq`).
Karena itu **kedua Sheet feed harus dibagikan "Siapa saja yang memiliki link" (Viewer)**.
Siapa pun yang tahu ID Sheet bisa membaca isinya, termasuk angka revenue dan nama coach.
Sheet Exboard sendiri tetap privat (feed mengambilnya lewat `IMPORTRANGE`).

Kalau dibuka lewat claude.ai (artifact), `slot-map/` tetap membaca lewat connector Google Drive seperti sebelumnya.

## Pasang

1. Bagikan kedua Sheet feed: Bagikan → Akses umum → *Siapa saja yang memiliki link* → Viewer.
2. GitHub → repo ini → **Settings → Pages** → *Deploy from a branch* → pilih branch yang berisi file ini, folder `/ (root)`.
3. Buka `https://mrifqidata-code.github.io/Dashboard/` di HP:
   - Android (Chrome): menu ⋮ → **Tambahkan ke layar utama / Instal aplikasi**.
   - iPhone (Safari): tombol Bagikan → **Tambahkan ke Layar Utama**.

## Catatan

- Rekap versi aplikasi hanya untuk melihat data. Input manual, catatan coaching SA/CM, mode bulanan, dan
  pengaturan threshold tetap di dashboard Rekap Mingguan Area di claude.ai.
- Threshold status memakai nilai bawaan (Acquisition/Revenue hijau ≥100%, merah <90%; CVR hijau ≥45%, merah <35%;
  DP → FP hijau ≥80%, merah <70%).
- Data ditarik ulang otomatis saat aplikasi dibuka lagi setelah 5 menit.
