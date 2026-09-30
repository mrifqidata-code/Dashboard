# Dashboard Area

Aplikasi web ringan (tanpa server, tanpa build) untuk membuka dashboard dari HP. Satu aplikasi, tiga kategori
(menu bawah), masing-masing dengan sub-menu (menu atas). Pilihan center (Semua / KLM / …) diingat di semua halaman.

| Kategori | Sub-menu | Folder | Sumber data |
|---|---|---|---|
| **Center** | Rekap Acquisition (harian / mingguan / bulanan) | `center/` | Feed Rekap Harian, Feed Rekap Bulanan, Feed Rekap Mingguan Area |
| | Retention, Cash & Piutang | segera | Feed Dashboard Tambahan (tab 03, 12) |
| **Student Advisor** | Acquisition (CVR SU→Paid, CVR Baby Tiny, Full Pay Rate, FP SD, Revenue) | `sa/` | Feed Acquisition SA (tab 02B) |
| | Retention, PIP | segera | Feed Dashboard Tambahan (tab 03, 05) |
| **Kelas** | Slot Map | `kelas/` | Feed Slot Map (Reguler Class Availability) |
| | Utilisasi, Coach | segera | Feed Dashboard Tambahan (tab 08, 11) |

- `assets/shell.js` + `assets/shell.css`: menu kategori & sub-menu (daftar di `NAV`), filter center bersama,
  dan utilitas baca Sheet (`DA.fetchPublicCsv`, `DA.parseCSV`, `DA.sheetNum`). Sub-menu baru: tambah folder,
  muat `../assets/shell.js`, panggil `DA.mountNav('<id>')`, lalu isi `href` di `NAV`.
- `index.html` langsung membuka halaman terakhir yang dilihat. `rekap/` dan `slot-map/` hanya mengalihkan
  ke `center/` dan `kelas/` (alamat lama).
- Bisa dipasang ke home screen (manifest + service worker); data terakhir tetap tampil saat offline.

## Cara data dibaca

Halaman membaca Sheet feed langsung dari browser (`/export?format=csv`, cadangan `/gviz/tq`).
Karena itu **semua Sheet feed harus dibagikan "Siapa saja yang memiliki link" (Viewer)**.
Siapa pun yang tahu ID Sheet bisa membaca isinya, termasuk angka revenue dan nama coach.
Sheet Exboard sendiri tetap privat (feed mengambilnya lewat `IMPORTRANGE`).

Kalau dibuka lewat claude.ai (artifact), Slot Map (`kelas/`) tetap membaca lewat connector Google Drive seperti sebelumnya.

## Pasang

1. Bagikan semua Sheet feed: Bagikan → Akses umum → *Siapa saja yang memiliki link* → Viewer.
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
- Rekap: begitu tab **Feed Harian** di Exboard terisi (rumus di `docs/feed-harian.md`),
  harian, mingguan dan bulanan semuanya dihitung dari tabel itu (Acquisition = FP SD, Revenue dari Revenue Record).
  Sebelum itu: mingguan dari feed mingguan (Acquisition = Paid DP + Paid Full), bulanan dari tab 02 Acquisition
  (hanya Reporting Month yang sedang dipilih di Exboard). Target bulan/minggu berjalan dipro-rata sampai hari ini.
- Data ditarik ulang otomatis saat aplikasi dibuka lagi setelah 5 menit.
