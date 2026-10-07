# Dashboard Area

Aplikasi web ringan (tanpa server, tanpa build) untuk membuka dashboard dari HP. Satu aplikasi, tiga kategori
(menu bawah), masing-masing dengan sub-menu (menu atas). Pilihan center (Semua / KLM / …) diingat di semua halaman.

| Kategori | Sub-menu | Folder | Sumber data |
|---|---|---|---|
| **Center** | Rekap Acquisition (harian / mingguan / bulanan) | `center/` | Feed Rekap Harian, Feed Rekap Bulanan, Feed Rekap Mingguan Area |
| | Class Trial (booking → confirmed → show up → paid DP / regis full; per center & age group, periode bebas) | `center/trial/` | Feed Class Trial (QUERY atas sheet Class Trial) |
| | SUNP (show up not paid: alasan belum DP, potential; per center & age group, periode bebas) | `center/sunp/` | Feed SUNP (tab Feed SUNP di Exboard, rumus di `docs/feed-sunp.md`) |
| | Retention (cohort, status follow-up, tren, alasan churn) | `center/retention/` | Feed Dashboard Tambahan (tab 03) |
| | Cash & Piutang (penerimaan, piutang, cash bertahap) | `center/cash/` | Feed Dashboard Tambahan (tab 12) |
| **Student Advisor** | Acquisition (CVR SU→Paid, CVR Baby Tiny, Full Pay Rate, FP SD, Revenue) | `sa/` | Feed Acquisition SA (tab 02B) |
| | Retention per SA Retention (SAR) | `sa/retention/` | Feed Dashboard Tambahan (tab 03) |
| | PIP (band & status PIP / Watch / OK per advisor) | `sa/pip/` | Feed Dashboard Tambahan (tab 05) |
| **Kelas** | Slot Map | `kelas/` | Feed Slot Map (Reguler Class Availability) |
| | Utilisasi per age group & hari, headroom, kelas trial | `kelas/utilisasi/` | Feed Dashboard Tambahan (tab 08) |
| | Coach (beban murid, sesi per hari, level murid) | `kelas/coach/` | Feed Dashboard Tambahan (tab 11, dan tabel G tab 08) |

- `assets/shell.js` + `assets/shell.css`: menu kategori & sub-menu (daftar di `NAV`), filter center bersama,
  dan utilitas baca Sheet (`DA.fetchPublicCsv`, `DA.parseCSV`, `DA.sheetNum`).
- `assets/report.js` + `assets/report.css`: dipakai keenam sub-menu yang membaca **Feed Dashboard Tambahan**.
  Feed itu memuat beberapa tab Exboard berdampingan; baris 1 berisi penanda `##<nama tab>` di kolom awal tiap tab.
  Tabel dicari lewat judulnya ("A. …", "B. …") dan kolom lewat nama header, jadi menggeser baris di Exboard aman
  selama judul dan nama header tidak diubah. Sub-menu baru: tambah folder,
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
- Class Trial: feed menjumlahkan sheet Class Trial (satu baris per sesi trial, tanpa data anak) per tanggal trial ×
  center × age group × jam, mulai awal bulan 6 bulan lalu. Filter jam muncul bila feed punya kolom Jam. Pilihan sumber (Tanpa crosspol / Crosspol saja / Semua) muncul bila feed punya blok Kunci/Jumlah dari tab Feed Crosspol Exboard; bawaannya Tanpa crosspol. Label umur ditampilkan sebagai Baby (6-17 Mo), Tiny (18-36 Mo),
  Little (3-5 Yo), Kids (5-8+ Yo), Star (9-12 Yo). Periode cepat (Hari ini … Bulan lalu) dihitung ulang setiap dibuka;
  tanggal bebas diingat di perangkat. Rentang ≤31 hari dirinci per hari, lebih panjang per bulan.
- SA Retention: tab 03 belum mencatat center tiap SAR, jadi halaman itu selalu menampilkan semua center.
- Coach: beban dinilai terhadap rata-rata coach di center yang sama (merah bila >30% di atasnya);
  sesi per hari merah ≥7, oranye 6. Potensi revenue di tab 08 tercatat dalam ribuan rupiah (dikali ×1.000).
- Data ditarik ulang otomatis saat aplikasi dibuka lagi setelah 5 menit.
