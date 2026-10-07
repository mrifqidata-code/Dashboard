# Halaman Student (privat)

Cari nama anak lalu tampilkan semua kolom dari tab **Student Database** dan **Student Management** Exboard.
Berjalan sebagai web app Apps Script di dalam Exboard dengan akses **Only myself**, jadi data pribadi
tidak lewat feed publik dan tidak tersimpan di repo ini.

1. Exboard → **Extensions → Apps Script**.
2. Ganti isi `Code.gs` dengan [Code.gs](Code.gs). Tambah file HTML bernama `Index` (+ → HTML) dan isi dengan [Index.html](Index.html).
3. **Deploy → New deployment** → jenis **Web app** → Execute as: **Me**, Who has access: **Only myself** → Deploy → izinkan akses.
4. Salin **Web app URL** (berakhiran `/exec`) → pasang sebagai `href` sub-menu Student di `assets/shell.js`.

Kolom nama anak dideteksi otomatis dari judul kolom (mis. "Kids Name", "Student Name", "Nama Anak").
Kalau salah, isi `NAME_COL` di bagian atas `Code.gs`. Setelah mengubah kode: Deploy → Manage deployments → edit → Version: New version.
