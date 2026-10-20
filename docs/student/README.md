# Halaman Student (privat)

Cari nama anak lalu tampilkan semua kolom dari tab **Student Database** dan **Student Management** Exboard.
Berjalan sebagai web app Apps Script di dalam Exboard. Tidak perlu login Google: data hanya dikirim
setelah PIN benar (dicek di server, dikunci 15 menit setelah 10 kali salah). Data pribadi tidak lewat
feed publik dan tidak tersimpan di repo ini; PIN juga tidak ada di repo.

1. Exboard → **Extensions → Apps Script**.
2. Ganti isi `Code.gs` dengan [Code.gs](Code.gs). Tambah file HTML bernama `Index` (+ → HTML) dan isi dengan [Index.html](Index.html).
3. **Project Settings (⚙) → Script Properties → Add script property**: `STUDENT_PIN` = PIN pilihanmu (minimal 6 angka).
3b. **Deploy → New deployment** (atau Manage deployments → edit → New version) → **Web app** → Execute as: **Me**, Who has access: **Anyone** → Deploy.
4. Halaman app `kelas/student/` memanggil web app ini lewat `doPost` (tanpa akun Google, jadi aman dari masalah banyak akun di HP).
   URL `/exec`-nya ada di konstanta `API` di halaman itu. Halaman `Index.html` tetap bisa dibuka langsung dari browser.

Web app yang sama juga melayani **daftar murid perlu ditagih** di Cash & Piutang → Cash Bertahap (aksi `cb`):
semua murid di tab **CB Tracker** yang belum lunas (kurang bayar di bawah Rp 300.000 dianggap lunas), hanya kolom yang
perlu untuk menagih dan menghitung sesi. App yang memilih siapa yang tampil (telat, tunggakan, jatuh tempo, sesi lolos),
jadi aturan itu bisa diubah tanpa deploy ulang.
Judul kolom CB Tracker dibaca dari baris 7; jalankan `cekSetup` untuk melihat kolom yang terdeteksi.

Kolom nama anak dideteksi otomatis dari judul kolom (mis. "Kids Name", "Student Name", "Nama Anak").
Kalau salah, isi `NAME_COL` di bagian atas `Code.gs`. Setelah mengubah kode: Deploy → Manage deployments → edit → Version: New version.

Kecepatan: indeks (nama, center, nomor baris) disimpan di cache Apps Script hingga 6 jam. Jalankan sekali
`pasangPemanasan` dari editor supaya indeks dibangun ulang tiap 30 menit di belakang layar. App mengambil
daftar nama sekali per sesi (hanya nama, center, jumlah baris) lalu mencari langsung di HP; detail diambil
lewat nomor baris, tanpa membaca ulang seluruh tab.
