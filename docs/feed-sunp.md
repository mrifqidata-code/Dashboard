# Tab "Feed SUNP" di Exboard

SUNP = Show Up Not Paid. Tab ini menjumlahkan Register Trial **tanpa nama / nomor HP**: hanya anak yang show up,
dikelompokkan per tanggal × center × age group × jenis bayar × potential × alasan belum DP.
Sheet **Feed SUNP (dari Exboard)** mengambil tabel ini lewat `IMPORTRANGE 'Feed SUNP'!A1:G`.

Buat tab baru bernama persis `Feed SUNP`, lalu tempel rumus ini di sel **A1**:

```
=QUERY('Register Trial'!A2:U, "select B, A, E, O, T, U, count(A) where A is not null and N = true and B >= date '"&TEXT(EOMONTH(TODAY(),-7)+1,"yyyy-mm-dd")&"' group by B, A, E, O, T, U label B 'Tanggal', A 'Center', E 'Age Group', O 'Paid Fee', T 'Potential', U 'Reason', count(A) 'Show Up' format B 'yyyy-mm-dd'", 0)
```

| Kolom Register Trial | Dipakai sebagai |
|---|---|
| A Center, B Date, E Age Group | pengelompokan (tanggal sama dengan Rekap Harian) |
| N * Show Up = TRUE | hanya anak yang show up |
| O * Type of Paid Fee | "Regis Fee only" = DP, "Regis + Member Fee" / "Member Fee" = Full, lainnya / kosong = SUNP |
| T Potential, U * Reasons Not DP | rincian SUNP |

- Rentang: awal bulan 6 bulan lalu s.d. data terakhir. Ubah `-7` untuk lebih panjang/pendek.
- Kalau Exboard memakai format angka Indonesia dan rumus error, ganti pemisah `,` di luar tanda kutip menjadi `;`.
