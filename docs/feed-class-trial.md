# Feed Class Trial (dari Exboard)

Halaman **Center › Class Trial** membaca feed ini untuk menghitung funnel trial
(booking → confirmed → show up → paid DP / regis full) per tanggal × center × age group × jam.

Feed menjumlahkan sheet `Class Trial` di Exboard — **tanpa data anak** — dan **ikut menarik kolom `Class`**
supaya halaman bisa memisahkan sumber lead: `Swim` / `swim` = **Reguler**, `Swim Crosspoll` = **Crosspol**,
`Star` = **Star**.

Tempel rumus ini di sel **A1** spreadsheet feed:

```
=QUERY(IMPORTRANGE("1CLm1NYROYjPEGisRGw-A2ZMEk7OP6d3pDJywxTBJDXA","'Class Trial'!A2:P"),"select Col4, Col1, Col7, Col6, Col5, sum(Col11), sum(Col12), sum(Col13), sum(Col14), sum(Col15), sum(Col16) where Col1 is not null and Col4 >= date '"&TEXT(EOMONTH(TODAY(),-7)+1,"yyyy-mm-dd")&"' group by Col4, Col1, Col7, Col6, Col5 label Col4 'Tanggal', Col1 'Center', Col7 'Age Group', Col6 'Jam', Col5 'Class', sum(Col11) 'Booking', sum(Col12) 'Confirmed', sum(Col13) 'Show Up', sum(Col14) 'Paid DP', sum(Col15) 'Regis Full', sum(Col16) 'Total Paid' format Col4 'yyyy-mm-dd'",0)
```

Hasilnya blok A (A:K):

| Kolom | Isi |
|---|---|
| Tanggal | `Trial Date` (format `yyyy-mm-dd`) |
| Center | kode center |
| Age Group | teks asli sheet (halaman yang menerjemahkan ke Baby/Tiny/Little/Kids/Star) |
| Jam | jam sesi |
| **Class** | `Swim`, `swim`, `Swim Crosspoll`, atau `Star` |
| Booking, Confirmed, Show Up, Paid DP, Regis Full, Total Paid | `sum` kolom # Booking … Total Paid |

- Rentang: awal bulan 6 bulan lalu s.d. data terakhir. Ubah `-7` untuk lebih panjang/pendek.
- Blok `Kunci | Jumlah` lama (kolom M:N) tidak dipakai lagi oleh halaman Class Trial. Biarkan saja; tidak mengganggu
  karena blok A hanya memakai A:K.
- Pemisahan crosspol mengikuti kolom `Class` yang diisi sistem booking sejak **7 Sep 2026**. Sebelum tanggal itu
  semua trial tercatat `Swim`, jadi masuk **Reguler**.
- Kalau Exboard memakai format angka Indonesia dan rumus error, ganti pemisah `,` di luar tanda kutip menjadi `;`.
