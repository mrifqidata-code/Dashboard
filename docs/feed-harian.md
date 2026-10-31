# Tab "Feed Harian" di Exboard

Satu baris per tanggal per center, dihitung di dalam Exboard dari data mentahnya.
Sheet **Feed Rekap Harian (dari Exboard)** hanya mengambil tabel angka ini (`IMPORTRANGE 'Feed Harian'!A1:K`),
jadi data pribadi di Register Trial / Student Database / Revenue Record tidak ikut keluar.

Tempel rumus ini di sel **A1** tab `Feed Harian` (kosongkan sisa tab):

```
=ARRAYFORMULA(LET(
  mulai, EOMONTH(TODAY(),-4)+1,
  i, SEQUENCE((TODAY()-mulai+1)*6),
  tgl, mulai+INT((i-1)/6),
  c, CHOOSE(MOD(i-1,6)+1, "KLM","KWC","PML","TMP","BTU","HIB"),
  su, MAP(tgl, c, LAMBDA(t, k, COUNTIFS('Register Trial'!$A$2:$A, k, 'Register Trial'!$B$2:$B, ">="&t, 'Register Trial'!$B$2:$B, "<"&(t+1), 'Register Trial'!$N$2:$N, TRUE))),
  dp, MAP(tgl, c, LAMBDA(t, k, COUNTIFS('Register Trial'!$A$2:$A, k, 'Register Trial'!$B$2:$B, ">="&t, 'Register Trial'!$B$2:$B, "<"&(t+1), 'Register Trial'!$O$2:$O, "Regis Fee only"))),
  fp, MAP(tgl, c, LAMBDA(t, k, COUNTIFS('Register Trial'!$A$2:$A, k, 'Register Trial'!$B$2:$B, ">="&t, 'Register Trial'!$B$2:$B, "<"&(t+1), 'Register Trial'!$O$2:$O, "Regis + Member Fee")
    + COUNTIFS('Register Trial'!$A$2:$A, k, 'Register Trial'!$B$2:$B, ">="&t, 'Register Trial'!$B$2:$B, "<"&(t+1), 'Register Trial'!$O$2:$O, "Member Fee"))),
  acq, MAP(tgl, c, LAMBDA(t, k, COUNTIFS('Student Database'!$A$2:$A, k, 'Student Database'!$BZ$2:$BZ, ">="&t, 'Student Database'!$BZ$2:$BZ, "<"&(t+1), 'Student Database'!$E$2:$E, "Full"))),
  rev, MAP(tgl, c, LAMBDA(t, k, SUMIFS('Revenue Record'!$H$2:$H, 'Revenue Record'!$A$2:$A, k, 'Revenue Record'!$S$2:$S, ">="&t, 'Revenue Record'!$S$2:$S, "<"&(t+1)))),
  cic, MAP(tgl, c, LAMBDA(t, k, SUMIFS('Revenue Record'!$H$2:$H, 'Revenue Record'!$A$2:$A, k, 'Revenue Record'!$S$2:$S, ">="&t, 'Revenue Record'!$S$2:$S, "<"&(t+1), 'Revenue Record'!$K$2:$K, "Cash Bertahap*Cicilan*"))),
  pel, MAP(tgl, c, LAMBDA(t, k, SUMIFS('Revenue Record'!$H$2:$H, 'Revenue Record'!$A$2:$A, k, 'Revenue Record'!$S$2:$S, ">="&t, 'Revenue Record'!$S$2:$S, "<"&(t+1), 'Revenue Record'!$K$2:$K, "Cash Bertahap*Pelunasan*"))),
  hari, DAY(EOMONTH(tgl,0)),
  ta, VLOOKUP(c, '02 Acquisition'!$A$7:$N$12, 14, FALSE)/hari,
  tr, VLOOKUP(c, '02 Acquisition'!$A$7:$L$12, 12, FALSE)/hari,
  VSTACK(
    {"Tanggal","Center","Show Up","Paid DP","Paid Full","Acquisition","Revenue","Target Acquisition","Target Revenue","Revenue CB Cicilan","Revenue CB Pelunasan"},
    HSTACK(TEXT(tgl,"yyyy-mm-dd"), c, su, dp, fp, acq, rev, ta, tr, cic, pel))))
```

| Kolom | Sama dengan di Exboard |
|---|---|
| Show Up | Register Trial kolom N = TRUE |
| Paid DP / Paid Full | Register Trial kolom O: "Regis Fee only" / "Regis + Member Fee" + "Member Fee" |
| Acquisition | FP SD: Student Database kolom E = "Full", tanggal kolom BZ (rumus H7 tab 02 Acquisition) |
| Revenue | Revenue Record kolom H, tanggal kolom S (rumus K7 tab 02 Acquisition) |
| Target | Target bulanan tab 02 Acquisition (kolom N dan L) ÷ jumlah hari bulan itu |
| Revenue CB Cicilan / Pelunasan | Bagian dari Revenue dengan Revenue Record kolom K (Type Payment) = "Cash Bertahap: Cicilan" / "Cash Bertahap: Pelunasan". Dipakai opsi Revenue "Tanpa cicilan & pelunasan" di Rekap (akad tetap dihitung) |

- COUNTIFS/SUMIFS/INDEX tidak ikut ARRAYFORMULA di Google Sheets, jadi tiap kolom dihitung per baris dengan MAP.
- Kolom J–K (Revenue CB Cicilan / Pelunasan) ditambahkan belakangan. Kalau Feed Harian sudah terpasang dengan rumus lama:
  ganti rumus di A1 tab Feed Harian dengan rumus di atas, lalu di Sheet Feed Rekap Harian ubah rentang
  `'Feed Harian'!A1:I` di sel A1 menjadi `'Feed Harian'!A1:K`. Selama kolom itu belum ada, opsi Revenue
  di Rekap tetap memakai revenue penuh. Bila Type Payment tidak lagi di kolom K Revenue Record, ganti `$K$2:$K`.
- Rentang: awal bulan 3 bulan lalu s.d. hari ini (± 4 bulan). Ubah `-4` di `EOMONTH(TODAY(),-4)` untuk lebih panjang/pendek.
- Kalau Exboard memakai format angka Indonesia dan rumus error, ganti pemisah `,` menjadi `;`
  (dan `{"KLM";"KWC";…}` menjadi `{"KLM"\"KWC"\…}` hanya bila Sheets meminta).
