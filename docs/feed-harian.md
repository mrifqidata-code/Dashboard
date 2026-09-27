# Tab "Feed Harian" di Exboard

Satu baris per tanggal per center, dihitung di dalam Exboard dari data mentahnya.
Sheet **Feed Rekap Harian (dari Exboard)** hanya mengambil tabel angka ini (`IMPORTRANGE 'Feed Harian'!A1:I`),
jadi data pribadi di Register Trial / Student Database / Revenue Record tidak ikut keluar.

Tempel rumus ini di sel **A1** tab `Feed Harian` (kosongkan sisa tab):

```
=ARRAYFORMULA(LET(
  ctr, {"KLM";"KWC";"PML";"TMP";"BTU";"HIB"},
  mulai, EOMONTH(TODAY(),-4)+1,
  i, SEQUENCE((TODAY()-mulai+1)*6),
  tgl, mulai+INT((i-1)/6),
  c, INDEX(ctr, MOD(i-1,6)+1),
  su, COUNTIFS('Register Trial'!$A$2:$A, c, 'Register Trial'!$B$2:$B, ">="&tgl, 'Register Trial'!$B$2:$B, "<"&(tgl+1), 'Register Trial'!$N$2:$N, TRUE),
  dp, COUNTIFS('Register Trial'!$A$2:$A, c, 'Register Trial'!$B$2:$B, ">="&tgl, 'Register Trial'!$B$2:$B, "<"&(tgl+1), 'Register Trial'!$O$2:$O, "Regis Fee only"),
  fp, COUNTIFS('Register Trial'!$A$2:$A, c, 'Register Trial'!$B$2:$B, ">="&tgl, 'Register Trial'!$B$2:$B, "<"&(tgl+1), 'Register Trial'!$O$2:$O, "Regis + Member Fee")
    + COUNTIFS('Register Trial'!$A$2:$A, c, 'Register Trial'!$B$2:$B, ">="&tgl, 'Register Trial'!$B$2:$B, "<"&(tgl+1), 'Register Trial'!$O$2:$O, "Member Fee"),
  acq, COUNTIFS('Student Database'!$A$2:$A, c, 'Student Database'!$BZ$2:$BZ, ">="&tgl, 'Student Database'!$BZ$2:$BZ, "<"&(tgl+1), 'Student Database'!$E$2:$E, "Full"),
  rev, SUMIFS('Revenue Record'!$H$2:$H, 'Revenue Record'!$A$2:$A, c, 'Revenue Record'!$S$2:$S, ">="&tgl, 'Revenue Record'!$S$2:$S, "<"&(tgl+1)),
  hari, DAY(EOMONTH(tgl,0)),
  ta, VLOOKUP(c, '02 Acquisition'!$A$7:$N$12, 14, FALSE)/hari,
  tr, VLOOKUP(c, '02 Acquisition'!$A$7:$L$12, 12, FALSE)/hari,
  VSTACK(
    {"Tanggal","Center","Show Up","Paid DP","Paid Full","Acquisition","Revenue","Target Acquisition","Target Revenue"},
    HSTACK(TEXT(tgl,"yyyy-mm-dd"), c, su, dp, fp, acq, rev, ta, tr))))
```

| Kolom | Sama dengan di Exboard |
|---|---|
| Show Up | Register Trial kolom N = TRUE |
| Paid DP / Paid Full | Register Trial kolom O: "Regis Fee only" / "Regis + Member Fee" + "Member Fee" |
| Acquisition | FP SD: Student Database kolom E = "Full", tanggal kolom BZ (rumus H7 tab 02 Acquisition) |
| Revenue | Revenue Record kolom H, tanggal kolom S (rumus K7 tab 02 Acquisition) |
| Target | Target bulanan tab 02 Acquisition (kolom N dan L) ÷ jumlah hari bulan itu |

- Rentang: awal bulan 3 bulan lalu s.d. hari ini (± 4 bulan). Ubah `-4` di `EOMONTH(TODAY(),-4)` untuk lebih panjang/pendek.
- Kalau Exboard memakai format angka Indonesia dan rumus error, ganti pemisah `,` menjadi `;`
  (dan `{"KLM";"KWC";…}` menjadi `{"KLM"\"KWC"\…}` hanya bila Sheets meminta).
