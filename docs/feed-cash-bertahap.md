# Feed Cash Bertahap v2 (dari Exboard)

Sheet feed untuk sub-tab **Cash & Piutang → Cash Bertahap** (`center/cash/bertahap/`). Isinya hanya angka gabungan
per center; daftar murid tetap di Exboard. Sheet ini harus dibagikan *Siapa saja yang memiliki link* (Viewer) dan
sekali klik **Allow access** di sel A2 supaya `IMPORTRANGE` bisa membaca Exboard.

| Sel | Isi |
|---|---|
| A1 | `##16 Cash Bertahap` |
| A2 | ringkasan tab 16 Cash Bertahap baris 1–40 (daftar murid tab itu mulai baris 99, jadi tidak ikut) |
| Q1 | `##CB Tracker` |
| Q2 | satu rumus yang menghitung tabel A–E dari CB Tracker (judul kolom di baris 7, data mulai baris 8) |

Kolom CB Tracker dicari lewat **judulnya** di baris 7 (Sparks ID, Center, Nilai Akad, Total Dibayar, Bayar Terakhir,
Hari Sejak Bayar, Jatuh Tempo Berikut, STATUS, Total Tagihan), jadi menyisipkan kolom baru di CB Tracker tidak
menggeser hitungan. Kalau ada judul yang tidak ketemu, sel Q2 menuliskan judul mana yang hilang.
(Versi pertama feed membaca kolom berdasarkan urutan dan ikut salah ketika kolom Status Student Management disisipkan.)

Tabel di Q2 (kolom KLM, KWC, PML, TMP, BTU, HIB, TOTAL):

- **A. Status cicilan (murid)**: total, lunas (= LUNAS di CB Tracker + kurang bayar di bawah batas), masih berjalan,
  on track / telat / warning / default / lainnya.
- **B. Nilai (Rp)**: nilai akad, total tagihan, sudah dibayar, sisa tagihan murid yang masih berjalan,
  sisa berisiko (telat + warning + default), dan selisih yang dianggap diskon.
- **C. Hari sejak bayar per status**: rentang hari tiap status CB Tracker (dipakai sebagai keterangan di halaman).
- **D. Jatuh tempo berikut per tanggal**: murid yang masih berjalan per tanggal jatuh tempo (+ sisa tagihannya).
  Halaman menghitung sendiri "jatuh tempo bulan ini", "sudah lewat", "tunggakan", dan per pekan dari tanggal HP.
- **E. Bayar terakhir per tanggal** (3 bulan terakhir): untuk "sudah bayar bulan ini" dan berapa yang lunas.

Kurang bayar = Total Tagihan − Total Dibayar. Di bawah **Rp 300.000** dianggap lunas (diskon). Untuk mengubah batas,
ganti `lim,300000` di rumus Q2.

Rumus A2:

```
=IMPORTRANGE("<ID Exboard>","'16 Cash Bertahap'!A1:O40")
```

Rumus Q2:

```
=ARRAYFORMULA(LET(exk,"<ID Exboard>",cbRaw,IMPORTRANGE(exk,"'CB Tracker'!A7:AD"),IF(ISERROR(INDEX(cbRaw,1,1)),INDEX(cbRaw,1,1),LET(lim,300000,cbAll,ARRAYFORMULA(IFERROR(cbRaw,"")),hn,ARRAYFORMULA(LOWER(TRIM(REGEXREPLACE(CHOOSEROWS(cbAll,1)&"","[\s*]+"," ")))),cId,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^sparks id")),0),0),cCen,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^center")),0),0),cAkad,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^nilai akad")),0),0),cPaid,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^total dibayar")),0),0),cLast,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^bayar terakhir")),0),0),cDays,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^hari sejak bayar")),0),0),cDue,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^jatuh tempo")),0),0),cStat,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^status$")),0),0),cBill,IFERROR(MATCH(TRUE,ARRAYFORMULA(REGEXMATCH(hn,"^total tagihan")),0),0),miss,TEXTJOIN(", ",TRUE,IF(cId=0,"Sparks ID",""),IF(cCen=0,"Center",""),IF(cAkad=0,"Nilai Akad",""),IF(cPaid=0,"Total Dibayar",""),IF(cLast=0,"Bayar Terakhir",""),IF(cDays=0,"Hari Sejak Bayar",""),IF(cDue=0,"Jatuh Tempo Berikut",""),IF(cStat=0,"STATUS",""),IF(cBill=0,"Total Tagihan","")),IF(LEN(miss),"Judul kolom CB Tracker (baris 7) tidak ditemukan: "&miss,LET(dat,DROP(cbAll,1),okc,ARRAYFORMULA(REGEXMATCH(UPPER(TRIM(CHOOSECOLS(dat,cCen)&"")),"^(KLM|KWC|PML|TMP|BTU|HIB)$")*(LEN(CHOOSECOLS(dat,cId)&"")>0)),cb,FILTER(dat,okc=1),nr,ROWS(cb),ones,SEQUENCE(nr,1,1,0),cen,ARRAYFORMULA(UPPER(TRIM(CHOOSECOLS(cb,cCen)&""))),st,ARRAYFORMULA(UPPER(TRIM(CHOOSECOLS(cb,cStat)&""))),akad,ARRAYFORMULA(IF(ISNUMBER(CHOOSECOLS(cb,cAkad)),CHOOSECOLS(cb,cAkad),0)),paid,ARRAYFORMULA(IF(ISNUMBER(CHOOSECOLS(cb,cPaid)),CHOOSECOLS(cb,cPaid),0)),billRaw,CHOOSECOLS(cb,cBill),hasBill,ARRAYFORMULA(ISNUMBER(billRaw)*(IF(ISNUMBER(billRaw),billRaw,0)>0)),bill,ARRAYFORMULA(IF(hasBill=1,billRaw,0)),kurang,ARRAYFORMULA(bill-paid),lun,ARRAYFORMULA(--(st="LUNAS")),tolr,ARRAYFORMULA((1-lun)*hasBill*(kurang<lim)),lunAdj,ARRAYFORMULA(lun+tolr),act,ARRAYFORMULA(1-lunAdj),isOT,ARRAYFORMULA((st="ON TRACK")*act),isTL,ARRAYFORMULA((st="TELAT")*act),isWR,ARRAYFORMULA((st="WARNING")*act),isDF,ARRAYFORMULA((st="DEFAULT")*act),isOth,ARRAYFORMULA(act-isOT-isTL-isWR-isDF),sisa,ARRAYFORMULA(IF(kurang>0,kurang,0)*act),berisiko,ARRAYFORMULA(sisa*(isTL+isWR+isDF)),diskon,ARRAYFORMULA(IF(kurang>0,kurang,0)*tolr),hariRaw,CHOOSECOLS(cb,cDays),hari,ARRAYFORMULA(IF(ISNUMBER(hariRaw),hariRaw,"")),jtRaw,CHOOSECOLS(cb,cDue),jt,ARRAYFORMULA(IF(ISNUMBER(jtRaw),INT(jtRaw),0)),lpRaw,CHOOSECOLS(cb,cLast),lp,ARRAYFORMULA(IF(ISNUMBER(lpRaw),INT(lpRaw),0)),inKLM,ARRAYFORMULA(--(cen="KLM")),inKWC,ARRAYFORMULA(--(cen="KWC")),inPML,ARRAYFORMULA(--(cen="PML")),inTMP,ARRAYFORMULA(--(cen="TMP")),inBTU,ARRAYFORMULA(--(cen="BTU")),inHIB,ARRAYFORMULA(--(cen="HIB")),cmt,HSTACK(inKLM,inKWC,inPML,inTMP,inBTU,inHIB,ones),matA,HSTACK(ones,lunAdj,lun,tolr,act,isOT,isTL,isWR,isDF,isOth),resA,MMULT(TRANSPOSE(matA),cmt),matB,HSTACK(akad,bill,paid,sisa,berisiko,diskon),resB,ARRAYFORMULA(ROUND(MMULT(TRANSPOSE(matB),cmt),0)),stCol,TRANSPOSE(SPLIT("ON TRACK,TELAT,WARNING,DEFAULT",",")),cntC,MAP(stCol,LAMBDA(sv,SUMPRODUCT(--(st=sv)))),minC,MAP(stCol,LAMBDA(sv,IFERROR(MIN(FILTER(hari,(st=sv)*ISNUMBER(hari))),""))),maxC,MAP(stCol,LAMBDA(sv,IFERROR(MAX(FILTER(hari,(st=sv)*ISNUMBER(hari))),""))),dueOk,ARRAYFORMULA(act*(jt>0)),jtD,ARRAYFORMULA(IF(dueOk=1,jt,-1)),dts,SORT(UNIQUE(FILTER(jt,dueOk=1))),dm,ARRAYFORMULA(--(jtD=TRANSPOSE(dts))),outD,IFERROR(HSTACK(ARRAYFORMULA(TEXT(dts,"yyyy-mm-dd")),ARRAYFORMULA(IF(SEQUENCE(ROWS(dts)),"")),MMULT(TRANSPOSE(dm),cmt),ARRAYFORMULA(ROUND(MMULT(TRANSPOSE(dm),HSTACK(ARRAYFORMULA(inKLM*sisa),ARRAYFORMULA(inKWC*sisa),ARRAYFORMULA(inPML*sisa),ARRAYFORMULA(inTMP*sisa),ARRAYFORMULA(inBTU*sisa),ARRAYFORMULA(inHIB*sisa),sisa)),0))),"-"),lpMin,EOMONTH(MAX(lp),-3)+1,lpOk,ARRAYFORMULA((lp>=lpMin)*(lp>0)),lpD,ARRAYFORMULA(IF(lpOk=1,lp,-1)),dtsE,SORT(UNIQUE(FILTER(lp,lpOk=1))),dmE,ARRAYFORMULA(--(lpD=TRANSPOSE(dtsE))),outE,IFERROR(HSTACK(ARRAYFORMULA(TEXT(dtsE,"yyyy-mm-dd")),ARRAYFORMULA(IF(SEQUENCE(ROWS(dtsE)),"")),MMULT(TRANSPOSE(dmE),cmt),MMULT(TRANSPOSE(dmE),HSTACK(ARRAYFORMULA(inKLM*lunAdj),ARRAYFORMULA(inKWC*lunAdj),ARRAYFORMULA(inPML*lunAdj),ARRAYFORMULA(inTMP*lunAdj),ARRAYFORMULA(inBTU*lunAdj),ARRAYFORMULA(inHIB*lunAdj),lunAdj))),"-"),IFNA(VSTACK("CB TRACKER — CASH BERTAHAP PER CENTER","Hanya angka gabungan per center dari tab CB Tracker. Kolom dicari lewat judulnya di baris 7. Selisih tagihan di bawah batas dianggap lunas (diskon).",HSTACK("Bayar terakhir paling baru:",TEXT(MAX(lp),"yyyy-mm-dd")),HSTACK("Batas selisih dianggap lunas (Rp):",lim),"","A. STATUS CICILAN (MURID)",HSTACK("Metric","","KLM","KWC","PML","TMP","BTU","HIB","TOTAL"),HSTACK(VSTACK("Total murid","Lunas","Lunas menurut CB Tracker","Lunas karena selisih < 300rb","Masih berjalan","On track","Telat","Warning","Default","Lainnya"),ARRAYFORMULA(IF(SEQUENCE(10),"")),resA),"","B. NILAI CASH BERTAHAP (RP)",HSTACK("Metric","","KLM","KWC","PML","TMP","BTU","HIB","TOTAL"),HSTACK(VSTACK("Nilai akad","Total tagihan","Sudah dibayar","Sisa tagihan (masih berjalan)","Sisa tagihan berisiko (telat + warning + default)","Selisih dianggap diskon (< 300rb)"),ARRAYFORMULA(IF(SEQUENCE(6),"")),resB),"","C. HARI SEJAK BAYAR PER STATUS CB TRACKER",HSTACK("Status","","Murid","Min hari","Max hari"),HSTACK(stCol,ARRAYFORMULA(IF(SEQUENCE(4),"")),cntC,minC,maxC),"","D. JATUH TEMPO BERIKUT PER TANGGAL (MASIH BERJALAN)",HSTACK("Tanggal","","KLM","KWC","PML","TMP","BTU","HIB","TOTAL","Rp KLM","Rp KWC","Rp PML","Rp TMP","Rp BTU","Rp HIB","Rp TOTAL"),outD,"","E. BAYAR TERAKHIR PER TANGGAL (3 BULAN TERAKHIR)",HSTACK("Tanggal","","KLM","KWC","PML","TMP","BTU","HIB","TOTAL","Lunas KLM","Lunas KWC","Lunas PML","Lunas TMP","Lunas BTU","Lunas HIB","Lunas TOTAL"),outE),"")))))))
```
