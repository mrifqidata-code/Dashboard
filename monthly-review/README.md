# Monthly Review · Center Manager Meeting

Deck meeting bulanan Area Manager × Center Manager.

- `Monthly_Review_September_2026_CM.pptx` — review September 2026 dan draft komitmen Oktober (24 slide, ada speaker notes).
- `build_deck.js` — generator (pptxgenjs). Semua angka ada di blok `DATA` di awal file.

Font: **Poppins** (diset di tema). Google Slides sudah punya Poppins. Untuk PowerPoint, install Poppins dulu
(gratis di Google Fonts); kalau tidak terpasang, PowerPoint mengganti dengan font lain dan layout bisa bergeser.

Alur deck:
1. Apa yang terjadi di September: revenue vs target, funnel, alasan SUNP, highlight.
2. Center yang perlu improvement: scorecard dan prioritas.
3. Apa yang akan kita lakukan: 4 strategi (volume, closing, cash & ticket, kapasitas).
4. Komitmen per center: model "jalan ke Rp 600 jt", satu slide per center, ringkasan, ritme monitoring.

## Sumber data

- Exboard › 02 Acquisition, 08 Class Utilization — lewat deck SA League Monthly September 2026 (data per 4 Okt 2026).
- Agustus (leads, show up, paid, revenue): Center Performance Agustus 2026.
- Kelas 0 murid: `../working-sheet/Slot_Map_Weekly_Meeting_W40.xlsx` (data 26 Sep 2026).
- Harga & promo: deck Priority Package & New Pitching Style, Pricing War Strategy, Sparks Swim Competitor Analysis.

## Model komitmen

Target Oktober Rp 600 jt per center. Per center:

- FP butuh = Rp 600 jt ÷ revenue per FP September.
- Trial baru = FP butuh − 60% DP belum lunas − 10% SUNP September (`ASSUME`).
- Show up butuh = trial baru ÷ (target CVR × target FP progress) — target rasio di `GOAL`.
- Komitmen show up dan leads tidak di bawah angka September; dibulatkan ke 5.

## Generate ulang

```
npm install pptxgenjs react react-dom react-icons sharp
NODE_PATH=./node_modules node build_deck.js Monthly_Review_Oktober_2026_CM.pptx
```

Untuk bulan berikutnya: ganti blok `DATA`, `AREA`, `SUNP_REASONS`, teks per center di `CENTER_PAGES`, dan label bulan.
