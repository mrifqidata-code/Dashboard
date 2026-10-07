/* Halaman Student: ketik nama anak → pilih → tampil semua data dari tab Student Database
   dan Student Management di Exboard. Data dibaca di server Google dan hanya dikirim setelah PIN benar;
   tidak ada feed publik.

   Pasang: buka Exboard → Extensions → Apps Script, tempel file ini sebagai Code.gs dan Index.html.
   PIN: Project Settings (ikon ⚙) → Script Properties → Add script property → STUDENT_PIN = PIN pilihanmu
   (minimal 6 angka). PIN tidak pernah ditulis di kode ini.
   Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone. */

const SOURCES = ['Student Database', 'Student Management'];
// Kolom nama anak dideteksi otomatis dari judul kolom. Kalau salah, isi nama judul kolomnya persis, mis.
// const NAME_COL = {'Student Database': 'Kids Name', 'Student Management': 'Student Name'};
const NAME_COL = {};
const MAX_SUGGEST = 15;
const CACHE_SEC = 600;
const MAX_FAIL = 10, LOCK_SEC = 900;   // 10 PIN salah dalam 15 menit → semua percobaan dikunci 15 menit

// Semua fungsi yang mengembalikan data wajib lolos cek PIN di server
function checkPin_(pin) {
  const want = PropertiesService.getScriptProperties().getProperty('STUDENT_PIN');
  if (!want) throw new Error('PIN belum diatur. Isi Script Property STUDENT_PIN di Apps Script.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinfail') || 0);
  if (fails >= MAX_FAIL) throw new Error('Terlalu banyak PIN salah. Coba lagi 15 menit lagi.');
  if (String(pin || '').trim() !== String(want).trim()) {   // spasi tak sengaja di Script Property diabaikan
    cache.put('pinfail', String(fails + 1), LOCK_SEC);
    throw new Error('PIN_SALAH');
  }
}

function verify(pin) { checkPin_(pin); return true; }

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Student')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const norm_ = s => String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim();

function sheetInfo_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return null;
  const lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  const head = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h).trim());
  const find = re => head.findIndex(h => re.test(h));
  let nameIdx = NAME_COL[name] ? head.indexOf(NAME_COL[name]) : -1;
  if (nameIdx < 0) nameIdx = find(/(student|kid|child|anak|murid|siswa).*(name|nama)|(name|nama).*(student|kid|child|anak|murid|siswa)/i);
  if (nameIdx < 0) nameIdx = find(/^[\s*]*(nama|name)\s*$/i);
  const centerIdx = find(/^[\s*]*(center|cabang|branch)\b/i);
  return {sh, head, nameIdx, centerIdx, lastRow, lastCol};
}

// Indeks ringan: hanya kolom nama + center, disimpan di cache 10 menit (dipecah karena batas 100 KB per kunci).
// Cache per pengguna, supaya kalau web app dijalankan sebagai "User accessing the web app", orang yang tidak
// punya akses ke Exboard tidak bisa membaca indeks milik orang lain.
function index_() {
  const cache = CacheService.getUserCache();
  const meta = cache.get('idx:n');
  if (meta) {
    const parts = cache.getAll(Array.from({length: Number(meta)}, (_, i) => 'idx:' + i));
    const txt = Array.from({length: Number(meta)}, (_, i) => parts['idx:' + i]).join('');
    if (txt) return JSON.parse(txt);
  }
  const idx = {};   // kunci "nama|center" → {n, c, hits: {sumber: jumlah baris}}
  const cols = {};
  SOURCES.forEach(src => {
    const info = sheetInfo_(src);
    if (!info || info.nameIdx < 0 || info.lastRow < 2) { cols[src] = info ? info.head[info.nameIdx] || null : null; return; }
    cols[src] = info.head[info.nameIdx];
    const names = info.sh.getRange(2, info.nameIdx + 1, info.lastRow - 1, 1).getDisplayValues();
    const centers = info.centerIdx >= 0 ? info.sh.getRange(2, info.centerIdx + 1, info.lastRow - 1, 1).getDisplayValues() : null;
    names.forEach((r, i) => {
      const n = String(r[0]).trim(); if (!n) return;
      const c = centers ? String(centers[i][0]).trim().toUpperCase() : '';
      const k = norm_(n) + '|' + c;
      const e = idx[k] || (idx[k] = {n, c, hits: {}});
      e.hits[src] = (e.hits[src] || 0) + 1;
    });
  });
  const out = {entries: Object.values(idx), cols};
  const txt = JSON.stringify(out), size = 90000, put = {};
  for (let i = 0; i * size < txt.length; i++) put['idx:' + i] = txt.slice(i * size, (i + 1) * size);
  put['idx:n'] = String(Object.keys(put).length);
  try { cache.putAll(put, CACHE_SEC); } catch (e) {}
  return out;
}

// Saran nama: semua kata yang diketik harus ada di nama; yang diawali kata pertama tampil lebih dulu
function search(q, pin) {
  checkPin_(pin);
  const words = norm_(q).split(' ').filter(Boolean);
  if (!words.length || norm_(q).length < 2) return {items: [], cols: null};
  const {entries, cols} = index_();
  const hits = entries.filter(e => { const n = norm_(e.n); return words.every(w => n.includes(w)); });
  hits.sort((a, b) => (norm_(b.n).startsWith(words[0]) - norm_(a.n).startsWith(words[0])) || a.n.localeCompare(b.n));
  return {items: hits.slice(0, MAX_SUGGEST), more: Math.max(0, hits.length - MAX_SUGGEST), cols};
}

// Semua baris dengan nama (dan center, bila ada) yang sama, dari kedua tab
function detail(name, center, pin) {
  checkPin_(pin);
  const target = norm_(name), c = String(center || '').toUpperCase();
  return SOURCES.map(src => {
    const info = sheetInfo_(src);
    if (!info || info.nameIdx < 0 || info.lastRow < 2) return {src, head: [], rows: [], missing: true};
    const names = info.sh.getRange(2, info.nameIdx + 1, info.lastRow - 1, 1).getDisplayValues();
    const centers = info.centerIdx >= 0 ? info.sh.getRange(2, info.centerIdx + 1, info.lastRow - 1, 1).getDisplayValues() : null;
    const rows = [];
    names.forEach((r, i) => {
      if (norm_(r[0]) !== target) return;
      if (centers && c && String(centers[i][0]).trim().toUpperCase() !== c) return;
      rows.push(info.sh.getRange(i + 2, 1, 1, info.lastCol).getDisplayValues()[0]);
    });
    return {src, head: info.head, rows};
  });
}
