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
// Kolom yang TIDAK dikirim ke halaman Student (judul kolom; besar-kecil huruf dan tanda * diabaikan)
const HIDE = {
  'Student Database': ['No', 'Pre Sparks ID Formula', "Child's Nick Name", 'Day of Birth', 'Month of Birth', 'Year of Birth',
    'Regist Day', 'Regist Month', 'Regist Year', 'Main Phone Number', 'Alamat', 'Referral Code', 'Uang Pangkal', 'Member Fee',
    'Weekend Fee', 'Potongan Uang Pangkal', 'Potongan Member Fee', 'DP', 'Pembayaran Selain DP', 'English Student?',
    'Source Lead', 'Get Bag?', 'Contract (Months)', 'Total Sesi'],
  'Student Management': ['Center', 'Student ID', 'Name', 'Age', 'Parents Name', 'Package', 'Day', 'Time (hh:mm)',
    'Next Schedule (yyyy-mm-dd)', "Child's Nickname", 'Main Phone Number', 'Unique Key + Name (copy-paste ke kolom B Attendance Log)',
    'Last attendance log', 'Actual last class']
};
const hkey_ = h => String(h || '').replace(/[\u2018\u2019`]/g, "'").replace(/[*]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
const MAX_SUGGEST = 15;
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

// API untuk halaman Kelas → Student di app (GitHub Pages). App mengirim POST tanpa akun Google:
// body JSON {action: 'verify' | 'list' | 'search' | 'detail', pin, q, name, center}; balasan JSON {ok, data} / {ok: false, error}.
function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const data = b.action === 'verify' ? verify(b.pin)
      : b.action === 'list' ? list(b.pin)
      : b.action === 'search' ? search(b.q, b.pin)
      : b.action === 'detail' ? detail(b.name, b.center, b.pin)
      : (() => { throw new Error('Aksi tidak dikenal'); })();
    out = {ok: true, data};
  } catch (err) {
    out = {ok: false, error: String(err && err.message || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
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

// Indeks: nama + center + nomor baris di tiap tab. Disimpan di cache (dipecah karena batas 100 KB per kunci),
// jadi pencarian dan detail tidak perlu membaca ulang seluruh tab. Dibangun ulang bila cache habis
// atau lewat pemicu waktu (lihat pasangPemanasan).
const IDX_SEC = 21600;   // maks. cache Apps Script = 6 jam; pemicu 30 menit menjaganya tetap baru

function buildIndex_() {
  const idx = {};   // kunci "nama|center" → {n, c, r: {sumber: [nomor baris]}}
  SOURCES.forEach(src => {
    const info = sheetInfo_(src);
    if (!info || info.nameIdx < 0 || info.lastRow < 2) return;
    const names = info.sh.getRange(2, info.nameIdx + 1, info.lastRow - 1, 1).getValues();
    const centers = info.centerIdx >= 0 ? info.sh.getRange(2, info.centerIdx + 1, info.lastRow - 1, 1).getValues() : null;
    names.forEach((r, i) => {
      const n = String(r[0]).trim(); if (!n) return;
      const c = centers ? String(centers[i][0]).trim().toUpperCase() : '';
      const k = norm_(n) + '|' + c;
      const e = idx[k] || (idx[k] = {n, c, r: {}});
      (e.r[src] || (e.r[src] = [])).push(i + 2);
    });
  });
  const out = {entries: Object.values(idx), at: Date.now()};
  const cache = CacheService.getScriptCache(), txt = JSON.stringify(out), size = 90000, put = {};
  let n = 0;
  for (; n * size < txt.length; n++) put['idx:' + n] = txt.slice(n * size, (n + 1) * size);
  put['idx:n'] = String(n);
  try { cache.putAll(put, IDX_SEC); } catch (e) {}
  return out;
}

function index_() {
  const cache = CacheService.getScriptCache(), meta = cache.get('idx:n');
  if (meta) {
    const keys = Array.from({length: Number(meta)}, (_, i) => 'idx:' + i), parts = cache.getAll(keys);
    if (keys.every(k => parts[k] != null)) return JSON.parse(keys.map(k => parts[k]).join(''));
  }
  return buildIndex_();
}

// Nomor kolom → huruf (1 → A, 81 → CC)
const col_ = n => { let t = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) t = String.fromCharCode(65 + (n - 1) % 26) + t; return t; };

const hitsOf_ = e => Object.fromEntries(Object.entries(e.r).map(([s, rows]) => [s, rows.length]));

// Daftar semua nama (untuk pencarian langsung di HP): [nama, center, jumlah baris per tab]
function list(pin) {
  checkPin_(pin);
  return {sources: SOURCES, rows: index_().entries.map(e => [e.n, e.c].concat(SOURCES.map(s => (e.r[s] || []).length)))};
}

// Saran nama (dipakai Index.html): semua kata yang diketik harus ada di nama
function search(q, pin) {
  checkPin_(pin);
  const words = norm_(q).split(' ').filter(Boolean);
  if (!words.length || norm_(q).length < 2) return {items: []};
  const hits = index_().entries.filter(e => { const n = norm_(e.n); return words.every(w => n.includes(w)); });
  hits.sort((a, b) => (norm_(b.n).startsWith(words[0]) - norm_(a.n).startsWith(words[0])) || a.n.localeCompare(b.n));
  return {items: hits.slice(0, MAX_SUGGEST).map(e => ({n: e.n, c: e.c, hits: hitsOf_(e)})), more: Math.max(0, hits.length - MAX_SUGGEST)};
}

// Semua baris dengan nama (dan center) yang sama dari kedua tab. Baris diambil langsung lewat nomor baris
// di indeks; kalau isinya sudah bergeser (data berubah sejak indeks dibuat), indeks dibangun ulang sekali.
function detail(name, center, pin) {
  checkPin_(pin);
  const target = norm_(name), c = String(center || '').toUpperCase();
  const read = idx => {
    const e = idx.entries.find(x => norm_(x.n) === target && x.c === c);
    let stale = false;
    const parts = SOURCES.map(src => {
      const info = sheetInfo_(src);
      if (!info || info.nameIdx < 0) return {src, head: [], rows: [], missing: true};
      const nums = (e && e.r[src]) || [];
      const rows = nums.length ? info.sh.getRangeList(nums.map(r => 'A' + r + ':' + col_(info.lastCol) + r)).getRanges().map(rg => rg.getDisplayValues()[0]) : [];
      if (rows.some(r => norm_(r[info.nameIdx]) !== target)) stale = true;
      // Kolom tersembunyi dibuang di server, jadi tidak pernah terkirim ke HP
      const hide = new Set((HIDE[src] || []).map(hkey_));
      const keep = info.head.map((h, j) => j).filter(j => !hide.has(hkey_(info.head[j])));
      return {src, head: keep.map(j => info.head[j]), rows: rows.map(r => keep.map(j => r[j]))};
    });
    return {parts, stale};
  };
  let res = read(index_());
  if (res.stale) res = read(buildIndex_());
  return res.parts;
}

// Jalankan SEKALI dari editor (pilih "pasangPemanasan" → Run): indeks dibangun ulang tiap 30 menit di belakang
// layar, jadi pencarian dari HP tidak perlu menunggu indeks dibuat.
function pasangPemanasan() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'buildIndex_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('buildIndex_').timeBased().everyMinutes(30).create();
  buildIndex_();
  console.log('Pemanasan terpasang: indeks dibangun ulang tiap 30 menit.');
}

// Jalankan fungsi ini dari editor (pilih "cekSetup" di daftar fungsi → Run) untuk memeriksa pemasangan.
// Tidak menampilkan PIN maupun data murid; hanya status pengaturan dan kolom yang terdeteksi.
function cekSetup() {
  const want = PropertiesService.getScriptProperties().getProperty('STUDENT_PIN');
  console.log(want ? `STUDENT_PIN terisi (${String(want).trim().length} karakter).` : 'STUDENT_PIN BELUM diisi di Script Properties.');
  CacheService.getScriptCache().remove('pinfail');
  console.log('Hitungan PIN salah direset.');
  SOURCES.forEach(src => {
    const info = sheetInfo_(src);
    if (!info) return console.log(`${src}: tab tidak ditemukan.`);
    console.log(`${src}: ${info.lastRow - 1} baris; kolom nama = "${info.nameIdx >= 0 ? info.head[info.nameIdx] : 'TIDAK TERDETEKSI'}"; kolom center = "${info.centerIdx >= 0 ? info.head[info.centerIdx] : '-'}"`);
  });
}
