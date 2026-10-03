/* Bersama untuk halaman laporan (sub-menu): baca "Feed Dashboard Tambahan", pecah per tab Exboard,
   cari tabel berdasarkan judulnya, dan komponen tampilan (angka, batang, tabel). Butuh shell.js (window.DA). */
(function(){
'use strict';
const {esc, lsGet, lsSet, fmtTime, parseCSV, sheetNum, fetchPublicCsv} = DA;

// Satu Sheet berisi beberapa tab Exboard berdampingan; baris 1 berisi penanda "##<nama tab>" di kolom awalnya.
const EXTRA = {id:'1378IAR8Wan0nTg9diXcok6_Q8c1rqkDakXbpS3JPvKc', name:'Feed Dashboard Tambahan (dari Exboard)', cache:'dashboardApp.extra.v1'};
const STALE_MS = 5 * 60 * 1000;

/* ---------- Angka & format ---------- */
const T = s => String(s==null?'':s).replace(/\s+/g,' ').trim();
// "(424,148,337)" = negatif (format akuntansi Sheets)
function num(s){
  s = T(s); if (!s) return null;
  const neg = /^\(.*\)$/.test(s);
  const n = sheetNum(neg ? s.slice(1,-1) : s);
  return n==null ? null : neg ? -n : n;
}
const fmt = (v, d=1) => (v==null || isNaN(v)) ? '—' : Number(v).toLocaleString('id-ID',{maximumFractionDigits:d});
const fmtPct = (v, d=1) => (v==null || isNaN(v)) ? '—' : fmt(v,d) + '%';
function fmtRp(v){
  if (v==null || isNaN(v)) return '—';
  const a = Math.abs(v), s = v < 0 ? '−' : '';
  if (a >= 1e9) return s + 'Rp ' + fmt(a/1e9, 2) + ' M';
  if (a >= 1e6) return s + 'Rp ' + fmt(a/1e6, 1) + ' jt';
  if (a >= 1e3) return s + 'Rp ' + fmt(a/1e3, 0) + ' rb';
  return s + 'Rp ' + fmt(a, 0);
}

/* ---------- Pecah feed per tab & cari tabel ---------- */
function sections(text){
  const rows = parseCSV(text), top = rows[0] || [], marks = [];
  top.forEach((c, i) => { if (/^##/.test(T(c))) marks.push({name: T(c).slice(2).trim(), col: i}); });
  const out = {};
  marks.forEach((m, k) => {
    const end = k+1 < marks.length ? marks[k+1].col : undefined;
    out[m.name] = rows.slice(1).map(r => r.slice(m.col, end).map(T));
  });
  return out;
}
const isTitle = s => /^[A-Z](\.\d+)?\.\s+\S/.test(s);
// Tabel = baris judul (mis. "A. COHORT BULAN BERJALAN …"), baris header tepat di bawahnya, lalu baris data
// sampai judul berikutnya. Kolom dicari lewat nama header, mulai dari kolom judul.
function table(sec, re){
  if (!sec) return null;
  for (let r = 0; r < sec.length; r++) for (let c = 0; c < sec[r].length; c++){
    if (!re.test(sec[r][c] || '')) continue;
    const head = sec[r+1] || [], rows = [];
    let last = c; head.forEach((h, j) => { if (j >= c && h) last = j; });
    // Baris kosong dilewati (mis. TOTAL retention ada di bawah beberapa baris kosong); selesai di judul berikutnya
    // atau setelah 12 baris kosong berturut-turut. Baris "No data" berisi "-" tetap ikut.
    for (let i = r+2, gap = 0; i < sec.length && gap < 12; i++){
      const row = sec[i] || [];
      if (!row.slice(c, last+1).some(x => x)){ gap++; continue; }
      if (isTitle(row[c] || '')) break;
      gap = 0; rows.push(row);
    }
    const t = {col:c, head, rows, title: sec[r][c]};
    t.idx = name => { const n = T(name).toLowerCase(); for (let j = c; j < head.length; j++) if (T(head[j]).toLowerCase()===n) return j; return -1; };
    t.get = (row, name) => { const j = t.idx(name); return j < 0 ? '' : row[j] || ''; };
    // label: teks persis (tanpa beda huruf besar/kecil) atau RegExp
    t.row = label => t.rows.find(r => label instanceof RegExp ? label.test(r[c] || '') : T(r[c]).toLowerCase()===T(label).toLowerCase()) || null;
    // Tabel "Metric, , KLM, KWC, … TOTAL": nilai metrik untuk satu center (ALL = TOTAL)
    t.val = (metric, center) => { const r = t.row(metric); return r ? num(t.get(r, center==='ALL' ? 'TOTAL' : center)) : null; };
    return t;
  }
  return null;
}
const monthOf = sec => { for (const r of (sec || []).slice(0, 6)) { const i = r.findIndex(x => /month:$/i.test(x)); if (i >= 0) return r[i+1] || ''; } return ''; };

/* ---------- Komponen ---------- */
const kpi = (label, val, foot, tone) => `<div class="kpi"><div class="k-label">${esc(label)}</div><div class="k-val ${tone||''}">${val}</div>${foot ? `<div class="k-foot">${foot}</div>` : ''}</div>`;
// items: [{label, value, text, tone, sub}]; max opsional (default nilai terbesar), mark = garis pembanding (nilai)
function bars(items, opts){
  opts = opts || {};
  const max = opts.max || Math.max(1, ...items.map(x => Math.abs(x.value || 0)));
  if (!items.length) return '<p class="note">Belum ada data.</p>';
  return `<div class="bars">${items.map(x => {
    const w = x.value==null ? 0 : Math.max(0, Math.min(100, Math.abs(x.value)/max*100));
    const mk = opts.mark!=null ? `<b style="left:calc(${Math.min(100, opts.mark/max*100)}% - 1px)"></b>` : '';
    return `<div class="brow"><span class="bl" title="${esc(x.label)}">${esc(x.label)}</span><div class="track"><i class="${x.tone||''}" style="width:${w}%"></i>${mk}</div><span class="bv">${x.text!=null ? x.text : fmt(x.value,0)}${x.sub ? `<small>${x.sub}</small>` : ''}</span></div>`;
  }).join('')}</div>`;
}
function tbl(head, rows, opts){
  opts = opts || {};
  return `<div class="table-wrap"><table class="tbl"><thead><tr>${head.map((h, i) => `<th class="${i===0?'l':''}">${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r =>
    `<tr class="${r.cls||''}">${r.cells.map((c, i) => `<td class="${i===0?'l ':''}${(r.tones||[])[i]||''}">${c}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
const empty = (title, text) => `<div class="empty-state"><h2>${esc(title)}</h2><p>${text}</p></div>`;

/* ---------- Muat feed ---------- */
// page: {render(secs), foot?, feed?, parse?} — render dipanggil dengan hasil parse feed. Mengelola status, error, cache.
// Default: Feed Dashboard Tambahan dipecah per tab. Halaman lain boleh memberi feed {id,name,cache} dan
// parse(text) sendiri yang mengembalikan data, atau 'ref' / 'empty' / pesan error (string berawalan '!').
function start(page){
  const FEED = page.feed || EXTRA;
  let data = null, pulling = false;
  const $ = s => document.querySelector(s);
  const setStatus = (k, t) => { const el = $('#status'); el.className = 'sync ' + k; el.textContent = t; el.title = t; };
  const showErr = h => { const b = $('#errBox'); b.innerHTML = h || ''; b.hidden = !h; };
  const url = `https://docs.google.com/spreadsheets/d/${FEED.id}/edit`;
  function draw(){
    if (!data){ if (!pulling) $('#view').innerHTML = empty('Belum ada data', 'Data belum pernah berhasil ditarik di perangkat ini. Periksa pesan di atas, lalu ketuk tombol tarik ulang.'); return; }
    try { $('#view').innerHTML = page.render(data.secs); }
    catch(e){ $('#view').innerHTML = empty('Tabel belum terbaca', 'Susunan tab di Exboard mungkin berubah. ' + esc(e && e.message || '')); }
    $('#foot').innerHTML = `<span>Ditarik ${esc(fmtTime(data.pulledAt))}. Sumber: <a href="${url}" target="_blank" rel="noopener">${esc(FEED.name)}</a>.</span>${page.foot ? `<span>${page.foot}</span>` : ''}`;
  }
  // 'ok' | 'ref' (IMPORTRANGE belum diizinkan) | 'empty'
  function parseSections(text){
    const secs = sections(text), names = Object.keys(secs);
    if (!names.length) return '!<b>Isi feed tidak dikenali.</b> Pastikan baris 1 feed masih berisi penanda ##nama tab.';
    if (names.every(n => ((secs[n][0] || [])[0] || '').startsWith('#'))) return 'ref';
    return secs;
  }
  function load(text, pulledAt){
    const r = (page.parse || parseSections)(text);
    if (typeof r==='string') return r;
    data = {secs: r, pulledAt}; return 'ok';
  }
  async function pull(){
    if (pulling) return;
    pulling = true; $('#pullBtn').disabled = true; $('#pullBtn').classList.add('spin'); setStatus('busy', 'Menarik data…');
    try {
      const text = await fetchPublicCsv(FEED.id), now = Date.now(), st = load(text, now);
      if (st==='ok'){ lsSet(FEED.cache, {text, pulledAt: now}); setStatus('ok', 'Tersinkron · ' + fmtTime(now)); showErr(''); }
      else {
        setStatus('err', 'Feed belum siap');
        showErr(st==='ref' ? `<b>Feed belum bisa membaca Exboard.</b> Buka <a href="${url}" target="_blank" rel="noopener">${esc(FEED.name)}</a>, klik sel yang berisi #REF!, lalu pilih "Allow access".`
          : st[0]==='!' ? st.slice(1) : '<b>Isi feed tidak dikenali.</b>');
      }
    } catch(err){
      setStatus('err', data ? 'Data ' + fmtTime(data.pulledAt) : 'Gagal menarik');
      const why = err && err.code==='offline' ? 'HP sedang offline.'
        : `<a href="${url}" target="_blank" rel="noopener">${esc(FEED.name)}</a> belum bisa dibaca dari aplikasi ini. Buka Sheet-nya → Bagikan → Akses umum → <b>Siapa saja yang memiliki link</b> (Viewer).`;
      showErr(`<b>Data terbaru belum bisa ditarik.</b> ${why}${data ? ` Menampilkan data terakhir (${esc(fmtTime(data.pulledAt))}).` : ''}`);
    } finally {
      pulling = false; $('#pullBtn').disabled = false; $('#pullBtn').classList.remove('spin'); draw();
    }
  }
  $('#pullBtn').addEventListener('click', pull);
  document.addEventListener('visibilitychange', () => { if (document.visibilityState==='visible' && (!data || Date.now() - data.pulledAt > STALE_MS)) pull(); });
  const c = lsGet(FEED.cache);
  if (c && typeof c.text==='string' && load(c.text, c.pulledAt)==='ok') setStatus('', 'Data ' + fmtTime(c.pulledAt));
  draw(); pull();
  return {redraw: draw};
}

window.DA.report = {T, num, fmt, fmtPct, fmtRp, sections, table, monthOf, kpi, bars, tbl, empty, start};
})();
