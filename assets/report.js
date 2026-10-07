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

/* ---------- Periode & age group (Class Trial, SUNP) ---------- */
const MON = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
const pad2 = n => String(n).padStart(2, '0');
const isoDate = d => `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
const addDays = (s, n) => { const d = new Date(s + 'T12:00:00'); d.setDate(d.getDate() + n); return isoDate(d); };
const niceDate = s => { const [y, m, d] = s.split('-').map(Number); return `${d} ${MON[m-1]} ${y}`; };
const niceRange = (a, b) => a===b ? niceDate(a) : a.slice(0,7)===b.slice(0,7) ? `${Number(a.slice(8))}–${niceDate(b)}` : `${niceDate(a)} – ${niceDate(b)}`;
const dayLabel = k => { const d = new Date(k + 'T12:00:00'); return `${['Min','Sen','Sel','Rab','Kam','Jum','Sab'][d.getDay()]} ${d.getDate()} ${MON[d.getMonth()]}`; };
const monthLabel = k => `${MON[Number(k.slice(5,7))-1]} ${k.slice(0,4)}`;
// "2026-09-01" atau "9/1/2026" (format bulan/tanggal dari Sheets) → "2026-09-01"
function toIso(s){
  s = T(s); let m;
  if ((m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s))) return `${m[1]}-${m[2]}-${m[3]}`;
  if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s))) return `${m[3]}-${pad2(m[1])}-${pad2(m[2])}`;
  return null;
}
// Pilihan periode di atas filter center: periode cepat (dihitung ulang dari hari ini setiap dibuka) + tanggal dari/sampai.
// Pilihan disimpan di perangkat dengan kunci `key`; onChange dipanggil setiap periode berubah.
function period(key, onChange){
  const today = isoDate(new Date());
  const PRESETS = [
    ['today', 'Hari ini', () => [today, today]],
    ['yesterday', 'Kemarin', () => [addDays(today,-1), addDays(today,-1)]],
    ['week', 'Minggu ini', () => [addDays(today, -((new Date(today + 'T12:00:00').getDay() + 6) % 7)), today]],
    ['7d', '7 hari', () => [addDays(today,-6), today]],
    ['month', 'Bulan ini', () => [today.slice(0,8) + '01', today]],
    ['lastmonth', 'Bulan lalu', () => { const e = addDays(today.slice(0,8) + '01', -1); return [e.slice(0,8) + '01', e]; }],
    ['30d', '30 hari', () => [addDays(today,-29), today]]
  ];
  let range = lsGet(key) || {preset:'month'};
  const ok = s => /^\d{4}-\d{2}-\d{2}$/.test(s || '');
  function resolve(){
    const p = PRESETS.find(x => x[0]===range.preset);
    if (p){ const [a, b] = p[2](); range.from = a; range.to = b; }
    if (!ok(range.from) || !ok(range.to)){ range = {preset:'month'}; return resolve(); }
    if (range.from > range.to) [range.from, range.to] = [range.to, range.from];
  }
  resolve();
  const $ = s => document.querySelector(s);
  $('#cChips').insertAdjacentHTML('beforebegin', `<div class="dbar">
    <div class="cchips" id="pre" role="group" aria-label="Periode">${PRESETS.map(p => `<button type="button" class="cchip" data-p="${p[0]}">${p[1]}</button>`).join('')}</div>
    <div class="drange"><label>Dari<input type="date" id="dFrom"></label><span aria-hidden="true">–</span><label>Sampai<input type="date" id="dTo"></label></div>
  </div>`);
  function sync(){
    $('#dFrom').value = range.from; $('#dTo').value = range.to;
    document.querySelectorAll('#pre .cchip').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.p===range.preset)));
  }
  const set = r => { range = r; resolve(); lsSet(key, range); sync(); onChange(); };
  $('#pre').addEventListener('click', e => { const b = e.target.closest('[data-p]'); if (b) set({preset: b.dataset.p}); });
  ['#dFrom','#dTo'].forEach(id => $(id).addEventListener('change', () => {
    const a = $('#dFrom').value, b = $('#dTo').value; if (a && b) set({preset:'custom', from:a, to:b});
  }));
  sync();
  return {
    today,
    get from(){ return range.from; }, get to(){ return range.to; },
    label: () => niceRange(range.from, range.to),
    days: () => (new Date(range.to + 'T12:00:00') - new Date(range.from + 'T12:00:00')) / 864e5 + 1,
    // Batasi pilihan tanggal ke rentang data yang ada
    bounds(dates){ if (!dates.length) return; const s = dates.slice().sort(); $('#dFrom').min = $('#dTo').min = s[0]; $('#dFrom').max = $('#dTo').max = s[s.length-1]; }
  };
}
// Label umur di Exboard ("6-17 Mo", …) → nama kelas; label lain ditampilkan apa adanya
const AGE_NAMES = [[/^6\s*-\s*1[78]\s*mo/i, 'Baby'], [/^18\s*-\s*36\s*mo/i, 'Tiny'], [/^3\s*-\s*5\s*yo/i, 'Little'], [/^5\s*-\s*8/i, 'Kids'], [/^9\s*-\s*12\s*yo/i, 'Star']];
const ageName = s => { s = T(s); if (!s) return 'Lainnya'; const m = AGE_NAMES.find(([re, n]) => re.test(s) || s.toLowerCase()===n.toLowerCase()); return m ? m[1] : s; };
// Urut Baby → Star, lalu label lain dari umur termuda ("x Mo" = bulan, "x Yo" = tahun)
const ageMonths = s => { const m = /(\d+)\D*?(mo|yo|bulan|tahun|th)?/i.exec(s); if (!m) return 1e9; return Number(m[1]) * (/^(yo|tahun|th)/i.test(m[2] || s.replace(/^[^a-z]*/i,'')) ? 12 : 1); };
const AGE_ORDER = AGE_NAMES.map(x => x[1]);
const ageSort = (a, b) => (AGE_ORDER.indexOf(a) + 1 || 99) - (AGE_ORDER.indexOf(b) + 1 || 99) || ageMonths(a) - ageMonths(b) || a.localeCompare(b);
// Baris chip pilihan (age group, jam, …); pilihan disimpan dengan kunci `key`. Mengembalikan pilihan saat ini ('ALL' atau nilai).
function pickChips(el, items, key, opts, onChange){
  let cur = lsGet(key) || 'ALL';
  if (cur !== 'ALL' && !items.includes(cur)) cur = 'ALL';
  el.className = 'cbar cchips'; el.setAttribute('role', 'group'); el.setAttribute('aria-label', opts.aria);
  el.innerHTML = ['ALL'].concat(items).map(a => `<button type="button" class="cchip" data-v="${esc(a)}" aria-pressed="${a===cur}">${a==='ALL' ? esc(opts.all) : esc(a)}</button>`).join('');
  el.onclick = e => { const b = e.target.closest('[data-v]'); if (!b) return; lsSet(key, b.dataset.v); onChange(b.dataset.v); };
  return cur;
}
const ageChips = (el, ages, key, onChange) => pickChips(el, ages, key, {all:'Semua umur', aria:'Pilih age group'}, onChange);
// Jam sesi dari Sheets ("9:00:00", "09.00", "1:30:00 PM", "09.00 - 10.00") → "09:00"; teks lain apa adanya
function toTime(s){
  s = T(s); const m = /(\d{1,2})[:.](\d{2})(?::\d{2})?\s*(am|pm)?/i.exec(s);
  if (!m) return s || 'Lainnya';
  let h = Number(m[1]); const ap = (m[3] || '').toLowerCase();
  if (ap==='pm' && h < 12) h += 12; if (ap==='am' && h===12) h = 0;
  return `${pad2(h)}:${m[2]}`;
}
const timeSort = (a, b) => (/^\d\d:\d\d$/.test(b) - /^\d\d:\d\d$/.test(a)) || a.localeCompare(b);

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

window.DA.report = {T, num, fmt, fmtPct, fmtRp, sections, table, monthOf, kpi, bars, tbl, empty, start,
  MON, niceDate, niceRange, dayLabel, monthLabel, toIso, period, ageName, ageSort, ageChips, pickChips, toTime, timeSort};
})();
