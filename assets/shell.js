/* Kerangka bersama Dashboard: menu kategori & sub-menu, filter center yang diingat di semua halaman,
   dan utilitas baca Sheet feed. Dimuat sebelum skrip halaman; halaman memakai window.DA. */
(function(){
'use strict';
const ROOT = new URL('../', document.currentScript.src).href;
const CENTERS = ['KLM','KWC','PML','BTU','TMP','HIB'];
const ICON = {
  center: '<svg viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/></svg>',
  sa: '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3.2"/><path d="M3.5 19c.8-3.2 3-5 5.5-5s4.7 1.8 5.5 5"/><path d="M16 11l2 2 3.5-4"/></svg>',
  kelas: '<svg viewBox="0 0 24 24"><path d="M3 7h18M3 12h18M3 17h18"/><path d="M8 4v16M16 4v16"/></svg>'
};
// Kategori (menu bawah) dan sub-menu (menu atas). Sub-menu tanpa href = belum tersedia.
const NAV = [
  {id:'center', label:'Center', subs:[
    {id:'center', label:'Rekap Acquisition', href:'center/'},
    {id:'center-trial', label:'Class Trial', href:'center/trial/'},
    {id:'center-sunp', label:'SUNP', href:'center/sunp/'},
    {id:'center-retention', label:'Retention', href:'center/retention/'},
    {id:'center-cash', label:'Cash & Piutang', href:'center/cash/'}]},
  {id:'sa', label:'Student Advisor', short:'SA', subs:[
    {id:'sa', label:'Acquisition', href:'sa/'},
    {id:'sa-retention', label:'Retention', href:'sa/retention/'},
    {id:'sa-pip', label:'PIP', href:'sa/pip/'}]},
  {id:'kelas', label:'Kelas', subs:[
    {id:'kelas', label:'Slot Map', href:'kelas/'},
    {id:'kelas-utilisasi', label:'Utilisasi', href:'kelas/utilisasi/'},
    {id:'kelas-coach', label:'Coach', href:'kelas/coach/'},
    {id:'kelas-student', label:'Student', href:'https://script.google.com/a/macros/seven-retail.com/s/AKfycbzHc-PlOTKY5QIqG78-P9EgC9dVjeC2qNW-gdSApZy9XK0mV-I9FTgnMv8pv0LdgUUjGQ/exec'}]}
];

/* ---------- Utilitas ---------- */
const esc = s => String(s==null?'':s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function lsGet(k){ try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch(e){ return null; } }
function lsSet(k, v){ try { localStorage.setItem(k, JSON.stringify(v)); } catch(e){} }
function fmtTime(ts){
  if (!ts) return '';
  const d = new Date(ts);
  return d.toLocaleDateString('id-ID',{day:'numeric',month:'short'}) + ' ' + d.toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit'});
}
function parseCSV(text){
  const rows = []; let row = [], cell = '', q = false;
  for (let i=0; i<text.length; i++){
    const ch = text[i];
    if (q){
      if (ch==='"'){ if (text[i+1]==='"'){ cell += '"'; i++; } else q = false; }
      else cell += ch;
    } else if (ch==='"') q = true;
    else if (ch===','){ row.push(cell); cell = ''; }
    else if (ch==='\n'){ row.push(cell); rows.push(row); row = []; cell = ''; }
    else if (ch!=='\r') cell += ch;
  }
  if (cell!=='' || row.length){ row.push(cell); rows.push(row); }
  return rows;
}
// Angka dari CSV Sheets: "34.9", "34,9", "328553000", "328,553,000", "328.553.000", "-" (kosong)
function sheetNum(s){
  s = String(s==null?'':s).trim().replace(/\s|Rp|%/g,'');
  if (s==='' || s==='-' || s[0]==='#') return null;
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  if (/^-?\d+,\d{1,2}$/.test(s)) return Number(s.replace(',', '.'));
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return Number(s.replace(/,/g,''));
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) return Number(s.replace(/\./g,'').replace(',', '.'));
  const n = Number(s); return isFinite(n) ? n : null;
}
// Sheet yang dibagikan "siapa saja yang punya link" bisa dibaca langsung dari browser.
// export?format=csv membaca nilai apa adanya; gviz cadangan bila export diblokir.
async function fetchPublicCsv(id){
  const urls = [
    `https://docs.google.com/spreadsheets/d/${id}/export?format=csv`,
    `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv`
  ];
  let lastErr = null;
  for (const u of urls){
    try {
      const r = await fetch(u, {cache:'no-store', credentials:'omit'});
      if (!r.ok){ lastErr = {code:'http', status:r.status}; continue; }
      const t = await r.text();
      if (/^\s*<(!doctype|html)/i.test(t)){ lastErr = {code:'private'}; continue; }   // halaman login Google
      return t;
    } catch(e){ lastErr = {code: navigator.onLine === false ? 'offline' : 'blocked', message: e && e.message}; }
  }
  throw lastErr || {code:'blocked'};
}

/* ---------- Filter center (dipakai bersama semua halaman) ---------- */
const CENTER_KEY = 'dashboardApp.center';
const center = {
  get(){ const c = lsGet(CENTER_KEY); return c && (c==='ALL' || /^[A-Z]{2,5}$/.test(c)) ? c : 'ALL'; },
  set(c){ lsSet(CENTER_KEY, c); }
};
// Baris tombol Semua / KLM / KWC / …; onChange dipanggil dengan kode center atau 'ALL'
function centerChips(el, list, onChange, opts){
  opts = opts || {};
  const cur = center.get(), items = (opts.all===false ? [] : ['ALL']).concat(list || CENTERS);
  const sel = items.includes(cur) ? cur : items[0];
  el.classList.add('cchips'); el.setAttribute('role', 'group'); el.setAttribute('aria-label', 'Pilih center');
  el.innerHTML = items.map(c => `<button type="button" class="cchip" data-c="${esc(c)}" aria-pressed="${c===sel}">${c==='ALL' ? 'Semua' : esc(c)}</button>`).join('');
  if (!el._daBound){
    el._daBound = true;
    el.addEventListener('click', e => {
      const b = e.target.closest('button[data-c]'); if (!b) return;
      center.set(b.dataset.c);
      el.querySelectorAll('button[data-c]').forEach(x => x.setAttribute('aria-pressed', String(x===b)));
      el._daChange && el._daChange(b.dataset.c);
    });
  }
  el._daChange = onChange;
  return sel;
}

/* ---------- Font (uji coba) ---------- */
// Pilihan font diingat di perangkat ini. Default Poppins; Plus Jakarta Sans dimuat hanya bila dipilih.
const FONT_KEY = 'dashboardApp.font';
const FONTS = {
  poppins: {label:'Poppins'},
  jakarta: {label:'Plus Jakarta Sans', family:'"Plus Jakarta Sans"', css:'https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap'}
};
const font = {
  get(){ const f = lsGet(FONT_KEY); return FONTS[f] ? f : 'poppins'; },
  apply(k){
    const f = FONTS[k] || FONTS.poppins;
    if (f.css && !document.getElementById('da-font-css')){
      const l = document.createElement('link'); l.id = 'da-font-css'; l.rel = 'stylesheet'; l.href = f.css; document.head.appendChild(l);
    }
    let st = document.getElementById('da-font');
    if (!f.family){ if (st) st.remove(); return; }
    if (!st){ st = document.createElement('style'); st.id = 'da-font'; document.head.appendChild(st); }
    st.textContent = `body{font-family:${f.family},system-ui,-apple-system,"Segoe UI",Roboto,sans-serif !important}`;
  },
  set(k){ lsSet(FONT_KEY, k); font.apply(k); }
};
font.apply(font.get());

/* ---------- Navigasi ---------- */
function mountNav(pageId){
  const cat = NAV.find(c => c.subs.some(s => s.id===pageId)) || NAV[0];
  const top = document.querySelector('.top-inner');
  if (top){
    const nav = document.createElement('nav');
    nav.className = 'appnav'; nav.setAttribute('aria-label', 'Kategori');
    nav.innerHTML = NAV.map(c => `<a href="${ROOT + c.subs[0].href}"${c===cat ? ' aria-current="page"' : ''}>${ICON[c.id]}<span>${esc(c.short && innerWidth < 900 ? c.short : c.label)}</span></a>`).join('');
    top.appendChild(nav);
    document.body.classList.add('has-appnav');
  }
  const header = document.querySelector('header.top');
  if (header && cat.subs.length > 1){
    const sub = document.createElement('nav');
    sub.className = 'subnav'; sub.setAttribute('aria-label', cat.label);
    sub.innerHTML = `<div class="wrap">${cat.subs.map(s => s.href
      ? `<a href="${/^https?:/.test(s.href) ? s.href : ROOT + s.href}"${s.id===pageId ? ' aria-current="page"' : ''}>${esc(s.label)}</a>`
      : `<span class="soon" title="Segera hadir">${esc(s.label)} <small>segera</small></span>`).join('')}<button type="button" class="fontbtn" title="Ganti font (uji coba)"></button></div>`;
    header.insertAdjacentElement('afterend', sub);
    const fb = sub.querySelector('.fontbtn');
    const label = () => { fb.innerHTML = `<b>Aa</b> ${esc(FONTS[font.get()].label)}`; };
    fb.addEventListener('click', () => { font.set(font.get()==='jakarta' ? 'poppins' : 'jakarta'); label(); });
    label();
  }
  const me = cat.subs.find(s => s.id===pageId);
  if (me && me.href) { try { localStorage.setItem('dashboardApp.last', me.href); } catch(e){} }
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register(ROOT + 'sw.js').catch(() => {});
}

window.DA = {ROOT, CENTERS, NAV, esc, lsGet, lsSet, fmtTime, parseCSV, sheetNum, fetchPublicCsv, center, centerChips, mountNav, font};
})();
