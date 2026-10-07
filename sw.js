// Service worker: halaman dashboard tetap bisa dibuka saat offline.
// Network-first untuk file aplikasi (supaya update langsung terpakai); data Sheet tidak dicache di sini,
// karena tiap halaman sudah menyimpan data terakhir di localStorage.
const CACHE = 'dashboard-app-v9';
const SHELL = ['./', 'index.html', 'center/', 'center/trial/', 'center/sunp/', 'center/retention/', 'center/cash/', 'sa/', 'sa/retention/', 'sa/pip/', 'kelas/', 'kelas/utilisasi/', 'kelas/coach/', 'assets/shell.css', 'assets/shell.js', 'assets/report.css', 'assets/report.js', 'manifest.webmanifest', 'icon.svg', 'icon-192.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).catch(() => {}));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(
    fetch(req).then(res => {
      if (res.ok){ const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req).then(r => r || caches.match('./')))
  );
});
