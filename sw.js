const V = 'bib-v4';
const SHELL = ['./', 'index.html', 'style.css', 'app.js', 'config.js', 'manifest.json', 'icon-192.png', 'icon-512.png', 'capa-padrao.svg', 'capa-audio.svg', 'fone-marca.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// Só o "casco" do app é guardado; Graph/OneDrive e CDNs passam direto pela rede.
const FONTES = ['fonts.googleapis.com', 'fonts.gstatic.com'];
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (FONTES.includes(u.hostname)) {            // fontes: guarda na primeira vez e depois funciona sem internet
    e.respondWith(caches.open(V).then(c => c.match(e.request).then(r => r || fetch(e.request).then(n => { c.put(e.request, n.clone()); return n; }))));
    return;
  }
  if (u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => {
    const copia = r.clone();
    caches.open(V).then(c => c.put(e.request, copia));
    return r;
  }).catch(() => caches.match(e.request)));
});
