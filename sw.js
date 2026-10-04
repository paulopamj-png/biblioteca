const V = 'bib-v11';
const OFFLINE = 'bib-offline';                       // downloads do usuario: nunca apagar ao atualizar o app
const SHELL = ['./', 'index.html', 'style.css', 'app.js', 'config.js', 'manifest.json', 'icon-192.png', 'icon-512.png', 'capa-padrao.svg', 'capa-audio.svg', 'fone-marca.svg'];
const CDN = ['https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.28.1/lib/msal-browser.min.js',
             'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
             'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
             'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
             'https://cdn.jsdelivr.net/npm/epubjs@0.3.93/dist/epub.min.js'];
const EXTERNOS = ['fonts.googleapis.com', 'fonts.gstatic.com', 'cdn.jsdelivr.net', 'cdnjs.cloudflare.com'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(async c => {
    await c.addAll(SHELL);
    await Promise.allSettled(CDN.map(u => fetch(u, { mode: 'no-cors' }).then(r => c.put(u, r))));   // bibliotecas para abrir sem internet
  }).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V && k !== OFFLINE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  if (EXTERNOS.includes(u.hostname)) {               // fontes e bibliotecas: guarda na primeira vez; depois funciona sem internet
    e.respondWith(caches.open(V).then(c => c.match(e.request).then(r => r || fetch(e.request).then(n => { c.put(e.request, n.clone()); return n; }))));
    return;
  }
  if (u.origin !== location.origin) return;           // Graph/OneDrive passam direto
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(r => {      // sempre confere com o servidor se há versão nova
    const copia = r.clone();
    caches.open(V).then(c => c.put(e.request, copia));
    return r;
  }).catch(() => caches.match(e.request)));
});
