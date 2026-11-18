// La app abre al instante desde la caché del móvil y se actualiza por detrás.
// Los datos siempre se piden a la hoja (la API nunca se cachea).
const CACHE = 'incidencias-v17';
const ARCHIVOS = ['./', './index.html', './config.js', './manifest.webmanifest', './icon-192.png', './icon-512.png',
  './icon-maskable-512.png', './apple-touch-icon.png', './fotos/equipo.jpg', './fotos/equipo-movil.jpg',
  './logos/fo-blanco.png', './logos/fo-color.png', './logos/escudo-color.png', './logos/escudo-blanco.png', './lib/html2canvas.min.js', './lib/jspdf.umd.min.js'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  const esPagina = e.request.mode === 'navigate';
  const clave = esPagina ? './index.html' : e.request;
  e.respondWith(caches.open(CACHE).then(async c => {
    const enCache = await c.match(clave, { ignoreSearch: true });
    const deRed = fetch(e.request, { cache: 'no-cache' }).then(async r => {
      if (r.ok) {
        if (esPagina && enCache) {
          const [a, b] = await Promise.all([enCache.clone().text(), r.clone().text()]);
          if (a !== b) (await self.clients.matchAll()).forEach(cl => cl.postMessage('nueva-version'));
        }
        await c.put(clave, r.clone());
      }
      return r;
    }).catch(() => enCache || c.match('./index.html'));
    if (enCache) { e.waitUntil(deRed); return enCache; }    // instantáneo; se actualiza por detrás
    return deRed;
  }));
});
