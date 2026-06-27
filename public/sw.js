/* ALHEE — Saisie PPT · Service Worker (Phase 3 — PWA)
   Objectif : application installable et utilisable hors-ligne en visite.
   Stratégies :
     - navigations (HTML)        : network-first, repli sur le cache (puis /index.html)
     - assets same-origin        : stale-while-revalidate (rapide + mise à jour en fond)
     - cross-origin (Google Fonts) : cache-first opportuniste
   Le modèle PPTX (/modele-alhee.pptx) est mis en cache au 1er chargement en ligne,
   ce qui permet la génération PowerPoint hors-ligne ensuite.

   NB : pas de synchronisation réseau ici — les données restent en IndexedDB local
   (comme aujourd'hui). La synchro SharePoint viendra avec la Phase 4. */

const CACHE = 'alhee-ppt-v4';
const APP_SHELL = [
  '/',
  '/index.html',
  '/manifest.webmanifest',
  '/icon.svg',
  '/icon-192.png',
  '/icon-512.png',
  '/apple-touch-icon-180.png',
  '/modele-alhee.pptx',   // pré-caché dès l'installation → génération PowerPoint hors-ligne
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(APP_SHELL).catch(() => {}))   // tolérant si un asset manque
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('message', (e) => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Navigations : network-first
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req)
        .then((r) => { putInCache(req, r.clone()); return r; })
        .catch(() => caches.match(req).then((m) => m || caches.match('/index.html')))
    );
    return;
  }

  // Same-origin : stale-while-revalidate
  if (url.origin === self.location.origin) {
    e.respondWith(
      caches.match(req).then((cached) => {
        const net = fetch(req)
          .then((r) => { if (r && r.status === 200) putInCache(req, r.clone()); return r; })
          .catch(() => cached);
        return cached || net;
      })
    );
    return;
  }

  // Cross-origin (ex. Google Fonts) : cache-first opportuniste
  e.respondWith(
    caches.match(req).then((cached) =>
      cached || fetch(req).then((r) => {
        if (r && (r.status === 200 || r.type === 'opaque')) putInCache(req, r.clone());
        return r;
      }).catch(() => cached)
    )
  );
});

function putInCache(req, res) {
  caches.open(CACHE).then((c) => c.put(req, res)).catch(() => {});
}
