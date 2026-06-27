// src/pwa.js — enregistrement du service worker (PWA).
// Activé uniquement dans le build de production (vite build / vite preview / hébergement)
// afin de ne pas interférer avec le Hot Module Replacement de Vite en développement.
// En dev (npm run dev), ce module ne fait rien.

if (import.meta.env.PROD && 'serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js')
      .then((reg) => { console.log('[PWA] service worker enregistré :', reg.scope); })
      .catch((err) => { console.warn('[PWA] échec d’enregistrement du service worker :', err); });
  });
}
