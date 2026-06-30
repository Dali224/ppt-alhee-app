// src/pwa.js — enregistrement du service worker + MISE À JOUR AUTOMATIQUE.
//
// En production : dès qu'une nouvelle version est déployée, le service worker se met à jour
// et l'application se recharge toute seule pour servir la dernière version. Les données en
// cours de saisie sont préservées (sauvegarde automatique dans IndexedDB).
// En développement (npm run dev) : ce module ne fait rien (n'interfère pas avec le HMR de Vite).

if (import.meta.env.PROD && 'serviceWorker' in navigator) {
  let refreshing = false;
  const hadController = !!navigator.serviceWorker.controller;

  // Quand un nouveau service worker prend le contrôle → on recharge pour passer à la nouvelle version.
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (refreshing) return;
    refreshing = true;
    if (hadController) window.location.reload(); // pas de reload inutile à la toute 1re installation
  });

  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').then((reg) => {
      console.log('[PWA] service worker enregistré :', reg.scope);
      // Cherche une mise à jour à l'ouverture, à chaque retour sur l'app, et toutes les heures.
      reg.update();
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') reg.update();
      });
      setInterval(() => reg.update(), 60 * 60 * 1000);
    }).catch((err) => {
      console.warn('[PWA] échec d’enregistrement du service worker :', err);
    });
  });
}
