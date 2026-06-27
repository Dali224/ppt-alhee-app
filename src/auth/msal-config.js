// src/auth/msal-config.js — configuration MSAL.js (Entra ID).
// Toutes les valeurs viennent de src/config.js (lui-même alimenté par .env). Aucun secret en dur.

import { config } from '../config.js';

export const msalConfig = {
  auth: {
    clientId: config.entra.clientId,
    authority: config.entra.authority, // https://login.microsoftonline.com/<tenantId>
    redirectUri: config.entra.redirectUri,
    postLogoutRedirectUri: config.entra.redirectUri,
  },
  cache: {
    cacheLocation: 'localStorage', // session conservée entre rechargements (utile en visite mobile)
    storeAuthStateInCookie: false,
  },
  system: {
    // Le renouvellement silencieux par iframe est souvent bloqué (cookies tiers) → on échoue
    // vite pour basculer sur une redirection pleine page (fiable, contexte first-party).
    iframeHashTimeout: 6000,
    loadFrameTimeout: 6000,
  },
};

// On demande toutes les permissions dès la connexion (consentement déjà accordé) : le jeton
// est mis en cache d'avance et le 1er appel Graph les récupère sans iframe de renouvellement.
// Sites.Manage.All = créer listes/bibliothèques ; Sites.ReadWrite.All = écrire items/fichiers.
export const loginRequest = { scopes: config.entra.scopes };
export const graphTokenRequest = { scopes: ['Sites.ReadWrite.All', 'Sites.Manage.All'] };
