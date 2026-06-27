// src/cloud.js — point d'entrée « cloud » (Phase 4).
// Charge l'authentification Entra ID + SharePoint UNIQUEMENT si la configuration est
// présente (.env rempli → isCloudConfigured = true). Sinon : mode local, rien n'est chargé
// (MSAL et le code Graph restent hors du bundle principal grâce à l'import dynamique).

import { isCloudConfigured } from './config.js';

// MSAL ouvre une iframe cachée (même origine) pour renouveler les jetons silencieusement.
// Cette iframe charge à nouveau l'app : il ne faut SURTOUT PAS y réinitialiser MSAL, sinon
// elle consomme la réponse d'auth avant la fenêtre parente → erreur "timed_out".
// On n'initialise donc le cloud que dans la fenêtre de premier plan.
const inIframe = window.self !== window.top;

if (isCloudConfigured && !inIframe) {
  import('./auth/auth-ui.js')
    .then((m) => m.initCloudUI())
    .catch((e) => console.error('[cloud] échec d’initialisation', e));
} else if (!isCloudConfigured) {
  console.info('[cloud] mode local — Entra ID / SharePoint non configuré (.env vide)');
}
