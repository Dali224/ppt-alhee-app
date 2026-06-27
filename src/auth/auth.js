// src/auth/auth.js — connexion / déconnexion / jetons via MSAL.js (Entra ID).
// Chargé dynamiquement par src/cloud.js UNIQUEMENT si la config cloud est présente,
// donc @azure/msal-browser reste hors du bundle en mode local.

import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import { msalConfig, loginRequest, graphTokenRequest } from './msal-config.js';

let pca = null;
let initialized = false;

export async function initMsal() {
  if (initialized) return pca;
  pca = new PublicClientApplication(msalConfig);
  await pca.initialize();
  // Récupère un éventuel retour de redirection (si on bascule un jour en loginRedirect)
  try {
    const res = await pca.handleRedirectPromise();
    if (res && res.account) pca.setActiveAccount(res.account);
  } catch (e) {
    console.warn('[auth] handleRedirectPromise', e);
  }
  if (!pca.getActiveAccount()) {
    const accts = pca.getAllAccounts();
    if (accts.length) pca.setActiveAccount(accts[0]);
  }
  initialized = true;
  return pca;
}

export function getActiveAccount() {
  return pca ? pca.getActiveAccount() : null;
}

export async function signIn() {
  await initMsal();
  // Flux par REDIRECTION (pas de popup) : la page navigue vers Microsoft puis revient
  // sur redirectUri, où handleRedirectPromise() (dans initMsal) traite la réponse.
  await pca.loginRedirect(loginRequest);
}

export async function signOut() {
  await initMsal();
  const account = pca.getActiveAccount();
  await pca.logoutRedirect({ account });
}

// Jeton d'accès Microsoft Graph : silencieux si possible, popup sinon.
export async function getGraphToken() {
  await initMsal();
  const account = pca.getActiveAccount();
  if (!account) throw new Error('Non connecté');
  try {
    const r = await pca.acquireTokenSilent({ ...graphTokenRequest, account });
    return r.accessToken;
  } catch (e) {
    // Silencieux impossible (consentement à confirmer, ou renouvellement iframe bloqué par
    // les cookies tiers → timed_out). On bascule en REDIRECTION pleine page : fiable, en
    // contexte first-party. La page navigue puis revient ; le jeton est alors en cache.
    console.warn('[auth] jeton silencieux impossible → redirection', e && e.errorCode);
    await pca.acquireTokenRedirect(graphTokenRequest);
    return await new Promise(() => {}); // la page va naviguer : on suspend ici
  }
}
