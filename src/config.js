// src/config.js — configuration Phase 4 (Entra ID + SharePoint via Microsoft Graph).
//
// Les valeurs proviennent des variables d'environnement Vite (VITE_*), définies dans un
// fichier `.env` à la racine (JAMAIS commité — voir .gitignore). Si la configuration est
// incomplète, `isCloudConfigured` vaut false et l'application fonctionne en MODE LOCAL
// (IndexedDB seul), exactement comme avant la Phase 4. Aucun secret n'est codé en dur ici.

const env = import.meta.env;
const clean = (v) => (typeof v === 'string' && v.trim() ? v.trim() : '');

export const config = {
  entra: {
    tenantId: clean(env.VITE_ENTRA_TENANT_ID),
    clientId: clean(env.VITE_ENTRA_CLIENT_ID),
    // redirect URI : par défaut l'origine courante (fonctionne en dev localhost et en prod)
    redirectUri:
      clean(env.VITE_ENTRA_REDIRECT_URI) ||
      (typeof window !== 'undefined' ? window.location.origin : ''),
    // autorité OIDC pour le tenant
    get authority() {
      return this.tenantId ? `https://login.microsoftonline.com/${this.tenantId}` : '';
    },
    // permissions Graph déléguées demandées au login.
    // Sites.Manage.All est nécessaire pour CRÉER des listes/bibliothèques/colonnes
    // (Sites.ReadWrite.All ne permet que d'écrire dans des listes existantes).
    scopes: ['User.Read', 'Sites.ReadWrite.All', 'Sites.Manage.All'],
  },
  sharepoint: {
    hostname: clean(env.VITE_SP_HOSTNAME), // ex. alhee.sharepoint.com
    sitePath: clean(env.VITE_SP_SITE_PATH), // ex. /sites/PPTALHEE
    indexList: clean(env.VITE_SP_INDEX_LIST) || 'PPT-Index',
    dataLibrary: clean(env.VITE_SP_DATA_LIBRARY) || 'PPT-Data',
  },
  graphBase: 'https://graph.microsoft.com/v1.0',
};

// true uniquement si tout le nécessaire est présent → bascule en mode SharePoint.
export const isCloudConfigured = !!(
  config.entra.tenantId &&
  config.entra.clientId &&
  config.sharepoint.hostname &&
  config.sharepoint.sitePath
);
