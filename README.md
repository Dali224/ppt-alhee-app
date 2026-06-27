# ALHEE — Saisie PPT (Plan Pluriannuel de Travaux)

Application web de saisie de PPT — projet restructuré depuis l'outil HTML autonome
`Saisie_PPT_ALHEE_12_1.html`. **Phase 1 terminée.**

## Phase 1 — Cartographie + restructuration locale ✓

L'outil HTML monolithique a été décomposé en projet propre :
- **moteur PPTX gelé** (JSZip + générateur copiés caractère pour caractère)
- **modèle .pptx extrait** en asset binaire (chargé via `fetch` au démarrage)
- **CSS séparé**, **markup HTML minimal**, **couche applicative en module ES**
- IndexedDB conservée à l'identique (même nom de base, même store, même clés) → les
  données saisies dans l'ancien outil restent compatibles
- Test de non-régression automatisé (régénération + diff binaire)

## Démarrer

```bash
npm install
npm run extract              # extrait les assets depuis le HTML source (idempotent)
npm run dev                  # démarre Vite sur http://localhost:5173
```

> Le HTML source est cherché par défaut à
> `C:/Users/Dali/Downloads/Saisie_PPT_ALHEE_12_1.html`. Sur un autre poste,
> définir la variable `SOURCE` : `SOURCE=/chemin/source.html npm run extract`.

## Tests de non-régression (CRITIQUE — à lancer avant et après toute modif)

L'invariant : pour un même `tests/reference-project.json`, l'export PPTX doit produire
un fichier **strictement identique** à `tests/reference-output.pptx` (l'oracle figé).

```bash
npm run test:no-regression   # régénère un candidat + diff binaire vs l'oracle
```

Workflow d'évolution :
1. **Avant** toute modif (visuelle, mobile, SharePoint…) → vérifier que
   `npm run test:no-regression` passe.
2. Faire la modif.
3. **Après** la modif → relancer `npm run test:no-regression`. Doit passer.
4. Si la modif touche **légitimement** le moteur ou le catalogue (rare en Phases 2-3) :
   `npm run generate-reference` pour mettre à jour l'oracle, puis commiter.

## Structure

```
public/
├── modele-alhee.pptx              Modèle PPT extrait (binaire, fetch au démarrage)
├── manifest.webmanifest           PWA — manifeste (installable)
├── sw.js                          PWA — service worker (cache app-shell + pptx, offline)
├── icon.svg / icon-192.png / icon-512.png / apple-touch-icon-180.png   Icônes de marque
└── vendor/
    ├── jszip.min.js               JSZip vendored verbatim (window.JSZip)
    └── pptx-engine.js             🔒 Moteur PPTX GELÉ — copie verbatim (window.GEN)

src/
├── styles.css                     CSS (palette ALHEE, tokens, responsive mobile)
├── catalog.js                     BDD ALHEE + obligations/contrats/docs + smileys
├── app.js                         Couche applicative (UI, IA, photo, IndexedDB, exports)
└── pwa.js                         Enregistrement du service worker (prod uniquement)

tests/
├── reference-project.json         Jeu de données de référence (toutes sections)
└── reference-output.pptx          Oracle figé — l'invariant à préserver

scripts/
├── extract-assets.mjs             Extrait tout depuis le HTML source (reproductible)
├── make-icons.mjs                 Génère les icônes PNG de la PWA (via zlib, sans dépendance)
├── generate-reference.mjs         Génère le PPTX de référence (Node, moteur gelé)
├── diff-pptx.mjs                  Compare 2 PPTX (zip + diff XML/binaire)
└── test-no-regression.mjs         Pipeline : génère candidat → diff vs oracle

index.html                         Markup body + PWA + chargement des scripts
package.json                       vite + jszip (devDependency)
vite.config.js
.gitignore
```

## Mobile / PWA (Phase 3)

- **Responsive** : sur petit écran, les grilles s'empilent en 1 colonne, les zones tactiles
  sont agrandies (≥ 46 px), les champs passent à 16 px (anti-zoom iOS), l'action principale
  « Exporter PPT » est pleine largeur. Mêmes écrans, juste réorganisés.
- **Photo** : compression à la capture (max 1600 px, JPEG ~82 %, orientation EXIF gérée),
  stockée en dataURL dans `it.photo` — format inchangé pour le moteur PPTX.
- **PWA** : installable (manifeste + icônes), service worker activé **en production
  uniquement** (pas en `npm run dev`, pour ne pas gêner le HMR de Vite).

Tester la PWA (le SW ne s'active qu'en build de prod) :

```bash
npm run build
npm run preview          # sert dist/ ; ouvrir l'URL, le SW s'enregistre et met en cache
```

Le service worker pré-cache l'app-shell **et le modèle `.pptx`** dès l'installation : après
une première ouverture en ligne, la saisie, la capture photo et la génération PowerPoint
fonctionnent **hors-ligne**. Les données restent en IndexedDB local (pas de synchro réseau
avant la Phase 4).

## Contrats à NE PAS rompre

- **`public/vendor/pptx-engine.js`** est la boîte noire gelée du moteur. Aucune
  modification, même cosmétique.
- **`public/vendor/jszip.min.js`** : idem, vendored verbatim.
- **Structure de `project` runtime** : `meta / car / lots / oblig / docs / synthGen /
  synthese / conclusion / zone / tva`. Tout changement de schéma casse la
  compatibilité avec les sauvegardes IndexedDB existantes des utilisateurs.
- **Onglets** (ordre, libellés, ids) : conservés tels quels.
- **IndexedDB** : base `alhee_ppt`, store `kv`, clé `project`, versions préfixées `ver:`.

## Décisions enregistrées (Phase 1)

- **IA Anthropic** : conservée telle quelle (clé API utilisateur côté navigateur,
  appel direct `api.anthropic.com`). À ré-arbitrer à la Phase 4 (routage via backend
  Azure Function + Entra ID si gouvernance souhaitée).
- **JSZip** : vendored verbatim depuis le HTML source (pas la version npm) pour
  garantir l'invariance du PPTX généré.
- **Hébergement Phase 1** : local uniquement (Vite dev / build statique). Le déploiement
  Azure Static Web Apps arrive en Phase 5.

## Décisions enregistrées (Phase 3)

- **Mode photo** : caméra **+ galerie au choix** (`accept="image/*"` sans `capture`) —
  choix de l'utilisateur, qui prime sur le `capture="environment"` du cahier des charges §6,
  pour permettre aussi la réutilisation de photos prises hors de l'app.
- **Compression photo** : max 1600 px, JPEG 0,82 (constantes `PHOTO_MAX_DIM` / `PHOTO_QUALITY`
  en tête de `src/app.js`, facilement ajustables).
- **PWA hors-ligne sans synchro** : cache local complet, mais aucune synchro réseau — il n'y
  a rien à synchroniser tant que SharePoint n'est pas branché (Phase 4 ; réconciliation
  robuste en Phase 6).

## Roadmap

- [x] **Phase 1** — Cartographie + restructuration locale
- [x] **Phase 2** — Refonte visuelle (Montserrat + Manrope, tokens, états ; structure inchangée, export PPTX identique)
- [x] **Phase 3** — Mobile (responsive + capture photo compressée + PWA installable & hors-ligne)
- [ ] **Phase 4** — Backend Entra ID + SharePoint via Microsoft Graph (multi-utilisateur)
- [ ] **Phase 5** — Hébergement Azure Static Web Apps
- [ ] **Phase 6** (optionnelle) — Synchronisation hors-ligne robuste

À chaque phase touchant le visuel ou le pipeline d'export :
**lancer `npm run test:no-regression` d'abord**, puis le re-lancer après.
