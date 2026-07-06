// scripts/extract-assets.mjs
// Extrait depuis le HTML autonome `Saisie_PPT_ALHEE_12_1.html` :
//   - public/modele-alhee.pptx          (binaire décodé depuis TEMPLATE_B64)
//   - src/vendor/jszip.min.js           (texte verbatim du <script>/* JSZip */)
//   - src/pptx-engine.js                (texte verbatim du <script> moteur)
//   - src/catalog.js                    (BDD + DEFAULT_OBLIG + DEFAULT_DOCS + SMILEY + SMILEY_FULL, exports ES)
//
// Usage :
//   SOURCE=/chemin/vers/Saisie_PPT_ALHEE_12_1.html node scripts/extract-assets.mjs
//   (défaut : C:/Users/Dali/Downloads/Saisie_PPT_ALHEE_12_1.html)

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const DEFAULT_SOURCE = 'C:/Users/Dali/Downloads/Saisie_PPT_ALHEE_12_1.html';
const SOURCE = process.env.SOURCE || DEFAULT_SOURCE;

if (!fs.existsSync(SOURCE)) {
  console.error(`✗ Source HTML introuvable : ${SOURCE}`);
  console.error('  Définis SOURCE=... dans l\'environnement ou modifie DEFAULT_SOURCE.');
  process.exit(1);
}

console.log(`→ Lecture de ${SOURCE}`);
const html = fs.readFileSync(SOURCE, 'utf8');
const lines = html.split('\n');
console.log(`  ${lines.length} lignes lues.`);

function findLine(predicate, from = 0, direction = 1) {
  if (direction > 0) {
    for (let i = from; i < lines.length; i++) if (predicate(lines[i], i)) return i;
  } else {
    for (let i = from; i >= 0; i--) if (predicate(lines[i], i)) return i;
  }
  return -1;
}

// 1) JSZip --------------------------------------------------------------------
const jszipMarker = findLine(l => l.trim() === '<script>/* JSZip */');
if (jszipMarker < 0) throw new Error('Bloc JSZip introuvable (<script>/* JSZip */)');
const jszipEnd = findLine(l => l.trim() === '</script>', jszipMarker + 1);
if (jszipEnd < 0) throw new Error('Fermeture </script> du bloc JSZip introuvable');
const jszipCode = lines.slice(jszipMarker + 1, jszipEnd).join('\n');

// 2) Moteur PPTX --------------------------------------------------------------
const engineMarker = findLine(l => l.includes('/* Moteur de génération PPTX ALHEE */'));
if (engineMarker < 0) throw new Error('Bloc moteur PPTX introuvable');
const engineOpen = findLine(l => l.trim() === '<script>', engineMarker - 1, -1);
if (engineOpen < 0) throw new Error('Ouverture <script> du moteur introuvable');
const engineClose = findLine(l => l.trim() === '</script>', engineMarker + 1);
if (engineClose < 0) throw new Error('Fermeture </script> du moteur introuvable');
const engineCode = lines.slice(engineOpen + 1, engineClose).join('\n');

// 3) Extraire la déclaration complète d'une constante (parseur équilibré)
//    Gère le cas où plusieurs constantes sont sur la même ligne (ex : L1195 SMILEY + SMILEY_FULL).
function extractDecl(name) {
  const re = new RegExp(`(?:^|[\\s;])(const|let|var)\\s+${name}\\s*=`, 'g');
  let m;
  while ((m = re.exec(html))) {
    // Position du début de `const|let|var` (en sautant le séparateur initial s'il est de longueur 1)
    const sepLen = m[0].length - (`${m[1]} ${name}=`.length); // approx; mais plus simple :
    const declKwIdx = html.indexOf(m[1], m.index);
    const eqIdx = html.indexOf('=', declKwIdx);
    // Parse depuis eqIdx + 1 jusqu'au `;` au niveau 0
    let i = eqIdx + 1;
    let depth = 0;
    let inStr = null;
    let escaped = false;
    while (i < html.length) {
      const c = html[i];
      if (inStr) {
        if (escaped) { escaped = false; }
        else if (c === '\\') { escaped = true; }
        else if (c === inStr) { inStr = null; }
      } else {
        if (c === '"' || c === '\'' || c === '`') inStr = c;
        else if (c === '{' || c === '[' || c === '(') depth++;
        else if (c === '}' || c === ']' || c === ')') depth--;
        else if (c === ';' && depth === 0) {
          // Localisation pour message diagnostic (numéro de ligne)
          const lineNo = html.slice(0, declKwIdx).split('\n').length;
          return { decl: html.slice(declKwIdx, i + 1), lineNo };
        }
      }
      i++;
    }
    throw new Error(`Fin de déclaration ${name} introuvable (pas de ';' au niveau 0)`);
  }
  throw new Error(`Déclaration ${name} introuvable`);
}

const bdd = extractDecl('BDD');
const dOblig = extractDecl('DEFAULT_OBLIG');
const dDocs = extractDecl('DEFAULT_DOCS');
const smileyB64 = extractDecl('SMILEY_B64');
const smileyFullB64 = extractDecl('SMILEY_FULL_B64');
const smiley = extractDecl('SMILEY');
const smileyFull = extractDecl('SMILEY_FULL');
const templateDecl = extractDecl('TEMPLATE_B64');

// 4) Décoder TEMPLATE_B64 → binaire .pptx
const tplMatch = templateDecl.decl.match(/TEMPLATE_B64\s*=\s*"([A-Za-z0-9+/=]+)"/);
if (!tplMatch) throw new Error('Impossible d\'extraire la chaîne base64 depuis TEMPLATE_B64');
const templateBytes = Buffer.from(tplMatch[1], 'base64');

// 5) Réécrire chaque déclaration en `export const NAME = ...;`
function toExport(name, decl) {
  const re = new RegExp(`^(const|let|var)\\s+${name}\\b`);
  if (!re.test(decl)) {
    throw new Error(`Impossible de réécrire la déclaration ${name} en export ES`);
  }
  return decl.replace(re, `export const ${name}`);
}

// SMILEY/SMILEY_FULL réfèrent SMILEY_B64/SMILEY_FULL_B64 et `_b64ToU8`. On les laisse
// internes (pas d'export) pour ne pas polluer le module.
function toLocal(name, decl) {
  // les déclarations internes restent `const NAME = ...;` sans export
  return decl;
}

const catalogPieces = [
  '// src/catalog.js',
  '// Catalogue ALHEE — extrait verbatim depuis Saisie_PPT_ALHEE_12_1.html',
  '// NE PAS MODIFIER À LA MAIN. Régénérer via `npm run extract`.',
  '',
  '// helper interne pour décoder un base64 en Uint8Array (utilisé par SMILEY / SMILEY_FULL)',
  'function _b64ToU8(b){const s=atob(b);const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u;}',
  '',
  toExport('BDD', bdd.decl),
  toExport('DEFAULT_OBLIG', dOblig.decl),
  toExport('DEFAULT_DOCS', dDocs.decl),
  '',
  '// Smileys PNG (données b64 + objets décodés Uint8Array consommés par le moteur PPTX)',
  toLocal('SMILEY_B64', smileyB64.decl),
  toLocal('SMILEY_FULL_B64', smileyFullB64.decl),
  toExport('SMILEY', smiley.decl),
  toExport('SMILEY_FULL', smileyFull.decl),
  ''
];
const catalogCode = catalogPieces.join('\n');

// 6) Extraire le 3e <script> (couche applicative) → src/app.js
//    Stratégie : COPIE VERBATIM du code, avec uniquement :
//      a) imports ES en tête pour BDD, DEFAULT_OBLIG, DEFAULT_DOCS, SMILEY, SMILEY_FULL
//      b) lignes des constantes extraites commentées
//      c) ligne de templateBytes remplacée par un fetch du .pptx
const appMarker = findLine(l => l.includes('CATALOGUE (BDD ALHEE'));
if (appMarker < 0) throw new Error('Bloc applicatif (CATALOGUE) introuvable');
const appOpen = findLine(l => l.trim() === '<script>', appMarker - 1, -1);
if (appOpen < 0) throw new Error('Ouverture <script> du bloc applicatif introuvable');
const appClose = findLine(l => l.trim() === '</script>', appMarker + 1);
if (appClose < 0) throw new Error('Fermeture </script> du bloc applicatif introuvable');

// Numéros de ligne (1-based) des déclarations à retirer/remplacer
const linesToBlank = new Set([
  bdd.lineNo,
  dOblig.lineNo,
  dDocs.lineNo,
  smileyB64.lineNo,
  smileyFullB64.lineNo,
  smiley.lineNo,
  smileyFull.lineNo,
  templateDecl.lineNo
]);

const appBody = [];
for (let i = appOpen + 1; i < appClose; i++) {
  const lineNo = i + 1; // 1-based
  if (linesToBlank.has(lineNo)) {
    // Commenter intégralement (l'utilisateur peut retrouver le contenu via le HTML source)
    appBody.push(`// (L${lineNo} extrait — voir catalog.js ou public/modele-alhee.pptx)`);
    continue;
  }
  // Remplacement spécial : init de templateBytes basée sur TEMPLATE_B64 → fetch
  if (lines[i].includes('templateBytes=') && lines[i].includes('TEMPLATE_B64')) {
    appBody.push('// Init templateBytes par fetch du .pptx (au lieu de _b64ToU8(TEMPLATE_B64))');
    appBody.push('let templateBytes = null;');
    appBody.push("fetch('/modele-alhee.pptx').then(r=>{if(!r.ok)throw new Error('HTTP '+r.status);return r.arrayBuffer();}).then(b=>{templateBytes=new Uint8Array(b);const ts=document.getElementById('tplStatus');if(ts&&!ts.textContent.includes('\\u2713')){ts.textContent='Modèle ALHEE intégré \\u2713';ts.style.color='var(--vf)';}}).catch(e=>{console.error('Modèle PPT injoignable :',e);const ts=document.getElementById('tplStatus');if(ts){ts.textContent='⚠ Modèle PPT injoignable';ts.style.color='var(--rouge)';}});");
    continue;
  }
  appBody.push(lines[i]);
}

const appCode = [
  '// src/app.js',
  '// Couche applicative ALHEE — copie verbatim du 3e <script> de Saisie_PPT_ALHEE_12_1.html',
  '// (lignes ' + (appOpen + 2) + '..' + appClose + ' du HTML source)',
  '//',
  '// NE PAS MODIFIER À LA MAIN sans réfléchir : ce fichier reproduit fidèlement le comportement',
  '// de l\'outil original. Toute modification visuelle relève de la Phase 2 ; toute modification',
  '// de logique métier doit être validée par le test de non-régression du PPTX généré.',
  '//',
  '// Dépendances globales attendues sur window :',
  '//   - JSZip  (chargé via <script src="/vendor/jszip.min.js"> dans index.html)',
  '//   - GEN    (chargé via <script src="/vendor/pptx-engine.js"> dans index.html)',
  '',
  "import { BDD, DEFAULT_OBLIG, DEFAULT_DOCS, SMILEY, SMILEY_FULL } from './catalog.js';",
  '',
  ...appBody,
  ''
].join('\n');

// 7) Extraire le CSS (entre <style> et </style>) → src/styles.css
const styleOpen = findLine(l => l.trim() === '<style>');
if (styleOpen < 0) throw new Error('Ouverture <style> introuvable');
const styleClose = findLine(l => l.trim() === '</style>', styleOpen + 1);
if (styleClose < 0) throw new Error('Fermeture </style> introuvable');
const cssCode = lines.slice(styleOpen + 1, styleClose).join('\n');

// 8) Extraire le markup body : tout ce qui est entre <body> et le premier <script>
const bodyOpen = findLine(l => l.trim() === '<body>');
if (bodyOpen < 0) throw new Error('Ouverture <body> introuvable');
const bodyClose = findLine(l => l.trim() === '</body>', bodyOpen + 1);
if (bodyClose < 0) throw new Error('Fermeture </body> introuvable');
const firstScriptInBody = findLine(l => /^\s*<script\b/.test(l), bodyOpen + 1);
const markupEnd = firstScriptInBody > 0 && firstScriptInBody < bodyClose ? firstScriptInBody : bodyClose;
const bodyMarkup = lines.slice(bodyOpen + 1, markupEnd).join('\n').replace(/\s+$/, '');

const indexHtml = `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>ALHEE — Saisie PPT (Plan Pluriannuel de Travaux)</title>
<meta name="description" content="Saisie de Plans Pluriannuels de Travaux en copropriété — outil ALHEE.">
<meta name="theme-color" content="#18483C">
<link rel="manifest" href="/manifest.webmanifest">
<link rel="icon" type="image/svg+xml" href="/icon.svg">
<link rel="apple-touch-icon" href="/apple-touch-icon-180.png">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="PPT ALHEE">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700;800&family=Manrope:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="/src/styles.css">
</head>
<body>
${bodyMarkup}
<!-- JSZip + moteur PPTX : IIFE qui posent window.JSZip et window.GEN. Gelés, ne pas modifier. -->
<script src="/vendor/jszip.min.js"></script>
<script src="/vendor/pptx-engine.js"></script>
<!-- Couche applicative : module ES, importe catalog.js -->
<script type="module" src="/src/app.js"></script>
<!-- PWA : enregistrement du service worker (production uniquement) -->
<script type="module" src="/src/pwa.js"></script>
</body>
</html>
`;

// 9) Écrire les sorties
//    JSZip et le moteur PPTX sont servis depuis /public/vendor/ (assets bruts via <script>),
//    pour préserver à 100% le comportement IIFE qui pose window.JSZip / window.GEN.
const outPptx = path.join(ROOT, 'public', 'modele-alhee.pptx');
const outJszip = path.join(ROOT, 'public', 'vendor', 'jszip.min.js');
const outEngine = path.join(ROOT, 'public', 'vendor', 'pptx-engine.js');
const outCatalog = path.join(ROOT, 'src', 'catalog.js');
const outApp = path.join(ROOT, 'src', 'app.js');
const outCss = path.join(ROOT, 'src', 'styles.css');
const outHtml = path.join(ROOT, 'index.html');

// Fichiers « vivants » : édités à la main à partir de la Phase 2 (refonte visuelle).
// L'extraction ne les écrase PLUS s'ils existent, sauf --force / FORCE=1.
// Les assets dérivés (pptx, jszip, engine, catalog) sont toujours régénérés.
const FORCE = process.argv.includes('--force') || process.env.FORCE === '1';
// catalog.js est désormais « vivant » lui aussi : les prix du catalogue y sont maintenus à la
// main (mise à jour via Excel), il ne doit donc plus être écrasé par une ré-extraction du HTML.
// modele-alhee.pptx aussi « vivant » : trame enrichie à la main (7 slides 1.6/1.7 + analyse
// énergétique, mono-master, ids sldId valides), ne plus l'écraser par TEMPLATE_B64.
const LIVING = new Set([outApp, outCss, outHtml, outCatalog, outPptx]);

for (const p of [outPptx, outJszip, outEngine, outCatalog, outApp, outCss, outHtml]) {
  fs.mkdirSync(path.dirname(p), { recursive: true });
}

function writeOut(p, data, label, note) {
  const living = LIVING.has(p);
  if (living && fs.existsSync(p) && !FORCE) {
    console.log(`• ${label.padEnd(28)} CONSERVÉ (fichier vivant existant — utilise --force pour régénérer)`);
    return;
  }
  fs.writeFileSync(p, data);
  console.log(`✓ ${label.padEnd(28)} ${note}`);
}

console.log('');
writeOut(outPptx, templateBytes, 'public/modele-alhee.pptx', `${templateBytes.length.toLocaleString('fr-FR')} octets`);
writeOut(outJszip, jszipCode, 'public/vendor/jszip.min.js', `${jszipCode.length.toLocaleString('fr-FR')} caractères (HTML L${jszipMarker + 2}..${jszipEnd})`);
writeOut(outEngine, engineCode, 'public/vendor/pptx-engine.js', `${engineCode.length.toLocaleString('fr-FR')} caractères (HTML L${engineOpen + 2}..${engineClose}) — GELÉ`);
writeOut(outCatalog, catalogCode, 'src/catalog.js', `BDD L${bdd.lineNo}, OBLIG L${dOblig.lineNo}, DOCS L${dDocs.lineNo}, SMILEY L${smiley.lineNo}`);
writeOut(outApp, appCode, 'src/app.js', `${appCode.length.toLocaleString('fr-FR')} caractères (HTML L${appOpen + 2}..${appClose})`);
writeOut(outCss, cssCode, 'src/styles.css', `${cssCode.length.toLocaleString('fr-FR')} caractères (HTML L${styleOpen + 2}..${styleClose})`);
writeOut(outHtml, indexHtml, 'index.html', `markup body HTML L${bodyOpen + 2}..${markupEnd}`);
console.log('');
console.log('→ Assets gelés régénérés. Fichiers vivants (css/html/app) préservés sauf --force.');
