// scripts/diff-pptx.mjs
// Compare deux fichiers .pptx (zips Office) en décompressant le contenu et en diffant
// caractère par caractère chaque fichier interne (XML, médias, rels).
//
// Usage : node scripts/diff-pptx.mjs <reference.pptx> <candidat.pptx>
//   ex. : npm run diff-pptx tests/reference-output.pptx out.pptx
//
// Sortie :
//   - Code de sortie 0  : pas de différence (test de non-régression OK)
//   - Code de sortie 1  : différences détectées (régression)

import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';

const [refPath, candPath] = process.argv.slice(2);
if (!refPath || !candPath) {
  console.error('Usage : node scripts/diff-pptx.mjs <reference.pptx> <candidat.pptx>');
  process.exit(2);
}
if (!fs.existsSync(refPath)) { console.error(`✗ Référence introuvable : ${refPath}`); process.exit(2); }
if (!fs.existsSync(candPath)) { console.error(`✗ Candidat introuvable : ${candPath}`); process.exit(2); }

async function loadZip(p) {
  const z = await JSZip.loadAsync(fs.readFileSync(p));
  const entries = {};
  for (const name of Object.keys(z.files)) {
    if (z.files[name].dir) continue;
    entries[name] = await z.file(name).async('uint8array');
  }
  return entries;
}

function bytesEqual(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function isText(name) {
  return /\.(xml|rels|json|txt)$/i.test(name);
}

console.log(`→ Référence : ${refPath}`);
console.log(`→ Candidat  : ${candPath}`);
console.log('');

const [ref, cand] = await Promise.all([loadZip(refPath), loadZip(candPath)]);

const refNames = new Set(Object.keys(ref));
const candNames = new Set(Object.keys(cand));
const onlyRef = [...refNames].filter(n => !candNames.has(n)).sort();
const onlyCand = [...candNames].filter(n => !refNames.has(n)).sort();
const common = [...refNames].filter(n => candNames.has(n)).sort();

let diffs = 0;
let identical = 0;

for (const name of common) {
  const a = ref[name], b = cand[name];
  if (bytesEqual(a, b)) { identical++; continue; }
  diffs++;
  console.log(`✗ ${name}  (référence ${a.length} octets / candidat ${b.length} octets)`);
  if (isText(name)) {
    const decA = new TextDecoder('utf-8', { fatal: false }).decode(a);
    const decB = new TextDecoder('utf-8', { fatal: false }).decode(b);
    // Diff simple : trouver le 1er offset divergent et afficher un extrait autour
    let i = 0;
    while (i < decA.length && i < decB.length && decA[i] === decB[i]) i++;
    const ctxStart = Math.max(0, i - 80);
    const ctxEnd = i + 120;
    console.log(`   1er écart à l'offset ${i} :`);
    console.log(`   réf : …${JSON.stringify(decA.slice(ctxStart, ctxEnd))}…`);
    console.log(`   cand: …${JSON.stringify(decB.slice(ctxStart, ctxEnd))}…`);
  }
}

for (const name of onlyRef)  { diffs++; console.log(`✗ Présent dans référence uniquement : ${name}`); }
for (const name of onlyCand) { diffs++; console.log(`✗ Présent dans candidat uniquement  : ${name}`); }

console.log('');
console.log(`Résumé : ${identical} fichiers identiques / ${common.length} en commun, ${diffs} divergences, ${onlyRef.length} en réf seule, ${onlyCand.length} en cand seul`);
console.log('');

if (diffs === 0) {
  console.log('✓ AUCUNE RÉGRESSION détectée — les deux PPTX sont strictement identiques.');
  process.exit(0);
} else {
  console.log('✗ RÉGRESSION : différences entre la référence et le candidat.');
  process.exit(1);
}
