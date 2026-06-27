// scripts/generate-reference.mjs
// Génère le PPTX de référence pour les tests de non-régression.
//
// Charge le moteur PPTX gelé (public/vendor/pptx-engine.js) dans un sandbox Node,
// applique tests/reference-project.json sur public/modele-alhee.pptx,
// et écrit le résultat dans tests/reference-output.pptx.
//
// Usage : npm run generate-reference  (ou : node scripts/generate-reference.mjs)

import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { SMILEY, SMILEY_FULL } from '../src/catalog.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

// 1) Charger le moteur PPTX gelé dans un sandbox vm.
//    Le moteur fait : `if (module.exports) { root.NODE=true; module.exports=api } else { root.GEN=api }`.
//    On force `module.exports = {}` pour passer le test truthy → expose via `module.exports`.
const engineSrc = fs.readFileSync(path.join(ROOT, 'public', 'vendor', 'pptx-engine.js'), 'utf8');
const ctx = { module: { exports: {} }, exports: {}, console };
vm.createContext(ctx);
vm.runInContext(engineSrc, ctx);
const GEN = ctx.module.exports;
if (!GEN || typeof GEN.generate !== 'function') {
  throw new Error('Moteur PPTX non chargé : module.exports.generate manquant');
}

// 2) Charger le template et le projet de référence
const templateBytes = fs.readFileSync(path.join(ROOT, 'public', 'modele-alhee.pptx'));
const project = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'reference-project.json'), 'utf8'));

console.log(`→ Template :        ${templateBytes.length.toLocaleString('fr-FR')} octets`);
console.log(`→ Projet réf. :     ${project.lots.length} lots, ${project.lots.reduce((s, L) => s + L.subs.reduce((t, S) => t + (S.items?.length || 0), 0), 0)} items`);

// 3) Générer le PPTX
const t0 = Date.now();
const out = await GEN.generate(JSZip, templateBytes, project, { smiley: SMILEY, smileyFull: SMILEY_FULL });
const dt = Date.now() - t0;

// Path de sortie : par défaut tests/reference-output.pptx, peut être overridé via env OUTPUT
const outPath = process.env.OUTPUT
  ? path.resolve(ROOT, process.env.OUTPUT)
  : path.join(ROOT, 'tests', 'reference-output.pptx');
fs.writeFileSync(outPath, out);

const relOutPath = path.relative(ROOT, outPath).replace(/\\/g, '/');
console.log('');
console.log(`✓ ${relOutPath}     ${out.length.toLocaleString('fr-FR')} octets, généré en ${dt} ms`);
if (!process.env.OUTPUT) {
  console.log('');
  console.log('→ Vérifie que ce fichier s\'ouvre correctement dans PowerPoint.');
  console.log('→ Sauvegarde-le. Toute modification ultérieure du moteur ou du catalogue doit');
  console.log('  reproduire EXACTEMENT le même fichier — sinon : régression à corriger.');
}
