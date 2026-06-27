// scripts/test-no-regression.mjs
// Test de non-régression de l'export PPTX.
//
// Workflow :
//   1. Régénère un PPTX depuis tests/reference-project.json → tests/.candidate.pptx
//   2. Diff binaire contre tests/reference-output.pptx (l'oracle figé)
//   3. Nettoie le fichier candidat
//   4. Exit 0 si identique, 1 si régression

import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const oracle = path.join(ROOT, 'tests', 'reference-output.pptx');
const candidate = path.join(ROOT, 'tests', '.candidate.pptx');

if (!fs.existsSync(oracle)) {
  console.error(`✗ Oracle introuvable : ${oracle}`);
  console.error('  Lance d\'abord : npm run generate-reference');
  process.exit(2);
}

console.log('=== Test de non-régression du moteur PPTX ===');
console.log('');

// 1) Régénérer
const gen = spawnSync('node', ['scripts/generate-reference.mjs'], {
  cwd: ROOT,
  stdio: 'inherit',
  env: { ...process.env, OUTPUT: candidate }
});
if (gen.status !== 0) {
  console.error('✗ Échec de la génération du candidat');
  process.exit(2);
}

// 2) Diff
console.log('');
const diff = spawnSync('node', ['scripts/diff-pptx.mjs', oracle, candidate], {
  cwd: ROOT,
  stdio: 'inherit'
});

// 3) Cleanup (toujours, même si erreur)
try { fs.unlinkSync(candidate); } catch {}

// 4) Code de sortie
process.exit(diff.status ?? 1);
