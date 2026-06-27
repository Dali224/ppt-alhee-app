// scripts/make-icons.mjs
// Génère les icônes PNG de la PWA sans aucune dépendance (encodeur PNG via zlib).
// Design de marque : fond vert ALHEE, « diapositive » blanche, bandeau orange,
// trois lignes (évoque un PPT). Remplaçables par un logo définitif si besoin.
//
// Sorties : public/icon-192.png, public/icon-512.png, public/apple-touch-icon-180.png

import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

const GREEN = [0x18, 0x48, 0x3c];
const ORANGE = [0xe6, 0x7e, 0x22];
const WHITE = [0xff, 0xff, 0xff];
const GREY = [0xcd, 0xd8, 0xcf];

function render(S) {
  const buf = Buffer.alloc(S * S * 4);
  const set = (x, y, c) => {
    if (x < 0 || y < 0 || x >= S || y >= S) return;
    const i = (y * S + x) * 4;
    buf[i] = c[0]; buf[i + 1] = c[1]; buf[i + 2] = c[2]; buf[i + 3] = 255;
  };
  const rect = (x0, y0, x1, y1, c) => {
    for (let y = Math.round(y0); y < Math.round(y1); y++)
      for (let x = Math.round(x0); x < Math.round(x1); x++) set(x, y, c);
  };
  // fond vert plein
  rect(0, 0, S, S, GREEN);
  // diapositive blanche (zone de sécurité maskable : centre ~56%)
  const sx0 = .22 * S, sx1 = .78 * S, sy0 = .26 * S, sy1 = .74 * S;
  rect(sx0, sy0, sx1, sy1, WHITE);
  // bandeau orange (en-tête de diapo)
  rect(sx0, sy0, sx1, .38 * S, ORANGE);
  // lignes de contenu
  rect(.29 * S, .47 * S, .71 * S, .50 * S, GREY);
  rect(.29 * S, .55 * S, .63 * S, .58 * S, GREY);
  rect(.29 * S, .63 * S, .50 * S, .66 * S, ORANGE);
  return buf;
}

// — Encodeur PNG minimal (RGBA, filtre 0) —
function png(S, rgba) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length, 0);
    const tb = Buffer.from(type, 'ascii');
    const body = Buffer.concat([tb, data]);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(body) >>> 0, 0);
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(S, 0); ihdr.writeUInt32BE(S, 4);
  ihdr[8] = 8;   // bit depth
  ihdr[9] = 6;   // color type RGBA
  ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  // scanlines avec octet de filtre 0
  const raw = Buffer.alloc(S * (S * 4 + 1));
  for (let y = 0; y < S; y++) {
    raw[y * (S * 4 + 1)] = 0;
    rgba.copy(raw, y * (S * 4 + 1) + 1, y * S * 4, (y + 1) * S * 4);
  }
  const idat = zlib.deflateSync(raw, { level: 9 });
  return Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', idat), chunk('IEND', Buffer.alloc(0))]);
}

const targets = [
  { size: 512, file: 'icon-512.png' },
  { size: 192, file: 'icon-192.png' },
  { size: 180, file: 'apple-touch-icon-180.png' },
];

for (const t of targets) {
  const out = path.join(ROOT, 'public', t.file);
  fs.writeFileSync(out, png(t.size, render(t.size)));
  console.log(`✓ public/${t.file.padEnd(26)} ${t.size}×${t.size}`);
}
console.log('\n→ Icônes générées (placeholders de marque, remplaçables par un logo définitif).');
