// src/graph/graph-client.js — appels Microsoft Graph (fetch + jeton MSAL).
// Pas de SDK : un simple wrapper fetch avec en-tête Bearer, gestion d'erreurs et
// un retry anti-throttling (429/503).

import { getGraphToken } from '../auth/auth.js';
import { config } from '../config.js';

export async function graphFetch(path, { method = 'GET', body = null, headers = {}, raw = false } = {}) {
  const token = await getGraphToken();
  const url = path.startsWith('http') ? path : config.graphBase + path;
  const opts = { method, headers: { Authorization: 'Bearer ' + token, ...headers } };
  if (body != null) {
    if (raw) {
      opts.body = body; // ArrayBuffer / Blob (upload de fichiers)
    } else {
      opts.body = JSON.stringify(body);
      opts.headers['Content-Type'] = 'application/json';
    }
  }

  let res = await fetch(url, opts);
  if (res.status === 429 || res.status === 503) {
    const wait = (Number(res.headers.get('Retry-After')) || 2) * 1000;
    await new Promise((r) => setTimeout(r, wait));
    res = await fetch(url, opts);
  }
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Graph ${method} ${path} → ${res.status} ${txt.slice(0, 400)}`);
  }
  if (res.status === 204) return null;
  const ct = res.headers.get('Content-Type') || '';
  return ct.includes('application/json') ? res.json() : res;
}
