// src/graph/missions.js — CRUD des missions sur SharePoint.
//   • JSON de la mission → bibliothèque PPT-Data : missions/{missionId}/mission.json
//   • Photos → fichiers : missions/{missionId}/photos/{photoId}.jpg (jamais en base64 dans le JSON)
//   • Une ligne par mission dans la liste PPT-Index (pour lister/filtrer)
//   • Conflit « dernier qui écrit gagne » : avertissement si l'eTag serveur a changé depuis le chargement

import { graphFetch } from './graph-client.js';
import { config } from '../config.js';
import { getSiteId, getListId, getDriveId, ensureIndexSchema } from './sharepoint.js';

let _schemaReady = null;
function ensureSchemaOnce() {
  if (!_schemaReady) _schemaReady = ensureIndexSchema();
  return _schemaReady;
}

/* ------------------------------ Photos ------------------------------ */
function dataUrlToBytes(d) {
  const b64 = d.split(',')[1] || '';
  const bin = atob(b64);
  const u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return u;
}
function hashStr(s) {
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}
async function uploadPhoto(driveId, base, photoId, dataUrl) {
  const path = `${base}/photos/${photoId}.jpg`;
  await graphFetch(`/drives/${driveId}/root:/${encodeURI(path)}:/content`, {
    method: 'PUT', raw: true, body: dataUrlToBytes(dataUrl), headers: { 'Content-Type': 'image/jpeg' },
  });
  return `photos/${photoId}.jpg`;
}
async function fetchPhotoDataUrl(driveId, base, ref) {
  const res = await graphFetch(`/drives/${driveId}/root:/${encodeURI(base + '/' + ref)}:/content`);
  const blob = res instanceof Response ? await res.blob() : null;
  if (!blob) return '';
  return await new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(blob);
  });
}
// Remplace les photos dataURL par des références (et les téléverse). Mute la copie `save`.
async function extractPhotos(driveId, base, live, save) {
  const warnings = [];
  const liveLots = live.lots || {}, saveLots = save.lots || {};
  for (const num of Object.keys(saveLots)) {
    const liveSubs = ((liveLots[num] || {}).subs) || {}, saveSubs = ((saveLots[num] || {}).subs) || {};
    for (const name of Object.keys(saveSubs)) {
      const liveItems = ((liveSubs[name] || {}).items) || [], saveItems = ((saveSubs[name] || {}).items) || [];
      for (let i = 0; i < saveItems.length; i++) {
        const sit = saveItems[i], lit = liveItems[i] || {};
        if (sit && typeof sit.photo === 'string' && sit.photo.startsWith('data:')) {
          const h = hashStr(sit.photo);
          const photoId = lit.photoId || ('p_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6));
          try {
            if (!(lit.photoRef && lit.photoHash === h)) {
              await uploadPhoto(driveId, base, photoId, sit.photo);
            }
            sit.photoId = photoId; sit.photoRef = `photos/${photoId}.jpg`; sit.photoHash = h;
            delete sit.photo; // pas de base64 dans le JSON serveur
          } catch (e) {
            warnings.push(`Photo non envoyée (${name} #${i + 1}) : ${e.message || e}`);
            // repli : on laisse le dataURL dans le JSON pour ne pas perdre la photo
          }
        }
      }
    }
  }
  return warnings;
}
async function hydratePhotos(driveId, base, data) {
  const lots = data.lots || {};
  for (const num of Object.keys(lots)) {
    const subs = ((lots[num] || {}).subs) || {};
    for (const name of Object.keys(subs)) {
      const items = ((subs[name] || {}).items) || [];
      for (const it of items) {
        if (it && it.photoRef && !it.photo) {
          try { it.photo = await fetchPhotoDataUrl(driveId, base, it.photoRef); } catch (_) { /* photo manquante */ }
        }
      }
    }
  }
}

/* ------------------------------ Index ------------------------------ */
async function upsertIndexRow(siteId, listId, missionId, fields) {
  const r = await graphFetch(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`);
  const existing = (r.value || []).find((it) => it.fields && it.fields.MissionId === missionId);
  if (existing) {
    await graphFetch(`/sites/${siteId}/lists/${listId}/items/${existing.id}/fields`, { method: 'PATCH', body: fields });
  } else {
    await graphFetch(`/sites/${siteId}/lists/${listId}/items`, { method: 'POST', body: { fields } });
  }
}

/* ------------------------------ API publique ------------------------------ */
export async function listMissions() {
  const siteId = await getSiteId();
  const listId = await getListId(config.sharepoint.indexList);
  const r = await graphFetch(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500`);
  return (r.value || [])
    .filter((it) => it.fields && it.fields.MissionId)
    .map((it) => ({
      itemId: it.id,
      missionId: it.fields.MissionId,
      copro: it.fields.Title || '(sans nom)',
      syndic: it.fields.Syndic || '',
      statut: it.fields.Statut || '',
      auteur: it.fields.Auteur || '',
      modified: it.lastModifiedDateTime || '',
    }))
    .sort((a, b) => (b.modified || '').localeCompare(a.modified || ''));
}

export async function loadMission(missionId) {
  const driveId = await getDriveId(config.sharepoint.dataLibrary);
  const base = `missions/${missionId}`;
  const path = `${base}/mission.json`;
  const meta = await graphFetch(`/drives/${driveId}/root:/${encodeURI(path)}`);
  const etag = meta.eTag || meta.cTag || '';
  const res = await graphFetch(`/drives/${driveId}/root:/${encodeURI(path)}:/content`);
  const data = res instanceof Response ? JSON.parse(await res.text()) : res;
  data.missionId = missionId;
  data.__serverETag = etag;
  await hydratePhotos(driveId, base, data);
  return data;
}

// project : mission courante (clone). opts.force = écraser malgré un conflit.
export async function saveMission(project, opts = {}) {
  const force = !!opts.force;
  const driveId = await getDriveId(config.sharepoint.dataLibrary);
  const listId = await ensureSchemaOnce();
  const siteId = await getSiteId();

  const missionId = project.missionId || ('m_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7));
  const base = `missions/${missionId}`;
  const path = `${base}/mission.json`;

  // Détection de conflit (mission déjà connue + eTag mémorisé au chargement)
  if (project.missionId && project.__serverETag && !force) {
    let serverTag = '';
    try {
      const meta = await graphFetch(`/drives/${driveId}/root:/${encodeURI(path)}`);
      serverTag = meta.eTag || meta.cTag || '';
    } catch (_) { serverTag = ''; }
    if (serverTag && serverTag !== project.__serverETag) {
      const e = new Error('CONFLICT'); e.code = 'CONFLICT'; throw e;
    }
  }

  // Copie pour le JSON : photos → références
  const toSave = JSON.parse(JSON.stringify(project));
  toSave.missionId = missionId;
  delete toSave.__serverETag;
  const photoWarnings = await extractPhotos(driveId, base, project, toSave);

  // Écrit le JSON
  const put = await graphFetch(`/drives/${driveId}/root:/${encodeURI(path)}:/content`, {
    method: 'PUT', raw: true,
    body: new TextEncoder().encode(JSON.stringify(toSave)),
    headers: { 'Content-Type': 'application/json' },
  });
  const etag = put.eTag || put.cTag || '';

  // Met à jour / crée la ligne d'index
  await upsertIndexRow(siteId, listId, missionId, {
    Title: (project.meta && project.meta.copro) || '(sans nom)',
    MissionId: missionId,
    Syndic: (project.meta && project.meta.syndic) || '',
    Statut: project.statut || 'En cours',
    DateMission: new Date().toISOString(),
    Auteur: (project.meta && project.meta.auditeur) || '',
    LienDocument: path,
  });

  return { missionId, etag, photoWarnings };
}

export async function deleteMission(missionId, itemId) {
  const driveId = await getDriveId(config.sharepoint.dataLibrary);
  const siteId = await getSiteId();
  const listId = await getListId(config.sharepoint.indexList);
  try { await graphFetch(`/drives/${driveId}/root:/${encodeURI('missions/' + missionId)}`, { method: 'DELETE' }); } catch (_) {}
  if (itemId) {
    try { await graphFetch(`/sites/${siteId}/lists/${listId}/items/${itemId}`, { method: 'DELETE' }); } catch (_) {}
  }
}
