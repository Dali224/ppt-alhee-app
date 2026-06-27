// src/graph/sharepoint.js — accès au site : résolution site/liste/drive + provisionnement.

import { graphFetch } from './graph-client.js';
import { config } from '../config.js';

let _siteId = null;
const _driveCache = {};

// Colonnes de la liste-index (Title existe déjà → copropriété). MissionId relie la ligne au JSON.
const INDEX_COLUMNS = [
  { name: 'MissionId', text: {} },
  { name: 'Syndic', text: {} },
  { name: 'Statut', text: {} },
  { name: 'DateMission', dateTime: {} },
  { name: 'Auteur', text: {} },
  { name: 'LienDocument', text: {} },
];

export async function getSiteId() {
  if (_siteId) return _siteId;
  const { hostname, sitePath } = config.sharepoint;
  const site = await graphFetch(`/sites/${hostname}:${sitePath}`);
  _siteId = site.id;
  return _siteId;
}

async function findList(displayName) {
  const siteId = await getSiteId();
  const q = encodeURIComponent(`displayName eq '${displayName.replace(/'/g, "''")}'`);
  const r = await graphFetch(`/sites/${siteId}/lists?$filter=${q}&$select=id,displayName`);
  return (r.value && r.value[0]) || null;
}

export async function getListId(displayName) {
  const l = await findList(displayName);
  if (!l) throw new Error(`Liste « ${displayName} » introuvable — lance « ⚙️ Initialiser SharePoint ».`);
  return l.id;
}

// Drive (= bibliothèque de documents) par nom. PPT-Data n'est pas la bibliothèque par défaut.
export async function getDriveId(libraryName) {
  if (_driveCache[libraryName]) return _driveCache[libraryName];
  const siteId = await getSiteId();
  const r = await graphFetch(`/sites/${siteId}/drives?$select=id,name`);
  const d = (r.value || []).find((x) => x.name === libraryName);
  if (!d) throw new Error(`Bibliothèque « ${libraryName} » introuvable — lance « ⚙️ Initialiser SharePoint ».`);
  _driveCache[libraryName] = d.id;
  return d.id;
}

async function ensureColumn(siteId, listId, def) {
  try {
    await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, { method: 'POST', body: def });
  } catch (e) {
    if (!/already\s*exists|nameAlreadyExists|exists/i.test(e.message)) throw e;
  }
}

// Garantit que la liste-index a toutes ses colonnes (idempotent). Renvoie son id.
export async function ensureIndexSchema() {
  const siteId = await getSiteId();
  const listId = await getListId(config.sharepoint.indexList);
  for (const c of INDEX_COLUMNS) await ensureColumn(siteId, listId, c);
  return listId;
}

// Crée la liste PPT-Index (+colonnes) et la bibliothèque PPT-Data si absentes. Idempotent.
export async function ensureProvisioned() {
  const siteId = await getSiteId();
  const log = [];

  let idx = await findList(config.sharepoint.indexList);
  if (!idx) {
    idx = await graphFetch(`/sites/${siteId}/lists`, {
      method: 'POST',
      body: { displayName: config.sharepoint.indexList, list: { template: 'genericList' } },
    });
    log.push(`✓ Liste « ${config.sharepoint.indexList} » créée`);
  } else {
    log.push(`• Liste « ${config.sharepoint.indexList} » déjà présente`);
  }
  for (const c of INDEX_COLUMNS) await ensureColumn(siteId, idx.id, c);
  log.push('✓ Colonnes vérifiées (MissionId, Syndic, Statut, DateMission, Auteur, LienDocument)');

  let lib = await findList(config.sharepoint.dataLibrary);
  if (!lib) {
    lib = await graphFetch(`/sites/${siteId}/lists`, {
      method: 'POST',
      body: { displayName: config.sharepoint.dataLibrary, list: { template: 'documentLibrary' } },
    });
    log.push(`✓ Bibliothèque « ${config.sharepoint.dataLibrary} » créée`);
  } else {
    log.push(`• Bibliothèque « ${config.sharepoint.dataLibrary} » déjà présente`);
  }

  return log;
}
