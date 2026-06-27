// src/auth/auth-ui.js — UI cloud (barre) : connexion, Enregistrer sur SharePoint,
// navigateur de missions, et bouton d'initialisation. La logique métier reste dans app.js ;
// on passe par window.ALHEE_BRIDGE pour lire/charger la mission courante.

import { initMsal, signIn, signOut, getActiveAccount } from './auth.js';
import { listMissions, loadMission, saveMission, deleteMission } from '../graph/missions.js';
import { ensureProvisioned } from '../graph/sharepoint.js';

const RESUME_KEY = 'alhee_resume_init';

export async function initCloudUI() {
  await initMsal();
  injectUI();
  injectModal();
  renderAuthState();
  if (getActiveAccount() && sessionStorage.getItem(RESUME_KEY)) {
    sessionStorage.removeItem(RESUME_KEY);
    onInitClick(true);
  }
}

/* ------------------------------ Barre ------------------------------ */
function injectUI() {
  const bar = document.getElementById('bar');
  if (!bar || document.getElementById('authWrap')) return;
  const wrap = document.createElement('span');
  wrap.id = 'authWrap';
  wrap.style.cssText = 'display:flex;align-items:center;gap:8px;margin-left:8px;flex-wrap:wrap';
  wrap.innerHTML =
    '<span id="authWho" class="sub" style="opacity:.9"></span>' +
    '<button id="cloudSaveBtn" class="btn" style="display:none">☁️ Enregistrer</button>' +
    '<button id="cloudMissionsBtn" class="btn" style="display:none">🗂️ Missions</button>' +
    '<button id="initSpBtn" class="btn" title="Créer/réparer la liste PPT-Index et la bibliothèque PPT-Data" style="display:none">⚙️ Initialiser</button>' +
    '<button id="authBtn" class="btn">Se connecter</button>';
  bar.appendChild(wrap);
  document.getElementById('authBtn').onclick = onAuthClick;
  document.getElementById('initSpBtn').onclick = () => onInitClick(false);
  document.getElementById('cloudSaveBtn').onclick = onSaveClick;
  document.getElementById('cloudMissionsBtn').onclick = openMissions;
}

function renderAuthState() {
  const acct = getActiveAccount();
  const who = document.getElementById('authWho');
  const btn = document.getElementById('authBtn');
  if (!btn || !who) return;
  const cloudBtns = ['cloudSaveBtn', 'cloudMissionsBtn', 'initSpBtn'].map((id) => document.getElementById(id));
  if (acct) {
    who.textContent = acct.name || acct.username || '';
    btn.textContent = 'Se déconnecter';
    cloudBtns.forEach((b) => b && (b.style.display = ''));
  } else {
    who.textContent = '';
    btn.textContent = 'Se connecter';
    cloudBtns.forEach((b) => b && (b.style.display = 'none'));
  }
}

async function onAuthClick() {
  const btn = document.getElementById('authBtn');
  const acct = getActiveAccount();
  try {
    if (btn) btn.disabled = true;
    if (acct) await signOut();
    else await signIn();
  } catch (e) {
    console.warn('[auth] action', e);
    alert('Connexion impossible : ' + (e && e.message ? e.message : e));
  } finally {
    if (btn) btn.disabled = false;
    renderAuthState();
  }
}

async function onInitClick(isResume) {
  const btn = document.getElementById('initSpBtn');
  const old = btn ? btn.textContent : '';
  try {
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Initialisation…'; }
    if (!isResume) sessionStorage.setItem(RESUME_KEY, '1');
    const log = await ensureProvisioned();
    sessionStorage.removeItem(RESUME_KEY);
    alert('SharePoint initialisé :\n\n' + log.join('\n'));
  } catch (e) {
    sessionStorage.removeItem(RESUME_KEY);
    console.error('[sharepoint] init', e);
    alert('Échec de l’initialisation SharePoint :\n\n' + (e && e.message ? e.message : e));
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = old; }
  }
}

/* ------------------------------ Enregistrer ------------------------------ */
async function onSaveClick() {
  const btn = document.getElementById('cloudSaveBtn');
  const old = btn.textContent;
  const bridge = window.ALHEE_BRIDGE;
  if (!bridge) { alert('Pont applicatif indisponible.'); return; }
  const project = bridge.getProject();
  if (!project.meta || !(project.meta.copro || '').trim()) {
    if (!confirm('La mission n’a pas de nom de copropriété (onglet Généralités).\nEnregistrer quand même ?')) return;
  }
  try {
    btn.disabled = true; btn.textContent = '⏳ Enregistrement…';
    let res;
    try {
      res = await saveMission(project);
    } catch (e) {
      if (e && e.code === 'CONFLICT') {
        if (!confirm('⚠️ Une version plus récente existe sur SharePoint (un collègue l’a peut-être modifiée).\n\nÉcraser avec ta version actuelle ?')) return;
        res = await saveMission(project, { force: true });
      } else throw e;
    }
    bridge.markSaved(res.missionId, res.etag);
    bridge.toast('☁️ Enregistré sur SharePoint');
    if (res.photoWarnings && res.photoWarnings.length) {
      alert('Mission enregistrée, mais certaines photos n’ont pas pu être envoyées :\n\n' + res.photoWarnings.join('\n'));
    }
  } catch (e) {
    console.error('[missions] save', e);
    alert('Enregistrement SharePoint impossible :\n\n' + (e && e.message ? e.message : e));
  } finally {
    btn.disabled = false; btn.textContent = old;
  }
}

/* ------------------------------ Navigateur de missions ------------------------------ */
function injectModal() {
  if (document.getElementById('cloudModal')) return;
  const m = document.createElement('div');
  m.id = 'cloudModal';
  m.style.cssText =
    'position:fixed;inset:0;z-index:1000;display:none;align-items:flex-start;justify-content:center;' +
    'background:rgba(11,26,18,.55);backdrop-filter:blur(2px);padding:40px 14px;overflow-y:auto';
  m.innerHTML =
    '<div class="vrs-card">' +
    '<div class="vrs-head"><span class="vrs-title">🗂️ Missions sur SharePoint</span>' +
    '<button id="cloudClose" class="vrs-x" title="Fermer">✕</button></div>' +
    '<div class="vrs-hint">Ouvre une mission partagée pour la modifier, ou enregistre la mission courante via « ☁️ Enregistrer ». Les missions sont communes à toute l’équipe (selon les accès SharePoint).</div>' +
    '<div id="cloudList" class="vrs-list"></div>' +
    '</div>';
  document.body.appendChild(m);
  document.getElementById('cloudClose').onclick = closeMissions;
  m.addEventListener('click', (e) => { if (e.target === m) closeMissions(); });
}

function closeMissions() {
  const m = document.getElementById('cloudModal');
  if (m) m.style.display = 'none';
}

async function openMissions() {
  const m = document.getElementById('cloudModal');
  const list = document.getElementById('cloudList');
  if (!m || !list) return;
  m.style.display = 'flex';
  list.innerHTML = '<div class="vrs-empty">Chargement…</div>';
  let missions = [];
  try {
    missions = await listMissions();
  } catch (e) {
    list.innerHTML = '<div class="vrs-empty">Erreur de lecture : ' + (e && e.message ? e.message : e) + '</div>';
    return;
  }
  if (!missions.length) {
    list.innerHTML = '<div class="vrs-empty">Aucune mission sur SharePoint pour l’instant.<br>Ouvre/saisis une mission puis clique « ☁️ Enregistrer ».</div>';
    return;
  }
  list.innerHTML = '';
  missions.forEach((mi) => {
    const row = document.createElement('div');
    row.className = 'vrs-row';
    const meta = [mi.syndic, mi.auteur, _fmtDate(mi.modified)].filter(Boolean).join(' · ');
    row.innerHTML =
      '<div class="vrs-info"><div class="vrs-name"></div><div class="vrs-meta"></div></div>' +
      '<div class="vrs-actions">' +
      '<button class="vrs-b vrs-load">Ouvrir</button>' +
      '<button class="vrs-b vrs-del">Supprimer</button></div>';
    row.querySelector('.vrs-name').textContent = mi.copro + (mi.statut ? ' — ' + mi.statut : '');
    row.querySelector('.vrs-meta').textContent = meta;
    row.querySelector('.vrs-load').onclick = () => onOpenMission(mi);
    row.querySelector('.vrs-del').onclick = () => onDeleteMission(mi);
    list.appendChild(row);
  });
}

async function onOpenMission(mi) {
  const bridge = window.ALHEE_BRIDGE;
  if (!bridge) return;
  try {
    const data = await loadMission(mi.missionId);
    bridge.setProject(data);
    bridge.toast('✓ Mission chargée : ' + (mi.copro || ''));
    closeMissions();
  } catch (e) {
    console.error('[missions] open', e);
    alert('Ouverture impossible :\n\n' + (e && e.message ? e.message : e));
  }
}

async function onDeleteMission(mi) {
  if (!confirm('Supprimer définitivement la mission « ' + (mi.copro || '') + ' » de SharePoint ?')) return;
  try {
    await deleteMission(mi.missionId, mi.itemId);
    if (window.ALHEE_BRIDGE) window.ALHEE_BRIDGE.toast('🗑 Mission supprimée');
    openMissions();
  } catch (e) {
    console.error('[missions] delete', e);
    alert('Suppression impossible :\n\n' + (e && e.message ? e.message : e));
  }
}

function _fmtDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return '';
  const p = (n) => String(n).padStart(2, '0');
  return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear() + ' ' + p(d.getHours()) + 'h' + p(d.getMinutes());
}
