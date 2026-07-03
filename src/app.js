// src/app.js
// Couche applicative ALHEE — copie verbatim du 3e <script> de Saisie_PPT_ALHEE_12_1.html
// (lignes 623..1266 du HTML source)
//
// NE PAS MODIFIER À LA MAIN sans réfléchir : ce fichier reproduit fidèlement le comportement
// de l'outil original. Toute modification visuelle relève de la Phase 2 ; toute modification
// de logique métier doit être validée par le test de non-régression du PPTX généré.
//
// Dépendances globales attendues sur window :
//   - JSZip  (chargé via <script src="/vendor/jszip.min.js"> dans index.html)
//   - GEN    (chargé via <script src="/vendor/pptx-engine.js"> dans index.html)

import { BDD, DEFAULT_OBLIG, DEFAULT_DOCS, SMILEY, SMILEY_FULL } from './catalog.js';

/* ====== COMPRESSION PHOTO (Phase 3 — mobile) ======
   Redimensionne à PHOTO_MAX_DIM px max et ré-encode en JPEG (PHOTO_QUALITY).
   Respecte l'orientation EXIF (createImageBitmap when available). Le résultat est
   un dataURL stocké tel quel dans it.photo (format inchangé pour le moteur PPTX).
   En cas d'échec : repli sur la lecture brute (comportement d'origine). */
const PHOTO_MAX_DIM = 1600;
const PHOTO_QUALITY = 0.82;
async function compressImageFile(file){
  try{
    if(!file || !/^image\//.test(file.type||'')) throw new Error('not-an-image');
    let bmp;
    if('createImageBitmap' in window){
      try{ bmp = await createImageBitmap(file, {imageOrientation:'from-image'}); }
      catch(_){ bmp = await createImageBitmap(file); }
    } else {
      bmp = await new Promise((res,rej)=>{ const img=new Image(); img.onload=()=>res(img); img.onerror=rej; img.src=URL.createObjectURL(file); });
    }
    let w = bmp.width, h = bmp.height;
    const scale = Math.min(1, PHOTO_MAX_DIM/Math.max(w,h));
    w = Math.max(1, Math.round(w*scale)); h = Math.max(1, Math.round(h*scale));
    const canvas = document.createElement('canvas'); canvas.width=w; canvas.height=h;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff'; ctx.fillRect(0,0,w,h);   // évite le fond noir si PNG transparent
    ctx.drawImage(bmp, 0, 0, w, h);
    if(bmp.close) bmp.close();
    return canvas.toDataURL('image/jpeg', PHOTO_QUALITY);
  }catch(e){
    return await new Promise((res,rej)=>{ const r=new FileReader(); r.onload=ev=>res(ev.target.result); r.onerror=rej; r.readAsDataURL(file); });
  }
}

/* ====== CATALOGUE (BDD ALHEE, fixe) ====== */
// (L624 extrait — voir catalog.js ou public/modele-alhee.pptx)

/* ====== ÉTAT PROJET ====== */
// (L627 extrait — voir catalog.js ou public/modele-alhee.pptx)
// (L628 extrait — voir catalog.js ou public/modele-alhee.pptx)
function rid(){ return 'r'+Math.random().toString(36).slice(2,9); }
let project = newProject();
function newProject(){
  const p={meta:{},car:{},lots:{},zone:'paris',tva:10,synthese:'',conclusion:'',synthGen:{note:'',etats:{}},p0:[]};
  p.oblig=JSON.parse(JSON.stringify(DEFAULT_OBLIG)); p.docs=JSON.parse(JSON.stringify(DEFAULT_DOCS));
  BDD.lots.forEach(L=>{ p.lots[L.num]={subs:{}}; L.subs.forEach(s=>{ p.lots[L.num].subs[s.name]={remark:'',items:[]}; }); });
  return p;
}

/* ====== SMILEYS ====== */
function smiley(kind,size){ size=size||34;
  const c=kind==='bon'?'#3E7C4F':kind==='moyen'?'#3E7C4F':'#C0392B';
  let m=kind==='bon'?'<path d="M16 36 Q32 50 48 36" fill="none" stroke="'+c+'" stroke-width="4" stroke-linecap="round"/>'
    :kind==='moyen'?'<path d="M16 46 Q32 32 48 46" fill="none" stroke="'+c+'" stroke-width="4" stroke-linecap="round"/>'
    :'<path d="M16 46 Q32 32 48 46" fill="none" stroke="'+c+'" stroke-width="4" stroke-linecap="round"/>';
  return '<svg class="smiley" viewBox="0 0 64 64" width="'+size+'" height="'+size+'"><circle cx="32" cy="32" r="29" fill="none" stroke="'+c+'" stroke-width="4"/><circle cx="23" cy="26" r="3.4" fill="'+c+'"/><circle cx="41" cy="26" r="3.4" fill="'+c+'"/>'+m+'</svg>';
}

/* ====== SYNTHÈSE GÉNÉRALE (3.1) ====== */
function _normName(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,''); }
function aggregateEtat(sub){ const items=(sub&&sub.items)||[]; let w=null; items.forEach(it=>{ const e=it.etat; if(e==='urgent')w='urgent'; else if(e==='moyen'&&w!=='urgent')w='moyen'; else if(e==='bon'&&!w)w='bon'; }); return w||'na'; }
function _emoP(e){ return (e==='bon'||e==='moyen'||e==='urgent')? smiley(e,22) : '<span style="color:#bbb;font-size:20px">–</span>'; }
function computeSynthGen(){ const sg=project.synthGen||{note:'',etats:{}}; const out={}; BDD.lots.forEach(function(L){ L.subs.forEach(function(s){ var key=L.num+'|'+s.name; var e=sg.etats[key]||'auto'; if(e==='auto') e=aggregateEtat(project.lots[L.num].subs[s.name]); out[_normName(s.name)]=e; }); }); return {note:(sg.note||''),etats:out}; }
function renderSynG(m){
  if(!project.synthGen) project.synthGen={note:'',etats:{}};
  var sg=project.synthGen;
  var h='<div class="lot-h">Synthèse générale (3.1)</div>'
    +'<div class="subcard open"><div class="sh">Note générale<span class="chev"> </span></div><div class="sb">'
    +'<p class="hint">Petite note affichée en tête de la page de synthèse, au-dessus des tableaux. L\'état de chaque sous-lot est déduit des anomalies saisies (le plus grave l\'emporte) ; vous pouvez le forcer.</p>'
    +'<textarea id="sgNote" class="ai-ta" style="min-height:70px" placeholder="Note générale…">'+escAttr(sg.note||'')+'</textarea></div></div>';
  var lbl={na:'N/A',bon:'Bon',moyen:'Moyen',urgent:'Urgent'};
  BDD.lots.forEach(function(L){
    h+='<div class="subcard open"><div class="sh">'+L.num+' '+escAttr(L.label)+'<span class="chev"> </span></div><div class="sb"><table style="width:100%;border-collapse:collapse">';
    L.subs.forEach(function(s){
      var key=L.num+'|'+s.name; var auto=aggregateEtat(project.lots[L.num].subs[s.name]); var cur=sg.etats[key]||'auto'; var eff=(cur==='auto')?auto:cur;
      h+='<tr><td style="padding:4px 6px;width:58%;border-bottom:1px solid #eee">'+escAttr(s.name)+'</td>'
        +'<td style="padding:4px 6px;border-bottom:1px solid #eee"><select class="sel sgSel" data-k="'+key+'" data-auto="'+auto+'">'
        +'<option value="auto"'+(cur==='auto'?' selected':'')+'>Auto ('+(lbl[auto]||'N/A')+')</option>'
        +'<option value="bon"'+(cur==='bon'?' selected':'')+'>Bon</option>'
        +'<option value="moyen"'+(cur==='moyen'?' selected':'')+'>Moyen</option>'
        +'<option value="urgent"'+(cur==='urgent'?' selected':'')+'>Urgent</option>'
        +'<option value="na"'+(cur==='na'?' selected':'')+'>N/A</option></select></td>'
        +'<td class="sgEmo" style="padding:4px 6px;width:46px;text-align:center;border-bottom:1px solid #eee">'+_emoP(eff)+'</td></tr>';
    });
    h+='</table></div></div>';
  });
  m.innerHTML=h;
  var nt=document.getElementById('sgNote'); if(nt) nt.oninput=function(){ project.synthGen.note=this.value; };
  m.querySelectorAll('.sgSel').forEach(function(sel){ sel.onchange=function(){ project.synthGen.etats[this.dataset.k]=this.value; var eff=(this.value==='auto')?this.dataset.auto:this.value; var cell=this.closest('tr').querySelector('.sgEmo'); if(cell) cell.innerHTML=_emoP(eff); }; });
  if(typeof bindAccordions==='function') bindAccordions(m);
}

/* ====== ONGLETS ====== */
const TABS=[
  {id:'gen',label:'Généralités'},
  {id:'car',label:'Caractéristiques'},
  {id:'obl',label:'Obligations'},
  {id:'doc',label:'Documents'},
];
TABS.push({id:'synthg',label:'Synthèse générale'});
BDD.lots.forEach(L=> TABS.push({id:'lot:'+L.num,label:L.num+' '+shortLabel(L.label)}));
TABS.push({id:'p0',label:'Énergie & confort d\'été'});
TABS.push({id:'synth',label:'Synthèse'});
TABS.push({id:'chif',label:'Chiffrage'});
TABS.push({id:'plan',label:'Plan 10 ans'});
TABS.push({id:'concl',label:'Conclusion'});
function shortLabel(l){ return l.replace(' (SSI)','').replace(' (PMR)',''); }

let activeTab='gen';
function renderTabs(){
  const t=document.getElementById('tabs'); t.innerHTML='';
  TABS.forEach(tab=>{
    const d=document.createElement('div');
    d.className='tab'+(tab.id===activeTab?' active':'')+(tab.soon?' soon':'');
    d.textContent=tab.label; d.onclick=()=>{activeTab=tab.id;renderTabs();renderMain();};
    t.appendChild(d);
  });
}

/* ====== RENDU PRINCIPAL ====== */
function renderMain(){
  const m=document.getElementById('main'); m.innerHTML='';
  if(activeTab==='gen') return renderGen(m);
  if(activeTab==='car') return renderCar(m);
  if(activeTab.startsWith('lot:')) return renderLot(m, activeTab.slice(4));
  if(activeTab==='obl') return renderObl(m);
  if(activeTab==='doc') return renderDoc(m);
  if(activeTab==='p0') return renderP0(m);
  if(activeTab==='chif') return renderChif(m);
  if(activeTab==='plan') return renderPlan(m);
  if(activeTab==='synthg') return renderSynG(m);
  if(activeTab==='synth') return renderSyn(m);
  if(activeTab==='concl') return renderConcl(m);
  // onglets à venir
  const lab=(TABS.find(t=>t.id===activeTab)||{}).label||'';
  m.innerHTML='<div class="soon-box"><h2 style="color:var(--vf)">'+lab+'</h2>'
    +'<p>Onglet prévu dans la prochaine étape.<br>On valide d\'abord l\'interface de saisie des lots techniques.</p></div>';
}

/* --- Généralités --- */
function field(lab,obj,key,big){
  const id='f_'+Math.random().toString(36).slice(2);
  setTimeout(()=>{ const el=document.getElementById(id); if(el)el.oninput=()=>obj[key]=el.value; },0);
  const v=(obj[key]||'').replace(/"/g,'&quot;');
  if(big) return '<div class="field"><label>'+lab+'</label><textarea id="'+id+'">'+(obj[key]||'')+'</textarea></div>';
  return '<div class="field"><label>'+lab+'</label><input id="'+id+'" value="'+v+'"></div>';
}

/* --- Obligations / Documents / Chiffrage --- */
function dispoRow(r, withComment){
  const i1=rid(), i2=rid();
  setTimeout(()=>{ const s=document.getElementById(i1); if(s){ s.value=r.dispo||''; s.onchange=()=>r.dispo=s.value; }
    if(withComment){ const c=document.getElementById(i2); if(c){ c.value=r.comment||''; c.oninput=()=>r.comment=c.value; } } },0);
  return '<div class="dispo-row'+(withComment?'':' nc')+'"><div class="dl">'+escAttr(r.label)+'</div>'
    +'<select id="'+i1+'"><option value="">\u2014</option><option>Oui</option><option>Non</option></select>'
    +(withComment?'<input id="'+i2+'" placeholder="Commentaire">':'')+'</div>';
}
/* ============ IA (clé API utilisateur — appel Anthropic direct navigateur) ============ */
const AI={ key:'', model:'claude-sonnet-4-6', docs:[] };
function aiKeyBar(onRun, runLabel){
  const idK=rid(), idM=rid(), idF=rid(), idB=rid(), idS=rid();
  setTimeout(()=>{
    const k=document.getElementById(idK); if(k){ k.value=AI.key; k.oninput=()=>AI.key=k.value.trim(); }
    const mo=document.getElementById(idM); if(mo){ mo.value=AI.model; mo.onchange=()=>AI.model=mo.value; }
    const fi=document.getElementById(idF); if(fi){ fi.onchange=e=>readDocs(e.target.files, idS); }
    const b=document.getElementById(idB); if(b){ b.onclick=()=>onRun(b, idS); }
    const st=document.getElementById(idS); if(st) st.textContent = AI.docs.length? (AI.docs.length+' document(s) chargé(s)') : '';
  },0);
  return '<div class="ai-bar"><div class="ai-row">'
    +'<div class="field" style="flex:1"><label>Clé API Anthropic (sk-ant-…)</label><input id="'+idK+'" type="password" placeholder="sk-ant-..." autocomplete="off"></div>'
    +'<div class="field" style="max-width:240px"><label>Modèle</label><select id="'+idM+'"><option value="claude-sonnet-4-6">Claude Sonnet 4.6 (rapide)</option><option value="claude-opus-4-8">Claude Opus (qualité max)</option></select></div></div>'
    +'<div class="ai-row"><label class="btn-doc">📎 Documents sources<input id="'+idF+'" type="file" accept=".pdf,.png,.jpg,.jpeg,.txt" multiple style="display:none"></label>'
    +'<span id="'+idS+'" class="ai-status"></span>'
    +'<button id="'+idB+'" class="ai-go">🤖 '+escAttr(runLabel)+'</button></div>'
    +'<p class="hint">La clé reste locale (mémoire du navigateur, jamais enregistrée dans le fichier). Nécessite un accès réseau à api.anthropic.com.</p></div>';
}
function readDocs(files, statusId){
  AI.docs=[]; const arr=Array.from(files||[]); if(!arr.length){ const s=document.getElementById(statusId); if(s)s.textContent=''; return; }
  let done=0;
  arr.forEach(f=>{ const r=new FileReader(); r.onload=ev=>{
    const b64=String(ev.target.result).split(',')[1];
    let mt=f.type||(f.name.match(/\.pdf$/i)?'application/pdf':f.name.match(/\.png$/i)?'image/png':f.name.match(/\.txt$/i)?'text/plain':'image/jpeg');
    AI.docs.push({name:f.name, media_type:mt, data:b64, isText:mt==='text/plain'});
    done++; const s=document.getElementById(statusId); if(s)s.textContent=done+'/'+arr.length+' document(s) chargé(s)';
  }; r.readAsDataURL(f); });
}
function decodeB64Text(b64){ try{ return decodeURIComponent(escape(atob(b64))); }catch(e){ try{ return atob(b64); }catch(_){ return ''; } } }
function docBlocks(){
  return AI.docs.map(d=>{
    if(d.isText) return {type:'text', text:'--- '+d.name+' ---\n'+decodeB64Text(d.data)};
    if(d.media_type==='application/pdf') return {type:'document', source:{type:'base64', media_type:'application/pdf', data:d.data}};
    return {type:'image', source:{type:'base64', media_type:d.media_type, data:d.data}};
  });
}
async function callClaude(system, contentBlocks, maxTokens){
  if(!AI.key) throw new Error('Renseigne ta clé API Anthropic.');
  const res=await fetch('https://api.anthropic.com/v1/messages',{ method:'POST',
    headers:{'content-type':'application/json','x-api-key':AI.key,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
    body:JSON.stringify({model:AI.model||'claude-sonnet-4-5', max_tokens:maxTokens||1500, system:system, messages:[{role:'user', content:contentBlocks}]}) });
  let data; try{ data=await res.json(); }catch(e){ throw new Error('Réponse API illisible (HTTP '+res.status+').'); }
  if(!res.ok || data.error) throw new Error((data.error&&data.error.message)||('Erreur API HTTP '+res.status));
  return (data.content||[]).filter(b=>b.type==='text').map(b=>b.text).join('\n');
}
function parseJSON(txt){ let t=String(txt).trim().replace(/^```json\s*/i,'').replace(/^```\s*/,'').replace(/```$/,'').trim();
  const a=t.indexOf('{'), b=t.lastIndexOf('}'); if(a>=0&&b>a) t=t.slice(a,b+1); return JSON.parse(t); }
function normKey(s){ return String(s).toLowerCase().replace(/[^a-z0-9]/g,''); }
function matchLoose(obj, label){ const nl=normKey(label);
  for(const k in obj){ const nk=normKey(k); if(nk===nl||nk.indexOf(nl)>=0||nl.indexOf(nk)>=0) return obj[k]; } return null; }
function partie3Text(){
  let out='ÉTAT DES LIEUX (Partie 3) — données saisies :\n';
  BDD.lots.forEach(L=>{ const PL=project.lots[L.num]; if(!PL) return; let buf='\n## '+L.num+' '+L.label+'\n'; let has=false;
    Object.entries(PL.subs||{}).forEach(([nm,sub])=>{ const items=(sub.items||[]); if(!items.length && !(sub.remark||'').trim()) return; has=true;
      buf+='### '+nm+'\n'; if((sub.remark||'').trim()) buf+='Remarque: '+sub.remark.trim()+'\n';
      items.forEach(it=>{ const et={urgent:'U1/U2 (urgent)',moyen:'U3 (moyen terme)',bon:'U4 (bon état)'}[it.etat]||'non qualifié';
        buf+='- '+(it.desig||'(poste)')+' : '+(it.constat||'')+' | état: '+et+(it.preco?(' | préconisation: '+it.preco):'')+'\n'; }); });
    if(has) out+=buf; });
  return out;
}
function sumP(list){ return (list||[]).reduce((s,w)=>s+(w.ttc||0),0); }
async function runObligAI(btn, statusId){
  const old=btn.textContent; btn.textContent='⏳ Analyse…'; btn.disabled=true;
  try{
    const labels=[].concat(project.oblig.reglementaires.map(r=>r.label), project.oblig.contrats.map(r=>r.label));
    const sys="Tu es l'assistant d'un cabinet de diagnostic en copropriété. À partir des DOCUMENTS fournis, indique pour chaque LIBELLÉ s'il est disponible : \"Oui\", \"Non\", ou \"N/A\" si non applicable à cette copropriété. Rédige aussi un commentaire court (max 12 mots, en français). Réponds UNIQUEMENT par un objet JSON valide, sans aucun texte autour, de la forme {\"<libellé exact>\":{\"dispo\":\"Oui|Non|N/A\",\"comment\":\"...\"}}.";
    const blocks=[{type:'text', text:'Libellés à statuer :\n- '+labels.join('\n- ')+'\n\nLes documents sources de la copropriété suivent.'}].concat(docBlocks());
    const obj=parseJSON(await callClaude(sys, blocks, 1500));
    let n=0;
    [].concat(project.oblig.reglementaires, project.oblig.contrats).forEach(r=>{ const v=obj[r.label]||matchLoose(obj, r.label);
      if(v){ if(v.dispo){ r.dispo = v.dispo==='N/A' ? '' : v.dispo; } if(v.comment!=null) r.comment=v.comment; n++; } });
    renderMain(); alert('IA : '+n+' ligne(s) renseignée(s) sur '+labels.length+'.');
  }catch(e){ alert('IA — '+e.message); }
  btn.textContent=old; btn.disabled=false;
}
async function runSyntheseAI(btn, statusId, taId){
  const old=btn.textContent; btn.textContent='⏳ Rédaction…'; btn.disabled=true;
  try{
    const sys="Tu es ingénieur thermicien-pathologiste senior (cabinet ALHEE). Rédige la SYNTHÈSE de l'état des lieux (section 3.8) d'un DTG/PPT de copropriété, en français technique et concis. FORME imposée : une suite de sections thématiques, chacune = un intertitre en gras suivi de deux-points, puis un paragraphe. Utilise le markdown **Intertitre :** texte. Sections à couvrir si les données le permettent : État général de la copropriété, Voiries et espaces extérieurs, Espaces communs, Clos et couvert, Fluides et réseaux, Sécurité incendie, Accessibilité PMR. Base-toi STRICTEMENT sur l'état des lieux fourni (désordres, niveaux d'urgence U1–U4). Environ 250 à 350 mots. Pas d'introduction hors sujet, pas de liste à puces.";
    const blocks=[{type:'text', text:partie3Text()+'\n\nRédige maintenant la synthèse 3.8 selon la forme imposée.'}].concat(docBlocks());
    project.synthese=(await callClaude(sys, blocks, 1600)).trim();
    const ta=document.getElementById(taId); if(ta) ta.value=project.synthese;
  }catch(e){ alert('IA — '+e.message); }
  btn.textContent=old; btn.disabled=false;
}
async function runConclusionAI(btn, statusId, taId){
  const old=btn.textContent; btn.textContent='⏳ Rédaction…'; btn.disabled=true;
  try{
    const cp=computePlan(); const ch=computeChiffrage();
    const ctx='CARACTÉRISTIQUES GÉNÉRALES:\n'+JSON.stringify(project.car||{})+'\n\n'+partie3Text()
      +'\n\nSYNTHÈSE EXISTANTE:\n'+(project.synthese||'(non rédigée)')
      +'\n\nCHIFFRAGE par programme (€TTC) : P1='+sumP(ch.P1)+' | P2='+sumP(ch.P2)+' | P3='+sumP(ch.P3)+' | P0='+sumP(ch.P0)
      +'\nTOTAL plan 10 ans : '+cp.totals.grand+' €TTC';
    const sys="Tu es ingénieur (cabinet ALHEE). Rédige la CONCLUSION GÉNÉRALE (section 7) d'un DTG/PPT de copropriété, en français. Couvre : état général du bâti, performance énergétique (déperditions / classe DPE si pertinent), principaux désordres et leur urgence, scénarios et programme de travaux avec montants, priorisation sur 10 ans, recommandations finales. FORME : 4 à 6 paragraphes narratifs fluides, avec emphase en gras (**…**) sur les points clés (montants, classes DPE, urgences). Environ 300 à 400 mots. Pas de liste à puces.";
    const blocks=[{type:'text', text:ctx+'\n\nRédige maintenant la conclusion générale (section 7).'}].concat(docBlocks());
    project.conclusion=(await callClaude(sys, blocks, 1800)).trim();
    const ta=document.getElementById(taId); if(ta) ta.value=project.conclusion;
  }catch(e){ alert('IA — '+e.message); }
  btn.textContent=old; btn.disabled=false;
}
/* ============ Onglet PLAN 10 ANS (croix manuelles) ============ */
function effYears(it){ return (it.years==null)? [effAnnee(it)] : it.years; }
function updatePlanTotals(){ const el=document.getElementById('planTotals'); if(!el) return; const cp=computePlan();
  let h='<table class="chif-tbl"><thead><tr>'; for(let y=1;y<=10;y++) h+='<th>An '+y+'</th>'; h+='</tr></thead><tbody><tr>';
  for(let y=1;y<=10;y++) h+='<td class="num">'+fmtEur(cp.totals.years[y]||0)+'</td>'; h+='</tr></tbody></table>'
    +'<div class="totline" style="margin-top:10px"><b>TOTAL : '+fmtEur(cp.totals.grand)+' TTC</b></div>';
  el.innerHTML=h; }
function renderPlan(m){
  const flat=[];
  let html='<div class="lot-h">Plan pluriannuel sur 10 ans</div>'
   +'<p class="hint">Coche les années d\'intervention de chaque poste (croix de la partie 6 du PPT). Pré-rempli selon la priorité — ajuste librement.</p>';
  let any=false;
  BDD.lots.forEach(L=>{ const PL=project.lots[L.num]; if(!PL) return; const rows=[];
    Object.values(PL.subs||{}).forEach(sub=> (sub.items||[]).forEach(it=>{ if((it.preco||it.constat||'').trim()) rows.push(it); }));
    if(!rows.length) return; any=true;
    html+='<div class="subcard open"><div class="sh">'+escAttr(L.num+' '+shortLabel(L.label))+'<span class="chev"> </span></div><div class="sb"><div class="plan-wrap"><table class="plan-tbl"><thead><tr><th class="pt-trav">Travaux</th><th class="pt-loc">Localisation</th>';
    for(let y=1;y<=10;y++) html+='<th>'+y+'</th>'; html+='</tr></thead><tbody>';
    rows.forEach(it=>{ const trav=(it.preco||it.constat||'').trim(); const ys=effYears(it); const pi=flat.length; flat.push(it);
      html+='<tr><td class="pt-trav" title="'+escAttr(trav)+'">'+escAttr(trav.length>90?trav.slice(0,90)+'…':trav)+'</td><td>'+escAttr(it.loc||'')+'</td>';
      for(let y=1;y<=10;y++) html+='<td><input type="checkbox" data-pi="'+pi+'" data-y="'+y+'" '+(ys.indexOf(y)>=0?'checked':'')+'></td>';
      html+='</tr>'; });
    html+='</tbody></table></div></div></div>';
  });
  const _p0plan=p0Works();
  if(_p0plan.length){ any=true;
    html+='<div class="subcard open"><div class="sh">'+escAttr(P0_LABEL)+' <span class="badge">'+_p0plan.length+' poste(s)</span><span class="chev"> </span></div><div class="sb"><div class="plan-wrap"><table class="plan-tbl"><thead><tr><th class="pt-trav">Travaux P0</th><th class="pt-loc">Localisation</th>';
    for(let y=1;y<=10;y++) html+='<th>'+y+'</th>'; html+='</tr></thead><tbody>';
    _p0plan.forEach(w=>{ const yy=Number(w.annee)||1; const trav=(w.desig||'').trim();
      html+='<tr><td class="pt-trav" title="'+escAttr(trav)+'">'+escAttr(trav.length>90?trav.slice(0,90)+'…':trav)+'</td><td>'+escAttr(w.loc||'')+'</td>';
      for(let y=1;y<=10;y++) html+='<td>'+(y===yy?'✓':'')+'</td>';
      html+='</tr>'; });
    html+='</tbody></table></div><p class="hint" style="margin-top:8px">Ces postes se saisissent et se modifient dans l\'onglet « Énergie &amp; confort d\'été ».</p></div></div>';
  }
  if(!any) html+='<div class="soon-box"><p>Aucun poste de travaux saisi. Renseigne des postes dans les onglets des lots techniques (3.2 à 3.7) ou dans « Énergie & confort d\'été ».</p></div>';
  html+='<div class="subcard open"><div class="sh">Totaux par année<span class="chev"> </span></div><div class="sb"><div id="planTotals"></div></div></div>';
  m.innerHTML=html;
  m.querySelectorAll('.plan-tbl input[type=checkbox]').forEach(cb=>{ const it=flat[+cb.dataset.pi], y=+cb.dataset.y;
    cb.onchange=()=>{ if(it.years==null) it.years=effYears(it).slice(); const i=it.years.indexOf(y);
      if(cb.checked){ if(i<0) it.years.push(y); } else if(i>=0) it.years.splice(i,1); updatePlanTotals(); }; });
  updatePlanTotals(); bindAccordions(m);
}
/* ============ Onglets SYNTHÈSE & CONCLUSION ============ */
function renderSyn(m){
  const taId=rid();
  m.innerHTML='<div class="lot-h">Synthèse de l\'état des lieux (3.8)</div>'
   +aiKeyBar((b,sid)=>runSyntheseAI(b,sid,taId),'Générer la synthèse depuis l\'état des lieux')
   +'<div class="subcard open"><div class="sh">Texte de la synthèse<span class="chev"> </span></div><div class="sb">'
   +'<p class="hint">Format : <b>**Intertitre :**</b> texte (le gras est rendu dans le PPT). Entièrement modifiable.</p>'
   +'<textarea id="'+taId+'" class="ai-ta" placeholder="Clique sur Générer, ou rédige directement ici…"></textarea></div></div>';
  setTimeout(()=>{ const ta=document.getElementById(taId); if(ta){ ta.value=project.synthese||''; ta.oninput=()=>project.synthese=ta.value; } },0);
  bindAccordions(m);
}
function renderConcl(m){
  const taId=rid();
  m.innerHTML='<div class="lot-h">Conclusion générale (7)</div>'
   +aiKeyBar((b,sid)=>runConclusionAI(b,sid,taId),'Générer la conclusion générale')
   +'<div class="subcard open"><div class="sh">Texte de la conclusion<span class="chev"> </span></div><div class="sb">'
   +'<p class="hint">Format narratif ; emphase en <b>**gras**</b> (rendue dans le PPT). Entièrement modifiable.</p>'
   +'<textarea id="'+taId+'" class="ai-ta" placeholder="Clique sur Générer, ou rédige directement ici…"></textarea></div></div>';
  setTimeout(()=>{ const ta=document.getElementById(taId); if(ta){ ta.value=project.conclusion||''; ta.oninput=()=>project.conclusion=ta.value; } },0);
  bindAccordions(m);
}
function renderObl(m){
  const o=project.oblig;
  m.innerHTML='<div class="lot-h">Obligations r\u00e9glementaires et entretien</div>'
   +aiKeyBar((b,sid)=>runObligAI(b,sid),'Remplir disponibilit\u00e9s & commentaires par IA')
   +'<div class="subcard open"><div class="sh">Obligations r\u00e9glementaires<span class="chev"> </span></div><div class="sb">'
     +'<div class="dispo-head"><span>Document</span><span>Disponibilit\u00e9</span><span>Commentaire</span></div>'
     +o.reglementaires.map(r=>dispoRow(r,true)).join('')+'</div></div>'
   +'<div class="subcard open"><div class="sh">Contrats d\'entretien<span class="chev"> </span></div><div class="sb">'
     +'<div class="dispo-head"><span>Document</span><span>Disponibilit\u00e9</span><span>Commentaire</span></div>'
     +o.contrats.map(r=>dispoRow(r,true)).join('')+'</div></div>';
  bindAccordions(m);
}
function renderDoc(m){
  const d=project.docs;
  m.innerHTML='<div class="lot-h">Documents r\u00e9cup\u00e9r\u00e9s</div>'
   +'<div class="subcard open"><div class="sh">Documents administratifs<span class="chev"> </span></div><div class="sb">'
     +'<div class="dispo-head nc"><span>Document</span><span>Disponibilit\u00e9</span></div>'
     +d.admin.map(r=>dispoRow(r,false)).join('')+'</div></div>'
   +'<div class="subcard open"><div class="sh">Plans<span class="chev"> </span></div><div class="sb">'
     +'<div class="dispo-head nc"><span>Document</span><span>Disponibilit\u00e9</span></div>'
     +d.plans.map(r=>dispoRow(r,false)).join('')+'</div></div>';
  bindAccordions(m);
}
function effPrio(it){ if(it.prio) return it.prio; const c=(it.cat||''); if(c.indexOf('nerg')>=0) return 'P0'; if(it.etat==='urgent') return 'P1'; if(it.etat==='moyen') return 'P2'; return 'P3'; }
function effAnnee(it){ if(it.annee) return Number(it.annee); const p=effPrio(it); return p==='P1'?1:p==='P2'?3:p==='P3'?6:2; }
function iterItems(cb){ Object.entries(project.lots||{}).forEach(([num,L])=>{ Object.values(L.subs||{}).forEach(sub=>{ (sub.items||[]).forEach(it=>cb(it,num)); }); }); }
const P0_LABEL='Performance énergétique & confort d\'été';
function p0Works(){ return (project.p0||[]).filter(w=>(w.desig||'').trim()); }
function computePreco(){ const byLot={};
  iterItems((it,num)=>{ const trav=(it.preco||it.constat||'').trim(); if(!trav) return; (byLot[num]=byLot[num]||[]).push({travaux:trav, loc:it.loc||'', prio:effPrio(it)}); });
  const groups=BDD.lots.filter(L=>byLot[L.num]).map(L=>({label:L.num+' '+shortLabel(L.label), rows:byLot[L.num]}));
  const p0=p0Works().map(w=>({travaux:w.desig.trim(), loc:w.loc||'', prio:'P0'}));
  if(p0.length) groups.push({label:P0_LABEL, rows:p0});
  return groups; }
function computePlan(){ const tva=Number(project.tva||10)/100; const byLot={}; const years={};
  iterItems((it,num)=>{ const trav=(it.preco||it.constat||'').trim(); if(!trav) return; const ys=effYears(it); const ttc=Math.round(Number(it.total||0)*(1+tva));
    (byLot[num]=byLot[num]||[]).push({travaux:trav, loc:it.loc||'', years:ys, ttc}); ys.forEach(y=>{ years[y]=(years[y]||0)+ttc; }); });
  const plan=BDD.lots.filter(L=>byLot[L.num]).map(L=>({label:L.num+' '+shortLabel(L.label), rows:byLot[L.num]}));
  const p0rows=p0Works().map(w=>{ const y=Number(w.annee)||1; const ttc=Math.round(Number(w.total||0)*(1+tva)); years[y]=(years[y]||0)+ttc; return {travaux:w.desig.trim(), loc:w.loc||'', years:[y], ttc}; });
  if(p0rows.length) plan.push({label:P0_LABEL, rows:p0rows});
  const grand=Object.values(years).reduce((a,b)=>a+b,0);
  return { plan, totals:{years, grand} }; }
function computeChiffrage(){
  const tva=Number(project.tva||10)/100; const g={P0:[],P1:[],P2:[],P3:[]};
  iterItems(it=>{ const trav=(it.preco||it.constat||'').trim(); if(!trav) return;
    const ht=Number(it.total||0); g[effPrio(it)].push({travaux:trav, loc:it.loc||'', ht, ttc:Math.round(ht*(1+tva))}); });
  p0Works().forEach(w=>{ const ht=Number(w.total||0); g.P0.push({travaux:w.desig.trim(), loc:w.loc||'', ht, ttc:Math.round(ht*(1+tva))}); });
  return g;
}
function renderChif(m){
  const g=computeChiffrage();
  const names={P1:'Programme \u00e0 court terme (P1)',P2:'Programme \u00e0 moyen terme (P2)',P3:'Programme \u00e0 long terme (P3)',P0:'Performance \u00e9nerg\u00e9tique & confort (P0)'};
  let html='<div class="lot-h">Chiffrage \u2014 programme de travaux</div>'
   +'<div class="subcard open"><div class="sh">Param\u00e8tres<span class="chev"> </span></div><div class="sb">'
   +'<div class="field" style="max-width:200px"><label>TVA appliqu\u00e9e (%)</label><input id="tvaIn" type="number" value="'+(project.tva||10)+'"></div>'
   +'<p class="hint">Chiffrage calcul\u00e9 automatiquement \u00e0 partir des postes saisis dans les lots (P.U. \u00d7 quantit\u00e9), r\u00e9partis par programme selon la <b>priorit\u00e9</b> choisie par poste (ou, \u00e0 d\u00e9faut, d\u00e9duite : \ud83d\udea8 \u2192 P1, \ud83d\udd27 \u2192 P2, \u2713 \u2192 P3, \u26a1 \u2192 P0). Montants TTC = HT \u00d7 (1 + TVA).</p></div></div>';
  let grand=0;
  ['P1','P2','P3','P0'].forEach(p=>{ const list=g[p]; const tot=list.reduce((s,w)=>s+(w.ttc||0),0); grand+=tot;
    html+='<div class="subcard open"><div class="sh">'+names[p]+' <span class="catbadge c-sauv" style="margin-left:8px">'+fmtEur(tot)+' TTC</span><span class="chev"> </span></div><div class="sb">';
    if(!list.length) html+='<p class="hint">Aucun poste pour ce programme.</p>';
    else html+='<table class="chif-tbl"><thead><tr><th>Travaux</th><th>Localisation</th><th>Co\u00fbt \u20acTTC</th></tr></thead><tbody>'
      +list.map(w=>'<tr><td>'+escAttr(w.travaux)+'</td><td>'+escAttr(w.loc||'')+'</td><td class="num">'+(w.ht?fmtEur(w.ttc):'\u00e0 chiffrer')+'</td></tr>').join('')
      +'<tr class="tot"><td>TOTAL</td><td></td><td class="num">'+fmtEur(tot)+'</td></tr></tbody></table>';
    html+='</div></div>';
  });
  html+='<div class="totline" style="font-size:15px;margin-top:14px"><b>TOTAL G\u00c9N\u00c9RAL : '+fmtEur(grand)+' TTC</b></div>';
  m.innerHTML=html;
  const tv=document.getElementById('tvaIn'); if(tv) tv.onchange=()=>{ project.tva=Number(tv.value)||10; renderMain(); };
  bindAccordions(m);
}

/* --- Onglet P0 : performance énergétique & confort d'été (travaux proactifs) --- */
function renderP0(m){
  if(!Array.isArray(project.p0)) project.p0=[];
  m.innerHTML='';
  const h=document.createElement('div'); h.className='lot-h'; h.textContent='Performance énergétique & confort d\'été'; m.appendChild(h);
  const hint=document.createElement('div'); hint.className='hint';
  hint.innerHTML='Listez ici les <b>travaux P0</b> (performance énergétique, confort d\'été) — des préconisations <b>proactives</b>, indépendantes des anomalies constatées. Ils alimentent automatiquement le <b>chiffrage (programme P0)</b>, le <b>plan pluriannuel</b> et les <b>préconisations</b> du PowerPoint.';
  m.appendChild(hint);
  const card=document.createElement('div'); card.className='subcard open';
  const sh=document.createElement('div'); sh.className='sh'; sh.innerHTML='Travaux P0 à prévoir <span class="badge">'+project.p0.length+' poste(s)</span><span class="chev"> </span>';
  const body=document.createElement('div'); body.className='sb';
  project.p0.forEach((w,idx)=> body.appendChild(buildP0Row(w,idx)));
  const add=document.createElement('button'); add.className='addbtn'; add.textContent='+ Ajouter un travail P0';
  add.onclick=()=>{ project.p0.push({desig:'',loc:'',pu:'',unite:'',qty:'',total:0,annee:''}); if(typeof _scheduleAutosave==='function') _scheduleAutosave(); renderMain(); };
  body.appendChild(add);
  card.appendChild(sh); card.appendChild(body); m.appendChild(card);
  const tva=Number(project.tva||10)/100;
  const totHT=project.p0.reduce((s,w)=>s+Number(w.total||0),0);
  const recap=document.createElement('div'); recap.className='totline'; recap.style.cssText='font-size:15px;margin-top:14px';
  recap.innerHTML='<b>Total P0 : '+fmtEur(totHT)+' HT · '+fmtEur(Math.round(totHT*(1+tva)))+' TTC</b>';
  m.appendChild(recap);
  bindAccordions(m);
}
function buildP0Row(w,idx){
  const a=document.createElement('div'); a.className='anom';
  const del=document.createElement('button'); del.className='del'; del.textContent='✕';
  del.onclick=()=>{ project.p0.splice(idx,1); if(typeof _scheduleAutosave==='function') _scheduleAutosave(); renderMain(); }; a.appendChild(del);
  const fD=document.createElement('div'); fD.className='field'; fD.innerHTML='<label>Travaux / préconisation</label>';
  const taD=document.createElement('textarea'); taD.value=w.desig||''; taD.placeholder='ex : Isolation des combles perdus par soufflage'; taD.oninput=()=>w.desig=taD.value; fD.appendChild(taD); a.appendChild(fD);
  const fr=document.createElement('div'); fr.className='inline3';
  const puF=document.createElement('div'); puF.className='field'; puF.innerHTML='<label>P.U. € HT</label>';
  const inPU=document.createElement('input'); inPU.type='number'; inPU.value=(w.pu!=null?w.pu:''); inPU.oninput=()=>{ w.pu=inPU.value; rec(); }; puF.appendChild(inPU);
  const uF=document.createElement('div'); uF.className='field'; uF.innerHTML='<label>Unité</label>';
  const uEl=document.createElement('input'); uEl.value=w.unite||''; uEl.placeholder='U / m² / ml'; uEl.oninput=()=>w.unite=uEl.value; uF.appendChild(uEl);
  const qF=document.createElement('div'); qF.className='field'; qF.innerHTML='<label>Quantité</label>';
  const inQ=document.createElement('input'); inQ.type='number'; inQ.value=w.qty||''; inQ.placeholder='ex: 1'; inQ.oninput=()=>{ w.qty=inQ.value; rec(); }; qF.appendChild(inQ);
  fr.appendChild(puF); fr.appendChild(uF); fr.appendChild(qF); a.appendChild(fr);
  const totF=document.createElement('div'); totF.className='totline';
  function rec(){ const pu=Number(w.pu||0), q=Number(w.qty||0), t=pu*q; w.total=t;
    totF.innerHTML='Total estimé : <b>'+(t?fmtEur(t):'—')+'</b>'+((pu&&w.unite)?' <span class="sub">('+fmtEur(pu)+' / '+escAttr(w.unite)+' × '+(w.qty||0)+')</span>':''); }
  rec(); a.appendChild(totF);
  const fr2=document.createElement('div'); fr2.className='inline2';
  const locF=document.createElement('div'); locF.className='field'; locF.innerHTML='<label>Localisation</label>';
  const inLoc=document.createElement('input'); inLoc.value=w.loc||''; inLoc.placeholder='ex: Toiture, façades…'; inLoc.oninput=()=>w.loc=inLoc.value; locF.appendChild(inLoc);
  const anF=document.createElement('div'); anF.className='field'; anF.innerHTML='<label>Année (plan)</label>';
  const selAnnee=document.createElement('select');
  { const op=document.createElement('option'); op.value=''; op.textContent='Année 1 (défaut)'; selAnnee.appendChild(op); }
  for(let y=1;y<=10;y++){ const op=document.createElement('option'); op.value=String(y); op.textContent='Année '+y; selAnnee.appendChild(op); }
  selAnnee.value=w.annee||''; selAnnee.onchange=()=>w.annee=selAnnee.value; anF.appendChild(selAnnee);
  fr2.appendChild(locF); fr2.appendChild(anF); a.appendChild(fr2);
  return a;
}

function renderGen(m){
  const g=project.meta;
  m.innerHTML='<div class="lot-h">Généralités</div>'
   +'<div class="subcard open"><div class="sh">Syndic & destinataire<span class="chev"> </span></div><div class="sb"><div class="grid2">'
   +field('Nom de la copropriété',g,'copro')+field('Typologie',g,'typologie')
   +field('Nom du syndic',g,'syndic')+field('Adresse du syndic',g,'adresseSyndic')
   +field('Destinataire',g,'destinataire')+field('Fonction',g,'fonction')
   +'</div></div></div>'
   +'<div class="subcard open"><div class="sh">Bureau d\'études<span class="chev"> </span></div><div class="sb"><div class="grid2">'
   +field("Numéro d'affaires",g,'affaire')+field('Auditeur',g,'auditeur')
   +field('Relecture',g,'relecture')+field('Date de visite',g,'dateVisite')
   +field('Date de restitution',g,'dateResti')
   +'</div></div></div>';
  bindAccordions(m);
}
function renderCar(m){
  const c=project.car;
  m.innerHTML='<div class="lot-h">Caractéristiques générales</div>'
   +'<div class="subcard open"><div class="sh">Site<span class="chev"> </span></div><div class="sb"><div class="grid2">'
   +field('Adresse',c,'adresse')+field('Année de construction',c,'annee')
   +field('Nombre de bâtiments',c,'nbBat')+field('Nombre de lots',c,'nbLots')
   +field('Nombre de niveaux',c,'nbNiv')+field('Surface',c,'surface')
   +'</div></div></div>'
   +'<div class="subcard open"><div class="sh">Caractéristiques énergétiques<span class="chev"> </span></div><div class="sb"><div class="grid2">'
   +field('Type de chauffage',c,'chauffage')+field('Énergie chauffage',c,'enChauf')
   +field('Type ECS',c,'ecs')+field('Énergie ECS',c,'enEcs')
   +'</div></div></div>';
  bindAccordions(m);
}

/* --- Lot technique (cœur : sélection BDD) --- */
function renderLot(m, num){
  const L=BDD.lots.find(x=>x.num===num); if(!L)return;
  m.innerHTML='<div class="lot-h">'+L.num+' '+L.label+'</div>'
    +'<div class="hint">Pour chaque sous-lot : renseignez la <b>remarque générale</b>, puis ajoutez les anomalies. '
    +'Choisissez la <b>désignation</b> puis l\'<b>anomalie</b> → le constat se remplit automatiquement (modifiable). '
    +'La <b>recommandation</b>, la <b>catégorie</b> et le <b>prix unitaire</b> (selon la zone) se remplissent alors automatiquement ; ajustez le P.U., l\'<b>unité</b> et la <b>quantité</b> pour obtenir le total, puis l\'<b>état</b> et la <b>photo</b>.</div>';
  L.subs.forEach(sub=>{
    const st=project.lots[num].subs[sub.name];
    const card=document.createElement('div'); card.className='subcard'+(st.items.length?' open':'');
    card.innerHTML='<div class="sh"><span class="chev">▶</span> '+sub.name+'<span class="badge">'+st.items.length+' anomalie(s)</span></div>';
    const sb=document.createElement('div'); sb.className='sb';
    // remarque
    const rf=document.createElement('div'); rf.className='field';
    rf.innerHTML='<label>Remarque générale</label>';
    const ta=document.createElement('textarea'); ta.value=st.remark||''; ta.oninput=()=>st.remark=ta.value;
    rf.appendChild(ta); sb.appendChild(rf);
    // liste anomalies
    const list=document.createElement('div');
    st.items.forEach(it=> list.appendChild(anomCard(L,sub,st,it,card)));
    sb.appendChild(list);
    const add=document.createElement('button'); add.className='addbtn'; add.textContent='+ Ajouter une anomalie';
    add.onclick=()=>{ const it={desig:'',desigFree:'',anom:'',constat:'',preco:'',cat:'',unite:'',pP:null,pV:null,pu:'',qty:'',total:0,etat:null,photo:null};
      st.items.push(it); list.appendChild(anomCard(L,sub,st,it,card)); card.classList.add('open'); updateBadge(card,st); };
    sb.appendChild(add);
    card.appendChild(sb);
    card.querySelector('.sh').onclick=(e)=>{ if(e.target.closest('.sb'))return; card.classList.toggle('open'); };
    m.appendChild(card);
  });
}
function updateBadge(card,st){ card.querySelector('.badge').textContent=st.items.length+' anomalie(s)'; }

function anomCard(L,sub,st,it,card){
  const a=document.createElement('div'); a.className='anom';
  const del=document.createElement('button'); del.className='del'; del.textContent='✕';
  del.onclick=()=>{ const i=st.items.indexOf(it); if(i>=0)st.items.splice(i,1); a.remove(); updateBadge(card,st); };
  a.appendChild(del);

  const desigNames=Object.keys(sub.designations||{});
  // --- ligne 1 : désignation + observation ---
  const top=document.createElement('div'); top.className='top';
  const fD=document.createElement('div'); fD.className='field'; fD.innerHTML='<label>Désignation</label>';
  const selD=document.createElement('select');
  selD.innerHTML='<option value="">— choisir —</option>'+desigNames.map(d=>'<option>'+escAttr(d)+'</option>').join('')+'<option value="__autre">Autre…</option>';
  selD.value=it.desig||''; fD.appendChild(selD);
  const inD=document.createElement('input'); inD.placeholder='Désignation libre'; inD.style.marginTop='6px';
  inD.style.display = it.desig==='__autre'?'block':'none'; inD.value = it.desig==='__autre'?(it.desigFree||''):'';
  inD.oninput=()=>{ it.desigFree=inD.value; }; fD.appendChild(inD);
  top.appendChild(fD);
  const fA=document.createElement('div'); fA.className='field'; fA.innerHTML='<label>Anomalie / observation</label>';
  const selA=document.createElement('select'); fA.appendChild(selA);
  const inA=document.createElement('input'); inA.placeholder='Observation libre'; inA.style.marginTop='6px'; inA.style.display='none';
  inA.oninput=()=>{ it.anom=inA.value; }; fA.appendChild(inA);
  top.appendChild(fA);
  a.appendChild(top);

  function fillAnomOptions(){
    const arr=(sub.designations||{})[selD.value]||[];
    selA.innerHTML='<option value="">— choisir —</option>'+arr.map((x,i)=>'<option value="'+i+'">'+escAttr(x.l)+'</option>').join('')+'<option value="__autre">Autre…</option>';
  }
  selD.onchange=()=>{ it.desig=selD.value; inD.style.display=selD.value==='__autre'?'block':'none';
    if(selD.value==='__autre'){ selA.innerHTML='<option value="">(saisie libre)</option>'; inA.style.display='block'; }
    else { fillAnomOptions(); inA.style.display='none'; } };
  selA.onchange=()=>{
    if(selA.value==='__autre'){ inA.style.display='block'; it.anom=''; return; }
    inA.style.display='none';
    const arr=(sub.designations||{})[selD.value]||[]; const x=arr[+selA.value];
    if(x){ it.anom=x.l; taC.value=x.c; it.constat=x.c;
      taW.value=x.r||''; it.preco=x.r||'';
      it.cat=x.cat||''; it.unite=x.u||''; it.pP=(x.pP==null?null:x.pP); it.pV=(x.pV==null?null:x.pV); it.pu='';
      catBadge.textContent=it.cat||'—'; catBadge.className='catbadge '+catClass(it.cat);
      unitEl.value=it.unite||''; inPU.value=''; inPU.placeholder=(defaultPU()===''?'—':defaultPU()+''); recompute();
    }
  };
  if(it.desig && it.desig!=='__autre'){ fillAnomOptions(); }
  else if(it.desig==='__autre'){ selA.innerHTML='<option value="">(saisie libre)</option>'; inA.style.display='block'; inA.value=it.anom||''; }

  // --- constat ---
  const fC=document.createElement('div'); fC.className='field'; fC.innerHTML='<label>Constat</label>';
  const taC=document.createElement('textarea'); taC.value=it.constat||''; taC.oninput=()=>it.constat=taC.value;
  fC.appendChild(taC); a.appendChild(fC);

  // --- bloc bas : gauche (photo+état), droite (reco + catégorie + prix) ---
  const row=document.createElement('div'); row.className='row-pe';
  const left=document.createElement('div');
  const box=document.createElement('div'); box.className='photobox';
  const refresh=()=>box.innerHTML = it.photo? '<img src="'+it.photo+'">':'📷<br>Photo';
  refresh();
  const fin=document.createElement('input'); fin.type='file'; fin.accept='image/*'; box.onclick=()=>fin.click();
  fin.onchange=e=>{ const f=e.target.files[0]; if(!f)return; box.innerHTML='<span style="font-size:11px">⏳ Traitement…</span>';
    compressImageFile(f).then(d=>{ it.photo=d; refresh(); if(typeof _scheduleAutosave==='function') _scheduleAutosave(); })
      .catch(()=>{ refresh(); }); };
  left.appendChild(box); left.appendChild(fin);
  const ep=document.createElement('div'); ep.className='etat-pick';
  ['bon','moyen','urgent','na'].forEach(k=>{ const o=document.createElement('div');
    o.className='opt'+(k==='na'?' na':'')+(it.etat===k?' sel':''); o.innerHTML=k==='na'?'—':smiley(k,30);
    o.title=k; o.onclick=()=>{ it.etat=k; ep.querySelectorAll('.opt').forEach(x=>x.classList.remove('sel')); o.classList.add('sel'); };
    ep.appendChild(o); });
  left.appendChild(ep);
  row.appendChild(left);

  const right=document.createElement('div');
  // recommandation
  const fW=document.createElement('div'); fW.className='field'; fW.innerHTML='<label>Entretien / travaux à prévoir (recommandation)</label>';
  const taW=document.createElement('textarea'); taW.value=it.preco||''; taW.oninput=()=>it.preco=taW.value; fW.appendChild(taW);
  right.appendChild(fW);
  // catégorie
  const fCat=document.createElement('div'); fCat.className='field'; fCat.innerHTML='<label>Catégorie de travaux</label>';
  const catBadge=document.createElement('span'); catBadge.className='catbadge '+catClass(it.cat); catBadge.textContent=it.cat||'—';
  fCat.appendChild(catBadge); right.appendChild(fCat);
  // prix
  function defaultPU(){ const v = project.zone==='prov'? it.pV : it.pP; return (v==null)?'':v; }
  function puEff(){ const base=(it.pu!==''&&it.pu!=null)? it.pu : defaultPU(); const n=Number(base); return isFinite(n)?n:0; }
  const zoneLbl = (project.zone==='prov'?'province':'Paris');
  const fr=document.createElement('div'); fr.className='inline3';
  const puF=document.createElement('div'); puF.className='field'; puF.innerHTML='<label>P.U. € HT ('+zoneLbl+')</label>';
  const inPU=document.createElement('input'); inPU.type='number'; inPU.value=(it.pu!=null?it.pu:''); inPU.placeholder=(defaultPU()===''?'—':defaultPU()+'');
  inPU.oninput=()=>{ it.pu=inPU.value; recompute(); }; puF.appendChild(inPU);
  const uF=document.createElement('div'); uF.className='field'; uF.innerHTML='<label>Unité</label>';
  const unitEl=document.createElement('input'); unitEl.value=it.unite||''; unitEl.placeholder='U / m² / ml'; unitEl.oninput=()=>it.unite=unitEl.value; uF.appendChild(unitEl);
  const qF=document.createElement('div'); qF.className='field'; qF.innerHTML='<label>Quantité</label>';
  const inQ=document.createElement('input'); inQ.type='number'; inQ.value=it.qty||''; inQ.placeholder='ex: 36'; inQ.oninput=()=>{ it.qty=inQ.value; recompute(); }; qF.appendChild(inQ);
  fr.appendChild(puF); fr.appendChild(uF); fr.appendChild(qF); right.appendChild(fr);
  // total
  const totF=document.createElement('div'); totF.className='totline';
  function recompute(){ const pu=puEff(); const q=Number(it.qty||0); const t=pu*q; it.total=t;
    totF.innerHTML='Total estimé : <b>'+(t?fmtEur(t):'—')+'</b>'+((pu&&it.unite)?' <span class="sub">('+fmtEur(pu)+' / '+escAttr(it.unite)+' × '+(it.qty||0)+')</span>':''); }
  recompute();
  right.appendChild(totF);
  // localisation / priorité / année (Partie 4, Chiffrage, Plan 10 ans)
  const fr2=document.createElement('div'); fr2.className='inline3';
  const locF=document.createElement('div'); locF.className='field'; locF.innerHTML='<label>Localisation</label>';
  const inLoc=document.createElement('input'); inLoc.value=it.loc||''; inLoc.placeholder="ex: Cage d'escalier"; inLoc.oninput=()=>it.loc=inLoc.value; locF.appendChild(inLoc);
  const prioF=document.createElement('div'); prioF.className='field'; prioF.innerHTML='<label>Priorité</label>';
  const selPrio=document.createElement('select');
  [['','Auto'],['P1','P1 \u2014 court'],['P2','P2 \u2014 moyen'],['P3','P3 \u2014 long'],['P0','P0 \u2014 énergie']].forEach(o=>{ const op=document.createElement('option'); op.value=o[0]; op.textContent=o[1]; selPrio.appendChild(op); });
  selPrio.value=it.prio||''; selPrio.onchange=()=>it.prio=selPrio.value; prioF.appendChild(selPrio);
  const anF=document.createElement('div'); anF.className='field'; anF.innerHTML='<label>Année (plan)</label>';
  const selAnnee=document.createElement('select');
  { const op=document.createElement('option'); op.value=''; op.textContent='Auto'; selAnnee.appendChild(op); }
  for(let y=1;y<=10;y++){ const op=document.createElement('option'); op.value=String(y); op.textContent='Année '+y; selAnnee.appendChild(op); }
  selAnnee.value=it.annee||''; selAnnee.onchange=()=>it.annee=selAnnee.value; anF.appendChild(selAnnee);
  fr2.appendChild(locF); fr2.appendChild(prioF); fr2.appendChild(anF); right.appendChild(fr2);
  row.appendChild(right);
  a.appendChild(row);
  return a;
}
function catClass(c){ if(!c)return ''; if(c.indexOf('nerg')>=0)return 'c-energie'; if(c.indexOf('Sécur')>=0||c.indexOf('Santé')>=0)return 'c-secu'; return 'c-sauv'; }
function fmtEur(n){ n=Math.round(Number(n)||0); return n.toLocaleString('fr-FR')+' €'; }

/* ====== HELPERS ====== */
function escAttr(s){ return (s==null?'':String(s)).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;'); }
function trim(s,n){ s=String(s||''); return s.length>n?s.slice(0,n)+'…':s; }
function bindAccordions(scope){ scope.querySelectorAll('.subcard .sh').forEach(sh=>{ sh.onclick=()=>sh.parentElement.classList.toggle('open'); }); }

/* ====== SAUVEGARDE / CHARGEMENT ====== */
(function(){var z=document.getElementById('zoneSel'); if(z){ z.value=project.zone||'paris'; z.onchange=function(){ project.zone=z.value; renderMain(); }; } })();
document.getElementById('btnBackup').onclick=()=>{
  const blob=new Blob([JSON.stringify(project)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  const nm=(project.meta.copro||'projet').replace(/[^a-z0-9]+/gi,'_').toLowerCase();
  a.download='PPT_'+nm+'.json'; a.click();
};
function mergeProject(loaded){ const base=newProject();
  base.meta=loaded.meta||{}; base.car=loaded.car||{};
  base.zone=loaded.zone||base.zone; base.tva=(loaded.tva!=null?loaded.tva:base.tva);
  base.synthese=loaded.synthese||''; base.conclusion=loaded.conclusion||'';
  if(loaded.oblig) base.oblig=loaded.oblig; if(loaded.docs) base.docs=loaded.docs;
  if(loaded.synthGen) base.synthGen=loaded.synthGen;
  if(Array.isArray(loaded.p0)) base.p0=loaded.p0;
  if(loaded.lots) for(const k in loaded.lots) if(base.lots[k]){ for(const s in loaded.lots[k].subs||{}) if(base.lots[k].subs[s]) base.lots[k].subs[s]=loaded.lots[k].subs[s]; }
  return base; }

/* ====== SAUVEGARDE LOCALE NAVIGATEUR (IndexedDB) ====== */
const _IDB_DB='alhee_ppt',_IDB_STORE='kv',_IDB_KEY='project';
function _idbOpen(){return new Promise((res,rej)=>{const r=indexedDB.open(_IDB_DB,1);r.onupgradeneeded=()=>{r.result.createObjectStore(_IDB_STORE);};r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);});}
async function _idbSet(k,v){const db=await _idbOpen();return new Promise((res,rej)=>{const t=db.transaction(_IDB_STORE,'readwrite');t.objectStore(_IDB_STORE).put(v,k);t.oncomplete=()=>res();t.onerror=()=>rej(t.error);});}
async function _idbGet(k){const db=await _idbOpen();return new Promise((res,rej)=>{const t=db.transaction(_IDB_STORE,'readonly');const rq=t.objectStore(_IDB_STORE).get(k);rq.onsuccess=()=>res(rq.result);rq.onerror=()=>rej(rq.error);});}
let _savedSnap='';
function _isDirty(){ try{ return JSON.stringify(project)!==_savedSnap; }catch(e){ return false; } }
function _updateStatus(){ const el=document.getElementById('saveStatus'); if(!el)return; if(_isDirty()){ el.textContent='\u25CF Enregistrement\u2026'; el.style.color='#e8c07a'; } else { el.textContent='\u2713 Sauvegarde auto'; el.style.color='#9fe0a0'; } }
let _statusTimer=null; function _scheduleStatus(){ clearTimeout(_statusTimer); _statusTimer=setTimeout(_updateStatus,400); }
function _toast(msg){ let t=document.getElementById('_toast'); if(!t){ t=document.createElement('div'); t.id='_toast'; t.style.cssText='position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#18483C;color:#fff;padding:10px 18px;border-radius:10px;font-size:14px;box-shadow:0 6px 18px rgba(0,0,0,.3);z-index:99999;opacity:0;transition:opacity .2s;pointer-events:none'; document.body.appendChild(t);} t.textContent=msg; t.style.opacity='1'; clearTimeout(t._h); t._h=setTimeout(()=>{t.style.opacity='0';},1900); }
async function saveLocal(){ try{ const s=JSON.stringify(project); await _idbSet(_IDB_KEY,s); _savedSnap=s; _updateStatus(); _toast('\uD83D\uDCBE Mission sauvegard\u00E9e'); }catch(e){ alert('Sauvegarde impossible : '+(e&&e.message||e)); } }
async function loadLocal(initial){
  try{ const s=await _idbGet(_IDB_KEY);
    if(!s){ if(!initial) alert('Aucune sauvegarde trouv\u00E9e dans ce navigateur.'); return false; }
    if(!initial && _isDirty() && !confirm('Charger la derni\u00E8re sauvegarde ?\nLes modifications non enregistr\u00E9es seront perdues.')) return false;
    project=mergeProject(JSON.parse(s)); _savedSnap=JSON.stringify(project);
    if(!initial){ activeTab='gen'; renderTabs(); renderMain(); _updateStatus(); _toast('\u21A9\uFE0E Derni\u00E8re sauvegarde charg\u00E9e'); }
    return true;
  }catch(e){ if(!initial) alert('Chargement impossible : '+(e&&e.message||e)); return false; }
}
function newMission(){
  if(_isDirty() && !confirm('D\u00E9marrer une nouvelle mission ?\nLes modifications non enregistr\u00E9es seront perdues.')) return;
  project=newProject(); _savedSnap=JSON.stringify(project); activeTab='gen';
  renderTabs(); renderMain(); _updateStatus(); _toast('\u2795 Nouvelle mission');
}
let _autosaveTimer=null;
function _scheduleAutosave(){ clearTimeout(_autosaveTimer); _autosaveTimer=setTimeout(async()=>{ try{ const s=JSON.stringify(project); if(s===_savedSnap){ _updateStatus(); return; } await _idbSet(_IDB_KEY,s); _savedSnap=s; _updateStatus(); }catch(e){} },800); }
document.addEventListener('input',_scheduleAutosave,true);
document.addEventListener('change',_scheduleAutosave,true);
setInterval(function(){ try{ const s=JSON.stringify(project); if(s!==_savedSnap){ _idbSet(_IDB_KEY,s).then(function(){ _savedSnap=s; _updateStatus(); }).catch(function(){}); } }catch(e){} },5000);
window.addEventListener('beforeunload',function(e){ if(_isDirty()){ e.preventDefault(); e.returnValue=''; return ''; } });

/* ====== EXPORT POWERPOINT ====== */
// (L1192 extrait — voir catalog.js ou public/modele-alhee.pptx)
function _b64ToU8(b){const s=atob(b);const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u;}
function _dataURLToU8(d){const s=atob(d.split(',')[1]);const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u;}
// (L1195 extrait — voir catalog.js ou public/modele-alhee.pptx)
// (L1196 extrait — voir catalog.js ou public/modele-alhee.pptx)
// Init templateBytes par fetch du .pptx (au lieu de _b64ToU8(TEMPLATE_B64))
let templateBytes = null;
fetch('/modele-alhee.pptx').then(r=>{if(!r.ok)throw new Error('HTTP '+r.status);return r.arrayBuffer();}).then(b=>{templateBytes=new Uint8Array(b);const ts=document.getElementById('tplStatus');if(ts&&!ts.textContent.includes('\u2713')){ts.textContent='Modèle ALHEE intégré \u2713';ts.style.color='var(--vf)';}}).catch(e=>{console.error('Modèle PPT injoignable :',e);const ts=document.getElementById('tplStatus');if(ts){ts.textContent='⚠ Modèle PPT injoignable';ts.style.color='var(--rouge)';}});
document.getElementById('fileTpl').onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();
  r.onload=ev=>{templateBytes=ev.target.result;document.getElementById('tplStatus').textContent='Modèle : '+f.name+' ✓';document.getElementById('tplStatus').style.color='var(--vf)';};
  r.readAsArrayBuffer(f);};
function buildExportProject(){
  const m=project.meta||{};
  const meta={copro:m.copro,typologie:m.typologie,syndic:m.syndic,adresseSyndic:m.adresseSyndic,destinataire:m.destinataire,fonction:m.fonction,dateVisite:m.dateVisite,affaire:m.affaire,auditeur:m.auditeur,relecture:m.relecture};
  const lots=BDD.lots.map(L=>{const subs=[];
    L.subs.forEach(sub=>{const st=project.lots[L.num].subs[sub.name];
      const has=(st.items&&st.items.length)||(st.remark&&st.remark.trim());if(!has)return;
      const items=(st.items||[]).map(it=>({desig:it.desig||'',constat:it.constat||'',preco:it.preco||'',
        etat:(it.etat&&it.etat!=='na')?it.etat:null, photoBytes: it.photo?_dataURLToU8(it.photo):null}));
      subs.push({name:sub.name,remark:st.remark||'',items});});
    return {num:L.num,label:L.label,subs};}).filter(L=>L.subs.length);
  const _cp=computePlan();
  return {meta,lots,oblig:project.oblig,docs:project.docs,chiffrage:computeChiffrage(),tva:project.tva||10,preco:computePreco(),plan:_cp.plan,planTotals:_cp.totals,synthese:project.synthese||'',conclusion:project.conclusion||'',synthGen:computeSynthGen()};
}
document.getElementById('btnExport').onclick=async()=>{
  if(typeof JSZip==='undefined'||typeof GEN==='undefined'){alert('Moteur non chargé.');return;}
  if(!templateBytes){alert('Charge d\'abord ton modèle PowerPoint (.pptx) via le bouton « Modèle PPT ».');return;}
  const btn=document.getElementById('btnExport');const old=btn.textContent;btn.textContent='⏳ Génération…';btn.disabled=true;
  try{
    const proj=buildExportProject();
    if(!proj.lots.length){alert('Aucun sous-lot rempli : renseigne au moins une remarque ou une anomalie.');return;}
    const blob=await GEN.generate(JSZip, templateBytes.slice(0), proj, {smiley:SMILEY,smileyFull:SMILEY_FULL});
    const a=document.createElement('a');a.href=URL.createObjectURL(blob);
    const nm=(project.meta.copro||'projet').replace(/[^a-z0-9]+/gi,'_').toLowerCase();
    a.download='PPT_'+nm+'.pptx';a.click();
  }catch(err){alert('Erreur de génération : '+err.message);console.error(err);}
  finally{btn.textContent=old;btn.disabled=false;}
};

/* ====== VERSIONS MULTIPLES (IndexedDB) ====== */
const _VER_PREFIX='ver:';
function _idbDelete(k){return _idbOpen().then(db=>new Promise((res,rej)=>{const t=db.transaction(_IDB_STORE,'readwrite');t.objectStore(_IDB_STORE).delete(k);t.oncomplete=()=>res();t.onerror=()=>rej(t.error);}));}
function _idbAllEntries(){return _idbOpen().then(db=>new Promise((res,rej)=>{const t=db.transaction(_IDB_STORE,'readonly');const st=t.objectStore(_IDB_STORE);const out=[];const rq=st.openCursor();rq.onsuccess=e=>{const c=e.target.result;if(c){out.push({key:c.key,value:c.value});c.continue();}else res(out);};rq.onerror=()=>rej(rq.error);}));}
function _verNewId(){return _VER_PREFIX+Date.now()+'_'+Math.random().toString(36).slice(2,7);}
function _verSuggestName(){const c=(project.meta&&project.meta.copro)?project.meta.copro:'PPT';const d=new Date();const p=n=>String(n).padStart(2,'0');return c+' — '+p(d.getDate())+'/'+p(d.getMonth()+1)+'/'+d.getFullYear()+' '+p(d.getHours())+'h'+p(d.getMinutes());}
function _verFmtDate(iso){if(!iso)return '';const d=new Date(iso);if(isNaN(d))return '';const p=n=>String(n).padStart(2,'0');return p(d.getDate())+'/'+p(d.getMonth()+1)+'/'+d.getFullYear()+' à '+p(d.getHours())+'h'+p(d.getMinutes());}

async function saveVersion(name){const rec={id:_verNewId(),name:name||_verSuggestName(),savedAt:new Date().toISOString(),data:JSON.parse(JSON.stringify(project))};await _idbSet(rec.id,JSON.stringify(rec));return rec;}
async function listVersions(){const entries=(await _idbAllEntries()).filter(e=>typeof e.key==='string'&&e.key.indexOf(_VER_PREFIX)===0);const out=[];for(const e of entries){try{const r=JSON.parse(e.value);out.push({key:e.key,name:r.name||'(sans nom)',savedAt:r.savedAt||'',size:(e.value||'').length});}catch(_){}}out.sort((a,b)=>(b.savedAt||'').localeCompare(a.savedAt||''));return out;}
async function loadVersion(key){if(_isDirty() && !confirm('Charger cette version ?\nLes modifications non enregistrées de la session en cours seront perdues.'))return;try{const s=await _idbGet(key);if(!s){alert('Version introuvable.');return;}const r=JSON.parse(s);project=mergeProject(r.data||{});_savedSnap=JSON.stringify(project);await _idbSet(_IDB_KEY,JSON.stringify(project));activeTab='gen';renderTabs();renderMain();_updateStatus();_toast('\u2713 Version chargée : '+(r.name||''));closeVersions();}catch(e){alert('Chargement impossible : '+(e&&e.message||e));}}
async function deleteVersion(key,name){if(!confirm('Supprimer définitivement la version « '+(name||'')+' » ?'))return;try{await _idbDelete(key);_toast('\uD83D\uDDD1 Version supprimée');renderVersions();}catch(e){alert('Suppression impossible : '+(e&&e.message||e));}}
async function renameVersion(key,oldName){const nn=prompt('Nouveau nom de la version :',oldName||'');if(nn==null)return;try{const s=await _idbGet(key);if(!s)return;const r=JSON.parse(s);r.name=(nn.trim()||r.name);await _idbSet(key,JSON.stringify(r));_toast('\u270E Version renommée');renderVersions();}catch(e){alert('Renommage impossible : '+(e&&e.message||e));}}
async function duplicateVersion(key){try{const s=await _idbGet(key);if(!s)return;const r=JSON.parse(s);const copy={id:_verNewId(),name:(r.name||'Version')+' (copie)',savedAt:new Date().toISOString(),data:r.data};await _idbSet(copy.id,JSON.stringify(copy));_toast('\u29C9 Version dupliquée');renderVersions();}catch(e){alert('Duplication impossible : '+(e&&e.message||e));}}
async function downloadVersion(key){try{const s=await _idbGet(key);if(!s)return;const r=JSON.parse(s);const blob=new Blob([JSON.stringify(r.data||{})],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);const nm=((r.name||'ppt')+'').replace(/[^a-z0-9]+/gi,'_').toLowerCase();a.download='PPT_'+nm+'.json';a.click();}catch(e){alert('Téléchargement impossible : '+(e&&e.message||e));}}

function openVersions(){const m=document.getElementById('verModal');if(!m)return;m.classList.add('open');const inp=document.getElementById('verName');if(inp)inp.value=_verSuggestName();renderVersions();}
function closeVersions(){const m=document.getElementById('verModal');if(m)m.classList.remove('open');}
async function renderVersions(){const list=document.getElementById('verList');if(!list)return;list.innerHTML='<div class="vrs-empty">Chargement…</div>';let vers=[];try{vers=await listVersions();}catch(e){list.innerHTML='<div class="vrs-empty">Erreur de lecture.</div>';return;}if(!vers.length){list.innerHTML='<div class="vrs-empty">Aucune version enregistrée pour l\u2019instant.<br>Saisissez un nom ci-dessus puis « ➕ Enregistrer la version actuelle ».</div>';return;}list.innerHTML='';vers.forEach(v=>{const kb=Math.max(1,Math.round((v.size||0)/1024));const row=document.createElement('div');row.className='vrs-row';row.innerHTML='<div class="vrs-info"><div class="vrs-name"></div><div class="vrs-meta">'+_verFmtDate(v.savedAt)+' · '+kb+' Ko</div></div><div class="vrs-actions"><button class="vrs-b vrs-load">Charger</button><button class="vrs-b vrs-ren">Renommer</button><button class="vrs-b vrs-dup">Dupliquer</button><button class="vrs-b vrs-dl">⬇︎ JSON</button><button class="vrs-b vrs-del">Supprimer</button></div>';row.querySelector('.vrs-name').textContent=v.name;row.querySelector('.vrs-load').onclick=()=>loadVersion(v.key);row.querySelector('.vrs-ren').onclick=()=>renameVersion(v.key,v.name);row.querySelector('.vrs-dup').onclick=()=>duplicateVersion(v.key);row.querySelector('.vrs-dl').onclick=()=>downloadVersion(v.key);row.querySelector('.vrs-del').onclick=()=>deleteVersion(v.key,v.name);list.appendChild(row);});}

document.getElementById('btnVersions').onclick=openVersions;
document.getElementById('verClose').onclick=closeVersions;
document.getElementById('verModal').addEventListener('click',function(e){if(e.target===this)closeVersions();});
document.addEventListener('keydown',function(e){if(e.key==='Escape')closeVersions();});
document.getElementById('verSaveBtn').onclick=async()=>{const inp=document.getElementById('verName');const nm=(inp&&inp.value||'').trim();try{await saveVersion(nm);_toast('\uD83C\uDFF7 Version enregistrée');if(inp)inp.value=_verSuggestName();renderVersions();}catch(e){alert('Enregistrement impossible : '+(e&&e.message||e));}};
document.getElementById('verImport').onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=async ev=>{try{const loaded=JSON.parse(ev.target.result);const merged=mergeProject(loaded);const rec={id:_verNewId(),name:(f.name||'Import').replace(/\.json$/i,''),savedAt:new Date().toISOString(),data:merged};await _idbSet(rec.id,JSON.stringify(rec));_toast('\u2B06\uFE0E Importé comme version');renderVersions();}catch(err){alert('JSON invalide');}finally{e.target.value='';}};r.readAsText(f);};

/* ====== INIT ====== */
document.getElementById('btnNew').onclick=newMission;
(async function(){
  try{ await loadLocal(true); }catch(e){}
  var ts=document.getElementById('tplStatus');
  if(ts && templateBytes){ ts.textContent='Modèle ALHEE intégré ✓'; ts.style.color='var(--vf)'; }
  (function(){var z=document.getElementById('zoneSel'); if(z){ z.value=project.zone||'paris'; }})();
  renderTabs(); renderMain();
  if(!_savedSnap) _savedSnap=JSON.stringify(project);
  _updateStatus();
})();

/* ====== PONT CLOUD (Phase 4) ======
   Petite passerelle utilisée par la couche SharePoint (src/graph/*) pour lire la mission
   courante et en charger une depuis le serveur, SANS toucher à la logique métier. */
window.ALHEE_BRIDGE = {
  // mission courante (clone JSON, photos en dataURL + éventuels photoRef/missionId/__serverETag)
  getProject(){ return JSON.parse(JSON.stringify(project)); },
  // remplace la mission courante par celle chargée depuis SharePoint et ré-affiche
  setProject(obj){
    project = mergeProject(obj || {});
    if(obj && obj.missionId) project.missionId = obj.missionId;
    if(obj && obj.__serverETag) project.__serverETag = obj.__serverETag;
    if(obj && obj.statut) project.statut = obj.statut;
    _savedSnap = JSON.stringify(project);
    activeTab = 'gen';
    renderTabs(); renderMain(); _updateStatus();
  },
  // après un enregistrement réussi : mémorise l'id mission + l'eTag serveur sur la mission courante
  markSaved(missionId, etag){
    if(missionId) project.missionId = missionId;
    if(etag) project.__serverETag = etag;
    _savedSnap = JSON.stringify(project);
  },
  newMission(){ try{ newMission(); }catch(_){ } },
  toast(m){ try{ _toast(m); }catch(_){ } },
};
