/* Moteur de génération PPTX ALHEE */
// genEngine.js — moteur de génération PPTX (Node + navigateur via JSZip)
// Entrées : template (bytes), project (saisie), assets {smiley:{bon,moyen,urgent}}
// project.lots[].subs[] = {name, num?, remark, items:[{desig,constat,preco,etat,photoBytes}]}
(function(root){
const CHUNK = 2; // anomalies max par diapo
const esc = s => String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

function setShapeText(xml, anchor, newText){
  const sps = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g)||[];
  for(const sp of sps){
    if(sp.indexOf(anchor)!==-1 && sp.indexOf('<a:t>')!==-1){
      const pM = sp.match(/<a:p>[\s\S]*?<\/a:p>/); if(!pM) continue;
      const p = pM[0];
      const rM = p.match(/<a:r>[\s\S]*?<\/a:r>/); if(!rM) continue;
      const rprM = rM[0].match(/<a:rPr(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:rPr>)/);
      const rpr = rprM?rprM[0]:'<a:rPr lang="fr-FR"/>';
      const newRun = '<a:r>'+rpr+'<a:t>'+esc(newText)+'</a:t></a:r>';
      const newP = p.replace(/(?:<a:r>[\s\S]*?<\/a:r>\s*)+/, newRun);
      return xml.replace(sp, sp.replace(p, newP));
    }
  }
  return xml;
}
// remplace le texte d'une cellule (collapse des runs), en gardant le rPr du 1er run si présent
function setCellText(tr, text){
  const pM = tr.match(/<a:p>[\s\S]*?<\/a:p>/); if(!pM) return tr;
  const p = pM[0];
  const rM = p.match(/<a:r>[\s\S]*?<\/a:r>/);
  let rpr = '<a:rPr lang="fr-FR" sz="1400"><a:latin typeface="Montserrat"/></a:rPr>';
  if(rM){ const x=rM[0].match(/<a:rPr(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:rPr>)/); if(x) rpr=x[0]; }
  const newRun = '<a:r>'+rpr+'<a:t>'+esc(text)+'</a:t></a:r>';
  let newP;
  if(rM) newP = p.replace(/(?:<a:r>[\s\S]*?<\/a:r>\s*)+/, newRun);
  else   newP = p.replace('</a:p>', newRun+'</a:p>'); // cellule sans run
  return tr.replace(p, newP);
}
const blip = r => '<a:blipFill rotWithShape="1"><a:blip r:embed="'+r+'"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>';
const blipInset = (r,p) => '<a:blipFill rotWithShape="1"><a:blip r:embed="'+r+'"/><a:stretch><a:fillRect l="'+p+'" t="'+p+'" r="'+p+'" b="'+p+'"/></a:stretch></a:blipFill>';
function setText(tc, t){
  if(/<a:t>[\s\S]*?<\/a:t>/.test(tc)) return tc.replace(/<a:t>[\s\S]*?<\/a:t>/, '<a:t>'+esc(t)+'</a:t>');
  return tc.replace('</a:p>', '<a:r><a:rPr lang="fr-FR" sz="1400"><a:latin typeface="Montserrat"/></a:rPr><a:t>'+esc(t)+'</a:t></a:r></a:p>');
}
const setFill = (tc,r)=> tc.replace('<a:noFill/>', blip(r));
const setFillSmiley = (tc,r)=> tc.replace('<a:noFill/>', blip(r));
function _totalCell(tc, val){
  tc=_setTc(tc, val);
  tc=tc.replace(/<a:rPr\b([^>]*?)(\/?)>/, (m,a,sc)=>{ a=a.replace(/\sb="[01]"/,''); return '<a:rPr'+a+' b="1"'+(sc?'/':'')+'>'; });
  tc=tc.replace(/(<a:rPr\b[^>]*?)\ssz="\d+"/, '$1 sz="2000"');
  if(/<a:rPr\b[^>]*>\s*<a:solidFill>/.test(tc)) tc=tc.replace(/(<a:rPr\b[^>]*>)<a:solidFill>[\s\S]*?<\/a:solidFill>/, '$1<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>');
  else tc=tc.replace(/(<a:rPr\b[^>]*[^\/]>)/, '$1<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>');
  if(/<a:pPr\b[^>]*\salgn="/.test(tc)) tc=tc.replace(/(<a:pPr\b[^>]*?)\salgn="[^"]*"/, '$1 algn="ctr"');
  else if(/<a:pPr\b/.test(tc)) tc=tc.replace(/<a:pPr\b/, '<a:pPr algn="ctr"');
  else tc=tc.replace(/<a:p>/, '<a:p><a:pPr algn="ctr"/>');
  tc=tc.replace(/<a:noFill\/>(\s*<\/a:tcPr>)/, '<a:solidFill><a:srgbClr val="18483C"/></a:solidFill>$1');
  return tc;
}
function imgExt(b){ if(!b)return 'png'; if(b[0]===0x89&&b[1]===0x50)return 'png'; if(b[0]===0xFF&&b[1]===0xD8)return 'jpg'; if(b[0]===0x47&&b[1]===0x49)return 'gif'; return 'png'; }

// ===== Remplisseurs de sections (Obligations / Documents / Chiffrage) =====
function _normLabel(s){
  return String(s||'').replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&amp;/g,'&')
    .replace(/&#8217;|&#x2019;|’/g,"'").normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().replace(/[^a-z0-9]+/g,'');
}
function _tcText(tc){ const ts=tc.match(/<a:t>[\s\S]*?<\/a:t>/g)||[]; return _normLabel(ts.map(t=>t.replace(/<\/?a:t>/g,'')).join('')); }
function _setTc(tc, val){
  if(!/<a:t>[\s\S]*?<\/a:t>/.test(tc))
    return tc.replace('</a:p>','<a:r><a:rPr lang="fr-FR" sz="1600"><a:latin typeface="Montserrat"/></a:rPr><a:t>'+esc(val)+'</a:t></a:r></a:p>');
  let n=0; return tc.replace(/<a:t>[\s\S]*?<\/a:t>/g, ()=>{ n++; return n===1?'<a:t>'+esc(val)+'</a:t>':'<a:t></a:t>'; });
}
async function fillDispoTables(zip, project){
  const o=project.oblig||{}, d=project.docs||{};
  const mk=arrs=>{ const m={}; arrs.forEach(a=>(a||[]).forEach(r=>{ if(r&&r.label) m[_normLabel(r.label)]={dispo:r.dispo||'',comment:r.comment||''}; })); return m; };
  const regMap=mk([o.reglementaires]), conMap=mk([o.contrats]), docMap=mk([d.admin,d.plans]);
  if(!Object.keys(regMap).length && !Object.keys(conMap).length && !Object.keys(docMap).length) return;
  const dec=s=>s.replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&amp;/g,'&');
  const files=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f));
  for(const f of files){
    let s=await zip.file(f).async('string'); if(s.indexOf('Disponibilité')===-1) continue;
    let chg=false;
    s=s.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, tbl=>{
      if(tbl.indexOf('Disponibilité')===-1) return tbl;
      const firstT=(tbl.match(/<a:t>([\s\S]*?)<\/a:t>/)||[,''])[1];
      const title=dec(firstT).toLowerCase();
      let map=null;
      if(title.indexOf('obligation')!==-1) map=regMap;
      else if(title.indexOf('contrat')!==-1) map=conMap;
      else if(title.indexOf('document')!==-1) map=docMap;
      if(!map) return tbl;
      return tbl.replace(/<a:tr[\s\S]*?<\/a:tr>/g, row=>{
        const cells=row.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g)||[];
        if(cells.length<2) return row;
        const rec=map[_tcText(cells[0])]; if(!rec) return row;
        let nr=row.replace(cells[1], _setTc(cells[1], rec.dispo));
        if(cells.length>=3) nr=nr.replace(cells[2], _setTc(cells[2], rec.comment));
        chg=true; return nr;
      });
    });
    if(chg) zip.file(f, s);
  }
}
function _progOf(s){ const t=s.toLowerCase();
  if(t.indexOf('performance')!==-1||t.indexOf('(p0)')!==-1) return 'P0';
  if(t.indexOf('court terme')!==-1||t.indexOf('(p1)')!==-1) return 'P1';
  if(t.indexOf('moyen terme')!==-1||t.indexOf('(p2)')!==-1) return 'P2';
  if(t.indexOf('long terme')!==-1||t.indexOf('(p3)')!==-1) return 'P3';
  return null; }
function _fmtTTC(n){ n=Math.round(Number(n)||0); return n? n.toLocaleString('fr-FR').replace(/\u202f|\u00a0/g,' ') : '0'; }
function _rowOpen(row){ return row.slice(0, row.indexOf('<a:tc')); }
function _grid3(open){
  return open.replace(/<a:tblGrid>[\s\S]*?<\/a:tblGrid>/,
    '<a:tblGrid><a:gridCol w="6000000"/><a:gridCol w="3410700"/><a:gridCol w="2095500"/></a:tblGrid>');
}
async function fillChiffrage(zip, project){
  const ch=project.chiffrage; if(!ch) return;
  const files=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f));
  for(const f of files){
    let s=await zip.file(f).async('string'); if(s.indexOf('Coût')===-1) continue;
    const prog=_progOf(s); if(!prog) continue;
    const works=ch[prog]||[];
    s=s.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/, tbl=>{
      if(tbl.indexOf('Coût')===-1) return tbl;
      const rows=tbl.match(/<a:tr[\s\S]*?<\/a:tr>/g)||[]; if(!rows.length) return tbl;
      const header=rows[0].replace(/<a:extLst>[\s\S]*?<\/a:extLst>/,'');
      const model=(rows[1]||rows[0]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/,'');
      const open=_grid3(tbl.slice(0, tbl.indexOf('<a:tr')));
      const hc=header.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g);
      const newHeader=_rowOpen(header)+hc[0]+_setTc(hc[0],'Localisation')+hc[1]+'</a:tr>';
      const mc=model.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g);
      const row3=(a,b,c)=> _rowOpen(model)+_setTc(mc[0],a)+_setTc(mc[0],b)+_setTc(mc[1],c)+'</a:tr>';
      const list=works.length? works : [{travaux:'Aucun poste identifié à ce stade.', loc:'', ttc:0, ht:0}];
      let body='', tot=0;
      list.forEach(w=>{ tot+=(w.ttc||0); body+=row3(w.travaux, w.loc||'', (w.ht? _fmtTTC(w.ttc):'À chiffrer')); });
      body+=_rowOpen(model)+_totalCell(mc[0],'TOTAL')+_totalCell(mc[0],'')+_totalCell(mc[1],_fmtTTC(tot))+'</a:tr>';
      return open+newHeader+body+'</a:tbl>';
    });
    zip.file(f, s);
  }
}

// ===== Sections paginées clonées par page (Partie 4 préconisations, Plan 10 ans) =====
function _tcRaw(tc){ const ts=tc.match(/<a:t>([\s\S]*?)<\/a:t>/g)||[]; return ts.map(t=>t.replace(/<\/?a:t>/g,'')).join('').replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&amp;/g,'&'); }
function _setTitleSuffix(s, oldLot, newLot){
  return s.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, sp=>{
    if(sp.indexOf('Préconisation')===-1 && sp.indexOf('Plan ')===-1) return sp;
    return sp.replace(/<a:t>([\s\S]*?)<\/a:t>/g, m=>{
      const t=m.replace(/<\/?a:t>/g,'');
      return (t.trim()===oldLot.trim()) ? '<a:t>'+esc(newLot)+'</a:t>' : m;
    });
  });
}
async function buildPaginatedSection(zip, modelFile, kind, sectionData){
  if(!sectionData || !sectionData.length) return [];
  const model = await zip.file('ppt/slides/'+modelFile).async('string');
  let rels = await zip.file('ppt/slides/_rels/'+modelFile+'.rels').async('string');
  rels = rels.replace(/<Relationship [^>]*notesSlide[^>]*\/>/,'');
  const tbl = model.match(/<a:tbl>[\s\S]*?<\/a:tbl>/)[0];
  const trs = tbl.match(/<a:tr[\s\S]*?<\/a:tr>/g);
  const ghModel = trs[0].replace(/<a:extLst>[\s\S]*?<\/a:extLst>/,'');
  const dataModel = (trs[1]||trs[0]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/,'');
  const tblOpen = tbl.slice(0, tbl.indexOf('<a:tr'));
  const ghCells = ghModel.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g);
  const dCells  = dataModel.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g);
  const modelLot = _tcRaw(ghCells[0]);
  const groupHeaderRow = label => _rowOpen(ghModel)+_setTc(ghCells[0],label)+ghCells.slice(1).join('')+'</a:tr>';
  function dataRow(r){
    if(kind==='plan'){
      let cells=_setTc(dCells[0], r.travaux)+_setTc(dCells[1], r.loc||'');
      const ys=r.years||(r.annee?[Number(r.annee)]:[]);
      for(let y=1;y<=10;y++){ const ci=1+y; if(dCells[ci]) cells+=_setTc(dCells[ci], (ys.indexOf(y)>=0)?'X':''); }
      return _rowOpen(dataModel)+cells+'</a:tr>';
    }
    return _rowOpen(dataModel)+_setTc(dCells[0],r.travaux)+_setTc(dCells[1],r.loc||'')+_setTc(dCells[2], r.prio||'')+'</a:tr>';
  }
  const BUDGET = kind==='plan'?6:7;
  const pages=[]; let cur=[], curN=0, curFirst=null;
  sectionData.forEach(grp=>{
    if(curN && curN>=BUDGET-1){ pages.push({rows:cur,first:curFirst}); cur=[]; curN=0; curFirst=null; }
    cur.push(groupHeaderRow(grp.label)); curN++; if(!curFirst) curFirst=grp.label;
    grp.rows.forEach(r=>{
      if(curN>=BUDGET){ pages.push({rows:cur,first:curFirst}); cur=[]; curN=0; curFirst=grp.label+' (suite)'; cur.push(groupHeaderRow(grp.label+' (suite)')); curN++; }
      cur.push(dataRow(r)); curN++;
    });
  });
  if(cur.length) pages.push({rows:cur,first:curFirst});
  return pages.map(pg=>{
    let s = model.replace(/<p:pic>[\s\S]*?<\/p:pic>/g,'');
    s = s.replace(tbl, tblOpen+pg.rows.join('')+'</a:tbl>');
    s = _setTitleSuffix(s, modelLot, pg.first||'');
    return {xml:s, relsXml:rels};
  });
}
async function fillPlanTotals(zip, project){
  const pt=project.planTotals; if(!pt) return;
  const files=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f));
  for(const f of files){
    let s=await zip.file(f).async('string'); if(s.indexOf('TTC)')===-1) continue;
    let chg=false;
    // totaux par année (ligne TOTAL d'un tableau à 11 colonnes avec 'Année 1')
    s=s.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, tbl=>{
      const rows=tbl.match(/<a:tr[\s\S]*?<\/a:tr>/g)||[];
      if(tbl.indexOf('Année 1')===-1 || rows.length<2) return tbl;
      const tr=rows[1]; const cells=tr.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g);
      if(!cells||cells.length<11) return tbl;
      let ntr=tr;
      for(let y=1;y<=10;y++){ const v=(pt.years&&pt.years[y])?_fmtTTC(pt.years[y]):'0'; ntr=ntr.replace(cells[y], _setTc(cells[y], v)); }
      chg=true; return tbl.replace(tr, ntr);
    });
    // total général (petit tableau 'TOTAL' / 'xxx €TTC')
    s=s.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, tbl=>{
      const rows=tbl.match(/<a:tr[\s\S]*?<\/a:tr>/g)||[];
      if(rows.length!==2) return tbl;
      const c0=rows[0].match(/<a:tc[ >][\s\S]*?<\/a:tc>/g);
      if(!c0||c0.length!==1||_tcText(c0[0])!=='total') return tbl;
      const c1=rows[1].match(/<a:tc[ >][\s\S]*?<\/a:tc>/g); if(!c1) return tbl;
      chg=true; return tbl.replace(rows[1], rows[1].replace(c1[0], _setTc(c1[0], _fmtTTC(pt.grand)+' €TTC')));
    });
    // montant figé dans le texte narratif (« … s'élève à 216 100 € TTC »)
    const ns=s.replace(/(<a:t>)([\d\u00a0\u202f ]+)€ TTC(<\/a:t>)/, '$1'+_fmtTTC(pt.grand)+' € TTC$3');
    if(ns!==s){ s=ns; chg=true; }
    if(chg) zip.file(f,s);
  }
}

// ===== Slides texte : Synthèse (3.8) + Conclusion (7) — mini-markdown **gras** =====
function _mdRuns(line, rprN, rprB){
  let out='', last=0, m; const re=/\*\*([^*]+)\*\*/g;
  while((m=re.exec(line))){ if(m.index>last) out+='<a:r>'+rprN+'<a:t>'+esc(line.slice(last,m.index))+'</a:t></a:r>';
    out+='<a:r>'+rprB+'<a:t>'+esc(m[1])+'</a:t></a:r>'; last=re.lastIndex; }
  if(last<line.length) out+='<a:r>'+rprN+'<a:t>'+esc(line.slice(last))+'</a:t></a:r>';
  if(!out) out='<a:endParaRPr lang="fr-FR"/>'; return out;
}
function _fillTextShape(xml, text){
  const sps=xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g)||[]; let best=null,bl=-1;
  for(const sp of sps){ const t=(sp.match(/<a:t>([^<]*)<\/a:t>/g)||[]).join('').length; if(t>bl){bl=t;best=sp;} }
  if(!best) return xml;
  const rM=best.match(/<a:rPr(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:rPr>)/);
  let rprN=rM?rM[0]:'<a:rPr lang="fr-FR" sz="1100"><a:latin typeface="Montserrat"/></a:rPr>';
  rprN=rprN.replace(/\sb="1"/,' b="0"');
  let rprB = /\sb="0"/.test(rprN) ? rprN.replace(/\sb="0"/,' b="1"') : rprN.replace('<a:rPr','<a:rPr b="1"');
  const pM=best.match(/<a:pPr[^>]*\/>|<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/); const ppr=pM?pM[0]:'';
  const paras=String(text).split('\n').map(l=>'<a:p>'+ppr+_mdRuns(l.trim(),rprN,rprB)+'</a:p>').join('');
  const nb=best.replace(/(<p:txBody>[\s\S]*?)<a:p>[\s\S]*<\/a:p>([\s\S]*?<\/p:txBody>)/, '$1'+paras+'$2');
  return xml.replace(best, nb);
}
async function fillTextSlide(zip, marker, text){
  const files=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f));
  for(const f of files){ let s=await zip.file(f).async('string');
    if(s.indexOf(marker)===-1 || s.indexOf('SOMMAIRE')!==-1) continue;
    zip.file(f, _fillTextShape(s, text)); return true; }
  return false;
}

// ===== Page 3.1 Synthèse générale (tableaux par section + emoji d'état + note) =====
function _synPic(id, rid, x, y, w, h){
  return '<p:pic><p:nvPicPr><p:cNvPr id="'+id+'" name="emoji_synthese"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
    +'<p:blipFill><a:blip r:embed="'+rid+'"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
    +'<p:spPr><a:xfrm><a:off x="'+Math.round(x)+'" y="'+Math.round(y)+'"/><a:ext cx="'+w+'" cy="'+h+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
}
function _synNote(id, text){
  const runs=String(text||'').split(/\n/).map(ln=>'<a:p><a:pPr/><a:r><a:rPr lang="fr-FR" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="3E5B3F"/></a:solidFill><a:latin typeface="Montserrat"/></a:rPr><a:t>'+esc(ln)+'</a:t></a:r></a:p>').join('');
  return '<p:sp><p:nvSpPr><p:cNvPr id="'+id+'" name="note_synthese"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>'
    +'<p:spPr><a:xfrm><a:off x="1219200" y="1300000"/><a:ext cx="12115800" cy="1080000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    +'<p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>'+runs+'</p:txBody></p:sp>';
}
async function fillSynthGen(zip, project, assets){
  const sg=project.synthGen; if(!sg) return;
  const emo=(assets&&assets.smileyFull)||(assets&&assets.smiley); if(!emo) return;
  const etats=sg.etats||{};
  const alias={facades:'facade',plancherbas:'planchersbas',balconloggiasterrasses:'balconsloggiasterrasses',efsetecs:'eaufroideeteauchaude',partiescommunes:'securiteincendie',interieur:'interieure'};
  const norm=s=>String(s||'').replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&amp;/g,'&').replace(/&#8217;|&#x2019;|’/g,"'").normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'');
  const etatOf=name=>{ let k=norm(name); if(alias[k])k=alias[k]; return etats[k]||'na'; };
  const PIC=560000;
  const files=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f)).sort();
  let noteAdded=false;
  for(const f of files){
    const relPath='ppt/slides/_rels/'+f.split('/').pop()+'.rels';
    let rels=''; if(zip.file(relPath)) rels=await zip.file(relPath).async('string');
    if(!/Target="\.\.\/media\/image(4[7-9]|5[0-9])\.png"/.test(rels)) continue; // slides synthèse uniquement
    let s=await zip.file(f).async('string');
    const r2t={}; for(const mm of rels.matchAll(/Id="(rId\d+)"[^>]*Target="\.\.\/media\/([^"]+)"/g)) r2t[mm[1]]=mm[2];
    // 1) retirer les emojis flottants d'origine
    s=s.replace(/<p:pic>[\s\S]*?<\/p:pic>/g, pic=>{ const e=pic.match(/r:embed="(rId\d+)"/); const t=e?r2t[e[1]]:''; return /^image(4[7-9]|5[0-9])\.png$/.test(t||'')?'':pic; });
    // 2) rels + media pour nos 3 emojis
    let maxRid=0; for(const k in r2t){ const n=parseInt(k.slice(3))||0; if(n>maxRid)maxRid=n; }
    const rid={bon:'rId'+(maxRid+1),moyen:'rId'+(maxRid+2),urgent:'rId'+(maxRid+3)};
    const mediaN={bon:'syn_bon.png',moyen:'syn_moyen.png',urgent:'syn_urgent.png'};
    let add=''; ['bon','moyen','urgent'].forEach(k=>{ add+='<Relationship Id="'+rid[k]+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/'+mediaN[k]+'"/>'; });
    rels=rels.replace('</Relationships>', add+'</Relationships>'); zip.file(relPath, rels);
    ['bon','moyen','urgent'].forEach(k=>{ if(emo[k]) zip.file('ppt/media/'+mediaN[k], emo[k]); });
    // 3) parcourir chaque table : poser emoji (pic) ou "-" selon l'état du sous-lot
    let idN=920; let newPics='';
    s=s.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g, gf=>{
      const off=gf.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/); if(!off) return gf;
      const cols=[...gf.matchAll(/<a:gridCol w="(\d+)"/g)].map(x=>+x[1]); if(cols.length<2) return gf;
      const ox=+off[1], oy=+off[2], col0=cols[0], col1=cols[1];
      const tblM=gf.match(/<a:tbl>[\s\S]*?<\/a:tbl>/); if(!tblM) return gf;
      const trs=tblM[0].match(/<a:tr h="\d+">[\s\S]*?<\/a:tr>/g)||[];
      let y=oy, newTbl=tblM[0];
      trs.forEach(tr=>{
        const h=+(tr.match(/h="(\d+)"/)[1]);
        const cells=tr.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g)||[];
        if(cells.length>=2){
          const n0=_tcRaw(cells[0]).trim(), t1=_tcRaw(cells[1]).trim();
          if(n0 && t1!=='État' && t1!=='Etat'){
            const e=etatOf(n0);
            const newC1=_setTc(cells[1], e==='na'?'-':'');
            if(newC1!==cells[1]){ const ntr=tr.replace(cells[1], newC1); newTbl=newTbl.replace(tr, ntr); }
            if(e==='bon'||e==='moyen'||e==='urgent'){
              const cx=ox+col0+col1/2, cy=y+h/2;
              newPics+=_synPic(idN++, rid[e], cx-PIC/2, cy-PIC/2, PIC, PIC);
            }
          }
        }
        y+=h;
      });
      return gf.replace(tblM[0], newTbl);
    });
    if(newPics) s=s.replace('</p:spTree>', newPics+'</p:spTree>');
    if(!noteAdded && sg.note && sg.note.trim()){ s=s.replace('</p:spTree>', _synNote(idN++, sg.note.trim())+'</p:spTree>'); noteAdded=true; }
    zip.file(f, s);
  }
}

async function generate(JSZip, templateBytes, project, assets){
  const zip = await JSZip.loadAsync(templateBytes);
  const RD = p => zip.file(p).async('string');

  // détection des slides clés (robuste au numéro)
  async function findSlide(test){
    const files=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f)).sort();
    for(const f of files){ const x=await zip.file(f).async('string'); if(test(x)) return f.split('/').pop(); }
    return null;
  }
  const modelFile = await findSlide(x=> x.indexOf('<a:t>Désignation</a:t>')!==-1 && x.indexOf('<a:t>Constat</a:t>')!==-1 && (x.indexOf('<a:t>Etat</a:t>')!==-1 || x.indexOf('<a:t>État</a:t>')!==-1));
  if(!modelFile) throw new Error("Slide-modèle introuvable (tableau Désignation/Constat/État). Vérifiez le modèle PPT.");

  // 1) Remplacer les tokens méta sur TOUTES les slides (page de garde + Généralités + autres)
  const m=project.meta||{};
  const map={'{{COPRO}}':m.copro,'{{TYPO}}':m.typologie,'{{SYNDIC}}':m.syndic,'{{SYNDIC_ADR}}':m.adresseSyndic,
    '{{DEST}}':m.destinataire,'{{FONCTION}}':m.fonction,'{{DATE_VISITE}}':m.dateVisite,'{{AFFAIRE}}':m.affaire,
    '{{AUDITEUR}}':m.auditeur,'{{RELECTURE}}':m.relecture};
  {
    const slideFiles=Object.keys(zip.files).filter(f=>/^ppt\/slides\/slide\d+\.xml$/.test(f));
    for(const f of slideFiles){
      let s=await zip.file(f).async('string'); let changed=false;
      for(const k in map){ if(s.indexOf(k)!==-1){ s=s.split(k).join(esc(map[k]||'')); changed=true; } }
      if(changed) zip.file(f, s);
    }
  }

  // 1bis) Sections à tableaux : Obligations / Documents (Disponibilité + Commentaire) + Chiffrage par programme
  await fillDispoTables(zip, project);
  await fillChiffrage(zip, project);
  await fillPlanTotals(zip, project);
  await fillTextSlide(zip, '3.8', (project.synthese&&project.synthese.trim())? project.synthese : "Synthèse de l'état des lieux à compléter.");
  await fillTextSlide(zip, '7. Conclusion', (project.conclusion&&project.conclusion.trim())? project.conclusion : "Conclusion générale à compléter.");
  await fillSynthGen(zip, project, assets);

  // 2) MODÈLE technique
  const model = await RD('ppt/slides/'+modelFile);
  let modelRels = await RD('ppt/slides/_rels/'+modelFile+'.rels');
  modelRels = modelRels.replace(/<Relationship [^>]*notesSlide[^>]*\/>/,''); // retirer la note (évite partage entre clones)

  // médias globaux
  const mediaSink = { files:{} };
  const smileyName = {bon:'gen_sm_bon.png',moyen:'gen_sm_moyen.png',urgent:'gen_sm_urgent.png'};
  let photoIdx=0;

  function buildSlide(lot, sub, subIndex, items, isCont){
    let s = model.replace(/<p:pic>[\s\S]*?<\/p:pic>/g,'');
    s = setShapeText(s, 'Voieries', lot.num+' '+lot.label);
    s = setShapeText(s, '3.2.1', lot.num+'.'+(subIndex+1)+' '+sub.name+(isCont?' (suite)':''));
    const tbl = s.match(/<a:tbl>[\s\S]*?<\/a:tbl>/)[0];
    const trs = tbl.match(/<a:tr h="\d+">[\s\S]*?<\/a:tr>/g);
    const tcT = trs[4].match(/<a:tc>[\s\S]*?<\/a:tc>/g);
    // rels du slide : on PART des rels d'origine (logo/layout préservés), images en rId100+
    let myRels=''; let ridN=100; const smCache={};
    const localImg=name=>{ const id='rId'+(ridN++); myRels+='<Relationship Id="'+id+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/'+name+'"/>'; return id; };
    const smileyRid=state=>{ if(!state||!smileyName[state])return null;
      if(!mediaSink.files[smileyName[state]]) mediaSink.files[smileyName[state]]=assets.smiley[state];
      if(!smCache[state]) smCache[state]=localImg(smileyName[state]); return smCache[state]; };
    const dataRows = items.map(an=>{
      let pr=null;
      if(an.photoBytes){ const ext=imgExt(an.photoBytes); const nm='gen_ph_'+(photoIdx++)+'.'+ext; mediaSink.files[nm]=an.photoBytes; pr=localImg(nm); }
      const img = pr? setFill(tcT[0], pr) : tcT[0];
      const eta = (function(){ const r=smileyRid(an.etat); return r? setFillSmiley(tcT[4], r) : tcT[4]; })();
      return '<a:tr h="2200150">'+img+setText(tcT[1],an.desig)+setText(tcT[2],an.constat)+setText(tcT[3],an.preco)+eta+'</a:tr>';
    }).join('');
    const row0 = setCellText(trs[0], sub.name);
    const row2 = setCellText(trs[2], sub.remark||'');
    const remarkBlock = isCont ? '' : (trs[1]+row2);        // remarque non répétée sur les pages de suite
    const colHeader = items.length ? trs[3] : '';           // pas d'en-tête colonnes si la section n'a que des remarques
    const newTbl = tbl.replace(trs.join(''), row0+remarkBlock+colHeader+dataRows);
    s = s.replace(tbl, newTbl);
    const relsXml = modelRels.replace('</Relationships>', myRels+'</Relationships>');
    return {xml:s, relsXml};
  }

  // liste des slides à générer
  const gen=[];
  (project.lots||[]).forEach(lot=>{
    (lot.subs||[]).forEach((sub,si)=>{
      const has = (sub.items&&sub.items.length) || (sub.remark&&sub.remark.trim());
      if(!has) return; // sous-lot vide -> pas de slide (variabilité)
      const items = sub.items||[];
      const chunks = items.length ? chunkArr(items, CHUNK) : [[]];
      chunks.forEach((part,ci)=> gen.push(buildSlide(lot, sub, si, part, ci>0)));
    });
  });

  // sections « slide-modèle -> N slides » (technique + Partie 4 + Plan)
  const sections=[{modelFile, slides:gen}];
  const precoModel = await findSlide(x=> x.indexOf('Préconisation')!==-1 && x.indexOf('Priorité')!==-1 && x.indexOf('<a:tbl>')!==-1);
  if(precoModel) sections.push({modelFile:precoModel, slides: await buildPaginatedSection(zip, precoModel, 'preco', project.preco||[])});
  const planModel = await findSlide(x=> x.indexOf('Année 1')!==-1 && x.indexOf('Localisation')!==-1 && x.indexOf('<a:tbl>')!==-1);
  if(planModel) sections.push({modelFile:planModel, slides: await buildPaginatedSection(zip, planModel, 'plan', project.plan||[])});

  // 3) ASSEMBLAGE (multi-sections, compteurs partagés)
  let pres = await RD('ppt/presentation.xml');
  let prels = await RD('ppt/_rels/presentation.xml.rels');
  let ct = await RD('[Content_Types].xml');
  const rel2file={}; for(const mm of prels.matchAll(/<Relationship Id="(rId\d+)"[^>]*Target="slides\/(slide\d+\.xml)"/g)) rel2file[mm[1]]=mm[2];
  let fileN=1001, ridN=900, idN=730;
  for(const sec of sections){
    let modelRid=null; for(const k in rel2file) if(rel2file[k]===sec.modelFile) modelRid=k;
    if(!modelRid) continue;
    const sldIdRe = new RegExp('<p:sldId[^>]*r:id="'+modelRid+'"\\s*/>');
    let newSldIds='';
    (sec.slides||[]).forEach(g=>{
      const file='slide'+(fileN++)+'.xml';
      zip.file('ppt/slides/'+file, g.xml);
      zip.file('ppt/slides/_rels/'+file+'.rels', g.relsXml);
      ct = ct.replace('</Types>', '<Override PartName="/ppt/slides/'+file+'" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
      const rid='rId'+(ridN++);
      prels = prels.replace('</Relationships>', '<Relationship Id="'+rid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/'+file+'"/></Relationships>');
      newSldIds += '<p:sldId id="'+(idN++)+'" r:id="'+rid+'"/>';
    });
    pres = pres.replace(sldIdRe, newSldIds);
    prels = prels.replace(new RegExp('<Relationship Id="'+modelRid+'"[^>]*/>'),'');
    ct = ct.replace(new RegExp('<Override PartName="/ppt/slides/'+sec.modelFile.replace('.','\\.')+'"[^>]*/>'),'');
    zip.remove('ppt/slides/'+sec.modelFile); zip.remove('ppt/slides/_rels/'+sec.modelFile+'.rels');
  }

  zip.file('ppt/presentation.xml', pres);
  zip.file('ppt/_rels/presentation.xml.rels', prels);
  ['png','jpg','jpeg'].forEach(x=>{ if(ct.indexOf('Extension="'+x+'"')===-1) ct=ct.replace('</Types>','<Default Extension="'+x+'" ContentType="image/'+(x==='png'?'png':'jpeg')+'"/></Types>'); });
  zip.file('[Content_Types].xml', ct);

  // 4) MÉDIAS
  for(const name in mediaSink.files) zip.file('ppt/media/'+name, mediaSink.files[name]);

  return await zip.generateAsync(root.NODE?{type:'nodebuffer'}:{type:'blob'});
}
function chunkArr(a,n){ const o=[]; for(let i=0;i<a.length;i+=n)o.push(a.slice(i,i+n)); return o; }

const api={generate, _setShapeText:setShapeText};
if(typeof module!=='undefined'&&module.exports){ root.NODE=true; module.exports=api; }
else root.GEN=api;
})(typeof window!=='undefined'?window:globalThis);
