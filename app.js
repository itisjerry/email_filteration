'use strict';
// ---------- Error reporter ----------
(function(){
  function report(msg,src,line,col,err){
    try{
      var box=document.getElementById('testOut');
      if(!box) return false;
      var details=(err&&err.stack)?('\n'+err.stack):'';
      box.innerHTML+='\n❌ Runtime error: '+msg+(src?('\n at '+src+':'+line+':'+col):'')+details+'\n';
    }catch(_){ }
    return false;
  }
  window.onerror=report;
  window.addEventListener('unhandledrejection',function(e){report('Unhandled promise rejection: '+(e.reason&&e.reason.message?e.reason.message:e.reason),'',0,0,e.reason);});
})();

// ---------- Icons ----------
function refreshIcons(){ if(window.feather && typeof window.feather.replace==='function'){ window.feather.replace({width:18,height:18}); } }
document.addEventListener('DOMContentLoaded', refreshIcons);

// ---------- Elements (Part 1) ----------
var fileInput=document.getElementById('file');
var p1Badges=document.getElementById('p1Badges');
var drop=document.getElementById('drop');
var workArea=document.getElementById('workArea');
var statusText=document.getElementById('statusText');
var cleanBtn=document.getElementById('cleanBtn');
var downloadXLSX=document.getElementById('downloadXLSX');
var downloadXLSX2=document.getElementById('downloadXLSX2');
var sticky=document.getElementById('stickyActions');
var stickyMsg=document.getElementById('stickyMsg');
var rowsBeforeEl=document.getElementById('rowsBefore');
var rowsAfterEl=document.getElementById('rowsAfter');
var dupePctEl=document.getElementById('dupePct');
var uniqueCompaniesEl=document.getElementById('uniqueCompanies');
var segmentCountEl=document.getElementById('segmentCount');
var sampleBtn=document.getElementById('sampleBtn');
var testBtn=document.getElementById('testBtn');
var resetBtn=document.getElementById('resetBtn');

// ---------- Elements (Settings) ----------
var kwInput=document.getElementById('kwInput');
var denyInput=document.getElementById('denyInput');
var kwChips=document.getElementById('kwChips');
var denyChips=document.getElementById('denyChips');
var kwCount=document.getElementById('kwCount');
var denyCount=document.getElementById('denyCount');
var regexPreview=document.getElementById('regexPreview');

// ---------- Elements (Part 2) ----------
var prevFileEl=document.getElementById('prevFile');
var newFileEl=document.getElementById('newFile');
var clearPrev=document.getElementById('clearPrev');
var clearNew=document.getElementById('clearNew');
var prevLabel=document.getElementById('prevLabel');
var newLabel=document.getElementById('newLabel');
var prevBadge=document.getElementById('prevBadge');
var newBadge=document.getElementById('newBadge');
var workArea2=document.getElementById('workArea2');
var statusText2=document.getElementById('statusText2');
var compareBtn=document.getElementById('compareBtn');
var downloadCompare=document.getElementById('downloadCompare');
var prevCountEl=document.getElementById('prevCount');
var newCountEl=document.getElementById('newCount');
var newOnlyRelCountEl=document.getElementById('newOnlyRelCount');
var newOnlyIrrelCountEl=document.getElementById('newOnlyIrrelCount');

// ---------- Config ----------
var REQUIRED_COLS=['Company ID','Company Name','Contact Name','First Name','Last Name','Email','Contact Phone','Industry Value','Website','City','State'];
var FRIENDLY_MSGS=['Merging your files…','Scanning sheets…','Cleaning duplicate emails…','Removing empty rows…','Optimizing your list…','Polishing headers…','Almost there…','Preparing download…'];

// ---------- State (Part 1) ----------
var rawRows=[];      // merged rows before cleaning
var cleanedRows=[];  // after cleaning
var segmentRows=[];  // irrelevant subset
var removedDupRows=[]; // removed duplicates with reason

// ---------- State (Part 2) ----------
var prevEmailsSet=null;
var newRelevantRowsP2=[];
var newIrrelevantRowsP2=[];
var newOnlyRel=[];
var newOnlyIrrel=[];

// ---------- Helpers ----------
function ensureXLSX(){ if(typeof XLSX==='undefined'){ throw new Error('SheetJS (XLSX) failed to load. Check network or CDN.'); } }
function setProgress(p){ var el=document.getElementById('bar'); if(!el) return; el.style.width=Math.max(0,Math.min(100,p))+'%'; el.parentElement.setAttribute('aria-valuenow',Math.round(p)); }
function setProgress2(p){ var el=document.getElementById('bar2'); if(!el) return; el.style.width=Math.max(0,Math.min(100,p))+'%'; el.parentElement.setAttribute('aria-valuenow',Math.round(p)); }
function showWorkArea(show){ if(!workArea) return; workArea.classList.toggle('hidden',!show); workArea.setAttribute('aria-busy', show ? 'true' : 'false'); }
function showWorkArea2(show){ if(!workArea2) return; workArea2.classList.toggle('hidden',!show); workArea2.setAttribute('aria-busy', show ? 'true' : 'false'); }
function setButtonsEnabled(ready){ if(cleanBtn) cleanBtn.disabled=!ready; }
function afterCleanButtons(){ var ready=(cleanedRows.length>0); if(downloadXLSX) downloadXLSX.disabled=!ready; if(downloadXLSX2) downloadXLSX2.disabled=!ready; sticky.classList.toggle('show',ready); stickyMsg.textContent= ready ? 'Cleaned results ready.' : 'Results ready.'; }
function cycleStatus(i){ if(statusText) statusText.textContent=FRIENDLY_MSGS[i%FRIENDLY_MSGS.length]; }
function formatNum(n){ return new Intl.NumberFormat().format(n==null?0:n); }
function normalizeHeader(h){ if(h==null) return ''; return String(h).trim().replace(/\s+/g,' ').replace(/[_.-]+/g,' ').toLowerCase(); }
var HEADER_ALIASES={'company id':'Company ID','company name':'Company Name','contact name':'Contact Name','first name':'First Name','lastname':'Last Name','last name':'Last Name','surname':'Last Name','email':'Email','email address':'Email','e-mail':'Email','phone':'Contact Phone','contact phone':'Contact Phone','phone number':'Contact Phone','industry value':'Industry Value','industry':'Industry Value','website':'Website','url':'Website','domain':'Website','city':'City','state':'State','province':'State'};
function coerceRowToSchema(row){ var out={}; for(var key in row){ var nk=normalizeHeader(key); var mapped=HEADER_ALIASES[nk]; if(mapped) out[mapped]=row[key]; } for(var i=0;i<REQUIRED_COLS.length;i++){ var c=REQUIRED_COLS[i]; if(!(c in out)) out[c]=''; } return out; }

// ---- Sheet name helpers ----
var FORBIDDEN_CHARS=':/\\?*[]';
function hasForbidden(name){ var s=String(name||''); for(var i=0;i<s.length;i++){ if(FORBIDDEN_CHARS.indexOf(s[i])!==-1) return true; } return false; }
function isValidSheetName(name){ var s=String(name||''); return s.length>0 && s.length<=31 && !hasForbidden(s); }
function sanitizeSheetName(name){
  var s=String(name||''); var out='';
  for(var i=0;i<s.length;i++){ var ch=s[i]; if(FORBIDDEN_CHARS.indexOf(ch)!==-1) ch=' '; out+=ch; }
  s=out.trim();
  while(s.length && (s[0]==="'" || s[s.length-1]==="'")){
    if(s[0]==="'") s=s.substring(1);
    if(s.endsWith("'")) s=s.substring(0,s.length-1);
    s=s.trim();
  }
  if(s.length===0) s='Sheet';
  if(s.length>31) s=s.slice(0,31);
  return s;
}

// ---- Static irrelevant (kept for tests) ----
function matchesIndustry(val){ if(val==null) return false; return /(architect|owner|planroom|engineer|equipment)/i.test(String(val)); }
function isEduEmail(email){ var e=String(email||'').trim().toLowerCase(); return e.endsWith('.edu'); }
function isMilEmail(email){ var e=String(email||'').trim().toLowerCase(); return e.endsWith('.mil'); }
const EMAIL_RX=/^[^@\s]+@[^@\s]+\.[^@\s]+$/i; // basic validity
function isValidEmail(e){ return EMAIL_RX.test(String(e||'')); }

// ---- Dynamic prospect regex & denylist ----
function escapeRegex(s){ return String(s).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function splitCleanComma(text){ return (text||'').split(',').map(function(s){return s.trim();}).filter(Boolean); }
function buildKeywordRegex(){
  var raw=splitCleanComma(kwInput.value).map(function(s){return escapeRegex(s);});
  return new RegExp('\\b(' + raw.join('|') + ')\\b','i');
}
function buildDenySet(){
  var raw=splitCleanComma(denyInput.value).map(function(s){return s.toLowerCase();});
  var set={}; for(var i=0;i<raw.length;i++){ set[raw[i]]=1; } return set;
}
var DYNAMIC_PROSPECT_RE; var DENY_SET;
function refreshSettings(){ DYNAMIC_PROSPECT_RE=buildKeywordRegex(); DENY_SET=buildDenySet(); renderChips(); renderCounts(); renderRegexPreview(); }
kwInput.addEventListener('input', refreshSettings);
denyInput.addEventListener('input', refreshSettings);
refreshSettings();

function renderCounts(){ if(kwCount) kwCount.textContent = splitCleanComma(kwInput.value).length; if(denyCount) denyCount.textContent = splitCleanComma(denyInput.value).length; }
function renderChips(){
  if(kwChips){ kwChips.innerHTML=''; splitCleanComma(kwInput.value).slice(0,14).forEach(function(t){ var el=document.createElement('span'); el.className='chip'; el.textContent=t; kwChips.appendChild(el); }); }
  if(denyChips){ denyChips.innerHTML=''; splitCleanComma(denyInput.value).slice(0,14).forEach(function(t){ var el=document.createElement('span'); el.className='chip'; el.textContent=t; denyChips.appendChild(el); }); }
}
function renderRegexPreview(){ if(!regexPreview) return; var src=buildKeywordRegex().source; regexPreview.textContent='Regex preview: '+src; regexPreview.classList.toggle('hidden', splitCleanComma(kwInput.value).length===0); }

// Preset handling
Array.prototype.forEach.call(document.querySelectorAll('.preset'), function(el){ el.addEventListener('click', function(){ var add=splitCleanComma(el.getAttribute('data-preset')); var label=el.textContent.toLowerCase(); if(label.indexOf('temp')>-1||label.indexOf('disposable')>-1){ var cur=splitCleanComma(denyInput.value); cur=cur.concat(add).filter(function(v,i,a){return a.indexOf(v)===i;}); denyInput.value=cur.join(', ');} else { var curK=splitCleanComma(kwInput.value); curK=curK.concat(add).filter(function(v,i,a){return a.indexOf(v)===i;}); kwInput.value=curK.join(', ');} refreshSettings(); }); });

function domainFromEmail(e){ var m=String(e||'').toLowerCase().match(/@([^@]+)$/); return m?m[1]:''; }
function denyDomain(email){ var d=domainFromEmail(email); return !!DENY_SET[d]; }

function isProspectIndustry(val){ if(val==null) return false; return DYNAMIC_PROSPECT_RE.test(String(val)); }
function isIrrelevantRow(row){
  var ind=String(row['Industry Value']||'');
  var em=String(row['Email']||'');
  return matchesIndustry(ind) || isProspectIndustry(ind) || isEduEmail(em) || isMilEmail(em) || denyDomain(em);
}

// ---- Sheet selection (Part 1) ----
function findTargetSheetNames(wb){
  var names=wb.SheetNames||[]; var lower=[]; for(var i=0;i<names.length;i++){ lower.push(String(names[i]).toLowerCase()); }
  var targets=[];
  for(i=0;i<names.length;i++){ if(lower[i]==='complete leads' && isValidSheetName(names[i])) targets.push(names[i]); }
  for(i=0;i<names.length;i++){ if(lower[i]==='leads title' && isValidSheetName(names[i])) targets.push(names[i]); }
  if(!targets.length){
    for(i=0;i<names.length;i++){ if(isValidSheetName(names[i])){ targets=[names[i]]; break; } }
    if(!targets.length && names.length) targets=[names[0]];
  }
  return targets;
}

// ---------- UI reset ----------
function clearBadges(){ if(p1Badges) p1Badges.innerHTML=''; if(prevBadge) prevBadge.classList.add('hidden'); if(newBadge) newBadge.classList.add('hidden'); prevLabel.textContent='No file selected'; newLabel.textContent='No file selected'; }
function resetUI(){
  setButtonsEnabled(false);
  if(downloadXLSX) downloadXLSX.disabled=true; if(downloadXLSX2) downloadXLSX2.disabled=true; sticky.classList.remove('show');
  if(rowsBeforeEl) rowsBeforeEl.textContent='—';
  if(rowsAfterEl)  rowsAfterEl.textContent='—';
  if(dupePctEl)    dupePctEl.textContent='—';
  if(uniqueCompaniesEl) uniqueCompaniesEl.textContent='—';
  if(segmentCountEl) segmentCountEl.textContent='—';
  setProgress(0);
  if(statusText) statusText.textContent='Merging your files…';
  showWorkArea(true);
  // Part 2
  prevEmailsSet=null; newRelevantRowsP2=[]; newIrrelevantRowsP2=[]; newOnlyRel=[]; newOnlyIrrel=[];
  if(prevCountEl) prevCountEl.textContent='—';
  if(newCountEl) newCountEl.textContent='—';
  if(newOnlyRelCountEl) newOnlyRelCountEl.textContent='—';
  if(newOnlyIrrelCountEl) newOnlyIrrelCountEl.textContent='—';
  setProgress2(0);
  if(statusText2) statusText2.textContent='Waiting for files…';
  showWorkArea2(false);
}

function startOver(){
  try{
    rawRows=[]; cleanedRows=[]; segmentRows=[]; removedDupRows=[];
    if(fileInput) fileInput.value=''; clearBadges();
    if(drop) drop.classList.remove('dragging');
    setButtonsEnabled(false);
    if(downloadXLSX) downloadXLSX.disabled=true; if(downloadXLSX2) downloadXLSX2.disabled=true; sticky.classList.remove('show');
    if(rowsBeforeEl) rowsBeforeEl.textContent='—';
    if(rowsAfterEl) rowsAfterEl.textContent='—';
    if(dupePctEl) dupePctEl.textContent='—';
    if(uniqueCompaniesEl) uniqueCompaniesEl.textContent='—';
    if(segmentCountEl) segmentCountEl.textContent='—';
    setProgress(0);
    if(statusText) statusText.textContent='Waiting for files…';
    if(workArea) workArea.classList.add('hidden');

    // Part 2 reset
    if(prevFileEl) prevFileEl.value=''; if(newFileEl) newFileEl.value='';
    prevEmailsSet=null; newRelevantRowsP2=[]; newIrrelevantRowsP2=[]; newOnlyRel=[]; newOnlyIrrel=[];
    if(prevCountEl) prevCountEl.textContent='—'; if(newCountEl) newCountEl.textContent='—';
    if(newOnlyRelCountEl) newOnlyRelCountEl.textContent='—'; if(newOnlyIrrelCountEl) newOnlyIrrelCountEl.textContent='—';
    if(compareBtn) compareBtn.disabled=true; if(downloadCompare) downloadCompare.disabled=true;
    setProgress2(0); if(workArea2) workArea2.classList.add('hidden');

    refreshIcons();
  }catch(e){ var box=document.getElementById('testOut'); if(box){ box.innerHTML+='\n❌ Reset error: '+(e&&e.message?e.message:e)+'\n'; } }
}

// ---------- File badges ----------
function renderP1Badges(files){ if(!p1Badges) return; p1Badges.innerHTML='';
  for(var i=0;i<files.length;i++){
    var span=document.createElement('span'); span.className='badge'; span.innerHTML='<i data-feather="file"></i>'+files[i].name+' <span class="x" title="Remove">✖</span>';
    (function(idx){ span.querySelector('.x').addEventListener('click', function(){ fileInput.value=''; p1Badges.innerHTML=''; }); })(i);
    p1Badges.appendChild(span);
  }
  refreshIcons();
}

// ---------- File ingestion (Part 1) ----------
function readFileAsArrayBuffer(file){ return new Promise(function(resolve,reject){ var reader=new FileReader(); reader.onload=function(){ resolve(reader.result); }; reader.onerror=function(err){ reject(err); }; reader.readAsArrayBuffer(file); }); }

async function handleFiles(files){
  try{
    if(!files||!files.length) return;
    if(typeof XLSX==='undefined') ensureXLSX();
    resetUI(); cycleStatus(0); setProgress(2); rawRows=[]; removedDupRows=[];
    renderP1Badges(files);
    for(var i=0;i<files.length;i++){
      var f=files[i];
      cycleStatus(i);
      if(statusText) statusText.textContent='Merging your files… ('+(i+1)+'/'+files.length+')';
      var wb; try{ wb=XLSX.read(await readFileAsArrayBuffer(f),{type:'array'}); }catch(err){ console.error('XLSX read failed',err); continue; }
      var targets=findTargetSheetNames(wb);
      for(var t=0;t<targets.length;t++){
        var sname=targets[t]; var ws=wb.Sheets[sname]; if(!ws) continue;
        var rows=XLSX.utils.sheet_to_json(ws,{defval:'',raw:true});
        for(var r=0;r<rows.length;r++){ rawRows.push(coerceRowToSchema(rows[r])); }
      }
      setProgress(((i+1)/files.length)*45);
    }
    if(rowsBeforeEl) rowsBeforeEl.textContent=formatNum(rawRows.length);
    setButtonsEnabled(rawRows.length>0);
    if(statusText) statusText.textContent = rawRows.length ? 'Files merged. Click “Clean Data”.' : 'No rows detected. Check your files.';
  }catch(err){ var box=document.getElementById('testOut'); if(box){ box.innerHTML+='\n❌ Ingestion error: '+(err&&err.message?err.message:err)+'\n'; } throw err; }
}

if(drop){
  ['dragenter','dragover'].forEach(function(evt){ drop.addEventListener(evt,function(e){ e.preventDefault(); drop.classList.add('dragging'); }); });
  ['dragleave','drop'].forEach(function(evt){ drop.addEventListener(evt,function(e){ e.preventDefault(); drop.classList.remove('dragging'); }); });
  drop.addEventListener('drop',function(e){ var files=(e&&e.dataTransfer&&e.dataTransfer.files)?Array.prototype.slice.call(e.dataTransfer.files):[]; if(files.length) handleFiles(files); });
}
if(fileInput){ fileInput.addEventListener('change',function(e){ var files=(e&&e.target&&e.target.files)?Array.prototype.slice.call(e.target.files):[]; if(files.length) handleFiles(files); }); }

// ---------- Cleaning (Part 1) with Web Worker ----------
function createWorker(){
  var code=`self.onmessage=e=>{
  const {rows, keywordSource, denyDomains}=e.data;
  const EMAIL_RX=/^[^@\\s]+@[^@\\s]+\\.[^@\\s]+$/i;
  const kw = new RegExp('\\\\b('+keywordSource+')\\\\b','i');
  const denySet={}; denyDomains.split(',').map(s=>s.trim().toLowerCase()).filter(Boolean).forEach(d=>denySet[d]=1);
  function isEduEmail(v){v=(v||'').toLowerCase().trim(); return v.endsWith('.edu');}
  function isMilEmail(v){v=(v||'').toLowerCase().trim(); return v.endsWith('.mil');}
  function denyDomain(email){const m=(email||'').toLowerCase().match(/@([^@]+)$/);return m?!!denySet[m[1]]:false;}
  function matchesIndustry(val){return /(architect|owner|planroom|engineer|equipment)/i.test(String(val||''));}
  const seen={}; const out=[]; const irr=[]; const removed=[];
  const n=rows.length;
  for(let i=0;i<n;i++){
    const r=rows[i];
    const email=String(r['Email']||'').trim();
    if(!email || !EMAIL_RX.test(email)){ removed.push({...r,_Reason: !email? 'Empty email':'Invalid email'}); continue; }
    const key=email.toLowerCase();
    if(seen[key]){ removed.push({...r,_Reason:'Duplicate email'}); continue; }
    seen[key]=1;
    const ind=String(r['Industry Value']||'');
    const irrelevant = matchesIndustry(ind) || kw.test(ind) || isEduEmail(email) || isMilEmail(email) || denyDomain(email);
    out.push(r); if(irrelevant) irr.push(r);
    if((i%4000)===0) postMessage({progress: Math.round((i/n)*100)});
  }
  postMessage({done:true, out, irr, removed});
};`;
  return new Worker(URL.createObjectURL(new Blob([code],{type:'text/javascript'})));
}

function cleanData(){
  if(!rawRows.length) return;
  if(rawRows.length>200000){ if(!confirm('Large dataset ('+formatNum(rawRows.length)+' rows). Proceed with full diagnostics?')){ return; } }
  showWorkArea(true); setProgress(5);
  if(statusText) statusText.textContent='Initializing cleaner…';
  var worker=createWorker();
  worker.onmessage=function(e){
    if(e.data.progress!=null){ setProgress(5 + e.data.progress*0.9); if(statusText) statusText.textContent='Optimizing your list… '+e.data.progress+'%'; return; }
    if(e.data.done){
      cleanedRows=e.data.out||[]; segmentRows=e.data.irr||[]; removedDupRows=e.data.removed||[];
      var before=rawRows.length, after=cleanedRows.length, removed=before-after, pct=before?(removed/before*100):0;
      if(rowsAfterEl) rowsAfterEl.textContent=formatNum(after);
      if(dupePctEl) dupePctEl.textContent = removed + ' (' + pct.toFixed(1) + '%)';
      var compSet={}; for(var k=0;k<cleanedRows.length;k++){ var id=String(cleanedRows[k]['Company ID']||'').trim(); if(id) compSet[id]=1; }
      var uniq=0; for(var kk in compSet){ if(Object.prototype.hasOwnProperty.call(compSet,kk)) uniq++; }
      if(uniqueCompaniesEl) uniqueCompaniesEl.textContent=formatNum(uniq);
      if(segmentCountEl) segmentCountEl.textContent=formatNum(segmentRows.length);
      if(statusText) statusText.textContent= segmentRows.length? ('Clean complete! Irrelevant matches: '+formatNum(segmentRows.length)) : 'Clean complete! No irrelevant matches under current rules.';
      setProgress(100); afterCleanButtons(); refreshIcons();
    }
  };
  worker.postMessage({ rows: rawRows, keywordSource: buildKeywordRegex().source.replace(/^(?:\(\?:)?|\)$/g,''), denyDomains: denyInput.value||'' });
}
if(cleanBtn){ cleanBtn.addEventListener('click', function(){ showWorkArea(true); cleanData(); }); }
if(resetBtn){ resetBtn.addEventListener('click', startOver); }

// ---------- Download workbook (Part 1) ----------
function simpleHash(s){ s=String(s); var h=0x811c9dc5; for(var i=0;i<s.length;i++){ h^=s.charCodeAt(i); h = (h>>>0) * 16777619; } return ('0000000'+(h>>>0).toString(16)).slice(-8); }
function rowsToSheet(rows, includeHash){
  var data=rows.map(function(r){ if(!includeHash) return r; var copy=Object.assign({},r); copy._RowHash=simpleHash(JSON.stringify(r)); return copy; });
  return XLSX.utils.json_to_sheet(data,{header: includeHash? (REQUIRED_COLS.concat(['_RowHash'])) : REQUIRED_COLS});
}
function downloadWorkbook(){
  try{
    ensureXLSX(); if(!cleanedRows.length) return;
    var relevant=[], irrelevant=segmentRows.slice();
    for(var i=0;i<cleanedRows.length;i++){ var rr=cleanedRows[i]; if(!isIrrelevantRow(rr)) relevant.push(rr); }
    var ws1=rowsToSheet(relevant,false);
    var ws2=rowsToSheet(irrelevant,false);
    var ws3=rowsToSheet(removedDupRows,true);
    var wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, sanitizeSheetName('Relevant Data'));
    XLSX.utils.book_append_sheet(wb, ws2, sanitizeSheetName('Irrelevant Data'));
    XLSX.utils.book_append_sheet(wb, ws3, sanitizeSheetName('Removed Duplicates'));
    XLSX.writeFile(wb,'EmailList_Relevant_vs_Irrelevant.xlsx');
  }catch(err){ var box=document.getElementById('testOut'); if(box){ box.innerHTML+='\n❌ Download error: '+(err&&err.message?err.message:err)+'\n'; } throw err; }
}
if(downloadXLSX){ downloadXLSX.addEventListener('click', downloadWorkbook); }
if(downloadXLSX2){ downloadXLSX2.addEventListener('click', downloadWorkbook); }

// ---------- Sample data ----------
if(sampleBtn){ sampleBtn.addEventListener('click', function(){
  try{
    ensureXLSX();
    var csv=[
      'Company ID,Company Name,Contact Name,First Name,Last Name,Email,Contact Phone,Industry Value,Website,City,State',
      'C-1001,Acme Inc.,Jane Roe,Jane,Roe,jane@acme.com,555-1234,Owner Operator,acme.com,Denver,CO',
      'C-1002,Globex,John Smith,John,Smith,JOHN@globex.com,555-5678,Architectural Services,globex.com,Austin,TX',
      'C-1002,Globex,John Smith,John,Smith,john@globex.com,555-5678,Planroom Portal,globex.com,Austin,TX',
      'C-1003,Initech,Peter Gibbons,Peter,Gibbons,,555-8888,Tech,initech.com,Dallas,TX',
      'C-1004,Alpha Co.,Amy Lee,Amy,Lee,amy@alpha.edu,555-2222,Education,alpha.edu,Boston,MA',
      'C-1005,Bravo Equip,Mark Toner,Mark,Toner,mark@bravoequip.com,555-3333,Heavy Equipment Rentals,bravo.com,Phoenix,AZ',
      'C-1006,City Works,Sam Ray,Sam,Ray,sam@army.mil,555-4444,municipal,city.gov,DC,DC'
    ].join('\n');
    var wb=XLSX.read(csv,{type:'string'}); var ws=wb.Sheets[wb.SheetNames[0]]; var rows=XLSX.utils.sheet_to_json(ws,{defval:'',raw:true});
    rawRows=[]; for(var i=0;i<rows.length;i++){ rawRows.push(coerceRowToSchema(rows[i])); }
    if(rowsBeforeEl) rowsBeforeEl.textContent=formatNum(rawRows.length);
    if(statusText) statusText.textContent='Sample loaded. Click “Clean Data”.';
    showWorkArea(true); setButtonsEnabled(true); refreshIcons();
  }catch(err){ var box=document.getElementById('testOut'); if(box){ box.innerHTML+='\n❌ Sample load error: '+(err&&err.message?err.message:err)+'\n'; } throw err; }
}); }

// ====================== Part 2: Compare New vs Previous ======================
function setP2Enabled(){
  var ready = prevFileEl && prevFileEl.files && prevFileEl.files.length>0 &&
              newFileEl && newFileEl.files && newFileEl.files.length>0;
  if(compareBtn) compareBtn.disabled = !ready;
  if(!ready){ if(downloadCompare) downloadCompare.disabled=true; }
}

function clearPrevSel(){ prevFileEl.value=''; prevBadge.classList.add('hidden'); prevLabel.textContent='No file selected'; setP2Enabled(); refreshIcons(); }
function clearNewSel(){ newFileEl.value=''; newBadge.classList.add('hidden'); newLabel.textContent='No file selected'; setP2Enabled(); refreshIcons(); }
if(clearPrev){ clearPrev.addEventListener('click', clearPrevSel); }
if(clearNew){ clearNew.addEventListener('click', clearNewSel); }

if(prevFileEl){
  prevFileEl.addEventListener('change', function(){
    var f = prevFileEl.files && prevFileEl.files[0];
    prevLabel.textContent = f ? f.name : 'No file selected';
    if(f){ prevBadge.textContent=f.name; prevBadge.classList.remove('hidden'); } else { prevBadge.classList.add('hidden'); }
    setP2Enabled();
  });
}
if(newFileEl){
  newFileEl.addEventListener('change', function(){
    var f = newFileEl.files && newFileEl.files[0];
    newLabel.textContent = f ? f.name : 'No file selected';
    if(f){ newBadge.textContent=f.name; newBadge.classList.remove('hidden'); } else { newBadge.classList.add('hidden'); }
    setP2Enabled();
  });
}

function setP2Progress(p){ setProgress2(p); }
function setP2Status(txt){ if(statusText2) statusText2.textContent = txt; }

function emailKey(v){ return String(v||'').trim().toLowerCase(); }

// Find sheet by name fragment (case-insensitive), fallback to first match with headers
function findSheetByNameFragment(wb, fragment){
  var frag = String(fragment||'').toLowerCase();
  var names = wb.SheetNames || [];
  for(var i=0;i<names.length;i++){
    var nm=String(names[i]||''); if(nm.toLowerCase().indexOf(frag)!==-1) return nm;
  }
  // fallback: choose a sheet that has at least Email header
  for(i=0;i<names.length;i++){
    var ws=wb.Sheets[names[i]]; if(!ws) continue;
    var rows = XLSX.utils.sheet_to_json(ws,{defval:'',raw:true,range:0});
    if(rows && rows.length){
      var headers = Object.keys(rows[0] || {});
      var hasEmail=false;
      for(var h=0;h<headers.length;h++){ if(normalizeHeader(headers[h])==='email'){ hasEmail=true; break; } }
      if(hasEmail) return names[i];
    }
  }
  return names[0] || null;
}

// Read ALL emails from a workbook (union of sheets)
function readAllEmailsFromWorkbook(wb){
  var names = wb.SheetNames || [];
  var set = {};
  for(var i=0;i<names.length;i++){
    var ws = wb.Sheets[names[i]]; if(!ws) continue;
    var rows = XLSX.utils.sheet_to_json(ws,{defval:'',raw:true});
    for(var r=0;r<rows.length;r++){
      var coerced = coerceRowToSchema(rows[r]);
      var em = emailKey(coerced.Email);
      if(em) set[em]=1;
    }
  }
  return set;
}

// Read rows from “Relevant” and “Irrelevant” sheets in the new processed workbook
function readRelevantIrrelevantFromNewProcessed(wb){
  var relName = findSheetByNameFragment(wb,'relevant');
  var irrName = findSheetByNameFragment(wb,'irrelevant');
  var relRows=[], irrRows=[];

  if(relName){
    var ws1 = wb.Sheets[relName];
    var rows1 = XLSX.utils.sheet_to_json(ws1,{defval:'',raw:true});
    for(var i=0;i<rows1.length;i++){ relRows.push(coerceRowToSchema(rows1[i])); }
  }
  if(irrName){
    var ws2 = wb.Sheets[irrName];
    var rows2 = XLSX.utils.sheet_to_json(ws2,{defval:'',raw:true});
    for(var j=0;j<rows2.length;j++){ irrRows.push(coerceRowToSchema(rows2[j])); }
  }
  return {rel:relRows, irr:irrRows};
}

async function comparePreviousAndNew(){
  try{
    if(!(prevFileEl && prevFileEl.files && prevFileEl.files.length) ||
       !(newFileEl && newFileEl.files && newFileEl.files.length)){
      alert('Please select both Previous and New Processed workbooks.'); return;
    }
    ensureXLSX();
    showWorkArea2(true);
    setP2Progress(5); setP2Status('Reading previous workbook…');

    // Read prev workbook
    var prevWb = XLSX.read(await readFileAsArrayBuffer(prevFileEl.files[0]),{type:'array'});
    prevEmailsSet = readAllEmailsFromWorkbook(prevWb);
    var prevTotal = 0; for(var k in prevEmailsSet){ if(Object.prototype.hasOwnProperty.call(prevEmailsSet,k)) prevTotal++; }
    if(prevCountEl) prevCountEl.textContent = formatNum(prevTotal);

    setP2Progress(35); setP2Status('Reading new processed workbook…');
    var newWb = XLSX.read(await readFileAsArrayBuffer(newFileEl.files[0]),{type:'array'});
    var groups = readRelevantIrrelevantFromNewProcessed(newWb);
    newRelevantRowsP2 = groups.rel;
    newIrrelevantRowsP2 = groups.irr;

    var newTotal = 0;
    // count unique emails in new (union across both sheets)
    var tmpSet={};
    for(var i=0;i<newRelevantRowsP2.length;i++){ var e1 = emailKey(newRelevantRowsP2[i].Email); if(e1) tmpSet[e1]=1; }
    for(i=0;i<newIrrelevantRowsP2.length;i++){ var e2 = emailKey(newIrrelevantRowsP2[i].Email); if(e2) tmpSet[e2]=1; }
    for(var z in tmpSet){ if(Object.prototype.hasOwnProperty.call(tmpSet,z)) newTotal++; }
    if(newCountEl) newCountEl.textContent = formatNum(newTotal);

    setP2Progress(55); setP2Status('Computing New-Only (by Email)…');
    newOnlyRel = [];
    newOnlyIrrel = [];
    for(i=0;i<newRelevantRowsP2.length;i++){
      var r = newRelevantRowsP2[i]; var em = emailKey(r.Email);
      if(em && !prevEmailsSet[em]) newOnlyRel.push(r); // include entire row from NEW
    }
    for(i=0;i<newIrrelevantRowsP2.length;i++){
      var rr = newIrrelevantRowsP2[i]; var em2 = emailKey(rr.Email);
      if(em2 && !prevEmailsSet[em2]) newOnlyIrrel.push(rr); // include entire row from NEW
    }
    if(newOnlyRelCountEl) newOnlyRelCountEl.textContent = formatNum(newOnlyRel.length);
    if(newOnlyIrrelCountEl) newOnlyIrrelCountEl.textContent = formatNum(newOnlyIrrel.length);

    setP2Progress(85); setP2Status('Ready to download New-Only workbook.');
    if(downloadCompare) downloadCompare.disabled = (newOnlyRel.length+newOnlyIrrel.length)===0 ? true : false;
  }catch(err){
    var box=document.getElementById('testOut');
    if(box){ box.innerHTML+='\n❌ Compare error: '+(err&&err.message?err.message:err)+'\n'; }
    throw err;
  }finally{
    setP2Progress(100);
  }
}

function downloadNewOnlyWorkbook(){
  try{
    ensureXLSX();
    var ws1 = XLSX.utils.json_to_sheet(newOnlyRel,{header:REQUIRED_COLS});
    var ws2 = XLSX.utils.json_to_sheet(newOnlyIrrel,{header:REQUIRED_COLS});
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws1, sanitizeSheetName('Relevant — New Only'));
    XLSX.utils.book_append_sheet(wb, ws2, sanitizeSheetName('Irrelevant — New Only'));
    XLSX.writeFile(wb, 'EmailList_NewOnly.xlsx');
  }catch(err){
    var box=document.getElementById('testOut');
    if(box){ box.innerHTML+='\n❌ Download New-Only error: '+(err&&err.message?err.message:err)+'\n'; }
    throw err;
  }
}
if(compareBtn){ compareBtn.addEventListener('click', comparePreviousAndNew); }
if(downloadCompare){ downloadCompare.addEventListener('click', downloadNewOnlyWorkbook); }

// ---------- Self-tests ----------
function fastCleanRows(rows){
  var seen={}; var out=[];
  for(var i=0;i<rows.length;i++){
    var r=rows[i]; var o={};
    for(var cIdx=0;cIdx<REQUIRED_COLS.length;cIdx++){ var c=REQUIRED_COLS[cIdx]; o[c]=r[c]??''; }
    var email=(o.Email||'').trim(); if(!email || !isValidEmail(email)) continue;
    var k=email.toLowerCase(); if(seen[k]) continue; seen[k]=1;
    out.push(o);
  }
  return out;
}
function assertEqual(name,a,b){
  var pass=JSON.stringify(a)===JSON.stringify(b);
  var el=document.getElementById('testOut');
  el.innerHTML+="\n"+(pass?'✅':'❌')+" "+name+(pass?'':("\n   expected: "+JSON.stringify(b)+"\n   got     : "+JSON.stringify(a)));
  return pass;
}
function runSelfTests(){
  var out=document.getElementById('testOut'); out.textContent='Running self-tests...'; var passed=0,total=0;
  total++; passed += assertEqual('T1 removes empty email rows', fastCleanRows([{Email:'a@x.com','Company ID':'C1'},{Email:'','Company ID':'C2'}]).length, 1)?1:0;
  var out2=fastCleanRows([{'Company ID':'C1',Email:'A@x.com'},{'Company ID':'C2',Email:'a@x.com'}]); total++; passed += assertEqual('T2 case-insensitive dedupe keeps first', out2.map(function(r){return r['Company ID'];}), ['C1'])?1:0;
  var out3=fastCleanRows([{'Company ID':'U1',Email:'u1@x.com'},{'Company ID':'U2',Email:'u2@x.com'},{'Company ID':'U2',Email:'dup@x.com'}]); var compSet={}; for(var i=0;i<out3.length;i++){ var id=String(out3[i]['Company ID']).trim(); if(id) compSet[id]=1; } var uniq=0; for(var k in compSet){ if(Object.prototype.hasOwnProperty.call(compSet,k)) uniq++; }
  total++; passed += assertEqual('T3 unique company IDs', uniq, 2)?1:0;
  var out4=fastCleanRows([{Email:'p@x.com'}])[0]; var allCols=true; for(var i2=0;i2<REQUIRED_COLS.length;i2++){ if(!(REQUIRED_COLS[i2] in out4)) { allCols=false; break; } } total++; passed += assertEqual('T4 ensures required columns exist', allCols, true)?1:0;
  var pref1=findTargetSheetNames({SheetNames:['Leads Title','Complete Leads']}); var pref2=findTargetSheetNames({SheetNames:['Random','Something']}); total++; passed += assertEqual('T5a prefers Complete Leads', pref1[0],'Complete Leads')?1:0; total++; passed += assertEqual('T5b falls back to first sheet', pref2[0],'Random')?1:0;
  var segSample=[
    {'Industry Value':'Owner Operator',Email:'1@x.com'},
    {'Industry Value':'civil ENGINEER',Email:'2@x.com'},
    {'Industry Value':'Architectural firm',Email:'3@x.com'},
    {'Industry Value':'Heavy Equipment Rentals',Email:'4@x.com'},
    {'Industry Value':'Retail',Email:'5@x.com'},
    {'Industry Value':'PLANROOM Updates',Email:'6@x.com'},
    {'Industry Value':'Project Architect',Email:'7@x.com'},
    {'Industry Value':'Owner-Rep Services',Email:'8@x.com'}
  ];
  var segOut=segSample.filter(function(r){ return matchesIndustry(r['Industry Value']); });
  total++; passed += assertEqual('T6 industry segment filter', segOut.map(function(r){return r.Email;}), ['1@x.com','2@x.com','3@x.com','4@x.com','6@x.com','7@x.com','8@x.com'])?1:0;

  // Prospect & TLD rules
  total++; passed += assertEqual('T13 prospect keywords (legal/designer/food service/photo graphic)', [
    isProspectIndustry('Legal Services'),
    isProspectIndustry('Senior Designer'),
    isProspectIndustry('Food Service Supply'),
    isProspectIndustry('Photo Graphic Studio')
  ], [true,true,true,true])?1:0;
  total++; passed += assertEqual('T14 municipal and mmunicipal variants', [isProspectIndustry('municipal'), isProspectIndustry('mmunicipal')], [true,true])?1:0;
  total++; passed += assertEqual('T15 .mil emails are irrelevant', [isMilEmail('user@army.mil'), isMilEmail('user@army.mil.com')], [true,false])?1:0;
  total++; passed += assertEqual('T19 email validity rejects bad formats', [isValidEmail('a@b.com'), isValidEmail('no-at-domain'), isValidEmail('a@b')], [true,false,false])?1:0;
  total++; passed += assertEqual('T20 configurable keywords (unicorn)', (function(){ var old=kwInput.value; kwInput.value='unicorn'; refreshSettings(); var ok=isProspectIndustry('great unicorn firm'); kwInput.value=old; refreshSettings(); return ok; })(), true)?1:0;

  // Start over reset check
  total++; (function(){
    startOver();
    var ok=(cleanBtn && cleanBtn.disabled===true) &&
           (downloadXLSX && downloadXLSX.disabled===true) &&
           (rowsBeforeEl && rowsBeforeEl.textContent==='—') &&
           (workArea && workArea.classList.contains('hidden'));
    passed += assertEqual('T16 startOver resets UI', ok, true)?1:0;
  })();

  // Extra safety tests
  total++; var dwEnabled=!downloadXLSX.disabled; passed += assertEqual('T7 button disabled before clean', dwEnabled, false)?1:0;
  total++; passed += assertEqual('T8 planroom lower/upper', [matchesIndustry('planroom'), matchesIndustry('PLANROOM')], [true,true])?1:0;
  total++; passed += assertEqual('T9 sheet name sanitization', sanitizeSheetName('Irrelevant Data : * [ ] / \\ ?'), 'Irrelevant Data')?1:0;
  total++; passed += assertEqual('T10 sheet name validity', [isValidSheetName('Relevant Data'), isValidSheetName('Bad/Name')], [true,false])?1:0;
  var long='ThisIsAVeryLongWorksheetNameThatExceeds31Characters'; total++; passed += assertEqual('T11 sheet name max length', sanitizeSheetName(long).length<=31, true)?1:0;
  var eduRow={Email:'dept@college.edu','Industry Value':'Retail'}; total++; passed += assertEqual('T12 .edu falls into irrelevant', isIrrelevantRow(eduRow), true)?1:0;

  out.innerHTML += '\n\nResult: '+passed+'/'+total+' tests passed.';
}
if(testBtn){ testBtn.addEventListener('click', runSelfTests); }

// Final icon refresh after dynamic DOM ops
document.addEventListener('DOMContentLoaded', refreshIcons);

// Hook UI
if(cleanBtn){ cleanBtn.addEventListener('click', function(){ showWorkArea(true); cleanData(); }); }
if(resetBtn){ resetBtn.addEventListener('click', startOver); }
if(downloadXLSX){ downloadXLSX.addEventListener('click', downloadWorkbook); }
if(downloadXLSX2){ downloadXLSX2.addEventListener('click', downloadWorkbook); }
if(compareBtn){ compareBtn.addEventListener('click', comparePreviousAndNew); }
if(downloadCompare){ downloadCompare.addEventListener('click', downloadNewOnlyWorkbook); }
