let workbookMeta, qsData, timeData, qualityData;
const qsMetrics = ["Academic Reputation","Employer Reputation","Faculty Student","Citations per Faculty","International Faculty","International Students","International Students Diversity","International Research Network","Employment Outcomes","Sustainability"];

async function api(url, opts={}){ const r=await fetch(url,opts); if(!r.ok){let d={};try{d=await r.json()}catch{};throw new Error(d.detail||`HTTP ${r.status}`)} return r.json(); }
function esc(v){return String(v ?? "").replace(/[&<>"']/g,s=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[s]));}
function showToast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),1800)}
function parseRank(r){ if(!r) return null; const m=String(r).match(/(\d+)/); return m?Number(m[1]):null; }

async function reloadAll(){
  [workbookMeta,qsData,timeData,qualityData]=await Promise.all([api('/api/workbook'),api('/api/qs'),api('/api/time'),api('/api/quality')]);
  document.getElementById('source-file').textContent=workbookMeta.file;
  renderOverview(); renderQS(); renderTime(); renderQuality(); showToast('Excel data refreshed');
}

function renderOverview(){
  const latestQS=[...qsData.records].sort((a,b)=>b.Year-a.Year)[0];
  const latestTHE=[...timeData.records].sort((a,b)=>b.Year-a.Year)[0];
  const populated2027=latestTHE ? Object.entries(latestTHE).filter(([k,v])=>!['Year','Rank','Name','excel_row'].includes(k)&&v!==null&&v!=='').length : 0;
  const kpis=[
    ['Latest QS Rank',latestQS?.Rank ?? '—',`QS ${latestQS?.Year ?? ''}`],
    ['Latest THE Rank',latestTHE?.Rank ?? '—',`THE ${latestTHE?.Year ?? ''}`],
    ['QS Metrics',qsData.groups.length,'Score + rank indicator groups'],
    ['Excel Formulas',workbookMeta.formula_count,workbookMeta.formula_count?'Read directly from workbook':'None in current workbook']
  ];
  document.getElementById('kpis').innerHTML=kpis.map(k=>`<div class="kpi"><div class="label">${esc(k[0])}</div><div class="value">${esc(k[1])}</div><div class="meta">${esc(k[2])}</div></div>`).join('');

  const ranks=[...qsData.records].sort((a,b)=>a.Year-b.Year); const max=Math.max(...ranks.map(x=>parseRank(x.Rank)||1));
  document.getElementById('qs-rank-trend').innerHTML=ranks.map(x=>{const n=parseRank(x.Rank)||max;const pct=Math.max(15,100-(n/max*70));return `<div class="rank-row"><span class="rank-year">${x.Year}</span><div class="track"><span style="width:${pct}%"></span></div><span class="rank-value">${esc(x.Rank)}</span></div>`}).join('');

  const t=[...timeData.records].sort((a,b)=>b.Year-a.Year).find(r=>r.Teaching!==null&&r.Teaching!=='')||{};
  const keys=['Teaching','Research Environment','Research Quality','Industry','International Outlook'];
  document.getElementById('time-bars').innerHTML=keys.map(k=>{let n=Number(String(t[k]).replace('..','.'))||0;return `<div class="bar-row"><span>${esc(k)}</span><div class="bar-track"><span style="width:${Math.min(100,n)}%"></span></div><b>${esc(t[k]??'—')}</b></div>`}).join('');

  document.getElementById('formula-notice').innerHTML=workbookMeta.formula_count
    ? `<strong>Excel formula integration active</strong><p>${workbookMeta.formula_count} formula cell(s) were detected. Formula text and Excel's cached values are fetched from the workbook.</p>`
    : `<strong>No formulas detected in the supplied workbook</strong><p>The application does not invent formulas. If formulas are later added to the Excel file, the backend will expose them automatically and protect those cells from web edits.</p>`;
}

function renderQS(){
  const top=`<thead><tr><th rowspan="2">Rank</th><th rowspan="2">Year</th>${qsData.groups.map(g=>`<th colspan="2">${esc(g.metric)}</th>`).join('')}</tr><tr>${qsData.groups.map(()=>'<th>Score</th><th>Rank</th>').join('')}</tr></thead>`;
  const body=qsData.records.map(r=>{
    let html=`<tr><td class="editable" data-sheet="QS Ranking" data-cell="A${r.excel_row}">${esc(r.Rank)}</td><td class="editable" data-sheet="QS Ranking" data-cell="B${r.excel_row}">${esc(r.Year)}</td>`;
    qsData.groups.forEach((g,i)=>{const m=r.metrics[g.metric]||{}; const ci=3+i*2; const sc=col(ci)+r.excel_row, rc=col(ci+1)+r.excel_row; html+=cellHTML('QS Ranking',sc,m.Score)+cellHTML('QS Ranking',rc,m.Rank)}); return html+'</tr>';
  }).join('');
  const table=document.getElementById('qs-table'); table.innerHTML=top+`<tbody>${body}</tbody>`; bindEditable(table);
}
function renderTime(){
  const cols=timeData.headers; const top=`<thead><tr>${cols.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead>`;
  const body=timeData.records.map(r=>`<tr>${cols.map((h,i)=>cellHTML('TIME Rankings',col(i+1)+r.excel_row,r[h])).join('')}</tr>`).join('');
  const table=document.getElementById('time-table'); table.innerHTML=top+`<tbody>${body}</tbody>`; bindEditable(table);
}
function cellHTML(sheet,cell,v){const s=String(v??'');const cls=['editable']; if(v===null||v==='')cls.push('missing'); if(/^\d+\.\.\d+$/.test(s)||/\d+=$/.test(s))cls.push('issue');return `<td class="${cls.join(' ')}" data-sheet="${esc(sheet)}" data-cell="${cell}">${v===null||v===''?'—':esc(v)}</td>`}
function col(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}

function bindEditable(table){
 table.querySelectorAll('td.editable').forEach(td=>td.ondblclick=()=>{
   if(td.querySelector('input'))return; const old=td.textContent==='—'?'':td.textContent; const inp=document.createElement('input');inp.className='cell-input';inp.value=old;td.textContent='';td.appendChild(inp);inp.focus();inp.select();
   const save=async()=>{let val=inp.value.trim(); if(/^-?\d+(\.\d+)?$/.test(val)) val=Number(val); try{await api('/api/cell',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({sheet:td.dataset.sheet,cell:td.dataset.cell,value:val===''?null:val})});showToast(`Saved ${td.dataset.cell}`);await reloadAll()}catch(e){showToast(e.message);td.textContent=old||'—'}};
   inp.onkeydown=e=>{if(e.key==='Enter')save();if(e.key==='Escape')td.textContent=old||'—'}; inp.onblur=()=>{setTimeout(()=>{if(document.body.contains(inp))td.textContent=old||'—'},150)};
 });
}
function renderQuality(){
 const list=document.getElementById('quality-list'); const issues=qualityData.issues||[];
 if(!issues.length){list.innerHTML='<div class="quality-empty"><b>No obvious text-format anomalies detected.</b></div>';return}
 list.innerHTML=issues.map(x=>`<div class="quality-item"><b>${esc(x.sheet)} · ${esc(x.cell)} · “${esc(x.value)}”</b><p>${esc(x.reason)}. Review the original source before changing it.</p></div>`).join('');
}

document.querySelectorAll('.nav-btn').forEach(btn=>btn.addEventListener('click',()=>{document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');document.querySelectorAll('.view').forEach(v=>v.classList.remove('active-view'));document.getElementById(btn.dataset.view).classList.add('active-view');document.getElementById('page-title').textContent=btn.dataset.view==='overview'?'2027 Rankings Analysis':btn.textContent; }));
reloadAll().catch(e=>showToast(e.message));
