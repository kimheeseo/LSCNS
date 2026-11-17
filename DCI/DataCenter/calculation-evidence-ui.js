(function(){'use strict';
const I=window.DCAuditI18n,V=window.DCDesignValidation;if(!I||!V)return;
const $=id=>document.getElementById(id),esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let result=null,validation=null,mode='waiting',page=0,query='',focusId='',opener=null;
const lang=()=>I.languages.includes(document.documentElement.lang)?document.documentElement.lang:'en';
const t=(k,p)=>I.translate(lang(),k,p),number=n=>Number.isFinite(n)?n.toLocaleString(lang(),{maximumFractionDigits:6}):String(n??'—');
function value(v){if(Array.isArray(v))return '['+v.slice(0,8).map(x=>typeof x==='object'?'…':number(x)).join(', ')+(v.length>8?', …':'')+'] ('+v.length+' '+t('entries')+')';if(typeof v==='object'&&v)return '{…}';return typeof v==='number'?number(v):String(v??'—');}
function ready(){return mode==='current'&&result?.evidence?.kind==='calculated';}
function inputs(){
 const target=$('engineering-inputs');if(!target||$('audit-facility-controls'))return;
 const box=document.createElement('div');box.id='audit-facility-controls';box.className='audit-facility-controls';
 box.innerHTML='<label><span data-audit-text="upsScope"></span><select id="upsLoadScope"><option value="it" data-audit-text="it"></option><option value="facility" data-audit-text="facility"></option></select></label>'+['ups','generator','transformer'].map(k=>'<label><span data-audit-text="'+k+'Installed"></span><input id="'+k+'InstalledQty" type="number" min="0" step="1" value="0"><small data-audit-text="manualHint"></small></label>').join('');target.append(box);
 const previous=window.getEngineeringInputs;if(typeof previous==='function')window.getEngineeringInputs=()=>{const i=previous();i.upsLoadScope=$('upsLoadScope').value;for(const k of ['ups','generator','transformer'])i[k+'InstalledQty']=Number($(k+'InstalledQty').value);return i;};
}
function mount(){
 const parent=document.querySelector('.resultPanel');if(!parent||$('calculation-audit'))return;
 const section=document.createElement('section');section.id='calculation-audit';section.setAttribute('aria-live','polite');
 section.innerHTML='<header><h3 data-audit-text="validation"></h3><span class="audit-kind"></span><button id="audit-open" type="button" aria-haspopup="dialog" aria-controls="audit-evidence-dialog" data-audit-text="open"></button></header><p class="audit-state" role="status"></p><div class="audit-body"><div class="audit-counts"></div><details><summary data-audit-text="validationDetails"></summary><div class="tableWrap"><table class="audit-table"><thead><tr><th data-audit-text="rule"></th><th data-audit-text="status"></th><th data-audit-text="reason"></th><th data-audit-text="suggestion"></th><th data-audit-text="show"></th></tr></thead><tbody id="audit-rules"></tbody></table></div></details></div><p class="audit-disclaimer" data-audit-text="disclaimer"></p>';
 parent.append(section);
 const dialog=document.createElement('dialog');dialog.id='audit-evidence-dialog';dialog.setAttribute('aria-labelledby','audit-dialog-title');
 dialog.innerHTML='<div class="audit-drawer"><header class="audit-drawer-head"><h2 id="audit-dialog-title" data-audit-text="title"></h2><button id="audit-close" type="button" data-audit-text="close"></button></header><p id="audit-drawer-state" class="audit-state"></p><div class="audit-drawer-tools"><input id="audit-search" type="search"><button id="audit-download" type="button" data-audit-text="download"></button></div><nav class="audit-jumps">'+[['racks','compute.units'],['fabric','fabric.leaf'],['optical','server.route'],['itPower','power.IT'],['facilityPower','power.facility'],['electrical','facility.upsDesign']].map(([key,id])=>'<button type="button" data-audit-ref="'+id+'" data-audit-text="'+key+'"></button>').join('')+'</nav><main class="audit-drawer-content" id="audit-steps"></main><nav class="audit-paging"><button id="audit-prev" type="button" data-audit-text="prev"></button><span id="audit-page"></span><button id="audit-next" type="button" data-audit-text="next"></button></nav></div>';
 document.body.append(dialog);
 $('audit-open').addEventListener('click',()=>open());$('audit-close').addEventListener('click',()=>dialog.close());
 dialog.addEventListener('close',()=>opener?.focus());dialog.addEventListener('click',e=>{if(e.target===dialog){const b=dialog.getBoundingClientRect();if(e.clientX<b.left||e.clientX>b.right||e.clientY<b.top||e.clientY>b.bottom)dialog.close();}});
 document.addEventListener('click',e=>{const a=e.target.closest('[data-audit-ref]');if(a){focusId=a.dataset.auditRef;open(focusId);}});
 $('audit-search').addEventListener('input',e=>{query=e.target.value.toLowerCase();page=0;drawSteps();});
 $('audit-prev').addEventListener('click',()=>{page=Math.max(0,page-1);drawSteps();});$('audit-next').addEventListener('click',()=>{page++;drawSteps();});
 $('audit-download').addEventListener('click',()=>{if(!ready())return;const blob=new Blob([JSON.stringify({evidence:result.evidence,validation},null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='DC-BOM-v8.2.0-calculation-evidence.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);});
 dialog.addEventListener('toggle',e=>{const raw=e.target;if(raw.matches?.('[data-audit-raw]')&&raw.open&&!raw.querySelector('pre')){const pre=document.createElement('pre');pre.className='audit-raw';const step=result?.evidence?.steps?.find(x=>x.id===raw.dataset.auditRaw);pre.textContent=JSON.stringify(step||result?.evidence?.inputSources,null,2);raw.append(pre);}},true);
 inputs();paint();if(window.DCDesign)accept(window.DCDesign);
}
function localize(){document.querySelectorAll('[data-audit-text]').forEach(e=>e.textContent=t(e.dataset.auditText));if($('audit-search')){$('audit-search').placeholder=t('search');$('audit-search').setAttribute('aria-label',t('search'));}const sample=$('entry-sample-result');if(sample){sample.dataset.auditKind='sample';const label=sample.querySelector('.entry-sample-heading span');if(label)label.textContent=t('sample');sample.setAttribute('aria-label',t('sample'));}}
function paint(){if(!$('calculation-audit'))return;localize();const root=$('calculation-audit');root.dataset.mode=mode;root.querySelector('.audit-kind').textContent=t(ready()?'current':result?'previous':'sample');root.querySelector('.audit-state').textContent=t(mode==='current'?'current':mode);$('audit-open').disabled=!ready();$('audit-download').disabled=!ready();$('audit-drawer-state').textContent=t(mode==='current'?'current':mode);
 root.querySelector('.audit-counts').innerHTML=validation?Object.entries(validation.counts).map(([s,n])=>'<span class="audit-'+s+'">'+s+' <b>'+number(n)+'</b></span>').join(''):'';
 $('audit-rules').innerHTML=validation?validation.rules.map(row=>{const params=Object.fromEntries(Object.entries(row.values).map(([k,v])=>[k,typeof v==='number'?number(v):k==='scope'?t(v):String(v)]));return '<tr><th scope="row">'+esc(row.id)+'</th><td class="audit-'+row.status+'">'+row.status+'</td><td>'+esc(t(row.reason,params))+'</td><td>'+esc(t(row.suggestion))+'</td><td>'+row.refs.map(id=>'<button type="button" data-audit-ref="'+esc(id)+'" '+(!ready()?'disabled':'')+'>'+esc(t('show'))+'</button>').join(' ')+'</td></tr>';}).join(''):'';
 if($('audit-evidence-dialog').open)drawSteps();
}
function accept(r){result=r;validation=V.evaluate(r);mode=r?.evidence?.kind==='calculated'?'current':'unavailable';page=0;query='';if($('audit-search'))$('audit-search').value='';paint();}
function open(id=''){if(!ready())return;opener=document.activeElement;focusId=id;if(id){query='';$('audit-search').value='';const idx=result.evidence.steps.findIndex(s=>s.id===id);page=idx>=0?Math.floor(idx/40):0;}drawSteps();const d=$('audit-evidence-dialog');if(!d.open)d.showModal();if(id)requestAnimationFrame(()=>document.querySelector('[data-step-id="'+CSS.escape(id)+'"]')?.scrollIntoView({block:'start'}));}
function drawSteps(){
 if(!result?.evidence)return;
 const steps=result.evidence.steps.filter(s=>!query||JSON.stringify([s.id,s.label,s.expression,s.inputs.map(v=>v.name)]).toLowerCase().includes(query)),pages=Math.max(1,Math.ceil(steps.length/40));page=Math.max(0,Math.min(page,pages-1));
 const block=steps.slice(page*40,page*40+40).map(s=>{
  let substituted=s.expression;for(const a of [...s.inputs].sort((a,b)=>b.name.length-a.name.length))substituted=substituted.split(a.name).join(value(a.value));
  return '<article class="audit-step" data-step-id="'+esc(s.id)+'"><h4>'+esc(s.id)+' · '+esc(s.label)+'</h4><dl><dt>'+t('formula')+'</dt><dd><code>'+esc(s.expression)+'</code></dd><dt>'+t('substitution')+'</dt><dd><code>'+esc(substituted)+'</code></dd><dt>'+t('result')+'</dt><dd><b>'+number(s.value)+'</b> '+esc(s.unit)+'</dd><dt>'+t('source')+'</dt><dd>'+esc(s.source)+'</dd></dl><ul class="audit-inputs">'+s.inputs.map(a=>'<li><b>'+esc(a.name)+'</b> = '+esc(value(a.value))+(a.assumption?'<span class="audit-badge">'+t('assumption')+'</span>':'')+'<br>'+t('source')+': '+esc(a.origin==='input'?t('project')+' · '+(a.field||a.name):a.source)+(a.ref?' <button type="button" data-audit-ref="'+esc(a.ref)+'">'+t('derived')+' '+esc(a.ref)+'</button>':'')+'</li>').join('')+'</ul><details data-audit-raw="'+esc(s.id)+'"><summary>'+t('raw')+'</summary></details></article>';
 }).join('');
 $('audit-steps').innerHTML=(page===0?'<details data-audit-raw="inputs"><summary>'+t('normalized')+'</summary></details>':'')+block;
 $('audit-page').textContent=number(page+1)+' / '+number(pages)+' · '+number(steps.length);$('audit-prev').disabled=page===0;$('audit-next').disabled=page+1>=pages;
}
document.addEventListener('dc:design',e=>accept(e.detail));document.addEventListener('dc:calculating',()=>{mode='pending';paint();});document.addEventListener('dc:stale',()=>{mode='stale';paint();});
for(const event of ['input','change'])document.addEventListener(event,e=>{if(e.target.closest('.formPanel')&&!e.target.dataset.referenceOnly){mode='stale';paint();}});
document.addEventListener('dc:inputs-ready',()=>{mount();inputs();localize();});new MutationObserver(()=>paint()).observe(document.documentElement,{attributes:true,attributeFilter:['lang']});
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount,{once:true});else mount();
})();
