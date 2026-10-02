(function browserPatch(){
'use strict';
const AGG_12=1200,AGG_16=1600;
const MEDIA=['auto','DAC','AEC','AOC','SR','DR','FR'];
const BREAKOUT=[
 ['auto','Auto / native mapping'],
 ['800-2x400','800G cage → 2×400G logical'],
 ['1200-3x400','1.2T aggregate → 3×400G physical'],
 ['1600-2x800','1.6T aggregate → 2×800G physical']
];
function q(id){return document.getElementById(id)}
function text(e){return(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function uiLang(){
 const raw=String(window.__dcBomUiLang||document.documentElement.getAttribute('data-dc-bom-ui-lang')||document.documentElement.lang||'ko').toLowerCase();
 if(raw.includes('en'))return'en';if(raw.includes('ja')||raw.includes('jp'))return'ja';if(raw.includes('zh')||raw.includes('cn'))return'zh';if(raw.includes('de'))return'de';return'ko';
}
const KOREA_LABEL={ko:'한국 채널',en:'Korea Channel',ja:'韓国チャネル',zh:'韩国渠道',de:'Korea-Kanal'};
const KOREA_TITLE={ko:'한국 구매 / 기술 문의 채널',en:'Korea Purchasing / Technical Contacts',ja:'韓国の購入 / 技術問い合わせ先',zh:'韩国采购 / 技术联系渠道',de:'Korea Einkauf / Technische Kontakte'};
function selectedSpeed(role){
 const e=q('v48-'+role+'-speed');if(!e||e.value==='auto')return null;return Number(e.value)||null;
}
function media(role){const e=q('v67-'+role+'-media');return e?e.value:'auto'}
function breakout(role){const e=q('v67-'+role+'-breakout');return e?e.value:'auto'}
function policy(role){return{speed:selectedSpeed(role),media:media(role),breakout:breakout(role)}}
function allPolicies(){return{compute:policy('compute'),storage:policy('storage'),inband:policy('inband')}}
window.__dcBomLinkPolicy=window.__dcBomLinkPolicy||allPolicies();

function ensure1600Options(){
 ['compute','storage','inband'].forEach(role=>{
  const e=q('v48-'+role+'-speed');if(!e)return;
  let o=Array.from(e.options).find(x=>x.value==='1600');
  if(!o){o=document.createElement('option');o.value='1600';o.textContent='1.6T Aggregate · 2×800G';e.appendChild(o)}
  o.disabled=false;
 });
}
function addLinkControls(){
 const host=q('v48-role-model');if(!host)return false;
 ensure1600Options();
 let box=q('v67-link-controls');
 if(!box){box=document.createElement('section');box.id='v67-link-controls';host.insertAdjacentElement('afterend',box)}
 const old={};
 ['compute','storage','inband'].forEach(r=>{old[r]={media:media(r),breakout:breakout(r)}});
 box.innerHTML='<h3>Network Link Engineering <span class="v67-review">role-based</span></h3>'+
 '<div class="v67-sub">Speed, link media/optic class and breakout are independent per network role. 1.2T/1.6T are aggregate planning profiles unless an exact native product is selected.</div>'+
 '<div class="v67-role-grid">'+['compute','storage','inband'].map(r=>'<div class="v67-role-box"><b>'+r.toUpperCase()+'</b>'+
 '<label>Link media / optic class</label><select id="v67-'+r+'-media">'+MEDIA.map(x=>'<option value="'+x+'">'+(x==='auto'?'Auto by distance / product':x)+'</option>').join('')+'</select>'+
 '<label>Breakout / aggregation</label><select id="v67-'+r+'-breakout">'+BREAKOUT.map(x=>'<option value="'+x[0]+'">'+x[1]+'</option>').join('')+'</select></div>').join('')+
 '</div><table class="v67-cap-table"><thead><tr><th>Logical service</th><th>Nominal replacement capacity</th><th>Physical planning realization</th><th>Meaning</th></tr></thead><tbody>'+
 '<tr><td>400G</td><td>1 × 400G</td><td>400G native</td><td>Base unit</td></tr>'+
 '<tr><td>800G</td><td>2 × 400G</td><td>800G native or verified 2×400G breakout</td><td>Product support required for breakout</td></tr>'+
 '<tr><td>1.2T Aggregate</td><td>3 × 400G</td><td>3 × 400G physical links</td><td><span class="v67-review">non-native planning</span></td></tr>'+
 '<tr><td>1.6T Aggregate</td><td>4 × 400G = 2 × 800G</td><td>2 × 800G physical links</td><td><span class="v67-review">non-native planning</span></td></tr></tbody></table>'+
 '<div class="v67-sub" id="v67-policy-note" style="margin-top:8px"></div>';
 ['compute','storage','inband'].forEach(r=>{
  const me=q('v67-'+r+'-media'),be=q('v67-'+r+'-breakout');
  if(me)me.value=old[r].media||'auto';if(be)be.value=old[r].breakout||'auto';
  [me,be].filter(Boolean).forEach(e=>e.addEventListener('change',()=>{window.__dcBomLinkPolicy=allPolicies();renderPolicyNote()}));
 });
 renderPolicyNote();
 return true;
}
function renderPolicyNote(){
 const e=q('v67-policy-note');if(!e)return;
 const p=allPolicies();window.__dcBomLinkPolicy=p;
 e.innerHTML='<b>Active link policy:</b> Compute '+(p.compute.speed||'Auto')+'G / '+p.compute.media+' / '+p.compute.breakout+
 ' · Storage '+(p.storage.speed||'Auto')+'G / '+p.storage.media+' / '+p.storage.breakout+
 ' · In-Band '+(p.inband.speed||'Auto')+'G / '+p.inband.media+' / '+p.inband.breakout+
 '<br><b>Rule:</b> DAC/AEC/AOC/SR/DR/FR and breakout selections are engineering constraints; exact optics/connectors/cages remain product-compatibility gated.';
}

if(typeof window.getSelectedSystem==='function'&&!window.getSelectedSystem.__v67wrapped){
 const prev=window.getSelectedSystem;
 const wrapped=function(){
  const s=prev.apply(this,arguments);if(!s)return s;
  const speed=selectedSpeed('compute');
  if(speed===AGG_16){
   const base=Number(s.__v64_baseLinks!=null?s.__v64_baseLinks:s.links||0);
   s.logicalLinkSpeed=AGG_16;s.linkSpeed=800;s.links=base*2;s.physicalClusterCages=s.links;
   s.aggregateNetworkProfile={mode:'1.6T Aggregate',logicalSpeed:1600,physicalSpeed:800,physicalLinksPerLogical:2,native:false};
  }
  s.linkPolicies=allPolicies();
  return s;
 };
 wrapped.__v67wrapped=true;window.getSelectedSystem=wrapped;
}
if(typeof window.resolveAuxSystemProfile==='function'&&!window.resolveAuxSystemProfile.__v67wrapped){
 const prev=window.resolveAuxSystemProfile;
 const wrapped=function(){
  const p=prev.apply(this,arguments);if(!p)return p;
  const out={...p,storage:{...(p.storage||{})},inband:{...(p.inband||{})},oob:{...(p.oob||{})}};
  ['storage','inband'].forEach(role=>{
   const v=selectedSpeed(role);
   if(v===AGG_16){out[role].logicalSpeed=1600;out[role].speed=800;out[role].aggregate=true;out[role].lanes=2;out[role].native=false;out[role].media='1.6T Aggregate · 2×800G physical links'}
   out[role].requestedMedia=media(role);out[role].requestedBreakout=breakout(role);
  });
  return out;
 };
 wrapped.__v67wrapped=true;window.resolveAuxSystemProfile=wrapped;
}

function patchPhysicalAudit(){
 const table=q('v48-physical-audit');if(!table)return;
 const compute=selectedSpeed('compute');
 const rows=table.querySelectorAll('tbody tr');
 if(compute===AGG_16&&rows[0]){
  const cells=rows[0].children;
  if(cells[6])cells[6].innerHTML='1.6T logical service → 2×800G physical <span class="v67-review">aggregate</span>';
 }
 const p=allPolicies();
 let note=table.querySelector('.v67-audit-note');
 if(!note){note=document.createElement('div');note.className='v67-audit-note v67-sub';note.style.marginTop='8px';table.appendChild(note)}
 note.textContent='Requested media: Compute '+p.compute.media+' / Storage '+p.storage.media+' / In-Band '+p.inband.media+'. Breakout mapping is valid only when the selected endpoint and switch explicitly support it.';
}

function ensureKoreaButton(){
 const bar=q('v63-top-links');if(!bar)return false;
 let b=q('v67-korea-channel-btn');
 if(!b){b=document.createElement('button');b.type='button';b.id='v67-korea-channel-btn';bar.appendChild(b);b.addEventListener('click',openKorea)}
 b.textContent=KOREA_LABEL[uiLang()]||KOREA_LABEL.ko;
 return true;
}
function modal(){
 let m=q('v67-korea-modal');if(m)return m;
 m=document.createElement('div');m.id='v67-korea-modal';m.hidden=true;
 m.innerHTML='<div class="v67-modal-card"><div class="v67-modal-head"><h3></h3><button class="v67-close" type="button">×</button></div><div class="v67-modal-body"></div></div>';
 document.body.appendChild(m);
 const close=()=>m.hidden=true;m.querySelector('.v67-close').onclick=close;m.addEventListener('click',e=>{if(e.target===m)close()});document.addEventListener('keydown',e=>{if(e.key==='Escape')close()});
 return m;
}
function openKorea(){
 const src=q('v47-channels'),m=modal();m.querySelector('h3').textContent=KOREA_TITLE[uiLang()]||KOREA_TITLE.ko;
 const body=m.querySelector('.v67-modal-body');
 body.innerHTML=src?src.innerHTML:'<p>Channel data is not available in the current build.</p>';
 m.hidden=false;
}
function refresh(){
 ensure1600Options();addLinkControls();renderPolicyNote();patchPhysicalAudit();ensureKoreaButton();
}
function start(){
 let attempts=0;const boot=setInterval(()=>{attempts++;refresh();if((q('v48-role-model')&&q('v63-top-links'))||attempts>30)clearInterval(boot)},400);
 document.addEventListener('change',e=>{if(e.target&&(/^v48-(compute|storage|inband)-speed$/.test(e.target.id)||/^v67-/.test(e.target.id)))setTimeout(()=>{refresh();try{if(typeof window.refreshSystem==='function')window.refreshSystem();else if(typeof window.calc==='function')window.calc()}catch(_){ }},0)},true);
 document.addEventListener('click',e=>{if(e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],.lang,.language'))setTimeout(refresh,80)},true);
 setInterval(()=>{ensureKoreaButton();patchPhysicalAudit()},1400);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();