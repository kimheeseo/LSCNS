(function browserPatch(){
'use strict';
const PHYSICAL_SPEED=400;
const AGG_SPEED=1200;
const LANES=3;

function q(id){return document.getElementById(id)}
function selected(role){
 const el=q('v48-'+role+'-speed');
 return el&&el.value!=='auto'?Number(el.value):null;
}
function isAgg(role){return selected(role)===AGG_SPEED}
function label(role){
 const v=selected(role);
 if(v===AGG_SPEED)return'1.2T Aggregate · 3×400G';
 if(v)return v+'G';
 return'Auto · system reference';
}
function rememberBase(obj,key,val){
 const k='__v64_'+key;
 if(!Object.prototype.hasOwnProperty.call(obj,k)){
   try{Object.defineProperty(obj,k,{value:val,writable:true,configurable:true,enumerable:false})}
   catch(_){obj[k]=val}
 }
 return obj[k];
}

if(typeof window.getSelectedSystem==='function'&&!window.getSelectedSystem.__v64wrapped){
 const prev=window.getSelectedSystem;
 const wrapped=function(){
   const s=prev.apply(this,arguments);
   if(!s)return s;
   const baseLinks=rememberBase(s,'baseLinks',Number(s.links||0));
   const baseCages=rememberBase(s,'baseCages',s.physicalClusterCages==null?null:Number(s.physicalClusterCages));
   if(isAgg('compute')){
     s.logicalLinkSpeed=AGG_SPEED;
     s.linkSpeed=PHYSICAL_SPEED;
     s.links=baseLinks*LANES;
     s.physicalClusterCages=s.links;
     s.aggregateNetworkProfile={mode:'1.2T Aggregate',logicalSpeed:AGG_SPEED,physicalSpeed:PHYSICAL_SPEED,physicalLinksPerLogical:LANES,native:false};
   }else{
     s.links=baseLinks;
     if(baseCages==null)delete s.physicalClusterCages;else s.physicalClusterCages=baseCages;
     s.logicalLinkSpeed=s.linkSpeed;
     s.aggregateNetworkProfile=null;
   }
   s.networkRoles=s.networkRoles||{};
   ['compute','storage','inband'].forEach(role=>{
     const v=selected(role);
     if(v===AGG_SPEED)s.networkRoles[role]={speed:AGG_SPEED,logicalSpeed:AGG_SPEED,physicalSpeed:PHYSICAL_SPEED,aggregate:true,lanes:LANES,native:false};
     else if(v)s.networkRoles[role]={speed:v,logicalSpeed:v,physicalSpeed:v,aggregate:false,lanes:1,native:true};
   });
   return s;
 };
 wrapped.__v64wrapped=true;
 window.getSelectedSystem=wrapped;
}

if(typeof window.resolveAuxSystemProfile==='function'&&!window.resolveAuxSystemProfile.__v64wrapped){
 const prev=window.resolveAuxSystemProfile;
 const wrapped=function(){
   const p=prev.apply(this,arguments);if(!p)return p;
   const out={...p,storage:{...(p.storage||{})},inband:{...(p.inband||{})},oob:{...(p.oob||{})}};
   if(isAgg('storage')){
     out.storage.logicalSpeed=AGG_SPEED;out.storage.speed=PHYSICAL_SPEED;out.storage.aggregate=true;out.storage.lanes=LANES;out.storage.native=false;
     out.storage.autoSwitch='sn5610';out.storage.ethernetSwitch='sn5610';out.storage.ibSwitch='qm9700';out.storage.media='1.2T Aggregate · 3×400G physical links';
   }
   if(isAgg('inband')){
     out.inband.logicalSpeed=AGG_SPEED;out.inband.speed=PHYSICAL_SPEED;out.inband.aggregate=true;out.inband.lanes=LANES;out.inband.native=false;
     out.inband.switch='sn5610';out.inband.media='1.2T Aggregate · 3×400G Ethernet';
   }
   return out;
 };
 wrapped.__v64wrapped=true;
 window.resolveAuxSystemProfile=wrapped;
}

function numberFrom(id){
 const e=q(id);if(!e)return 0;
 const m=((e.innerText||e.textContent||e.value||'')+'').replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);
 return m?Number(m[0]):0;
}
function ensureOptions(){
 ['compute','storage','inband'].forEach(role=>{
   const el=q('v48-'+role+'-speed');if(!el)return;
   let o=Array.from(el.options).find(x=>x.value==='1200');
   if(!o){o=document.createElement('option');o.value='1200';o.textContent='1.2T Aggregate · 3×400G';el.appendChild(o)}
   o.disabled=false;
 });
}
function updateRoleUi(){
 const host=q('v48-role-model');if(!host)return;
 ensureOptions();
 const note=q('v48-role-note');
 if(note){
   const any=isAgg('compute')||isAgg('storage')||isAgg('inband');
   note.innerHTML='<b>Active roles:</b> Compute '+label('compute')+' · Storage '+label('storage')+' · In-Band '+label('inband')+' · OOB 1G'+
   (any?'<br><span class="v48-warn"><b>1.2T Aggregate:</b> modeled as 3 × 400G physical links per 1.2 Tb/s logical service. Native 1.2TbE optics/cages are not assumed.</span>':'');
 }
 let x=q('v64-1p2t-note');
 if(!x){x=document.createElement('div');x.id='v64-1p2t-note';host.appendChild(x)}
 x.innerHTML='<b>1.2T modeling rule</b> · 1.2T Aggregate = 3 × 400G physical. Switch cages, optics and cables are therefore sized from 400G physical links. Exact vendor breakout/harness support remains a verification item.';
}
function updatePhysicalAudit(){
 const p=q('v48-physical-audit');if(!p)return;
 const rows=Array.from(p.querySelectorAll('tbody tr'));
 const row=rows.find(r=>(r.cells[0]&&r.cells[0].textContent||'').includes('System'));
 if(!row||!isAgg('compute'))return;
 const units=numberFrom('mUnits');
 let s=null;try{s=window.getSelectedSystem&&window.getSelectedSystem()}catch(_){}
 if(!s)return;
 const physical=units*Number(s.links||0);
 const spare=1+(Number(q('sparePct')&&q('sparePct').value||0)/100);
 const designPhysical=Math.ceil(physical*spare);
 const logical=Math.ceil(physical/LANES),designLogical=Math.ceil(designPhysical/LANES);
 if(row.cells[1])row.cells[1].innerHTML=logical+' × 1.2T logical / '+physical+' × 400G physical<br>'+designLogical+' × 1.2T incl. spare / '+designPhysical+' × 400G physical';
 if(row.cells[6])row.cells[6].innerHTML='1.2T Aggregate → 3×400G <span class="v64-agg-tag">non-native</span>';
}
function refresh(){
 updateRoleUi();
 setTimeout(updatePhysicalAudit,0);
}
document.addEventListener('change',e=>{if(e.target&&/^v48-(compute|storage|inband)-speed$/.test(e.target.id)){setTimeout(()=>{try{if(typeof window.refreshSystem==='function')window.refreshSystem();else if(typeof window.calc==='function')window.calc()}catch(_){}refresh()},0)}},true);
document.addEventListener('click',()=>setTimeout(refresh,60),true);
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',refresh,{once:true});else refresh();
setInterval(refresh,1200);
console.log('v6.4.1 1.2T Aggregate model active: 1.2T logical = 3x400G physical.');
})();