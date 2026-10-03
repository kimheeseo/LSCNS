(() => {
'use strict';
const q=id=>document.getElementById(id);
const VERSION='7.3.2';
const POWER_SOURCES={
 h200:'https://docs.nvidia.com/dgx/dgxh100-user-guide/introduction-to-dgxh100.html',
 b200:'https://docs.nvidia.com/dgx/dgxb200-user-guide/introduction-to-dgxb200.html',
 b300:'https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html',
 b300Peak:'https://docs.nvidia.com/dgx-pdf/data-center-best-practices-with-dgx-b300-v1.pdf',
 b300Ra:'https://docs.nvidia.com/dgx-superpod/reference-architecture/scalable-infrastructure-b300-xdr/latest/dgx-superpod-architecture.html'
};
const ROLE_DEFAULTS={
 h200:{compute:400,storage:400,inband:100,oob:1},
 b200:{compute:400,storage:400,inband:200,oob:1},
 b300:{compute:800,storage:400,inband:200,oob:1},
 l4host:{compute:400,storage:100,inband:100,oob:1},
 l40shost:{compute:400,storage:100,inband:100,oob:1},
 rubin:{compute:800,storage:400,inband:200,oob:1}
};
const ROLE_CAPS={
 h200:{compute:[400],storage:[400],inband:[100,200],oob:[1]},
 b200:{compute:[400],storage:[400],inband:[200,400],oob:[1]},
 b300:{compute:[800],storage:[400],inband:[200],oob:[1]},
 l4host:{compute:[400],storage:[100,200,400],inband:[100,200,400],oob:[1]},
 l40shost:{compute:[400],storage:[100,200,400],inband:[100,200,400],oob:[1]},
 rubin:{compute:[800],storage:[400,800],inband:[200,400,800],oob:[1]}
};
const POWER_ENVELOPES={
 h200:{typical:null,designMax:10.2,peak:10.2,typicalBasis:'Not separately published in the referenced NVIDIA system power table',peakBasis:'Uses vendor max until a distinct verified provisioning envelope is supplied'},
 b200:{typical:null,designMax:14.3,peak:14.3,typicalBasis:'Not separately published in the referenced NVIDIA system power table',peakBasis:'Uses vendor max until a distinct verified provisioning envelope is supplied'},
 b300:{typical:14.5,designMax:15.0,peak:19.7,typicalBasis:'DGX B300 user guide power consumption',peakBasis:'Low-density 2-system rack peak 39.4 kW ÷ 2 = 19.7 kW/system'}
};

function txt(e){return (e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function systemId(){return q('systemId')?q('systemId').value:'h200'}
function roleDefault(role){const d=ROLE_DEFAULTS[systemId()]||{};return d[role]||400}
function roleCaps(role){const d=ROLE_CAPS[systemId()]||{},base=d[role]||[roleDefault(role)];return role==='oob'?base:Array.from(new Set(base.concat([1200,1600])))}
function syncRoleSelectors(){
 ['compute','storage','inband'].forEach(role=>{
   const el=document.getElementById('v48-'+role+'-speed');if(!el)return;
   const allowed=roleCaps(role);
   Array.from(el.options).forEach(o=>{if(o.value!=='auto')o.disabled=!allowed.includes(Number(o.value))});
   if(el.value!=='auto'&&!allowed.includes(Number(el.value)))el.value='auto';
 });
}
function selectedRoleSpeed(role){
 const el=document.getElementById('v48-'+role+'-speed');
 const def=roleDefault(role),allowed=roleCaps(role);
 if(!el||el.value==='auto')return def;
 const requested=Number(el.value);
 return allowed.includes(requested)?requested:def;
}
function powerEnvelope(s){
 const id=systemId();
 const e=POWER_ENVELOPES[id];
 if(e)return e;
 const d=Number(s&&s.power||0);
 return {typical:null,designMax:d,peak:d,typicalBasis:'Partner/OEM profile required',peakBasis:'No separate peak envelope; design value used'};
}
function applyPowerProfiles(){
 if(typeof systems==='undefined')return;
 if(systems.h200)Object.assign(systems.h200,{powerTypical:null,powerDesignMax:10.2,powerPeakProvisioning:10.2,power:10.2,powerNominal:null});
 if(systems.b200)Object.assign(systems.b200,{powerTypical:null,powerDesignMax:14.3,powerPeakProvisioning:14.3,power:14.3,powerNominal:null});
 if(systems.b300)Object.assign(systems.b300,{powerTypical:14.5,powerDesignMax:15.0,powerPeakProvisioning:19.7,power:15.0,powerNominal:14.5});
 if(typeof auxSwitches!=='undefined'&&typeof switches!=='undefined'&&switches.q3400&&!auxSwitches.q3400){
   auxSwitches.q3400={name:switches.q3400.name,fabric:'InfiniBand XDR',ru:switches.q3400.ru,power:switches.q3400.power,ports:{800:144,400:144},media:{800:'OSFP twin-port XDR',400:'OSFP twin-port'},confidence:'Confirmed NVIDIA'};
 }
}
function mountRoleModel(){
 if(document.getElementById('v48-role-model'))return;
 const storage=q('storageFabricType');if(!storage)return;
 const card=storage.closest('.card');if(!card)return;
 const div=document.createElement('div');div.id='v48-role-model';
 div.innerHTML='<h3>Network Role Model <span class="v48-badge">role-based</span></h3><div class="v48-sub">GPU 이름이나 “800G data center” 한 값으로 모든 포트를 바꾸지 않습니다. Compute / Storage / In-Band / OOB를 독립 role로 계산합니다.</div><div class="v48-role-grid">'+
 '<div class="v48-role-card"><label>Compute Fabric</label><select id="v48-compute-speed"><option value="auto">Auto · system reference</option><option value="400">400G</option><option value="800">800G</option><option value="1200">1.2T Aggregate · 3×400G</option><option value="1600">1.6T Aggregate · 2×800G</option></select></div>'+
 '<div class="v48-role-card"><label>Storage Fabric</label><select id="v48-storage-speed"><option value="auto">Auto · role reference</option><option value="100">100G</option><option value="200">200G</option><option value="400">400G</option><option value="800">800G</option><option value="1200">1.2T Aggregate · 3×400G</option><option value="1600">1.6T Aggregate · 2×800G</option></select></div>'+
 '<div class="v48-role-card"><label>In-Band Ethernet</label><select id="v48-inband-speed"><option value="auto">Auto · role reference</option><option value="100">100G</option><option value="200">200G</option><option value="400">400G</option><option value="800">800G</option><option value="1200">1.2T Aggregate · 3×400G</option><option value="1600">1.6T Aggregate · 2×800G</option></select></div>'+
 '<div class="v48-role-card"><label>Out-of-Band</label><input value="1G management" readonly></div></div><div class="v48-sub" id="v48-role-note" style="margin-top:8px"></div>';
 const warn=q('networkWarning');if(warn)card.insertBefore(div,warn);else card.appendChild(div);
 ['compute','storage','inband'].forEach(role=>document.getElementById('v48-'+role+'-speed').addEventListener('change',()=>{try{refreshSystem()}catch(e){try{calc()}catch(_){}}}));
 syncRoleSelectors();
 updateRoleNote();
}
function updateRoleNote(){
 const el=document.getElementById('v48-role-note');if(!el)return;
 const b300Note=systemId()==='b300'?' · B300 In-Band: 200G logical links (bonded); NVIDIA may use 400G-capable QSFP / 800G twin-port OSFP physical optics.':'';
 el.innerHTML='<b>Active roles:</b> Compute '+selectedRoleSpeed('compute')+'G · Storage '+selectedRoleSpeed('storage')+'G · In-Band '+selectedRoleSpeed('inband')+'G · OOB 1G'+b300Note+'.';
}

applyPowerProfiles();
mountRoleModel();

if(typeof getSelectedSystem==='function'){
 const _getSelectedSystem=getSelectedSystem;
 getSelectedSystem=function(){
   const s=_getSelectedSystem();
   syncRoleSelectors();
   const id=systemId(), env=powerEnvelope(s);
   s.linkSpeed=selectedRoleSpeed('compute');
   s.networkRoles={compute:{speed:s.linkSpeed},storage:{speed:selectedRoleSpeed('storage')},inband:{speed:selectedRoleSpeed('inband')},oob:{speed:1}};
   s.powerTypical=env.typical;s.powerDesignMax=env.designMax;s.powerPeakProvisioning=env.peak;s.power=env.designMax;
   return s;
 };
}
if(typeof resolveAuxSystemProfile==='function'){
 const _resolveAuxSystemProfile=resolveAuxSystemProfile;
 resolveAuxSystemProfile=function(s){
   const p=_resolveAuxSystemProfile(s);if(!p)return p;
   const out={...p,storage:{...p.storage},inband:{...p.inband},oob:{...p.oob}};
   out.storage.speed=selectedRoleSpeed('storage');
   out.inband.speed=selectedRoleSpeed('inband');
   if(out.storage.speed===800){
     out.storage.autoSwitch='sn5610';out.storage.ethernetSwitch='sn5610';out.storage.ibSwitch='q3400';out.storage.media='800G role-selected optical';
   }else if(out.storage.speed===400){
     out.storage.autoSwitch=out.storage.autoSwitch||'sn5610';out.storage.ethernetSwitch='sn5610';out.storage.ibSwitch='qm9700';out.storage.media='400G role-selected optical';
   }else{
     out.storage.autoSwitch='sn5610';out.storage.ethernetSwitch='sn5610';out.storage.ibSwitch='qm9700';out.storage.media=out.storage.speed+'G role-selected optical';
   }
   out.inband.switch=out.inband.speed<=100?'sn4600c':'sn5610';
   out.inband.media=out.inband.speed+'G Ethernet';
   return out;
 };
}
if(typeof resolveStorageSwitch==='function'){
 const _resolveStorageSwitch=resolveStorageSwitch;
 resolveStorageSwitch=function(profile){
   if(!profile)return null;
   const sel=q('storageFabricType')?q('storageFabricType').value:'auto';
   if(sel==='infiniband')return profile.storage.speed>=800?'q3400':'qm9700';
   if(sel==='ethernet')return 'sn5610';
   return profile.storage.autoSwitch||_resolveStorageSwitch(profile);
 };
}

if(typeof sizeRackPower==='function'){
 sizeRackPower=function(name,deviceGroups){
   const outletPerPdu=Math.max(1,+q('pduOutletCount').value||24);
   const pduRatedKw=Math.max(1,+q('pduRatedKw').value||60);
   const util=Math.min(1,Math.max(.01,(+q('pduUtilPct').value||80)/100));
   const usablePerPdu=pduRatedKw*util;
   let typicalPower=0,designPower=0,peakPower=0,typicalKnown=true,cordsA=0,cordsB=0,totalCords=0,assumptions=[],survival=[];
   deviceGroups.forEach(g=>{
     if(!g||!g.qty||!g.profile)return;const p=g.profile,qty=g.qty;
     const typical=Object.prototype.hasOwnProperty.call(p,'powerTypical')?p.powerTypical:(p.power!=null?p.power:null);
     const design=p.powerDesignMax!=null?p.powerDesignMax:(p.powerDesign!=null?p.powerDesign:(p.power||0));
     const peak=p.powerPeakProvisioning!=null?p.powerPeakProvisioning:design;
     if(typical==null)typicalKnown=false;else typicalPower+=typical*qty;
     designPower+=design*qty;peakPower+=peak*qty;
     if(p.cordCount==null){assumptions.push(g.label+': vendor rack/busbar power architecture required');survival.push(g.label+': '+(p.feedSurvival||'vendor review'));return}
     const cords=p.cordCount*qty,sp=splitAB(cords);totalCords+=cords;cordsA+=sp.a;cordsB+=sp.b;
     if(p.feedSurvival)survival.push(g.label+': '+p.feedSurvival);
   });
   const normalA=designPower/2,normalB=designPower/2,peakA=peakPower/2,peakB=peakPower/2;
   const pduA=Math.max(cordsA?Math.ceil(cordsA/outletPerPdu):0,peakA?Math.ceil(peakA/usablePerPdu):0);
   const pduB=Math.max(cordsB?Math.ceil(cordsB/outletPerPdu):0,peakB?Math.ceil(peakB/usablePerPdu):0);
   const capA=pduA*usablePerPdu,capB=pduB*usablePerPdu;
   const singleFeedFullLoad=peakPower===0?true:(capA>=peakPower&&capB>=peakPower);
   return{name,rackPower:designPower,typicalPower:typicalKnown?typicalPower:null,designPower,peakPower,totalCords,cordsA,cordsB,pduA,pduB,normalA,normalB,peakA,peakB,capA,capB,singleFeedFullLoad,outletPerPdu,pduRatedKw,util,assumptions,survival};
 };
}
if(typeof renderPowerSizing==='function'){
 renderPowerSizing=function(powerSizing){
   const el=q('powerSummary'),notes=q('powerNotes');if(!el||!notes)return;
   const rows=powerSizing.racks.map(r=>[r.name,r.typicalPower==null?'—':r.typicalPower.toFixed(1)+' kW',r.designPower.toFixed(1)+' kW',r.peakPower.toFixed(1)+' kW',r.totalCords||'Vendor rack power',r.totalCords?(r.cordsA+' / '+r.cordsB):'—',r.totalCords?(r.pduA+' / '+r.pduB):'Vendor design',(r.capA.toFixed(1)+' / '+r.capB.toFixed(1)+' kW'),!r.totalCords?'VENDOR REVIEW':(r.singleFeedFullLoad?'PASS':'REVIEW')]);
   el.innerHTML='<table><thead><tr><th>Rack</th><th>Typical IT</th><th>Design-Max</th><th>Peak Provisioning</th><th>Cords</th><th>Outlets A/B</th><th>PDU A/B</th><th>Usable A/B</th><th>Single-feed peak</th></tr></thead><tbody>'+rows.map(r=>'<tr>'+r.map(c=>'<td>'+c+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
   notes.innerHTML='• <b>Three-tier power model:</b> Typical IT = operating/energy reference, Design-Max = rack placement/cooling basis, Peak-Provisioning = upstream electrical/PDU provisioning envelope.<br>• H200/B200 official system tables provide max values; a separate vendor-published typical value is not assumed. B300 uses 14.5 kW operating, 15.0 kW design max, and 19.7 kW/system peak-provisioning from the 39.4 kW / 2-system rack peak reference.';
 };
}

function mountPowerPanel(){
 let p=document.getElementById('v48-power-model');if(p)return p;
 const warning=q('resultWarning');if(!warning)return null;
 p=document.createElement('section');p.id='v48-power-model';warning.insertAdjacentElement('afterend',p);return p;
}
function renderPowerEnvelope(){
 const p=mountPowerPanel();if(!p)return;const s=getSelectedSystem(),e=powerEnvelope(s),units=Number(txt(q('mUnits')))||0;
 const typical=e.typical==null?null:e.typical*units,design=e.designMax*units,peak=e.peak*units;
 p.innerHTML='<h3>3-Level Power Model <span class="v48-badge">electrical basis</span></h3><div class="v48-power-grid">'+
 '<div class="v48-power-cell metric"><div class="k">Typical IT Power</div><div class="v">'+(typical==null?'—':typical.toFixed(1)+' kW')+'</div><div class="s">'+e.typicalBasis+'</div></div>'+
 '<div class="v48-power-cell metric"><div class="k">Design-Max Power</div><div class="v">'+design.toFixed(1)+' kW</div><div class="s">Rack placement and cooling design basis · '+e.designMax.toFixed(1)+' kW/system</div></div>'+
 '<div class="v48-power-cell metric"><div class="k">Peak-Provisioning Power</div><div class="v">'+peak.toFixed(1)+' kW</div><div class="s">'+e.peakBasis+'</div></div></div>'+
 '<div class="v48-sub" style="margin-top:8px">Sources: <a href="'+(systemId()==='b300'?POWER_SOURCES.b300:POWER_SOURCES[systemId()]||'#')+'" target="_blank" rel="noopener">system power</a>'+(systemId()==='b300'?' · <a href="'+POWER_SOURCES.b300Peak+'" target="_blank" rel="noopener">B300 rack peak envelope</a>':'')+'.</div>';
 if(q('mPower'))q('mPower').innerHTML=(typical==null?'Typical —':'Typical '+typical.toFixed(1))+'<br><span style="font-size:10px">Design '+design.toFixed(1)+' / Peak '+peak.toFixed(1)+' kW</span>';
 return{typical,design,peak,e};
}

function logicalPerEndpointCage(s){const cages=Number(s.physicalClusterCages||0),links=Number(s.links||0);return cages>0?Math.max(1,links/cages):1}
function physicalize(){
 const s=getSelectedSystem(),sw=switches[q('switchId').value],units=Number(txt(q('mUnits')))||0,leafCount=Number(txt(q('mLeafs')))||0,spineCount=Number(txt(q('mSpines')))||0;
 const computeSpeed=s.linkSpeed,upSpeed=Number(q('uplinkSpeed').value)||computeSpeed,spare=1+(+q('sparePct').value||0)/100;
 const logical=units*(s.links||0),designLogical=Math.ceil(logical*spare);
 const endpointLpc=logicalPerEndpointCage(s),switchLpc=Math.max(1,portsPerCage(sw,computeSpeed)||1);
 const endpointCages=Math.ceil(designLogical/endpointLpc);
 const perLeaf=leafCount?Math.ceil(designLogical/leafCount):designLogical;
 const leafDownCages=leafCount?leafCount*Math.ceil(perLeaf/switchLpc):0;
 let totalLeafUplinks=0;
 const engineText=txt(q('spineEngineText'));const m=engineText.match(/Leaf↔Spine=(\d+)/);if(m)totalLeafUplinks=Number(m[1]);
 if(!totalLeafUplinks&&leafCount){const cap=maxDownlinks(sw,computeSpeed,upSpeed,+q('oversub').value||1);if(cap)totalLeafUplinks=leafCount*cap.uplinks}
 const designLeafUplinks=Math.ceil(totalLeafUplinks*spare);
 const upLpc=Math.max(1,portsPerCage(sw,upSpeed)||1);
 const leafUpCages=Math.ceil(designLeafUplinks/upLpc);
 const spineDownCages=spineCount?spineCount*Math.ceil(Math.ceil(designLeafUplinks/spineCount)/upLpc):0;
 const sfProfile=mediaProfile(computeSpeed,+q('distServer').value),lsProfile=mediaProfile(upSpeed,+q('distSpine').value);
 const endpointOptics=sfProfile.optical?endpointCages:0,leafDownOptics=sfProfile.optical?leafDownCages:0;
 const leafUpOptics=lsProfile.optical?leafUpCages:0,spineOptics=lsProfile.optical?spineDownCages:0;
 // Cable legs remain link-level unless an exact breakout/harness SKU proves multiple logical links per assembly.
 // This prevents silently treating a twin-port OSFP cage as one cable assembly.
 const sfCableLegs=designLogical,lsCableLegs=designLeafUplinks;
 return{logical,designLogical,endpointLpc,switchLpc,endpointCages,leafDownCages,endpointOptics,leafDownOptics,totalSfOptics:endpointOptics+leafDownOptics,sfCables:sfCableLegs,totalLeafUplinks,designLeafUplinks,leafUpCages,spineDownCages,leafUpOptics,spineOptics,totalLsOptics:leafUpOptics+spineOptics,lsCables:lsCableLegs,computeSpeed,upSpeed};
}
function mountPhysicalPanel(){
 let p=document.getElementById('v48-physical-audit');if(p)return p;const sf=q('sfLinks');if(!sf)return null;const card=sf.closest('.card');if(!card)return null;p=document.createElement('section');p.id='v48-physical-audit';card.appendChild(p);return p;
}
function upsertBom(category,item,qty){
 const body=q('bom');if(!body)return;let row=Array.from(body.querySelectorAll('tr')).find(tr=>{const td=tr.querySelectorAll('td');return td[0]&&td[0].textContent===category});
 if(!row){row=document.createElement('tr');row.innerHTML='<td></td><td></td><td></td>';body.appendChild(row)}
 const td=row.querySelectorAll('td');td[0].textContent=category;td[1].textContent=item;td[2].textContent=qty;
}
function renderPhysicalization(){
 const p=mountPhysicalPanel();if(!p)return null;const x=physicalize(),s=getSelectedSystem(),sw=switches[q('switchId').value];
 p.innerHTML='<h3>Logical Link ↔ Physical Cage / Optic Audit <span class="v48-badge">no 1:1 assumption</span></h3><div class="v48-sub">Logical links, physical cages, optic modules and cable assemblies are counted separately. Twin-port OSFP is packed by <b>logicalPerCage</b>, not by line rate.</div><table><thead><tr><th>Segment</th><th>Logical links</th><th>Endpoint cages</th><th>Switch-side cages</th><th>Optic modules</th><th>Cable assemblies</th><th>Port mode</th></tr></thead><tbody>'+
 '<tr><td>System → Leaf</td><td>'+x.logical+' active / '+x.designLogical+' incl. spare</td><td>'+x.endpointCages+'</td><td>'+x.leafDownCages+'</td><td>'+x.endpointOptics+' endpoint + '+x.leafDownOptics+' switch</td><td>'+x.sfCables+' link-level cable legs</td><td>'+x.computeSpeed+'G · endpoint '+x.endpointLpc.toFixed(1)+' logical/cage · switch '+x.switchLpc+' logical/cage</td></tr>'+
 '<tr><td>Leaf → Spine</td><td>'+x.totalLeafUplinks+' active / '+x.designLeafUplinks+' incl. spare</td><td>'+x.leafUpCages+' leaf cages</td><td>'+x.spineDownCages+' spine cages</td><td>'+x.leafUpOptics+' leaf + '+x.spineOptics+' spine</td><td>'+x.lsCables+' link-level cable legs</td><td>'+x.upSpeed+'G · '+(portMode(sw,x.upSpeed)?.mode||'vendor mode')+'</td></tr></tbody></table><div class="v48-sub" style="margin-top:8px">Cable quantity is intentionally reported as <b>link-level cable legs</b> until an exact twin-port/breakout harness SKU is selected. Physical cage count is not silently reused as cable-assembly count.</div>';
 if(q('sfOC'))q('sfOC').textContent=(x.totalSfOptics?x.totalSfOptics+' optics (physicalized)':'integrated')+' / '+x.sfCables+' cable(s)';
 if(q('lsOC'))q('lsOC').textContent=(x.totalLsOptics?x.totalLsOptics+' optics (physicalized)':'integrated')+' / '+x.lsCables+' cable(s)';
 const sfRow=Array.from(q('bom')?.querySelectorAll('tr')||[]).find(tr=>txt(tr.querySelector('td'))==='Server-facing optics');if(sfRow){const td=sfRow.querySelectorAll('td');td[1].textContent='Endpoint + switch-side optics · physical cage aware';td[2].textContent=x.totalSfOptics}
 const lsRow=Array.from(q('bom')?.querySelectorAll('tr')||[]).find(tr=>txt(tr.querySelector('td'))==='Leaf↔Spine optics');if(lsRow){const td=lsRow.querySelectorAll('td');td[1].textContent='Leaf + spine optics · physical cage aware';td[2].textContent=x.totalLsOptics}
 upsertBom('Connectivity Audit','System→Leaf physical cages (endpoint / switch)',x.endpointCages+' / '+x.leafDownCages);
 upsertBom('Connectivity Audit','Leaf→Spine physical cages (leaf / spine)',x.leafUpCages+' / '+x.spineDownCages);
 return x;
}

function scoreV49(expected,actual,keys){
 const errs=keys.map(k=>expected[k]?Math.abs(actual[k]-expected[k])/expected[k]*100:0);
 const mape=errs.reduce((a,b)=>a+b,0)/Math.max(1,errs.length),max=Math.max(...errs,0);
 return{mape,max,status:max<10?'PASS':'FAIL'};
}
function deriveB300Fabric(nodes){
 const sw=switches.q3400,speed=800,cap=maxDownlinks(sw,speed,speed,1);
 const gpus=nodes*8,nodeLeaf=nodes*8;
 const leaf=cap?Math.ceil(nodeLeaf/cap.downlinks):0;
 const leafSpine=cap?leaf*cap.uplinks:0;
 const spine=Math.ceil(leafSpine/Math.max(1,maxLogical(sw,speed)));
 return{nodes,gpus,leaf,spine,nodeLeaf,leafSpine,cap};
}
function runB300ValidationSuite(){
 const f1=deriveB300Fabric(72),f2=deriveB300Fabric(144),f18=deriveB300Fabric(1296);
 const p=POWER_ENVELOPES.b300;
 const cases=[];
 let expected={nodes:72,gpus:576,leaf:8,spine:4,nodeLeaf:576,leafSpine:576};
 let actual={nodes:f1.nodes,gpus:f1.gpus,leaf:f1.leaf,spine:f1.spine,nodeLeaf:f1.nodeLeaf,leafSpine:f1.leafSpine};
 cases.push({id:31,name:'DGX B300 · 1 SU golden',scope:'Topology',expected,actual,keys:Object.keys(expected),...scoreV49(expected,actual,Object.keys(expected))});
 expected={nodes:144,gpus:1152,leaf:16,spine:8,nodeLeaf:1152,leafSpine:1152};
 actual={nodes:f2.nodes,gpus:f2.gpus,leaf:f2.leaf,spine:f2.spine,nodeLeaf:f2.nodeLeaf,leafSpine:f2.leafSpine};
 cases.push({id:32,name:'DGX B300 · 2 SU hold-out',scope:'Frozen-engine topology',expected,actual,keys:Object.keys(expected),...scoreV49(expected,actual,Object.keys(expected))});
 expected={typical:30,design:30,peak:39.4};
 actual={typical:p.typical*2,design:p.designMax*2,peak:p.peak*2};
 cases.push({id:33,name:'DGX B300 · 2-system rack power',scope:'3-level power hold-out',expected,actual,keys:['typical','design','peak'],...scoreV49(expected,actual,['typical','design','peak'])});
 expected={typical:58,design:60,peak:76};
 actual={typical:p.typical*4,design:p.designMax*4,peak:p.peak*4};
 cases.push({id:34,name:'DGX B300 · 4-system rack power',scope:'3-level power hold-out',expected,actual,keys:['typical','design','peak'],...scoreV49(expected,actual,['typical','design','peak'])});
 expected={nodes:1296,leaf:144,spine:72,nodeLeaf:10368,leafSpine:10368};
 actual={nodes:f18.nodes,leaf:f18.leaf,spine:f18.spine,nodeLeaf:f18.nodeLeaf,leafSpine:f18.leafSpine};
 cases.push({id:35,name:'DGX B300 · 18 SU hold-out',scope:'Frozen-engine topology',expected,actual,keys:Object.keys(expected),...scoreV49(expected,actual,Object.keys(expected)),sourceNote:'Published GPU total is internally inconsistent with 1,296 nodes × 8 GPUs, so GPU count is excluded from exact scoring.'});
 return{cases,cap:f1.cap};
}
function compactCaseValues(o,keys){
 return keys.map(k=>k+'='+((typeof o[k]==='number'&&Math.abs(o[k]-Math.round(o[k]))>.0001)?o[k].toFixed(1):o[k])).join(' · ');
}
if(typeof renderGoldenValidation==='function'){
 const _renderGoldenValidation=renderGoldenValidation;
 renderGoldenValidation=function(){
   _renderGoldenValidation();
   const el=q('goldenSummary'),note=q('goldenNotes');if(!el)return;
   const suite=runB300ValidationSuite();
   let box=document.getElementById('v49-b300-suite');
   if(!box){box=document.createElement('div');box.id='v49-b300-suite';el.insertAdjacentElement('afterend',box)}
   box.innerHTML='<div class="guideTableWrap"><table><thead><tr><th>Case</th><th>Validation scope</th><th>Engine output</th><th>Reference</th><th>MAPE</th><th>Max error</th><th>Status</th></tr></thead><tbody>'+
     suite.cases.map(r=>'<tr><td><b>'+r.id+'</b> · '+r.name+'</td><td>'+r.scope+'</td><td>'+compactCaseValues(r.actual,r.keys)+'</td><td>'+compactCaseValues(r.expected,r.keys)+'</td><td>'+r.mape.toFixed(4)+'%</td><td>'+r.max.toFixed(4)+'%</td><td class="'+(r.status==='PASS'?'good':'bad')+'"><b>'+r.status+'</b></td></tr>').join('')+
     '</tbody></table></div><div class="v48-sub" style="margin-top:8px"><b>Case 31</b> is the new DGX B300 golden/reference case. <b>Cases 32–35</b> are recalculated in the running UI with the same generalized fabric / power equations and no case-specific production answer table. Case 35 intentionally excludes the inconsistent published GPU total from exact scoring. <a href="'+POWER_SOURCES.b300Ra+'" target="_blank" rel="noopener">NVIDIA B300 SuperPOD RA</a>.</div>';
   if(note&&!note.dataset.v49HoldoutNote){note.innerHTML+='<br>• <b>v4.9 validation:</b> DGX B300 Case 31 + hold-out Cases 32–35 are now visible as live recalculation results. Q3400 packing for the 1-SU basis resolves to '+suite.cap.downlinks+' logical downlinks + '+suite.cap.uplinks+' uplinks per leaf.';note.dataset.v49HoldoutNote='1'}
 };
}

const RACK_VIEW_STATE={mode:'2d',side:'front'};
function rackModel(){
 const s=getSelectedSystem(),units=Number(txt(q('mUnits')))||1,racks=Math.max(1,Number(txt(q('mComputeRacks')))||1),perRack=Math.max(1,Math.ceil(units/racks));
 return{s,perRack,rackU:Number(q('rackRU').value)||48,sw:switches[q('switchId').value],power:powerEnvelope(s)};
}
function mountRackTwin(){
 let p=document.getElementById('v48-rack-twin');if(p)return p;
 const grid=q('rackGrid');if(!grid)return null;
 p=document.createElement('section');p.id='v48-rack-twin';grid.insertAdjacentElement('beforebegin',p);return p;
}
function renderRackTwin(){
 const p=mountRackTwin();if(!p)return;
 const m=rackModel(),s=m.s,n=m.perRack,U=m.rackU,ru=s.ru||10;
 const mode=RACK_VIEW_STATE.mode,side=RACK_VIEW_STATE.side;
 const modeLabel=mode==='2d'?'2D ENGINEERING ELEVATION':'3D ISOMETRIC / SERVICE VIEW';
 const sideLabel=side==='front'?'FRONT':'REAR';
 p.innerHTML='<h3>Rack 구성 해설도 <span class="v48-badge">'+s.name+'</span></h3>'+
 '<div class="v48-sub">아래의 <b>2D / 3D</b> 셀과 <b>Front / Rear</b> 셀을 눌러 한 화면씩 전환합니다. 2D는 실제 랙 정면/후면의 rail·bezel·vent·port·PDU·patching을 묘사한 engineering elevation, 3D는 cabinet depth·service side·배선 경로·raised-floor perspective를 포함한 isometric view입니다. 실제 장비와 유사한 시각 표현이지만 제조사 CAD/IFC 치수 도면은 아닙니다.</div>'+
 '<div class="v491-rack-toolbar"><div class="v491-rack-mode">'+
 '<button type="button" class="v491-rack-tab '+(mode==='2d'?'active':'')+'" data-rack-mode="2d">2D</button>'+
 '<button type="button" class="v491-rack-tab '+(mode==='3d'?'active':'')+'" data-rack-mode="3d">3D</button>'+
 '</div><div class="v491-rack-side">'+
 '<button type="button" class="v491-rack-tab '+(side==='front'?'active':'')+'" data-rack-side="front">Front</button>'+
 '<button type="button" class="v491-rack-tab '+(side==='rear'?'active':'')+'" data-rack-side="rear">Rear</button>'+
 '</div></div>'+
 '<div class="v491-rack-stage"><div class="cap"><span>'+modeLabel+'</span><span>'+sideLabel+'</span></div><svg id="v491-rack-canvas" viewBox="0 0 920 820"></svg></div>'+
 '<div class="v491-rack-help">Front: bezel / vent / drive·I/O / switch optic-port / patching 중심 · Rear: fan / PSU / NIC·OSFP / management / power inlet / service cabling 중심</div>';

 p.querySelectorAll('[data-rack-mode]').forEach(btn=>btn.addEventListener('click',()=>{RACK_VIEW_STATE.mode=btn.dataset.rackMode;renderRackTwin()}));
 p.querySelectorAll('[data-rack-side]').forEach(btn=>btn.addEventListener('click',()=>{RACK_VIEW_STATE.side=btn.dataset.rackSide;renderRackTwin()}));

 const svg=document.getElementById('v491-rack-canvas');if(!svg)return;
 const NS='http://www.w3.org/2000/svg';
 const E=(t,o,parent=svg)=>{const e=document.createElementNS(NS,t);Object.entries(o||{}).forEach(([k,v])=>e.setAttribute(k,v));parent.appendChild(e);return e};
 const T=(x,y,str,size=10,fill='#26384d',anchor='start',weight='600',parent=svg)=>{const e=E('text',{x,y,fill,'font-size':size,'text-anchor':anchor,'font-weight':weight},parent);e.textContent=str;return e};
 const D=E('defs',{});
 D.innerHTML='<pattern id="v492floor" width="24" height="24" patternUnits="userSpaceOnUse"><path d="M 24 0 L 0 0 0 24" fill="none" stroke="#b9c3cd" stroke-width="0.6"/></pattern><pattern id="v492mesh" width="8" height="8" patternUnits="userSpaceOnUse"><circle cx="2" cy="2" r="0.8" fill="#8293a3"/><circle cx="6" cy="6" r="0.8" fill="#8293a3"/></pattern><linearGradient id="v492glass" x1="0" x2="1"><stop offset="0" stop-color="#dbe5ee" stop-opacity=".15"/><stop offset=".45" stop-color="#ffffff" stop-opacity=".42"/><stop offset="1" stop-color="#9fb1c0" stop-opacity=".10"/></linearGradient><linearGradient id="v491cab" x1="0" x2="1"><stop offset="0" stop-color="#0b1117"/><stop offset=".5" stop-color="#3b4652"/><stop offset="1" stop-color="#10171e"/></linearGradient><linearGradient id="v491dev" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#526579"/><stop offset=".5" stop-color="#2b3b4a"/><stop offset="1" stop-color="#15202a"/></linearGradient><linearGradient id="v491sw" x1="0" x2="1"><stop offset="0" stop-color="#123a58"/><stop offset=".5" stop-color="#3b7daa"/><stop offset="1" stop-color="#112f49"/></linearGradient><linearGradient id="v491rear" x1="0" x2="1"><stop offset="0" stop-color="#26313b"/><stop offset=".5" stop-color="#475561"/><stop offset="1" stop-color="#202a33"/></linearGradient><filter id="v491shadow" x="-30%" y="-30%" width="160%" height="160%"><feDropShadow dx="4" dy="6" stdDeviation="5" flood-opacity=".28"/></filter>';
 const screw=(x,y)=>E('circle',{cx:x,cy:y,r:2,fill:'#b7c0c9',stroke:'#29323a','stroke-width':.6});
 const callout=(x,y,tx,ty,label,align='start')=>{E('path',{d:'M '+x+' '+y+' L '+(tx+(align==='end'?-7:7))+' '+(ty-3),fill:'none',stroke:'#75879a','stroke-width':1.2});E('circle',{cx:x,cy:y,r:2.4,fill:'#2c6fa5'});T(tx,ty,label,9,'#34495e',align,'700')};
 const reserved=8,maxSystems=Math.min(n,Math.max(1,Math.floor((U-reserved)/Math.max(1,ru))));

 function draw2D(which){
   E('rect',{x:38,y:20,width:844,height:760,rx:18,fill:'#eef2f5',stroke:'#c8d0d8'});E('rect',{x:38,y:705,width:844,height:75,fill:'url(#v492floor)',opacity:.72});
   const ox=210,oy=38,cw=420,ch=690,innerX=246,innerY=68,innerW=348,innerH=624;
   E('rect',{x:ox,y:oy,width:cw,height:ch,rx:15,fill:'url(#v491cab)',stroke:'#070b0f','stroke-width':6,filter:'url(#v491shadow)'});
   E('rect',{x:innerX,y:innerY,width:innerW,height:innerH,rx:3,fill:'#d6dce2',stroke:'#687684','stroke-width':2});E('rect',{x:innerX+4,y:innerY+4,width:innerW-8,height:innerH-8,rx:2,fill:'url(#v492glass)',stroke:'#ffffff','stroke-opacity':.28});
   E('rect',{x:252,y:74,width:10,height:608,fill:'#263039'});E('rect',{x:578,y:74,width:10,height:608,fill:'#263039'});
   for(let y=82;y<678;y+=11){E('rect',{x:255,y,width:4,height:4,fill:'#8995a0'});E('rect',{x:581,y,width:4,height:4,fill:'#8995a0'})}
   E('rect',{x:220,y:80,width:18,height:590,rx:5,fill:'#4e1717',stroke:'#d15b5b'});E('rect',{x:602,y:80,width:18,height:590,rx:5,fill:'#123252',stroke:'#4e99d3'});
   for(let y=96;y<653;y+=29){E('rect',{x:224,y,width:10,height:15,rx:2,fill:'#eccccc'});E('rect',{x:606,y,width:10,height:15,rx:2,fill:'#c9e0f2'})}
   T(229,63,'PDU-A',9,'#ef7777','middle','800');T(611,63,'PDU-B',9,'#73b9ee','middle','800');
   const usable=598,uy=usable/U,startY=680;
   for(let u=1;u<=U;u++){const y=startY-u*uy;E('line',{x1:264,y1:y,x2:576,y2:y,stroke:u%2?'#b3bbc4':'#949fa9','stroke-width':.34});if(u%2===0)T(242,y+3,'U'+u,7,'#526475','end','600')}
   let cursor=1;
   const blocks=[];
   function frontServer(label,devRu){
     const y=startY-(cursor+devRu-1)*uy,h=Math.max(22,devRu*uy-2),x=266,w=308;
     E('rect',{x,y,width:w,height:h,rx:4,fill:'url(#v491dev)',stroke:'#091017','stroke-width':1.4});
     screw(x+7,y+7);screw(x+w-7,y+7);screw(x+7,y+h-7);screw(x+w-7,y+h-7);
     E('rect',{x:x+12,y:y+7,width:16,height:Math.max(8,h-14),rx:4,fill:'#0c141b'});E('rect',{x:x+w-28,y:y+7,width:16,height:Math.max(8,h-14),rx:4,fill:'#0c141b'});
     for(let yy=y+10;yy<y+h-8;yy+=8)for(let xx=x+39;xx<x+137;xx+=9)E('circle',{cx:xx,cy:yy,r:1.4,fill:'#8194a6'});
     for(let k=0;k<5;k++)E('rect',{x:x+151+k*25,y:y+10,width:19,height:Math.min(24,h-19),rx:2,fill:'#111a23',stroke:'#758494','stroke-width':.6});
     for(let k=0;k<8;k++){E('rect',{x:x+150+k*17,y:y+h-17,width:13,height:8,rx:1,fill:k<4?'#24445b':'#182b3b',stroke:'#7b8a97','stroke-width':.4});E('circle',{cx:x+154+k*17,cy:y+h-13,r:1.2,fill:k===0?'#50df8c':'#88a1b6'})}
     T(x+w/2,y+17,label,9,'#f6f9fb','middle','800');blocks.push({kind:'server',x,y,w,h});cursor+=devRu;
   }
   function rearServer(label,devRu){
     const y=startY-(cursor+devRu-1)*uy,h=Math.max(22,devRu*uy-2),x=266,w=308;
     E('rect',{x,y,width:w,height:h,rx:4,fill:'url(#v491rear)',stroke:'#091017','stroke-width':1.4});
     screw(x+7,y+7);screw(x+w-7,y+7);screw(x+7,y+h-7);screw(x+w-7,y+h-7);
     // fan wall
     for(let k=0;k<5;k++){const cx=x+36+k*33,cy=y+Math.max(18,h/2);E('circle',{cx,cy,r:12,fill:'#101820',stroke:'#72808d','stroke-width':1});E('circle',{cx,cy,r:4,fill:'#3b4b58'});for(let q=0;q<6;q++){const ang=q*Math.PI/3;E('line',{x1:cx+5*Math.cos(ang),y1:cy+5*Math.sin(ang),x2:cx+10*Math.cos(ang),y2:cy+10*Math.sin(ang),stroke:'#70808e','stroke-width':1})}}
     // NIC/OSFP cages
     for(let k=0;k<8;k++)E('rect',{x:x+184+(k%4)*24,y:y+9+Math.floor(k/4)*16,width:18,height:10,rx:1,fill:'#0c1b27',stroke:'#9bc3dc','stroke-width':.7});
     // PSU modules / power inlets
     for(let k=0;k<4;k++){E('rect',{x:x+183+k*27,y:y+h-22,width:22,height:13,rx:2,fill:'#171f27',stroke:'#a1aab2','stroke-width':.6});E('circle',{cx:x+188+k*27,cy:y+h-15,r:1.5,fill:k%2?'#64a7d6':'#df6969'})}
     E('rect',{x:x+w-24,y:y+8,width:12,height:9,rx:1,fill:'#182a36',stroke:'#77a8c6'});T(x+w/2,y+17,label,9,'#f6f9fb','middle','800');blocks.push({kind:'server',x,y,w,h});cursor+=devRu;
   }
   function frontAux(label,devRu,kind){
     const y=startY-(cursor+devRu-1)*uy,h=Math.max(19,devRu*uy-2),x=266,w=308;
     E('rect',{x,y,width:w,height:h,rx:3,fill:kind==='switch'?'url(#v491sw)':'#344451',stroke:'#0a1118','stroke-width':1.1});
     if(kind==='switch'){for(let rr=0;rr<(h>31?2:1);rr++)for(let k=0;k<18;k++)E('rect',{x:x+38+k*13,y:y+8+rr*14,width:9,height:7,rx:1,fill:k%4===0?'#a6d8ef':'#c5d1da',stroke:'#0b2535','stroke-width':.45});E('circle',{cx:x+w-18,cy:y+13,r:2.2,fill:'#4cdd88'})}
     else{for(let k=0;k<16;k++)E('rect',{x:x+34+k*15,y:y+7,width:9,height:8,rx:1,fill:k%2?'#c0d8e7':'#8db7d0'})}
     T(x+w/2,y+15,label,8.5,'#fff','middle','800');blocks.push({kind,x,y,w,h});cursor+=devRu;
   }
   function rearAux(label,devRu,kind){
     const y=startY-(cursor+devRu-1)*uy,h=Math.max(19,devRu*uy-2),x=266,w=308;
     E('rect',{x,y,width:w,height:h,rx:3,fill:'#293742',stroke:'#0a1118','stroke-width':1.1});
     if(kind==='switch'){for(let k=0;k<5;k++){E('rect',{x:x+34+k*42,y:y+7,width:32,height:Math.max(8,h-14),rx:2,fill:'#15202a',stroke:'#7d8994','stroke-width':.5});E('circle',{cx:x+44+k*42,cy:y+h/2,r:4,fill:'#3b4a55'})}for(let k=0;k<2;k++)E('rect',{x:x+246+k*25,y:y+7,width:19,height:Math.max(8,h-14),rx:2,fill:'#121a21',stroke:'#9aa5ae','stroke-width':.6})}
     else{for(let k=0;k<10;k++)E('rect',{x:x+42+k*22,y:y+7,width:15,height:9,rx:1,fill:'#8199aa'})}
     T(x+w/2,y+15,label,8.5,'#fff','middle','800');blocks.push({kind,x,y,w,h});cursor+=devRu;
   }

   for(let i=0;i<maxSystems;i++){if(which==='front')frontServer((s.gpu||s.name)+' SYSTEM '+(i+1),ru);else rearServer((s.gpu||s.name)+' SYSTEM '+(i+1),ru)}
   if(cursor+1<U){if(which==='front')frontAux('HORIZONTAL CABLE MANAGER',1,'manager');else rearAux('REAR CABLE MANAGER',1,'manager')}
   if(cursor+1<U){if(which==='front')frontAux('FIBER PATCH · MPO / VSFF',1,'patch');else rearAux('PATCH REAR · TRUNK ENTRY',1,'patch')}
   if(cursor+Math.min(2,m.sw.ru||2)<U){if(which==='front')frontAux((m.sw.model||m.sw.name)+' · LEAF',Math.min(2,m.sw.ru||2),'switch');else rearAux((m.sw.model||m.sw.name)+' · PSU / FAN SIDE',Math.min(2,m.sw.ru||2),'switch')}

   if(which==='front'){
     for(let k=0;k<5;k++)E('path',{d:'M 566 '+(185+k*16)+' C 640 '+(185+k*16)+', 647 '+(305+k*20)+', 590 '+(340+k*21),fill:'none',stroke:k<3?'#16b7c9':'#8a65c7','stroke-width':2.5,'stroke-linecap':'round',opacity:.92});
     if(blocks[0]){callout(blocks[0].x+50,blocks[0].y+blocks[0].h/2,90,180,'GPU SYSTEM · bezel / vent','start');callout(blocks[0].x+225,blocks[0].y+blocks[0].h-13,90,213,'Front I/O / optic cage','start')}
     const swb=blocks.find(x=>x.kind==='switch');if(swb)callout(swb.x+200,swb.y+swb.h/2,690,560,'Leaf switch · port field','start');
     callout(229,300,90,310,'A-feed 0U PDU','start');callout(611,300,690,310,'B-feed 0U PDU','start');
     T(420,756,'2D Front · equipment faceplates / rack U / optical patching',11,'#273b4e','middle','800');
   }else{
     for(let k=0;k<6;k++)E('path',{d:'M 565 '+(180+k*18)+' C 662 '+(185+k*18)+', 674 '+(330+k*16)+', 612 '+(390+k*22),fill:'none',stroke:k<4?'#12a9b9':'#7b59b2','stroke-width':3,'stroke-linecap':'round',opacity:.94});
     for(let k=0;k<4;k++)E('path',{d:'M 278 '+(285+k*35)+' C 170 '+(300+k*34)+', 164 '+(480+k*25)+', 225 '+(520+k*25),fill:'none',stroke:k%2?'#4e91c6':'#c35454','stroke-width':3,'stroke-linecap':'round',opacity:.9});
     if(blocks[0]){callout(blocks[0].x+85,blocks[0].y+blocks[0].h/2,80,180,'Fan modules','start');callout(blocks[0].x+225,blocks[0].y+18,690,190,'NIC / OSFP cages','start');callout(blocks[0].x+230,blocks[0].y+blocks[0].h-15,690,225,'PSU / power inlet','start')}
     callout(611,360,690,380,'Rear 0U PDU outlets','start');
     T(420,756,'2D Rear · fan / PSU / NIC / management / power & fiber routing',11,'#273b4e','middle','800');
   }
   T(420,781,n+' system/rack · '+ru+'U/system · '+U+'U cabinet · selected system: '+s.name,8.5,'#647487','middle','600');
 }

 function draw3D(which){
   const rear=which==='rear';
   E('rect',{x:35,y:24,width:850,height:744,rx:20,fill:'#edf1f4',stroke:'#c6cfd7'});E('polygon',{points:'90,690 760,690 875,625 205,625',fill:'url(#v492floor)',stroke:'#aeb9c3','stroke-width':1});
   const fx=rear?305:185,fy=125,fw=340,fh=548,dx=rear?-115:115,dy=-62;
   const sideX=rear?fx:fx+fw;
   const topPts=fx+','+fy+' '+(fx+dx)+','+(fy+dy)+' '+(fx+fw+dx)+','+(fy+dy)+' '+(fx+fw)+','+fy;
   const sidePts=rear
     ? fx+','+fy+' '+(fx+dx)+','+(fy+dy)+' '+(fx+dx)+','+(fy+fh+dy)+' '+fx+','+(fy+fh)
     : (fx+fw)+','+fy+' '+(fx+fw+dx)+','+(fy+dy)+' '+(fx+fw+dx)+','+(fy+fh+dy)+' '+(fx+fw)+','+(fy+fh);
   E('polygon',{points:topPts,fill:'#697682',stroke:'#101820','stroke-width':2,filter:'url(#v491shadow)'});
   E('polygon',{points:sidePts,fill:'#27323c',stroke:'#101820','stroke-width':2});
   E('rect',{x:fx,y:fy,width:fw,height:fh,rx:5,fill:'#d0d7de',stroke:'#111820','stroke-width':5});E('rect',{x:fx+7,y:fy+7,width:fw-14,height:fh-14,rx:4,fill:'url(#v492glass)',stroke:'#f7fbff','stroke-opacity':.34});
   E('line',{x1:fx+13,y1:fy+18,x2:fx+13,y2:fy+fh-18,stroke:'#d65353','stroke-width':9});E('line',{x1:fx+fw-13,y1:fy+18,x2:fx+fw-13,y2:fy+fh-18,stroke:'#3e90cf','stroke-width':9});
   const u3=(fh-36)/U;let cursor=1;
   function device(label,devRu,kind){
     const y=fy+fh-18-(cursor+devRu-1)*u3,h=Math.max(19,devRu*u3-2),x=fx+29,w=fw-58,depth=rear?-55:55,dd=-29;
     const front=kind==='switch'?'#245f85':(rear?'#33424d':'#2b3d4d'),top=kind==='switch'?'#5794ba':'#586c7e',side='#182630';
     E('polygon',{points:x+','+y+' '+(x+depth)+','+(y+dd)+' '+(x+w+depth)+','+(y+dd)+' '+(x+w)+','+y,fill:top,stroke:'#111820'});
     const faceSide=rear
       ? x+','+y+' '+(x+depth)+','+(y+dd)+' '+(x+depth)+','+(y+h+dd)+' '+x+','+(y+h)
       : (x+w)+','+y+' '+(x+w+depth)+','+(y+dd)+' '+(x+w+depth)+','+(y+h+dd)+' '+(x+w)+','+(y+h);
     E('polygon',{points:faceSide,fill:side,stroke:'#111820'});
     E('rect',{x,y,width:w,height:h,rx:3,fill:front,stroke:'#101820','stroke-width':1.2});
     if(!rear&&kind==='server'){
       for(let yy=y+9;yy<y+h-7;yy+=9)for(let xx=x+35;xx<x+135;xx+=10)E('circle',{cx:xx,cy:yy,r:1.4,fill:'#8498a9'});
       for(let k=0;k<7;k++)E('rect',{x:x+160+k*16,y:y+h/2-5,width:11,height:9,rx:1,fill:'#10212d',stroke:'#86a9bd','stroke-width':.45});
     }else if(rear&&kind==='server'){
       for(let k=0;k<5;k++){const cx=x+31+k*31,cy=y+h/2;E('circle',{cx,cy,r:10,fill:'#111a22',stroke:'#7b8995'});E('circle',{cx,cy,r:3.5,fill:'#43525e'})}
       for(let k=0;k<8;k++)E('rect',{x:x+180+(k%4)*21,y:y+8+Math.floor(k/4)*15,width:16,height:9,rx:1,fill:'#0d1c27',stroke:'#9bc4dd','stroke-width':.6});
       for(let k=0;k<3;k++)E('rect',{x:x+182+k*30,y:y+h-20,width:24,height:12,rx:2,fill:'#141c23',stroke:'#9ba5ad','stroke-width':.5});
     }else if(!rear&&kind==='switch'){
       for(let rr=0;rr<(h>31?2:1);rr++)for(let k=0;k<17;k++)E('rect',{x:x+34+k*13,y:y+8+rr*13,width:9,height:6,rx:1,fill:k%4===0?'#98cee8':'#c0ccd5'});
     }else if(rear&&kind==='switch'){
       for(let k=0;k<5;k++)E('rect',{x:x+35+k*42,y:y+7,width:32,height:Math.max(8,h-14),rx:2,fill:'#16222b',stroke:'#7b8995','stroke-width':.5});
       for(let k=0;k<2;k++)E('rect',{x:x+247+k*24,y:y+7,width:18,height:Math.max(8,h-14),rx:2,fill:'#131b22',stroke:'#a1a9b0','stroke-width':.5});
     }else{
       for(let k=0;k<11;k++)E('rect',{x:x+45+k*19,y:y+7,width:13,height:8,rx:1,fill:'#89a8ba'});
     }
     T(x+w/2,y+15,label,8.5,'#fff','middle','800');cursor+=devRu;
   }
   for(let i=0;i<maxSystems;i++)device((s.gpu||s.name)+' '+(i+1),ru,'server');
   if(cursor+1<U)device(rear?'REAR CABLE MGR':'CABLE MGR',1,'manager');
   if(cursor+1<U)device(rear?'PATCH REAR / TRUNK ENTRY':'PATCH PANEL',1,'patch');
   if(cursor+2<U)device(rear?'LEAF · PSU / FAN SIDE':'LEAF SWITCH',Math.min(2,m.sw.ru||2),'switch');

   if(!rear){
     for(let k=0;k<7;k++)E('path',{d:'M '+(fx+fw-38)+' '+(225+k*17)+' C '+(fx+fw+72)+' '+(220+k*17)+', '+(fx+fw+90)+' '+(118+k*22)+', '+(fx+fw+dx-8)+' '+(125+k*22),fill:'none',stroke:k<4?'#19adbd':'#7d5bb7','stroke-width':3.2,'stroke-linecap':'round',opacity:.93});
     callout(fx+fw-30,250,720,205,'Front optic / patch cables','start');
     callout(fx+20,340,90,340,'A/B rack power rails','start');
     T(460,744,'3D Front · equipment depth + front optical service routing',11,'#293d50','middle','800');
   }else{
     for(let k=0;k<7;k++)E('path',{d:'M '+(fx+32)+' '+(215+k*17)+' C '+(fx-80)+' '+(225+k*17)+', '+(fx-104)+' '+(345+k*18)+', '+(fx+dx+10)+' '+(390+k*20),fill:'none',stroke:k<4?'#18a9b9':'#7654ae','stroke-width':3.2,'stroke-linecap':'round',opacity:.93});
     for(let k=0;k<4;k++)E('path',{d:'M '+(fx+38)+' '+(330+k*33)+' C '+(fx-50)+' '+(350+k*31)+', '+(fx-80)+' '+(520+k*20)+', '+(fx+dx+8)+' '+(560+k*14),fill:'none',stroke:k%2?'#488fc5':'#c35353','stroke-width':3,'stroke-linecap':'round',opacity:.92});
     callout(fx+65,250,700,210,'Rear NIC / OSFP / PSU zone','start');
     callout(fx+20,460,80,520,'Fiber + power service loops','start');
     T(430,744,'3D Rear · fan / PSU / NIC + rear fiber & power service routing',11,'#293d50','middle','800');
   }
   T(455,770,'Conceptual engineering geometry — realistic rack/service representation, not manufacturer CAD / IFC',8.5,'#667687','middle','600');
 }

 if(mode==='2d')draw2D(side);else draw3D(side);
}

function syncSnapshot(power,physical){
 if(!window.currentDesignSnapshot)return;const s=getSelectedSystem();
 window.currentDesignSnapshot.inputs.networkRoles=s.networkRoles;
 window.currentDesignSnapshot.results.powerEnvelopeKW={typical:power?.typical??null,designMax:power?.design??null,peakProvisioning:power?.peak??null};
 window.currentDesignSnapshot.results.physicalConnectivity=physical;
 if(typeof bomRowsFromDom==='function')window.currentDesignSnapshot.bom=bomRowsFromDom();
}
function afterCalcV48(){
 updateRoleNote();const power=renderPowerEnvelope(),physical=renderPhysicalization();renderRackTwin();
 if(power){upsertBom('Power Model','Typical IT Power',power.typical==null?'-':power.typical.toFixed(1)+' kW');upsertBom('Power Model','Design-Max Power',power.design.toFixed(1)+' kW');upsertBom('Power Model','Peak-Provisioning Power',power.peak.toFixed(1)+' kW')}
 syncSnapshot(power,physical);
}
if(typeof calc==='function'){
 const _calc=calc;
 calc=function(){const r=_calc.apply(this,arguments);try{afterCalcV48()}catch(e){console.error('v4.8 post-calc',e)}return r};
}
if(typeof renderProfileDb==='function'){
 const _renderProfileDb=renderProfileDb;
 renderProfileDb=function(){_renderProfileDb();try{const s=getSelectedSystem(),e=powerEnvelope(s),sum=q('selectedProfileSummary');if(sum)sum.innerHTML+='<br><b>Power envelope:</b> Typical '+(e.typical==null?'—':e.typical.toFixed(1)+' kW')+' / Design-Max '+e.designMax.toFixed(1)+' kW / Peak-Provisioning '+e.peak.toFixed(1)+' kW.'}catch(e){}};
}
try{refreshSystem()}catch(e){try{calc()}catch(_){}}
})();