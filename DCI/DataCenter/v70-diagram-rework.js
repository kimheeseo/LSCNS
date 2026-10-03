(function browserPatch(){
'use strict';
const VERSION='7.0.0';
const I18N={
 ko:{
   logicTitle:'Logic Diagram · 현재 설계 토폴로지',
   logicSub:'계산된 Leaf / Spine / Core / Compute Rack 수량을 사용한 2D 네트워크 토폴로지입니다.',
   cabTitle:'Structured Cabling · 현재 설계 배선 구조',
   cabSub:'Corning AEN185의 Level A / B / C 구조를 현재 계산값과 연결해 단순 2D 배선도로 표시합니다.',
   noteLogic:'실사 렌더링이 아닌 2D 논리도입니다. 화면에는 대표 노드만 표시하고, 각 계층의 ×N 값은 현재 BOM 계산 수량입니다.',
   noteCab:'Level A는 Node→Leaf, Level B는 Leaf→Spine structured cabling, Level C는 Spine→Core 구조를 나타냅니다. Level B는 jumper / adapter panel / housing / multifiber trunk의 백본 개념을 반영합니다.',
   source:'참고 구조: Corning Optical Communications AEN185 · NVIDIA AI Architecture Cabling Guide · Figures 19–23.',
   systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',racks:'Compute racks',
   levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',
   jumper:'Individual / bundled jumper',structured:'Patch panel + housing + multifiber trunk',p2p:'Point-to-point / bundled jumper / DAC',
   current:'Current design',notSelected:'Not selected'
 },
 en:{
   logicTitle:'Logic Diagram · Current Design Topology',
   logicSub:'2D network topology using the calculated Leaf / Spine / Core / Compute Rack quantities.',
   cabTitle:'Structured Cabling · Current Design Cabling',
   cabSub:'Simplified 2D cabling view mapping the current calculation to Corning AEN185 Levels A / B / C.',
   noteLogic:'This is a 2D logic diagram, not a photorealistic render. Only representative nodes are drawn; each ×N label is the current BOM quantity.',
   noteCab:'Level A represents Node→Leaf, Level B represents Leaf→Spine structured cabling, and Level C represents Spine→Core. Level B reflects jumper / adapter panel / housing / multifiber-trunk backbone concepts.',
   source:'Reference structure: Corning Optical Communications AEN185 · NVIDIA AI Architecture Cabling Guide · Figures 19–23.',
   systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',racks:'Compute racks',
   levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',
   jumper:'Individual / bundled jumper',structured:'Patch panel + housing + multifiber trunk',p2p:'Point-to-point / bundled jumper / DAC',
   current:'Current design',notSelected:'Not selected'
 },
 ja:{
   logicTitle:'Logic Diagram · 現在の設計トポロジ',logicSub:'計算された Leaf / Spine / Core / Compute Rack 数量を使用した2Dネットワークトポロジです。',
   cabTitle:'Structured Cabling · 現在の配線構成',cabSub:'Corning AEN185 Level A / B / C を現在の計算値に対応させた簡潔な2D配線図です。',
   noteLogic:'実写表現ではなく2D論理図です。代表ノードのみ表示し、×Nは現在のBOM計算数量です。',
   noteCab:'Level AはNode→Leaf、Level BはLeaf→Spine structured cabling、Level CはSpine→Coreを示します。',
   source:'参考構成: Corning Optical Communications AEN185 · NVIDIA AI Architecture Cabling Guide · Figures 19–23.',
   systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',racks:'Compute racks',
   levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',
   jumper:'Individual / bundled jumper',structured:'Patch panel + housing + multifiber trunk',p2p:'Point-to-point / bundled jumper / DAC',
   current:'Current design',notSelected:'Not selected'
 },
 zh:{
   logicTitle:'Logic Diagram · 当前设计拓扑',logicSub:'使用已计算的 Leaf / Spine / Core / Compute Rack 数量生成2D网络拓扑。',
   cabTitle:'Structured Cabling · 当前布线结构',cabSub:'将当前计算结果映射到 Corning AEN185 Level A / B / C 的简化2D布线图。',
   noteLogic:'该图为2D逻辑图而非写实渲染。仅显示代表节点，×N表示当前BOM计算数量。',
   noteCab:'Level A表示Node→Leaf，Level B表示Leaf→Spine structured cabling，Level C表示Spine→Core。',
   source:'参考结构: Corning Optical Communications AEN185 · NVIDIA AI Architecture Cabling Guide · Figures 19–23.',
   systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',racks:'Compute racks',
   levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',
   jumper:'Individual / bundled jumper',structured:'Patch panel + housing + multifiber trunk',p2p:'Point-to-point / bundled jumper / DAC',
   current:'Current design',notSelected:'Not selected'
 },
 de:{
   logicTitle:'Logic Diagram · Aktuelle Topologie',logicSub:'2D-Netztopologie mit den berechneten Leaf-/Spine-/Core-/Compute-Rack-Mengen.',
   cabTitle:'Structured Cabling · Aktuelle Verkabelung',cabSub:'Vereinfachte 2D-Verkabelung nach Corning AEN185 Level A / B / C mit den aktuellen Berechnungswerten.',
   noteLogic:'2D-Logikdiagramm statt fotorealistischer Darstellung. Es werden repräsentative Knoten gezeigt; ×N entspricht der aktuellen BOM-Menge.',
   noteCab:'Level A zeigt Node→Leaf, Level B Leaf→Spine Structured Cabling und Level C Spine→Core.',
   source:'Referenzstruktur: Corning Optical Communications AEN185 · NVIDIA AI Architecture Cabling Guide · Figures 19–23.',
   systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',racks:'Compute racks',
   levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',
   jumper:'Individual / bundled jumper',structured:'Patch panel + housing + multifiber trunk',p2p:'Point-to-point / bundled jumper / DAC',
   current:'Current design',notSelected:'Not selected'
 }
};
function q(id){return document.getElementById(id)}
function txt(e){return(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function lang(){const x=String(window.__dcBomUiLang||document.documentElement.getAttribute('data-dc-bom-ui-lang')||document.documentElement.lang||'ko').toLowerCase();if(x.includes('en'))return'en';if(x.includes('ja')||x.includes('jp'))return'ja';if(x.includes('zh')||x.includes('cn'))return'zh';if(x.includes('de'))return'de';return'ko'}
function T(){return I18N[lang()]||I18N.ko}
function esc(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function num(id,fallback){const e=q(id);if(!e)return fallback||0;const m=txt(e).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);return m?Number(m[0]):fallback||0}
function inputVal(ids,fallback){for(const id of ids){const e=q(id);if(e){if(e.tagName==='SELECT'){const o=e.options&&e.options[e.selectedIndex];return(o&&txt(o))||e.value||fallback}return e.value||txt(e)||fallback}}return fallback}
function design(){const r=window.DCDesign;if(r?.usable){const o=r.optical.leafSpine;return {systems:r.summary.systemUnits,gpus:r.summary.targetGPU,leaf:r.fabric.leafCount,spine:r.fabric.spineCount,core:r.fabric.coreCount,racks:r.summary.computeRacks||r.summary.totalRacks,speed:r.systemProfile?.linkSpeed||'RFQ',topology:r.input.topology,connector:o?.profile.connector||'RFQ',connectorDetail:'Exact pinning / polarity RFQ',product:o?.profile.media||'RFQ',trunk:o?.trunk.installedCableCount+' installed trunks',chain:'See calculated BOM; illustration only'};}
 return{
   systems:Math.max(0,num('mUnits',num('osaGpu',0))),
   gpus:Math.max(0,num('mGpu',num('osaGpu',0))),
   leaf:Math.max(0,num('mLeafs',num('osaLeaf',0))),
   spine:Math.max(0,num('mSpines',num('osaSpine',0))),
   core:Math.max(0,num('mCore',num('osaCore',0))),
   racks:Math.max(1,num('mComputeRacks',1)),
   speed:inputVal(['v48-compute-speed'],'Auto'),
   topology:inputVal(['topology'],'Clos / rail'),
   connector:txt(q('osaMpoLabel'))||inputVal(['osaMpoBase'],'Auto / exact interface'),
   connectorDetail:txt(q('osaMpoConnector'))||'-',
   product:txt(q('osaMpoProduct'))||'-',
   trunk:txt(q('osaTrunk'))||inputVal(['trunkFibers'],'-'),
   chain:inputVal(['osaChain'],'Structured')
 }
}
function positions(n,left,right){
 n=Math.max(1,n);if(n===1)return[(left+right)/2];
 const a=[];for(let i=0;i<n;i++)a.push(left+(right-left)*i/(n-1));return a
}
function switchNode(x,y,label,count,color){
 let s='<g transform="translate('+x+','+y+')"><rect x="-62" y="-22" width="124" height="44" rx="6" fill="'+color+'" stroke="#1b3e59" stroke-width="2"/>';
 for(let i=0;i<8;i++)s+='<rect x="'+(-48+i*13)+'" y="-5" width="8" height="7" rx="1" fill="#d8edf9"/>';
 s+='<circle cx="-49" cy="9" r="3" fill="#64d4ff"/><text x="0" y="-32" text-anchor="middle" font-size="12" font-weight="900" fill="#193d5d">'+esc(label)+'</text><rect x="25" y="6" width="48" height="22" rx="11" fill="#ffffff" stroke="#8faec3"/><text x="49" y="21" text-anchor="middle" font-size="11" font-weight="900" fill="#31546e">×'+esc(count)+'</text></g>';return s
}
function serverNode(x,y,label,count){
 return'<g transform="translate('+x+','+y+')"><rect x="-34" y="-42" width="68" height="84" rx="5" fill="#e8edf1" stroke="#687f91" stroke-width="2"/><rect x="-24" y="-30" width="48" height="10" fill="#b5c1ca"/><rect x="-24" y="-12" width="48" height="10" fill="#b5c1ca"/><rect x="-24" y="6" width="48" height="10" fill="#b5c1ca"/><circle cx="22" cy="31" r="4" fill="#49b8ee"/><text x="0" y="-54" text-anchor="middle" font-size="11" font-weight="800" fill="#264860">'+esc(label)+'</text><text x="0" y="61" text-anchor="middle" font-size="11" font-weight="900" fill="#31546e">×'+esc(count)+'</text></g>'
}
function summary(d){
 const L=T(),pairs=[[L.systems,d.systems],[L.gpus,d.gpus],[L.leaf,d.leaf],[L.spine,d.spine],[L.core,d.core],[L.racks,d.racks]];
 return'<div class="v70-summary">'+pairs.map(x=>'<div class="v70-kpi"><div class="k">'+esc(x[0])+'</div><div class="v">'+esc(x[1])+'</div></div>').join('')+'</div>'
}
function modal(){
 let m=q('v68-modal');
 if(!m){m=document.createElement('div');m.id='v68-modal';m.hidden=true;m.innerHTML='<div class="v68-card"><div class="v68-head"><div><h2></h2><p></p></div><button class="v68-close" type="button">×</button></div><div class="v68-body"></div></div>';document.body.appendChild(m);const close=()=>m.hidden=true;m.querySelector('.v68-close').onclick=close;m.addEventListener('click',e=>{if(e.target===m)close()})}
 return m
}
function logicSvg(d){
 const L=T(),coreN=d.core>0?Math.min(3,d.core):0,spineN=Math.max(1,Math.min(4,d.spine||1)),leafN=Math.max(1,Math.min(6,d.leaf||1)),rackN=Math.max(1,Math.min(6,d.racks||1));
 const coreX=positions(coreN||1,420,1080),spineX=positions(spineN,280,1220),leafX=positions(leafN,210,1290),rackX=positions(rackN,210,1290);
 let s='<svg viewBox="0 0 1500 850" role="img" aria-label="'+esc(L.logicTitle)+'"><rect width="1500" height="850" rx="18" fill="#f7f9fb"/>';
 s+='<text x="750" y="40" text-anchor="middle" fill="#173d69" font-size="24" font-weight="900">'+esc(L.logicTitle)+'</text><text x="750" y="67" text-anchor="middle" fill="#62778b" font-size="12">'+esc(d.topology)+' · '+esc(d.speed)+'</text>';
 const line=(x1,y1,x2,y2,w,op)=>'<line x1="'+x1+'" y1="'+y1+'" x2="'+x2+'" y2="'+y2+'" stroke="#365e7d" stroke-width="'+(w||2)+'" opacity="'+(op||.48)+'"/>';
 if(coreN){for(const sx of spineX)for(const cx of coreX)s+=line(sx,252,cx,142,2.2,.45)}
 for(const lx of leafX)for(const sx of spineX)s+=line(lx,472,sx,292,1.8,.34);
 for(let i=0;i<rackX.length;i++){const lx=leafX[Math.min(leafX.length-1,Math.floor(i*leafX.length/rackX.length))];s+=line(rackX[i],690,lx,512,2.2,.62)}
 if(coreN)coreX.forEach((x,i)=>{s+=switchNode(x,120,'Core '+(i+1),i===0?d.core:'…','#2c7fb8')});
 spineX.forEach((x,i)=>{s+=switchNode(x,270,'Spine '+(i+1),i===0?d.spine:'…','#2280aa')});
 leafX.forEach((x,i)=>{s+=switchNode(x,490,'Leaf '+(i+1),i===0?d.leaf:'…','#2c94b8')});
 rackX.forEach((x,i)=>{s+=serverNode(x,710,'Compute '+(i+1),i===0?d.racks:'…')});
 s+='<rect x="35" y="94" width="150" height="42" rx="8" fill="#dcecf8"/><text x="110" y="120" text-anchor="middle" font-size="13" font-weight="900" fill="#275a7c">CORE</text>';
 s+='<rect x="35" y="244" width="150" height="42" rx="8" fill="#dcecf8"/><text x="110" y="270" text-anchor="middle" font-size="13" font-weight="900" fill="#275a7c">SPINE</text>';
 s+='<rect x="35" y="464" width="150" height="42" rx="8" fill="#e3f2f6"/><text x="110" y="490" text-anchor="middle" font-size="13" font-weight="900" fill="#2d687a">LEAF / ToR</text>';
 s+='<rect x="35" y="684" width="150" height="42" rx="8" fill="#eef1f4"/><text x="110" y="710" text-anchor="middle" font-size="13" font-weight="900" fill="#4a6172">COMPUTE</text>';
 if(!coreN)s+='<text x="750" y="128" text-anchor="middle" fill="#8193a3" font-size="12">'+esc(L.core)+': '+esc(L.notSelected)+'</text>';
 s+='<text x="750" y="815" text-anchor="middle" fill="#708498" font-size="11">'+esc(L.noteLogic)+'</text></svg>';
 return s
}
function deviceBox(x,y,w,h,title,sub,color){
 return'<g transform="translate('+x+','+y+')"><rect width="'+w+'" height="'+h+'" rx="9" fill="'+color+'" stroke="#496579" stroke-width="2"/><text x="'+(w/2)+'" y="'+(h/2-5)+'" text-anchor="middle" font-size="13" font-weight="900" fill="#173d58">'+esc(title)+'</text><text x="'+(w/2)+'" y="'+(h/2+17)+'" text-anchor="middle" font-size="10" fill="#4f687a">'+esc(sub)+'</text></g>'
}
function cablingSvg(d){
 const L=T(),coreText=d.core>0?('×'+d.core):L.notSelected;
 let s='<svg viewBox="0 0 1500 900" role="img" aria-label="'+esc(L.cabTitle)+'"><rect width="1500" height="900" rx="18" fill="#fafbfc"/>';
 s+='<text x="750" y="40" text-anchor="middle" fill="#173d69" font-size="24" font-weight="900">'+esc(L.cabTitle)+'</text><text x="750" y="67" text-anchor="middle" fill="#667b8d" font-size="12">'+esc(d.chain)+' · '+esc(d.connector)+' · '+esc(d.trunk)+'</text>';
 // level bands
 s+='<rect x="45" y="120" width="1410" height="205" rx="14" fill="#eef6fb" stroke="#bdd2e3"/><text x="70" y="148" font-size="15" font-weight="900" fill="#24577b">'+esc(L.levelA)+'</text><text x="70" y="170" font-size="11" fill="#587185">'+esc(L.jumper)+'</text>';
 s+=deviceBox(155,205,230,78,'Compute nodes / racks','×'+d.racks,'#ffffff')+deviceBox(585,205,230,78,'Leaf switches','×'+d.leaf,'#dceef7');
 for(let i=0;i<5;i++)s+='<path d="M385 '+(220+i*12)+' C460 '+(220+i*12)+' 500 '+(220+i*12)+' 585 '+(220+i*12)+'" fill="none" stroke="#4a9fd0" stroke-width="2.5"/>';
 // Level B
 s+='<rect x="45" y="350" width="1410" height="275" rx="14" fill="#fff9e8" stroke="#e2d3a5"/><text x="70" y="380" font-size="15" font-weight="900" fill="#725a14">'+esc(L.levelB)+'</text><text x="70" y="402" font-size="11" fill="#7d6a35">'+esc(L.structured)+'</text>';
 s+=deviceBox(110,455,180,80,'Leaf','×'+d.leaf,'#e7f2f8')+deviceBox(375,445,190,100,'Patch Panel A','adapter / housing','#f1f3f5');
 s+='<g transform="translate(680,438)"><ellipse cx="105" cy="58" rx="92" ry="42" fill="none" stroke="#d5a629" stroke-width="12"/><ellipse cx="105" cy="58" rx="68" ry="29" fill="none" stroke="#edc75e" stroke-width="8"/><text x="105" y="55" text-anchor="middle" font-size="13" font-weight="900" fill="#725916">'+esc(d.trunk)+'</text><text x="105" y="75" text-anchor="middle" font-size="10" fill="#7d6a35">multifiber trunk</text></g>';
 s+=deviceBox(985,445,190,100,'Patch Panel B','adapter / housing','#f1f3f5')+deviceBox(1270,455,150,80,'Spine','×'+d.spine,'#dfeaf3');
 const bline=(x1,y1,x2,y2)=>'<path d="M'+x1+' '+y1+' C'+((x1+x2)/2)+' '+y1+' '+((x1+x2)/2)+' '+y2+' '+x2+' '+y2+'" fill="none" stroke="#d49a24" stroke-width="3"/>';
 for(let i=0;i<5;i++){const yy=470+i*12;s+=bline(290,yy,375,yy)+bline(565,yy,680,480+i*9)+bline(890,480+i*9,985,yy)+bline(1175,yy,1270,yy)}
 // Level C
 s+='<rect x="45" y="655" width="1410" height="165" rx="14" fill="#f1f3fb" stroke="#c7cde4"/><text x="70" y="685" font-size="15" font-weight="900" fill="#4d5686">'+esc(L.levelC)+'</text><text x="70" y="707" font-size="11" fill="#686f93">'+esc(L.p2p)+'</text>';
 s+=deviceBox(320,725,220,70,'Spine switches','×'+d.spine,'#e6edf5')+deviceBox(945,725,220,70,'Core switches',coreText,'#e7e9f8');
 for(let i=0;i<4;i++)s+='<path d="M540 '+(742+i*10)+' L945 '+(742+i*10)+'" stroke="#6f78a7" stroke-width="2.5" stroke-dasharray="'+(i%2?'8 6':'0')+'"/>';
 // current product receipt
 s+='<rect x="1120" y="116" width="300" height="170" rx="12" fill="#ffffff" stroke="#b9c9d5"/><text x="1142" y="145" font-size="13" font-weight="900" fill="#264b68">'+esc(L.current)+'</text><text x="1142" y="172" font-size="10.5" fill="#536d80">Connector: '+esc(d.connector)+'</text><text x="1142" y="195" font-size="10.5" fill="#536d80">Detail: '+esc(d.connectorDetail)+'</text><text x="1142" y="218" font-size="10.5" fill="#536d80">Trunk: '+esc(d.trunk)+'</text><text x="1142" y="241" font-size="10.5" fill="#536d80">Product: '+esc(d.product)+'</text><text x="1142" y="264" font-size="10.5" fill="#536d80">Leaf / Spine / Core: '+d.leaf+' / '+d.spine+' / '+d.core+'</text>';
 s+='<text x="750" y="866" text-anchor="middle" fill="#708498" font-size="11">'+esc(L.source)+'</text></svg>';
 return s
}
function openLogic(){
 const L=T(),d=design(),m=modal();m.querySelector('h2').textContent=L.logicTitle;m.querySelector('p').textContent=L.logicSub;m.querySelector('.v68-body').innerHTML=summary(d)+'<div class="v70-svg-wrap">'+logicSvg(d)+'</div><div class="v70-note">'+esc(L.noteLogic)+'</div>';m.hidden=false
}
function openStructured(){
 const L=T(),d=design(),m=modal();m.querySelector('h2').textContent=L.cabTitle;m.querySelector('p').textContent=L.cabSub;m.querySelector('.v68-body').innerHTML=summary(d)+'<div class="v70-svg-wrap">'+cablingSvg(d)+'</div><div class="v70-note">'+esc(L.noteCab)+'</div><div class="v70-source">'+esc(L.source)+'</div>';m.hidden=false
}
function coolingSvg(){
 return'<svg class="v70-flow-svg" viewBox="0 0 1600 900" preserveAspectRatio="xMidYMid meet" aria-hidden="true">'+
 '<defs><marker id="v70coldArrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#35b9ff"/></marker><marker id="v70warmArrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#ff6b59"/></marker></defs>'+
 '<path class="cold-main" marker-end="url(#v70coldArrow)" d="M330 390 L430 390 L665 420 L735 420 L1350 420 L1418 390"/>'+
 '<path class="warm-main" marker-end="url(#v70warmArrow)" d="M1418 445 L1350 465 L665 465 L430 465 L330 465"/>'+
 '<path class="cold-branch" marker-end="url(#v70coldArrow)" d="M808 615 L808 515"/><path class="warm-branch" marker-end="url(#v70warmArrow)" d="M958 515 L958 615"/>'+
 '<path class="cold-branch" marker-end="url(#v70coldArrow)" d="M1018 615 L1018 515"/><path class="warm-branch" marker-end="url(#v70warmArrow)" d="M1168 515 L1168 615"/>'+
 '<path class="cold-branch" marker-end="url(#v70coldArrow)" d="M1228 615 L1228 515"/><path class="warm-branch" marker-end="url(#v70warmArrow)" d="M1378 515 L1378 615"/>'+
 '<text class="flow-tag cold-tag" x="760" y="402">COLD SUPPLY →</text><text class="flow-tag warm-tag" x="1030" y="488">← WARM RETURN</text></svg>'
}
function ensureCoolingOverlay(){
 const host=document.querySelector('.v631-cool3d');if(!host)return false;
 const old=host.querySelector('.v68-flow-overlay');if(old)old.style.display='none';
 if(!host.querySelector('.v70-flow-svg'))host.insertAdjacentHTML('beforeend',coolingSvg());
 return true
}
function showCooling(){
 const p=q('v53-cooling-architecture');if(!p)return;
 p.classList.add('v70-open');const b=q('v68-cooling-btn');if(b)b.setAttribute('aria-expanded','true');ensureCoolingOverlay();setTimeout(()=>p.scrollIntoView({behavior:'smooth',block:'start'}),20)
}
function bind(){
 const logic=q('v68-logic-btn'),cab=q('v68-structured-btn'),cool=q('v68-cooling-btn');
 if(logic&&!logic.dataset.v70){logic.dataset.v70='1';logic.onclick=openLogic}
 if(cab&&!cab.dataset.v70){cab.dataset.v70='1';cab.onclick=openStructured}
 if(cool&&!cool.dataset.v70){cool.dataset.v70='1';cool.setAttribute('aria-expanded','false');cool.onclick=showCooling}
 ensureCoolingOverlay()
}
function start(){
 const p=q('v53-cooling-architecture');if(p)p.classList.remove('v70-open');
 let n=0;const t=setInterval(()=>{n++;bind();if(n>35)clearInterval(t)},300);
 document.addEventListener('click',e=>{if(e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],.lang,.language'))setTimeout(bind,90)},true);
 setInterval(bind,1800)
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start()
})();