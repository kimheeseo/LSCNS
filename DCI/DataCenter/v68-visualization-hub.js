(function browserPatch(){
'use strict';
const VERSION='6.8.0';
const I18N={
 ko:{cool:'Cooling Architecture',logic:'Logic Diagram',structured:'Structured Cabling',logicTitle:'Logic Diagram · 현재 설계 토폴로지',logicSub:'현재 BOM 계산 결과를 Leaf / Spine / Core / Compute Rack의 논리 구조로 재구성합니다.',structuredTitle:'Structured Cabling · 현재 설계 배선 구조',structuredSub:'현재 설계의 Node→Leaf / Leaf→Spine / Spine→Core 경로를 patch panel, bundled jumper, trunk 관점으로 보여줍니다.',systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',speed:'Compute speed',noteLogic:'대표 Rail 관점의 논리도입니다. 실제 port-to-port mapping은 선택한 switch/optic의 physical cage, breakout 및 rail 정책으로 확정됩니다.',noteCab:'AEN185의 Level A/B/C 구조를 설계 시각화에 적용했습니다. 현재 BOM의 connector/trunk 선택을 사용하지만 실제 시공 전 polarity, tray/port mapping, 거리, fire rating 및 제품 part number를 확인해야 합니다.',levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',refA:'서버 rack에서 Leaf rack으로 개별/번들 jumper',refB:'Patch panel + high-fiber trunk 기반 structured cabling',refC:'Point-to-point, bundled jumper 또는 조건부 structured cabling',current:'현재 선택',trunk:'Trunk',connector:'Connector',product:'Product',open:'Open'},
 en:{cool:'Cooling Architecture',logic:'Logic Diagram',structured:'Structured Cabling',logicTitle:'Logic Diagram · Current Design Topology',logicSub:'Reconstructs the current BOM result as a Leaf / Spine / Core / Compute Rack logic view.',structuredTitle:'Structured Cabling · Current Design Cabling',structuredSub:'Shows Node→Leaf / Leaf→Spine / Spine→Core using patch panels, bundled jumpers and trunks.',systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',speed:'Compute speed',noteLogic:'Representative rail logic view. Exact port-to-port mapping is determined by physical cages, breakout rules and the selected rail policy.',noteCab:'Applies the AEN185 Level A/B/C structure to the current design. Connector and trunk labels come from the BOM, but polarity, tray/port mapping, distance, fire rating and exact part numbers still require implementation review.',levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',refA:'Individual/bundled jumpers from server racks to leaf racks',refB:'Structured cabling using patch panels and high-fiber-count trunks',refC:'Point-to-point, bundled jumper, or conditional structured cabling',current:'Current selection',trunk:'Trunk',connector:'Connector',product:'Product',open:'Open'},
 ja:{cool:'Cooling Architecture',logic:'Logic Diagram',structured:'Structured Cabling',logicTitle:'Logic Diagram · 現在の設計トポロジ',logicSub:'現在のBOM計算結果を Leaf / Spine / Core / Compute Rack の論理構造で表示します。',structuredTitle:'Structured Cabling · 現在の配線構成',structuredSub:'Node→Leaf / Leaf→Spine / Spine→Core を patch panel、bundled jumper、trunk の観点で表示します。',systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',speed:'Compute speed',noteLogic:'代表Railの論理図です。実際のport-to-port mappingはphysical cage、breakout、rail policyで確定します。',noteCab:'AEN185のLevel A/B/C構造を現在設計に適用しています。Connector/trunk表記はBOMから取得しますが、施工前にpolarity、tray/port mapping、距離、難燃仕様、part numberを確認してください。',levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',refA:'Server rackからLeaf rackへの個別/ bundled jumper',refB:'Patch panel + high-fiber trunkによるstructured cabling',refC:'Point-to-point、bundled jumper、または条件付きstructured cabling',current:'現在の選択',trunk:'Trunk',connector:'Connector',product:'Product',open:'Open'},
 zh:{cool:'Cooling Architecture',logic:'Logic Diagram',structured:'Structured Cabling',logicTitle:'Logic Diagram · 当前设计拓扑',logicSub:'将当前BOM计算结果重构为 Leaf / Spine / Core / Compute Rack 逻辑视图。',structuredTitle:'Structured Cabling · 当前布线结构',structuredSub:'以patch panel、bundled jumper和trunk展示Node→Leaf / Leaf→Spine / Spine→Core。',systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',speed:'Compute speed',noteLogic:'代表性Rail逻辑图。实际port-to-port mapping由physical cage、breakout和rail policy决定。',noteCab:'将AEN185的Level A/B/C结构应用于当前设计。Connector/trunk标签来自BOM，但施工前仍需确认polarity、tray/port mapping、距离、阻燃等级和准确part number。',levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',refA:'Server rack到Leaf rack的individual/bundled jumper',refB:'Patch panel + high-fiber trunk的structured cabling',refC:'Point-to-point、bundled jumper或条件式structured cabling',current:'当前选择',trunk:'Trunk',connector:'Connector',product:'Product',open:'Open'},
 de:{cool:'Cooling Architecture',logic:'Logic Diagram',structured:'Structured Cabling',logicTitle:'Logic Diagram · Aktuelle Topologie',logicSub:'Stellt das aktuelle BOM-Ergebnis als Leaf-/Spine-/Core-/Compute-Rack-Logik dar.',structuredTitle:'Structured Cabling · Aktuelle Verkabelung',structuredSub:'Zeigt Node→Leaf / Leaf→Spine / Spine→Core mit Patch Panels, Bundled Jumpern und Trunks.',systems:'Systems',gpus:'GPUs',leaf:'Leaf',spine:'Spine',core:'Core',speed:'Compute speed',noteLogic:'Repräsentative Rail-Logik. Das exakte Port-Mapping wird durch Physical Cages, Breakout und Rail-Policy bestimmt.',noteCab:'Überträgt die AEN185-Level-A/B/C-Struktur auf das aktuelle Design. Connector/Trunk kommen aus dem BOM; vor Umsetzung sind Polarity, Tray/Port-Mapping, Distanz, Brandschutzklasse und Part Number zu prüfen.',levelA:'Level A · Node → Leaf',levelB:'Level B · Leaf → Spine',levelC:'Level C · Spine → Core',refA:'Individual/Bundled Jumper von Server-Racks zu Leaf-Racks',refB:'Structured Cabling mit Patch Panels und High-Fiber-Count Trunks',refC:'Point-to-point, Bundled Jumper oder bedingtes Structured Cabling',current:'Aktuelle Auswahl',trunk:'Trunk',connector:'Connector',product:'Product',open:'Open'}
};
function q(id){return document.getElementById(id)}
function text(e){return(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function lang(){
 const raw=String(window.__dcBomUiLang||document.documentElement.getAttribute('data-dc-bom-ui-lang')||document.documentElement.lang||'ko').toLowerCase();
 if(raw.includes('en'))return'en';if(raw.includes('ja')||raw.includes('jp'))return'ja';if(raw.includes('zh')||raw.includes('cn'))return'zh';if(raw.includes('de'))return'de';return'ko';
}
function T(){return I18N[lang()]||I18N.ko}
function num(id,fallback=0){const e=q(id);if(!e)return fallback;const m=text(e).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);return m?Number(m[0]):fallback}
function inputVal(ids,fallback='-'){for(const id of ids){const e=q(id);if(e){if(e.tagName==='SELECT'){const o=e.options&&e.options[e.selectedIndex];return(o&&text(o))||e.value||fallback}return e.value||text(e)||fallback}}return fallback}
function current(){
 const systems=Math.max(0,num('mUnits',num('osaGpu',0)));
 const gpus=Math.max(0,num('mGpu',num('osaGpu',0)));
 const leaf=Math.max(0,num('mLeafs',num('osaLeaf',0)));
 const spine=Math.max(0,num('mSpines',num('osaSpine',0)));
 const core=Math.max(0,num('mCore',num('osaCore',0)));
 const speed=inputVal(['v48-compute-speed'],'Auto');
 const topology=inputVal(['topology'],'Clos / rail');
 const connector=text(q('osaMpoLabel'))||inputVal(['osaMpoBase'],'Auto / exact interface');
 const connectorDetail=text(q('osaMpoConnector'))||'-';
 const product=text(q('osaMpoProduct'))||'-';
 const trunk=text(q('osaTrunk'))||inputVal(['trunkFibers'],'-');
 const chain=inputVal(['osaChain'],'Structured');
 const racks=Math.max(1,num('mComputeRacks',1));
 return{systems,gpus,leaf,spine,core,speed,topology,connector,connectorDetail,product,trunk,chain,racks};
}
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function modal(){
 let m=q('v68-modal');if(m)return m;
 m=document.createElement('div');m.id='v68-modal';m.hidden=true;
 m.innerHTML='<div class="v68-card"><div class="v68-head"><div><h2></h2><p></p></div><button class="v68-close" type="button">×</button></div><div class="v68-body"></div></div>';
 document.body.appendChild(m);
 const close=()=>m.hidden=true;m.querySelector('.v68-close').onclick=close;m.addEventListener('click',e=>{if(e.target===m)close()});document.addEventListener('keydown',e=>{if(e.key==='Escape')close()});
 return m;
}
function summaryHtml(d){
 const L=T(),pairs=[[L.systems,d.systems],[L.gpus,d.gpus],[L.leaf,d.leaf],[L.spine,d.spine],[L.core,d.core],[L.speed,String(d.speed).includes('Auto')?'Auto':String(d.speed).replace(/\s+/g,' ')]];
 return'<div class="v68-summary">'+pairs.map(x=>'<div class="v68-kpi"><div class="k">'+esc(x[0])+'</div><div class="v">'+esc(x[1])+'</div></div>').join('')+'</div>';
}
function rackSvg(x,y,label,count,kind){
 const w=165,h=230,top='#7c8790',front=kind==='core'?'#244a75':kind==='spine'?'#2b3442':'#253642',accent=kind==='core'?'#6fa6ff':kind==='spine'?'#87a6c4':'#47b2ff';
 let s='<g transform="translate('+x+','+y+')"><polygon points="0,24 25,0 '+(w+25)+',0 '+w+',24" fill="'+top+'" stroke="#2e3942"/><polygon points="'+w+',24 '+(w+25)+',0 '+(w+25)+','+(h-24)+' '+w+','+h+'" fill="#3d4851" stroke="#28323a"/><rect x="0" y="24" width="'+w+'" height="'+(h-24)+'" rx="4" fill="'+front+'" stroke="#111820" stroke-width="3"/><rect x="13" y="38" width="'+(w-26)+'" height="'+(h-54)+'" fill="#111820" stroke="#536474"/>';
 const n=Math.min(10,Math.max(3,count||3));for(let i=0;i<n;i++){const yy=48+i*((h-76)/n);s+='<rect x="20" y="'+yy.toFixed(1)+'" width="'+(w-40)+'" height="13" rx="2" fill="#334653" stroke="#72808a"/><circle cx="31" cy="'+(yy+6.5).toFixed(1)+'" r="2.5" fill="'+accent+'"/>';for(let p=0;p<5;p++)s+='<rect x="'+(88+p*11)+'" y="'+(yy+4).toFixed(1)+'" width="7" height="4" fill="#879aa9"/>'}
 s+='<text x="'+(w/2)+'" y="-10" text-anchor="middle" font-size="15" font-weight="800" fill="#173b69">'+esc(label)+'</text><text x="'+(w/2)+'" y="'+(h+18)+'" text-anchor="middle" font-size="11" font-weight="700" fill="#3b526a">×'+esc(count)+'</text></g>';return s;
}
function logicSvg(d){
 const L=T(),leafR=Math.max(1,Math.ceil(d.leaf/8)),spineR=Math.max(0,Math.ceil(d.spine/16)),coreR=Math.max(0,Math.ceil(d.core/16));
 let s='<svg viewBox="0 0 1600 900" role="img" aria-label="'+esc(L.logicTitle)+'"><defs><filter id="v68sh"><feDropShadow dx="4" dy="6" stdDeviation="4" flood-opacity=".25"/></filter><marker id="v68arr" markerWidth="9" markerHeight="9" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#3f658b"/></marker></defs><rect width="1600" height="900" rx="18" fill="#f4f7f9"/><text x="800" y="40" text-anchor="middle" fill="#173d6b" font-size="23" font-weight="900">'+esc(L.logicTitle)+'</text><text x="800" y="66" text-anchor="middle" fill="#566d83" font-size="12">'+esc(d.topology)+' · representative rail view · '+esc(d.speed)+'</text>';
 // representative racks
 s+=rackSvg(210,590,'Leaf Rack 1',Math.min(8,d.leaf||1),'leaf')+rackSvg(1215,590,'Leaf Rack '+Math.max(1,leafR),Math.min(8,d.leaf||1),'leaf');
 if(d.spine>0){s+=rackSvg(330,160,'Spine Rack 1',Math.min(16,d.spine),'spine')+rackSvg(1080,160,'Spine Rack '+Math.max(1,spineR),Math.min(16,d.spine),'spine')}
 if(d.core>0){s+=rackSvg(715,145,'Core Rack',Math.min(16,d.core),'core')}
 // compute rack groups
 const groups=[{x:45,label:'Compute Rack 1'},{x:1375,label:'Compute Rack '+Math.max(1,d.racks)}];for(const g of groups){s+='<g transform="translate('+g.x+',620)"><rect width="130" height="180" rx="6" fill="#202a32" stroke="#0d141b" stroke-width="3"/><rect x="12" y="14" width="106" height="146" fill="#0c1319" stroke="#667786"/>';for(let i=0;i<5;i++){const yy=24+i*26;s+='<rect x="20" y="'+yy+'" width="90" height="18" rx="2" fill="#2b3c49"/><circle cx="30" cy="'+(yy+9)+'" r="3" fill="#47b2ff"/>'}s+='<text x="65" y="-12" text-anchor="middle" font-size="13" font-weight="800" fill="#173b69">'+esc(g.label)+'</text></g>'}
 // links: compute->leaf
 const line=(x1,y1,x2,y2,dash,w=2)=>'<path d="M'+x1+' '+y1+' L'+x2+' '+y2+'" fill="none" stroke="#46637d" stroke-width="'+w+'" '+(dash?'stroke-dasharray="6 5"':'')+' opacity=".82"/>';
 s+=line(175,705,292,705,false,3)+line(1375,705,1380,705,false,3);
 // leaf-spine dense representative
 if(d.spine>0){const starts=[[375,610],[1340,610]],ends=[[412,390],[1162,390]];for(const a of starts)for(const b of ends)s+=line(a[0],a[1],b[0],b[1],true,2)}
 if(d.core>0&&d.spine>0){s+=line(493,390,797,375,false,3)+line(1162,390,880,375,false,3)}
 // labels and ellipses
 s+='<g fill="#b3bbc3"><circle cx="660" cy="265" r="15"/><circle cx="705" cy="265" r="15"/><circle cx="950" cy="265" r="15"/><circle cx="995" cy="265" r="15"/><circle cx="705" cy="695" r="15"/><circle cx="750" cy="695" r="15"/><circle cx="950" cy="695" r="15"/><circle cx="995" cy="695" r="15"/></g>';
 s+='<rect x="590" y="500" width="420" height="84" rx="12" fill="#e7edf3" stroke="#9fb0bf"/><text x="800" y="527" text-anchor="middle" font-size="14" font-weight="900" fill="#234766">Current design summary</text><text x="800" y="552" text-anchor="middle" font-size="12" fill="#536a7f">Leaf '+d.leaf+' · Spine '+d.spine+' · Core '+d.core+' · Compute racks '+d.racks+'</text><text x="800" y="572" text-anchor="middle" font-size="11" fill="#61778b">Representative rail paths are compressed with ellipses; counts are current BOM values.</text>';
 s+='<g><line x1="600" y1="845" x2="665" y2="845" stroke="#46637d" stroke-width="2" stroke-dasharray="6 5"/><text x="675" y="849" fill="#385470" font-size="11">Leaf ↔ Spine logical link</text><line x1="930" y1="845" x2="995" y2="845" stroke="#46637d" stroke-width="3"/><text x="1005" y="849" fill="#385470" font-size="11">Spine ↔ Core / Node ↔ Leaf</text></g></svg>';
 return s;
}
function cablingSvg(d){
 const L=T(),conn=esc(d.connector),trunk=esc(d.trunk),prod=esc(d.product),detail=esc(d.connectorDetail);
 let s='<svg viewBox="0 0 1600 860" role="img" aria-label="'+esc(L.structuredTitle)+'"><defs><linearGradient id="v68cab" x1="0" x2="1"><stop offset="0" stop-color="#26333c"/><stop offset=".5" stop-color="#5a6873"/><stop offset="1" stop-color="#26333c"/></linearGradient><filter id="v68ds"><feDropShadow dx="3" dy="5" stdDeviation="4" flood-opacity=".23"/></filter></defs><rect width="1600" height="860" rx="18" fill="#f7f8f9"/><text x="800" y="38" text-anchor="middle" fill="#1d4676" font-size="22" font-weight="900">'+esc(L.structuredTitle)+'</text><text x="800" y="64" text-anchor="middle" fill="#607389" font-size="12">'+esc(d.chain)+' · '+conn+' · '+trunk+'</text>';
 // active racks
 s+=rackSvg(55,235,'Source Rack',8,'leaf')+rackSvg(1375,235,'Destination Rack',Math.max(1,Math.min(16,d.spine||d.core||8)),'spine');
 // jumpers left/right
 const fiberColors=['#29a3ff','#f8b84d','#54c36d','#a678ff','#ef6575','#2cc5bb','#ff8d3c','#6a8cff'];
 for(let i=0;i<8;i++){const y=330+i*13;s+='<path d="M220 '+y+' C300 '+y+' 300 '+(270+i*12)+' 385 '+(270+i*12)+'" fill="none" stroke="'+fiberColors[i]+'" stroke-width="2.4"/><path d="M1215 '+(270+i*12)+' C1300 '+(270+i*12)+' 1300 '+y+' 1375 '+y+'" fill="none" stroke="'+fiberColors[i]+'" stroke-width="2.4"/>'}
 // panels
 function panel(x,label){let p='<g transform="translate('+x+',220)"><rect width="220" height="230" rx="8" fill="#202930" stroke="#576774" stroke-width="3" filter="url(#v68ds)"/><rect x="16" y="22" width="188" height="170" fill="#0d151b" stroke="#7d8a95"/>';for(let r=0;r<4;r++)for(let c=0;c<4;c++){const xx=30+c*43,yy=40+r*36;p+='<rect x="'+xx+'" y="'+yy+'" width="31" height="20" rx="3" fill="#424f58" stroke="#9aa5ae"/><circle cx="'+(xx+8)+'" cy="'+(yy+10)+'" r="3" fill="#4eb9ff"/>'}p+='<text x="110" y="216" text-anchor="middle" font-size="13" font-weight="900" fill="#eaf2f8">'+label+'</text></g>';return p}
 s+=panel(385,'Patch Panel A')+panel(995,'Patch Panel B');
 // trunk coil
 s+='<g transform="translate(670,240)"><circle cx="130" cy="105" r="100" fill="none" stroke="#d6a82e" stroke-width="15"/><circle cx="130" cy="105" r="74" fill="none" stroke="#f1cd5c" stroke-width="10"/><circle cx="130" cy="105" r="49" fill="none" stroke="#d8a226" stroke-width="8"/><text x="130" y="103" text-anchor="middle" font-size="16" font-weight="900" fill="#6a5215">'+trunk+'</text><text x="130" y="125" text-anchor="middle" font-size="11" fill="#7c6828">High-fiber trunk backbone</text></g>';
 // panel to trunk bundles
 for(let i=0;i<8;i++){const y=270+i*12;s+='<path d="M605 '+y+' C640 '+y+' 640 '+(280+i*12)+' 670 '+(280+i*12)+'" fill="none" stroke="'+fiberColors[i]+'" stroke-width="3"/><path d="M930 '+(280+i*12)+' C965 '+(280+i*12)+' 965 '+y+' 995 '+y+'" fill="none" stroke="'+fiberColors[i]+'" stroke-width="3"/>'}
 // level banners
 const band=(x,y,w,title,body,color)=>'<g><rect x="'+x+'" y="'+y+'" width="'+w+'" height="74" rx="11" fill="'+color+'" stroke="#b7c3cd"/><text x="'+(x+14)+'" y="'+(y+24)+'" font-size="14" font-weight="900" fill="#173b69">'+esc(title)+'</text><text x="'+(x+14)+'" y="'+(y+46)+'" font-size="11" fill="#52697e">'+esc(body)+'</text></g>';
 s+=band(55,545,460,L.levelA,L.refA,'#eaf4fb')+band(570,545,460,L.levelB,L.refB,'#fff8e4')+band(1085,545,460,L.levelC,L.refC,'#edf2ff');
 // current product details
 s+='<rect x="175" y="665" width="1250" height="130" rx="13" fill="#eef2f5" stroke="#afbdc9"/><text x="205" y="695" font-size="14" font-weight="900" fill="#244968">'+esc(L.current)+'</text><text x="205" y="723" font-size="12" fill="#536a7e">'+esc(L.connector)+': '+conn+'</text><text x="205" y="747" font-size="12" fill="#536a7e">Detail: '+detail+'</text><text x="205" y="771" font-size="12" fill="#536a7e">'+esc(L.product)+': '+prod+'</text><text x="900" y="723" font-size="12" fill="#536a7e">'+esc(L.trunk)+': '+trunk+'</text><text x="900" y="747" font-size="12" fill="#536a7e">Leaf / Spine / Core: '+d.leaf+' / '+d.spine+' / '+d.core+'</text><text x="900" y="771" font-size="12" fill="#536a7e">Compute racks: '+d.racks+' · Compute speed: '+esc(d.speed)+'</text></svg>';
 return s;
}
function openLogic(){
 const L=T(),d=current(),m=modal();m.querySelector('h2').textContent=L.logicTitle;m.querySelector('p').textContent=L.logicSub;m.querySelector('.v68-body').innerHTML=summaryHtml(d)+'<div class="v68-svg-wrap">'+logicSvg(d)+'</div><div class="v68-note">'+esc(L.noteLogic)+'</div>';m.hidden=false;
}
function openStructured(){
 const L=T(),d=current(),m=modal();m.querySelector('h2').textContent=L.structuredTitle;m.querySelector('p').textContent=L.structuredSub;m.querySelector('.v68-body').innerHTML=summaryHtml(d)+'<div class="v68-svg-wrap">'+cablingSvg(d)+'</div><div class="v68-levels"><div class="v68-level"><b>'+esc(L.levelA)+'</b><span>'+esc(L.refA)+'</span></div><div class="v68-level"><b>'+esc(L.levelB)+'</b><span>'+esc(L.refB)+'</span></div><div class="v68-level"><b>'+esc(L.levelC)+'</b><span>'+esc(L.refC)+'</span></div></div><div class="v68-note">'+esc(L.noteCab)+'</div>';m.hidden=false;
}
function gotoCooling(){
 const p=q('v53-cooling-architecture');if(p){p.scrollIntoView({behavior:'smooth',block:'start'});p.classList.remove('v68-cooling-pulse');void p.offsetWidth;p.classList.add('v68-cooling-pulse');return}
 for(let i=0;i<12;i++)setTimeout(()=>{const x=q('v53-cooling-architecture');if(x){x.scrollIntoView({behavior:'smooth',block:'start'});x.classList.add('v68-cooling-pulse')}},i*250);
}
function addTopButtons(){
 const bar=q('v63-top-links');if(!bar)return false;
 const L=T();
 const defs=[['v68-cooling-btn',L.cool,gotoCooling],['v68-logic-btn',L.logic,openLogic],['v68-structured-btn',L.structured,openStructured]];
 defs.forEach(([id,label,fn])=>{let b=q(id);if(!b){b=document.createElement('button');b.type='button';b.id=id;b.innerHTML='<span class="v68-icon-dot"></span><span class="v68-label"></span>';bar.appendChild(b);b.onclick=fn}const t=b.querySelector('.v68-label');if(t)t.textContent=label});
 return true;
}

// Animated cold-supply / warm-return particles over the existing realistic cooling canvas.
const FLOW={
 cold:[[330,390],[430,390],[665,420],[735,420],[1350,420],[1418,390]],
 warm:[[1418,445],[1350,465],[665,465],[430,465],[330,465]],
 coldRisers:[[[808,615],[808,515]],[[1018,615],[1018,515]],[[1228,615],[1228,515]]],
 warmRisers:[[[958,515],[958,615]],[[1168,515],[1168,615]],[[1378,515],[1378,615]]]
};
let anim=null;
function lenSeg(a,b){return Math.hypot(b[0]-a[0],b[1]-a[1])}
function pointOn(path,t){let total=0;for(let i=0;i<path.length-1;i++)total+=lenSeg(path[i],path[i+1]);let d=((t%1)+1)%1*total;for(let i=0;i<path.length-1;i++){const a=path[i],b=path[i+1],l=lenSeg(a,b);if(d<=l){const r=l?d/l:0;return[a[0]+(b[0]-a[0])*r,a[1]+(b[1]-a[1])*r]}d-=l}return path[path.length-1]}
function mountCoolingFlow(){
 const host=document.querySelector('.v631-cool3d');if(!host)return false;
 let c=host.querySelector('.v68-flow-overlay');if(!c){c=document.createElement('canvas');c.className='v68-flow-overlay';host.appendChild(c)}
 const dpr=Math.min(2.5,window.devicePixelRatio||1);if(c.width!==Math.round(1600*dpr)||c.height!==Math.round(900*dpr)){c.width=Math.round(1600*dpr);c.height=Math.round(900*dpr)}
 if(anim)return true;const ctx=c.getContext('2d');let start=performance.now();
 function dot(x,y,color,r,phase){ctx.save();ctx.shadowColor=color;ctx.shadowBlur=18+7*Math.sin(phase);ctx.fillStyle=color;ctx.globalAlpha=.94;ctx.beginPath();ctx.arc(x,y,r,0,Math.PI*2);ctx.fill();ctx.globalAlpha=.7;ctx.strokeStyle='#ffffff';ctx.lineWidth=1.2;ctx.beginPath();ctx.moveTo(x-r*1.7,y);ctx.lineTo(x+r*1.7,y);ctx.moveTo(x,y-r*1.7);ctx.lineTo(x,y+r*1.7);ctx.stroke();ctx.restore()}
 function frame(now){const hostNow=document.querySelector('.v631-cool3d'),canvas=hostNow&&hostNow.querySelector('.v68-flow-overlay');if(!canvas){anim=null;return}const dprNow=Math.min(2.5,window.devicePixelRatio||1);const cx=canvas.getContext('2d');cx.setTransform(dprNow,0,0,dprNow,0,0);cx.clearRect(0,0,1600,900);const t=(now-start)/4200;for(let i=0;i<13;i++){let p=pointOn(FLOW.cold,t+i/13);dot(p[0],p[1],'#56c9ff',4.2,now/240+i)}for(let i=0;i<13;i++){let p=pointOn(FLOW.warm,t+i/13);dot(p[0],p[1],'#ff7a67',4.2,now/240+i)}FLOW.coldRisers.forEach((p,j)=>{const x=pointOn(p,t+j*.17);dot(x[0],x[1],'#56c9ff',3.5,now/220+j)});FLOW.warmRisers.forEach((p,j)=>{const x=pointOn(p,t+j*.17);dot(x[0],x[1],'#ff7a67',3.5,now/220+j)});anim=requestAnimationFrame(frame)}
 anim=requestAnimationFrame(frame);return true;
}
function refresh(){addTopButtons();mountCoolingFlow()}
function start(){
 let tries=0;const boot=setInterval(()=>{tries++;refresh();if((q('v63-top-links')&&document.querySelector('.v631-cool3d'))||tries>35)clearInterval(boot)},350);
 document.addEventListener('click',e=>{if(e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],.lang,.language'))setTimeout(addTopButtons,80)},true);
 setInterval(()=>{addTopButtons();if(!anim)mountCoolingFlow()},1800);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();