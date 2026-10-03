(() => {
'use strict';
const V53_VERSION='7.3.2';

function textV53(e){return(e&&(e.innerText||e.textContent)||'').replace(/\s+/g,' ').trim()}
function nV53(v,fallback=0){const m=String(v==null?'':v).replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);return m?Number(m[0]):fallback}
function elV53(id){return document.getElementById(id)}
function langTokenV53(v){
 v=String(v||'').trim().toLowerCase();
 if(v==='ko'||v==='kr'||v.includes('한국')||v.includes('korean'))return'ko';
 if(v==='zh'||v==='cn'||v.includes('中文')||v.includes('chinese'))return'zh';
 if(v==='ja'||v==='jp'||v.includes('日本')||v.includes('japanese'))return'ja';
 if(v==='de'||v.includes('deutsch')||v.includes('german'))return'de';
 return'en';
}
function currentLangV53(){
 if(window.__dcBomUiLang)return langTokenV53(window.__dcBomUiLang);
 const d=document.documentElement.getAttribute('data-dc-bom-ui-lang');if(d)return langTokenV53(d);
 const active=document.querySelector('[data-lang].active,[data-language].active,[aria-pressed="true"],.lang.active,.language.active');
 if(active)return langTokenV53((active.getAttribute('data-lang')||'')+' '+(active.getAttribute('data-language')||'')+' '+textV53(active));
 return langTokenV53(document.documentElement.lang||'ko');
}

const COOLING_I18N={
 ko:{
  title:'냉각 아키텍처',sub:'데이터센터 설비에서 GPU 콜드 플레이트까지의 End-to-End 액체 냉각 경로',
  supply:'공급 (저온, 예: 20~25°C)',return:'환수 (고온, 예: 30~40°C)',
  facility:'데이터센터 설비',facilitySub:'대기 또는 외부 열원으로 방열',plant:'Cooling Tower / Dry Cooler / Chiller',facilityLoop:'설비수 1차 루프 (예: 25~35°C)',
  cdu:'CDU (냉각수 분배 장치)',cduSub:'설비측과 IT측 루프 사이 열교환',hx:'열교환기',primary:'설비수 (1차 루프)',secondary:'IT 냉각수 (2차 루프)',
  manifold:'랙 매니폴드',manifoldSub:'각 서버로 냉각수 공급·회수',supplyManifold:'공급 매니폴드',returnManifold:'환수 매니폴드',
  rack:'컴퓨트 랙',example:'예시',gpuServers:'GPU 서버',gpuServer:'GPU 서버',coldPlate:'콜드 플레이트',qd:'QD (공급/환수)',
  how:'동작 원리',design:'설계 예시',notes:'주요 메모',
  steps:['GPU/CPU가 동작 중 열을 발생시킵니다.','콜드 플레이트가 GPU/CPU의 열을 흡수합니다.','가열된 냉각수가 랙 매니폴드로 환수됩니다.','CDU 열교환기가 IT 루프의 열을 설비수로 전달합니다.','Cooling tower / dry cooler / chiller가 외부로 열을 방출합니다.','냉각된 냉각수가 다시 GPU로 공급됩니다.'],
  noteList:['설비수와 IT 냉각수는 열교환기로 분리되며 서로 혼합되지 않습니다.','QD(Quick Disconnect)는 랙 정비와 서버 교체를 쉽게 합니다.','유량과 ΔT는 냉각수 종류, GPU 열부하, 설계 목표에 따라 달라집니다.','인프라 계획은 Design-Max를 기본으로 하고 Peak-Provisioning은 전력·냉각 여유 검토에 사용합니다.'],
  parameter:'항목',typical:'일반 운용 (Typical)',designMax:'설계 최대 (Design-Max)',peak:'피크 인프라 (Peak)',
  pSystem:'시스템당 IT 전력',pRack:'랙 IT 전력',deltaT:'냉각수 ΔT (공급→환수)',flow:'필요 유량 (물 환산)',cduCapacity:'CDU 초기 용량',
  tableNote:'* 유량은 물 기준 cp=4.186 kJ/kg·K로 환산한 1차 추정값입니다. CDU 초기 용량은 열부하 × 1.15 여유율로 계산하며 실제 선정은 N+1, 현장 조건, 냉각수 종류와 제조사 곡선을 확인해야 합니다.',
  more:'추가 시스템은 도식에서 생략'
 },
 en:{
  title:'Cooling architecture',sub:'End-to-end liquid cooling path from facility equipment to GPU cold plates',
  supply:'Supply (Cold, e.g. 20–25°C)',return:'Return (Warm, e.g. 30–40°C)',
  facility:'DATA CENTER FACILITY',facilitySub:'Reject heat to atmosphere / external sink',plant:'Cooling Tower / Dry Cooler / Chiller',facilityLoop:'Facility water primary loop (e.g. 25–35°C)',
  cdu:'CDU (Coolant Distribution Unit)',cduSub:'Heat exchanger between facility and IT loops',hx:'Heat Exchanger',primary:'Facility Water (Primary Loop)',secondary:'IT Coolant (Secondary Loop)',
  manifold:'Rack Manifold',manifoldSub:'Distribute and collect coolant',supplyManifold:'Supply Manifold',returnManifold:'Return Manifold',
  rack:'Compute Rack',example:'Example',gpuServers:'GPU Servers',gpuServer:'GPU Server',coldPlate:'Cold Plate',qd:'QD (Supply/Return)',
  how:'How it works',design:'Design example',notes:'Key notes',
  steps:['GPU/CPU generates heat during operation.','Cold plate absorbs heat from GPU/CPU.','Heated coolant returns to the rack manifold.','CDU transfers heat to facility water through a heat exchanger.','Cooling tower / dry cooler / chiller rejects heat outside.','Cooled IT coolant is supplied back to the GPUs.'],
  noteList:['Facility water and IT coolant are separated by a heat exchanger and do not mix.','QD (Quick Disconnect) supports rack service and server replacement.','Flow rate and ΔT depend on coolant type, GPU heat load and design target.','Use Design-Max for infrastructure planning and Peak-Provisioning for power/cooling envelope review.'],
  parameter:'Parameter',typical:'Typical',designMax:'Design-Max',peak:'Peak-Provisioning',
  pSystem:'IT Power per system',pRack:'Rack IT Power',deltaT:'Coolant ΔT (supply→return)',flow:'Required Flow (water eq.)',cduCapacity:'CDU First-Pass Capacity',
  tableNote:'* Flow is a first-pass water-equivalent estimate using cp=4.186 kJ/kg·K. CDU first-pass capacity uses heat load × 1.15; actual selection requires N+1, site conditions, coolant properties and vendor curves.',
  more:'additional systems omitted from diagram'
 },
 zh:{
  title:'冷却架构',sub:'从数据中心设施到 GPU 冷板的端到端液冷路径',
  supply:'供液（低温，例如 20~25°C）',return:'回液（高温，例如 30~40°C）',
  facility:'数据中心设施',facilitySub:'将热量排放到外部环境',plant:'冷却塔 / 干冷器 / 冷水机',facilityLoop:'设施水一次回路（例如 25~35°C）',
  cdu:'CDU（冷却液分配单元）',cduSub:'设施侧与 IT 侧回路之间的换热器',hx:'换热器',primary:'设施水（一次回路）',secondary:'IT 冷却液（二次回路）',
  manifold:'机架歧管',manifoldSub:'向服务器分配并回收冷却液',supplyManifold:'供液歧管',returnManifold:'回液歧管',
  rack:'计算机架',example:'示例',gpuServers:'GPU 服务器',gpuServer:'GPU 服务器',coldPlate:'冷板',qd:'QD（供液/回液）',
  how:'工作原理',design:'设计示例',notes:'关键说明',
  steps:['GPU/CPU 在运行过程中产生热量。','冷板吸收 GPU/CPU 的热量。','升温后的冷却液返回机架歧管。','CDU 通过换热器将 IT 回路热量传递给设施水。','冷却塔 / 干冷器 / 冷水机将热量排到室外。','冷却后的 IT 冷却液再次供应给 GPU。'],
  noteList:['设施水与 IT 冷却液通过换热器隔离，不相互混合。','QD（快速接头）便于机架维护和服务器更换。','流量和 ΔT 取决于冷却液类型、GPU 热负载和设计目标。','基础设施规划采用 Design-Max，Peak-Provisioning 用于检查电力和冷却余量。'],
  parameter:'参数',typical:'典型运行',designMax:'设计最大',peak:'峰值预留',
  pSystem:'每系统 IT 功率',pRack:'机架 IT 功率',deltaT:'冷却液 ΔT（供→回）',flow:'所需流量（水当量）',cduCapacity:'CDU 初步容量',
  tableNote:'* 流量按水 cp=4.186 kJ/kg·K 进行初步估算。CDU 初步容量按热负载 × 1.15 计算；实际选型还需确认 N+1、现场条件、冷却液特性及厂商曲线。',
  more:'其余系统在图中省略'
 },
 ja:{
  title:'冷却アーキテクチャ',sub:'データセンター設備から GPU コールドプレートまでの End-to-End 液冷経路',
  supply:'供給（低温、例 20~25°C）',return:'戻り（高温、例 30~40°C）',
  facility:'データセンター設備',facilitySub:'大気・外部ヒートシンクへ放熱',plant:'Cooling Tower / Dry Cooler / Chiller',facilityLoop:'設備水一次ループ（例 25~35°C）',
  cdu:'CDU（Coolant Distribution Unit）',cduSub:'設備側と IT 側ループ間の熱交換',hx:'熱交換器',primary:'設備水（一次ループ）',secondary:'IT 冷却水（二次ループ）',
  manifold:'ラックマニホールド',manifoldSub:'各サーバーへ冷却水を分配・回収',supplyManifold:'供給マニホールド',returnManifold:'戻りマニホールド',
  rack:'コンピュートラック',example:'例',gpuServers:'GPU サーバー',gpuServer:'GPU サーバー',coldPlate:'コールドプレート',qd:'QD（供給/戻り）',
  how:'動作原理',design:'設計例',notes:'主な注意点',
  steps:['GPU/CPU が動作中に熱を発生します。','コールドプレートが GPU/CPU の熱を吸収します。','加熱された冷却水がラックマニホールドへ戻ります。','CDU の熱交換器が IT ループの熱を設備水へ移します。','Cooling tower / dry cooler / chiller が外部へ放熱します。','冷却された IT 冷却水が GPU へ再供給されます。'],
  noteList:['設備水と IT 冷却水は熱交換器で分離され、混合しません。','QD（Quick Disconnect）によりラック保守やサーバー交換が容易になります。','流量と ΔT は冷却液、GPU 熱負荷、設計目標で変化します。','インフラ計画は Design-Max を基本とし、Peak-Provisioning で電力・冷却余裕を確認します。'],
  parameter:'項目',typical:'Typical',designMax:'Design-Max',peak:'Peak-Provisioning',
  pSystem:'システム当たり IT 電力',pRack:'ラック IT 電力',deltaT:'冷却水 ΔT（供給→戻り）',flow:'必要流量（水換算）',cduCapacity:'CDU 初期容量',
  tableNote:'* 流量は水 cp=4.186 kJ/kg·K による一次推定です。CDU 初期容量は熱負荷 × 1.15 とし、実選定では N+1、現場条件、冷却液特性、ベンダーカーブを確認します。',
  more:'追加システムは図では省略'
 },
 de:{
  title:'Kühlarchitektur',sub:'End-to-End-Flüssigkeitskühlung von der Rechenzentrumsanlage bis zu den GPU-Cold-Plates',
  supply:'Vorlauf (kalt, z. B. 20~25°C)',return:'Rücklauf (warm, z. B. 30~40°C)',
  facility:'RECHENZENTRUMSANLAGE',facilitySub:'Wärme an Umgebung / Wärmesenke abgeben',plant:'Kühlturm / Trockenkühler / Chiller',facilityLoop:'Primärkreislauf Anlagenwasser (z. B. 25~35°C)',
  cdu:'CDU (Coolant Distribution Unit)',cduSub:'Wärmetauscher zwischen Anlagen- und IT-Kreis',hx:'Wärmetauscher',primary:'Anlagenwasser (Primärkreis)',secondary:'IT-Kühlmittel (Sekundärkreis)',
  manifold:'Rack-Verteiler',manifoldSub:'Kühlmittel verteilen und sammeln',supplyManifold:'Vorlaufverteiler',returnManifold:'Rücklaufverteiler',
  rack:'Compute-Rack',example:'Beispiel',gpuServers:'GPU-Server',gpuServer:'GPU-Server',coldPlate:'Cold Plate',qd:'QD (Vorlauf/Rücklauf)',
  how:'Funktionsweise',design:'Auslegungsbeispiel',notes:'Wichtige Hinweise',
  steps:['GPU/CPU erzeugt im Betrieb Wärme.','Die Cold Plate nimmt Wärme von GPU/CPU auf.','Erwärmtes Kühlmittel fließt zum Rack-Verteiler zurück.','Die CDU überträgt Wärme über den Wärmetauscher auf das Anlagenwasser.','Kühlturm / Trockenkühler / Chiller gibt Wärme nach außen ab.','Abgekühltes IT-Kühlmittel wird wieder zu den GPUs geführt.'],
  noteList:['Anlagenwasser und IT-Kühlmittel sind durch den Wärmetauscher getrennt und vermischen sich nicht.','QD (Quick Disconnect) erleichtert Rack-Service und Serveraustausch.','Volumenstrom und ΔT hängen von Kühlmittel, GPU-Wärmelast und Auslegungsziel ab.','Design-Max für Infrastrukturplanung verwenden; Peak-Provisioning für Leistungs- und Kühlreserve prüfen.'],
  parameter:'Parameter',typical:'Typical',designMax:'Design-Max',peak:'Peak-Provisioning',
  pSystem:'IT-Leistung je System',pRack:'Rack-IT-Leistung',deltaT:'Kühlmittel ΔT (Vorlauf→Rücklauf)',flow:'Erforderlicher Volumenstrom (Wasseräquiv.)',cduCapacity:'CDU-Vorauslegung',
  tableNote:'* Volumenstrom als Wasseräquivalent mit cp=4.186 kJ/kg·K. CDU-Vorauslegung = Wärmelast × 1.15; reale Auswahl erfordert N+1, Standortbedingungen, Kühlmitteleigenschaften und Herstellerkennlinien.',
  more:'weitere Systeme im Diagramm ausgelassen'
 }
};

function selectedSystemV53(){if(window.DCDesign?.systemProfile)return window.DCDesign.systemProfile;
 try{if(typeof getSelectedSystem==='function')return getSelectedSystem()}catch(e){}
 const id=elV53('systemId')?elV53('systemId').value:'system';
 try{if(typeof systems!=='undefined'&&systems[id])return systems[id]}catch(e){}
 return{name:String(id||'GPU System').toUpperCase(),gpu:String(id||'GPU').toUpperCase(),ru:10,power:10};
}
function rackModelV53(){
 const s=selectedSystemV53(),units=window.DCDesign?.summary?.systemUnits||0,racks=window.DCDesign?.summary?.computeRacks||1,perRack=window.DCDesign?.racks?.find(x=>x.role==='Compute')?.units||0;
 const rackU=Math.max(24,nV53(elV53('rackRU')?elV53('rackRU').value:48,48));
 let sw={name:'Leaf Switch',model:'Leaf Switch',ru:2};
 try{const id=elV53('switchId')?elV53('switchId').value:'';if(typeof switches!=='undefined'&&switches[id])sw=switches[id]}catch(e){}
 return{s,units,racks,perRack,rackU,sw};
}
function coolingModelV53(){
 const r=rackModelV53(),s=r.s,typical=(s.powerTypical==null?null:Number(s.powerTypical)),design=Number(s.powerDesignMax!=null?s.powerDesignMax:(s.power!=null?s.power:0)),peak=Number(s.powerPeakProvisioning!=null?s.powerPeakProvisioning:design);
 const dt=10,cp=4.186,margin=1.15,flow=k=>k==null?null:(k*60/(cp*dt)),cap=k=>k==null?null:k*margin;
 return{...r,typical,design,peak,dt,margin,rackTypical:typical==null?null:typical*r.perRack,rackDesign:design*r.perRack,rackPeak:peak*r.perRack,
  flowTypical:flow(typical==null?null:typical*r.perRack),flowDesign:flow(design*r.perRack),flowPeak:flow(peak*r.perRack),
  cduTypical:cap(typical==null?null:typical*r.perRack),cduDesign:cap(design*r.perRack),cduPeak:cap(peak*r.perRack)};
}
function fmtV53(v,u,d=1){return v==null||!Number.isFinite(v)?'—':v.toFixed(d)+(u?' '+u:'')}
function escV53(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}

function findCoolingHeadingV53(){
 const rx=/cooling\s*architecture|냉각\s*(아키텍처|구성|architecture)|冷却アーキテクチャ|冷却架构|kühlarchitektur/i;
 return Array.from(document.querySelectorAll('h1,h2,h3,h4,.section-title,.card-title')).find(x=>!x.closest('#v53-cooling-architecture')&&rx.test(textV53(x)))||null;
}
function mountCoolingArchitectureV53(){
 let p=elV53('v53-cooling-architecture');if(p)return p;
 const legacyV51=elV53('v51-cooling-architecture');
 p=document.createElement('section');p.id='v53-cooling-architecture';
 if(legacyV51){
  legacyV51.style.display='none';
  legacyV51.setAttribute('data-v53-replaced','1');
  legacyV51.insertAdjacentElement('afterend',p);
  return p;
 }
 const h=findCoolingHeadingV53();
 if(h){
  h.style.display='none';h.setAttribute('data-v53-legacy-title','1');
  let n=h.nextElementSibling,steps=0;
  while(n&&steps<4){const next=n.nextElementSibling,visual=(n.tagName==='SVG'||n.tagName==='CANVAS'||(n.querySelector&&n.querySelector('svg,canvas'))),interactive=n.querySelector&&n.querySelector('input,select,button,textarea');if(visual&&!interactive){n.style.display='none';n.setAttribute('data-v53-legacy-visual','1');break}n=next;steps++}
  h.insertAdjacentElement('afterend',p);
 }else{
  const anchor=elV53('v48-power-model')||elV53('v48-rack-twin')||document.querySelector('main,.wrap,.container')||document.body;
  if(anchor===document.body)anchor.appendChild(p);else anchor.insertAdjacentElement('afterend',p);
 }
 return p;
}
function drawCoolingSvgV53(svg,L,m){
 const blue='#28a9ff',red='#f3544e',panel='#0a2038',edge='#5b7894',white='#eef6ff',muted='#b8c9dc',shown=Math.min(4,Math.max(1,m.perRack)),sysName=escV53(m.s.name||m.s.gpu||'GPU');
 let s='<defs><marker id="v53-blue-arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="'+blue+'"/></marker><marker id="v53-red-arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="'+red+'"/></marker><linearGradient id="v53-metal" x1="0" x2="1"><stop offset="0" stop-color="#26394b"/><stop offset=".5" stop-color="#71808d"/><stop offset="1" stop-color="#233444"/></linearGradient><linearGradient id="v53-copper" x1="0" x2="1"><stop offset="0" stop-color="#7a3f25"/><stop offset=".5" stop-color="#d28a54"/><stop offset="1" stop-color="#8a4729"/></linearGradient><filter id="v53-ds"><feDropShadow dx="3" dy="5" stdDeviation="4" flood-opacity=".35"/></filter></defs>';
 s+='<rect x="5" y="5" width="1490" height="600" rx="18" fill="#06172a" stroke="#385b78"/>';
 [[20,65,300,480],[345,65,370,480],[740,65,230,480],[995,65,480,480]].forEach(p=>s+='<rect x="'+p[0]+'" y="'+p[1]+'" width="'+p[2]+'" height="'+p[3]+'" rx="13" fill="'+panel+'" stroke="'+edge+'" stroke-width="1.3"/>');
 s+='<text x="170" y="98" fill="'+white+'" font-size="20" font-weight="800" text-anchor="middle">'+escV53(L.facility)+'</text><text x="170" y="122" fill="'+muted+'" font-size="12" text-anchor="middle">'+escV53(L.facilitySub)+'</text>';
 for(let i=0;i<3;i++){const x=70+i*72;s+='<ellipse cx="'+x+'" cy="190" rx="32" ry="9" fill="#b6c2cc"/><rect x="'+(x-32)+'" y="190" width="64" height="120" fill="url(#v53-metal)" stroke="#8da0ae"/><ellipse cx="'+x+'" cy="310" rx="32" ry="9" fill="#3a4a58"/>';for(let q=-25;q<=25;q+=25)s+='<line x1="'+(x+q)+'" y1="203" x2="'+(x+q)+'" y2="295" stroke="#9eaeba" opacity=".55"/>'}
 s+='<rect x="52" y="340" width="236" height="55" rx="8" fill="#0b3553" stroke="#37759d"/><text x="170" y="362" fill="'+white+'" font-size="13" font-weight="700" text-anchor="middle">'+escV53(L.plant)+'</text><text x="170" y="383" fill="'+muted+'" font-size="11" text-anchor="middle">'+escV53(L.facilityLoop)+'</text>';
 s+='<path d="M286 360 C330 360 330 345 365 345" fill="none" stroke="'+blue+'" stroke-width="14" stroke-linecap="round"/><path d="M365 430 C330 430 330 410 286 410" fill="none" stroke="'+red+'" stroke-width="14" stroke-linecap="round"/><text x="332" y="325" fill="#9dd8ff" font-size="11" text-anchor="middle">'+escV53(L.primary)+'</text>';
 s+='<rect x="390" y="150" width="285" height="330" rx="6" fill="#172532" stroke="#7e8e9a" stroke-width="4" filter="url(#v53-ds)"/><rect x="408" y="172" width="249" height="286" fill="#0b1620" stroke="#445766"/><rect x="495" y="230" width="85" height="155" rx="5" fill="#a2adb5" stroke="#d5dde3"/>';
 for(let y=242;y<375;y+=9)s+='<line x1="503" y1="'+y+'" x2="572" y2="'+y+'" stroke="#6e7880" stroke-width="2"/>';
 s+='<text x="538" y="310" fill="#14202b" font-size="12" font-weight="800" text-anchor="middle">'+escV53(L.hx)+'</text>';
 [[448,405,blue],[620,405,red]].forEach(v=>{s+='<circle cx="'+v[0]+'" cy="'+v[1]+'" r="28" fill="#424f5a" stroke="#9aa8b3" stroke-width="3"/><circle cx="'+v[0]+'" cy="'+v[1]+'" r="11" fill="#232d35"/><path d="M'+v[0]+' '+(v[1]-28)+' V205" stroke="'+v[2]+'" stroke-width="12" fill="none" stroke-linecap="round"/>'});
 s+='<path d="M448 205 H493 V230" stroke="'+blue+'" stroke-width="12" fill="none"/><path d="M580 230 V205 H620" stroke="'+red+'" stroke-width="12" fill="none"/><text x="530" y="98" fill="'+white+'" font-size="20" font-weight="800" text-anchor="middle">'+escV53(L.cdu)+'</text><text x="530" y="122" fill="'+muted+'" font-size="12" text-anchor="middle">'+escV53(L.cduSub)+'</text><text x="530" y="515" fill="'+muted+'" font-size="11" text-anchor="middle">'+escV53(L.secondary)+'</text>';
 s+='<path d="M675 250 H785" fill="none" stroke="'+blue+'" stroke-width="14" marker-end="url(#v53-blue-arrow)"/><path d="M785 445 H675" fill="none" stroke="'+red+'" stroke-width="14" marker-end="url(#v53-red-arrow)"/>';
 s+='<rect x="775" y="150" width="160" height="330" rx="4" fill="#111a22" stroke="#71808d" stroke-width="4"/><line x1="825" y1="185" x2="825" y2="450" stroke="'+blue+'" stroke-width="16" stroke-linecap="round"/><line x1="885" y1="185" x2="885" y2="450" stroke="'+red+'" stroke-width="16" stroke-linecap="round"/>';
 for(let i=0;i<shown;i++){const y=220+i*60;s+='<line x1="825" y1="'+y+'" x2="930" y2="'+y+'" stroke="'+blue+'" stroke-width="9"/><line x1="885" y1="'+(y+22)+'" x2="930" y2="'+(y+22)+'" stroke="'+red+'" stroke-width="9"/><circle cx="930" cy="'+y+'" r="8" fill="#6dbbe9" stroke="#d8f0ff"/><circle cx="930" cy="'+(y+22)+'" r="8" fill="#e87069" stroke="#ffe0de"/>'}
 s+='<text x="855" y="98" fill="'+white+'" font-size="20" font-weight="800" text-anchor="middle">'+escV53(L.manifold)+'</text><text x="855" y="122" fill="'+muted+'" font-size="12" text-anchor="middle">'+escV53(L.manifoldSub)+'</text><text x="814" y="510" fill="#9dd8ff" font-size="11" text-anchor="middle">'+escV53(L.supplyManifold)+'</text><text x="895" y="510" fill="#ffaaa6" font-size="11" text-anchor="middle">'+escV53(L.returnManifold)+'</text>';
 s+='<rect x="1030" y="145" width="285" height="350" rx="5" fill="#10171e" stroke="#667582" stroke-width="5"/><rect x="1047" y="165" width="251" height="310" fill="#071018" stroke="#273745"/>';
 for(let i=0;i<shown;i++){const y=185+i*70;s+='<rect x="1060" y="'+y+'" width="225" height="56" rx="4" fill="#18232c" stroke="#60707c"/><rect x="1070" y="'+(y+9)+'" width="75" height="38" fill="#0d151b" stroke="#43515c"/>';for(let k=0;k<7;k++)s+='<circle cx="'+(1080+k*9)+'" cy="'+(y+28)+'" r="2" fill="#667b8d"/>';s+='<rect x="1195" y="'+(y+9)+'" width="55" height="38" rx="3" fill="url(#v53-copper)" stroke="#e1a272"/><path d="M1195 '+(y+19)+' C1168 '+(y+19)+' 1168 '+(y+12)+' 1148 '+(y+12)+'" fill="none" stroke="'+blue+'" stroke-width="8"/><path d="M1195 '+(y+38)+' C1168 '+(y+38)+' 1168 '+(y+45)+' 1148 '+(y+45)+'" fill="none" stroke="'+red+'" stroke-width="8"/><text x="1360" y="'+(y+20)+'" fill="'+white+'" font-size="12" font-weight="700">'+escV53(L.gpuServer)+' '+(i+1)+'</text><text x="1360" y="'+(y+39)+'" fill="'+muted+'" font-size="10">'+escV53(L.coldPlate)+' + '+escV53(L.qd)+'</text><path d="M930 '+(220+i*60)+' C980 '+(220+i*60)+' 986 '+(y+18)+' 1030 '+(y+18)+'" fill="none" stroke="'+blue+'" stroke-width="10" marker-end="url(#v53-blue-arrow)"/><path d="M1030 '+(y+40)+' C986 '+(y+40)+' 980 '+(242+i*60)+' 930 '+(242+i*60)+'" fill="none" stroke="'+red+'" stroke-width="10" marker-end="url(#v53-red-arrow)"/>'}
 s+='<text x="1170" y="98" fill="'+white+'" font-size="20" font-weight="800" text-anchor="middle">'+escV53(L.rack)+' ('+escV53(L.example)+': '+m.perRack+' '+escV53(L.gpuServers)+')</text><text x="1170" y="122" fill="'+muted+'" font-size="12" text-anchor="middle">'+sysName+' · '+escV53(L.coldPlate)+'</text>';
 if(m.perRack>shown)s+='<text x="1170" y="525" fill="'+muted+'" font-size="10" text-anchor="middle">+'+(m.perRack-shown)+' · '+escV53(L.more)+'</text>';
 svg.innerHTML=s;
}
function renderCoolingArchitectureV53(){
 const p=mountCoolingArchitectureV53();if(!p)return;const lang=currentLangV53(),L=COOLING_I18N[lang]||COOLING_I18N.ko,m=coolingModelV53(),systemName=escV53(m.s.name||m.s.gpu||'GPU'),val=(v,u)=>fmtV53(v,u,1);
 p.innerHTML='<h3>'+escV53(L.title)+'</h3><div class="v53-cool-sub">'+escV53(L.sub)+'</div><div class="v53-cool-legend"><span class="v53-legend-item"><i class="v53-line supply"></i>'+escV53(L.supply)+'</span><span class="v53-legend-item"><i class="v53-line return"></i>'+escV53(L.return)+'</span></div><div class="v53-cool-svg-wrap"><svg id="v53-cooling-svg" viewBox="0 0 1500 610"></svg></div><div class="v53-cool-bottom"><section class="v53-info-card"><h4>'+escV53(L.how)+'</h4><div class="v53-steps">'+L.steps.map((x,i)=>'<div class="v53-step"><b>'+(i+1)+'</b><span>'+escV53(x)+'</span></div>').join('')+'</div></section><section class="v53-info-card"><h4>'+escV53(L.design)+' ('+systemName+', '+m.perRack+' systems / rack)</h4><table class="v53-cool-table"><thead><tr><th>'+escV53(L.parameter)+'</th><th>'+escV53(L.typical)+'</th><th>'+escV53(L.designMax)+'</th><th>'+escV53(L.peak)+'</th></tr></thead><tbody><tr><td>'+escV53(L.pSystem)+'</td><td>'+val(m.typical,'kW')+'</td><td>'+val(m.design,'kW')+'</td><td>'+val(m.peak,'kW')+'</td></tr><tr><td>'+escV53(L.pRack)+'</td><td>'+val(m.rackTypical,'kW')+'</td><td>'+val(m.rackDesign,'kW')+'</td><td>'+val(m.rackPeak,'kW')+'</td></tr><tr><td>'+escV53(L.deltaT)+'</td><td>'+m.dt+' °C</td><td>'+m.dt+' °C</td><td>'+m.dt+' °C</td></tr><tr><td>'+escV53(L.flow)+'</td><td>'+val(m.flowTypical,'L/min')+'</td><td>'+val(m.flowDesign,'L/min')+'</td><td>'+val(m.flowPeak,'L/min')+'</td></tr><tr><td>'+escV53(L.cduCapacity)+'</td><td>'+val(m.cduTypical,'kW')+'</td><td>'+val(m.cduDesign,'kW')+'</td><td>'+val(m.cduPeak,'kW')+'</td></tr></tbody></table><div class="v53-table-note">'+escV53(L.tableNote)+'</div></section><section class="v53-info-card"><h4>'+escV53(L.notes)+'</h4><ul class="v53-notes">'+L.noteList.map(x=>'<li>'+escV53(x)+'</li>').join('')+'</ul></section></div>';
 drawCoolingSvgV53(elV53('v53-cooling-svg'),L,m);
 p.setAttribute('data-v53-lang',lang);
}

// Rack view with explicit 4 renderer functions and shared primitives
const RACK_VIEW_V53={mode:'2d',side:'front'};let RACK_CTX_V53=null;
function svgEV53(svg,t,o){const e=document.createElementNS('http://www.w3.org/2000/svg',t);Object.keys(o||{}).forEach(k=>e.setAttribute(k,o[k]));svg.appendChild(e);return e}
function svgTV53(svg,x,y,s,size=10,fill='#2c3f52',anchor='start',weight='600'){const e=svgEV53(svg,'text',{x,y,fill,'font-size':size,'text-anchor':anchor,'font-weight':weight});e.textContent=s;return e}
function rackDefsV53(svg){const d=svgEV53(svg,'defs',{});d.innerHTML='<linearGradient id="r410cab" x1="0" x2="1"><stop offset="0" stop-color="#0b1117"/><stop offset=".5" stop-color="#46515c"/><stop offset="1" stop-color="#10171e"/></linearGradient><linearGradient id="r410dev" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#55687a"/><stop offset=".5" stop-color="#293a49"/><stop offset="1" stop-color="#14202a"/></linearGradient><linearGradient id="r410sw" x1="0" x2="1"><stop offset="0" stop-color="#123b59"/><stop offset=".5" stop-color="#3c7eab"/><stop offset="1" stop-color="#112f48"/></linearGradient><filter id="r410shadow"><feDropShadow dx="4" dy="6" stdDeviation="5" flood-opacity=".28"/></filter>'}
function rackContextV53(){const r=rackModelV53();return{...r,ru:Number(r.s.ru||10),name:r.s.name||r.s.gpu||'GPU System',switchName:r.sw.model||r.sw.name||'Leaf Switch'}}
function drawRackShell2D(svg){const c=RACK_CTX_V53,x=260,y=55,w=370,h=680,innerX=292,innerY=82,innerW=306,innerH=620;svgEV53(svg,'rect',{x,y,width:w,height:h,rx:14,fill:'url(#r410cab)',stroke:'#070b0f','stroke-width':6,filter:'url(#r410shadow)'});svgEV53(svg,'rect',{x:innerX,y:innerY,width:innerW,height:innerH,fill:'#d4dae0',stroke:'#65727e','stroke-width':2});svgEV53(svg,'rect',{x:innerX+6,y:innerY+5,width:10,height:innerH-10,fill:'#27323b'});svgEV53(svg,'rect',{x:innerX+innerW-16,y:innerY+5,width:10,height:innerH-10,fill:'#27323b'});for(let yy=innerY+12;yy<innerY+innerH-10;yy+=11){svgEV53(svg,'rect',{x:innerX+9,y:yy,width:4,height:4,fill:'#8b97a2'});svgEV53(svg,'rect',{x:innerX+innerW-13,y:yy,width:4,height:4,fill:'#8b97a2'})}return{x,y,w,h,innerX:innerX+18,innerY,innerW:innerW-36,innerH,startY:innerY+innerH-10,uH:(innerH-20)/c.rackU}}
function drawRackShell3D(svg,rear=false){const x=rear?330:190,y=130,w=350,h=550,dx=rear?-120:120,dy=-62;svgEV53(svg,'polygon',{points:x+','+y+' '+(x+dx)+','+(y+dy)+' '+(x+w+dx)+','+(y+dy)+' '+(x+w)+','+y,fill:'#6b7782',stroke:'#111820','stroke-width':2,filter:'url(#r410shadow)'});const side=rear?x+','+y+' '+(x+dx)+','+(y+dy)+' '+(x+dx)+','+(y+h+dy)+' '+x+','+(y+h):(x+w)+','+y+' '+(x+w+dx)+','+(y+dy)+' '+(x+w+dx)+','+(y+h+dy)+' '+(x+w)+','+(y+h);svgEV53(svg,'polygon',{points:side,fill:'#26313a',stroke:'#111820','stroke-width':2});svgEV53(svg,'rect',{x,y,width:w,height:h,rx:5,fill:'#d0d6dc',stroke:'#111820','stroke-width':5});return{x,y,w,h,dx,dy,innerX:x+30,innerW:w-60,startY:y+h-18,uH:(h-36)/RACK_CTX_V53.rackU,rear}}
function drawRuScale(svg,g,rackU){for(let u=1;u<=rackU;u++){const yy=g.startY-u*g.uH;svgEV53(svg,'line',{x1:g.innerX,y1:yy,x2:g.innerX+g.innerW,y2:yy,stroke:u%2?'#b3bbc4':'#929da8','stroke-width':.35});if(u%2===0)svgTV53(svg,g.innerX-9,yy+3,'U'+u,7,'#566779','end','600')}}
function drawPduA(svg,g){svgEV53(svg,'rect',{x:g.x+10,y:g.y+40,width:18,height:g.h-80,rx:5,fill:'#4b1717',stroke:'#d35c5c'});for(let yy=g.y+58;yy<g.y+g.h-45;yy+=30)svgEV53(svg,'rect',{x:g.x+14,y:yy,width:10,height:15,rx:2,fill:'#eccccc'});svgTV53(svg,g.x+19,g.y+25,'PDU-A',9,'#e87373','middle','800')}
function drawPduB(svg,g){svgEV53(svg,'rect',{x:g.x+g.w-28,y:g.y+40,width:18,height:g.h-80,rx:5,fill:'#123252',stroke:'#4d99d3'});for(let yy=g.y+58;yy<g.y+g.h-45;yy+=30)svgEV53(svg,'rect',{x:g.x+g.w-24,y:yy,width:10,height:15,rx:2,fill:'#c9e0f2'});svgTV53(svg,g.x+g.w-19,g.y+25,'PDU-B',9,'#72b8ec','middle','800')}
function drawFrontServer(svg,x,y,w,h,label){svgEV53(svg,'rect',{x,y,width:w,height:h,rx:4,fill:'url(#r410dev)',stroke:'#081018','stroke-width':1.3});svgEV53(svg,'rect',{x:x+10,y:y+6,width:15,height:Math.max(8,h-12),rx:3,fill:'#0d151b'});svgEV53(svg,'rect',{x:x+w-25,y:y+6,width:15,height:Math.max(8,h-12),rx:3,fill:'#0d151b'});for(let yy=y+10;yy<y+h-7;yy+=8)for(let xx=x+35;xx<x+135;xx+=9)svgEV53(svg,'circle',{cx:xx,cy:yy,r:1.4,fill:'#8193a4'});for(let k=0;k<5;k++)svgEV53(svg,'rect',{x:x+151+k*24,y:y+9,width:18,height:Math.min(24,h-18),rx:2,fill:'#111a23',stroke:'#71808e','stroke-width':.5});svgTV53(svg,x+w/2,y+16,label,9,'#f5f9fc','middle','800')}
function drawRearServer(svg,x,y,w,h,label){svgEV53(svg,'rect',{x,y,width:w,height:h,rx:4,fill:'#33434f',stroke:'#081018','stroke-width':1.3});for(let k=0;k<5;k++){const cx=x+32+k*31,cy=y+h/2;svgEV53(svg,'circle',{cx,cy,r:11,fill:'#101820',stroke:'#7b8995'});svgEV53(svg,'circle',{cx,cy,r:3.5,fill:'#465662'})}for(let k=0;k<8;k++)svgEV53(svg,'rect',{x:x+178+(k%4)*22,y:y+8+Math.floor(k/4)*15,width:17,height:9,rx:1,fill:'#0d1c27',stroke:'#9bc4dd','stroke-width':.6});for(let k=0;k<3;k++)svgEV53(svg,'rect',{x:x+184+k*31,y:y+h-20,width:25,height:12,rx:2,fill:'#141c23',stroke:'#a1aab0','stroke-width':.5});svgTV53(svg,x+w/2,y+16,label,9,'#f5f9fc','middle','800')}
function drawFrontSwitch(svg,x,y,w,h,label){svgEV53(svg,'rect',{x,y,width:w,height:h,rx:3,fill:'url(#r410sw)',stroke:'#081018','stroke-width':1.2});const rows=h>30?2:1;for(let r=0;r<rows;r++)for(let k=0;k<18;k++)svgEV53(svg,'rect',{x:x+35+k*13,y:y+7+r*13,width:9,height:6,rx:1,fill:k%4===0?'#9fd4ed':'#c4d0da'});svgEV53(svg,'circle',{cx:x+w-16,cy:y+12,r:2.2,fill:'#4cdf88'});svgTV53(svg,x+w/2,y+15,label,8.5,'#fff','middle','800')}
function drawRearSwitch(svg,x,y,w,h,label){svgEV53(svg,'rect',{x,y,width:w,height:h,rx:3,fill:'#2b3944',stroke:'#081018','stroke-width':1.2});for(let k=0;k<5;k++){svgEV53(svg,'rect',{x:x+34+k*43,y:y+6,width:33,height:Math.max(8,h-12),rx:2,fill:'#15212a',stroke:'#7c8994','stroke-width':.5});svgEV53(svg,'circle',{cx:x+44+k*43,cy:y+h/2,r:4,fill:'#3d4c57'})}for(let k=0;k<2;k++)svgEV53(svg,'rect',{x:x+w-58+k*25,y:y+7,width:20,height:Math.max(8,h-14),rx:2,fill:'#121a21',stroke:'#9aa5ae','stroke-width':.6});svgTV53(svg,x+w/2,y+15,label,8.5,'#fff','middle','800')}
function drawPatchPanel(svg,x,y,w,h,label='Fiber Patch Panel'){svgEV53(svg,'rect',{x,y,width:w,height:h,rx:3,fill:'#314450',stroke:'#0a1118'});for(let k=0;k<16;k++)svgEV53(svg,'rect',{x:x+30+k*15,y:y+6,width:9,height:8,rx:1,fill:k%2?'#c0d8e7':'#8db7d0'});svgTV53(svg,x+w/2,y+15,label,8.2,'#fff','middle','800')}
function drawCableManager(svg,x,y,w,h,label='Horizontal Cable Manager'){svgEV53(svg,'rect',{x,y,width:w,height:h,rx:3,fill:'#222f38',stroke:'#0a1118'});for(let k=0;k<12;k++)svgEV53(svg,'line',{x1:x+28+k*20,y1:y+5,x2:x+28+k*20,y2:y+h-5,stroke:'#6f7e89','stroke-width':2});svgTV53(svg,x+w/2,y+15,label,8.2,'#fff','middle','800')}
function drawCallout(svg,x,y,tx,ty,text,anchor='start'){svgEV53(svg,'path',{d:'M '+x+' '+y+' L '+(tx+(anchor==='end'?-7:7))+' '+(ty-3),fill:'none',stroke:'#718498','stroke-width':1.2});svgEV53(svg,'circle',{cx:x,cy:y,r:2.4,fill:'#2c70a7'});svgTV53(svg,tx,ty,text,9,'#34495e',anchor,'700')}
function rackPlanV53(g){const r=window.DCDesign;if(r?.usable&&r.input.scenario!=='colo')return {shown:r.racks.find(x=>x.role==='Compute').units,items:r.rackLayout.components.filter(x=>x.group==='Compute').map(x=>({type:'server',u:x.startU,ru:x.endU-x.startU+1,label:x.label}))};const c=RACK_CTX_V53,reserve=7,shown=Math.min(c.perRack,Math.max(1,Math.floor((c.rackU-reserve)/Math.max(1,c.ru))));let cursor=1,items=[];for(let i=0;i<shown;i++){items.push({type:'server',u:cursor,ru:c.ru,label:'GPU Server '+(i+1)});cursor+=c.ru}if(cursor+1<c.rackU){items.push({type:'manager',u:cursor,ru:1,label:'Horizontal Cable Manager'});cursor++}if(cursor+1<c.rackU){items.push({type:'patch',u:cursor,ru:1,label:'Fiber Patch Panel'});cursor++}const swru=Math.min(2,Number(c.sw.ru||2));if(cursor+swru<c.rackU)items.push({type:'switch',u:cursor,ru:swru,label:'Leaf Switch'});return{items,shown}}
function renderRack2DFront(svg){rackDefsV53(svg);const g=drawRackShell2D(svg),c=RACK_CTX_V53,plan=rackPlanV53(g);drawRuScale(svg,g,c.rackU);drawPduA(svg,g);drawPduB(svg,g);let first=null,sw=null;plan.items.forEach(it=>{const y=g.startY-(it.u+it.ru-1)*g.uH,h=Math.max(19,it.ru*g.uH-2),x=g.innerX,w=g.innerW;if(it.type==='server'){drawFrontServer(svg,x,y,w,h,it.label);if(!first)first={x,y,w,h}}else if(it.type==='switch'){drawFrontSwitch(svg,x,y,w,h,it.label);sw={x,y,w,h}}else if(it.type==='patch')drawPatchPanel(svg,x,y,w,h,it.label);else drawCableManager(svg,x,y,w,h,it.label)});for(let k=0;k<5;k++)svgEV53(svg,'path',{d:'M '+(g.innerX+g.innerW-8)+' '+(190+k*17)+' C 700 '+(190+k*17)+', 710 '+(320+k*20)+', '+(g.x+g.w-28)+' '+(355+k*20),fill:'none',stroke:k<3?'#18afc1':'#805db8','stroke-width':2.6,'stroke-linecap':'round'});if(first){drawCallout(svg,first.x+55,first.y+first.h/2,65,190,'GPU Server front face');drawCallout(svg,first.x+220,first.y+first.h-12,65,222,'drive / vent / I/O / optic area')}if(sw)drawCallout(svg,sw.x+190,sw.y+sw.h/2,705,565,'Leaf Switch · QSFP/OSFP port field');drawCallout(svg,g.x+19,g.y+300,65,310,'PDU-A');drawCallout(svg,g.x+g.w-19,g.y+300,705,310,'PDU-B');svgTV53(svg,445,785,'2D FRONT · RU / rail / bezel / GPU server / patching',11,'#2c4053','middle','800')}
function renderRack2DRear(svg){rackDefsV53(svg);const g=drawRackShell2D(svg),c=RACK_CTX_V53,plan=rackPlanV53(g);drawRuScale(svg,g,c.rackU);drawPduA(svg,g);drawPduB(svg,g);let first=null;plan.items.forEach(it=>{const y=g.startY-(it.u+it.ru-1)*g.uH,h=Math.max(19,it.ru*g.uH-2),x=g.innerX,w=g.innerW;if(it.type==='server'){drawRearServer(svg,x,y,w,h,it.label);if(!first)first={x,y,w,h}}else if(it.type==='switch')drawRearSwitch(svg,x,y,w,h,'Leaf Switch · PSU/Fan side');else if(it.type==='patch')drawPatchPanel(svg,x,y,w,h,'Patch Rear / Trunk Entry');else drawCableManager(svg,x,y,w,h,'Rear Cable Manager')});for(let k=0;k<6;k++)svgEV53(svg,'path',{d:'M '+(g.innerX+g.innerW-5)+' '+(180+k*19)+' C 710 '+(185+k*19)+', 720 '+(350+k*17)+', '+(g.x+g.w-24)+' '+(390+k*21),fill:'none',stroke:k<4?'#16a9ba':'#7655ad','stroke-width':3,'stroke-linecap':'round'});for(let k=0;k<4;k++)svgEV53(svg,'path',{d:'M '+(g.innerX+8)+' '+(300+k*36)+' C 190 '+(315+k*33)+', 175 '+(490+k*24)+', '+(g.x+12)+' '+(525+k*24),fill:'none',stroke:k%2?'#4a91c5':'#c35454','stroke-width':3,'stroke-linecap':'round'});if(first){drawCallout(svg,first.x+85,first.y+first.h/2,65,188,'fan modules');drawCallout(svg,first.x+225,first.y+17,705,195,'rear NIC / network I/O');drawCallout(svg,first.x+225,first.y+first.h-14,705,228,'PSU / power inlet');drawCallout(svg,first.x+270,first.y+30,705,260,'management port')}drawCallout(svg,g.x+g.w-19,g.y+330,705,385,'vertical PDU');svgTV53(svg,445,785,'2D REAR · PSU / fan / NIC / management / cable routing',11,'#2c4053','middle','800')}
function draw3DDeviceV53(svg,g,it,rear){const y=g.startY-(it.u+it.ru-1)*g.uH,h=Math.max(19,it.ru*g.uH-2),x=g.innerX,w=g.innerW,dep=rear?-56:56,ddy=-29,front=it.type==='switch'?'#245f85':(rear?'#33434f':'#2a3d4d');svgEV53(svg,'polygon',{points:x+','+y+' '+(x+dep)+','+(y+ddy)+' '+(x+w+dep)+','+(y+ddy)+' '+(x+w)+','+y,fill:it.type==='switch'?'#5794ba':'#596d7e',stroke:'#111820'});const side=rear?x+','+y+' '+(x+dep)+','+(y+ddy)+' '+(x+dep)+','+(y+h+ddy)+' '+x+','+(y+h):(x+w)+','+y+' '+(x+w+dep)+','+(y+ddy)+' '+(x+w+dep)+','+(y+h+ddy)+' '+(x+w)+','+(y+h);svgEV53(svg,'polygon',{points:side,fill:'#182630',stroke:'#111820'});svgEV53(svg,'rect',{x,y,width:w,height:h,rx:3,fill:front,stroke:'#101820','stroke-width':1.1});if(it.type==='server'&&!rear){for(let yy=y+9;yy<y+h-7;yy+=9)for(let xx=x+32;xx<x+132;xx+=10)svgEV53(svg,'circle',{cx:xx,cy:yy,r:1.4,fill:'#8296a8'});for(let k=0;k<7;k++)svgEV53(svg,'rect',{x:x+158+k*16,y:y+h/2-5,width:11,height:9,rx:1,fill:'#10212d',stroke:'#86a9bd','stroke-width':.45})}else if(it.type==='server'&&rear){for(let k=0;k<5;k++){const cx=x+31+k*31,cy=y+h/2;svgEV53(svg,'circle',{cx,cy,r:10,fill:'#111a22',stroke:'#7b8995'});svgEV53(svg,'circle',{cx,cy,r:3.5,fill:'#43525e'})}for(let k=0;k<8;k++)svgEV53(svg,'rect',{x:x+180+(k%4)*21,y:y+8+Math.floor(k/4)*15,width:16,height:9,rx:1,fill:'#0d1c27',stroke:'#9bc4dd','stroke-width':.6})}else if(it.type==='switch'&&!rear){for(let r=0;r<(h>31?2:1);r++)for(let k=0;k<17;k++)svgEV53(svg,'rect',{x:x+34+k*13,y:y+8+r*13,width:9,height:6,rx:1,fill:k%4===0?'#98cee8':'#c0ccd5'})}else if(it.type==='switch'&&rear){for(let k=0;k<5;k++)svgEV53(svg,'rect',{x:x+35+k*42,y:y+7,width:32,height:Math.max(8,h-14),rx:2,fill:'#16222b',stroke:'#7b8995','stroke-width':.5})}else{for(let k=0;k<11;k++)svgEV53(svg,'rect',{x:x+45+k*19,y:y+7,width:13,height:8,rx:1,fill:'#89a8ba'})}svgTV53(svg,x+w/2,y+15,it.label,8.5,'#fff','middle','800')}
function renderRack3DFront(svg){rackDefsV53(svg);const g=drawRackShell3D(svg,false),plan=rackPlanV53(g);plan.items.forEach(it=>draw3DDeviceV53(svg,g,it,false));drawPduA(svg,g);drawPduB(svg,g);svgEV53(svg,'rect',{x:g.x-18,y:g.y-12,width:g.w+36,height:g.h+24,rx:8,fill:'none',stroke:'#91a9bb','stroke-width':2,'stroke-dasharray':'8 6',opacity:.55});for(let k=0;k<7;k++)svgEV53(svg,'path',{d:'M '+(g.x+g.w-35)+' '+(225+k*17)+' C '+(g.x+g.w+72)+' '+(220+k*17)+', '+(g.x+g.w+92)+' '+(120+k*22)+', '+(g.x+g.w+g.dx-8)+' '+(128+k*22),fill:'none',stroke:k<4?'#19adbd':'#7d5bb7','stroke-width':3.2,'stroke-linecap':'round'});drawCallout(svg,g.x+g.w-28,255,735,210,'front port field / optic cables');drawCallout(svg,g.x-10,340,70,345,'front door / cabinet');svgTV53(svg,455,780,'3D FRONT · equipment depth / patch panel / cable manager / service loop',11,'#2c4053','middle','800')}
function renderRack3DRear(svg){rackDefsV53(svg);const g=drawRackShell3D(svg,true),plan=rackPlanV53(g);plan.items.forEach(it=>draw3DDeviceV53(svg,g,it,true));drawPduA(svg,g);drawPduB(svg,g);svgEV53(svg,'polygon',{points:(g.x+g.dx-120)+','+(g.y+40)+' '+(g.x+g.dx)+','+(g.y+g.dy)+' '+(g.x+g.dx)+','+(g.y+g.h+g.dy)+' '+(g.x+g.dx-120)+','+(g.y+g.h-20),fill:'#7aa6c8','fill-opacity':.08,stroke:'#6c91ad','stroke-dasharray':'8 7'});svgTV53(svg,g.x+g.dx-70,g.y+g.h-5,'REAR SERVICE CLEARANCE',9,'#567188','middle','700');for(let k=0;k<7;k++)svgEV53(svg,'path',{d:'M '+(g.x+32)+' '+(215+k*17)+' C '+(g.x-82)+' '+(225+k*17)+', '+(g.x-106)+' '+(345+k*18)+', '+(g.x+g.dx+8)+' '+(390+k*20),fill:'none',stroke:k<4?'#18a9b9':'#7654ae','stroke-width':3.2,'stroke-linecap':'round'});for(let k=0;k<4;k++)svgEV53(svg,'path',{d:'M '+(g.x+38)+' '+(330+k*33)+' C '+(g.x-54)+' '+(350+k*31)+', '+(g.x-82)+' '+(520+k*20)+', '+(g.x+g.dx+8)+' '+(560+k*14),fill:'none',stroke:k%2?'#488fc5':'#c35353','stroke-width':3,'stroke-linecap':'round'});drawCallout(svg,g.x+65,250,715,215,'PSU / fan / NIC / transceiver zone');drawCallout(svg,g.x+18,460,70,525,'rear fiber bundle / power whip');svgTV53(svg,450,780,'3D REAR · PSU/fan side / rear I/O / service clearance',11,'#2c4053','middle','800')}
function mountRackExplainerV53(){let p=elV53('v48-rack-twin');if(p)return p;const grid=elV53('rackGrid');const power=elV53('v48-power-model');const role=elV53('v48-role-model');const host=grid||power||role||document.querySelector('main,.wrap,.container');if(!host)return null;p=document.createElement('section');p.id='v48-rack-twin';if(grid&&grid.parentNode)grid.parentNode.insertBefore(p,grid);else if(host.insertAdjacentElement)host.insertAdjacentElement('afterend',p);else document.body.appendChild(p);return p}
function renderRackExplainerV53(){const p=mountRackExplainerV53();if(!p)return;RACK_CTX_V53=rackContextV53();p.classList.add('v53-rack-upgraded');const mode=RACK_VIEW_V53.mode,side=RACK_VIEW_V53.side;p.innerHTML='<h3>Rack Digital Twin · 2D / 3D <span class="v48-badge">'+escV53(RACK_CTX_V53.name)+'</span></h3><div class="v48-sub">2D Front/Rear와 3D Front/Rear를 독립적으로 전환할 수 있습니다. 2D는 RU·rail·bezel·port·PDU·patching 중심의 상세 elevation, 3D는 장비 깊이·배선·service clearance 중심의 isometric view입니다. 제조사 CAD/IFC 치수도면이 아닌 conceptual engineering explanation view입니다.</div><div class="rack-view-toolbar"><div class="rack-mode-toggle"><button type="button" data-v53-mode="2d" class="'+(mode==='2d'?'active':'')+'">2D</button><button type="button" data-v53-mode="3d" class="'+(mode==='3d'?'active':'')+'">3D</button></div><div class="rack-side-toggle"><button type="button" data-v53-side="front" class="'+(side==='front'?'active':'')+'">Front</button><button type="button" data-v53-side="rear" class="'+(side==='rear'?'active':'')+'">Rear</button></div></div><div class="rack-view-frame"><div class="v53-rack-cap"><span>'+(mode==='2d'?'2D ENGINEERING ELEVATION':'3D ISOMETRIC / SERVICE VIEW')+'</span><span>'+side.toUpperCase()+'</span></div><svg id="v53-rack-svg" viewBox="0 0 920 820"></svg></div><div class="v53-rack-note">2D Front: RU / rail / bezel / GPU server / drive·vent·handle / Leaf port field / patch / cable manager / A·B PDU · 2D Rear: PSU / fan / rear NIC·I/O / management / power inlet / cabling · 3D: cabinet depth / service routing / clearance</div>';p.querySelectorAll('[data-v53-mode]').forEach(b=>b.addEventListener('click',()=>{RACK_VIEW_V53.mode=b.getAttribute('data-v53-mode');renderRackExplainerV53()}));p.querySelectorAll('[data-v53-side]').forEach(b=>b.addEventListener('click',()=>{RACK_VIEW_V53.side=b.getAttribute('data-v53-side');renderRackExplainerV53()}));const svg=elV53('v53-rack-svg');if(!svg)return;if(mode==='2d'&&side==='front')renderRack2DFront(svg);else if(mode==='2d')renderRack2DRear(svg);else if(side==='front')renderRack3DFront(svg);else renderRack3DRear(svg)}
function rackSignatureV53(){const r=rackModelV53();return[currentLangV53(),r.s.name||r.s.gpu,r.units,r.racks,r.perRack,r.rackU,r.sw.model||r.sw.name,RACK_VIEW_V53.mode,RACK_VIEW_V53.side].join('|')}
function coolingSignatureV53(){const m=coolingModelV53();return[currentLangV53(),m.s.name||m.s.gpu,m.perRack,m.typical,m.design,m.peak].join('|')}
function refreshV53(){try{renderRackExplainerV53()}catch(e){console.error('v5.3 rack view',e)}try{renderCoolingArchitectureV53()}catch(e){console.error('v5.3 cooling architecture',e)}}
if(typeof calc==='function'){const _calcV53=calc;calc=function(){const r=_calcV53.apply(this,arguments);setTimeout(refreshV53,0);return r}}
let lastLangV53='',lastRackSigV53='',lastCoolSigV53='';
function guardedRefreshV53(){const l=currentLangV53(),rs=rackSignatureV53(),cs=coolingSignatureV53();if(l!==lastLangV53||rs!==lastRackSigV53||cs!==lastCoolSigV53){lastLangV53=l;lastRackSigV53=rs;lastCoolSigV53=cs;refreshV53()}}
document.addEventListener('click',e=>{if(e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],.lang,.language'))setTimeout(guardedRefreshV53,50)},true);
document.addEventListener('change',()=>setTimeout(guardedRefreshV53,40),true);
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',()=>{refreshV53();setTimeout(refreshV53,600)},{once:true});else{refreshV53();setTimeout(refreshV53,600)}
setInterval(guardedRefreshV53,1200);
})();