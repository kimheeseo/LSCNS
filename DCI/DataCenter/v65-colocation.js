(function browserPatch(){
'use strict';

const L={
 ko:{
  title:'Data Center Operating Scenario',sub:'AI/GPU 전용 설계와 Colocation / Multi-tenant 설계를 동일 툴에서 분리해 계산합니다.',
  ai:'AI / GPU Cluster',colo:'Colocation / Multi-tenant',scenario:'Operating Scenario',
  tenants:'Tenant 수',racks:'Racks / Tenant',density:'평균 Rack 밀도 (kW)',reserve:'Capacity Reserve (%)',carriers:'Carrier Paths / Tenant',
  totalRacks:'Tenant Racks',itLoad:'Base IT Load',reserved:'Reserved IT Capacity',facility:'Facility Design Power',cross:'Cross-Connect Paths',
  planPower:'Power / Resilience',planPowerText:'A/B 전원 경로와 tenant별 계량을 기본 설계안으로 둡니다. 비어 있는 기존 설정은 2N / N+1 / Dual Fabric / 99.99% reference scenario로 채우되 사용자가 수정할 수 있습니다.',
  planNetwork:'Network / MMR',planNetworkText:'Carrier-diverse MMR/entrance → shared facility core → tenant VRF/VLAN 또는 전용 fabric → tenant ToR/leaf-spine 구조를 권장 설계안으로 표시합니다.',
  planCooling:'Cooling / Density Zones',planCoolingText:'공용 냉각 plant를 사용하되 tenant/rack 밀도별 thermal zone을 분리합니다. AI-ready 고밀도 zone은 liquid-cooling capacity를 별도 검증하도록 표시합니다.',
  planSecurity:'Tenant Isolation / Operations',planSecurityText:'Cage/Suite 물리 분리, 논리 segmentation, tenant별 OOB/계량, SLA 및 변경관리(MOP/SOP/EOP)를 설계 체크 항목으로 포함합니다.',
  note:'Colocation 값은 reference scenario이며 인증/가용성 보증값이 아닙니다. PUE는 현재 System Engineering의 editable PUE 값을 사용합니다. 기존 AI/GPU BOM은 colocation tenant workload의 한 모듈로 유지됩니다.',
  aiText:'현재 AI/GPU Cluster BOM 계산을 그대로 사용합니다. Colocation을 선택하면 tenant/facility/resilience 계층을 추가로 계산합니다.'
 },
 en:{
  title:'Data Center Operating Scenario',sub:'Switch between AI/GPU-cluster design and Colocation / Multi-tenant facility design.',
  ai:'AI / GPU Cluster',colo:'Colocation / Multi-tenant',scenario:'Operating Scenario',
  tenants:'Tenant count',racks:'Racks / Tenant',density:'Average Rack Density (kW)',reserve:'Capacity Reserve (%)',carriers:'Carrier Paths / Tenant',
  totalRacks:'Tenant Racks',itLoad:'Base IT Load',reserved:'Reserved IT Capacity',facility:'Facility Design Power',cross:'Cross-Connect Paths',
  planPower:'Power / Resilience',planPowerText:'Use A/B rack power paths and per-tenant metering. Empty existing settings are initialized to a reference 2N / N+1 / Dual Fabric / 99.99% scenario and remain editable.',
  planNetwork:'Network / MMR',planNetworkText:'Carrier-diverse MMR/entrance → shared facility core → tenant VRF/VLAN or dedicated fabric → tenant ToR/leaf-spine.',
  planCooling:'Cooling / Density Zones',planCoolingText:'Use a shared cooling plant with density-based thermal zones. AI-ready high-density zones require separate liquid-cooling capacity verification.',
  planSecurity:'Tenant Isolation / Operations',planSecurityText:'Include cage/suite separation, logical segmentation, tenant OOB/metering, SLA and change-control MOP/SOP/EOP checks.',
  note:'Colocation values are an editable reference scenario, not a certification or availability guarantee. PUE uses the current editable System Engineering value. The existing AI/GPU BOM remains available as one tenant workload module.',
  aiText:'The current AI/GPU Cluster BOM remains active. Selecting Colocation adds tenant, facility and resilience design layers.'
 },
 zh:{
  title:'Data Center Operating Scenario',sub:'在 AI/GPU 集群设计与 Colocation / Multi-tenant 设施设计之间切换。',
  ai:'AI / GPU Cluster',colo:'Colocation / Multi-tenant',scenario:'Operating Scenario',
  tenants:'租户数',racks:'每租户机架数',density:'平均机架功率 (kW)',reserve:'容量预留 (%)',carriers:'每租户 Carrier 路径',
  totalRacks:'Tenant Racks',itLoad:'Base IT Load',reserved:'Reserved IT Capacity',facility:'Facility Design Power',cross:'Cross-Connect Paths',
  planPower:'Power / Resilience',planPowerText:'采用 A/B 供电路径和租户独立计量；空白设置按参考 2N / N+1 / Dual Fabric / 99.99% 初始化并保持可编辑。',
  planNetwork:'Network / MMR',planNetworkText:'Carrier-diverse MMR/entrance → shared core → tenant VRF/VLAN 或独立 fabric → tenant ToR/leaf-spine。',
  planCooling:'Cooling / Density Zones',planCoolingText:'共享冷却系统并按机架密度划分 thermal zone；AI-ready 高密度区域需单独验证液冷容量。',
  planSecurity:'Tenant Isolation / Operations',planSecurityText:'包括 cage/suite 隔离、逻辑分段、租户 OOB/计量、SLA 与 MOP/SOP/EOP 变更管理。',
  note:'Colocation 数值为可编辑参考场景，并非认证或可用性保证。PUE 使用当前 System Engineering 可编辑值。',
  aiText:'保持当前 AI/GPU Cluster BOM。选择 Colocation 后增加租户、设施与冗余设计层。'
 },
 ja:{
  title:'Data Center Operating Scenario',sub:'AI/GPU クラスタ設計と Colocation / Multi-tenant 設計を切り替えます。',
  ai:'AI / GPU Cluster',colo:'Colocation / Multi-tenant',scenario:'Operating Scenario',
  tenants:'Tenant 数',racks:'Racks / Tenant',density:'平均 Rack 密度 (kW)',reserve:'Capacity Reserve (%)',carriers:'Carrier Paths / Tenant',
  totalRacks:'Tenant Racks',itLoad:'Base IT Load',reserved:'Reserved IT Capacity',facility:'Facility Design Power',cross:'Cross-Connect Paths',
  planPower:'Power / Resilience',planPowerText:'A/B 電源経路と tenant 別メータリングを基本とし、空欄設定は reference 2N / N+1 / Dual Fabric / 99.99% で初期化し編集可能です。',
  planNetwork:'Network / MMR',planNetworkText:'Carrier-diverse MMR/entrance → shared core → tenant VRF/VLAN または dedicated fabric → tenant ToR/leaf-spine。',
  planCooling:'Cooling / Density Zones',planCoolingText:'共有 cooling plant を使い密度別 thermal zone を分離。AI-ready 高密度 zone は液冷容量を別途検証します。',
  planSecurity:'Tenant Isolation / Operations',planSecurityText:'Cage/Suite 分離、論理 segmentation、tenant OOB/計量、SLA、MOP/SOP/EOP をチェック項目に含めます。',
  note:'Colocation 値は編集可能な reference scenario であり認証値ではありません。PUE は現在の System Engineering 値を使用します。',
  aiText:'現在の AI/GPU Cluster BOM を維持し、Colocation 選択時に tenant/facility/resilience レイヤーを追加します。'
 },
 de:{
  title:'Data Center Operating Scenario',sub:'Umschalten zwischen AI/GPU-Cluster und Colocation / Multi-tenant Design.',
  ai:'AI / GPU Cluster',colo:'Colocation / Multi-tenant',scenario:'Operating Scenario',
  tenants:'Tenant-Anzahl',racks:'Racks / Tenant',density:'Mittlere Rack-Dichte (kW)',reserve:'Kapazitätsreserve (%)',carriers:'Carrier-Pfade / Tenant',
  totalRacks:'Tenant Racks',itLoad:'Base IT Load',reserved:'Reserved IT Capacity',facility:'Facility Design Power',cross:'Cross-Connect Paths',
  planPower:'Power / Resilience',planPowerText:'A/B-Strompfade und tenant-spezifische Messung; leere Felder werden als editierbares 2N / N+1 / Dual Fabric / 99.99%-Referenzszenario vorbelegt.',
  planNetwork:'Network / MMR',planNetworkText:'Carrier-diverse MMR/Entrance → Shared Core → Tenant VRF/VLAN oder Dedicated Fabric → Tenant ToR/Leaf-Spine.',
  planCooling:'Cooling / Density Zones',planCoolingText:'Gemeinsame Kühlung mit dichteabhängigen Thermal Zones; AI-ready High-Density-Zonen benötigen separate Liquid-Cooling-Prüfung.',
  planSecurity:'Tenant Isolation / Operations',planSecurityText:'Cage/Suite-Trennung, logische Segmentierung, Tenant-OOB/Metering, SLA und MOP/SOP/EOP-Prüfungen.',
  note:'Colocation-Werte sind editierbare Referenzannahmen, keine Zertifizierung oder Verfügbarkeitsgarantie. PUE nutzt den aktuellen System-Engineering-Wert.',
  aiText:'Das aktuelle AI/GPU-Cluster-BOM bleibt aktiv; Colocation ergänzt Tenant-, Facility- und Resilience-Layer.'
 }
};
function lang(){
 const raw=((window.__dcBomUiLang||document.documentElement.getAttribute('data-dc-bom-ui-lang')||document.documentElement.lang||'ko')+'').toLowerCase();
 if(raw.includes('zh')||raw.includes('cn'))return'zh';if(raw.includes('ja')||raw.includes('jp'))return'ja';if(raw.includes('de'))return'de';if(raw.includes('en'))return'en';return'ko';
}
function T(){return L[lang()]||L.ko}
function root(){return document.querySelector('main')||document.querySelector('.wrap')||document.querySelector('.container')||document.body}
function field(label,id,type,value,min,step){
 return '<div class="v65-field field"><label for="'+id+'">'+label+'</label><input id="'+id+'" type="'+type+'" value="'+value+'"'+(min!=null?' min="'+min+'"':'')+(step!=null?' step="'+step+'"':'')+'></div>';
}
function metric(k,v,s){return '<div class="metric v65-metric"><div class="k">'+k+'</div><div class="v">'+v+'</div><div class="s">'+s+'</div></div>'}
function num(id,d){const e=document.getElementById(id),v=e?Number(e.value):NaN;return Number.isFinite(v)?v:d}
function pue(){const a=document.getElementById('pue');if(a)return Number(a.value)||1.2;const e=document.getElementById('v60-pue'),v=e?Number(e.value):NaN;return Number.isFinite(v)&&v>=1?v:1.2}
function scenario(){const e=document.getElementById('v65-mode');return e?e.value:'ai'}
function setIfBlank(id,value){
 const e=document.getElementById(id);if(!e||e.value)return;
 const ok=Array.from(e.options||[]).some(o=>o.value===value||o.textContent.trim()===value);if(!ok)return;
 e.value=value;e.dispatchEvent(new Event('change',{bubbles:true}));
}
function applyColoDefaults(){
 setIfBlank('v60-power-red','2N');
 setIfBlank('v60-cooling-red','N+1');
 setIfBlank('v60-network-red','Dual Fabric');
 setIfBlank('v60-avail-target','99.99%');
}
function snapshot(data){
 window.__dcBomScenario=data;
 try{
  window.currentDesignSnapshot=window.currentDesignSnapshot||{};
  window.currentDesignSnapshot.results=window.currentDesignSnapshot.results||{};
  window.currentDesignSnapshot.referenceScenario=data;
 }catch(_){}
}
function render(){
 const panel=document.getElementById('v65-scenario');if(!panel)return;
 const t=T(),mode=scenario(),colo=Array.from(panel.querySelectorAll('.v65-colo')),ai=panel.querySelector('.v65-ai-note');
 panel.querySelector('[data-title]').textContent=t.title;panel.querySelector('[data-sub]').textContent=t.sub;
 const modeLabel=panel.querySelector('label[for="v65-mode"]');if(modeLabel)modeLabel.textContent=t.scenario;
 const sel=document.getElementById('v65-mode');if(sel){sel.options[0].textContent=t.ai;sel.options[1].textContent=t.colo}
 [['v65-tenants','tenants'],['v65-racks','racks'],['v65-density','density'],['v65-reserve','reserve'],['v65-carriers','carriers']].forEach(x=>{const e=panel.querySelector('label[for="'+x[0]+'"]');if(e)e.textContent=t[x[1]]});
 if(mode==='colo'){
   colo.forEach(x=>x.hidden=false);ai.hidden=true;applyColoDefaults();
   const tenants=Math.max(1,Math.round(num('v65-tenants',4))),racks=Math.max(1,Math.round(num('v65-racks',20))),density=Math.max(1,num('v65-density',12)),reserve=Math.max(0,num('v65-reserve',20)),carriers=Math.max(1,Math.round(num('v65-carriers',2)));
   const totalRacks=tenants*racks,base=totalRacks*density,reserved=base*(1+reserve/100),facility=reserved*pue(),cross=tenants*carriers;
   document.getElementById('v65-metrics').innerHTML=
     metric(t.totalRacks,totalRacks.toLocaleString(),tenants+' tenants × '+racks+' racks')+
     metric(t.itLoad,base.toFixed(1)+' kW',totalRacks+' racks × '+density.toFixed(1)+' kW')+
     metric(t.reserved,reserved.toFixed(1)+' kW','+'+reserve.toFixed(0)+'% reserve')+
     metric(t.facility,facility.toFixed(1)+' kW','IT × PUE '+pue().toFixed(2))+
     metric(t.cross,cross.toLocaleString(),carriers+' carrier paths / tenant');
   document.getElementById('v65-plan').innerHTML=
     '<div class="v65-plan-card"><b>'+t.planPower+'</b><span>'+t.planPowerText+'</span></div>'+
     '<div class="v65-plan-card"><b>'+t.planNetwork+'</b><span>'+t.planNetworkText+'</span></div>'+
     '<div class="v65-plan-card"><b>'+t.planCooling+'</b><span>'+t.planCoolingText+'</span></div>'+
     '<div class="v65-plan-card"><b>'+t.planSecurity+'</b><span>'+t.planSecurityText+'</span></div>';
   document.getElementById('v65-note').textContent=t.note;
   snapshot({mode:'colocation-multitenant',tenantCount:tenants,racksPerTenant:racks,totalRacks,avgRackKw:density,reservePct:reserve,baseItKw:base,reservedItKw:reserved,pue:pue(),facilityDesignKw:facility,carrierPathsPerTenant:carriers,crossConnectPaths:cross,referenceDefaults:{power:'2N',cooling:'N+1',network:'Dual Fabric',availability:'99.99%'}});
 }else{
   colo.forEach(x=>x.hidden=true);ai.hidden=false;ai.textContent=t.aiText;snapshot({mode:'ai-gpu-cluster'});
 }
}
function mount(){
 let p=document.getElementById('v65-scenario');if(p)return p;
 const r=root();if(!r)return null;
 p=document.createElement('section');p.id='v65-scenario';
 const t=T();
 p.innerHTML='<h3 data-title>'+t.title+'</h3><div class="v65-sub" data-sub>'+t.sub+'</div>'+
 '<div class="v65-controls"><div class="v65-field field"><label for="v65-mode">'+t.scenario+'</label><select id="v65-mode"><option value="ai">'+t.ai+'</option><option value="colo">'+t.colo+'</option></select></div></div>'+
 '<div class="v65-controls v65-colo" hidden style="margin-top:8px">'+
 field(t.tenants,'v65-tenants','number',4,1,1)+field(t.racks,'v65-racks','number',20,1,1)+field(t.density,'v65-density','number',12,1,0.5)+field(t.reserve,'v65-reserve','number',20,0,1)+
 '<div class="v65-field field"><label for="v65-carriers">'+t.carriers+'</label><input id="v65-carriers" type="number" value="2" min="1" step="1"></div></div>'+
 '<div class="v65-ai-note v65-sub"></div><div class="v65-colo" hidden><div class="v65-grid" id="v65-metrics"></div><div class="v65-plan" id="v65-plan"></div><div class="v65-warn" id="v65-note"></div></div>';
 const anchor=document.getElementById('v48-role-model')||document.getElementById('storageFabricType')?.closest('.card')||r.querySelector('.card');
 if(anchor&&anchor.parentNode)anchor.parentNode.insertBefore(p,anchor);else r.insertBefore(p,r.firstChild);
 p.querySelectorAll('input,select').forEach(e=>{e.addEventListener('input',render);e.addEventListener('change',render)});
 return p;
}
function moveChannelsBottom(){
 const p=document.getElementById('v47-channels'),r=root();if(!p||!r)return;
 if(p.parentNode!==r||r.lastElementChild!==p)r.appendChild(p);
}
function start(){
 mount();render();moveChannelsBottom();
 document.addEventListener('change',()=>setTimeout(()=>{render();moveChannelsBottom()},80),true);
 document.addEventListener('click',e=>{if(e.target&&e.target.closest&&e.target.closest('[data-lang],[data-language],.lang,.language'))setTimeout(render,100)},true);
 setInterval(()=>{mount();render();moveChannelsBottom()},1500);
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})();