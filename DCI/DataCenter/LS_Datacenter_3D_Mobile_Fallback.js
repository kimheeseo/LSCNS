/* LS Datacenter 3D · mobile-first controls + WebGL-independent 2D schematic, 2026-10-26.
 * No external dependencies, GPU context or extra animation loop. Relies on existing
 * __LS3D_TEST__ authoritative zone, asset selection and metrics APIs. */
(function(){
'use strict';
const T=window.__LS3D_TEST__;
if(!T||document.getElementById('ls3dPortable2D'))return;
const $=id=>document.getElementById(id);
const portable=()=>window.matchMedia('(max-width:790px)').matches||document.body.classList.contains('display-mobile');
const NS='http://www.w3.org/2000/svg';
const svgEl=(name,attributes={},text)=>{const el=document.createElementNS(NS,name);for(const [k,v] of Object.entries(attributes))el.setAttribute(k,String(v));if(text!==undefined)el.textContent=text;return el};
const style=document.createElement('style');
style.textContent=`
/* Default mobile mode is 3D canvas first; all secondary inspection panels are opt-in. */
#ls3dMobileInspectorToggle,#ls3dScenarioSummaryToggle{display:none}
#ls3dPortable2D[hidden]{display:none!important}
#ls3dPortable2D{position:absolute;inset:0;z-index:14;background:#0a1b2b;color:#e9f8ff;display:flex;flex-direction:column;gap:8px;padding:10px;overflow:hidden}
#ls3dPortable2D .lite-header{display:flex;align-items:center;justify-content:space-between;gap:6px}
#ls3dPortable2D .lite-header strong{font-size:14px}
#ls3dPortable2D .lite-header small{display:block;font-size:10px;color:#a6c5d4}
#ls3dPortable2D button,#ls3dPortable2D select{min-height:38px;border:1px solid #4a7187;background:#16364a;border-radius:8px;color:#f1fcff;padding:7px 9px;font:650 11px system-ui;max-width:100%}
#ls3dPortable2D button{cursor:pointer}
#ls3dPortable2D .lite-controls{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.35fr);gap:6px}
#ls3dPortable2D .lite-controls[hidden]{display:none}
#ls3dPortable2D label{display:flex;flex-direction:column;gap:3px;font-size:10px;color:#a8c5d5;min-width:0}
#ls3dPortable2D .lite-map{min-height:100px;flex:1;overflow:hidden;border:1px solid #365c74;border-radius:9px;background:linear-gradient(140deg,#123348,#0b1d2d)}
#ls3dPortable2D svg{display:block;width:100%;height:100%;touch-action:manipulation}
#ls3dPortable2D svg [role=button]{cursor:pointer;outline:none}
#ls3dPortable2D svg [role=button]:focus-visible rect{stroke:#fff3b5;stroke-width:3}
#ls3dPortable2D .lite-selected{border-top:1px solid #36536c;padding-top:6px;display:grid;gap:2px;min-height:48px}
#ls3dPortable2D .lite-selected strong{font-size:12px;color:#9af2da}
#ls3dPortable2D .lite-selected span{font-size:11px;color:#b6d1dc;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#ls3dPortable2D .lite-footer{display:flex;justify-content:space-between;align-items:center;gap:5px;font-size:10px;color:#a9c1cf}
#ls3dPortable2D .lite-footer a{color:#91d7ef;text-decoration:underline}
#ls3dPortable2D .lite-notice{margin:0;font-size:10px;color:#b8d2dc}
body.ls3d-2d-active #scene,body.ls3d-2d-active #sceneThree,body.ls3d-2d-active #markers,
body.ls3d-2d-active #viewport .vp-top,body.ls3d-2d-active #viewport .vp-bottom,
body.ls3d-2d-active #viewport .worker-follow-hud,body.ls3d-2d-active #viewport #error,
body.ls3d-2d-active #viewport .ux-fallback{display:none!important}
@media(max-width:790px){
  #telemetryDrawer.is-closed{left:auto!important;right:8px!important;width:154px!important;max-width:calc(100vw - 16px)!important}
  #telemetryDrawer.is-closed #telemetryToggle{padding:7px 9px;min-height:35px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  #telemetryDrawer.is-closed #telemetryStatus{display:none}
  #ls3dMobileInspectorToggle,#ls3dScenarioSummaryToggle{display:block;cursor:pointer;width:100%;padding:11px 12px;background:#17354a;border:1px solid #426e83;border-radius:8px;color:#ecfaff;font:750 12px system-ui;text-align:left;min-height:42px}
  #ls3dMobileInspectorContent{display:none}
  .right.ls3d-inspector-open #ls3dMobileInspectorContent{display:block}
  .right .right-inner{padding:9px!important}
  .right .right-head{margin-bottom:5px}
  .right .right-head strong{font-size:12px}
  #ls3dScenarioSummaryToggle{margin-bottom:0}
  #scenarioStepPanel.ls3d-scenario-collapsed .scenario-step-head,
  #scenarioStepPanel.ls3d-scenario-collapsed .scenario-step-rail{display:none!important}
  #scenarioStepPanel{padding:4px 6px!important}
  #viewport{height:clamp(340px,60dvh,630px)!important;min-height:0!important}
  .viewport .marker:not(.active):not(.rack-marker){display:none!important}
  #viewport .vp-bottom .legend{display:none}
  #mobileInspectorTools:not([open]){padding:0 8px!important}
  #mobileInspectorTools summary{min-height:38px;padding:10px 0}
  .phase58-panel:not([hidden]){inset:auto 6px max(7px,env(safe-area-inset-bottom)) 6px!important;width:auto!important;max-height:min(62dvh,570px)!important;min-height:0!important;border-radius:15px!important;z-index:98!important;overscroll-behavior:contain!important}
  .detail-modal.show{z-index:97!important;align-items:flex-end!important;justify-content:center!important;padding:5px 5px max(5px,env(safe-area-inset-bottom))!important;background:#030e17a8!important;pointer-events:auto!important}
  .detail-modal.show .detail-card{width:100%!important;max-height:65dvh!important;border-radius:14px!important}
  .telemetry{max-height:55dvh;overflow:auto}
  .top-actions{overflow-x:auto}
}
@media(max-width:420px){
  #ls3dPortable2D{padding:7px}
  #ls3dPortable2D .lite-header strong{font-size:12px}
  #ls3dPortable2D .lite-header small{font-size:9px}
  #ls3dPortable2D button,#ls3dPortable2D select{font-size:10px}
}
`;
document.head.appendChild(style);
const right=$('viewport')?.closest('#shell')?.querySelector('.right')||document.querySelector('.right');
const inner=right?.querySelector('.right-inner'),head=inner?.querySelector('.right-head');
if(inner&&head){
  const content=document.createElement('div');content.id='ls3dMobileInspectorContent';
  while(head.nextSibling)content.appendChild(head.nextSibling);
  inner.appendChild(content);
  const button=document.createElement('button');button.id='ls3dMobileInspectorToggle';button.type='button';
  button.textContent='▣ 설비 검사·상세 펼치기';button.setAttribute('aria-controls',content.id);
  button.setAttribute('aria-expanded','false');head.insertAdjacentElement('afterend',button);
  const setExpanded=value=>{right.classList.toggle('ls3d-inspector-open',value);button.textContent=value?'▣ 설비 검사·상세 접기':'▣ 설비 검사·상세 펼치기';button.setAttribute('aria-expanded',String(value));};
  button.addEventListener('click',()=>setExpanded(!right.classList.contains('ls3d-inspector-open')));
  setExpanded(false);
}
const scenario=$('scenarioStepPanel');
if(scenario){
  const expand=document.createElement('button');expand.type='button';expand.id='ls3dScenarioSummaryToggle';expand.setAttribute('aria-controls','scenarioStepRail scenarioStepMetrics');
  expand.textContent='▤ 시뮬레이션 단계 보기 ▾';
  scenario.insertBefore(expand,scenario.firstChild);
  const setExpanded=v=>{scenario.classList.toggle('ls3d-scenario-collapsed',!v);expand.setAttribute('aria-expanded',String(v));expand.textContent=v?'▤ 시뮬레이션 단계 접기 ▴':'▤ 시뮬레이션 단계 보기 ▾';};
  expand.onclick=()=>setExpanded(scenario.classList.contains('ls3d-scenario-collapsed'));
  setExpanded(false);
}
const view=$('viewport'),layer=document.createElement('section');
layer.id='ls3dPortable2D';layer.setAttribute('aria-label','경량 2D 데이터센터 구역 및 랙 배치');layer.hidden=true;
layer.innerHTML='<div class="lite-header"><div><strong>▤ 2D 설비 배치도</strong><small id="ls3dLiteReason">3D를 대신하는 경량 도면</small></div><button type="button" id="ls3dLiteReturn">3D 돌아가기</button></div><div class="lite-controls"><label>구역 선택<select id="ls3dLiteZone"></select></label><label>설비·랙 선택<select id="ls3dLiteAsset"></select></label></div><div class="lite-map"><svg id="ls3dLiteSvg" role="img" aria-label="설비 개념 평면도" viewBox="0 0 760 470" preserveAspectRatio="xMidYMid meet"></svg></div><div class="lite-selected"><strong id="ls3dLiteSelected">구역 또는 설비를 선택하세요</strong><span id="ls3dLiteSelectedDesc">선택하면 기존 설비 검사와 정보가 연동됩니다.</span></div><div class="lite-footer"><span>개념 좌표 · BIM 실측 아님</span><a href="./LS_Datacenter_Campus.html">기존 2D Tour 보기</a></div>';
view?.appendChild(layer);
const zonePicker=$('ls3dLiteZone'),assetPicker=$('ls3dLiteAsset'),map=$('ls3dLiteSvg');
let active=false,reason='',lastSignature='',lastSynced=0,unsupported=false,selectedLitePart=null;
const colors={gpu:'#4d92c8',storage:'#4baa8c',network:'#d2a353',client:'#a690d3',other:'#6a9cad'};
function colorOf(a){
 if(a.type==='GPU server rack')return colors.gpu;
 if(a.type==='Storage rack')return colors.storage;
 if(a.type==='Top-of-Rack switch')return colors.network;
 if(a.type==='Co-location rack'||a.type==='Customer cage')return colors.client;
 return colors.other;
}
function shortName(a){return a.name.replace(/^GPU Rack /,'GPU ').replace(/^Storage Rack /,'STG ').replace(/^Network Rack /,'NET ').replace(/^Client Rack /,'Client ').replace(/ · ToR/,'').slice(0,18);}
function option(value,label){const op=document.createElement('option');op.value=value;op.textContent=label;return op;}
function rackParts(a){
 const gpu=a.type==='GPU server rack',storage=a.type==='Storage rack';
 return [
  {id:'tor',label:'ToR 스위치 · Top of Rack',description:'최상단 42U에 설치된 스위치로 서버 NIC와 상위 Leaf/Spine 망을 연결합니다.',uStart:42},
  {id:'manager',label:'케이블 매니저',description:'광·UTP·DAC 패치 코드를 정리하고 굴곡 반경과 작업 공간을 확보합니다.',uStart:38},
  {id:'module-a',label:gpu?'GPU 서버 트레이 · GPU/NIC':storage?'스토리지 모듈 · 드라이브':'네트워크 모듈 · 스위치',description:gpu?'GPU 카드와 NIC, 하우징이 장착된 서버 트레이입니다.':storage?'드라이브와 컨트롤러가 장착된 저장 모듈입니다.':'스위치 또는 라우터 모듈이 장착된 랙 구성품입니다.',uStart:30},
  {id:'module-b',label:gpu?'GPU 서버 트레이 · GPU/NIC':storage?'스토리지 모듈 · 드라이브':'라우터 / 광모듈',description:gpu?'두 번째 GPU 서버 트레이와 네트워크 모듈입니다.':storage?'두 번째 저장 모듈과 백업 연결입니다.':'상위 네트워크에 연결되는 라우터와 광모듈입니다.',uStart:19},
  {id:'dac',label:'DAC / 광 패치 케이블',description:'서버 NIC에서 최상단 ToR 스위치로 연결되는 짧은 DAC와 광 패치 코드입니다.',uStart:12},
  {id:'pdu',label:'랙 PDU · A/B 전원',description:'양쪽 전원 분배 장치가 서버에 이중 전원을 공급합니다.',uStart:3},
  {id:'housing',label:'42U 하우징 · 문/레일',description:'랙 프레임, 장착 레일과 통풍 도어를 포함한 외함입니다.',uStart:1}
 ].map(p=>({...p,id:a.id+'-lite-'+p.id,assetId:a.id,zoneId:a.zone,kind:'rack-part'}));
}
function drawRackInterior(a){
 const parts=rackParts(a),named=id=>parts.find(p=>p.id.endsWith('-'+id));
 map.replaceChildren();
 map.append(svgEl('rect',{x:25,y:18,width:710,height:428,rx:10,fill:'#11283c',stroke:'#43768d'}));
 map.append(svgEl('text',{x:380,y:48,'text-anchor':'middle',fill:'#dcf8ff','font-size':19,'font-weight':800},a.name+' · 42U 내부'));
 map.append(svgEl('rect',{x:220,y:66,width:320,height:348,rx:8,fill:'#172d41',stroke:'#9ac7d5','stroke-width':8}));
 for(const [id,y,h,color] of [['tor',78,42,'#116b71'],['manager',126,27,'#436071'],['module-a',162,80,'#2a607b'],['module-b',255,80,'#2a607b'],['dac',344,25,'#6b657f'],['pdu',377,27,'#7e6244']]){
  const p=named(id),g=svgEl('g',{role:'button',tabindex:0,'data-part':p.id,'aria-label':p.label});
  g.append(svgEl('rect',{x:255,y,width:250,height:h,rx:5,fill:color,stroke:selectedLitePart?.id===p.id?'#ffe5a0':'#8ab6c6','stroke-width':selectedLitePart?.id===p.id?3:1.5}));
  g.append(svgEl('text',{x:380,y:y+h/2+5,'text-anchor':'middle',fill:'#e7fbff','font-size':id==='tor'?15:13,'font-weight':800,'pointer-events':'none'},p.label));
  map.append(g);
 }
 for(const x of [236,520])map.append(svgEl('rect',{x,y:80,width:9,height:322,rx:3,fill:'#d2ad65'}));
 map.append(svgEl('path',{d:'M 493 355 C 586 309 585 135 498 101',fill:'none',stroke:'#72c0f8','stroke-width':6,'stroke-dasharray':'8 5'}));
 map.append(svgEl('text',{x:620,y:216,'text-anchor':'middle',fill:'#c6e9f7','font-size':12},'DAC / 광'));
 map.append(svgEl('text',{x:620,y:235,'text-anchor':'middle',fill:'#c6e9f7','font-size':12},'NIC → ToR'));
 map.setAttribute('aria-label',a.name+' 내부 구조 · 부품을 선택하면 설명 표시');
 const p=selectedLitePart;
 $('ls3dLiteSelected').textContent=p?.label||a.name+' · 내부 구조';
 $('ls3dLiteSelectedDesc').textContent=p?.description||'상단 ToR, GPU/저장 모듈, 케이블 매니저, DAC, A/B 전원을 선택해 확인하세요.';
 if(!window.LS3D_THREE?.adapter)window.LS3D_INSPECTOR?.setParts(parts);
}
function draw(){
 if(!active||!map)return;
 layer.querySelector('.lite-controls').hidden=!!T.rackIsolateActive;
 if(T.rackIsolateActive&&T.selected&&T.isMechanicalRack(T.selected)){drawRackInterior(T.selected);return}
 const z=T.Z.find(x=>x.id===T.currentZone)||T.Z[0],home=z.id==='campus';
 const border=home?T.Z[0]:z;
 const minX=border.pos[0]-border.size[0]/2,maxX=border.pos[0]+border.size[0]/2,minZ=border.pos[1]-border.size[1]/2,maxZ=border.pos[1]+border.size[1]/2;
 const x=world=>32+(world-minX)/(maxX-minX)*696;
 const y=world=>25+(world-minZ)/(maxZ-minZ)*412;
 const width=w=>w/(maxX-minX)*696,height=h=>h/(maxZ-minZ)*412;
 map.replaceChildren();
 map.appendChild(svgEl('rect',{x:25,y:18,width:710,height:428,rx:8,fill:'#102a40',stroke:'#37647a'}));
 const grid=svgEl('g',{'aria-hidden':'true',stroke:'#284257','stroke-width':'1'});
 for(let i=1;i<12;i++){let gx=32+i*696/12;grid.append(svgEl('line',{x1:gx,y1:25,x2:gx,y2:437}));}
 for(let i=1;i<8;i++){let gy=25+i*412/8;grid.append(svgEl('line',{x1:32,y1:gy,x2:728,y2:gy}));}
 map.appendChild(grid);
 const list=home?T.visibleZones().filter(v=>v.id!=='campus'):T.visibleAssets().filter(a=>a.zone===z.id);
 // Draw large footprints first, so smaller racks remain accessible over customer cages.
 list.sort((a,b)=>(b.w*b.d)-(a.w*a.d));
 for(const a of list){
   const isZone=home,selected=home?false:T.selected?.id===a.id;
   const rx=x(a.x??a.pos[0]),ry=y(a.z??a.pos[1]);
   const w=Math.max(isZone?46:14,width(a.w??a.size[0])),h=Math.max(isZone?28:12,height(a.d??a.size[1]));
   const g=svgEl('g',{role:'button',tabindex:0,'aria-label':isZone?a.n+' 구역으로 이동':a.name,'data-target':isZone?a.id:a.id,'data-type':isZone?'zone':'asset'});
   g.append(svgEl('rect',{x:rx-w/2,y:ry-h/2,width:w,height:h,rx:Math.min(5,w/7),fill:isZone?a.color:colorOf(a),'fill-opacity':isZone?.55:(selected?.85:.51),stroke:selected?'#fff2b0':'#a9d4e2','stroke-width':selected?3:1.1}));
   // Keep labels readable; the selector lists all assets even when a footprint is tiny.
   if(isZone||selected||((a.type.includes('rack')||a.type==='Top-of-Rack switch')&&w>36)){
     const label=isZone?a.n:(w<85?shortName(a).slice(0,11):shortName(a));
     g.append(svgEl('text',{x:rx,y:ry+3,'text-anchor':'middle',fill:'#effbff','font-size':isZone?12:10,'font-weight':selected?800:650,'pointer-events':'none'},label));
   }
   g.append(svgEl('title',{},isZone?a.n:a.name));
   map.appendChild(g);
 }
 $('ls3dLiteSelected').textContent=T.selected?.name||'선택된 설비 없음';
 $('ls3dLiteSelectedDesc').textContent=(T.selected?.type||'')+' · '+(T.selected?.desc||'설비를 선택하면 정보를 표시합니다.');
 map.setAttribute('aria-label',home?'데이터센터 전체 구역 배치도':z.n+' 설비 배치도');
}
function sync(force=false){
 if(!active)return;
 const now=performance.now();
 if(!force&&now-lastSynced<240)return;
 lastSynced=now;
 if(!T.rackIsolateActive||selectedLitePart?.assetId!==T.selected?.id)selectedLitePart=null;
 const zones=T.visibleZones(),zone=T.currentZone;
 const assets=T.visibleAssets().filter(a=>zone==='campus'||a.zone===zone);
 const signature=zone+'|'+T.facilityMode+'|'+(T.selected?.id||'')+'|'+T.rackIsolateActive+'|'+(selectedLitePart?.id||'')+'|'+assets.map(a=>a.id).join(',');
 if(!force&&signature===lastSignature)return;
 lastSignature=signature;
 const oldZone=zonePicker.value;
 zonePicker.replaceChildren(...zones.map(z=>option(z.id,z.n)));
 zonePicker.value=zone;
 assetPicker.replaceChildren(option('','설비 선택'),...assets.map(a=>option(a.id,a.name)));
 assetPicker.value=T.selected?.id||'';
 draw();
}
const liteAction=(kind,id)=>{
 if(kind==='zone')T.setZone(id,true);
 else {const a=T.assetMap.get(id);if(a)T.pickAsset(a,false);}
 sync(true);
};
function mapAction(e){const part=e.target.closest?.('[data-part]');if(part&&T.rackIsolateActive){selectedLitePart=rackParts(T.selected).find(p=>p.id===part.dataset.part)||null;if(selectedLitePart&&!window.LS3D_THREE?.adapter)window.LS3D_INSPECTOR?.setSelection(selectedLitePart);sync(true);return}const target=e.target.closest?.('[data-target]');if(target)liteAction(target.dataset.type,target.dataset.target)}
map?.addEventListener('click',mapAction);
map?.addEventListener('keydown',e=>{if(e.key!=='Enter'&&e.key!==' ')return;e.preventDefault();mapAction(e)});
zonePicker?.addEventListener('change',()=>liteAction('zone',zonePicker.value));
assetPicker?.addEventListener('change',()=>liteAction('asset',assetPicker.value));
function activate(why='manual'){
 active=true;reason=why;
 unsupported=!T.gl||why==='webglcontextlost';
 document.body.classList.add('ls3d-2d-active');layer.hidden=false;
 $('ls3dLiteReason').textContent=unsupported?'WebGL 미지원 또는 그래픽 오류 · 경량 도면':'수동 전환 · GPU 렌더링 일시 중단';
 $('ls3dLiteReturn').hidden=unsupported;
 sync(true);
}
function deactivate(){
 if(unsupported||!T.gl)return false;
 active=false;layer.hidden=true;document.body.classList.remove('ls3d-2d-active');
 return true;
}
$('ls3dLiteReturn')?.addEventListener('click',deactivate);
for(const id of ['scene','sceneThree'])$(id)?.addEventListener('webglcontextlost',e=>{e.preventDefault();activate('webglcontextlost')});
window.LS3D_FALLBACK={activate,deactivate,sync,get active(){return active},get reason(){return reason}};
function mobileDefaults(){
 const compact=portable();
 document.body.classList.toggle('ls3d-compact-mobile',compact);
 if(compact&&!window.__LS3D_MOBILE_QUALITY_SET&&window.LS3D_CONFIG){
   const c=window.LS3D_CONFIG;c.quality='low';window.__LS3D_MOBILE_QUALITY_SET=true;
   const b=$('qualityToggle');if(b){b.textContent='품질: 낮음';b.setAttribute('aria-label','화면 품질: 낮음 · 눌러 변경');}
 }
}
mobileDefaults();
window.addEventListener('resize',()=>{mobileDefaults();if(active)sync(true)});
window.addEventListener('ls3d-display-mode-changed',()=>{mobileDefaults();if(active)sync(true)});
window.addEventListener('ls3d-tick',()=>sync());
if(!T.gl)activate('webgl-unavailable');
else if(new URLSearchParams(window.location.search).get('view')==='2d')activate('manual');
// On phones, do not force an onboarding modal over the scene. Re-open with '? 안내'.
document.addEventListener('DOMContentLoaded',()=>{if(portable())$('guide')?.classList.remove('open')});
})();
