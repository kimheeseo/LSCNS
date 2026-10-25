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
#ls3dModeToggle{border:1px solid #5d9aa9;border-radius:8px;background:#174356;color:#e7fff9;padding:8px 10px;font:750 11px system-ui;min-height:36px;cursor:pointer;white-space:nowrap}
#ls3dPortable2D[hidden]{display:none!important}
#ls3dPortable2D{position:absolute;inset:0;z-index:14;background:#0a1b2b;color:#e9f8ff;display:flex;flex-direction:column;gap:8px;padding:10px;overflow:hidden}
#ls3dPortable2D .lite-header{display:flex;align-items:center;justify-content:space-between;gap:6px}
#ls3dPortable2D .lite-header strong{font-size:14px}
#ls3dPortable2D .lite-header small{display:block;font-size:10px;color:#a6c5d4}
#ls3dPortable2D button,#ls3dPortable2D select{min-height:38px;border:1px solid #4a7187;background:#16364a;border-radius:8px;color:#f1fcff;padding:7px 9px;font:650 11px system-ui;max-width:100%}
#ls3dPortable2D button{cursor:pointer}
#ls3dPortable2D .lite-controls{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1.35fr);gap:6px}
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
  #ls3dModeToggle{font-size:10px;padding:6px}
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
const actions=document.querySelector('.top-actions');
const toggle=document.createElement('button');
toggle.type='button';toggle.id='ls3dModeToggle';toggle.textContent='▤ 2D 경량';toggle.setAttribute('aria-label','2D 경량 배치도로 전환');
actions?.appendChild(toggle);
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
let active=false,reason='',lastSignature='',lastSynced=0,unsupported=false;
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
function draw(){
 if(!active||!map)return;
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
 const zones=T.visibleZones(),zone=T.currentZone;
 const assets=T.visibleAssets().filter(a=>zone==='campus'||a.zone===zone);
 const signature=zone+'|'+T.facilityMode+'|'+(T.selected?.id||'')+'|'+assets.map(a=>a.id).join(',');
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
map?.addEventListener('click',e=>{const target=e.target.closest?.('[data-target]');if(target)liteAction(target.dataset.type,target.dataset.target);});
map?.addEventListener('keydown',e=>{if(e.key!=='Enter'&&e.key!==' ')return;const target=e.target.closest?.('[data-target]');if(target){e.preventDefault();liteAction(target.dataset.type,target.dataset.target);}});
zonePicker?.addEventListener('change',()=>liteAction('zone',zonePicker.value));
assetPicker?.addEventListener('change',()=>liteAction('asset',assetPicker.value));
function activate(why='manual'){
 active=true;reason=why;
 unsupported=!T.gl||why==='webglcontextlost';
 document.body.classList.add('ls3d-2d-active');layer.hidden=false;
 $('ls3dLiteReason').textContent=unsupported?'WebGL 미지원 또는 그래픽 오류 · 경량 도면':'수동 전환 · GPU 렌더링 일시 중단';
 $('ls3dLiteReturn').hidden=unsupported;
 toggle.textContent='▤ 2D 보기 중';toggle.setAttribute('aria-pressed','true');toggle.setAttribute('aria-label','2D 경량 보기 활성화');
 sync(true);
}
function deactivate(){
 if(unsupported||!T.gl)return false;
 active=false;layer.hidden=true;document.body.classList.remove('ls3d-2d-active');
 toggle.textContent='▤ 2D 경량';toggle.setAttribute('aria-pressed','false');toggle.setAttribute('aria-label','2D 경량 배치도로 전환');
 return true;
}
toggle.addEventListener('click',()=>{if(active){deactivate();return}activate('manual');});
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
