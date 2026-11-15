import {drawRoomBase,drawRoomEquipment,drawIllustratedBase,drawIllustratedEquipment} from './room-drawing.mjs?v=20261116-dual-2d-v5';
// The plan uses the same host assets and zones as the 3D tour, never a separate inventory.
export function planBounds(zones,assets){
 const boxes=[...zones.map(z=>[z.pos[0]-z.size[0]/2,z.pos[1]-z.size[1]/2,z.size[0],z.size[1]]),...assets.map(a=>[a.x-a.w/2,a.z-a.d/2,a.w,a.d])].filter(b=>b.every(Number.isFinite));
 if(!boxes.length)return [-10,-10,20,20];
 const x=Math.min(...boxes.map(b=>b[0])),y=Math.min(...boxes.map(b=>b[1]));
 return [x-4,y-4,Math.max(8,Math.max(...boxes.map(b=>b[0]+b[2]))-x)+8,Math.max(8,Math.max(...boxes.map(b=>b[1]+b[3]))-y)+8];
}
export function planKind(a){const t=(a.type||'')+' '+a.id;if(/Storage/i.test(t))return ['스토리지','#e7dcff'];if(/rack|ODF|switch/i.test(t))return ['랙·네트워크','#cce9f3'];if(/ups|battery|generator|transformer|switchgear|pdu|ats/i.test(t))return ['전력','#ffedc7'];if(/cool|chiller|crah|cdu|pump/i.test(t))return ['냉각','#c9f0e2'];return ['기타 설비','#e1e6ec'];}
export function createFloorPlan(T){
 const viewport=document.getElementById('viewport'),NS='http://www.w3.org/2000/svg';
 const el=(tag,attrs={},text)=>{const n=document.createElementNS(NS,tag);for(const [k,v] of Object.entries(attrs))n.setAttribute(k,String(v));if(text!==undefined)n.textContent=text;return n;};
 const button=(text,fn)=>{const b=document.createElement('button');b.type='button';b.textContent=text;b.onclick=fn;return b;};
 const tabs=document.createElement('div');tabs.className='ls-plan-tabs';tabs.setAttribute('role','tablist');tabs.setAttribute('aria-label','도면 보기 방식');
 const tour=button('3D 투어',()=>setActive(false)),plan=button('2D 도면',()=>setActive(true));
 for(const [b,id] of [[tour,'tour'],[plan,'plan']]){b.id='ls-view-tab-'+id;b.setAttribute('role','tab');b.setAttribute('aria-controls',id==='plan'?'ls-floor-plan':'scene');}
 tabs.append(tour,plan);viewport.append(tabs);
 const panel=document.createElement('section');panel.className='ls-floor-plan';panel.id='ls-floor-plan';panel.hidden=true;panel.setAttribute('role','tabpanel');panel.setAttribute('aria-labelledby',plan.id);
 const tools=document.createElement('div');tools.className='ls-plan-tools';
 const drawingMode=document.createElement('select');drawingMode.setAttribute('aria-label','2D 도면 표시 방식');for(const [value,label] of [['drawing','상세 도면'],['illustrated','그림형 배치도']]){const o=document.createElement('option');o.value=value;o.textContent=label;drawingMode.append(o);}drawingMode.onchange=()=>{signature='';update();};
 const heading=document.createElement('strong');heading.textContent='2D 설비 배치도';
 const zone=document.createElement('select');zone.setAttribute('aria-label','2D 도면 구역');zone.onchange=()=>{T.closeRackScene();T.setZone(zone.value);update();};
 tools.append(heading,zone,drawingMode,button('확대 +',()=>zoom(.8)),button('축소 −',()=>zoom(1.25)),button('전체 맞춤',()=>{view=fit.slice();applyView();}),button('도면 크게 / 닫기',()=>{viewport.classList.toggle('ls-plan-expanded');}));
 const svg=el('svg',{role:'group','aria-label':'데이터센터 구역과 설비의 2D 평면 배치도',preserveAspectRatio:'xMidYMid meet',tabindex:0});
 const foot=document.createElement('div');foot.className='ls-plan-footer';
 const selected=document.createElement('span');selected.textContent='설비를 선택하면 이름과 유형이 표시됩니다.';
 const open=button('선택 설비 3D로 보기',()=>{if(T.selected){setActive(false);T.pickAsset(T.selected,true);}});open.disabled=true;
 const legend=document.createElement('small');legend.textContent='공간 선택: 상세 평면도 · 번호: 설비 일람표 대응 · 드래그/방향키: 이동 · 휠/두 손가락: 확대·축소 · u: 모델 좌표 단위';
 foot.append(selected,open,legend);panel.append(tools,svg,foot);viewport.append(panel);
 const style=document.createElement('style');style.textContent=`
 .ls-plan-tabs{position:absolute;top:96px;left:12px;z-index:46;display:flex;padding:4px;gap:4px;border:1px solid #54717e;border-radius:10px;background:#102b3bee}
 .ls-plan-tabs button,.ls-plan-tools button,.ls-plan-tools select,.ls-plan-footer button{font:inherit;font-size:12px;min-height:36px;padding:7px 11px;border:1px solid #9aabb9;border-radius:7px;background:white;color:#203849;cursor:pointer}
 .ls-plan-tabs button{background:transparent;color:white;border-color:transparent}.ls-plan-tabs [aria-selected=true]{background:#b6eee7;color:#173b43}.ls-plan-tabs button:focus-visible{outline:2px solid #35bcc1}
 .ls-floor-plan{position:absolute;inset:58px 12px 12px;z-index:40;background:#f7fafc;border:1px solid #bacbd6;border-radius:12px;display:flex;flex-direction:column;color:#203849;overflow:hidden}.ls-floor-plan[hidden]{display:none!important}
 .ls-plan-tools{display:flex;flex-wrap:wrap;gap:7px;align-items:center;padding:10px;border-bottom:1px solid #d2dde5}.ls-plan-tools strong{margin-right:auto}.ls-floor-plan svg{width:100%;flex:1;min-height:140px;touch-action:none;background:#f7fafc;cursor:grab}.ls-floor-plan svg:active{cursor:grabbing}.ls-floor-plan svg [data-room]{cursor:pointer}.ls-floor-plan svg text{user-select:none}.ls-plan-footer{display:flex;flex-wrap:wrap;align-items:center;gap:8px;padding:10px;border-top:1px solid #d2dde5;font-size:12px}.ls-plan-footer small{flex-basis:100%;font-size:11px;color:#516678}.ls-plan-footer span{flex:1}.ls-floor-plan [data-asset]{cursor:pointer}.ls-floor-plan [data-asset]:focus rect{stroke:#0b6980;stroke-width:.4}
 #viewport.ls-floor-plan-active > :not(.ls-plan-tabs):not(.ls-floor-plan){visibility:hidden}#viewport.ls-floor-plan-active .ls-plan-tabs{top:8px}
 #viewport.ls-plan-expanded::after{content:'';position:fixed;inset:0;z-index:10000;background:#102333}
 #viewport.ls-plan-expanded .ls-floor-plan{position:fixed;inset:62px 12px 12px;z-index:10001}
 #viewport.ls-plan-expanded .ls-plan-tabs{position:fixed;top:10px;left:12px;z-index:10002}
 @media(max-width:790px){.ls-plan-tabs{top:106px;left:8px}.ls-floor-plan{left:6px;right:6px;bottom:6px}.ls-plan-tools{padding:7px}.ls-plan-tools strong{flex-basis:100%}.ls-plan-footer{padding:7px}}
 `;document.head.append(style);
 let active=false,signature='',selection='',view=[-10,-10,20,20],fit=view.slice(),drag=null;
 const pointers=new Map();let gesture=null;
 function applyView(){svg.setAttribute('viewBox',view.join(' '));}
 function zoom(factor){const w=Math.max(fit[2]/30,Math.min(fit[2]*4,view[2]*factor)),f=w/view[2];view=[view[0]+view[2]*(1-f)/2,view[1]+view[3]*(1-f)/2,w,view[3]*f];applyView();}
 function setActive(on){if(!on)window.LS3D_FALLBACK?.deactivate?.();active=on;if(!on)viewport.classList.remove('ls-plan-expanded');panel.hidden=!on;viewport.classList.toggle('ls-floor-plan-active',on);tour.setAttribute('aria-selected',String(!on));plan.setAttribute('aria-selected',String(on));tour.tabIndex=on?-1:0;plan.tabIndex=on?0:-1;if(on){signature='';update();svg.focus();}else tour.focus();}
 tabs.addEventListener('keydown',e=>{if(['ArrowLeft','ArrowRight','Home','End'].includes(e.key)){e.preventDefault();setActive(e.key==='Home'?false:e.key==='End'?true:!active);(active?plan:tour).focus();}});
 function choose(a){T.pickAsset(a,false);selection='';update();}
 svg.addEventListener('pointerdown',e=>{e.preventDefault();svg.focus();pointers.set(e.pointerId,[e.clientX,e.clientY]);svg.setPointerCapture(e.pointerId);if(pointers.size===1)drag={x:e.clientX,y:e.clientY,view:view.slice(),moved:false,target:e.target};if(pointers.size===2){const [a,b]=[...pointers.values()];gesture={distance:Math.hypot(a[0]-b[0],a[1]-b[1]),view:view.slice()};if(drag)drag.moved=true;}});
 svg.addEventListener('pointermove',e=>{if(!pointers.has(e.pointerId))return;pointers.set(e.pointerId,[e.clientX,e.clientY]);if(pointers.size===2&&gesture){const [a,b]=[...pointers.values()],d=Math.hypot(a[0]-b[0],a[1]-b[1]);view=gesture.view.slice();zoom(gesture.distance/Math.max(1,d));return;}if(!drag||gesture)return;const r=svg.getBoundingClientRect(),scale=Math.min(r.width/drag.view[2],r.height/drag.view[3]),dx=e.clientX-drag.x,dy=e.clientY-drag.y;if(Math.hypot(dx,dy)>4)drag.moved=true;view=[drag.view[0]-dx/scale,drag.view[1]-dy/scale,...drag.view.slice(2)];applyView();});
 function end(e){if(drag&&!drag.moved&&!gesture&&e.type==='pointerup'){const id=drag.target.closest('[data-asset]')?.getAttribute('data-asset'),a=T.assetMap.get(id);if(a)choose(a);else{const room=drag.target.closest('[data-room]')?.getAttribute('data-room');if(room){T.closeRackScene();T.setZone(room);signature='';update();}}}pointers.delete(e.pointerId);drag=null;if(!pointers.size)gesture=null;}
 for(const event of ['pointerup','pointercancel','lostpointercapture'])svg.addEventListener(event,end);
 svg.addEventListener('wheel',e=>{e.preventDefault();zoom(Math.exp(Math.max(-.3,Math.min(.3,e.deltaY*.002))));},{passive:false});
 svg.addEventListener('keydown',e=>{if(['+','=','-'].includes(e.key)){e.preventDefault();zoom(e.key==='-'?1.2:1/1.2);}else{const d={ArrowLeft:[-1,0],ArrowRight:[1,0],ArrowUp:[0,-1],ArrowDown:[0,1]}[e.key];if(d){e.preventDefault();view[0]+=d[0]*view[2]*.06;view[1]+=d[1]*view[3]*.06;applyView();}}});
 function update(){if(!active)return;
 const allAssets=T.visibleAssets().filter(a=>[a.x,a.z,a.w,a.d].every(Number.isFinite));
 const assets=allAssets.filter(a=>T.currentZone==='campus'||(a.zone===T.currentZone&&a.type!=='Visitor tour'));
 const zoneIds=new Set(allAssets.map(a=>a.zone));const zones=T.Z.filter(z=>z.id!=='campus'&&(T.currentZone==='campus'?zoneIds.has(z.id):z.id===T.currentZone));
 const sig=drawingMode.value+'|'+T.facilityMode+'|'+T.currentZone+'|'+JSON.stringify(assets.map(a=>[a.id,a.name,a.x,a.z,a.w,a.d]));
 if(sig!==signature){signature=sig;svg.replaceChildren();
 const available=T.Z.filter(z=>z.id==='campus'||zoneIds.has(z.id)||z.id===T.currentZone);zone.replaceChildren(...available.map(z=>{const o=document.createElement('option');o.value=z.id;o.textContent=z.ko||z.n;return o;}));zone.value=T.currentZone;
 fit=planBounds(zones,assets);view=fit.slice();applyView();
 const detailed=T.currentZone!=='campus'&&zones.length===1;
 heading.textContent=detailed?(zones[0].ko||zones[0].n)+(drawingMode.value==='drawing'?' · 상세 평면도':' · 그림형 배치도'):'캠퍼스 배치도 · 공간을 선택하세요';
 if(detailed){fit=(drawingMode.value==='drawing'?drawRoomBase:drawIllustratedBase)(svg,el,zones[0],assets,T.facilityMode);view=fit.slice();applyView();}
 else for(const z of zones){svg.append(el('rect',{x:z.pos[0]-z.size[0]/2,y:z.pos[1]-z.size[1]/2,width:z.size[0],height:z.size[1],'data-room':z.id,fill:'#fff',stroke:'#647c90','stroke-width':.25}));svg.append(el('text',{x:z.pos[0]-z.size[0]/2+.6,y:z.pos[1]-z.size[1]/2+1.2,'font-size':1,fill:'#304e64'},z.ko||z.n));}
 for(const [index,a] of assets.entries()){const g=el('g',{'data-asset':a.id,tabindex:0,role:'button','aria-label':a.name+' · '+a.type});g.append(el('title',{},a.name+' · '+a.type));if(detailed){(drawingMode.value==='drawing'?drawRoomEquipment:drawIllustratedEquipment)(g,el,a,index);}else{g.append(el('rect',{x:a.x-a.w/2,y:a.z-a.d/2,width:a.w,height:a.d,rx:.1,fill:planKind(a)[1],stroke:'#647c90','stroke-width':.16}));g.append(el('text',{x:a.x,y:a.z,'text-anchor':'middle','dominant-baseline':'middle','font-size':Math.max(.28,Math.min(.75,a.w/Math.max(3,a.name.length)*1.4)),fill:'#19364b','pointer-events':'none'},a.name));}g.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();e.stopPropagation();choose(a);}});svg.append(g);}
 selection=null;}
 const currentSelected=assets.find(a=>a.id===T.selected?.id);const id=currentSelected?.id||'';if(id!==selection){selection=id;for(const g of svg.querySelectorAll('[data-asset]')){const on=g.getAttribute('data-asset')===id;g.setAttribute('aria-pressed',String(on));g.querySelector('rect').setAttribute('stroke',on?'#007d87':'#647c90');g.querySelector('rect').setAttribute('stroke-width',on?'.45':'.16');}selected.textContent=currentSelected?currentSelected.name+' · '+currentSelected.type:'설비를 선택하면 이름과 유형이 표시됩니다.';open.disabled=!currentSelected;}
 }
 document.addEventListener('keydown',e=>{if(e.key==='Escape')viewport.classList.remove('ls-plan-expanded');});
 tour.setAttribute('aria-selected','true');plan.setAttribute('aria-selected','false');plan.tabIndex=-1;return {update,setActive};
}

