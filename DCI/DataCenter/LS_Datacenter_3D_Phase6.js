/* Phase 6 v4.7.0 | schematic 3D MEP overlays and AABB-based clearance precheck.
 * NOT Revit/BIM, neither structural engineering approval nor construction dimensions. */
(function(){
'use strict';const T=window.__LS3D_TEST__,view=document.getElementById('viewport');if(!T||!view)return;
const $=id=>document.getElementById(id),assets=T.assets;
const el=document.createElement('aside');el.id='phase6-panel';el.className='phase58-panel';el.hidden=true;
el.innerHTML=[
'<header><b>PHASE 6 · 랙·MEP 정밀 개념도</b><button type="button" id="p6-close">✕</button></header>',
'<p class="p58-note">실측 CAD/BIM 아님. 랙 치수·배관·버스웨이는 기존 가상 좌표의 개념 구성입니다. 캡슐 선분–AABB 충돌·간격 검사는 실제 시공 승인 기준이 아닙니다.</p>',
'<label><input type="checkbox" id="p6-pipe" checked> 수랭 Supply/Return 배관 및 밸브</label>',
'<label><input type="checkbox" id="p6-power" checked> 상부 전력 Busway · 랙 분기</label>',
'<label><input type="checkbox" id="p6-rack" checked> 랙 서버/ToR 구획 · 케이블 경로</label>',
'<label>배관 경로 Z 오프셋 (가상 m)<input id="p6-offset" type="range" min="-4" max="4" value="0" step=".5"><output id="p6-offset-out">0 m</output></label>',
'<label>요구 이격 거리 (가정 m)<input id="p6-clearance" type="number" min=".1" max="2" step=".1" value=".5"></label>',
'<div class="p58-actions"><button id="p6-check" type="button">충돌·간섭 검사</button><button id="p6-go" type="button">Data Hall 이동</button><button id="p6-csv" type="button">검사 CSV</button></div>',
'<div id="p6-status" role="status">검사 전</div><div id="p6-results"></div>',
'<small class="p58-foot">검사 범위: 축 평행 간격이 지정된 이격 미만인 배관/버스웨이와 rack footprint. 3D 기계 조립체, 유체압력, 굽힘 반경, 화재구획 및 내진 검토는 미포함.</small>'
].join('');document.body.appendChild(el);
const btn=document.createElement('button');btn.id='phase6-open';btn.className='phase58-launch p6';btn.textContent='▥ 랙 / MEP 검사';btn.type='button';view.appendChild(btn);
let opened=false,last=[];
function options(){return {pipe:$('p6-pipe').checked,power:$('p6-power').checked,rack:$('p6-rack').checked,offset:Number($('p6-offset').value),clearance:Number($('p6-clearance').value)}}
function pipes(opt){const h=T.Z.find(z=>z.id==='hall'),x=h.pos[0],z=h.pos[1],o=opt.offset;
 return [{id:'DLC-SUPPLY',kind:'cooling',from:[x-18,6.7,z-2+o],to:[x+17,6.7,z-2+o],radius:.13,color:'#51c6eb'},
 {id:'DLC-RETURN',kind:'cooling',from:[x-18,6.45,z-1.4+o],to:[x+17,6.45,z-1.4+o],radius:.13,color:'#2c8cbb'},
 {id:'POWER-BUSWAY',kind:'power',from:[x-18,6.8,z-13],to:[x+17,6.8,z-13],radius:.22,color:'#efb964'}];
}
function gap1(a0,a1,b0,b1){return Math.max(0,b0-a1,a0-b1)}
function pointBoxDistance(p,a){
 const q=[Math.max(a.x-a.w/2-p[0],0,p[0]-(a.x+a.w/2)),Math.max(-p[1],0,p[1]-a.h),Math.max(a.z-a.d/2-p[2],0,p[2]-(a.z+a.d/2))];
 return Math.hypot(...q);
}
// True capsule-to-AABB distance for an axis-aligned cabinet, without treating a diagonal
// pipe segment as the entire enclosing rectangular prism. The distance to a convex box
// along a straight segment is convex; fixed-iteration ternary minimization converges.
function clearance(pipe,a){
 if(![...pipe.from,...pipe.to,pipe.radius,a.x,a.z,a.w,a.h,a.d].every(Number.isFinite))throw Error('MEP geometry must contain finite values');
 let lo=0,hi=1;
 for(let i=0;i<65;i++){
  const l=lo+(hi-lo)/3,r=hi-(hi-lo)/3;
  const d=t=>pointBoxDistance(pipe.from.map((v,k)=>v+(pipe.to[k]-v)*t),a);
  if(d(l)<d(r))hi=r;else lo=l;
 }
 const t=(lo+hi)/2;
 const p=pipe.from.map((v,k)=>v+(pipe.to[k]-v)*t);
 return Math.max(0,pointBoxDistance(p,a)-pipe.radius);
}
function check(){
 const opt=options();if(!Number.isFinite(opt.clearance)||opt.clearance<.1||opt.clearance>2)throw Error('이격 거리 .1–2m');
 const equipment=assets.filter(a=>a.zone==='hall'&&/rack|switch|ODF|CDU|CRAH|cage/i.test(a.type));
 const routes=pipes(opt).filter(p=>(p.kind==='power'?opt.power:opt.pipe));
 const table=routes.flatMap(route=>equipment.map(asset=>{
 const d=clearance(route,asset);
 return {route:route.id,asset:asset.id,name:asset.name,spacingM:Number(d.toFixed(3)),requiredM:opt.clearance,collision:d<opt.clearance};
 }));
 return{rows:table,violations:table.filter(x=>x.collision),testedAssets:equipment.length,routeCount:routes.length,units:'schematic metres',notice:'segment-to-AABB capsule precheck on assumed concept positions; not construction clearance'};
}
function draw(ctx){
 const opt=options();if(T.currentZone!=='hall')return;
 const route=pipes(opt);
 for(const p of route){
  if((p.kind==='power'&&!opt.power)||(p.kind==='cooling'&&!opt.pipe))continue;
  ctx.line(p.from,p.to,p.color);
  for(let j=0;j<6;j++){const f=j/5,x=p.from[0]+(p.to[0]-p.from[0])*f,z=p.from[2]+(p.to[2]-p.from[2])*f;
   ctx.cylinder(x,p.from[1]-.34,z,.055,.5,p.color,6);
   ctx.box(x,p.from[1]-.38,z,.5,.12,.5,'#6f8797');
  }
 }
 if(!opt.rack)return;
 for(const a of assets.filter(x=>x.zone==='hall'&&(/GPU server rack|Top-of-Rack switch/.test(x.type)))){
  // Draw as additive internals outside the original rack: service-line guides above cabinet only.
  ctx.line([a.x-a.w*.37,a.h+.2,a.z],[a.x+a.w*.37,a.h+.2,a.z],'#85dec5');
  ctx.box(a.x,a.h+.38,a.z-a.d*.25,a.w*.66,.12,.26,'#e6b75c');
  for(let k=-1;k<=1;k++)ctx.line([a.x+k*a.w*.24,a.h+.25,a.z+a.d*.14],[a.x+k*a.w*.24,a.h+.25,a.z+a.d*.48],'#4abbdd');
 }
}
if(Array.isArray(window.LS3D_RESPONSE_VISUALS))window.LS3D_RESPONSE_VISUALS.push(draw);
function render(){let v;try{v=check()}catch(e){$('p6-status').textContent=e.message;return}
 last=v.rows;$('p6-status').textContent=v.routeCount+'개 라우트 · '+v.testedAssets+'개 설비 · 이격 미달 '+v.violations.length+'건';
 const rows=v.violations.slice(0,14);$('p6-results').innerHTML=rows.length?rows.map(x=>'<p><b>'+x.route+'</b> / '+x.asset+' · '+x.spacingM.toFixed(2)+'m</p>').join(''):'<p>선택 경로에서 이격 미달이 검출되지 않았습니다. 단, 비검출은 시공 승인이 아닙니다.</p>';
}
function toggle(v){opened=v;el.hidden=!v;btn.setAttribute('aria-expanded',String(v));if(v)render()}
btn.onclick=()=>toggle(!opened);$('p6-close').onclick=()=>toggle(false);
$('p6-go').onclick=()=>{T.setZone('hall',true);render()};
$('p6-check').onclick=render;
$('p6-offset').oninput=e=>{$('p6-offset-out').value=e.target.value+' m';render()};
$('p6-clearance').onchange=render;for(const id of ['p6-pipe','p6-power','p6-rack'])$(id).onchange=render;
$('p6-csv').onclick=()=>{
 const data=last.length?last:check().rows,s='\uFEFFroute,asset,name,spacing_m,required_m,collision\n'+data.map(x=>[x.route,x.asset,'"'+String(x.name).replace(/"/g,'""')+'"',x.spacingM,x.requiredM,x.collision].join(',')).join('\n');
 const url=URL.createObjectURL(new Blob([s],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='LS_Datacenter_MEP_clearance_concept.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1200);
};
window.LS3D_PHASE6={open:()=>toggle(true),get options(){return options()},get checks(){return check()},get drawings(){return pipes(options())},clearance,run:render};
})();