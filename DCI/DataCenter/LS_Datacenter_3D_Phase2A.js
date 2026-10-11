/* v4.6.0 Phase 2-A | MEP field log and lightweight checked-work geometry | concept only */
(function(){
'use strict';
const T=window.__LS3D_TEST__,view=document.getElementById('viewport');
if(!T||!view)return;
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const tasks={
campus:['보안 게이트·수하물 검사','지붕 배수·우수 경로','외벽 루버·급배기','외부 케이블 지지대','장비 서비스 동선'],
utility:['수전·변압기 전력 덕트','변압기 접지·본딩','발전기 배기 루버','전력 버스덕트 지지대','ATS 유지보수 통로'],
cooling:['냉수 공급·환수 배관','CDU/펌프/HX 격리 밸브','냉각탑 급배기 루버','지붕 드레인·우수 배수','펌프 진동방지 지지대'],
gray:['UPS/PDU 전력 케이블 트레이','배터리실 환기·공조','전기실 유지보수 간격','천장 행거·관통부','배관 내진 지지대'],
hall:['랙 상부 광케이블 트레이','천장 배관·행거','CDU 수랭 Supply/Return','서버 레일 서비스 간격','ToR·MPO 카세트'],
network:['MMR ODF 광패치 동선','Leaf–Spine 경로 이중화','네트워크실 배기 루버','광패치패널 지지대','CPO 서비스 동선'],
ops:['NOC/DCIM 경보 인터페이스','관제실 천장 케이블','화재감지·소방연동','출입통제·CCTV','지붕·외벽 누수 점검']
};
const zones=T.Z.filter(x=>Object.hasOwn(tasks,x.id)),catalog=zones.flatMap(z=>tasks[z.id].map((label,i)=>({key:z.id+':'+i,zone:z.id,label})));
const key='lsdc-phase2-mep-v1';
function restore(){try{let s=JSON.parse(localStorage.getItem(key)||'{}');return{checked:s.checked&&typeof s.checked==='object'?s.checked:{},history:Array.isArray(s.history)?s.history.slice(-70):[]}}catch(e){return{checked:{},history:[]}}}
let store=restore(),zone=T.currentZone in tasks?T.currentZone:'hall',lastScenario=T.scenario,opened=false,renderAt=0;
function persist(){try{localStorage.setItem(key,JSON.stringify(store))}catch(e){}}
const checkedCount=()=>catalog.filter(x=>store.checked[x.key]).length;
function log(message,asset){
 const e={time:new Date().toISOString(),message};store.history.push(e);store.history=store.history.slice(-70);persist();
 if(typeof T.addEventLog==='function')try{T.addEventLog('INFO',message,asset||T.selected||null,'')}catch(err){console.warn('Phase2 MEP logging bridge',err)}
}
let bar=document.createElement('span');bar.className='phase2-buttons-top';bar.id='phase2-bar';
bar.innerHTML='<button type="button" id="phase2-mep-open" aria-controls="phase2-mep-panel" aria-expanded="false">▤ MEP 작업</button><button type="button" id="phase2-opt-open" aria-controls="phase2-opt-panel" aria-expanded="false">◎ 광 연결</button>';
view.appendChild(bar);
let el=document.createElement('aside');el.id='phase2-mep-panel';el.className='phase2-side-panel';el.setAttribute('role','dialog');el.setAttribute('aria-label','구역별 설비 작업 체크리스트');el.hidden=true;
el.innerHTML=[
'<header><b>PHASE 2-A · MEP 작업 로그</b><button id="phase2-mep-close" type="button" aria-label="닫기">✕</button></header>',
'<p class="phase2-notice">Revit/Astra 스타일 검토 보드 · 실측 CAD/BIM 아님 · 가상값/가정값</p>',
'<div class="phase2-progress"><strong id="phase2-mep-count">0 / 35</strong><span id="phase2-mep-percent">0%</span><div class="phase2-meter"><i id="phase2-mep-meter"></i></div></div>',
'<label>구역 <select id="phase2-mep-zone"></select></label>',
'<p id="phase2-mep-zone-desc"></p><div id="phase2-mep-items"></div>',
'<div class="phase2-actions"><button type="button" id="phase2-mep-go">구역 현장 이동</button><button type="button" id="phase2-mep-export">CSV 출력</button></div>',
'<h4>작업 진행 로그 · 이벤트 로그 연동</h4><div id="phase2-mep-log" aria-live="polite"></div>',
'<p class="phase2-notice">체크는 설계 검토 완료 표시이며 실제 시공/준공 확인이 아닙니다. 선택 항목의 일부 경량 3D 형상은 설비 위치를 설명하는 개념 표시입니다.</p>'
].join('');
document.body.appendChild(el);
const sel=$('phase2-mep-zone');sel.innerHTML=zones.map(z=>'<option value="'+esc(z.id)+'">'+esc(z.ko||z.n)+'</option>').join('');
function render(){
 sel.value=zone;const done=checkedCount(),total=catalog.length;
 $('phase2-mep-count').textContent=done+' / '+total;$('phase2-mep-percent').textContent=Math.round(done/total*100)+'%';
 $('phase2-mep-meter').style.width=done/total*100+'%';
 const z=zones.find(a=>a.id===zone);$('phase2-mep-zone-desc').textContent=(z?.sub||'')+' · '+tasks[zone].filter((_,i)=>store.checked[zone+':'+i]).length+'/'+tasks[zone].length;
 $('phase2-mep-items').innerHTML=tasks[zone].map((txt,i)=>'<label class="phase2-check'+(store.checked[zone+':'+i]?' checked':'')+'"><input type="checkbox" data-mep="'+esc(zone+':'+i)+'"'+(store.checked[zone+':'+i]?' checked':'')+'><span>'+esc(txt)+'</span><small>'+(store.checked[zone+':'+i]?'완료':'대기')+'</small></label>').join('');
 $('phase2-mep-log').innerHTML=store.history.length?store.history.slice(-18).reverse().map(o=>'<div><time>'+esc(o.time.slice(0,19).replace('T',' '))+'</time> '+esc(o.message)+'</div>').join(''):'<small>아직 작업 기록이 없습니다.</small>';
}
function change(key,val){
 const item=catalog.find(x=>x.key===key);if(!item)return false;
 store.checked[key]=!!val;persist();log((val?'MEP 검토 완료':'MEP 검토 재개')+' · '+item.zone+' · '+item.label,T.assets.find(a=>a.zone===item.zone));render();return true;
}
function open(){opened=true;el.hidden=false;$('phase2-mep-open').setAttribute('aria-expanded','true');render()}
function close(){opened=false;el.hidden=true;$('phase2-mep-open').setAttribute('aria-expanded','false')}
$('phase2-mep-open').onclick=()=>opened?close():open();$('phase2-mep-close').onclick=close;
sel.onchange=e=>{zone=e.target.value;render()};
$('phase2-mep-items').onchange=e=>{const id=e.target.dataset.mep;if(id)change(id,e.target.checked)};
$('phase2-mep-go').onclick=()=>{T.setZone(zone,true);log('MEP 대상 구역으로 이동 · '+zone);render()};
$('phase2-mep-export').onclick=()=>{
 const rows=[['구역','항목','검토 완료'],...catalog.map(x=>[x.zone,x.label,store.checked[x.key]?'YES':'NO'])];
 const csv='\uFEFF'+rows.map(row=>row.map(x=>'"'+String(x).replace(/"/g,'""')+'"').join(',')).join('\r\n');
 const url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));let link=document.createElement('a');link.href=url;link.download='LS_Datacenter_Phase2_MEP.csv';link.click();setTimeout(()=>URL.revokeObjectURL(url),1500);
};
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&opened)close()});
function mepGeometry(ctx){
 if(T.currentZone==='campus'||!tasks[T.currentZone])return;
 const z=zones.find(x=>x.id===T.currentZone);if(!z)return;
 const [x,p]=z.pos,id=z.id;
 if(store.checked[id+':0']){ctx.line([x-6,6.1,p-3],[x+6,6.1,p-3],'#f8cb6a');ctx.line([x-6,5.7,p-2.6],[x+6,5.7,p-2.6],'#42bdd9')}
 if(store.checked[id+':1'])ctx.line([x-6,6.5,p+3],[x+6,6.5,p+3],'#62d2c8');
 if(store.checked[id+':2'])for(let k=0;k<3;k++)ctx.box(x+8,2.5+k*.46,p+5,1.4,.09,.16,'#658c9c');
 if(store.checked[id+':3'])for(let dx of [-5,5])ctx.cylinder(x+dx,0,p+5,.11,6.1,'#7495a4',6);
 if(store.checked[id+':4'])ctx.line([x-4,.12,p-6],[x+4,.12,p-6],'#f9d777');
}
if(Array.isArray(window.LS3D_RESPONSE_VISUALS))window.LS3D_RESPONSE_VISUALS.push(mepGeometry);
window.addEventListener('ls3d-tick',()=>{
 const now=performance.now();if(now-renderAt<450)return;renderAt=now;
 if(T.scenario!==lastScenario){lastScenario=T.scenario;log('운영 시나리오 변경 · '+T.scenario,T.selected);if(opened)render()}
});
window.LS3D_PHASE2_A={open,close,change,get checkedCount(){return checkedCount()},get totalCount(){return catalog.length},get tasks(){return catalog.map(x=>({...x,checked:!!store.checked[x.key]}))},get history(){return store.history.slice()},get openState(){return opened},get activeZone(){return zone}};
render();
})();
