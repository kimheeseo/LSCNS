/* v4.6.0 Phase 2-B | optical copper/CPO journey and fault/reroute diagrams | concept only */
(function(){
'use strict';
const T=window.__LS3D_TEST__,C=window.LS3D_CONFIG,phase1=window.LS3D_PHASE1;
if(!T||!C||!phase1)return;
const $=id=>document.getElementById(id),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let openState=false,lastSig='',lastPaint=0;
const e=document.createElement('aside');e.id='phase2-opt-panel';e.className='phase2-side-panel phase2-opt-panel';e.hidden=true;e.setAttribute('role','dialog');e.setAttribute('aria-label','서버부터 Spine까지 광 인터커넥트 및 장애 경로');
e.innerHTML=[
'<header><b>PHASE 2-B · OPTICAL PATH</b><button id="phase2-opt-close" aria-label="닫기" type="button">✕</button></header>',
'<p class="phase2-notice">GPU / NIC → ToR/Leaf → Spine → Core · 교육용 광/구리 개념 구성</p>',
'<label>광 모듈 구성 <select id="phase2-opt-mode"><option value="pluggable">Pluggable 트랜시버</option><option value="cpo">CPO · 스위치 집적광학</option></select></label>',
'<div id="phase2-opt-status" aria-live="polite"></div>',
'<svg id="phase2-opt-svg" viewBox="0 0 360 220" role="img" aria-label="정상 및 우회 광 경로" preserveAspectRatio="xMidYMid meet"></svg>',
'<div class="phase2-opt-legend"><span>━━ 구리 DAC/AEC</span><span>━━ 광 주경로 A</span><span>┄┄ 광 우회경로 B</span></div>',
'<p class="phase2-opt-section"><b>현재 줌 단계</b> <span id="phase2-opt-zoom">캠퍼스</span></p>',
'<div id="phase2-opt-detail"></div>',
'<div class="phase2-actions"><button id="phase2-opt-rack" type="button">GPU 랙 조사</button><button id="phase2-opt-existing" type="button">기존 상세 광 토폴로지</button></div>',
'<div class="phase2-actions"><button id="phase2-opt-fault" type="button">광케이블 단선 시뮬레이션</button><button id="phase2-opt-reset" type="button">정상 복구</button></div>',
'<p class="phase2-notice">기본 서버 NIC–ToR 간은 DAC/AEC 구리 링크를 가정하므로 해당 구간의 광트랜시버는 계상하지 않습니다. 광 구간의 CPO는 스위치측 집적 광학 모듈과 MPO를 나타내며, GPU/HBM 내부 형상은 전기적 패키지 개념도입니다. 실제 BOM 산식·제품 선정은 별도 검증 필요.</p>'
].join('');
document.body.appendChild(e);
function state(){
 const broken=!!C.broken||T.scenario==='fiber-cut';
 return{broken,rerouted:broken&&!!C.reroute,link:C.broken||C.selected||'',speed:C.linkSpeed||'800G',mode:C.mode==='cpo'?'cpo':'pluggable',bandwidth:Number(C.brokenBandwidthPct||0),latency:Number(C.brokenLatencyUs||0)}
}
function zoom(){
 const inspect=!!(T.selected&&T.isMechanicalRack(T.selected)&&$('detailModal')?.classList.contains('show'));
 const slide=inspect&&!!(T.mechanical(T.selected).targetSlide||T.mechanical(T.selected).slide>.5);
 const dist=Number(T.camera?.distance||T.cameraDesired.distance);
 return{inspect,slide,distance:dist,level:phase1.stageFor(dist,inspect,slide)}
}
function model(){
 const s=state(),z=zoom();
 const steps=[
 ['GPU · NIC → ToR','기본 NIC→ToR 구간: DAC/AEC 구리 패치. 랙 내 다중 NIC와 ToR 포트를 개념적으로 표시합니다.'],
 ['ToR/Leaf → Spine',s.mode==='cpo'?'ToR/Spine 스위치측 CPO 광학 엔진 + MPO/SMF 트렁크를 표시합니다.':'ToR/Spine 양단 Pluggable 광모듈 + MPO/SMF 트렁크를 표시합니다.'],
 ['서버 트레이','서버 인출 · GPU/NIC 및 보드 단거리 전기적 연결 개념도.'],
 ['GPU 카드','GPU 보드의 PCB 신호 라인과 NIC 연결은 개념 표시이며 광섬유가 GPU 실리콘에 직접 연결된 것이 아닙니다.'],
 ['패키지 · HBM','GPU 칩렛/HBM과 패키지 전기적 연결 개념도입니다. CPO 모듈 위치를 실측한 것이 아닙니다.'],
 ['µm~nm 개념 구간','실측 아님. 광전송·트랜지스터 실물 형상 없이 크기만 개념적으로 확대합니다.']
 ];
 const idx=!z.inspect?(z.level<=2?0:1):z.level===4?2:z.level===5?3:z.level===6?4:z.level>=7?5:1;
 return{state:s,zoom:z,detail:steps[idx],index:idx};
}
const pos={gpu:[34,94],tor:[112,94],spA:[230,46],spB:[230,158],core:[330,94]};
function node(x,y,top,sub){return'<rect x="'+(x-30)+'" y="'+(y-20)+'" width="60" height="40" rx="7" class="phase2-opt-node"/><text x="'+x+'" y="'+(y-3)+'" text-anchor="middle">'+esc(top)+'</text><text x="'+x+'" y="'+(y+10)+'" class="phase2-opt-sub" text-anchor="middle">'+esc(sub)+'</text>'}
function paint(){
 if(!e||e.hidden)return;
 const m=model(),s=m.state;
 $('phase2-opt-mode').value=s.mode;
 let msg;
 if(!s.broken)msg='<b class="phase2-state-ok">정상 경로 A + 이중화 B</b>';
 else if(s.rerouted)msg='<b class="phase2-state-warn">광 링크 단선 · 우회 경로 B 활성화</b>';
 else msg='<b class="phase2-state-bad">광 링크 단선 · 우회 불가</b>';
 $('phase2-opt-status').innerHTML=msg+'<small>'+esc(s.link||'광 네트워크')+' · '+esc(s.speed)+(s.broken?' · '+(s.rerouted?'가용 대역폭 '+s.bandwidth.toFixed(0)+'% · 추가 지연 '+s.latency.toFixed(1)+'µs (가정)':'복구 필요'):' · 2중 경로 예시')+'</small>';
 const p=pos,copper='M 64 94 H 82';
 let svg='<path class="phase2-opt-copper" d="'+copper+'"/>';
 svg+='<path class="phase2-opt-primary'+(s.broken?' damaged':'')+'" d="M 142 88 L 200 48 M 260 48 L 300 88"/>';
 svg+='<path class="phase2-opt-backup'+(s.broken&&!s.rerouted?' inactive':'')+'" d="M 142 100 L 200 156 M 260 156 L 300 100"/>';
 if(s.broken)svg+='<circle class="phase2-opt-broken" cx="170" cy="67" r="9"/><text class="phase2-opt-alert" x="170" y="71" text-anchor="middle">×</text>';
 svg+=node(34,94,'GPU','NIC')+node(112,94,'ToR','Leaf')+node(230,46,'Spine','A')+node(230,158,'Spine','B')+node(330,94,'Core','MMR');
 svg+='<text class="phase2-opt-caption" x="72" y="67" text-anchor="middle">DAC/AEC</text>';
 svg+='<text class="phase2-opt-caption" x="171" y="31" text-anchor="middle">'+(s.mode==='cpo'?'CPO·MPO':'TRX·MPO')+'</text>';
 svg+='<text class="phase2-opt-caption" x="179" y="190" text-anchor="middle">B · redundant</text>';
 $('phase2-opt-svg').innerHTML=svg;
 $('phase2-opt-zoom').textContent=(phase1.stage||'캠퍼스').split(' · ')[0]+(m.zoom.inspect?' · 랙 상세':' · 전체 화면');
 $('phase2-opt-detail').innerHTML='<b>'+esc(m.detail[0])+'</b><p>'+esc(m.detail[1])+'</p><small>'+(m.index>=4?'개념도 · 실측 아님':'실제 장비·포트 연결은 공식 사양 확인 필요')+'</small>';
}
function open(){openState=true;e.hidden=false;$('phase2-opt-open').setAttribute('aria-expanded','true');paint()}
function close(){openState=false;e.hidden=true;$('phase2-opt-open').setAttribute('aria-expanded','false')}
$('phase2-opt-open').onclick=()=>openState?close():open();$('phase2-opt-close').onclick=close;
$('phase2-opt-mode').onchange=ev=>{
 C.mode=ev.target.value;
 const old=$('optMode');if(old){old.value=C.mode;old.dispatchEvent(new Event('change',{bubbles:true}))}
 paint();
 if(typeof T.addEventLog==='function')T.addEventLog('INFO','광 인터커넥트 '+C.mode+' 선택 · 개념 모델',T.selected,'')
};
$('phase2-opt-rack').onclick=()=>{
 const rack=T.assets.find(a=>a.zone==='hall'&&a.type==='GPU server rack'&&!a.hidden&&T.isMechanicalRack(a));
 if(!rack)return;
 T.setZone('hall',false);T.pickAsset(rack,true);T.openDetail();close();
};
$('phase2-opt-existing').onclick=()=>{const b=$('optOpen');if(b)b.click()};
function scenario(id){const select=$('scenarioSelect');if(select){select.value=id;select.dispatchEvent(new Event('change',{bubbles:true}))}else T.scenarioApply(id)}
$('phase2-opt-fault').onclick=()=>{scenario('fiber-cut');paint()};
$('phase2-opt-reset').onclick=()=>{scenario('normal');paint()};
document.addEventListener('keydown',ev=>{if(ev.key==='Escape'&&openState)close()});
const base=window.LS3D_PHASE1_RENDER;
if(typeof base==='function')window.LS3D_PHASE1_RENDER=function(ctx,a,d,mech){
 base(ctx,a,d,mech);
 if(!T.phase1Local)return;
 const m=model(),level=m.zoom.level;
 const color=m.state.broken&&!m.state.rerouted?'#ff7489':m.state.broken?'#f8cb6a':'#54d9c6';
 if(level===4){ctx.line([-.68,-.07,.96],[.68,-.07,.96],color);ctx.line([-.63,-.03,-.34],[.64,-.03,-.34],color)}
 if(level===5){ctx.line([-.18,-.006,-.14],[.18,-.006,-.14],color);ctx.line([.18,-.006,-.14],[.20,-.029,.13],color)}
 if(level===6)ctx.line([-.024,.008,-.009],[.024,.008,-.009],color);
};
let prev='',t0=0;
window.addEventListener('ls3d-tick',()=>{
 const now=performance.now();if(now-t0<170)return;t0=now;
 if(!openState)return;
 const s=state(),z=zoom(),sig=[s.broken,s.rerouted,s.link,s.mode,s.speed,z.level,z.inspect].join('|');
 if(prev!==sig){prev=sig;paint()}
});
window.LS3D_PHASE2_B={open,close,get openState(){return openState},get opticalState(){return state()},get model(){return model()},update:paint};
})();
