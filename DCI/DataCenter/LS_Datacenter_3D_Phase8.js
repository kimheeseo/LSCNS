/* Phase 8 v4.7.0 · non-destructive startup/integration checks and help */
(function(){
'use strict';const T=window.__LS3D_TEST__,view=document.getElementById('viewport');if(!T||!view)return;
const $=id=>document.getElementById(id);
const panel=document.createElement('aside');panel.id='phase8-panel';panel.className='phase58-panel';panel.hidden=true;
panel.innerHTML=[
'<header><b>PHASE 8 · 종합 진단·사용 가이드</b><button type="button" id="p8-close">✕</button></header>',
'<p class="p58-note">기존 시뮬레이터의 상태를 바꾸지 않는 간단 통합 검사입니다. 실제 브라우저 테스트 및 BOM 구매 승인 검증을 대체하지 않습니다.</p>',
'<div class="p58-actions"><button id="p8-check" type="button">통합 점검 실행</button><button id="p8-json" type="button">결과 JSON</button></div>',
'<p id="p8-summary" role="status">점검 전</p><div id="p8-cases"></div>',
'<h4>사용 가이드</h4>',
'<p><a target="_blank" rel="noopener" href="https://github.com/kimheeseo/LSCNS/blob/main/DCI/DataCenter/LS_Datacenter_3D_Phase8_Guide.md">Phase 1–8 종합 사용 가이드 열기 ↗</a></p>',
'<p><a target="_blank" rel="noopener" href="./">데이터센터 BOM 기본 도구 열기 ↗</a></p>',
'<p class="p58-note">실기기 성능 확인: 먼저 「FPS / 기기 검증」에서 Windows·Android·iPhone의 각 브라우저로 측정하고 JSON 저장. 실기기 데이터가 없는 경우 Phase 5는 미검증으로 유지됩니다.</p>',
'<small class="p58-foot">교육·설계 검토용 3D 개념 모델. 실제 NVIDIA 장비 인터페이스/광 링크 및 MEP 이격은 별도 검증 필요.</small>'
].join('');document.body.appendChild(panel);
const btn=document.createElement('button');btn.id='phase8-open';btn.className='phase58-launch p8';btn.type='button';btn.textContent='☑ 종합 점검 / 가이드';view.appendChild(btn);
let report=null,opened=false;
function item(name,test,detail){return{name,pass:!!test,detail:String(detail||'')}}
function checks(){
 const c=window.LS3D_CONFIG,p4=window.LS3D_PHASE4_CACHE,p5=window.LS3D_PHASE5,p6=window.LS3D_PHASE6,p7=window.LS3D_PHASE7;
 const list=[
 item('캠퍼스 7구역 표시 모델',T.Z.length===7,'zones='+T.Z.length),
 item('랙/설비 데이터',Array.isArray(T.assets)&&T.assets.length>=50,'assets='+T.assets.length),
 item('Phase 1 FOV / 확대',typeof window.LS3D_PHASE1?.stageFor==='function','semantic LOD stageFor'),
 item('Phase 2-A MEP 35개',window.LS3D_PHASE2_A?.totalCount===35,'35 reviewed work items'),
 item('Phase 2-B 광 연결 경로',!!window.LS3D_PHASE2_B?.opticalState,'optical path model'),
 item('Phase 4 GPU 정적 지오메트리 캐시',!!p4&&!!p4.gpu&&Number.isFinite(p4.gpu.uploads),'uploads='+p4?.gpu?.uploads),
 item('Phase 5 실기기 성능 진단 기능',!!p5&&!!p5.device?.viewport,'browser self-report only'),
 item('Phase 6 선분–AABB 거리: 비관통',typeof p6?.clearance==='function'&&Math.abs(p6.clearance({from:[-10,5,0],to:[10,5,0],radius:.1},{x:0,z:0,w:2,d:2,h:2})-2.9)<.001,'expected 2.9m'),
 item('Phase 6 선분–AABB 거리: 관통',typeof p6?.clearance==='function'&&p6.clearance({from:[-10,1,0],to:[10,1,0],radius:.1},{x:0,z:0,w:2,d:2,h:2})===0,'expected 0m'),
 item('Phase 6 MEP 형상·이격 검사',!!p6&&Array.isArray(p6.checks.rows)&&p6.checks.routeCount>=2,'capsule/AABB concept precheck'),
 item('Phase 7 NVIDIA 카탈로그 감사 함수',typeof p7?.estimate==='function'&&typeof p7?.load==='function','source catalog loader'),
 item('Phase 7 공유 토폴로지 snapshot API',typeof T.exportConfig==='function'&&T.exportConfig()?.schemaVersion==='lsdc-twin-bom/1.0','existing shared schema'),
 item('Phase 7 토폴로지 비교 함수',typeof p7?.reconcileTopology==='function','explicit quantity reconciliation'),
 item('기본 시뮬레이터 설정',Number(c?.rackCount)>0&&Number(c?.rackPowerKw)>0,'rackCount='+c?.rackCount),
 item('운영 시나리오 API',typeof T.scenarioApply==='function','power/cooling/network/fiber-cut'),
 item('기존 다운로드 및 이벤트 로그',typeof T.exportCSV==='function'&&typeof T.addEventLog==='function','CSV + event log')
 ];
 report={schema:'lsdc/phase8-integration/1.0',at:new Date().toISOString(),passed:list.filter(x=>x.pass).length,failed:list.filter(x=>!x.pass).length,cases:list,limitations:['Physical Windows GPU/Android/iPhone not measured in CI','BOM exact orderable SKU, reach, FEC and MPO polarity unverified','MEP is conceptual AABB only, not Revit construction approval','GPU instancing is not added to original batched WebGL renderer']};
 return report;
}
function run(){
 const d=checks();$('p8-summary').textContent='기능 상태 '+d.passed+'/'+d.cases.length+' · 미통과 '+d.failed+' · 실제 기기/구매 검증은 별도';
 $('p8-cases').innerHTML=d.cases.map(x=>'<p class="'+(x.pass?'p58-pass':'p58-err')+'">'+(x.pass?'✓ ':'× ')+x.name+' · '+x.detail+'</p>').join('');return d;
}
function toggle(v){opened=v;panel.hidden=!v;btn.setAttribute('aria-expanded',String(v));if(v)run()}
btn.onclick=()=>toggle(!opened);$('p8-close').onclick=()=>toggle(false);$('p8-check').onclick=run;
$('p8-json').onclick=()=>{const s=JSON.stringify(report||run(),null,2),url=URL.createObjectURL(new Blob([s],{type:'application/json'})),a=document.createElement('a');a.href=url;a.download='LS_Datacenter_Phase8_integration.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)};
window.LS3D_PHASE8={run,open:()=>toggle(true),get report(){return report}};
})();