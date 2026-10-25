/* Phase 7 v4.7.0 · shared Twin packet BOM summary and evidence-aware validator. */
(function(){
'use strict';
const T=window.__LS3D_TEST__,view=document.getElementById('viewport'),B=window.LSDCPhase7BOM,V=window.LSDCPhase7Compatibility;
if(!T||!view||!B||!V)return;
const $=id=>document.getElementById(id),esc=s=>String(s??'').replace(/[&<>"']/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[x]));
let last=null,openState=false,signature=null;
function packetSignature(p){const {createdAt,...stable}=p;return JSON.stringify(stable)}
const panel=document.createElement('aside');panel.id='phase7-panel';panel.className='phase58-panel';panel.hidden=true;
panel.innerHTML=[
 '<header><b>PHASE 7 · 공통 BOM 산정</b><button id="p7-close" type="button">✕</button></header>',
 '<p class="p58-note">현재 3D 설계의 lsdc-twin-bom/1.0 스냅샷을 공통 계산 엔진에 전달합니다. 설계 물량과 화면에 그려진 대표 객체 수는 별도로 유지합니다.</p>',
 '<div class="p58-actions"><button id="p7-sync" type="button">현재 설계 다시 계산</button><button id="p7-json" type="button">검토 JSON</button><button id="p7-to-bom" type="button">전체 설계 → BOM 도구</button></div>',
 '<p id="p7-status" aria-live="polite">실행 전</p><div id="p7-summary"></div><div id="p7-table"></div>',
 '<small class="p58-foot">BOM 수량은 입력한 랙·토폴로지·링크 설정에 대한 개념 산정입니다. 실제 구매 적합성은 별도 검증이 필요합니다.</small>'
].join('');document.body.appendChild(panel);
const btn=document.createElement('button');btn.id='phase7-open';btn.className='phase58-launch p7';btn.textContent='▣ 공통 BOM';btn.type='button';view.appendChild(btn);
function snapshot(){if(typeof T.exportConfig!=='function')throw Error('공유 설계 스키마 API T.exportConfig()를 찾을 수 없습니다.');return T.exportConfig()}
function estimate(packet){const p=packet||snapshot(),engine=B.calculate(p),compatibility=V.validate(p);return {...engine,compatibility,packet:p}}
function reconcileTopology(audit,packet){const fresh=B.calculate(packet||snapshot()),a=audit?.counts||{},b=fresh.counts;const keys=['rackCount','gpuServerCount','gpuCount','nicCount','leafCount','spineCount','coreCount','opticalLinkCount','totalOpticalLengthM'];const differences=keys.filter(k=>a[k]!==b[k]).map(k=>({key:k,audit:a[k]??null,packet:b[k]??null}));return {matched:differences.length===0,status:differences.length?'different':'matched',differences,source:'same lsdc-twin-bom/1.0 packet'};}
function render(){
 try{
  last=estimate();signature=packetSignature(last.packet);const c=last.counts,v=last.compatibility;
  $('p7-status').textContent='공통 BOM 계산 완료 · 수량 및 산정 근거 표시';
  $('p7-summary').innerHTML='<strong>랙 '+fmt(c.rackCount)+' · GPU 서버 '+fmt(c.gpuServerCount)+' · GPU '+fmt(c.gpuCount)+' · NIC '+fmt(c.nicCount)+'</strong><p class="p58-note">Leaf '+fmt(c.leafCount)+' · Spine '+fmt(c.spineCount)+' · Core '+fmt(c.coreCount)+' · 활성 광 링크 '+fmt(c.opticalLinkCount)+' · 광 길이 '+fmt(c.totalOpticalLengthM)+' m</p>'
   +last.assumptions.map(x=>'<p class="p58-note">가정: '+esc(x)+'</p>').join('');
  $('p7-table').innerHTML='<table><thead><tr><th>분류</th><th>BOM 품목</th><th>수량</th><th>산정 근거</th></tr></thead><tbody>'+last.rows.map(x=>'<tr><td>'+esc(x.category)+'</td><td>'+esc(x.part)+'</td><td>'+fmt(x.quantity)+' '+esc(x.unit)+'</td><td>'+esc(x.basis)+' · '+esc(x.source)+'</td></tr>').join('')+'</tbody></table>';

 }catch(e){last=null;$('p7-status').textContent='계산 불가 · '+e.message;$('p7-table').innerHTML='';$('p7-summary').innerHTML=''}
 return last;
}
function fmt(v){return v==null?'미확정':Number(v).toLocaleString('ko-KR',{maximumFractionDigits:2})}
function toggle(v){openState=v;panel.hidden=!v;btn.setAttribute('aria-expanded',String(v));if(v)render()}
btn.onclick=()=>toggle(!openState);$('p7-close').onclick=()=>toggle(false);$('p7-sync').onclick=render;
$('p7-json').onclick=()=>{const result=last||render();if(!result)return;const data={schema:'lsdc/phase7-review/1.0',bom:{...result,packet:undefined,compatibility:undefined}};const url=URL.createObjectURL(new Blob([JSON.stringify(data,null,2)],{type:'application/json'})),a=document.createElement('a');a.href=url;a.download='LS_Datacenter_Phase7_review.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1200)};
$('p7-to-bom').onclick=()=>{const out=$('bomOut');if(out){out.click();$('p7-status').textContent='전체 공유 설계를 BOM 도구로 전달했습니다.'}else $('p7-status').textContent='기존 BOM 공유 버튼을 찾지 못했습니다.'};
setInterval(()=>{if(openState&&!document.hidden){try{if(packetSignature(snapshot())!==signature)render()}catch(e){render()}}},400);
window.LS3D_PHASE7={open:()=>toggle(true),load:async()=>({source:'twin-bom-bridge.js',schema:'lsdc-twin-bom/1.0'}),estimate,reconcileTopology,get audit(){return last},get engine(){return B},get validator(){return V}};
})();
