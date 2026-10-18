/* LS Cable perspective on the current engineering result. No product SKU is inferred. */
(function(root){
  'use strict';
  const labels={server:'서버–Leaf',leafSpine:'Leaf–Spine',core:'Spine–Core'};
  const safeNumber=value=>Number.isFinite(Number(value))?Number(value):0;
  function summarize(design){
    if(!design?.usable||!design.optical||!design.summary||design.input?.scenario==='colo')return null;
    const rows=Object.entries(design.optical).filter(([,o])=>safeNumber(o?.links)>0).map(([segment,o])=>{
      const profile=o.profile||{},trunk=o.trunk||{};
      const patch=design.bom?.find(item=>item.segment===segment&&item.item?.includes('optical patch leads'));
      return {segment,label:labels[segment]||segment,links:safeNumber(o.links),lengthM:safeNumber(o.lengthM),
        cableMode:design.input?.[segment==='server'?'serverCabling':segment==='leafSpine'?'leafSpineCabling':'coreCabling']||'structured',
        installedTrunks:safeNumber(trunk.installedCableCount),purchasedTrunks:safeNumber(trunk.cableCount),
        installedPatch:safeNumber(patch?.installedQty),purchasedPatch:safeNumber(patch?.qty),
        fiber:profile.fiberType||'',activeFibers:safeNumber(profile.fibers),channelFibers:safeNumber(profile.installedChannelFibers),
        trunkFibers:safeNumber(profile.trunkFiberCount),connector:profile.connector||'',polish:profile.polish||'',
        opticSku:profile.opticSku||'',lossPass:o.loss?.pass===true,marginDb:safeNumber(o.loss?.marginDb),
        unusedFibers:safeNumber(profile.unusedTrunkFibers)};
    });
    const issues=[];
    for(const row of rows){
      if(!row.opticSku)issues.push(row.label+': 광모듈 SKU 미선정 · 커넥터/도달거리 RFQ');
      if(row.marginDb<=0)issues.push(row.label+': 광손실 여유 예산 0 dB · 프로젝트 기준 입력');
      if(!row.lossPass)issues.push(row.label+': 도달거리 또는 광손실 조건 실패');
      if(row.unusedFibers)issues.push(row.label+': '+row.trunkFibers+'F 트렁크의 '+row.unusedFibers+'F는 완전한 채널로 종단 불가');
      if(row.connector.includes('MPO')&&design.input?.mpoEndGender==='review')issues.push(row.label+': MPO 종단 핀 유무와 카세트 극성 확인');
    }
    for(const error of design.errors||[])if(/optic|fiber|polish|polarity|connector|reach|cable|trunk|media/i.test(error))issues.push(error);
    return {rows,issues:[...new Set(issues)],status:design.status||'REVIEW',
      totalInstalledTrunks:rows.reduce((n,row)=>n+row.installedTrunks,0),
      totalPurchasedTrunks:rows.reduce((n,row)=>n+row.purchasedTrunks,0),
      totalInstalledPatch:rows.reduce((n,row)=>n+row.installedPatch,0),
      totalRacks:safeNumber(design.summary.totalRacks),itPowerKw:safeNumber(design.summary.totalItPowerKw),
      coolingMode:design.cooling?.modeLabel||'',liquidFlowLMin:safeNumber(design.cooling?.flowLMin)};
  }
  const api={summarize};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;
  root.LSCablePerspective=api;
  if(typeof document==='undefined')return;
  const $=id=>document.getElementById(id);
  const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt=value=>safeNumber(value).toLocaleString('ko-KR',{maximumFractionDigits:1});
  function render(design){
    const host=$('lsCablePerspectiveContent');if(!host)return;
    if(root.DCDesignStale){host.textContent='입력값이 변경되었습니다. 새 설계 계산이 완료되면 요약이 갱신됩니다.';return;}
    const report=summarize(design);
    if(!report){host.textContent='현재 설계 계산 결과가 없습니다. 요구조건에서 설계 계산을 실행하세요.';return;}
    host.innerHTML=`<p class="lsPerspectiveNote">현재 계산 ${esc(report.status)} · 설치 수량과 구매 수량(예비 포함)을 구분했습니다. ${report.status.includes('ERROR')?'설계 조건 충돌이 있으므로 아래 수량은 발주에 사용할 수 없습니다. ':''}LS전선의 확정 제품·품목 코드를 뜻하지 않습니다.</p>
      <div class="lsPerspectiveCards">
        <article><h3>광배선 수요</h3><strong>트렁크 ${fmt(report.totalInstalledTrunks)}개 설치</strong><p>구매 ${fmt(report.totalPurchasedTrunks)}개 · 종단 패치리드 ${fmt(report.totalInstalledPatch)}개 설치 계획</p></article>
        <article><h3>랙·전력 배선</h3><strong>${fmt(report.totalRacks)}랙 · IT ${fmt(report.itPowerKw)} kW</strong><p>급전 전압·상수·A/B 부하·거리·포설 방식이 정해져야 전력 케이블 단면과 허용전류를 산정할 수 있습니다.</p></article>
        <article><h3>냉각 설비 연계</h3><strong>${esc(report.coolingMode||'검토 필요')}</strong><p>${report.liquidFlowLMin>0?'냉각수 계획 유량 '+fmt(report.liquidFlowLMin)+' L/min · CDU/펌프 급전 및 계장 케이블 확인':'공랭 및 보조 설비의 전원·제어 배선은 별도 검토'}</p></article>
      </div>
      <h3>광케이블·커넥터 우선 검토</h3>
      <div class="tableWrap"><table><thead><tr><th>구간</th><th>트렁크 설치 / 구매</th><th>규격·길이</th><th>활성 / 채널 / 트렁크</th><th>커넥터·모듈</th></tr></thead><tbody>${report.rows.map(row=>`<tr><th scope="row">${esc(row.label)} · ${fmt(row.links)}링크</th><td>${row.cableMode==='p2p'?'직결 · 트렁크 없음':fmt(row.installedTrunks)+' / '+fmt(row.purchasedTrunks)}</td><td>${esc(row.fiber||'RFQ')} · ${fmt(row.lengthM)} m</td><td>${fmt(row.activeFibers)} / ${fmt(row.channelFibers)} / ${fmt(row.trunkFibers)} F</td><td>${esc(row.connector||'RFQ')}/${esc(row.polish||'?')}<br><small>${esc(row.opticSku||'일반 가정 · SKU RFQ')}</small></td></tr>`).join('')}</tbody></table></div>
      <div class="lsPerspectiveChecks"><h3>LS전선 품목 선정 전 확인</h3><ul>${(report.issues.length?report.issues:['선택한 광모듈의 상대편 SKU, FEC, 카세트 극성 및 현장 광손실 측정 확인']).map(issue=>`<li>${esc(issue)}</li>`).join('')}</ul></div>
      <p class="lsPerspectiveFoot"><b>품목 검토 순서:</b> ① 광섬유·트렁크·패치코드의 심수와 길이 → MPO/LC 연마·극성·핀·삽입손실 → 피복·난연 등급·굴곡반경·트레이 경로 → LS전선 실제 SKU 매핑. ② 전력·제어 케이블은 부하 전류, A/B 급전, 전압강하·단락·포설 감쇄를 별도로 검토하세요. 이 화면은 제조 가능 품목이나 납품 적합성을 자동 인증하지 않습니다.</p>`;
  }
  function mount(){
    const button=$('lsCablePerspectiveToggle'),panel=$('lsCablePerspective');if(!button||!panel)return;
    function toggle(open){panel.hidden=!open;button.setAttribute('aria-expanded',String(open));if(open){render(root.DCDesign);panel.scrollIntoView({behavior:'smooth',block:'start'});}}
    button.addEventListener('click',()=>toggle(panel.hidden));
    $('lsCablePerspectiveClose')?.addEventListener('click',()=>{toggle(false);button.focus();});
    document.addEventListener('dc:design',event=>render(event.detail));
    document.addEventListener('dc:stale',()=>{const host=$('lsCablePerspectiveContent');if(host)host.textContent='입력값이 변경되었습니다. 새 설계 계산이 완료되면 요약이 갱신됩니다.';});
    render(root.DCDesign);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',mount,{once:true});else mount();
})(typeof window!=='undefined'?window:globalThis);
