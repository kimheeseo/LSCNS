(function(root){
  'use strict';
  const number = (value,label,integer=false) => {
    if(value==null || typeof value==='boolean' || String(value).trim()==='') throw Error(label+'을 입력하세요.');
    const n=Number(value);
    if(!Number.isFinite(n)||n<0||(integer&&!Number.isInteger(n))) throw Error(label+'은 '+(integer?'0 이상의 정수':'0 이상의 수')+'여야 합니다.');
    return n;
  };
  function calculate(input){
    const modules=number(input.modules,'물리 광모듈 수',true),assemblies=number(input.assemblies,'능동 케이블 수',true);
    const moduleW=modules?number(input.moduleW,'광모듈 1개 소비전력'):0;
    const assemblyW=assemblies?number(input.assemblyW,'능동 케이블 1개 전체 소비전력'):0;
    const hours=number(input.hours,'운영 시간');
    const totalW=modules*moduleW+assemblies*assemblyW,kw=totalW/1000,kwh=kw*hours;
    if(![totalW,kw,kwh].every(Number.isFinite))throw Error('계산 범위를 초과했습니다.');
    return {passiveFiberW:0,modulePowerW:modules*moduleW,assemblyPowerW:assemblies*assemblyW,totalW,kw,kwh};
  }
  const model={calculate};
  if(typeof module!=='undefined'&&module.exports)module.exports=model;
  root.DCCablePower=model;
  if(typeof document==='undefined')return;
  const $=id=>document.getElementById(id),esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt=n=>Number(n).toLocaleString('ko-KR',{maximumFractionDigits:4});
  let design=null;
  function snapshot(){
    const host=$('cp-design');if(!host)return;
    if(!design?.usable || root.DCDesignStale){host.textContent='유효한 설계 계산 후 현재 구간별 설치 수량을 확인할 수 있습니다.';return;}
    const rows=Object.entries(design.optical||{}).filter(([,o])=>o.links>0);
    host.innerHTML='<h3>현재 설계의 설치 수량 · 예비 재고 제외</h3><div class="tableWrap"><table><thead><tr><th>구간</th><th>미디어</th><th>논리 링크</th><th>물리 광모듈</th><th>설치 트렁크</th><th>W/개</th></tr></thead><tbody>'+rows.map(([key,o])=>'<tr><td>'+esc(({server:'서버–Leaf',leafSpine:'Leaf–Spine',core:'Spine–Core'})[key]||key)+'</td><td>'+esc(o.profile?.media)+'</td><td>'+fmt(o.links)+'</td><td>'+fmt(o.installedModules||0)+'</td><td>'+fmt(o.trunk?.installedCableCount||0)+'</td><td>제품별 입력 필요</td></tr>').join('')+'</tbody></table></div><p>논리 링크 수와 물리 모듈 수는 다릅니다. Twin-port·breakout과 AOC 분기 구성은 실제 설치 부품 수로 확인하세요. 현재 설계전력 '+esc(design.summary?.totalItPowerKw)+' kW에 광모듈·능동 케이블 전력이 별도로 포함됐는지는 제품·장비의 전력 측정 경계를 확인해야 합니다.</p>';
  }
  function update(){
    try{
      const result=calculate({modules:$('cp-modules').value,moduleW:$('cp-module-w').value,assemblies:$('cp-assemblies').value,assemblyW:$('cp-assembly-w').value,hours:$('cp-hours').value});
      $('cp-result').textContent='수동 광섬유·광케이블 자체: 0 W · 광모듈 '+fmt(result.modulePowerW)+' W + 능동 케이블 '+fmt(result.assemblyPowerW)+' W = '+fmt(result.totalW)+' W ('+fmt(result.kw)+' kW) · 해당 시간 에너지 '+fmt(result.kwh)+' kWh';
    }catch(e){$('cp-result').textContent='산정 대기: '+e.message;}
  }
  function mount(){
    const target=$('view-power');if(!target)return false;if($('cable-power-panel'))return true;
    const panel=document.createElement('section');panel.id='cable-power-panel';panel.className='panel';
    panel.innerHTML=`<h2>LS전선 관점 · 광섬유·케이블과 연결 전력</h2>
      <p><b>수동 광섬유 1심, 광 패치코드·트렁크 1개 자체의 전기적 소비전력은 0 W입니다.</b> 통신 전력은 양 끝의 광모듈·NIC·스위치 또는 능동 케이블 내부 회로에서 발생합니다. 광신호가 운반하는 광파워(dBm/mW)와 장비의 전기 소비전력(W)을 구분합니다.</p>
      <div class="tableWrap"><table><thead><tr><th>LS 제품 / 연결 유형</th><th>자체 소비전력 기준</th><th>산정·검토 항목</th></tr></thead><tbody>
      <tr><td>광섬유·패치코드·트렁크·수동 카세트</td><td>0 W/심 또는 0 W/개 (수동 부품)</td><td>심수·길이로 전력을 곱하지 않음. 감쇠·접속손실·극성·굽힘·허용 도달거리 확인</td></tr>
      <tr><td>수동 DAC / Cat6·RJ45 데이터 케이블</td><td>케이블 능동 회로 전력 0 W</td><td>PHY 전력은 장비 측. 도체 저항 손실은 별도이며 PoE 급전 시 I²R 손실·발열 확인</td></tr>
      <tr><td>플러거블 광연결</td><td>실제 물리 광모듈별 W</td><td>양단 포함 설치 모듈 수 × 모듈별 전력. Twin-port 공유 모듈 중복 계산 금지</td></tr>
      <tr><td>AOC / AEC / ACC</td><td>제품별 능동 회로 전력</td><td>케이블 어셈블리 전체(양 끝·모든 분기) W 기준. 별도 광모듈 전력을 다시 더하지 않음</td></tr>
      <tr><td>CPO·ELS·PM fiber</td><td>PM fiber 자체 0 W</td><td>광엔진·EIC·외부 레이저·스위치 통합 전력 경계 확인. 보편적인 W/심 없음</td></tr>
      <tr><td>전력 케이블</td><td>전달 전력과 자체 손실 구분</td><td>전달 용량 kW·허용전류 A와 도체 손실 ΣI²R W를 별도 산정. 저항·온도·길이·도체 경로 필요</td></tr>
      </tbody></table></div>
      <p><b>현재 도구:</b> 광섬유·수동 케이블 1심/1개당 W 계수는 없습니다. 서버·스위치 등의 장비 설계전력을 집계하며, 광모듈·AOC/AEC의 제품별 W 계수는 현재 공개 결과에 별도로 제공되지 않습니다. <b>미제공을 능동 부품의 0 W로 해석하면 안 됩니다.</b></p>
      <div id="cp-design"></div>
      <h3>연결 전력·에너지 참고 계산</h3><p>제품 데이터시트의 최대 또는 측정 평균 전력을 같은 기준으로 입력하세요. W는 소비전력, kWh는 일정 시간 사용한 에너지입니다. 이 계산은 기존 IT/PUE 결과에 자동 합산되지 않습니다.</p>
      <div class="fields sub">
      <label><span>설치 물리 광모듈 수 (양단 합계)</span><input id="cp-modules" data-reference-only="true" type="number" min="0" step="1" value="0"></label>
      <label><span>광모듈 1개 소비전력 (W)</span><input id="cp-module-w" data-reference-only="true" type="number" min="0" step="any" placeholder="제품별 입력 · 미확인"></label>
      <label><span>설치 능동 케이블 어셈블리 수</span><input id="cp-assemblies" data-reference-only="true" type="number" min="0" step="1" value="0"></label>
      <label><span>능동 케이블 1개 전체 소비전력 (W)</span><input id="cp-assembly-w" data-reference-only="true" type="number" min="0" step="any" placeholder="양 끝·분기 합계"></label>
      <label><span>운영 시간 (h)</span><input id="cp-hours" data-reference-only="true" type="number" min="0" step="any" value="24"></label></div>
      <p><button type="button" id="cp-example">400G DR4 단일 포트 예제: 2개 × 최대 9 W</button> <button type="button" id="cp-copy">현재 설계 설치 광모듈 수 가져오기</button></p>
      <p id="cp-result" role="status"></p><p>수식: P(W) = Σ[설치 물리 모듈 수 × 모듈 W] + Σ[능동 어셈블리 수 × 전체 W]; E(kWh) = P(W) ÷ 1,000 × 시간(h). 서로 다른 SKU는 각각 계산해 합산하세요. 최대 W를 넣으면 최대 전력 기준 예산이며 실제 평균 에너지 측정값은 아닙니다.</p>
      <p>예: 동일한 400G DR4 단일 포트 모듈 2개가 각각 최대 9 W이면 연결 광모듈 예산은 18 W, 24시간 최대 부하 가정은 0.432 kWh입니다. DR4의 활성 8심에 균등 배분하면 2.25 W/활성심이라는 <b>연결 장비 전력의 배분 지표</b>가 되지만, 광섬유가 2.25 W를 소비한다는 뜻은 아닙니다. 미사용·예비 심에 같은 계수를 곱하지 않습니다.</p>
      <p><b>LS 제품의 검토 가치:</b> 낮은 채널 손실과 적합한 인터페이스가 더 낮은 전력의 광연결 구성 선택에 기여할 수 있습니다. 케이블 감쇠 개선 dB를 전력 절감 W로 바로 환산할 수는 없습니다. 동일 속도·거리·BER/FEC·장비 조건의 실제 제품 비교로 절감량을 확인하세요.</p>
      <p>공식 예제 출처: <a href="https://networking-docs.nvidia.com/mms4x00ns400/specifications" target="_blank" rel="noopener">NVIDIA MMS4X00-NS400 · 400G 최대 9 W / DR4 활성 8심</a> · <a href="https://www.nvidia.com/en-au/networking/ethernet/direct-attach-copper-cables/" target="_blank" rel="noopener">NVIDIA 수동 DAC 전력 설명</a>. 확인일 2026-10-11. 위 9 W는 전체 제품에 적용하는 기본값이 아닙니다.</p>`;
    target.prepend(panel);
    panel.querySelectorAll('input').forEach(el=>el.addEventListener('input',update));
    $('cp-example').onclick=()=>{$('cp-modules').value='2';$('cp-module-w').value='9';$('cp-assemblies').value='0';$('cp-assembly-w').value='';$('cp-hours').value='24';update();};
    $('cp-copy').onclick=()=>{
      if(!design?.usable||root.DCDesignStale){$('cp-result').textContent='현재 입력으로 유효한 설계를 먼저 계산하세요.';return;}
      $('cp-modules').value=Object.values(design.optical||{}).reduce((n,o)=>n+(Number(o.installedModules)||0),0);
      $('cp-module-w').value='';update();
    };
    design=root.DCDesign;snapshot();update();return true;
  }
  document.addEventListener('dc:design',e=>{design=e.detail;snapshot();});
  document.addEventListener('dc:stale',()=>{design=null;snapshot();});
  function start(){if(mount())return;let tries=0;const timer=setInterval(()=>{if(mount()||++tries>30)clearInterval(timer);},100);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});else start();
})(typeof window!=='undefined'?window:globalThis);
