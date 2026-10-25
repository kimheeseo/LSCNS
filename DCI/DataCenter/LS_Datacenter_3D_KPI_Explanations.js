/* Interactive live KPI explanations for the LS DataCenter 3D digital twin · 2026-10-26 */
(function(){
'use strict';
const $=id=>document.getElementById(id),cfg=()=>window.LS3D_CONFIG,twin=()=>window.__LS3D_TEST__;
if(!document.querySelector('[data-kpi-explain]'))return;
const style=document.createElement('style');
style.textContent=`
.metric[data-kpi-explain]{cursor:pointer;transition:border-color .15s,background .15s}
.metric[data-kpi-explain]:hover,.metric[data-kpi-explain]:focus-visible{border-color:#5ddcca;background:#17334a;outline:2px solid #58ccb788;outline-offset:1px}
.metric[data-kpi-explain] .mt:after{content:'ⓘ';color:#77d8dc;font-size:13px;margin-left:4px}
#ls3dKpiExplainer[hidden]{display:none!important}
#ls3dKpiExplainer{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;z-index:150;background:#03101de0;padding:15px}
#ls3dKpiExplainer .kpi-dialog{width:min(570px,100%);max-height:86vh;overflow:auto;background:#10283c;border:1px solid #477b8b;border-radius:14px;padding:20px;box-shadow:0 16px 60px #0009;color:#dfedf6}
#ls3dKpiExplainer header{display:flex;align-items:center;justify-content:space-between;gap:12px}
#ls3dKpiExplainer h2{font-size:18px;margin:0}
#ls3dKpiExplainer h3{font-size:13px;color:#93e8d8;margin:14px 0 6px}
#ls3dKpiExplainer button{cursor:pointer;border:1px solid #58899a;border-radius:7px;background:#183b51;color:#f0fcff;padding:8px 12px}
#ls3dKpiExplainer .kpi-content{font-size:12px;line-height:1.75}
#ls3dKpiExplainer p{margin:8px 0}
#ls3dKpiExplainer code{color:#ffe0a8;white-space:normal;word-break:break-word}
#ls3dKpiExplainer table{width:100%;border-collapse:collapse;margin:8px 0}
#ls3dKpiExplainer th,#ls3dKpiExplainer td{border-bottom:1px solid #325269;padding:6px;text-align:left}
#ls3dKpiExplainer th{width:38%;color:#bbdce9}
#ls3dKpiExplainer footer{font-size:10px;line-height:1.7;color:#a7bccc;margin-top:13px}
@media(max-width:600px){#ls3dKpiExplainer .kpi-dialog{padding:14px}#ls3dKpiExplainer th{width:43%}}
`;
document.head.appendChild(style);
const overlay=document.createElement('div');
overlay.id='ls3dKpiExplainer';overlay.hidden=true;
overlay.innerHTML='<section class="kpi-dialog" role="dialog" aria-modal="true" aria-labelledby="ls3dKpiTitle"><header><h2 id="ls3dKpiTitle">계산 근거</h2><button type="button" id="ls3dKpiClose">닫기 ✕</button></header><div class="kpi-content" id="ls3dKpiContent"></div><footer>현재 설정값을 이용한 교육·설계 검토용 추정치이며 실제 계측값이나 시공 승인값이 아닙니다.</footer></section>';
document.body.appendChild(overlay);
const title=$('ls3dKpiTitle'),body=$('ls3dKpiContent'),close=$('ls3dKpiClose');
const fmt=n=>Number(n||0).toLocaleString('ko-KR',{maximumFractionDigits:2});
const row=(label,value)=>'<tr><th>'+label+'</th><td>'+value+'</td></tr>';
let previousFocus=null;
function open(kind){
  const c=cfg(),t=twin();if(!c||!t||typeof t.metrics!=='function')return;
  const m=t.metrics(),sc=$('scenarioSelect')?.value||'normal',elapsed=t.scenarioElapsed?.()||0;
  const rackAndOptical=m.rackLoad+m.opticalPower,scenarioMultiplier=rackAndOptical?m.load/rackAndOptical:1;
  const calculatedIT=(m.coolingPower||0)*(m.effectiveCop||0),baseFacility=calculatedIT+m.coolingPower+m.upsLoss+m.distributionLoss+m.otherPower;
  let content='';
  if(kind==='load'){
    title.textContent='IT 부하 · 계산식';
    content='<h3>장비 소비전력 모델</h3><p><code>IT 부하 = (랙 수 × 랙당 전력 + 광 포트 전력 + 네트워크 스위치 전력) × 시나리오 보정</code></p><table>'+
      row('랙 IT 부하',fmt(c.rackCount)+' × '+fmt(c.rackPowerKw)+' kW = '+fmt(m.rackLoad)+' kW')+
      row('광·네트워크',fmt(c.portCount)+' 포트 · '+(c.mode==='cpo'?'CPO':'Pluggable')+' · '+fmt(m.opticalPower)+' kW')+
      row('시나리오 보정',fmt(scenarioMultiplier)+'배 · '+sc)+
      row('현재 IT 부하',fmt(m.load)+' kW')+'</table><p>3D 랙은 대표 배치이며, 실제 계산에는 설정 메뉴의 랙 수와 랙당 IT 전력 값을 사용합니다.</p>';
  }else if(kind==='facility'){
    title.textContent='시설 전력 · 계산식';
    content='<h3>전체 시설 소비전력 모델</h3><p><code>시설 전력 = IT 부하 + 냉각 전력 + UPS 손실 + 배전 손실 + 기타 전력</code></p><table>'+
      row('계산 기준 IT 부하',fmt(calculatedIT)+' kW')+
      row('냉각 전력',fmt(calculatedIT)+' kW ÷ 유효 COP '+fmt(m.effectiveCop)+' = '+fmt(m.coolingPower)+' kW')+
      row('UPS 손실',fmt(m.upsLoss)+' kW · UPS 효율 '+fmt(c.upsEfficiencyPct)+'%')+
      row('배전 손실',fmt(m.distributionLoss)+' kW · 손실률 '+fmt(c.distributionLossPct)+'%')+
      row('기타 전력',fmt(m.otherPower)+' kW · 비율 '+fmt(c.otherPct)+'%')+
      row('시나리오 전력 조정',fmt(m.facility-baseFacility)+' kW · '+sc)+
      row('현재 시설 전력',fmt(m.facility)+' kW')+
      row('실제 계산 PUE',fmt(m.pue)+' = 시설 전력 ÷ IT 부하')+'</table><p>목표 PUE 설정과 구분하여, 현재 전력 구성에서 계산한 PUE를 사용합니다. 유효 COP는 냉각 방식·외기 온도·부하율 가정을 반영합니다.</p>';
  }else{
    title.textContent='냉각 상태 · 온도 계산식';
    const cap=Math.max(1,c.coolingKw),ratio=calculatedIT/cap,baseline=c.baselineRackInletC??22;
    const loadExtra=Math.max(0,ratio-.7)*4,outdoorExtra=Math.max(0,c.outdoorC-24)*.08;
    const scenarioExtra=m.temperature-baseline-loadExtra-outdoorExtra;
    content='<h3>현재 카드의 지표는 냉각 여유율이 아닌 랙 입구 추정 온도(°C)입니다.</h3><p><code>추정 온도 = 기준 온도 + max(0, IT 부하 / 냉각 용량 − 0.7) × 4 + max(0, 외기 − 24) × 0.08 + 시나리오 보정</code></p><table>'+
      row('기준 온도',fmt(baseline)+' °C')+
      row('냉각 설비 용량',fmt(c.coolingKw)+' kW')+
      row('계산 기준 IT 부하 / 냉각 용량',fmt(calculatedIT)+' / '+fmt(c.coolingKw)+' = '+fmt(ratio*100)+'%')+
      row('부하 온도 가산',fmt(loadExtra)+' °C')+
      row('외기 온도 가산',fmt(outdoorExtra)+' °C · 외기 '+fmt(c.outdoorC)+' °C')+
      row('시나리오 온도 가산',fmt(scenarioExtra)+' °C · '+sc)+
      row('현재 추정 온도',fmt(m.temperature)+' °C')+'</table><p>냉각 용량보다 열부하가 큰지는 부하율이 100%를 넘는지로 확인합니다. 이는 CFD 해석 또는 센서 실측 온도가 아닙니다.</p>';
  }
  body.innerHTML=content;
  previousFocus=document.activeElement;overlay.hidden=false;close.focus();
}
function dismiss(){overlay.hidden=true;previousFocus?.focus?.()}
document.querySelectorAll('[data-kpi-explain]').forEach(card=>{
  card.addEventListener('click',()=>open(card.dataset.kpiExplain));
  card.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();open(card.dataset.kpiExplain)}});
});
close.addEventListener('click',dismiss);
overlay.addEventListener('click',e=>{if(e.target===overlay)dismiss()});
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!overlay.hidden){e.stopPropagation();dismiss()}});
window.LS3D_KPI_EXPLAINER={open,dismiss,get isOpen(){return !overlay.hidden}};
})();
