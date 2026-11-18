/* Time-axis charts & click-to-view simulation assumptions.
 * This presentation is driven by LS3D_TIMELINE_MODEL, not measured sensors.
 */
(function(){
 'use strict';
 const model=window.LS3D_TIMELINE_MODEL;
 if(!model)return;
 const $=id=>document.getElementById(id);
 const parent=$('scenarioStepPanel');
 if(!parent)return;
 const style=document.createElement('style');
 style.textContent=`
#ls3dFailureTimeline{margin:10px 0 14px;padding:11px 13px;border:1px solid #36546d;border-radius:10px;background:#0d2133;color:#e4f5ff;font:12px/1.5 system-ui}
#ls3dFailureTimeline[hidden]{display:none}
#ls3dFailureTimeline .tline-top{display:flex;align-items:center;justify-content:space-between;gap:7px;flex-wrap:wrap}
#ls3dFailureTimeline .tline-top strong{font-size:13px}
#ls3dFailureTimeline .tline-edu{border:1px solid #668390;border-radius:40px;padding:2px 8px;background:#192f40;color:#d1e4ef;font-size:10px}
#ls3dFailureTimeline .tline-status{display:flex;gap:11px;flex-wrap:wrap;margin-top:5px;color:#c1dae8;font-size:11px}
#ls3dFailureTimeline .tline-chart-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin:9px 0}
#ls3dFailureTimeline figure{margin:0;min-width:0;border:1px solid #2b5166;border-radius:8px;padding:7px 6px 3px;background:#112b3e}
#ls3dFailureTimeline figcaption{color:#cde9f4;font-size:11px;font-weight:700}
#ls3dFailureTimeline svg{display:block;width:100%;height:auto;aspect-ratio:460 / 148}
#ls3dFailureTimeline .tline-reading{color:#b5d3e1;margin:5px 0 0;font-size:11px}
#ls3dFailureTimeline .tline-reading strong{color:#f7fbfd}
#ls3dFailureTimeline details{margin-top:8px;border-top:1px solid #3b5f71;padding-top:7px;font-size:11px}
#ls3dFailureTimeline summary{cursor:pointer;color:#a9efdf;font-weight:750;padding:5px 0}
#ls3dFailureTimeline .tline-assumptions{width:100%;border-collapse:collapse;margin-top:5px}
#ls3dFailureTimeline .tline-assumptions td{padding:6px 7px;border-top:1px solid #355264;vertical-align:top}
#ls3dFailureTimeline .tline-assumptions td:first-child{width:33%;color:#b9e9e4}
#ls3dFailureTimeline .tline-disclaimer{font-size:10px;color:#9eb9c8;margin:6px 0 0}
#comparisonResults .tline-compare{width:100%;border-collapse:collapse;text-align:left;font:11px/1.45 system-ui}
#comparisonResults .tline-compare th,#comparisonResults .tline-compare td{padding:7px 8px;border-bottom:1px solid #35556c;white-space:nowrap}
#comparisonResults .tline-compare thead th{font-weight:800;color:#add9e5}
#comparisonResults .tline-compare tr.is-active{background:#21455c}
#comparisonResults .tline-compare-wrap{width:100%;overflow-x:auto}
@media(max-width:690px){#ls3dFailureTimeline .tline-chart-grid{grid-template-columns:1fr}#ls3dFailureTimeline{padding:9px}#ls3dFailureTimeline .tline-status{gap:5px}}
`;
 document.head.appendChild(style);
 const root=document.createElement('section');
 root.id='ls3dFailureTimeline';
 root.setAttribute('aria-label','장애 시나리오 시간축 시뮬레이션');
 root.innerHTML=`
 <div class="tline-top"><strong>장애 시간축 시뮬레이션 · 초~분</strong>
  <span class="tline-edu" title="시간축 수치는 실측 데이터가 아닌 교육용 가정입니다">교육용 가정</span></div>
 <div class="tline-status" aria-live="polite"><span id="tlineNow"></span><span id="tlineWorst"></span><span id="tlineTime"></span></div>
 <div class="tline-chart-grid">
  <figure><figcaption id="tlineGraphTitleA"></figcaption><svg id="tlineGraphA" viewBox="0 0 460 148" role="img"></svg></figure>
  <figure><figcaption id="tlineGraphTitleB"></figcaption><svg id="tlineGraphB" viewBox="0 0 460 148" role="img"></svg></figure>
 </div>
 <p class="tline-reading" id="tlineReading"></p>
 <details id="tlineAssumptionDisclosure"><summary>교육용 가정 · 적용 수치와 계산식 확인</summary>
  <table class="tline-assumptions"><tbody id="tlineAssumptionRows"></tbody></table>
  <p class="tline-disclaimer">본 곡선은 사고 대응 절차나 실측 데이터가 아닌 학습용 모델이며, 보호계전·배터리·화재·냉각 제어를 대신하지 않습니다.</p>
 </details>`;
 parent.insertAdjacentElement('afterend',root);
 const allowed=new Set(['power','ups-battery','cooling','dlc-leak','fire','gpu-surge']);
 const fmtSec=t=>t==null?'도달하지 않음':t>=60?(t/60).toFixed(1)+'분':Math.round(t)+'초';
 const fmt=(v,d=1)=>Number(v).toFixed(d);
 const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
 function chartSpec(scenario,sum){
  const p=model.parameters(window.LS3D_CONFIG||{}),maxRacks=p.rackCount;
  const pct=row=>row.serviceFraction*100;
  const specs={
   power:[
    {title:'UPS 잔여시간 · 분',y:r=>r.upsRemaining,base:0,max:Math.max(1,p.batterySec/60),unit:'분'},
    {title:'IT 부하 공급률 · %',y:pct,base:0,max:105,unit:'%'}
   ],
   'ups-battery':[
    {title:'UPS 잔여시간 · 분',y:r=>r.upsRemaining,base:0,max:Math.max(1,p.batterySec/60),unit:'분'},
    {title:'IT 부하 공급률 · %',y:pct,base:0,max:105,unit:'%'}
   ],
   cooling:[
    {title:'랙 입구 온도 · °C',y:r=>r.temperature,base:Math.min(20,p.tempBaseline-3),
     max:Math.max(46,sum.maxTemperature+3),unit:'°C',thresholds:[27,32,41]},
    {title:'서비스 공급률 · %',y:pct,base:0,max:105,unit:'%'}
   ],
   'dlc-leak':[
    {title:'영향 랙 수 · 대',y:r=>r.affectedRacks,base:0,
     max:Math.max(2,maxRacks*.3,sum.maxAffectedRacks+2),unit:'랙'},
    {title:'영향 구간 IT 부하 공급률 · %',y:pct,base:0,max:105,unit:'%'}
   ],
   fire:[
    {title:'격리 구역 영향 랙 수 · 대',y:r=>r.affectedRacks,base:0,max:Math.max(2,sum.maxAffectedRacks+2),unit:'랙'},
    {title:'전체 IT 부하 공급률 · %',y:pct,base:0,max:105,unit:'%'}
   ],
   'gpu-surge':[
    {title:'GPU IT 요구전력/가용전력 · kW',y:r=>r.requestedKw,other:r=>r.availableCapacity,
     base:0,max:Math.max(sum.maxDemandKw,p.itBaseKw*1.7)*1.08,unit:'kW'},
    {title:'랙 입구 온도 · °C',y:r=>r.temperature,base:Math.min(20,p.tempBaseline-2),
     max:Math.max(37,sum.maxTemperature+3),unit:'°C',thresholds:[27,32]}
   ]
  };
  return specs[scenario];
 }
 function svgMarkup(spec,rows,now,horizon,label){
  const W=460,H=148,x0=43,x1=448,y0=11,y1=122;
  const low=spec.base,high=spec.max;
  const x=t=>x0+(x1-x0)*t/Math.max(1,horizon);
  const y=v=>y1-(y1-y0)*(Math.max(low,Math.min(high,v))-low)/Math.max(.01,high-low);
  const path=fn=>rows.map((r,i)=>(i?'L':'M')+fmt(x(r.t),2)+' '+fmt(y(fn(r)),2)).join(' ');
  const a=path(spec.y),b=spec.other?path(spec.other):'';
  const grid=[0,.5,1].map(k=>{const yy=y0+(y1-y0)*k,v=high-(high-low)*k;
   return '<line x1="'+x0+'" y1="'+yy+'" x2="'+x1+'" y2="'+yy+'" stroke="#355269" stroke-width="1"/>'+
    '<text x="3" y="'+(yy+4)+'" font-size="10" fill="#a9c0cf">'+fmt(v,high-low>100?0:1)+'</text>';
  }).join('');
  const marks=(spec.thresholds||[]).filter(v=>v>=low&&v<=high).map(v=>'<line x1="'+x0+'" y1="'+y(v)+'" x2="'+x1+'" y2="'+y(v)+
   '" stroke="#eebd73" stroke-dasharray="4 4" stroke-opacity=".8"/><text x="'+(x1-24)+'" y="'+(y(v)-3)+
   '" font-size="9" fill="#f1c581">'+v+'°</text>').join('');
  const nowRow=model.point($('scenarioSelect')?.value||'normal',$('redundancySelect')?.value||'N',now,window.LS3D_CONFIG||{});
  const currentValue=spec.y(nowRow),cx=x(Math.min(horizon,now)),cy=y(currentValue);
  const clock=t=>t>=60?Math.round(t/60)+'m':Math.round(t)+'s';
  return '<title>'+esc(label+' · 시간에 따른 교육용 추정값')+'</title>'+grid+marks+
    '<path d="'+a+'" fill="none" stroke="#68dcce" stroke-width="2.8" stroke-linejoin="round" />'+
    (b?'<path d="'+b+'" fill="none" stroke="#e9b96e" stroke-width="2.1" stroke-dasharray="6 4"/>':'')+
    '<line x1="'+cx+'" y1="'+y0+'" x2="'+cx+'" y2="'+y1+'" stroke="#f7f9ff" stroke-width="1.2" stroke-dasharray="3 4"/>'+
    '<circle cx="'+cx+'" cy="'+cy+'" r="3.8" fill="#f8f9ff" stroke="#26617b" stroke-width="1.5"/>'+
    '<text x="'+x0+'" y="140" font-size="10" fill="#a4c6d4">0s</text>'+
    '<text x="'+(x1-32)+'" y="140" font-size="10" fill="#a4c6d4">'+clock(horizon)+'</text>';
 }
 let cachedKey='',cached=null,lastDraw=0;
 const cfgKeys=['rackCount','rackPowerKw','portCount','mode','plugKw','plugW','cpoKw','cpoW','dlc','batteryMin',
  'upsKw','generatorKw','coolingKw','baselineRackInletC','outdoorC','gpuRatio'];
 function redraw(force=false){
  const selected=$('scenarioSelect')?.value||'normal',mode=$('redundancySelect')?.value||'N';
  root.hidden=!allowed.has(selected);
  if(root.hidden)return;
  const now=Math.max(0,window.__LS3D_TEST__?.scenarioElapsed?.()||0);
  const cfg=window.LS3D_CONFIG||{};
  const key=JSON.stringify([selected,mode,...cfgKeys.map(k=>cfg[k])]);
  if(force||key!==cachedKey){
   cachedKey=key;
   cached=model.summarize(selected,mode,cfg);
   const body=$('tlineAssumptionRows');
   body.replaceChildren(...model.scenarioAssumptions(selected,cfg).map(([k,v])=>{
    const tr=document.createElement('tr'),th=document.createElement('td'),td=document.createElement('td');
    th.textContent=k;td.textContent=v;tr.append(th,td);return tr;
   }));
   const specs=chartSpec(selected,cached);
   $('tlineGraphTitleA').textContent=specs[0].title;
   $('tlineGraphTitleB').textContent=specs[1].title;
  }
  const horizon=cached.horizonSec,elapsed=Math.min(horizon,now);
  const point=model.point(selected,mode,elapsed,cfg),specs=chartSpec(selected,cached);
  $('tlineGraphA').innerHTML=svgMarkup(specs[0],cached.rows,elapsed,horizon,specs[0].title);
  $('tlineGraphB').innerHTML=svgMarkup(specs[1],cached.rows,elapsed,horizon,specs[1].title);
  $('tlineNow').textContent='현재 ('+fmtSec(now)+') '+point.state;
  $('tlineWorst').textContent='전체 구간 결과: '+cached.state;
  $('tlineTime').textContent='예측 중단 '+fmtSec(cached.downtimeSec)+' · 최고온도 '+fmt(cached.maxTemperature)+'°C';
  const readings=[
   '현재 이벤트: '+point.event,
   '현재 랙 입구 '+fmt(point.temperature)+'°C',
   '영향 랙 '+point.affectedRacks+'대',
   '최초 27°C '+fmtSec(cached.firstWarnSec),
   '32°C '+fmtSec(cached.firstThrottleSec),
   '서비스 중단 '+fmtSec(cached.firstOutageSec)
  ];
  if(selected==='power'||selected==='ups-battery')readings.push('UPS 잔여 '+fmt(point.upsRemaining)+'분');
  if(selected==='gpu-surge')readings.push('GPU 요구 '+fmt(point.requestedKw,0)+' kW');
  $('tlineReading').textContent=readings.join(' · ');
 }
 const events=['change','input'];
 for(const id of ['scenarioSelect','redundancySelect'])for(const event of events)$(id)?.addEventListener(event,()=>redraw(true));
 window.addEventListener('ls3d-tick',()=>{
  const t=performance.now();if(t-lastDraw<450)return;lastDraw=t;redraw(false);
 });
 window.addEventListener('resize',()=>redraw(true));
 document.addEventListener('visibilitychange',()=>{if(!document.hidden)redraw(true);});
 window.LS3D_FAILURE_TIMELINE={redraw,get current(){return cached},get point(){return model.point($('scenarioSelect')?.value,$('redundancySelect')?.value,window.__LS3D_TEST__?.scenarioElapsed?.()||0,window.LS3D_CONFIG||{})}};
 redraw(true);
})();
