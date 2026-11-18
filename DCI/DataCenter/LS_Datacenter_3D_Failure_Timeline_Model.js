/* Educational time-axis incident model, deliberately not a thermal CFD,
 * electrical transient, fire-code or manufacturer validation tool.
 * Pure functions: identical inputs give identical curves and classifications.
 */
(function(root,factory){
 'use strict';
 const api=factory();
 if(typeof module==='object'&&module.exports)module.exports=api;
 if(root)root.LS3D_TIMELINE_MODEL=api;
})(typeof window!=='undefined'?window:null,function(){
 'use strict';
 const NORMAL='정상 유지', DEGRADED='성능 저하', OUTAGE='서비스 중단';
 const modes=['N','N+1','2N'];
 const horizon={normal:180,power:900,'ups-battery':960,cooling:900,'dlc-leak':360,fire:420,'gpu-surge':780,
  network:300,'rack-hotspot':300,'fiber-cut':240};
 const assumptions={
  global:[
   ['모델 범위','실측이 아닌 시각화용 연속 시간 계산. 설비 보호, 소화 방출, 배터리 수명·안전 판정에 사용 불가'],
   ['시간 간격','전체 기간을 2초 간격으로 계산(현 시점은 정확한 초값으로 별도 계산)'],
   ['결과 등급','정상 유지: 계획 부하 공급; 성능 저하: 온도 경고/스로틀링/일부 랙 영향; 서비스 중단: 지정 부하/구역 공급 중단'],
   ['이중화 정의','N은 예비 경로 없음, N+1은 예비 1계통, 2N은 독립된 이중 경로를 가정. 실제 공통원인고장은 미반영'],
   ['가용성 백분율','별도의 기존 교육용 KPI 가용성 표와 다름. 본 비교표의 상태는 시간축에서 발생한 서비스 영향 등급']
  ],
  power:[
   ['Utility 상실','t=0초에 상용전원 0 kW로 가정'],
   ['기동 및 ATS','발전기 기동 12초 + ATS 절체 6초 후 백업 가용; 시동/절체 성공 여부는 용량과 예비경로 가정'],
   ['UPS 자율시간','설정된 백업 분 × 60초 × min(1, UPS가용 kW / IT부하 kW)로 단순 계산'],
   ['N/N+1/2N','각각 기동 성공에 필요한 가용 용량 배율 1.00 / 1.00 / 1.00. 경로 수에 따른 절체 유예 0/2/4초는 교육용'],
   ['판정','UPS·발전기가 부하를 감당하면 정상, 일부만 감당하면 성능 저하, 전력 공급이 0이면 중단']
  ],
  'ups-battery':[
   ['발전기 고장','기동 실패가 지속되는 시나리오. Utility 복구 t=900초 가정'],
   ['UPS 잔여시간','설정 batteryMin × 60초 × min(1,upsKw/load)로 단순화'],
   ['이중화 한계','공통 Utility 상실 + 모든 발전기 기동 실패를 가정하므로 2N도 무한정 버티지 못함']
  ],
  cooling:[
   ['온도 1차 모델','기준 랙 입구온도 + dT/dt × 유효 고장 경과 시간; CFD·열관성 실측 아님'],
   ['N/N+1/2N 온도 상승률','0.055 / 0.021 / 0.006 °C/초; 600초에 백업 냉각 공급, 780초에 회복 시작'],
   ['임계값','27°C: 경고, 32°C: 스로틀링, 41°C: 일부 서비스 중단. 제조사 보증 한계가 아닌 교육용 가정']
  ],
  'dlc-leak':[
   ['대상 수','ceil(랙 수 × DLC 비율 × 영향계수); 영향계수 N=0.30,N+1=0.16,2N=0.08'],
   ['누수 처리','25초 공급 차단, 60초 제한, 150초 격리, 250초부터 복구 시작 가정'],
   ['부하 영향','누수 영향 랙이 전체 랙 대비 차지하는 비율로 서비스 감소를 계산하며 전체 DC 중단과 구분']
  ],
  fire:[
   ['가상 안전 이벤트','0초 감지, 20초 공조 정지, 100초 대피 완료 가정, 120초 가스 소화 방출 표시, 160~360초 구역 격리'],
   ['주의','실제 소화 방출 승인, 인명 안전, 설비 연동, 대피 시퀀스를 결정하는 로직이 아님'],
   ['이중화와 구역 격리','2N이더라도 안전 확보를 위해 격리되는 구역은 서비스 중단으로 계산']
  ],
  'gpu-surge':[
   ['학습 잡 부하','GPU 부분의 IT 부하가 기준 대비 최대 +55%로 180초에 선형 상승, 600초까지 유지, 720초에 복귀'],
   ['N/N+1/2N IT 여유율','1.35 / 1.48 / 1.65 × 기준 IT 부하. 다른 설비 사양과 무관한 교육용 가정'],
   ['온도와 부하 제어','부하 초과분을 제한. N/N+1/2N 추가 냉각 여유 가정은 1.00/1.12/1.35배, 온도상승 계수는 1.00/0.62/0.35배']
  ]
 };
 const clamp=(x,a,b)=>Math.min(b,Math.max(a,x));
 const num=(v,def)=>Number.isFinite(Number(v))?Number(v):def;
 const smooth=(a,b,t)=>clamp((t-a)/Math.max(1,b-a),0,1);
 function parameters(config={}){
  const rackCount=Math.max(1,Math.round(num(config.rackCount,20)));
  const rackPowerKw=Math.max(.1,num(config.rackPowerKw,60));
  const ports=Math.max(0,num(config.portCount,128));
  const mode=config.mode==='cpo'?'cpo':'pluggable';
  const opticalKw=Math.max(0,ports*(mode==='cpo'?num(config.cpoW,10):num(config.plugW,12))/1000
    +(mode==='cpo'?num(config.cpoKw,1):num(config.plugKw,2)));
  const itBaseKw=rackCount*rackPowerKw+opticalKw;
  const tempBaseline=num(config.baselineRackInletC,22)
    +Math.max(0,itBaseKw/Math.max(1,num(config.coolingKw,1600))-.7)*4
    +Math.max(0,num(config.outdoorC,24)-24)*.08;
  return {rackCount,itBaseKw,opticalKw,
   dlcFraction:clamp(num(config.dlc,60)/100,0,1),
   batterySec:Math.max(0,num(config.batteryMin,12))*60,
   upsKw:Math.max(0,num(config.upsKw,1600)),
   generatorKw:Math.max(0,num(config.generatorKw,1800)),
   coolingKw:Math.max(0,num(config.coolingKw,1600)),
   tempBaseline, gpuRatio:clamp(num(config.gpuRatio,80)/100,0,1)};
 }
 function point(scenario,mode,t,config={}){
  const p=parameters(config);
  const sec=clamp(num(t,0),0,86400);
  const r=modes.includes(mode)?mode:'N';
  const o={t:sec,temperature:p.tempBaseline,upsRemaining:p.batterySec/60,affectedRacks:0,
   itLoad:p.itBaseKw,requestedKw:p.itBaseKw,availableCapacity:p.itBaseKw,
   serviceFraction:1,state:NORMAL,event:'정상 운영',hint:'기준 상태'};
  if(scenario==='normal')return o;
  if(scenario==='power'||scenario==='ups-battery'){
   const failed=scenario==='ups-battery',boot=12,transfer=6,lag={N:0,'N+1':2,'2N':4}[r],
    readyAt=boot+transfer-lag,upsSec=p.batterySec*Math.min(1,p.upsKw/Math.max(1,p.itBaseKw)),
    genReady=!failed&&p.generatorKw>=p.itBaseKw&&sec>=readyAt;
   const upsSupported=p.upsKw>=p.itBaseKw, batteryAlive=sec<upsSec;
   const utilityRestored=failed&&sec>=900;
   const supplied=utilityRestored||genReady?1:batteryAlive?Math.min(1,p.upsKw/Math.max(1,p.itBaseKw)):0;
   o.itLoad=p.itBaseKw*supplied;
   o.serviceFraction=supplied;
   o.upsRemaining=genReady||utilityRestored?p.batterySec/60:Math.max(0,(upsSec-sec)/60);
   o.availableCapacity=utilityRestored?p.itBaseKw:genReady?p.generatorKw:batteryAlive?p.upsKw:0;
   o.event=utilityRestored?'Utility 복구':genReady?'발전기·ATS 절체 완료':batteryAlive?'UPS 공급':'UPS 소진 · 무전원';
   o.hint='UPS 자율시간 '+(upsSec/60).toFixed(1)+'분 · 절체 '+readyAt+'초';
  }else if(scenario==='cooling'){
   const slope={N:.055,'N+1':.021,'2N':.006}[r];
   const heat=sec<600?slope*sec:
     sec<780?slope*600*(1-.15*smooth(600,780,sec)):
     slope*600*.85*Math.exp(-(sec-780)/65);
   o.temperature=p.tempBaseline+heat;
   o.event=sec<600?'냉각 공급 부족':sec<780?'예비 냉각 공급': '냉각 회복 중';
   const dt=o.temperature;
   if(dt>=41){o.serviceFraction=.45;o.itLoad=p.itBaseKw*.45;}
   else if(dt>=32){o.serviceFraction=.82;o.itLoad=p.itBaseKw*.82;}
   o.hint='온도 상승률 '+slope.toFixed(3)+'°C/초 · 경고 27°C/제한 32°C/차단 41°C';
  }else if(scenario==='dlc-leak'){
   const frac={N:.30,'N+1':.16,'2N':.08}[r],
     maxRacks=Math.min(p.rackCount,Math.ceil(p.rackCount*p.dlcFraction*frac));
   const severity=sec<25?Math.max(.2,smooth(0,25,sec)):
      sec<150?1:sec<250?1-smooth(150,250,sec)*.65:
      (1-.65)*Math.exp(-(sec-250)/25);
   o.affectedRacks=sec>=350?0:Math.ceil(maxRacks*severity);
   o.temperature=p.tempBaseline+o.affectedRacks/Math.max(1,p.rackCount)*20;
   const impact=sec<25?.18:sec<60?.35:sec<150?.75:sec<250?.48:.30;
   o.serviceFraction=clamp(1-impact*o.affectedRacks/p.rackCount,0,1);
   o.itLoad=p.itBaseKw*o.serviceFraction;
   o.event=sec<25?'누수 감지':sec<150?'냉각수 차단·부하 제한':sec<250?'누수 격리':'냉각 경로 복구';
   o.hint='영향 랙 '+o.affectedRacks+'/'+p.rackCount+' · DLC 비율 '+(p.dlcFraction*100).toFixed(0)+'%';
  }else if(scenario==='fire'){
   const zoneRacks=Math.max(1,Math.ceil(p.rackCount*.20));
   o.affectedRacks=sec>=360?0:zoneRacks;
   const isolated=sec>=120&&sec<360;
   o.serviceFraction=isolated?1-zoneRacks/p.rackCount:1;
   o.itLoad=p.itBaseKw*o.serviceFraction;
   o.temperature=p.tempBaseline+(sec<360&&sec>=20?4:0);
   o.event=sec<20?'화재 감지·경보':sec<100?'공조 정지·대피 진행':sec<120?'대피 완료 가정':sec<160?'가스 소화 방출·격리':sec<360?'구역 서비스 차단':'복구 완료 가정';
   o.hint='격리 구역 '+zoneRacks+'랙(전체의 약 20%) · 소화 방출은 시각화일 뿐';
   o.safetyOutage=isolated;
  }else if(scenario==='gpu-surge'){
   const ramp=sec<180?smooth(0,180,sec):sec<600?1:1-smooth(600,720,sec);
   const requested=p.itBaseKw*(1+.55*p.gpuRatio*ramp);
   const available=p.itBaseKw*{N:1.35,'N+1':1.48,'2N':1.65}[r];
   const coolingCapacity=p.coolingKw*{N:1,'N+1':1.12,'2N':1.35}[r];
   const thermalFactor={N:1,'N+1':.62,'2N':.35}[r];
   const overload=Math.max(0,requested-available),coolingOver=Math.max(0,requested-coolingCapacity);
   o.requestedKw=requested;o.itLoad=Math.min(requested,available);
   o.serviceFraction=clamp(1-overload/Math.max(1,requested),0,1);
   o.availableCapacity=available;o.temperature=p.tempBaseline+5*ramp*thermalFactor
     +clamp(coolingOver/Math.max(1,coolingCapacity),0,1)*8;
   o.event=sec<180?'GPU 학습 잡 증가':sec<600?'GPU 고부하 유지':'학습 잡 종료';
   o.hint='IT 요구 '+requested.toFixed(0)+' kW · 공급 여유 '+available.toFixed(0)+' kW';
  }
  if(scenario==='fire'&&o.safetyOutage)o.state=OUTAGE;
  else if(scenario==='power'||scenario==='ups-battery'){
    o.state=o.serviceFraction<.05?OUTAGE:o.serviceFraction<.999?DEGRADED:NORMAL;
  }else if(scenario==='cooling'){
    o.state=o.temperature>=41?OUTAGE:o.temperature>=27?DEGRADED:NORMAL;
  }else if(scenario==='dlc-leak'){
    o.state=o.affectedRacks>0?DEGRADED:NORMAL;
  }else if(scenario==='gpu-surge'){
    o.state=o.serviceFraction<.6?OUTAGE:o.serviceFraction<.999||o.temperature>=27?DEGRADED:NORMAL;
  }
  return o;
 }
 function curve(scenario,mode,config={},stepSec=2){
  const end=horizon[scenario]||300,step=clamp(Math.round(stepSec)||2,1,60),rows=[];
  for(let t=0;t<=end;t+=step)rows.push(point(scenario,mode,t,config));
  if(rows.at(-1).t!==end)rows.push(point(scenario,mode,end,config));
  return rows;
 }
 function summarize(scenario,mode,config={}){
  const rows=curve(scenario,mode,config),p=parameters(config),duration=rows.at(-1).t,
    peaks={temperature:-Infinity,affectedRacks:0,requestedKw:0,itLoad:0},
    stageTime={[NORMAL]:0,[DEGRADED]:0,[OUTAGE]:0};
  let firstWarn=null,firstThrottle=null,firstStop=null,lastStop=null,criticalAt=null;
  for(let i=0;i<rows.length;i++){
   const row=rows[i],next=rows[i+1],dt=next?next.t-row.t:0;
   peaks.temperature=Math.max(peaks.temperature,row.temperature);
   peaks.affectedRacks=Math.max(peaks.affectedRacks,row.affectedRacks);
   peaks.requestedKw=Math.max(peaks.requestedKw,row.requestedKw);
   peaks.itLoad=Math.max(peaks.itLoad,row.itLoad);
   stageTime[row.state]=(stageTime[row.state]||0)+dt;
   if(row.temperature>=27&&firstWarn===null)firstWarn=row.t;
   if(row.temperature>=32&&firstThrottle===null)firstThrottle=row.t;
   if(row.state===OUTAGE&&firstStop===null)firstStop=row.t;
   if(row.state===OUTAGE)lastStop=row.t;
  }
  const state=stageTime[OUTAGE]>0?OUTAGE:stageTime[DEGRADED]>0?DEGRADED:NORMAL;
  return {scenario,mode,state,horizonSec:duration,firstWarnSec:firstWarn,firstThrottleSec:firstThrottle,
   firstOutageSec:firstStop,downtimeSec:stageTime[OUTAGE],degradedSec:stageTime[DEGRADED],
   maxTemperature:peaks.temperature,maxAffectedRacks:peaks.affectedRacks,
   maxDemandKw:peaks.requestedKw,maxITKw:peaks.itLoad,
   baseITKw:p.itBaseKw,batteryMin:p.batterySec/60,rows};
 }
 function scenarioAssumptions(scenario,config={}){
  const p=parameters(config);
  const dynamic=[
   ['현재 랙 수',p.rackCount+'개'],
   ['현재 기준 IT 부하',p.itBaseKw.toFixed(1)+' kW (랙수×랙전력 + 광네트워크)'],
   ['가정 랙 입구온도',p.tempBaseline.toFixed(1)+' °C'],
   ['냉각 설비 용량',p.coolingKw.toFixed(0)+' kW'],
   ['UPS/발전기 가용용량',p.upsKw.toFixed(0)+' / '+p.generatorKw.toFixed(0)+' kW'],
   ['UPS 배터리 백업',p.batterySec/60+' 분'],
   ['DLC 대상 비율',(p.dlcFraction*100).toFixed(0)+' %'],
   ['GPU 대상 비율',(p.gpuRatio*100).toFixed(0)+' %']
  ];
  return [...assumptions.global,...(assumptions[scenario]||[]),...dynamic];
 }
 return {point,curve,summarize,parameters,scenarioAssumptions,horizon,modes,assumptions,
  STATE:{NORMAL,DEGRADED,OUTAGE}};
});
