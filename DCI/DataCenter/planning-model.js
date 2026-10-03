/* Planning quantities: installed route lengths exclude spare stock and unmeasured patch leads. */
(() => {
  'use strict';
  const presets = {
    'amd-mi300x': {vendor:'AMD',name:'AMD MI300X · 8 가속기 계획 노드',gpus:8,ru:8,kw:10,links:8,speed:400,cords:6,source:'https://www.amd.com/en/products/accelerators/instinct/mi300/mi300x.html',note:'MI300X 750 W/가속기 공식 사양. 8U·10 kW·외부 8×400G·전원 6개는 OEM 미선정 계획 가정입니다.'},
    'intel-gaudi3': {vendor:'Intel',name:'Intel Gaudi 3 · 8 가속기 계획 노드',gpus:8,ru:8,kw:12,links:8,speed:400,cords:6,source:'https://www.intel.com/content/www/us/en/content-details/845118/intel-gaudi-3-ai-accelerator-30-3-30-pdf.html',note:'Gaudi 3 공랭 OAM 900 W/가속기 공식 사양. 8U·12 kW·외부 8×400G·전원 6개는 계획 가정이며 Gaudi 고유 포트 구성 재현이 아닙니다.'}
  };
  const regions = {
    kr:{name:'한국 · 380 V / 60 Hz 예시',voltage:380,hz:60,currency:'KRW',framework:'site'},
    eu:{name:'유럽 · 400 V / 50 Hz 예시',voltage:400,hz:50,currency:'EUR',framework:'iec'},
    us:{name:'미국 · 415 V / 60 Hz DC 예시',voltage:415,hz:60,currency:'USD',framework:'ul'},
    jp50:{name:'일본 · 200 V / 50 Hz 예시',voltage:200,hz:50,currency:'JPY',framework:'site'},
    jp60:{name:'일본 · 200 V / 60 Hz 예시',voltage:200,hz:60,currency:'JPY',framework:'site'},
    cn:{name:'중국 · 380 V / 50 Hz 예시',voltage:380,hz:50,currency:'CNY',framework:'iec'},
    custom:{name:'현장 지정',voltage:400,hz:50,currency:'USD',framework:'site'}
  };
  const references = {
    h200:{gpus:8,ru:8,kw:10.2,source:'https://docs.nvidia.com/dgx/dgxh100-user-guide/introduction-to-dgxh100.html'},
    b200:{gpus:8,ru:10,kw:14.3,source:'https://docs.nvidia.com/dgx/dgxb200-user-guide/introduction-to-dgxb200.html'},
    b300:{gpus:8,ru:10,kw:14.5,source:'https://docs.nvidia.com/dgx/dgxb300-user-guide/introduction-to-dgxb300.html'}
  };
  const errorPct=(actual,expected)=>expected===0?(actual===0?0:null):100*Math.abs(actual-expected)/Math.abs(expected);
  function totals(r){
    let routeLengthM=0,trunkLengthM=0,assemblyLengthM=0,p2pLengthM=0,activeFibers=0,installedFibers=0,hasRoutes=false;
    for(const [segment,o] of Object.entries(r.optical||{})){
      if(!o.links)continue;
      const d=Number(r.input?.[{server:'serverDistanceM',leafSpine:'leafSpineDistanceM',core:'coreDistanceM'}[segment]]);
      if(!Number.isFinite(d))continue;
      hasRoutes=true;
      if(o.trunk?.installedCableCount){const length=d*o.trunk.installedCableCount;trunkLengthM+=length;routeLengthM+=length;activeFibers+=Number(o.trunk.requiredFibers)||0;installedFibers+=Number(o.trunk.provisionedFibers)||0;}
      else {const length=d*o.links;routeLengthM+=length;if(o.profile?.optical){p2pLengthM+=length;activeFibers+=o.links*(Number(o.profile.fibers)||0);installedFibers+=o.links*(Number(o.profile.installedChannelFibers)||Number(o.profile.fibers)||0);}else assemblyLengthM+=length;}
    }
    return {routeLengthM,trunkLengthM,assemblyLengthM,p2pLengthM,activeFibers,installedFibers,hasRoutes,bomLines:(r.bom||[]).length,itKw:r.summary?.totalItPowerKw||0,facilityKw:r.facility?.facilityPowerKw||0};
  }
  function quote(r,prices,currency){let subtotal=0,priced=0;for(const x of r.bom||[]){const key=x.category+' · '+x.item,p=prices[currency]?.[key];if(p!==''&&p!=null&&Number.isFinite(Number(p))&&Number(p)>=0){subtotal+=Number(p)*Number(x.qty);priced++;}}return {subtotal,priced,total:(r.bom||[]).length,complete:priced===(r.bom||[]).length&&priced>0};}
  function checks(r){
    if(!r.usable)return [];
    const ref=references[r.input?.systemId],rows=[];
    const add=(label,actual,expected,unit,source,type)=>rows.push({label,actual,expected,unit,source,type,errorPct:errorPct(Number(actual),Number(expected))});
    if(ref&&r.systemProfile){const p=r.systemProfile;add('가속기 / 서버',p.gpus,ref.gpus,'개',ref.source,'공식 사양');add('서버 높이',p.ru,ref.ru,'RU',ref.source,'공식 사양');add('서버 최대 설계전력',p.power,ref.kw,'kW',ref.source,'공식 사양');add('Compute 전력',r.summary.computePowerKw,Math.ceil(r.input.targetGPU/ref.gpus)*ref.kw,'kW',ref.source,'공식 사양 × 서버 수');}
    for(const [key,o] of Object.entries(r.optical||{}))if(o.loss?.applicable){const i=r.input,d=Number(i[{server:'serverDistanceM',leafSpine:'leafSpineDistanceM',core:'coreDistanceM'}[key]]),structured=i[{server:'serverCabling',leafSpine:'leafSpineCabling',core:'coreCabling'}[key]]==='structured',pairs=structured?i.matedPairs:2,connector=/MPO|MTP/i.test(o.profile?.connector||'')?i.mpoLossDb:i.lcLossDb,expected=d/1000*i.fiberAttenDbKm+pairs*connector+i.spliceCount*i.spliceLossDb+i.marginDb;add(key+' 채널 손실',o.loss.estimatedDb,Math.round(expected*1000)/1000,'dB','', '입력 가정의 독립 산식');}
    if(r.portAudit){const routes=r.portAudit.routes||[];for(const [key,o] of Object.entries(r.optical||{}))add(key+' 링크 보존',routes.filter(x=>x.segment===key).reduce((s,x)=>s+x.links,0),o.links,'링크','','예약 경로 합계 점검');}
    return rows;
  }
  function regionalCapacity(i){const v=Number(i.siteVoltage),a=Number(i.siteCurrentA),pf=Number(i.powerFactor),phase=Number(i.sitePhases);if(![v,a,pf,phase].every(Number.isFinite)||v<=0||a<=0||pf<=0||pf>1||![1,3].includes(phase)||![50,60].includes(Number(i.siteFrequency)))throw Error('현장 전압·전류·역률·상수·주파수를 확인하세요.');return (phase===3?Math.sqrt(3):1)*v*a*pf/1000;}
  window.DCPlanning={presets,regions,references,totals,quote,checks,regionalCapacity};
})();
