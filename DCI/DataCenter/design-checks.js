/* Read-only checks against the calculated rack schedule, not nominal BOM quantities. */
(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;else root.DCDesignChecks=api;})(typeof window!=='undefined'?window:globalThis,function(){
 'use strict';
 const finite=n=>typeof n==='number'&&Number.isFinite(n);
 function evaluate(r){
  if(!r?.usable)return {valid:false,metrics:[],issues:[],errors:[r?.error||'입력값을 확인하고 다시 계산하세요.'],warnings:[]};
  const i=r.input||{},racks=r.racks||[],issues=[];
  const defs=[['power','전력',i.rackPowerKw,'kW',x=>x.powerKw],['cooling','냉각',i.rackCoolingKw,'kW',x=>x.powerKw],['ru','랙 공간',i.rackRU,'RU',x=>x.ru+(r.systemProfile?.rackScale&&x.role==='Compute'?0:Number(i.reservedRU||0))]];
  const metrics=defs.map(([key,label,limit,unit,value])=>{
   if(!racks.length||!finite(limit)||limit<=0||racks.some(x=>!finite(value(x))))return {key,label,unit,known:false};
   const peak=racks.reduce((a,b)=>value(a)>=value(b)?a:b),bad=racks.filter(x=>value(x)>limit+1e-6);
   const m={key,label,unit,known:true,max:value(peak),limit,rack:peak.id,count:bad.length};
   if(bad.length)issues.push({level:'error',key,message:`${label} 한도 초과: ${bad.length}개 랙 · ${peak.id} 최대 ${Number(m.max.toFixed(2))} / ${limit} ${unit}`,racks:bad.map(x=>x.id)});
   return m;
  });
  const head=1-Number(i.rackHeadroomPct||0)/100;
  for(const m of metrics.filter(m=>m.known&&m.key!=='ru')){
   // OEM rack kits use their full stated rating in the engineering model.
   const affected=racks.filter(x=>!(r.systemProfile?.rackScale&&x.role==='Compute')&&x.powerKw>m.limit*head+1e-6&&x.powerKw<=m.limit+1e-6);
   if(affected.length)issues.push({level:'warning',key:m.key+'-headroom',message:`${m.label} 예비율 ${i.rackHeadroomPct}% 미확보: ${affected.length}개 랙 · 계획 한도 ${Number((m.limit*head).toFixed(2))} kW`,racks:affected.map(x=>x.id)});
  }
  const threshold=r.cooling?.mode==='air'?i.airLimitKw:r.cooling?.mode==='rear-door'?i.rearDoorLimitKw:null;
  if(finite(threshold)){
   const affected=racks.filter(x=>x.role==='Compute'&&x.powerKw>threshold+1e-6);
   if(affected.length)issues.push({level:'warning',key:'cooling-method',message:`${r.cooling.mode==='air'?'공랭':'Rear-door'} 계획 기준 ${threshold} kW/랙 초과: ${affected.length}개 랙 · 냉각 방식과 제조사 사양 확인`,racks:affected.map(x=>x.id)});
  }
  const errors=[...new Set(r.errors||[])],warnings=[...new Set(r.warnings||[])].filter(x=>!errors.includes(x));
  return {valid:true,metrics,issues,errors,warnings};
 }
 return {evaluate};
});
