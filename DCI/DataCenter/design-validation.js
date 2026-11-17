/* Pure rules: no DOM or display strings. Every outcome has a reason and an action. */
(function(root,factory){const cfg=root.DCValidationConfig||(typeof require==='function'?require('./design-validation-config.js'):{}),calc=root.DCCalculations||(typeof require==='function'?require('./calculation-definitions.js'):null),api=factory(cfg,calc);if(typeof module==='object'&&module.exports)module.exports=api;root.DCDesignValidation=api;})(typeof window!=='undefined'?window:globalThis,function(defaults,calc){
'use strict';
const sum=a=>a.reduce((s,v)=>s+v,0),finite=Number.isFinite;
function evaluate(r,overrides={}){
 const c={...defaults,...overrides},rows=[],add=(id,status,reason,suggestion,values={},refs=[])=>rows.push({id,status,reason,suggestion,values,refs});
 if(!r?.usable||!r.summary||!r.portAudit){add('result','WARN','unavailable','recalculate');return finish();}
 const i=r.input,p=r.systemProfile||{},racks=r.racks||[],nodes=r.portAudit.nodes||[],routes=r.portAudit.routes||[],eps=c.tolerance;
 const above=(n,limit)=>!finite(n)||!finite(limit)||n>limit+eps;
 const peak=Math.max(0,...racks.map(x=>x.powerKw));
 add('rackCooling',racks.length&&!above(peak,i.rackCoolingKw)?'PASS':'FAIL','rackCooling','increaseCooling',{actual:peak,limit:i.rackCoolingKw,unit:'kW'},['rack.perRack','power.IT']);
 add('rackPower',racks.length&&!above(peak,i.rackPowerKw)?'PASS':'FAIL','rackPower','increaseRackPower',{actual:peak,limit:i.rackPowerKw,unit:'kW'},['rack.perRack']);
 const thermalLimit=r.cooling?.mode==='air'?i.airLimitKw:r.cooling?.mode==='rear-door'?i.rearDoorLimitKw:Infinity;
 const computePeak=Math.max(0,...racks.filter(x=>x.role==='Compute').map(x=>x.powerKw));
 const liquidInvalid=p.rackScale&&!['dlc','immersion'].includes(r.cooling?.mode);
 add('coolingMethod',liquidInvalid?'FAIL':(!finite(computePeak)||finite(thermalLimit)&&above(computePeak,thermalLimit))?'WARN':'PASS','coolingMethod','verifyCooling',{actual:computePeak,limit:finite(thermalLimit)?thermalLimit:'OEM / RFQ',mode:r.cooling?.mode,unit:'kW'},[p.rackScale?'rack.perRack':'rack.limitCooling']);
 const leaves=nodes.filter(n=>n.role==='Leaf'),switches=nodes.filter(n=>n.role!=='Endpoint');
 const overflow=switches.filter(n=>!finite(n.capacity)||sum(Object.values(n.used||{}))>n.capacity+eps);
 const missingUp=leaves.filter(n=>n.used.server>0&&!(n.used.leafSpine>0));
 add('leafPorts',overflow.length||missingUp.length?'FAIL':'PASS','leafPorts','increaseRadix',{overflow:overflow.length,missingUp:missingUp.length,switches:switches.length,radix:i.switchRadix},['fabric.down','fabric.up','fabric.leaf']);
 const ratio=Math.max(0,...leaves.map(n=>n.used.server/Math.max(eps,n.used.leafSpine)));
 add('oversubscription',ratio>i.oversub+eps?'FAIL':ratio>c.oversubscriptionWarnRatio+eps?'WARN':'PASS','oversubscription','reduceOversub',{actual:ratio,limit:i.oversub,warn:c.oversubscriptionWarnRatio,unit:':1'},['fabric.up']);
 const domains=r.fabric?.domains||[],capacityBad=domains.some(d=>d.hosts>(d.tier===3?d.threeLimit:d.twoLimit)+eps);
 const requested=r.evidence?.rawInput?.topology||i.topology,requestedTier=Number(r.evidence?.rawInput?.fabricTier||2),promoted=domains.some(d=>d.tier===3)&&requested!=='clos3'&&requestedTier!==3;
 add('topologyRadix',capacityBad||overflow.length||i.switchRadix<c.minimumRadix?'FAIL':promoted?'WARN':'PASS','topologyRadix','splitFabric',{requested,effective:i.topology,tier:r.fabric.stageCount,radix:i.switchRadix,domains:domains.length,promoted},['fabric.leaf','fabric.spine','fabric.core']);
 for(const [segment,o]of Object.entries(r.optical||{})){
  if(!o.links)continue;
  const profile=o.profile||{},reach=profile.reachM,reachKnown=finite(reach),reachFail=!finite(o.lengthM)||reachKnown&&above(o.lengthM,reach);
  add('reach.'+segment,reachFail?'FAIL':!reachKnown||!profile.opticSku?'WARN':'PASS','reach','selectReach',{segment,media:profile.media,actual:o.lengthM,limit:reachKnown?reach:'RFQ',unit:'m',assumption:!profile.opticSku},[segment+'.route',segment+'.length']);
  if(o.loss?.applicable)add('loss.'+segment,!o.loss.pass?'FAIL':i.marginDb<c.minimumOpticalMarginDb?'WARN':'PASS','loss','reviewLoss',{segment,actual:o.loss.estimatedDb,margin:i.marginDb,limit:o.loss.budgetDb,unit:'dB'},[segment+'.loss']);
  const structured=i[{server:'serverCabling',leafSpine:'leafSpineCabling',core:'coreCabling'}[segment]]==='structured';
  if(profile.optical&&structured){
   const channel=profile.installedChannelFibers,actual=profile.trunkFiberCount,requested=profile.requestedTrunkFiberCount;
   const capacity=channel>0?Math.floor(actual/channel):0,required=capacity?sum(routes.filter(q=>q.segment===segment).map(q=>Math.ceil(q.links/capacity))):Infinity;
   const fail=!finite(channel)||channel<=0||capacity<1||requested>0&&requested<channel||o.trunk.installedCableCount<required||o.trunk.provisionedFibers<o.links*channel;
   add('trunk.'+segment,fail?'FAIL':actual%channel?'WARN':'PASS','trunk','fixTrunk',{segment,requested,actual,channel,required,installed:o.trunk.installedCableCount,unit:'F / cable'},[segment+'.trunkCapacity',segment+'.trunksInstalled']);
   const buy=calc.evaluate('purchase',{installed:o.trunk.installedCableCount,sparePct:i.sparePct,pack:i.trunkBundle});
   add('trunkSpare.'+segment,o.trunk.cableCount<buy?'FAIL':i.sparePct<c.minimumSparePct?'WARN':'PASS','trunkSpare','increaseSpare',{segment,installed:o.trunk.installedCableCount,required:buy,actual:o.trunk.cableCount,sparePct:i.sparePct,unit:'cable'},[segment+'.trunksInstalled']);
  }
 }
 for(const [kind,unit]of [['ups','kVA'],['generator','kW'],['transformer','kVA']]){
  const mode=i.facilityRedundancy,qty=r.facility?.[kind+'Qty'],duty=r.facility?.[kind+'Duty'],usable=r.facility?.[kind+'UsableBlock'];
  const step=r.evidence?.steps?.find(s=>s.id==='facility.'+kind+'Design'),need=step?.value??r.facility?.[kind+'Design'+(kind==='generator'?'Kw':'Kva')];
  const surviving=mode==='2N'?Math.floor(qty/2):mode==='N+1'?Math.max(0,qty-1):qty;
  const valid=finite(qty)&&Number.isInteger(qty)&&finite(need)&&finite(usable)&&finite(duty)&&['N','N+1','2N'].includes(mode);
  add('facility.'+kind,!i.facilityEnabled?'WARN':!valid||surviving*usable+eps<need?'FAIL':'PASS',!i.facilityEnabled?'disabled':'facilityCapacity','increaseFacility',{kind,mode,qty,duty,surviving,usable,required:need,actual:surviving*usable,unit,scope:kind==='ups'?r.facility.upsLoadScope:'facility'},['facility.'+kind+'Design','facility.'+kind+'Duty','facility.'+kind+'Qty']);
 }
 const quantityBad=(r.bom||[]).filter(b=>!finite(b.qty)||b.qty<0||!Number.isInteger(b.qty)||b.qty<b.installedQty);
 add('bomQuantities',quantityBad.length?'FAIL':'PASS','bomQuantities','reviewInputs',{invalid:quantityBad.length,rows:r.bom?.length});
 add('qualification','WARN','qualification','verifyOEM');
 if(r.errors?.length)add('engine','FAIL','engine','reviewInputs',{count:r.errors.length,details:r.errors});
 function finish(){const counts={PASS:0,WARN:0,FAIL:0};rows.forEach(r=>counts[r.status]++);return {schemaVersion:1,status:counts.FAIL?'FAIL':counts.WARN?'WARN':'PASS',counts,rules:rows,config:c};}
 return finish();
}
return {evaluate};
});
