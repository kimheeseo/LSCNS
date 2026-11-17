/* One executable definition per formula. Engine and evidence UI consume this registry. */
(function(root,factory){const api=factory();if(typeof module==='object'&&module.exports)module.exports=api;root.DCCalculations=api;})(typeof window!=='undefined'?window:globalThis,function(){
'use strict';
const total=a=>a.reduce((s,n)=>s+n,0),ceil=Math.ceil,floor=Math.floor;
const definitions={};
function def(id,expression,unit,calculate){definitions[id]={id,expression,unit,calculate};}
def('multiply','count × rating','',a=>a.count*a.rating);
def('sum','Σ(values)','',a=>total(a.values));
def('ceilRatio','ceil(numerator / denominator)','count',a=>ceil(a.numerator/a.denominator));
def('rackFit','max(1, min(...limits))','systems/rack',a=>Math.max(1,Math.min(...a.limits)));
def('fraction','1 − reservePct / 100','ratio',a=>1-a.reservePct/100);
def('rackRU','floor(max(0, rackRU − reservedRU) / systemRU)','systems/rack',a=>floor(Math.max(0,a.rackRU-a.reservedRU)/a.systemRU));
def('rackPowerFit','floor(rackCapacity × headroom / systemPower)','systems/rack',a=>floor(a.rackCapacity*a.headroom/a.systemPower));
def('budget','inputMW × 1000 × headroom','kW',a=>a.inputMW*1000*a.headroom);
def('powerSystems','min(racks × perRack, max(1, floor(budget / power)))','system',a=>Math.min(a.racks*a.perRack,Math.max(1,floor(a.budget/a.power))));
def('redundancy','N: duty; N+1: duty + 1; 2N: 2 × duty (zero load: 0)','module',a=>!a.duty?0:a.mode==='2N'?2*a.duty:a.mode==='N+1'?a.duty+1:a.duty);
def('portsDown','radix / 2','port/switch',a=>a.radix/2);
def('portsUp','max(1, ceil(downPorts / oversubscription))','port/switch',a=>Math.max(1,ceil(a.downPorts/a.oversubscription)));
def('leafGroups','Σgroup ceil(groups / downPorts)','switch',a=>total(a.groups.map(n=>ceil(n/a.downPorts))));
def('compactPlanes','smallest divisor s of uplinks with fanIn × (uplinks / s) ≤ capacity; fallback uplinks','switch',a=>{let s=1;while(s<a.uplinks&&(a.uplinks%s!==0||a.fanIn*(a.uplinks/s)>a.capacity))s++;return s;});
def('countRole','count(nodes where role = selectedRole)','switch',a=>a.nodes.filter(n=>n.role===a.selectedRole).length);
def('routeLinks','Σ(route.links where segment = selectedSegment)','link',a=>total(a.routes.filter(r=>r.segment===a.selectedSegment).map(r=>r.links)));
def('spareUnits','ceil(installed × sparePct / 100)','quantity',a=>ceil(a.installed*a.sparePct/100));
def('withSpare','installed + spare','quantity',a=>a.installed+a.spare);
def('purchase','ceil((installed + ceil(installed × sparePct / 100)) / pack) × pack','quantity',a=>{const p=Math.max(1,ceil(a.pack));return ceil((definitions.withSpare.calculate({installed:a.installed,spare:definitions.spareUnits.calculate(a)}))/p)*p;});
def('routeLength','(route + rise + 2 × termination) × (1 + slackPct / 100)','m',a=>(a.route+a.rise+2*a.termination)*(1+a.slackPct/100));
def('standardLength','first standard length ≥ routeLength; otherwise ceil(routeLength / increment) × increment','m',a=>a.standardLengths.find(n=>n>=a.routeLength)||ceil(a.routeLength/a.increment)*a.increment);
def('channelsPerTrunk','floor(trunkFibers / allocatedFibersPerChannel)','channel/trunk',a=>floor(a.trunkFibers/a.allocatedFibersPerChannel));
def('trunks','Σroute ceil(routeLinks / channelsPerTrunk)','cable',a=>total(a.routeLinks.map(n=>ceil(n/a.channelsPerTrunk))));
def('opticalLoss','attenuation × length / 1000 + pairs × connectorLoss + splices × spliceLoss','dB',a=>a.attenuation*a.length/1000+a.pairs*a.connectorLoss+a.splices*a.spliceLoss);
def('storageNodes','max(ceil(capacity / (rawPerNode × usableRatio)), ceil(GPU × readPerGPU / readPerNode))','node',a=>Math.max(ceil(a.capacity/(a.rawPerNode*a.usableRatio)),ceil(a.GPU*a.readPerGPU/a.readPerNode)));
def('itPower','Compute + Network + Storage + Ops','kW',a=>a.Compute+a.Network+a.Storage+a.Ops);
def('facilityPower','ITpower × PUE','kW',a=>a.ITpower*a.PUE);
def('upsDesign','protectedLoad / powerFactor / efficiency × (1 + headroomPct / 100)','kVA',a=>a.protectedLoad/a.powerFactor/a.efficiency*(1+a.headroomPct/100));
def('generatorDesign','facilityPower × (1 + headroomPct / 100)','kW',a=>a.facilityPower*(1+a.headroomPct/100));
def('transformerDesign','facilityPower / powerFactor × (1 + headroomPct / 100)','kVA',a=>a.facilityPower/a.powerFactor*(1+a.headroomPct/100));
def('usableBlock','blockRating × utilizationPct / 100','',a=>a.blockRating*a.utilizationPct/100);
def('installedModules','override > 0 ? ceil(override) : autoQuantity','module',a=>a.override>0?ceil(a.override):a.autoQuantity);
def('round','round(value × 1000) / 1000','',a=>Math.round(a.value*1000)/1000);
def('flow','liquidKW × 1000 / (density × heatCapacity × deltaT) × 60000','L/min',a=>a.liquidKW*1000/(a.density*a.heatCapacity*a.deltaT)*60000);
def('bomQuantity','ceil(calculatedQuantity)','quantity',a=>ceil(a.calculatedQuantity));
function createTrace(raw,input,defaults){
 const steps=[],counts={};
 const sourceFor=k=>({origin:Object.hasOwn(raw,k)?'input':'assumption',assumption:!Object.hasOwn(raw,k),source:Object.hasOwn(raw,k)?'Project input':'engineering-v82.js / DEFAULTS',original:raw[k],normalized:Object.hasOwn(raw,k)&&JSON.stringify(raw[k])!==JSON.stringify(input[k])});
 const inputSources=()=>Object.fromEntries(Object.keys(input).map(k=>[k,{value:input[k],...sourceFor(k)}]));
 function run(key,args,meta={}){
  const d=definitions[key];if(!d)throw Error('Unknown calculation '+key);
  const value=d.calculate(args);if(typeof value!=='number'||!Number.isFinite(value))throw Error('Non-finite calculation '+key);
  const id=meta.id||key+'.'+(counts[key]=(counts[key]||0)+1);
  const inputs=Object.entries(args).map(([name,value])=>{
   const s=meta.inputs?.[name],ref=typeof s==='string'?s:s?.ref,field=typeof s==='object'?s.field:null,provenance=field?sourceFor(field):{};
   return {name,value,ref:ref||null,...provenance,...(typeof s==='object'?s:{}),assumption:typeof s==='object'&&s.assumption!==undefined?s.assumption:field?provenance.assumption:!ref,source:typeof s==='object'&&s.source?s.source:ref?'calculation:'+ref:field?provenance.source+' / '+field:'Model constant / planning assumption'};
  });
  const step={id,key,label:meta.label||key,scope:meta.scope||'design',expression:d.expression,value,unit:meta.unit??d.unit,inputs,source:'calculation-definitions.js#'+key};
  steps.push(step);return value;
 }
 return {run,steps,inputSources,snapshot:()=>({schemaVersion:1,kind:'calculated',rawInput:raw,normalizedInput:input,inputSources:inputSources(),steps,definitions:Object.fromEntries(Object.entries(definitions).map(([k,d])=>[k,{expression:d.expression,unit:d.unit}]))})};
}
return {definitions,createTrace,evaluate:(key,args)=>definitions[key].calculate(args)};
});
