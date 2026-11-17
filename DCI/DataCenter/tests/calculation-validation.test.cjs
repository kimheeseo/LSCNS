'use strict';
const {test}=require('node:test'),assert=require('node:assert/strict');
const E=require('../engineering-v82.js'),C=require('../calculation-definitions.js'),V=require('../design-validation.js'),I=require('../calculation-evidence-i18n.js');
const base={systemId:'b200',targetGPU:1024,topology:'rail',coolingMode:'dlc',facilityEnabled:true,marginDb:.2};
const calculate=x=>E.calculate({...base,...x});
const rule=(r,id,c)=>{const found=V.evaluate(r,c).rules.find(x=>x.id===id);assert.ok(found,id);return found.status;};
const scenarios=[
 {name:'H200 / 256 GPU / Dual ToR',input:{systemId:'h200',targetGPU:256,topology:'dual',rackPowerKw:30,rackCoolingKw:30,coolingMode:'air'},expected:[32,16,22,32,16,0,256,1024,0,426.9,512.28,2,2,2]},
 {name:'B200 / 1024 GPU / Rail',input:{},expected:[128,32,37,32,16,0,1024,1024,0,1937.4,2324.88,5,3,3]},
 {name:'B300 / 4096 GPU / 3-tier Clos',input:{systemId:'b300',targetGPU:4096,topology:'clos3'},expected:[512,128,150,128,128,128,4096,4096,4096,8495,10194,16,8,7]}
];
for(const s of scenarios)test(s.name+' independent regression and evidence replay',()=>{
 const r=calculate(s.input),a=r.summary,f=r.fabric,p=r.facility;
 assert.deepEqual([a.servers,a.computeRacks,a.totalRacks,f.leafCount,f.spineCount,f.coreCount,f.serverLinks,f.leafSpineLinks,f.spineCoreLinks,a.totalItPowerKw,a.facilityPowerKw,p.upsQty,p.generatorQty,p.transformerQty],s.expected);
 assert.equal(r.errors.length,0);assert.equal(V.evaluate(r).counts.FAIL,0);assert.equal(V.evaluate(r).status,'WARN');
 assert.ok(V.evaluate(r).rules.some(x=>x.id==='qualification'&&x.status==='WARN'));
 assert.equal(a.computeRacks,Math.ceil(a.servers/a.systemsPerRack));
 assert.ok(Math.abs(r.racks.reduce((sum,x)=>sum+x.powerKw,0)-a.totalItPowerKw)<.001);
 const ids=new Set(r.evidence.steps.map(x=>x.id));assert.equal(ids.size,r.evidence.steps.length);
 for(const step of r.evidence.steps){
  assert.ok(Number.isFinite(step.value),step.id);
  assert.equal(C.evaluate(step.key,Object.fromEntries(step.inputs.map(x=>[x.name,x.value]))),step.value,step.id);
  assert.equal(step.expression,C.definitions[step.key].expression);assert.ok(step.unit!==undefined);assert.ok(step.inputs.every(x=>x.source));
  for(const input of step.inputs)if(input.ref)assert.ok(ids.has(input.ref),'missing reference '+input.ref);
 }
 for(const b of r.bom){assert.ok(ids.has(b.evidenceId));assert.equal(r.evidence.steps.find(x=>x.id===b.evidenceId).value,b.qty);}
 for(const check of V.evaluate(r).rules)for(const ref of check.refs)assert.ok(ids.has(ref),check.id+' / '+ref);
 assert.equal(JSON.parse(JSON.stringify(r.evidence)).kind,'calculated');
});
for(const [mode,expected]of [['N',[4,2,2]],['N+1',[5,3,3]],['2N',[8,4,4]]])test(mode+' facility surviving capacity and manual shortage',()=>{
 const r=calculate({facilityRedundancy:mode});assert.deepEqual([r.facility.upsQty,r.facility.generatorQty,r.facility.transformerQty],expected);
 for(const kind of ['ups','generator','transformer']){assert.equal(rule(r,'facility.'+kind),'PASS');const bad=calculate({facilityRedundancy:mode,[kind+'InstalledQty']:expected[['ups','generator','transformer'].indexOf(kind)]-1});assert.equal(rule(bad,'facility.'+kind),'FAIL');}
});
test('UPS protection scope changes UPS load and facility units only',()=>{const it=calculate(),all=calculate({upsLoadScope:'facility'});assert.equal(all.facility.protectedLoadKw,all.summary.facilityPowerKw);assert.ok(Math.abs(all.facility.upsDesignKva/it.facility.upsDesignKva-it.input.pue)<.00001);assert.ok(all.facility.upsQty>it.facility.upsQty);assert.equal(all.facility.generatorQty,it.facility.generatorQty);});
test('facility disabled and malformed redundancy never pass',()=>{assert.equal(rule(calculate({facilityEnabled:false}),'facility.ups'),'WARN');assert.equal(rule(calculate({facilityRedundancy:'unknown'}),'facility.ups'),'FAIL');});
test('cooling and actual rack power failures',()=>{const r=calculate({rackPowerKw:10,rackCoolingKw:10});assert.equal(rule(r,'rackCooling'),'FAIL');assert.equal(rule(r,'rackPower'),'FAIL');});
test('air and rear-door thresholds, liquid mode not compared against Infinity',()=>{assert.equal(rule(calculate({coolingMode:'air'}),'coolingMethod'),'WARN');assert.equal(rule(calculate({coolingMode:'rear-door'}),'coolingMethod'),'WARN');assert.equal(rule(calculate(),'coolingMethod'),'PASS');});
test('radix overflow is FAIL; quantized 4:1 oversubscription is WARN',()=>{assert.equal(rule(calculate({switchRadix:4,topology:'clos3'}),'topologyRadix'),'FAIL');assert.equal(rule(calculate({oversub:4}),'oversubscription'),'WARN');assert.equal(rule(calculate({oversub:4}),'leafPorts'),'PASS');});
test('2-tier automatic promotion is exposed',()=>{const r=calculate({targetGPU:32776});assert.equal(rule(r,'topologyRadix'),'WARN');assert.ok(r.fabric.domains.every(d=>d.tier===3));});
for(const [media,speed,reach]of [['dac',400,3],['dac',800,2],['aoc',400,30],['sr',400,100],['dr',400,500],['fr',400,2000]])test(media+' '+speed+' reach uses rounded purchased length',()=>{
 const x={systemId:speed===800?'b300':'b200',serverMedia:media,serverCabling:'p2p',verticalRiseM:0,terminationM:0,slackPct:0,serverDistanceM:reach};
 const exact=calculate(x);assert.equal(exact.optical.server.lengthM,reach);assert.notEqual(rule(exact,'reach.server'),'FAIL');
 const exceed=calculate({...x,serverDistanceM:reach+.01});assert.equal(rule(exceed,'reach.server'),'FAIL');assert.ok(exceed.optical.server.lengthM>reach);
});
test('named SKU reach still limits generic optical media',()=>{const r=calculate({serverOpticSku:'nvidia-mma4z00-ns400',serverDistanceM:60,serverMedia:'sr',mpoEndGender:'female'});assert.equal(r.optical.server.profile.reachM,50);assert.equal(rule(r,'reach.server'),'FAIL');});
test('connector optical loss budget failure',()=>{assert.equal(rule(calculate({serverMedia:'sr',mpoLossDb:5}),'loss.server'),'FAIL');});
test('requested undersized trunk is FAIL even when engine normalizes capacity',()=>{const r=calculate({trunkFiberCount:8});assert.equal(rule(r,'trunk.server'),'FAIL');assert.ok(r.optical.server.profile.trunkFiberCount>=12);});
test('incomplete trunk packing is WARN with correct per-route channel capacity',()=>assert.equal(rule(calculate({trunkFiberCount:32}),'trunk.server'),'WARN'));
test('independent route trunks cannot pool fibers; bundle and spare match BOM',()=>{const r=calculate({trunkFiberCount:96,trunkBundle:8});const o=r.optical.server;assert.equal(o.trunk.installedCableCount,256);assert.equal(o.trunk.cableCount,288);const bom=r.bom.find(x=>x.segment==='server'&&/trunk/.test(x.item));assert.equal(bom.qty,o.trunk.cableCount);assert.equal(rule(r,'trunkSpare.server'),'PASS');assert.ok(r.evidence.steps.some(x=>x.key==='spareUnits'));assert.ok(r.evidence.steps.some(x=>x.key==='withSpare'));});
test('insufficient spare detected and threshold configurable',()=>{const r=calculate({sparePct:2});assert.equal(rule(r,'trunkSpare.server'),'WARN');assert.equal(rule(r,'trunkSpare.server',{minimumSparePct:1}),'PASS');r.optical.server.trunk.cableCount=r.optical.server.trunk.installedCableCount;assert.equal(rule(r,'trunkSpare.server'),'FAIL');});
test('storage capacity and bandwidth affect IT and facility loads',()=>{const r=calculate({storageEnabled:true,storageCapacityTb:5000});assert.equal(r.storage.systems,7);assert.equal(r.storage.switches,2);assert.equal(r.storage.usableTb,5600);assert.equal(r.summary.totalItPowerKw,2010.4);assert.equal(r.summary.totalRacks,39);});
test('inputs record defaults, explicit inputs and normalized original values',()=>{const r=E.calculate({systemId:'b200',topology:'clos3',switchRadix:63,auxOversub:0,storageSwitchPorts:0});assert.equal(r.evidence.rawInput.topology,'clos3');assert.equal(r.evidence.normalizedInput.topology,'rail');assert.equal(r.evidence.inputSources.switchRadix.value,62);assert.equal(r.evidence.inputSources.switchRadix.normalized,true);assert.equal(r.evidence.inputSources.auxOversub.original,0);const a=r.evidence.steps.find(x=>x.id==='rack.limitPower').inputs;assert.equal(a.find(x=>x.name==='rackCapacity').assumption,true);assert.equal(r.evidence.inputSources.systemId.origin,'input');});
test('rack-scale and power-input calculations remain traceable',()=>{for(const x of [{systemId:'gb200',targetGPU:288,rackPowerKw:150,rackCoolingKw:150},{inputMode:'power',itInputMw:1}]){const r=calculate(x);assert.ok(r.evidence.steps.some(x=>x.id==='compute.units'));assert.equal(rule(r,'bomQuantities'),'PASS');}});
test('every new translation, rule reason/action interpolates in five languages',()=>{const keys=Object.keys(I.packs.en);assert.deepEqual(I.languages,['ko','ja','en','zh','de']);for(const lang of I.languages){assert.deepEqual(Object.keys(I.packs[lang]),keys);for(const key of keys)assert.ok(I.packs[lang][key].length,key);for(const row of V.evaluate(calculate()).rules){for(const key of [row.reason,row.suggestion]){assert.ok(keys.includes(key),key);assert.ok(!I.translate(lang,key,row.values).includes('{'));}}}});
test('all browser engine entry points load shared dependencies before the engine',()=>{const fs=require('node:fs'),path=require('node:path');for(const file of ['index.html','engineering-preview.html','customer-designer.html']){const html=fs.readFileSync(path.join(__dirname,'..',file),'utf8'),engine=html.indexOf('src="./engineering-v82.js');for(const dep of ['design-validation-config.js','calculation-definitions.js']){const pos=html.indexOf('src="./'+dep);assert.ok(pos>=0&&pos<engine,file+' / '+dep);}}});
