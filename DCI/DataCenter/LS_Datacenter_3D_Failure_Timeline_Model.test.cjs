// Run: node --test DCI/DataCenter/LS_Datacenter_3D_Failure_Timeline_Model.test.cjs
const test=require('node:test');
const assert=require('node:assert/strict');
const model=require('./LS_Datacenter_3D_Failure_Timeline_Model.js');
const cfg={rackCount:20,rackPowerKw:60,portCount:128,plugW:12,plugKw:2,mode:'pluggable',
 coolingKw:1600,batteryMin:12,upsKw:1600,generatorKw:1800,dlc:60,gpuRatio:80,baselineRackInletC:22,outdoorC:24};
test('Utility: UPS bridges 18 s and generator/ATS takeover preserves service',()=>{
 assert.equal(model.point('power','N',0,cfg).state,model.STATE.NORMAL);
 assert.match(model.point('power','N',18,cfg).event,/ATS/);
 assert.equal(model.summarize('power','N',cfg).downtimeSec,0);
});
test('Generator start failure: battery countdown reaches outage at configured threshold',()=>{
 const p=model.point('ups-battery','2N',cfg.batteryMin*60+2,cfg);
 assert.equal(p.state,model.STATE.OUTAGE);
 assert.equal(p.upsRemaining,0);
 const s=model.summarize('ups-battery','N',cfg);
 assert.equal(s.firstOutageSec,720);
});
test('Cooling: N > N+1 > 2N temperature rise, threshold reach times and states',()=>{
 const n=model.summarize('cooling','N',cfg),spare=model.summarize('cooling','N+1',cfg),
  twice=model.summarize('cooling','2N',cfg);
 assert.ok(n.maxTemperature>spare.maxTemperature);
 assert.ok(spare.maxTemperature>twice.maxTemperature);
 assert.ok(n.firstThrottleSec!==null);
 assert.ok(n.firstOutageSec!==null);
 assert.equal(twice.state,model.STATE.NORMAL);
});
test('DLC leak: impact is constrained to eligible DLC rack count',()=>{
 const n=model.summarize('dlc-leak','N',cfg),twice=model.summarize('dlc-leak','2N',cfg);
 assert.ok(n.maxAffectedRacks>twice.maxAffectedRacks);
 assert.ok(n.maxAffectedRacks<=Math.ceil(cfg.rackCount*cfg.dlc/100));
 assert.equal(n.state,model.STATE.DEGRADED);
});
test('Fire: mandatory safety-zone isolation stops affected service even with 2N',()=>{
 for(const r of model.modes){
  assert.equal(model.point('fire',r,150,cfg).state,model.STATE.OUTAGE);
  assert.equal(model.point('fire',r,370,cfg).state,model.STATE.NORMAL);
 }
});
test('GPU job surge: model has explicit power and thermal margins for redundancy',()=>{
 const n=model.summarize('gpu-surge','N',cfg),twice=model.summarize('gpu-surge','2N',cfg);
 assert.ok(n.maxTemperature>twice.maxTemperature);
 assert.ok(n.degradedSec>=twice.degradedSec);
 assert.ok(n.maxDemandKw>n.baseITKw);
 assert.equal(model.point('gpu-surge','2N',750,cfg).state,model.STATE.NORMAL);
});
test('Curves contain monotonic time and all numerical outputs are finite',()=>{
 for(const sc of ['power','cooling','dlc-leak','fire','gpu-surge','ups-battery']){
  for(const mode of model.modes){
   const c=model.curve(sc,mode,cfg),s=model.summarize(sc,mode,cfg);
   assert.equal(c.at(0).t,0);assert.equal(c.at(-1).t,s.horizonSec);
   for(let i=0;i<c.length;i++){
    assert.ok(Number.isFinite(c[i].temperature)&&Number.isFinite(c[i].itLoad));
    assert.ok(c[i].serviceFraction>=0&&c[i].serviceFraction<=1);
    if(i)assert.ok(c[i].t>c[i-1].t);
   }
   assert.ok(s.downtimeSec>=0&&s.downtimeSec<=s.horizonSec);
  }
 }
});
test('Assumption disclosures contain numerical parameters and explanatory limits',()=>{
 const labels=model.scenarioAssumptions('power',cfg).map(row=>row[0]);
 for(const required of ['시간 간격','기동 및 ATS','UPS 자율시간','현재 기준 IT 부하','가정 랙 입구온도'])
  assert.ok(labels.includes(required));
});
