const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const context={window:{}};vm.createContext(context);vm.runInContext(fs.readFileSync(__dirname+'/dc-engine.min.js','utf8'),context);vm.runInContext(fs.readFileSync(__dirname+'/planning-model.js','utf8'),context);
const model=context.window.DCPlanning,engine=context.DCBOMEngine;
for(const systemId of ['h200','b200','b300']){const r=engine.design({targetGPU:1024,systemId,trunkFiberCount:systemId==='b300'?24:16});assert.ok(r.usable);if(systemId==='b300'){assert.ok(Math.abs(model.checks(r).find(x=>x.label==='서버 최대 설계전력').errorPct-100*.5/14.5)<1e-8);}else assert.ok(model.checks(r).every(x=>x.errorPct===0));}
const r=engine.design({targetGPU:1024,systemId:'b200'}),totals=model.totals(r);
assert.equal(totals.routeLengthM,133120);assert.equal(totals.trunkLengthM,102400);assert.equal(totals.assemblyLengthM,30720);assert.equal(totals.installedFibers,16384);assert.equal(totals.activeFibers,8192);
const noSpare=engine.design({targetGPU:1024,systemId:'b200',sparePct:0});assert.equal(model.totals(noSpare).routeLengthM,totals.routeLengthM);assert.equal(model.totals(noSpare).installedFibers,totals.installedFibers);
const key=r.bom[0].category+' · '+r.bom[0].item;
assert.equal(model.quote(r,{KRW:{[key]:1000}},'KRW').subtotal,128000);assert.equal(model.quote(r,{KRW:{[key]:1000}},'USD').priced,0);assert.equal(model.quote(r,{KRW:{[key]:0}},'KRW').priced,1);assert.equal(model.quote(r,{KRW:{[key]:-5}},'KRW').priced,0);
r.systemProfile.power=15;assert.ok(model.checks(r).find(x=>x.label==='서버 최대 설계전력').errorPct>0);
assert.ok(Math.abs(model.regionalCapacity({siteVoltage:400,siteCurrentA:64,powerFactor:.9,sitePhases:3,siteFrequency:50})-39.906450606386934)<1e-8);
assert.equal(model.regionalCapacity({siteVoltage:230,siteCurrentA:32,powerFactor:1,sitePhases:1,siteFrequency:50}),7.36);
assert.throws(()=>model.regionalCapacity({siteVoltage:0,siteCurrentA:64,powerFactor:.9,sitePhases:3,siteFrequency:50}));
const colo=engine.design({scenario:'colo'});assert.equal(model.totals(colo).hasRoutes,false);assert.equal(model.checks(colo).length,0);
console.log('PASS: official server references, installed lengths/fibers, spare exclusion, partial quotes/currencies, nonzero error detection, regional PDU math, colocation scope.');
