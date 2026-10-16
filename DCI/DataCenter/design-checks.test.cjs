const test=require('node:test');
const assert=require('node:assert/strict');
const E=require('./engineering-v82.js');
const C=require('./design-checks.js');
test('default example: physical limits pass, air planning threshold remains visible',()=>{
 const r=E.calculate({}),c=C.evaluate(r);
 assert.equal(r.summary.targetGPU,1024);assert.equal(r.summary.totalItPowerKw,1937.4);
 assert.equal(c.metrics.filter(m=>m.count).length,0);
 assert(c.issues.some(x=>x.key==='cooling-method'));
});
test('rack-scale power, cooling and RU exceedances are all reported',()=>{
 const c=C.evaluate(E.calculate({systemId:'gb200',inputMode:'racks',nvlRackCount:4,rackPowerKw:50,rackCoolingKw:40,rackRU:20}));
 assert.deepEqual(c.metrics.filter(m=>m.count).map(m=>m.key),['power','cooling','ru']);
 assert(c.issues.find(x=>x.key==='ru').racks.includes('compute-1'));
});
test('OEM rack RU is not double-counted with generic reserved RU',()=>{
 const c=C.evaluate(E.calculate({systemId:'gb200',inputMode:'racks',nvlRackCount:4,rackPowerKw:150,rackCoolingKw:150,rackRU:48}));
 assert.equal(c.metrics.find(m=>m.key==='ru').count,0);
});
test('headroom shortfall is a planning warning, not a physical overload',()=>{
 const r=E.calculate({});r.racks=[{id:'test',role:'Network',powerKw:90,ru:10}];
 const c=C.evaluate(r);assert.equal(c.metrics[0].count,0);
 assert(c.issues.some(x=>x.key==='power-headroom'&&x.level==='warning'));
});
test('missing rack schedule cannot be labelled a verified pass',()=>{
 const r=E.calculate({});delete r.racks;
 assert(C.evaluate(r).metrics.every(m=>!m.known));
});
test('failed calculation carries its error and no stale metrics',()=>{
 const c=C.evaluate({usable:false,error:'invalid input'});
 assert.equal(c.valid,false);assert.equal(c.metrics.length,0);assert.deepEqual(c.errors,['invalid input']);
});
