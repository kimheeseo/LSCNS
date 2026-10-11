const test=require('node:test'),assert=require('node:assert/strict');
const {calculate}=require('../cable-power.js');
test('passive fiber remains zero; missing active-device W is unknown',()=>{
 assert.equal(calculate({modules:0,assemblies:0,hours:24}).totalW,0);
 assert.throws(()=>calculate({modules:2,moduleW:'',assemblies:0,hours:24}));
});
test('physical modules, whole active assemblies and kWh use separate counting boundaries',()=>{
 const e=calculate({modules:2,moduleW:9,assemblies:0,hours:24});assert.equal(e.totalW,18);assert.equal(e.kw,0.018);assert.ok(Math.abs(e.kwh-0.432)<1e-12);
 const r=calculate({modules:3,moduleW:17,assemblies:2,assemblyW:8,hours:100});assert.equal(r.totalW,67);assert.equal(r.kwh,6.7);
});
test('fractional counts, negative values and invalid quantities do not produce results',()=>{
 for(const modules of [null,'',-1,1.5,Infinity])assert.throws(()=>calculate({modules,moduleW:9,assemblies:0,hours:24}));
 assert.throws(()=>calculate({modules:0,assemblies:2,assemblyW:-3,hours:24}));
 assert.throws(()=>calculate({modules:0,assemblies:0,hours:''}));
});
