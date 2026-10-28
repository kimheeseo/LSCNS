import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
const html=readFileSync(new URL('../LS_Datacenter_3D.html',import.meta.url),'utf8');
const assets=[];
const context={addAsset:(zone,id,name,type,x,z,w,d,h)=>assets.push({zone,id,x,z,w,d,h})};
for(const call of html.matchAll(/^addAsset\('(network|gray)'[^\n]+\);/gm))vm.runInNewContext(call[0],context);
const overlaps=(a,b)=>Math.abs(a.x-b.x)<(a.w+b.w)/2&&Math.abs(a.z-b.z)<(a.d+b.d)/2;
test('Carrier entry and cross-connect remain in Network/MMR with clear equipment footprints',()=>{
  for(const id of ['carrier-entry','cross-connect']){
    const a=assets.find(a=>a.id===id);assert.ok(a,id);assert.equal(a.zone,'network');
    assert.ok(Math.abs(a.x)+a.w/2<=14,id+' inside room width');
    assert.ok(Math.abs(a.z)+a.d/2<=10,id+' inside room depth');
    for(const b of assets.filter(b=>b.zone==='network'&&b.id!==id))assert.ok(!overlaps(a,b),id+' overlaps '+b.id);
  }
});
test('loading bay stays in Gray Zone while optical cabinets are excluded',()=>{
  assert.match(html,/\['gray','visitor-loading','LS CABLE · 하역장'/);
  assert.ok(!assets.some(a=>a.zone==='gray'&&['carrier-entry','cross-connect'].includes(a.id)));
});
