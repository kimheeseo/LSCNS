import test from 'node:test';import assert from 'node:assert/strict';
import {planRackFrontFill,createRackComponentKit} from './rack-components.mjs';
class Group{constructor(){this.children=[];this.userData={};this.position={set(x,y,z){Object.assign(this,{x,y,z})}};this.rotation={};}add(o){this.children.push(o);o.parent=this;}remove(o){this.children=this.children.filter(x=>x!==o);}traverse(fn){fn(this);this.children.forEach(o=>o.traverse?o.traverse(fn):fn(o));}}
function cube(parent,x,y,z,w,h,d,color){const m={isMesh:true,userData:{},position:{x,y,z},scale:{x:w,y:h,z:d},name:'',material:{color}};parent.add(m);return m;}
for(const [id,type,units,servers] of [['tor-a','Top-of-Rack switch',48,3],['odf','ODF',42,3],['gpu-0-0','GPU rack',48,7],['storage','Storage rack',42,7],['compute','Compute rack',42,6]]){
 test(id+' has no unassigned front U and preserves active sled count',()=>{
  const kit=createRackComponentKit({Group},{cube}),root=new Group(),built=kit.build(root,{id,type,w:3,d:3,h:6});
  assert.equal(built.rackUnits,units);assert.equal(built.serverModules,servers);
  assert.equal(built.frontFill.occupiedUnits,units);assert.ok(built.frontFill.coverage.every(Boolean));
  const added=built.frontFill.added;
  assert.ok(added.some(p=>p.kind==='blanking-panel'));
  if(/switch|ODF/.test(type))assert.ok(added.some(p=>p.kind==='distribution-panel'));
  else assert.ok(!added.some(p=>p.kind==='distribution-panel'));
  for(const p of added){assert.ok(p.uStart>=1&&p.uStart+p.uHeight-1<=units);assert.ok(built.parts.some(q=>q.uStart===p.uStart&&q.kind.startsWith(p.kind)));}
  assert.equal(new Set(built.parts.map(p=>p.id)).size,built.parts.length);
 });
}
test('overlapping or out-of-range active allocations are rejected',()=>{
 assert.throws(()=>planRackFrontFill(42,[{uStart:1,uHeight:3},{uStart:3,uHeight:2}]),/Overlapping/);
 assert.throws(()=>planRackFrontFill(42,[{uStart:42,uHeight:2}]),/out-of-range/);
});
