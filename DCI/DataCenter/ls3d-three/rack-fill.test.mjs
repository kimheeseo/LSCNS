import test from 'node:test';import assert from 'node:assert/strict';
import {planRackFrontFill,createRackComponentKit} from './rack-components.mjs';
import {createRackServerRails} from './rack-server-rails.mjs';
import {createRackNetworkCabling} from './rack-network-cabling.mjs';
class Group{constructor(){this.children=[];this.userData={};this.position={set(x,y,z){Object.assign(this,{x,y,z})}};this.rotation={};}add(o){this.children.push(o);o.parent=this;}remove(o){this.children=this.children.filter(x=>x!==o);}traverse(fn){fn(this);this.children.forEach(o=>o.traverse?o.traverse(fn):fn(o));}}
function cube(parent,x,y,z,w,h,d,color){const m={isMesh:true,userData:{},position:{x,y,z},scale:{x:w,y:h,z:d},name:'',material:{color}};parent.add(m);return m;}
for(const [id,type,units,servers] of [['tor-a','Top-of-Rack switch',48,3],['odf','ODF',42,3],['gpu-0-0','GPU rack',48,4],['storage','Storage rack',42,7],['compute','Compute rack',42,6]]){
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
test('GPU rack modules can be selected, extracted and reinstalled independently',()=>{
 const asset={id:'gpu-0-1',type:'GPU server rack',w:3,d:3,h:6};
 const root=new Group(),kit=createRackComponentKit({Group},{cube}),built=kit.build(root,asset);
 const rails=createRackServerRails({Group},{cube});rails.decorate(asset,built,root);
 assert.equal(built.serverModules,4);
 assert.equal(built.gpuGroups.length,16);
 const gpu=built.gpuGroups[5],other=built.gpuGroups[4],id=gpu.userData.rackPart.id;
 assert.equal(rails.selectGpu(asset.id,id).slot,6);
 assert.equal(rails.state(asset.id).selected,gpu.userData.serverPartId);
 assert.equal(rails.toggleGpu(asset.id).target,1);
 rails.update(asset.id,built.moving,{slide:1},.2);
 assert.ok(gpu.position.z<-asset.d*.5);
 assert.equal(other.position.z,0);
 assert.equal(rails.toggleGpu(asset.id).target,0);
 rails.update(asset.id,built.moving,{slide:1},.2);
 assert.equal(gpu.position.z,0);
});
test('GPU rack allocates two top Leaf switches above four server trays without overlapping the optical drawer',()=>{
 const a={id:'gpu-0-0',type:'GPU server rack',w:2.7,d:3,h:6};
 const built=createRackComponentKit({Group},{cube}).build(new Group(),a);
 const find=kind=>built.parts.find(p=>p.kind===kind);
 assert.equal(find('tor').label,'Leaf A · Top-of-Rack 스위치');
 assert.equal(find('leaf-b').label,'Leaf B · Top-of-Rack 스위치');
 assert.ok(find('tor').uStart>find('leaf-b').uStart);
 assert.ok(find('leaf-b').uStart>find('patch-panel').uStart);
 assert.ok(find('patch-panel').uStart>find('cassette-pull').uStart);
 assert.equal(built.cassetteU,42);
 assert.equal(built.serverModules,4);
 const cabling=createRackNetworkCabling({Group},{cube});
 cabling.decorate(a,built,new Group());
 assert.equal(cabling.diagnostic(a.id).links.filter(l=>l.to===find('leaf-b').id).length,4);
});
test('network rack exposes Spine 1/2 modules above its expansion and blanking space',()=>{
 const built=createRackComponentKit({Group},{cube}).build(new Group(),
   {id:'tor-a',type:'Top-of-Rack switch',w:2.7,d:3,h:6});
 const switchParts=built.parts.filter(p=>p.kind.startsWith('network-switch-'));
 assert.deepEqual(switchParts.map(p=>p.label),[
   'Spine 1 · 네트워크 스위치','Spine 2 · 네트워크 스위치','네트워크 확장 모듈']);
 assert.ok(switchParts[0].uStart>switchParts[1].uStart);
 assert.ok(switchParts[1].uStart>switchParts[2].uStart);
 assert.ok(built.frontFill.added.some(p=>p.kind==='horizontal-manager'));
 assert.equal(built.parts.filter(p=>p.kind.startsWith('optical-transceiver-')).length,4);
 assert.equal(built.parts.filter(p=>p.kind.startsWith('gpu-module-')||p.kind.startsWith('storage-drive-')).length,0);
});
test('storage drive bays are selectable and move independently of the other drives and rack',()=>{
 const asset={id:'storage',type:'Storage rack',w:3,d:3,h:6};
 const built=createRackComponentKit({Group},{cube}).build(new Group(),asset);
 const driveParts=built.parts.filter(p=>p.kind.startsWith('storage-drive-'));
 assert.equal(driveParts.length,28);
 assert.equal(built.parts.some(p=>p.kind.startsWith('gpu-module-')),false);
 const rails=createRackServerRails({Group},{cube});rails.decorate(asset,built,new Group());
 const first=built.storageDrives[0],second=built.storageDrives[1];
 assert.equal(rails.selectDrive(asset.id,first.userData.partId).slot,1);
 assert.equal(rails.toggleDrive(asset.id).target,1);
 rails.update(asset.id,built.moving,{slide:0},.2);
 assert.ok(first.position.z<0);
 assert.equal(second.position.z,0);
 assert.equal(rails.toggleDrive(asset.id).target,0);
 rails.update(asset.id,built.moving,{slide:0},.2);
 assert.equal(first.position.z,0);
});
test('compute rack exposes CPU and memory without GPU and storage-specific parts',()=>{
 const built=createRackComponentKit({Group},{cube}).build(new Group(),
   {id:'compute',type:'Compute rack',w:3,d:3,h:6});
 assert.equal(built.parts.filter(p=>p.kind.startsWith('compute-cpu-')).length,6);
 assert.equal(built.parts.filter(p=>p.kind.startsWith('compute-memory-')).length,6);
 assert.equal(built.parts.some(p=>p.kind.startsWith('gpu-module-')||p.kind.startsWith('storage-drive-')),false);
});
