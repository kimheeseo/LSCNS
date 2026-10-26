// LS3D STEP16 -- standalone model tests. Run from repo root:
// node --test DCI/DataCenter/ls3d-three/step16.unit.test.mjs
// Geometry-level THREE mock, not a Chromium/WebGL renderer or FPS benchmark.
import test from 'node:test';
import assert from 'node:assert/strict';
import {createColocationCages} from './colocation-cages.mjs';
import {createZoneSpecializations} from './zone-specializations.mjs';
class Group {
  constructor(){this.children=[];this.name='';this.userData={};this.visible=true;
    this.rotation={x:0,y:0,z:0};
    this.position={x:0,y:0,z:0,set(x,y,z){Object.assign(this,{x,y,z});}};}
  add(...arr){for(const o of arr){o.parent=this;this.children.push(o);}return this;}
  remove(o){const i=this.children.indexOf(o);if(i>=0)this.children.splice(i,1);o.parent=null;}
  traverse(fn){fn(this);for(const o of this.children)o.traverse(fn);}
  removeFromParent(){this.parent?.remove(this);}
}
class Vector3 {constructor(x,y,z){Object.assign(this,{x,y,z});}}
class CatmullRomCurve3 {constructor(points){this.points=points;}}
class TubeGeometry {constructor(path){this.parameters={path};this.disposed=false;}dispose(){this.disposed=true;}}
class MeshStandardMaterial {constructor(o){Object.assign(this,o);}dispose(){this.disposed=true;}}
class Mesh extends Group {constructor(g,m){super();this.geometry=g;this.material=m;this.isMesh=true;}}
const THREE={Group,Vector3,CatmullRomCurve3,TubeGeometry,MeshStandardMaterial,Mesh};
function cube(parent,x,y,z,w,h,d,color){
  const mesh=new Mesh(null,{color});mesh.position.set(x,y,z);
  mesh.scale={x:w,y:h,z:d};parent.add(mesh);return mesh;
}
const zone={id:'hall',pos:[18,7],size:[50,42]};
// Mirror the 2x-area Client Rack zone and its wider tenant aisles.
const cage={id:'colo',zone:'hall',x:32.7,z:17.8,w:18.4,d:15,h:4};
const racks=[26.3,32.7,39.1].map((x,i)=>({
  id:'colo-rack-'+['a','b','c'][i],zone:'hall',x,z:17.8,w:2.2,d:2.6,h:6
}));
const assets=[cage,...racks];
function kit(){return createColocationCages(THREE,{cube});}
function traverse(root){const result=[];root.traverse(o=>result.push(o));return result;}

test('rack coordinates determine safe independent A/B/C fence boundaries and 3 mesh gates',()=>{
  const sim=kit(),s=sim.build(assets,'aidc',{});
  assert.equal(s.ready,true);assert.equal(s.layoutIssue,null);
  assert.equal(s.bays.length,3);assert.equal(s.stats.doors,3);
  assert.equal(s.stats.meshPanels,4);assert.ok(s.stats.meshWires>=150);
  const stile=traverse(sim.root).find(o=>o.name==='colo:mesh-stile-0-0');
  assert.ok(stile.position.y+stile.scale.y>racks[0].h);
  for(let i=0;i<3;i++){
    assert.ok(racks[i].x-racks[i].w/2>=cage.x+s.boundaries[i]);
    assert.ok(racks[i].x+racks[i].w/2<=cage.x+s.boundaries[i+1]);
  }
  assert.deepEqual(s.bays.map(x=>x.rackId),racks.map(x=>x.id));
  sim.dispose();
});

test('expanded client area has independent spacious A/B/C rack bays and service aisles',()=>{
  const sim=kit(),state=sim.build(assets,'aidc',{});
  assert.equal(state.ready,true);assert.equal(state.layoutIssue,null);
  assert.equal(cage.w*cage.d,276);
  assert.deepEqual(state.boundaries.map(x=>Number(x.toFixed(2))),[-9.2,-3.2,3.2,9.2]);
  assert.ok(state.boundaries.slice(1).every((x,i)=>x-state.boundaries[i]>=6-1e-8));
  for(let i=1;i<racks.length;i++)
    assert.ok(racks[i].x-racks[i].w/2-(racks[i-1].x+racks[i-1].w/2)>=4.2-1e-8);
  const front=cage.z+cage.d/2-Math.max(...racks.map(x=>x.z+x.d/2));
  const rear=Math.min(...racks.map(x=>x.z-x.d/2))-(cage.z-cage.d/2);
  assert.ok(front>=6.2-1e-8);assert.ok(rear>=6.2-1e-8);
  assert.ok(cage.x+cage.w/2<=zone.pos[0]+zone.size[0]/2);
  assert.ok(cage.z+cage.d/2<=zone.pos[1]+zone.size[1]/2);
  assert.equal(state.stats.doors,3);
  sim.dispose();
});

test('tenant-specific fiber and PDU A/B physically exist, are pickable, and have unverified metadata',()=>{
  const sim=kit(),s=sim.build(assets,'aidc',{});
  assert.equal(s.stats.fiberPaths,3);assert.equal(s.stats.powerPaths,6);
  const nodes=traverse(sim.root),routes=nodes.filter(x=>x.name.startsWith('colo:wire:'));
  assert.equal(routes.length,9);
  assert.ok(routes.every(x=>x.geometry instanceof TubeGeometry&&
    x.userData.coloTenant&&x.userData.coloSystem&&x.userData.noPick!==true));
  for(const t of ['A','B','C']){
    assert.equal(nodes.filter(o=>o.name.startsWith('colo:fiber-adapter-'+t)).length,8);
    for(const side of ['A','B'])
      assert.ok(nodes.some(o=>o.name==='colo:rear-pdu-feed-'+t+'-'+side));
    assert.equal(s.bays.find(x=>x.id===t).circuits.length,3);
    assert.ok(s.bays.find(x=>x.id===t).circuits.every(c=>c.verified===false));
  }
  sim.dispose();
});

test('filter dims unrelated cables while preserving all routes; door hinge survives rebuild',()=>{
  const sim=kit();sim.build(assets,'aidc',{});
  const powerB=traverse(sim.root).find(o=>o.name==='colo:wire:power-B-B');
  const fiberA=traverse(sim.root).find(o=>o.name==='colo:wire:fiber-A');
  const metadata=sim.selectCircuit('B','power-B'),s=sim.summary();
  assert.equal(metadata.kind,'power-B');
  assert.equal(s.selected,'B');assert.equal(s.selectedCircuit,'power-B');
  assert.ok(s.bays.flatMap(b=>b.circuits).every(c=>c.visible));
  assert.equal(s.bays.flatMap(b=>b.circuits).filter(c=>c.highlighted).length,1);
  assert.equal(powerB.material.opacity,.95);assert.equal(fiberA.material.opacity,.11);
  sim.toggleGate('B');sim.tick(.1);
  const opened=sim.summary().bays.find(b=>b.id==='B').gateOpening;
  const door=traverse(sim.root).find(o=>o.name==='colo:gate-hinge:B');
  assert.ok(opened>.1&&door.rotation.y<-.1);
  sim.build(assets,'aidc',{showSection:true});
  assert.equal(sim.summary().selectedCircuit,'power-B');
  assert.ok(sim.summary().bays.find(b=>b.id==='B').gateOpening>=opened-.001);
  sim.selectTenant('all');
  assert.ok(sim.summary().bays.flatMap(b=>b.circuits).every(c=>c.highlighted));
  sim.dispose();
});

test('overlapping rack footprints never cause fences to cut through customer racks',()=>{
  const sim=kit();sim.build(assets,'aidc',{});
  const oldFiber=traverse(sim.root).find(o=>o.name==='colo:wire:fiber-A').geometry;
  sim.build([cage,racks[0],{...racks[1],x:racks[0].x+1.5},racks[2]],'aidc',{});
  const s=sim.summary();
  assert.equal(s.ready,false);assert.ok(s.layoutIssue);
  assert.equal(sim.root.visible,false);assert.equal(oldFiber.disposed,true);
  sim.build(assets,'server-room',{});
  assert.equal(sim.root.visible,false);
  sim.dispose();
});

test('scene bridge resolves clicked tenant circuit, hides hall on navigation and preserves selection',()=>{
  const T={Z:[zone],facilityMode:'aidc',currentZone:'hall',assetMap:new Map(assets.map(a=>[a.id,a])),
    visibleAssets:()=>assets,pickAsset(a){T.selected=a}};
  const scene=new Group(),bridge=createZoneSpecializations(THREE,T,{cube,world:scene,canvas:{parentElement:null}});
  bridge.update({},.016);
  assert.equal(bridge.colocation.root.visible,true);
  const wire=traverse(bridge.colocation.root).find(o=>o.name==='colo:wire:power-A-C');
  const result=bridge.pick(wire);
  assert.equal(result.kind,'customer-link');assert.equal(result.id,'C');
  assert.equal(result.system,'power-A');assert.equal(result.assetId,'colo-rack-c');
  assert.equal(bridge.colocation.summary().selectedCircuit,'power-A');
  T.currentZone='network';bridge.update({},.016);
  assert.equal(bridge.colocation.root.visible,false);
  T.currentZone='hall';bridge.update({},.016);
  assert.equal(bridge.colocation.summary().selectedCircuit,'power-A');
  bridge.dispose();assert.equal(bridge.colocation.root.children.length,0);
});
