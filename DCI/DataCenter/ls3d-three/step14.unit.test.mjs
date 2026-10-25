// STEP 14 isolated geometry tests: node --test DCI/DataCenter/ls3d-three/step14.unit.test.mjs
// Uses lightweight THREE-compatible objects; NOT a substitute for a real WebGL browser test.
import test from 'node:test';
import assert from 'node:assert/strict';
import {createRackNetworkCabling} from './rack-network-cabling.mjs';

class Group {
  constructor(){this.name='';this.children=[];this.position={z:0};this.userData={};this.visible=true;}
  add(o){this.children.push(o);o.parent=this;return this;}
}
class Vector3 {
  constructor(x,y,z){Object.assign(this,{x,y,z});}
}
class CatmullRomCurve3 {
  constructor(points){this.points=points;}
}
class TubeGeometry {
  constructor(path){this.parameters={path};this.disposed=false;}
  dispose(){this.disposed=true;}
}
class MeshStandardMaterial {
  constructor(opts){Object.assign(this,opts);}
}
class Mesh extends Group {
  constructor(geometry,material){super();this.geometry=geometry;this.material=material;this.isMesh=true;}
}
const THREE={Group,Vector3,CatmullRomCurve3,TubeGeometry,MeshStandardMaterial,Mesh};
function cube(parent,x,y,z,w,h,d,color){
  const mesh=new Mesh(null,{color});
  mesh.position={x,y,z};mesh.scale={x:w,y:h,z:d};
  parent.add(mesh);return mesh;
}
function fixture(units=48){
  const asset={id:'rack-A',w:3,d:3,h:6};
  const parts=[
    {id:'rack-A:tor',kind:'tor',uStart:units-1},
    {id:'rack-A:pdu-a',kind:'pdu-a'},
    {id:'rack-A:pdu-b',kind:'pdu-b'}
  ];
  const serverGroups=['server-1','server-2'].map((id,i)=>{
    const group=new Group();
    group.userData.rackPart={id,kind:'server-'+(i+1),uStart:4+i*5,uHeight:4};
    return group;
  });
  const built={rackUnits:units,parts,partMap:new Map(parts.map(x=>[x.id,x])),serverGroups};
  const high=new Group();
  const cabling=createRackNetworkCabling(THREE,{cube});
  cabling.decorate(asset,built,high);
  return {asset,built,high,cabling};
}
function cableGeometryFor(f,kind,server){
  const rec=f.cabling.diagnostic(f.asset.id);
  assert.ok(rec?.links.some(l=>l.kind===kind&&l.from===server));
  const assembly=f.high.children.find(o=>o.name.startsWith('rack-link:assembly'));
  const link=assembly.children.find(o=>o.userData.cableSystem===kind&&o.userData.from===server);
  assert.ok(link);
  return link.children.find(o=>o.isMesh).geometry;
}
test('rack ToR has eight independent selectable ports and 2× independent A/B power links',()=>{
  for(const units of [42,48]){
    const f=fixture(units);
    const ports=f.built.parts.filter(p=>p.kind==='tor-port');
    assert.equal(ports.length,8);
    assert.equal(new Set(ports.map(p=>p.id)).size,8);
    for(const p of ports)assert.equal(f.built.partMap.get(p.id),p);
    const diag=f.cabling.diagnostic(f.asset.id);
    assert.equal(diag.torU,units-1);
    assert.equal(diag.links.filter(x=>x.kind==='fiber').length,1);
    assert.equal(diag.links.filter(x=>x.kind==='dac').length,1);
    assert.equal(diag.links.filter(x=>x.kind==='power-A').length,2);
    assert.equal(diag.links.filter(x=>x.kind==='power-B').length,2);
    assert.ok(diag.links.every(x=>x.from&&x.to));
  }
});
test('filter keeps all routes visible and dims unrelated A/B/fiber routes',()=>{
  const f=fixture();
  f.cabling.update(f.asset.id,'server-2');
  f.cabling.focus(f.asset.id,'power-A');
  const diag=f.cabling.diagnostic(f.asset.id);
  assert.equal(diag.filter,'power-A');
  assert.equal(diag.links.filter(x=>x.emphasized).length,1);
  assert.ok(diag.links.every(x=>x.visible));
  assert.equal(diag.links.find(x=>x.kind==='power-A'&&x.from==='server-2').emphasized,true);
  const assembly=f.high.children.find(x=>x.name.startsWith('rack-link:assembly'));
  const links=assembly.children.filter(x=>x.userData.cableSystem);
  assert.ok(links.every(x=>{
    const expected=x.userData.cableSystem==='power-A'&&x.userData.from==='server-2'?.94:.11;
    return Math.abs(x.children[0].material.opacity-expected)<1e-9;
  }));
  f.cabling.focus(f.asset.id,'all');
  assert.equal(f.cabling.diagnostic(f.asset.id).links.filter(x=>x.emphasized).length,6);
});
test('pulled server moves NIC and its own cable source only',()=>{
  const f=fixture();
  const first='server-1',second='server-2';
  const before1=cableGeometryFor(f,'dac',first);
  const before2=cableGeometryFor(f,'fiber',second);
  const z1=before1.parameters.path.points[0].z;
  const z2=before2.parameters.path.points[0].z;
  f.built.serverGroups[1].position.z=-.62;
  f.cabling.update(f.asset.id,second);
  const after1=cableGeometryFor(f,'dac',first);
  const after2=cableGeometryFor(f,'fiber',second);
  assert.equal(after1,before1);
  assert.notEqual(after2,before2);
  assert.ok(before2.disposed);
  assert.ok(Math.abs(after2.parameters.path.points[0].z-z2+.62)<1e-9);
  assert.equal(after1.parameters.path.points[0].z,z1);
  assert.equal(f.cabling.stats.cableGeometryUpdates,3); // fiber + dual PDU cords
  f.built.serverGroups[1].position.z=0;
  f.cabling.update(f.asset.id,second);
  assert.ok(Math.abs(cableGeometryFor(f,'fiber',second).parameters.path.points[0].z-z2)<1e-9);
  assert.equal(f.cabling.stats.cableGeometryUpdates,6);
});
