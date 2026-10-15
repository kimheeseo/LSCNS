// STEP 15 source-level unit/integration regression; run: node --test DCI/DataCenter/ls3d-three/step15.unit.test.mjs
// Simulated THREE objects check scene hierarchy and transforms, not actual GPU rasterization.
import test from 'node:test';
import assert from 'node:assert/strict';
import {createMMRRoom} from './mmr-room.mjs';
import {createZoneSpecializations} from './zone-specializations.mjs';

class Group{
  constructor(){
    this.children=[];this.name='';this.userData={};this.visible=true;
    this.rotation={x:0,y:0,z:0};
    this.position={x:0,y:0,z:0,set(x,y,z){Object.assign(this,{x,y,z});}};
  }
  add(...objects){for(const obj of objects){obj.parent=this;this.children.push(obj);}return this;}
  remove(obj){const i=this.children.indexOf(obj);if(i>=0)this.children.splice(i,1);obj.parent=null;}
  removeFromParent(){this.parent?.remove(this);}
  traverse(visitor){visitor(this);for(const child of this.children)child.traverse(visitor);}
}
class Vector3{constructor(x,y,z){Object.assign(this,{x,y,z});}}
class CatmullRomCurve3{constructor(points){this.points=points;}}
class TubeGeometry{
  constructor(path){this.parameters={path};this.disposed=false;}
  dispose(){this.disposed=true;}
}
class MeshStandardMaterial{
  constructor(properties){Object.assign(this,properties);}
  dispose(){this.disposed=true;}
}
class Mesh extends Group{
  constructor(geometry,material){super();this.geometry=geometry;this.material=material;this.isMesh=true;}
}
const THREE={Group,Vector3,CatmullRomCurve3,TubeGeometry,MeshStandardMaterial,Mesh};
function cube(parent,x,y,z,w,h,d,color){
  const mesh=new Mesh(null,{color});mesh.position.set(x,y,z);
  mesh.scale={x:w,y:h,z:d};parent.add(mesh);return mesh;
}
const zone={id:'network',pos:[44,39],size:[28,20]};
const assets=['carrier','mmr','odf','core'].map((id,i)=>({
  id,zone:'network',x:44+[-9,-3,3,9][i],z:35,w:4,d:3,h:5
}));
function fixture(){const mmr=createMMRRoom(THREE,{cube});mmr.build(zone,assets,'aidc',{});return mmr;}
const all=root=>{const out=[];root.traverse(o=>out.push(o));return out;};

test('4 independent MMR rooms align to source racks, have access doors and differentiated rack panels',()=>{
  const mmr=fixture(),s=mmr.summary(),models=new Set(all(mmr.root).map(o=>o.name));
  assert.equal(s.ready,true);assert.equal(s.bays.length,4);assert.equal(s.doors.length,4);
  assert.equal(s.stats.patchPorts,128);assert.ok(s.stats.modules>=12);
  assert.equal(mmr.root.position.x,zone.pos[0]);assert.equal(mmr.root.position.z,zone.pos[1]);
  for(const n of ['carrier-demarc-plate','dwdm-muxponder-0','odf-patch-shelf-0',
    'core-interface-card-0'])assert.ok(models.has('mmr:'+n),n);
  mmr.dispose();
});

test('room doors swing, and selectable optical cable highlights retain 3D topology',()=>{
  const mmr=fixture(),before=mmr.summary().doors[0];
  assert.equal(before.opening,0);assert.equal(mmr.toggleDoor('carrier'),1);
  mmr.tick(.1);
  assert.ok(mmr.summary().doors.find(d=>d.id==='carrier').opening>.1);
  const door=all(mmr.root).find(o=>o.name==='mmr:door-hinge:carrier');
  assert.ok(door.rotation.y<-.1);
  mmr.selectLink('odf-dwdm');
  const info=mmr.summary();
  assert.equal(info.selectedRoute.from,'ODF Cross-connect');
  assert.equal(info.selectedRoute.to,'DWDM/OTN Handoff');
  assert.equal(info.links.filter(l=>!l.id.startsWith('vertical-drop-')).length,4);
  assert.equal(info.links.find(l=>l.id==='odf-dwdm').emphasized,true);
  assert.equal(info.links.find(l=>l.id==='carrier-odf').emphasized,false);
  assert.ok(info.links.every(l=>l.visible));
  const selected=all(mmr.root).find(o=>o.name==='mmr:fiber:odf-dwdm');
  const dimmed=all(mmr.root).find(o=>o.name==='mmr:fiber:carrier-odf');
  assert.equal(selected.userData.mmrSegment,'odf-dwdm');
  assert.ok(selected.material.opacity>.9);
  assert.ok(dimmed.material.opacity<.2);
  mmr.dispose();
});

test('scene rebuild preserves selected link, releases old geometries, and server-room mode isolates MMR',()=>{
  const mmr=fixture();
  mmr.selectLink('carrier-odf');
  const geometry=all(mmr.root).find(o=>o.name==='mmr:fiber:external-carrier').geometry;
  mmr.build(zone,assets,'aidc',{showSection:true});
  assert.equal(geometry.disposed,true);
  assert.equal(mmr.summary().selectedLink,'carrier-odf');
  assert.equal(mmr.summary().links.find(l=>l.id==='dwdm-core').emphasized,false);
  mmr.build(zone,assets,'server-room',{});
  assert.equal(mmr.root.visible,false);
  assert.equal(mmr.summary().ready,false);
  mmr.dispose();assert.equal(mmr.root.children.length,0);
});

test('MMR zone scene bridge identifies a picked 3D fiber path without touching host rack/BOM',()=>{
  const T={
    Z:[zone],facilityMode:'aidc',currentZone:'network',assetMap:new Map(assets.map(a=>[a.id,a])),
    visibleAssets:()=>assets,pickAsset(a){T.selected=a}
  };
  const adapter=createZoneSpecializations(THREE,T,{cube,world:new Group(),canvas:{parentElement:null}});
  adapter.update({},.016);
  assert.equal(adapter.mmr.root.visible,true);
  const cable=all(adapter.mmr.root).find(o=>o.name==='mmr:fiber:carrier-odf');
  const result=adapter.pick(cable);
  assert.equal(result.kind,'mmr-link');assert.equal(result.assetId,'odf');
  assert.equal(adapter.mmr.summary().selectedLink,'carrier-odf');
  T.currentZone='campus';adapter.update({},.016);
  assert.equal(adapter.mmr.root.visible,false);
  T.currentZone='network';adapter.update({},.016);
  assert.equal(adapter.mmr.summary().selectedLink,'carrier-odf');
  T.facilityMode='server-room';adapter.update({},.016);
  assert.equal(adapter.mmr.root.visible,false);
  adapter.dispose();
});
