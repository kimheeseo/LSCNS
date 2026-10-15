// STEP 17: standalone, deterministic Three.js-compatible MEP geometry tests.
// Run: node --test DCI/DataCenter/ls3d-three/step17.unit.test.mjs
// This is geometry/source-level validation; it does not execute GPU/WebGL/Chromium.
import test from 'node:test';
import assert from 'node:assert/strict';
import {createCampusMEPRoutes} from './campus-mep-routes.mjs';
class Group{
  constructor(){this.children=[];this.name='';this.visible=true;this.userData={};this.position={x:0,y:0,z:0,set(x,y,z){Object.assign(this,{x,y,z});}};}
  add(o){this.children.push(o);o.parent=this;return o;}
  remove(o){const n=this.children.indexOf(o);if(n>=0)this.children.splice(n,1);o.parent=null;}
  traverse(fn){fn(this);for(const child of this.children)child.traverse(fn);}
}
class Vector3{constructor(x,y,z){Object.assign(this,{x,y,z});}}
class CatmullRomCurve3{constructor(points){this.points=points;}}
class TubeGeometry{
  constructor(path){this.parameters={path};this.disposed=false;}
  dispose(){this.disposed=true;}
}
class MeshStandardMaterial{
  constructor(opts){Object.assign(this,opts);}
  dispose(){this.disposed=true;}
}
class Mesh extends Group{
  constructor(geom,mat){super();this.geometry=geom;this.material=mat;this.isMesh=true;}
}
const THREE={Group,Vector3,CatmullRomCurve3,TubeGeometry,MeshStandardMaterial,Mesh};
function cube(parent,x,y,z,w,h,d,color){
  const mesh=new Mesh(null,{color});mesh.position.set(x,y,z);mesh.scale={x:w,y:h,z:d};
  parent.add(mesh);return mesh;
}
const points=[
  ['xfmr-a','utility',-55,-31,5],['switchgear','utility',-39,-31,4],
  ['ups-a','gray',-36,2,5],['pdu','gray',-35,11,4],
  ['tor-a','hall',22,-2,6],
  ['tower-a','cooling',39,-31,8],['chiller','cooling',57,-31,4],
  ['pump','cooling',49,-20,3],['cdu-plant','cooling',57,-20,3],
  ['gray-cdu','gray',-26,11,4],['gpu-0-0','hall',0,-2,6],
  ['crah-a','hall',10,20,4],
  ['carrier','network',35,35,5],['mmr','network',41,35,5],
  ['odf','network',47,35,5],['core','network',53,35,5],
  ['spine-a','network',36,44,5],['leaf-a','network',48,44,5],
  ['mmr-hall','hall',26.3,13.5,6],['odf-hall','hall',3,20,5],
  ['colo-rack-a','hall',27.8,19.6,6],['colo-rack-b','hall',30.3,19.6,6],
  ['colo-rack-c','hall',32.8,19.6,6]
];
const assets=points.map(([id,zone,x,z,h])=>({id,name:id,zone,x,z,h,w:3,d:3}));
function fixture(){
  const T={facilityMode:'aidc',currentZone:'campus',visibleAssets:()=>assets};
  const root=new Group(),mod=createCampusMEPRoutes(THREE,T,{cube,root,canvas:{parentElement:null}});
  return {T,root,mod};
}
const all=root=>{const r=[];root.traverse(o=>r.push(o));return r;};

test('world links are anchored to real asset IDs; every system has 3D curves and no invented specs',()=>{
  const {mod,root}=fixture(),state=mod.update({flow:{power:true,cooling:true,fiber:true}});
  assert.ok(state.total>=17);
  assert.deepEqual(new Set(state.paths.map(x=>x.system)),new Set(['fiber','power','cooling']));
  const ids=new Set(assets.map(a=>a.id));
  assert.ok(state.paths.every(p=>ids.has(p.source)&&ids.has(p.target)&&
    p.lengthMeters===null&&p.verified===false));
  const tubes=all(root).filter(o=>o.name.startsWith('mep:tube:'));
  assert.equal(tubes.length,state.total);
  assert.ok(tubes.every(t=>t.geometry instanceof TubeGeometry&&t.userData.mepRouteId));
  assert.ok(state.stats.traySegments>0&&state.stats.pipeClamps>0);
  mod.dispose();
});

test('MMR → Core → Leaf → ToR, and MMR handoff to tenant A/B/C are represented explicitly',()=>{
  const {mod}=fixture(),paths=mod.update({}).paths.map(p=>p.id);
  for(const id of ['fiber:carrier:mmr','fiber:mmr:odf','fiber:odf:core',
    'fiber:core:spine-a','fiber:spine-a:leaf-a','fiber:leaf-a:tor-a',
    'fiber:odf:mmr-hall','fiber:mmr-hall:odf-hall',
    'fiber:mmr-hall:colo-rack-a','fiber:mmr-hall:colo-rack-b',
    'fiber:mmr-hall:colo-rack-c'])assert.ok(paths.includes(id),id);
  mod.dispose();
});

test('raycast a route tube, retain all others as faded 3D objects, restore power filter',()=>{
  const {mod,root}=fixture();mod.update({});
  const chosen='fiber:carrier:mmr';
  const mesh=all(root).find(o=>o.name==='mep:tube:'+chosen);
  const picked=mod.pick(mesh);
  assert.equal(picked.id,chosen);
  assert.equal(picked.source,'carrier');
  assert.equal(mod.summary().paths.filter(p=>p.highlighted).length,1);
  assert.ok(mod.records.every(r=>r.group.visible));
  assert.ok(all(root).filter(o=>o.name.startsWith('mep:tube:')&&o!==mesh)
    .every(o=>o.material.opacity===.09));
  mod.selectSystem('power');
  assert.ok(mod.summary().paths.every(p=>p.highlighted===(p.system==='power')));
  mod.dispose();
});

test('route rebuilds are cached and source positions / old curves do not leak',()=>{
  const {mod,T,root}=fixture(),on={flow:{power:true,cooling:true,fiber:true}};
  const a=mod.update(on),oldTube=all(root).find(o=>o.name.startsWith('mep:tube:'));
  assert.equal(mod.update(on).builds,a.builds);
  T.currentZone='hall';
  const b=mod.update(on);
  assert.ok(b.paths.every(p=>p.fromZone==='hall'||p.toZone==='hall'));
  assert.ok(oldTube.geometry.disposed&&oldTube.material.disposed);
  T.currentZone='campus';mod.update({flow:{power:false,cooling:true,fiber:true}});
  assert.ok(mod.summary().paths.every(p=>p.system!=='power'));
  const newTube=all(root).find(o=>o.name.startsWith('mep:tube:'));
  mod.dispose();assert.equal(root.children.length,0);
  assert.equal(newTube.geometry.disposed,true);
});

test('server-room and rack inspector hide campus routes without mutating host source assets',()=>{
  const {mod,T,root}=fixture();
  const before=JSON.stringify(assets),state=mod.update({});
  assert.ok(state.visible);
  T.facilityMode='server-room';
  assert.equal(mod.update({}).visible,false);assert.equal(root.visible,false);
  T.facilityMode='aidc';
  assert.equal(mod.update({inspect:true}).visible,false);
  assert.equal(JSON.stringify(assets),before);
  mod.dispose();
});
