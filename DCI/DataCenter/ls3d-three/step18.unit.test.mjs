// STEP 18 — pure policy + camera/overlay mock regression.
// node --test DCI/DataCenter/ls3d-three/step18.unit.test.mjs
// This DOES NOT benchmark Chromium, GPU, FPS or pixel correctness.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createViewportPolicy,createCameraOcclusion,layoutSceneOverlays
} from './scene-viewport-policy.mjs';

class Vector3 {
  constructor(x=0,y=0,z=0){Object.assign(this,{x,y,z});}
  set(x,y,z){Object.assign(this,{x,y,z});return this;}
  subVectors(a,b){this.set(a.x-b.x,a.y-b.y,a.z-b.z);return this;}
  length(){return Math.hypot(this.x,this.y,this.z);}
  normalize(){const n=this.length()||1;this.x/=n;this.y/=n;this.z/=n;return this;}
}
class Raycaster {
  set(origin,direction){this.origin=origin;this.direction=direction;}
  intersectObjects(objects){return objects.filter(o=>o.userData.hit===true).map(object=>({object}));}
}
const THREE={Vector3,Raycaster};
class Node{
  constructor(name='',kind=null,hit=false){
    this.name=name;this.isMesh=kind!==null;this.userData={archKind:kind,hit};
    this.visible=true;this.children=[];
  }
  add(o){this.children.push(o);return this;}
  traverse(fn){fn(this);this.children.forEach(c=>c.traverse(fn));}
  updateWorldMatrix(){}
}
function shape(){
  const root=new Node('root'),room=new Node('arch:zone:hall');root.add(room);
  const wall=new Node('arch:wall:side','wall',true);
  const roof=new Node('arch:roof:main','roof',true);
  const floor=new Node('arch:floor:main','floor',true);
  room.add(wall).add(roof).add(floor);
  return {root,wall,roof,floor};
}
const selected={id:'gpu-0-0',zone:'hall',x:3,z:2,h:6};
test('adaptive render quality reacts only after sustained bad frames',()=>{
  const perf=createViewportPolicy();
  assert.equal(perf.settings({width:1200,quality:'medium',devicePixelRatio:2}).dpr,1.1);
  for(let i=0;i<10;i++)perf.sample(.065);
  assert.equal(perf.summary().degraded,false);
  for(let i=0;i<95;i++)perf.sample(.075);
  assert.equal(perf.summary().degraded,true);
  assert.ok(perf.settings({width:1200,quality:'medium',devicePixelRatio:2}).dpr<.9);
  assert.equal(perf.settings({width:390,quality:'medium',devicePixelRatio:3}).mobile,true);
  assert.equal(perf.settings({width:390,quality:'high',devicePixelRatio:3}).dpr,1.3);
  for(let i=0;i<260;i++)perf.sample(.014);
  assert.equal(perf.summary().degraded,false);
  assert.equal(perf.settings({width:390,quality:'medium',devicePixelRatio:3}).dpr,1);
});

test('render quality accepts bounded fractional dpr on low-density screens',()=>{
  const perf=createViewportPolicy();
  assert.equal(perf.settings({width:390,quality:'low',devicePixelRatio:3}).dpr,.8);
  assert.equal(perf.settings({width:1200,quality:'high',devicePixelRatio:2}).dpr,1.55);
  assert.equal(perf.settings({width:1200,quality:'medium',devicePixelRatio:.75}).dpr,.75);
  perf.sample(NaN);perf.sample(-3);perf.sample(30);
  assert.equal(perf.summary().sampleCount,0);
});

test('only architectural room wall/roof occluders vanish close to a selected Hall rack',()=>{
  const {root,wall,roof,floor}=shape(),manager=createCameraOcclusion(THREE,{zoneRoot:root});
  const state={eye:[5,5,11],selected,currentZone:'hall',facilityMode:'aidc'};
  manager.update(state);
  assert.equal(manager.summary().hiddenOccluders,2);
  assert.equal(wall.visible,false);assert.equal(roof.visible,false);
  assert.equal(floor.visible,true);
  // Do not flicker hidden walls during the in-between throttled frames.
  manager.update(state);assert.equal(wall.visible,false);
  manager.update({...state,eye:[55,15,72]});
  assert.equal(manager.summary().hiddenOccluders,0);
  assert.equal(wall.visible,true);assert.equal(roof.visible,true);
  manager.dispose();
});

test('section, inspect, company server-room and zoning always restore wall visibility',()=>{
  for(const config of [
    {showSection:true},{inspect:true},{facilityMode:'server-room'},
    {currentZone:'network'},{selected:null}
  ]){
    const {root,wall,roof}=shape(),manager=createCameraOcclusion(THREE,{zoneRoot:root});
    const state={eye:[5,5,11],selected,currentZone:'hall',facilityMode:'aidc'};
    manager.update(state);manager.update({...state,...config});
    assert.ok(wall.visible&&roof.visible);
    assert.equal(manager.summary().hiddenOccluders,0);
    manager.dispose();
  }
});

function panel(){return {style:{},hidden:false,dataset:{},children:[],
  querySelector(selector){return selector==='details'?this.details:null;}};}
test('mobile HUD moves zone and MEP panels into separate docks without permanently closing details',()=>{
  const zone=panel(),mep=panel(),rack=panel();
  mep.details={style:{},open:true};
  let result=layoutSceneOverlays({width:390,height:540,colo:zone,mep,rack});
  assert.equal(result.mobile,true);assert.equal(result.zonePanel,true);
  assert.equal(zone.style.left,'8px');
  assert.equal(mep.style.top,'auto');
  assert.equal(mep.style.bottom,'calc(32% + 16px)');
  assert.equal(mep.details.open,false);
  mep.details.open=true;
  layoutSceneOverlays({width:390,height:540,colo:zone,mep,rack});
  assert.equal(mep.details.open,true);
  zone.hidden=true;
  layoutSceneOverlays({width:1040,height:800,colo:zone,mep,rack});
  assert.equal(mep.style.top,'83px');
  assert.equal(rack.style.overflowY,'auto');
  assert.equal(mep.dataset.step18Collapsed,'0');
});
