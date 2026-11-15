import test from 'node:test';import assert from 'node:assert/strict';
import * as THREE from '../vendor/three.module.js';
import {createRackHardwareDetail} from './rack-hardware-detail.mjs';
import {industrialColor} from './industrial-palette.mjs';
import {createRackPanelDisclosure} from './rack-panel-disclosure.mjs';
import {mechanicalButtons} from './rack-mechanical-policy.mjs';
import {layoutSceneOverlays} from './scene-viewport-policy.mjs';
const geometry=new THREE.BoxGeometry(1,1,1),cache=new Map();
test('Door/cassette controls remain outside the collapsed details panel',()=>{
 const panel={style:{},dataset:{}},body={hidden:false,id:'details'},pinned={hidden:false};
 const attrs={},button={setAttribute(k,v){attrs[k]=v;}};
 const control=createRackPanelDisclosure({panel,body,button,pinned,collapsed:true});
 assert.equal(body.hidden,true);assert.equal(pinned.hidden,false);assert.equal(panel.style.maxHeight,'132px');
 for(const width of [390,1200]){layoutSceneOverlays({width,height:700,rack:panel});assert.equal(panel.style.maxHeight,'132px');}
 panel.dataset.rackMechanicalRows='2';
 for(const width of [390,1200]){layoutSceneOverlays({width,height:700,rack:panel});assert.equal(panel.style.maxHeight,'184px');}
 assert.equal(attrs['aria-expanded'],'false');control.toggle();assert.equal(body.hidden,false);assert.equal(pinned.hidden,false);
 control.toggle();assert.equal(body.hidden,true);assert.equal(pinned.hidden,false);
});
function material(c){if(!cache.has(c))cache.set(c,new THREE.MeshStandardMaterial({color:c}));return cache.get(c);}
function cube(g,x,y,z,w,h,d,c){const m=new THREE.Mesh(geometry,material(c));m.position.set(x,y+h/2,z);m.scale.set(w,h,d);g.add(m);return m;}
test('Neutral finishes preserve cable and operational signal colors',()=>{
 assert.equal(industrialColor('front-door-opaque-steel-panel','#263b4d','rack'),'#dce0e1');
 assert.equal(industrialColor('board','#000000'),'#287858');
 assert.equal(industrialColor('coolant-pipe','#28ccef'),'#28ccef');
 assert.equal(industrialColor('status-led','#e74747'),'#e74747');
});
for(const [type,kind,units] of [['GPU server rack','gpu',48],['Compute rack','compute',42],['Storage rack','storage',42],['Top-of-Rack switch','network',48]]){
 test(type+' exposes and restores one tray without changing other modules',()=>{
  const kit=createRackHardwareDetail(THREE,{cube,material}),m=new THREE.Group(),neighbor=new THREE.Group();
  m.userData.rackPart={id:'test:server-1',uStart:3,uHeight:4};
  const sled=cube(m,0,0,0,1,1,1,'#555555');sled.name='rack-shape:sled0';
  const asset={id:'test',type,w:3,h:6,d:3};kit.update(asset,m,true,.1,units);
  assert.equal(kit.state().kind,kind);assert.equal(sled.visible,false);assert.equal(neighbor.children.length,0);
  assert.ok(m.getObjectByName('hardware:green-mainboard'));assert.ok(m.getObjectByName('hardware:fan-grille'));
  assert.equal(kit.explode(),true);for(let i=0;i<20;i++)kit.update(asset,m,true,.1,units);
  assert.ok(kit.state().progress>.99);assert.ok(m.getObjectByName('hardware-layer:cover').position.x < -1.5);assert.ok(m.getObjectByName('hardware-layer:cover').visible);
  assert.equal(kit.explode(),false);for(let i=0;i<20;i++)kit.update(asset,m,true,.1,units);
  assert.ok(kit.state().progress<.01);kit.update(null,null,false,.1);assert.equal(kit.state(),null);assert.equal(sled.visible,true);assert.equal(m.children.length,1);
 });
}

test('Storage rack controls its shelf and drive; server racks control server trays',()=>{
 for(const type of ['GPU server rack','Compute rack','Co-location rack']){
  const actions=mechanicalButtons(type,{targetSlide:1});
  assert.equal(actions[1].code,'server-slide');assert.equal(actions[1].label,'서버 트레이 넣기');
  assert.ok(!actions.some(a=>a.code==='cassette-slide'));
 }
 const storage=mechanicalButtons('Storage rack',{targetSlide:1},{driveTarget:1});
 assert.deepEqual(storage.map(a=>a.code),['front-door','server-slide','drive-toggle','next-drive']);
 assert.equal(storage[1].label,'스토리지 쉘프 넣기');assert.equal(storage[2].label,'드라이브 넣기');
 assert.equal(mechanicalButtons('Top-of-Rack switch')[1].code,'cassette-slide');
});
