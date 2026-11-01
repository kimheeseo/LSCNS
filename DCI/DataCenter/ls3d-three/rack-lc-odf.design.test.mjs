import test from 'node:test';
import assert from 'node:assert/strict';
import * as THREE from '../vendor/three.module.js';
import {createRackLcAdapterDesign} from './rack-lc-adapter-design.mjs';
import {createOdfSmallFiberBox} from './odf-small-fiber-box.mjs';

function cube(parent,x,y,z,w,h,d,color){
 const mesh=new THREE.Mesh(new THREE.BoxGeometry(w,h,d),new THREE.MeshStandardMaterial({color}));
 mesh.position.set(x,y+h/2,z);parent.add(mesh);return mesh;
}
function model(id,type){
 const high=new THREE.Group(),other=new THREE.Group();
 high.add(other);other.name='rack-frame';
 const part={id:id+':lc-panel',assetId:id,kind:'lc-adapter-panel',label:'LC panel',uStart:36,uHeight:1};
 const panel=new THREE.Group();panel.userData.rackPart=part;
 panel.add(cube(panel,0,2,0,1,.1,.1,'#aaaaaa'));high.add(panel);
 return {asset:{id,type,w:2.5,d:2,h:6},high,built:{parts:[part],partMap:new Map([[part.id,part]]),rackUnits:42},other,panel};
}
test('LC adapter panel is passive and exposes twelve duplex front/rear paths',()=>{
 const kit=createRackLcAdapterDesign(THREE,{cube}),rack=model('tor-a','Top-of-Rack switch');
 const s=kit.decorate(rack.asset,rack.built,rack.high);
 assert.ok(s);assert.equal(kit.state('tor-a').portCount,12);
 assert.equal(s.paths.length,24);
 assert.ok(s.root.getObjectByName('rack-lc:front-lc-socket-0-0'));
 assert.ok(s.root.getObjectByName('rack-lc:rear-lc-connector-11-1'));
 assert.equal(rack.panel.children[0].visible,false);
 kit.isolate('tor-a',true);assert.equal(rack.other.visible,false);
 kit.highlight('tor-a',3);assert.equal(kit.state('tor-a').selectedPort,3);
 assert.equal(s.paths.filter(p=>p.port===3&&p.mesh.material.emissiveIntensity>.5).length,2);
 kit.isolate('tor-a',false);assert.equal(rack.other.visible,true);
 kit.drop('tor-a');assert.equal(kit.state('tor-a'),null);
 assert.equal(kit.decorate(model('gpu','GPU rack').asset,model('gpu','GPU rack').built,new THREE.Group()),null);
});
test('ODF small enclosure closes over six fiber paths and opens on the left hinge',()=>{
 const kit=createOdfSmallFiberBox(THREE,{cube}),rack=model('odf','ODF');
 const s=kit.decorate(rack.asset,rack.built,rack.high);
 assert.ok(s);assert.equal(kit.state('odf').portCount,6);
 assert.equal(s.paths.length,6);assert.equal(s.portMeshes.length,6);
 assert.equal(s.cover.rotation.y,0);
 assert.ok(s.root.getObjectByName('odf-small:splice-sleeve-5'));
 assert.ok(s.root.getObjectByName('odf-small:cable-tie'));
 assert.ok(s.root.getObjectByName('odf-small:mount-bracket-1'));
 kit.isolate('odf',true);assert.equal(rack.other.visible,false);
 kit.toggle('odf');for(let n=0;n<60;n++)kit.update('odf',.016);
 assert.ok(s.cover.rotation.y>2.4);
 kit.highlight('odf',2);assert.equal(kit.state('odf').selectedPort,2);
 assert.ok(s.paths[2].material.emissiveIntensity>.5);
 assert.ok(s.paths[1].material.opacity<.5);
 kit.toggle('odf');for(let n=0;n<60;n++)kit.update('odf',.016);
 assert.ok(s.cover.rotation.y<.01);assert.equal(kit.state('odf').selectedPort,-1);
 kit.isolate('odf',false);assert.equal(rack.other.visible,true);
 kit.drop('odf');assert.equal(kit.state('odf'),null);
});
