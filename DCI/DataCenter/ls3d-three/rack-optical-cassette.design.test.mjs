import test from 'node:test';
import assert from 'node:assert/strict';
import * as THREE from '../vendor/three.module.js';
import {createRackOpticalCassette} from './rack-optical-cassette.mjs';

function cube(parent,x,y,z,w,h,d,color){
 const mesh=new THREE.Mesh(new THREE.BoxGeometry(w,h,d),new THREE.MeshStandardMaterial({color}));
 mesh.position.set(x,y+h/2,z);parent.add(mesh);return mesh;
}
function build(kit,id,type){
 const root=new THREE.Group(),other=new THREE.Group(),carrier=new THREE.Group(),drawer=new THREE.Group();
 other.name='rack-shell';carrier.add(drawer);root.add(other,carrier);
 const parts=[],partMap=new Map();
 const session=kit.decorate({id,type,w:2,d:1,h:4},{cassetteDrawer:drawer,cassetteGroup:carrier,
  rackUnits:42,cassetteU:38,parts,partMap});
 return {root,other,carrier,drawer,session,parts,partMap};
}

test('network cassette exposes distinct LC/MPO parts, slide, fan-out and reversible isolation',()=>{
 const kit=createRackOpticalCassette(THREE,{cube});
 const rack=build(kit,'net','Top-of-Rack switch'),s=rack.session;
 assert.equal(s.modules.length,4);
 assert.equal(s.portDefs.length,24);
 assert.ok(rack.drawer.getObjectByName('rack-optical:cassette-front-handle-0'));
 assert.ok(rack.drawer.getObjectByName('rack-optical:mpo-keyed-receptacle-3'));
 assert.ok(kit.owns(s.modules[0].p));
 kit.isolate('net',true);
 assert.equal(rack.other.visible,false);
 assert.equal(rack.carrier.visible,true);
 const closed=s.covers[0].position.z,moving={cassette:rack.carrier,reach:.67};
 kit.update('net',moving,{cassette:1});
 assert.ok(moving.cassette.position.z<0);
 assert.ok(s.covers[0].position.z>closed);
 assert.equal(s.covers[0].rotation.x,0);
 assert.equal(kit.state('net').fibers,48);
 assert.equal(rack.drawer.getObjectByName('rack-optical:fanout-fiber-0-0').userData.partId,s.portDefs[0].id);
 const fiberCount=kit.stats.fibers;
 kit.update('net',moving,{cassette:1});
 assert.equal(kit.stats.fibers,fiberCount); // lazy fiber geometry is created once
 kit.update('net',moving,{cassette:0});
 assert.ok(Math.abs(moving.cassette.position.z)<1e-12);
 assert.equal(s.covers[0].position.z,closed);
 kit.isolate('net',false);
 assert.equal(rack.other.visible,true);
 kit.drop('net');assert.equal(kit.state('net'),null);

 const gpu=build(kit,'gpu','GPU rack');
 kit.update('gpu',{cassette:gpu.carrier,reach:.67},{cassette:1});
 assert.equal(gpu.drawer.getObjectByName('rack-optical:cassette-front-handle-0'),undefined);
 assert.ok(gpu.session.covers[0].rotation.x<0);
 assert.equal(kit.isolate('gpu',true),false);
 kit.drop('gpu');
});
