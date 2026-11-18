import test from 'node:test';
import assert from 'node:assert/strict';
import {registerRackSurfaceParts} from './rack-universal-picking.mjs';

let seq=0;
class Object3D{
 constructor(name='',type='group'){
  this.name=name;this.uuid='fixture-'+(++seq);this.children=[];this.parent=null;
  this.userData={};this.isMesh=type==='mesh';this.isInstancedMesh=type==='instanced';
 }
 add(node){this.children.push(node);node.parent=this;return node;}
 remove(node){this.children=this.children.filter(child=>child!==node);node.parent=null;}
 traverse(fn){fn(this);for(const child of this.children)child.traverse(fn);}
}
function fixture(){
 const g=new Object3D('rack'),high=g.add(new Object3D('rack-high'));
 const parent={id:'rack-test:server-1:3',assetId:'rack-test',kind:'server-1',label:'서버 트레이 1',uStart:3,uHeight:4};
 const port={id:'rack-test:mpo-input-0',assetId:'rack-test',kind:'mpo-input',label:'MPO 포트'};
 const cable={id:'rack-test:cassette-trunk-0',assetId:'rack-test',kind:'cassette-cable',label:'12심 케이블'};
 const power={id:'rack-test:power-link:1:A',assetId:'rack-test',kind:'power-cable',label:'PDU 전력 케이블'};
 const map=new Map([parent,port,cable,power].map(part=>[part.id,part]));
 const parts=[parent,port,cable,power];
 const chassis=high.add(new Object3D('rack-shape:sled0','mesh'));
 const pcb=high.add(new Object3D('hardware:green-mainboard','mesh'));
 const fan=high.add(new Object3D('hardware:fan-group'));
 fan.userData.rackSelectableGroup=true;fan.userData.rackSelectableLabel='냉각팬 1';
 for(let i=0;i<6;i++)fan.add(new Object3D('hardware:fan-blade','mesh'));
 const mpo=high.add(new Object3D('rack-optical:mpo-adapter','mesh'));
 const fiber=high.add(new Object3D('rack-optical:fanout-fiber','mesh'));
 const pduWire=high.add(new Object3D('rack-link:cable:power-A','mesh'));
 for(const mesh of [chassis,pcb])mesh.userData.partId=parent.id;
 mpo.userData.partId=port.id;
 fiber.userData.partId=cable.id;fiber.userData.cassetteCable=true;
 pduWire.userData.partId=power.id;
 const record={g,rackLOD:{high},rackParts:parts,partMap:map};
 return {record,high,fan,chassis,pcb,mpo,fiber,pduWire,port,cable,power,parent};
}
test('every modeled non-decorative rack component has independent selectable ID',()=>{
 const f=fixture(),result=registerRackSurfaceParts({id:'rack-test'},f.record);
 assert.equal(result.added,3);
 assert.notEqual(f.chassis.userData.partId,f.pcb.userData.partId);
 assert.equal(f.record.partMap.get(f.pcb.userData.partId).insideHardwareDetail,true);
 assert.equal(f.record.partMap.get(f.chassis.userData.partId).insideHardwareDetail,false);
 assert.equal(f.record.partMap.get(f.pcb.userData.partId).uStart,3);
 const fanIds=f.fan.children.map(x=>x.userData.partId);
 assert.equal(new Set(fanIds).size,1);
 assert.match(f.record.partMap.get(fanIds[0]).label,/냉각팬/);
 assert.equal(f.fiber.userData.partId,f.cable.id);
 assert.equal(f.mpo.userData.partId,f.port.id);
 assert.equal(f.pduWire.userData.partId,f.power.id);
 assert.equal(registerRackSurfaceParts({id:'rack-test'},f.record).added,0);
});
test('removing a detailed component restores the picker inventory',()=>{
 const f=fixture();
 registerRackSurfaceParts({id:'rack-test'},f.record);
 f.high.remove(f.fan);
 const state=registerRackSurfaceParts({id:'rack-test'},f.record);
 assert.equal(state.removed,1);
 assert.equal(state.total,2);
 assert.equal(f.record.rackParts.filter(x=>x.universal).length,2);
});
