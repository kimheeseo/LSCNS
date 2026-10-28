import test from 'node:test';
import assert from 'node:assert/strict';
import {buildStagingRoom} from './staging-room.mjs';
function fixture(){
 const objects=[];
 const add=(type)=>(root,name,...args)=>{const o={type,name,args,rotation:{x:0,y:0,z:0}};objects.push(o);return o;};
 buildStagingRoom({part:add('part'),round:add('round'),decal:add('decal')},{},{w:9,d:5,h:3.5});
 return objects;
}
test('staging workflow has three ordered labeled work areas and Hall destination',()=>{
 const o=fixture(),stages=o.filter(o=>o.name.startsWith('staging-step-'));
 assert.equal(stages.length,3);assert.deepEqual(stages.map(o=>o.args[0]),['① 입고 · 개봉','② 검수 · 설정 · 시험','③ 서버홀 반입 대기']);
 assert.ok(stages[0].args[1]<stages[1].args[1]&&stages[1].args[1]<stages[2].args[1]);
 assert.ok(o.some(o=>o.name==='staging-hall-exit'&&o.args[0].includes('서버홀')));
});
test('receiving, live test bench, and completed transfer equipment are visually distinct',()=>{
 const o=fixture(),names=new Set(o.map(o=>o.name));
 for(const name of ['staging-open-flap-1','staging-packing-foam','staging-packaging-bin','staging-test-server','staging-laptop-screen','staging-test-cable','staging-checklist-paper','staging-transfer-cart','staging-ready-label'])assert.ok(names.has(name),name);
 assert.equal(o.filter(o=>o.name.startsWith('staging-test-port-')).length,6);
 assert.equal(o.filter(o=>o.name.startsWith('staging-cart-wheel-')).length,4);
 assert.ok(o.some(o=>o.type==='decal'&&o.args[0].includes('TEST PASS')));
 assert.ok(o.some(o=>o.type==='decal'&&o.args[0].includes('검수 완료')));
});
test('workshop parts stay inside the room footprint',()=>{
 for(const o of fixture().filter(o=>o.type==='part')){
  const [x,,z,w,,d]=o.args;
  assert.ok(Math.abs(x)+w/2<=4.51,o.name+' horizontal extent');
  assert.ok(Math.abs(z)+d/2<=2.51,o.name+' depth extent');
 }
});
