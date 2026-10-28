import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const context={};vm.runInNewContext(readFileSync(new URL('../visitor-tour-policy.js',import.meta.url),'utf8'),context);
const policy=context.LS3D_VISITOR_POLICY;
test('gate opens and closes once every five seconds, five times, then stays closed',()=>{
  for(let cycle=0;cycle<5;cycle++){
    assert.equal(policy.gateMotion(cycle*5).openness,0);
    assert.equal(policy.gateMotion(cycle*5+2).openness,1);
    assert.equal(policy.gateMotion(cycle*5+4).openness,0);
    assert.equal(policy.gateMotion(cycle*5+4).completed,cycle);
  }
  for(const t of [25,26,50,1000]){const d=policy.gateMotion(t);assert.equal(d.openness,0);assert.equal(d.completed,5);assert.equal(d.finished,true);}
});
test('gate aligns to Hall doorway centre and remains outside facade',()=>{
  const hall={pos:[18,7],size:[50,42]},p=policy.gatePosition(hall);
  assert.equal(p.x,18);assert.equal(p.z,30.3);
  assert.ok(p.z-4*.41>28.19);
});
test('loading camera faces the carton labels from the cargo side, clear of truck',()=>{
  const a={x:-37,z:18,w:12,d:7};const p=policy.loadingShot(a);
  assert.equal(p.eye[0],a.x-2.4);assert.equal(p.eye[0],p.target[0]);
  assert.ok(p.eye[2]>a.z+a.d/2);assert.ok(p.target[2]<p.eye[2]);
  assert.ok(p.eye[1]<2);assert.ok(p.target[1]<p.eye[1]);
});
test('HTML inline scripts remain valid and load tour policy before visitor controller',()=>{
  const html=readFileSync(new URL('../LS_Datacenter_3D.html',import.meta.url),'utf8');
  for(const [,attrs,body]of html.matchAll(/<script([^>]*)>([\s\S]*?)<\/script>/g))if(!/type=["'](?:module|importmap)/.test(attrs))new vm.Script(body);
  assert.ok(html.indexOf('src="./visitor-tour-policy.js')<html.indexOf('const tourPolicy=window.LS3D_VISITOR_POLICY'));
  assert.match(html,/gateDemo:gateDemo\(\)/);assert.match(html,/index===4\?tourPolicy.gateDuration\*1000/);
});
